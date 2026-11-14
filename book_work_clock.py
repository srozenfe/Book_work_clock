"""Book Work Clock: local, Windows-only time tracking for Word documents."""
from __future__ import annotations

import ctypes
import os
import winreg
import sqlite3
import sys
import time
from datetime import datetime
from pathlib import Path

from PySide6.QtNetwork import QLocalServer, QLocalSocket
from updater import APP_VERSION, check_for_updates

import psutil
from PySide6.QtCore import Qt, QTimer
from manual_time import ManualSessionDialog, ensure_manual_schema
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QDialog, QDialogButtonBox, QFileDialog,
    QFormLayout, QHBoxLayout, QInputDialog, QLabel, QListWidget,
    QMainWindow, QMessageBox, QPushButton, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget, QSystemTrayIcon, QMenu,
)
from PySide6.QtGui import QAction, QIcon

APP_DIR = Path(os.getenv("LOCALAPPDATA", str(Path.home()))) / "BookWorkClock"
DB_FILE = APP_DIR / "work_clock.sqlite3"
IDLE_SECONDS = 15
STARTUP_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
STARTUP_NAME = "BookWorkClock"


def startup_command():
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}" --background'
    return f'"{sys.executable}" "{Path(__file__).resolve()}" --background'


def startup_enabled():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_KEY) as key:
            value, _ = winreg.QueryValueEx(key, STARTUP_NAME)
            return value == startup_command()
    except OSError:
        return False


def configure_startup(enabled):
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, STARTUP_KEY) as key:
        if enabled:
            winreg.SetValueEx(key, STARTUP_NAME, 0, winreg.REG_SZ, startup_command())
        else:
            try:
                winreg.DeleteValue(key, STARTUP_NAME)
            except FileNotFoundError:
                pass


user32 = ctypes.windll.user32 if sys.platform == "win32" else None


def timestamp() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def format_duration(seconds: float) -> str:
    seconds = max(0, int(seconds))
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02}:{minutes:02}:{seconds:02}"


class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]


def last_input_tick() -> int:
    info = LASTINPUTINFO()
    info.cbSize = ctypes.sizeof(info)
    if user32 is not None and user32.GetLastInputInfo(ctypes.byref(info)):
        return info.dwTime
    return 0


def idle_seconds() -> float:
    if user32 is None:
        return 0
    info = LASTINPUTINFO()
    info.cbSize = ctypes.sizeof(info)
    if not user32.GetLastInputInfo(ctypes.byref(info)):
        return 0
    # Tick counts are 32-bit in LASTINPUTINFO and wrap after ~49.7 days.
    current = ctypes.windll.kernel32.GetTickCount64() & 0xFFFFFFFF
    return ((current - info.dwTime) & 0xFFFFFFFF) / 1000.0


DOCUMENT_APPS = {
    "winword.exe": "Word",
    "acrord32.exe": "PDF",
    "acrobat.exe": "PDF",
    "foxitpdfreader.exe": "PDF",
    "foxitreader.exe": "PDF",
    "sumatrapdf.exe": "PDF",
    "msedge.exe": "Browser",
    "chrome.exe": "Browser",
    "firefox.exe": "Browser",
}


def foreground_document():
    """Inspect just the foreground app/window, never typed keys or document contents.

    Return (process, title) even when unmatched so users can diagnose filename
    or application-title differences. Browsers count only registered PDF names.
    """
    if user32 is None:
        return None, ""
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return None, ""
    pid = ctypes.c_ulong()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    try:
        process = psutil.Process(pid.value).name().casefold()
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        return None, ""
    buf = ctypes.create_unicode_buffer(2048)
    user32.GetWindowTextW(hwnd, buf, len(buf))
    return process, buf.value.strip()


def title_matches_document(window_title, path, process):
    """Accept Word title variants with/without extension and common suffixes.

    For browser-based PDFs require .pdf in the selected title, to avoid tracking
    unrelated tabs accidentally sharing a book name.
    """
    ext = Path(path).suffix.casefold()
    app = DOCUMENT_APPS.get(process or "")
    if ext in (".doc", ".docx", ".docm"):
        if app != "Word":
            return False
    elif ext == ".pdf":
        if app not in ("PDF", "Browser"):
            return False
    else:
        return False
    title = window_title.casefold().strip()
    stem = Path(path).stem.casefold().strip()
    name = Path(path).name.casefold().strip()
    # Windows applications normally put the document name before ' - App'.
    # Word can hide the extension, while Acrobat often includes '.pdf'.
    candidates = (name, stem) if app != "Browser" else (name,)
    for candidate in candidates:
        if title == candidate:
            return True
        # Only match at the beginning and require a clear window-title boundary.
        if title.startswith(candidate):
            rest = title[len(candidate):]
            if rest.startswith((" - ", " — ", " – ", " | ", " —", " (")):
                return True
    return False


class Store:
    def __init__(self):
        APP_DIR.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(DB_FILE)
        self.db.execute("PRAGMA journal_mode=WAL")
        self.db.execute("PRAGMA foreign_keys=ON")
        self.db.executescript("""
            CREATE TABLE IF NOT EXISTS books (
              id INTEGER PRIMARY KEY, title TEXT NOT NULL,
              document_path TEXT NOT NULL, archived INTEGER NOT NULL DEFAULT 0
            );
            CREATE TABLE IF NOT EXISTS sessions (
              id INTEGER PRIMARY KEY, book_id INTEGER NOT NULL,
              started_at TEXT NOT NULL, ended_at TEXT NOT NULL,
              seconds REAL NOT NULL DEFAULT 0,
              FOREIGN KEY(book_id) REFERENCES books(id)
            );
            CREATE INDEX IF NOT EXISTS idx_sessions_book ON sessions(book_id);
        """)
        ensure_manual_schema(self.db)
        self.ensure_payment_schema()
        self.db.commit()

    def ensure_payment_schema(self):
        columns = {row[1] for row in self.db.execute("PRAGMA table_info(books)")}
        if "payment" not in columns:
            self.db.execute("ALTER TABLE books ADD COLUMN payment REAL NOT NULL DEFAULT 0")
        if "currency" not in columns:
            self.db.execute("ALTER TABLE books ADD COLUMN currency TEXT NOT NULL DEFAULT 'ILS'")
        self.db.commit()

    def book_payment(self, book_id):
        return self.db.execute(
            "SELECT payment,currency FROM books WHERE id=?", (book_id,)
        ).fetchone()

    def update_payment(self, book_id, payment, currency):
        self.db.execute("UPDATE books SET payment=?,currency=? WHERE id=?",
                        (payment, currency, book_id))
        self.db.commit()

    def books(self):
        return self.db.execute(
            "SELECT id,title,document_path FROM books WHERE archived=0 ORDER BY title"
        ).fetchall()

    def add_book(self, title, path):
        self.db.execute("INSERT INTO books(title,document_path) VALUES (?,?)", (title, path))
        self.db.commit()

    def edit_book(self, book_id, title, path):
        self.db.execute(
            "UPDATE books SET title=?,document_path=? WHERE id=?",
            (title, path, book_id),
        )
        self.db.commit()

    def archive_book(self, book_id):
        # Keep historical sessions, even when a book is removed from the UI.
        self.db.execute("UPDATE books SET archived=1 WHERE id=?", (book_id,))
        self.db.commit()

    def start_segment(self, book_id):
        now = timestamp()
        cur = self.db.execute(
            "INSERT INTO sessions(book_id,started_at,ended_at,seconds) VALUES(?,?,?,0)",
            (book_id, now, now),
        )
        self.db.commit()
        return cur.lastrowid

    def save_segment(self, segment_id, seconds):
        self.db.execute(
            "UPDATE sessions SET ended_at=?,seconds=? WHERE id=?",
            (timestamp(), max(0, seconds), segment_id),
        )
        self.db.commit()

    def add_manual(self, book_id, start, end, note):
        duration = (end - start).total_seconds()
        if duration <= 0:
            raise ValueError("End must be after start")
        self.db.execute(
            "INSERT INTO sessions (book_id,started_at,ended_at,seconds,is_manual,note) VALUES (?,?,?,?,1,?)",
            (book_id, start.isoformat(timespec="seconds"),
             end.isoformat(timespec="seconds"), duration, note)
        )
        self.db.commit()

    def get_manual(self, session_id):
        return self.db.execute(
            "SELECT id,book_id,started_at,ended_at,note FROM sessions WHERE id=? AND is_manual=1",
            (session_id,)
        ).fetchone()

    def edit_manual(self, session_id, book_id, start, end, note):
        duration = (end - start).total_seconds()
        if duration <= 0:
            raise ValueError("End must be after start")
        self.db.execute(
            """UPDATE sessions SET book_id=?, started_at=?, ended_at=?, seconds=?, note=?
               WHERE id=? AND is_manual=1""",
            (book_id, start.isoformat(timespec="seconds"),
             end.isoformat(timespec="seconds"), duration, note, session_id)
        )
        self.db.commit()

    def remove_manual(self, session_id):
        self.db.execute("DELETE FROM sessions WHERE id=? AND is_manual=1", (session_id,))
        self.db.commit()

    def history(self):
        return self.db.execute("""
            SELECT s.id,b.title,s.started_at,s.ended_at,s.seconds,s.is_manual,s.note
            FROM sessions s JOIN books b ON b.id=s.book_id
            ORDER BY s.id DESC LIMIT 300
        """).fetchall()

    def totals(self):
        return dict(self.db.execute(
            "SELECT book_id,COALESCE(SUM(seconds),0) FROM sessions GROUP BY book_id"
        ).fetchall())


class BookDialog(QDialog):
    def __init__(self, parent=None, title="", path=""):
        super().__init__(parent)
        self.setWindowTitle("Book details")
        self.title_label = QLabel(title or "New book")
        from PySide6.QtWidgets import QLineEdit
        self.title_input = QLineEdit(title)
        self.path_input = QLineEdit(path)
        browse = QPushButton("Browse…")
        browse.clicked.connect(self.browse)
        path_row = QHBoxLayout()
        path_row.addWidget(self.path_input)
        path_row.addWidget(browse)
        form = QFormLayout()
        form.addRow("Book title:", self.title_input)
        form.addRow("Word or PDF document:", path_row)
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.validate)
        buttons.rejected.connect(self.reject)
        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(buttons)

    def browse(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Choose book document", "", "Documents (*.docx *.doc *.docm *.pdf)"
        )
        if path:
            self.path_input.setText(path)
            if not self.title_input.text().strip():
                self.title_input.setText(Path(path).stem)

    def validate(self):
        if not self.title_input.text().strip() or not self.path_input.text().strip():
            QMessageBox.warning(self, "Missing information", "Enter a title and document.")
            return
        self.accept()

    def values(self):
        return self.title_input.text().strip(), self.path_input.text().strip()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.store = Store()
        self.current_book = None
        self.segment_id = None
        self.segment_started = None
        self.auto_enabled = True
        self.last_idle_alert = False
        self.last_nonword_alert = False
        self.continue_outside_word = False
        self.resume_guard = None  # last input when paused
        self.pause_needs_leave = False
        self.idle_prompt_suppressed = False
        self.last_foreground_signature = None
        self.setWindowTitle(f"Book Work Clock v{APP_VERSION}")
        self.resize(1030, 720)
        self.setStyleSheet("""
            QMainWindow, QWidget { background: #f4f7fb; color: #1d2d45; }
            QLabel { font-size: 13px; }
            QPushButton { background: #e5edfa; border: 1px solid #d5dfed;
                          border-radius: 7px; padding: 8px 12px; }
            QPushButton:hover { background: #d8e6fa; }
            QListWidget, QTableWidget { background: white;
                          border: 1px solid #d9e3f1; border-radius: 8px;
                          selection-background-color: #d6e7ff; }
            QHeaderView::section { background: #edf2fa; padding: 8px;
                                  border: none; font-weight: bold; }
        """)
        self.build_ui()
        self.refresh_books()
        self.refresh_history()

        self.ticker = QTimer(self)
        self.ticker.timeout.connect(self.tick)
        self.ticker.start(1000)

    def build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        heading = QLabel(f"Book Work Clock  •  Writing Dashboard  •  v{APP_VERSION}")
        heading.setStyleSheet("font-size: 26px; font-weight: bold; padding: 8px;")
        outer.addWidget(heading)
        self.status = QLabel("Ready — open a registered Word or PDF document to start")
        self.detection = QLabel("Detection: waiting for a document")
        self.detection.setStyleSheet("color: #65758b; font-size: 11px;")
        outer.addWidget(self.status)
        outer.addWidget(self.detection)
        self.clock = QLabel("00:00:00")
        self.clock.setStyleSheet("font-size: 40px; font-weight: bold; color: #2563c8; padding: 8px;")
        outer.addWidget(self.clock)

        controls = QHBoxLayout()
        self.start_btn = QPushButton("Start / Resume")
        self.pause_btn = QPushButton("Pause")
        self.stop_btn = QPushButton("Stop")
        self.auto_box = QCheckBox("Automatic tracking enabled")
        self.auto_box.setChecked(True)
        self.auto_box.toggled.connect(self.set_auto)
        self.startup_box = QCheckBox("Start minimized with Windows")
        self.startup_box.setChecked(startup_enabled())
        self.startup_box.toggled.connect(self.set_startup)
        self.start_btn.clicked.connect(self.manual_start)
        self.pause_btn.clicked.connect(self.user_pause)
        self.stop_btn.clicked.connect(self.stop)
        for widget in (self.start_btn, self.pause_btn, self.stop_btn, self.auto_box):
            controls.addWidget(widget)
        outer.addLayout(controls)
        outer.addWidget(self.startup_box)
        self.update_btn = QPushButton("Check GitHub for updates")
        self.update_btn.clicked.connect(self.check_updates)
        outer.addWidget(self.update_btn)

        outer.addWidget(QLabel("Books (select a book to start manually):"))
        self.books_list = QListWidget()
        outer.addWidget(self.books_list)
        book_buttons = QHBoxLayout()
        for name, handler in (
            ("Add book", self.add_book),
            ("Edit", self.edit_book),
            ("Remove", self.remove_book),
            ("Payment / earnings", self.edit_payment),
        ):
            button = QPushButton(name)
            button.clicked.connect(handler)
            book_buttons.addWidget(button)
        outer.addLayout(book_buttons)
        self.earnings_label = QLabel("Select a book to see earnings")
        self.earnings_label.setStyleSheet(
            "font-size: 16px; font-weight: bold; padding: 10px; color: #1762c0;"
        )
        outer.addWidget(self.earnings_label)
        self.books_list.currentItemChanged.connect(
            lambda current, previous: self.update_earnings()
        )
        outer.addWidget(QLabel("Recent tracked periods:"))
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ["Book", "Start", "End", "Duration", "Type", "Notes"]
        )
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        outer.addWidget(self.table)
        manual_row = QHBoxLayout()
        for label, fn in (
            ("+ Add manual time", self.add_manual_entry),
            ("Edit manual time", self.edit_manual_entry),
            ("Remove manual time", self.remove_manual_entry),
        ):
            button = QPushButton(label)
            button.clicked.connect(fn)
            manual_row.addWidget(button)
        outer.addLayout(manual_row)
        refresh = QPushButton("Refresh history")
        refresh.clicked.connect(self.refresh_history)
        outer.addWidget(refresh)

        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.style().standardIcon(self.style().StandardPixmap.SP_ComputerIcon))
        self.tray.setToolTip("Book Work Clock")
        tray_menu = QMenu()
        show = QAction("Show window", self)
        show.triggered.connect(self.showNormal)
        quit_action = QAction("Exit", self)
        quit_action.triggered.connect(self.exit_app)
        tray_menu.addAction(show)
        tray_menu.addAction(quit_action)
        self.tray.setContextMenu(tray_menu)
        self.tray.activated.connect(
            lambda reason: self.showNormal()
            if reason == QSystemTrayIcon.ActivationReason.DoubleClick else None
        )
        if QSystemTrayIcon.isSystemTrayAvailable():
            self.tray.show()

    def selected_book(self):
        item = self.books_list.currentItem()
        return item.data(Qt.ItemDataRole.UserRole) if item else None

    def refresh_books(self):
        selected = self.selected_book()
        self.books_list.clear()
        from PySide6.QtWidgets import QListWidgetItem
        totals = self.store.totals()
        for book_id, title, path in self.store.books():
            item = QListWidgetItem(
                f"{title}    —    {format_duration(totals.get(book_id, 0))}"
            )
            item.setData(Qt.ItemDataRole.UserRole, book_id)
            item.setToolTip(path)
            self.books_list.addItem(item)
            if book_id == selected:
                self.books_list.setCurrentItem(item)

    def update_earnings(self):
        book_id = self.selected_book()
        if book_id is None:
            self.earnings_label.setText("Select a book to see earnings")
            return
        payment, currency = self.store.book_payment(book_id)
        total_seconds = self.store.totals().get(book_id, 0)
        hours = total_seconds / 3600
        symbol = {"ILS": "₪", "USD": "$", "EUR": "€", "GBP": "£"}.get(currency, currency)
        hourly = payment / hours if hours > 0 else None
        self.earnings_label.setText(
            f"Payment: {symbol}{payment:,.2f}   |   "
            f"Tracked: {hours:.2f} hours   |   "
            f"Effective hourly earnings: "
            + (f"{symbol}{hourly:,.2f}/hour" if hourly is not None else "N/A (no work time yet)")
        )

    def edit_payment(self):
        book_id = self.selected_book()
        if book_id is None:
            QMessageBox.information(self, "Select a book", "Select a book first.")
            return
        payment, currency = self.store.book_payment(book_id)
        from PySide6.QtWidgets import QDoubleSpinBox, QComboBox
        dialog = QDialog(self)
        dialog.setWindowTitle("Book payment")
        form = QFormLayout(dialog)
        amount = QDoubleSpinBox()
        amount.setDecimals(2)
        amount.setRange(0, 999999999)
        amount.setValue(payment)
        money = QComboBox()
        for item in ("ILS", "USD", "EUR", "GBP"):
            money.addItem(item)
        money.setCurrentText(currency)
        form.addRow("Total payment for this book:", amount)
        form.addRow("Currency:", money)
        form.addRow(QLabel(
            "Effective hourly earnings = total payment / all tracked hours.\n"
            "Both automatic and manual work entries are included."
        ))
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save |
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        form.addRow(buttons)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.store.update_payment(book_id, amount.value(), money.currentText())
            self.update_earnings()

    def check_updates(self):
        if check_for_updates(self):
            self.exit_app()

    def refresh_history(self):
        rows = self.store.history()
        self.table.setRowCount(len(rows))
        for i, (session_id, title, start, end, seconds, manual, note) in enumerate(rows):
            values = (
                title, start.replace("T", " ")[:19], end.replace("T", " ")[:19],
                format_duration(seconds), "Manual" if manual else "Automatic", note
            )
            for j, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if j == 0:
                    item.setData(Qt.ItemDataRole.UserRole, session_id)
                self.table.setItem(i, j, item)
        self.refresh_books()
        self.update_earnings()

    def selected_session(self):
        row = self.table.currentRow()
        if row < 0:
            return None
        item = self.table.item(row, 0)
        return item.data(Qt.ItemDataRole.UserRole) if item else None

    def add_manual_entry(self):
        dialog = ManualSessionDialog(self.store.books(), self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.store.add_manual(*dialog.values())
            self.refresh_history()

    def edit_manual_entry(self):
        session_id = self.selected_session()
        record = self.store.get_manual(session_id) if session_id is not None else None
        if record is None:
            QMessageBox.information(self, "Select a manual entry",
                                    "Only manual time entries can be edited here.")
            return
        dialog = ManualSessionDialog(self.store.books(), self, record)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.store.edit_manual(session_id, *dialog.values())
            self.refresh_history()

    def remove_manual_entry(self):
        session_id = self.selected_session()
        record = self.store.get_manual(session_id) if session_id is not None else None
        if record is None:
            QMessageBox.information(self, "Select a manual entry",
                                    "Only manual time entries can be removed here.")
            return
        if QMessageBox.question(
            self, "Remove manual time", "Delete the selected manual entry?"
        ) == QMessageBox.StandardButton.Yes:
            self.store.remove_manual(session_id)
            self.refresh_history()

    def set_startup(self, enabled):
        try:
            configure_startup(enabled)
        except OSError as exc:
            QMessageBox.warning(self, "Startup setting failed", str(exc))
            self.startup_box.blockSignals(True)
            self.startup_box.setChecked(startup_enabled())
            self.startup_box.blockSignals(False)

    def add_book(self):
        dialog = BookDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.store.add_book(*dialog.values())
            self.refresh_books()

    def edit_book(self):
        book_id = self.selected_book()
        if book_id is None:
            return
        book = next((b for b in self.store.books() if b[0] == book_id), None)
        if not book:
            return
        dialog = BookDialog(self, book[1], book[2])
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.store.edit_book(book_id, *dialog.values())
            self.refresh_books()

    def remove_book(self):
        book_id = self.selected_book()
        if book_id is None:
            return
        reply = QMessageBox.question(
            self, "Remove book", "Archive this book? Its time history will be preserved."
        )
        if reply == QMessageBox.StandardButton.Yes:
            if self.current_book == book_id:
                self.pause()
                self.current_book = None
            self.store.archive_book(book_id)
            self.refresh_history()

    def set_auto(self, enabled):
        self.auto_enabled = enabled
        if not enabled and self.segment_id is not None:
            self.user_pause()
        self.status.setText(
            "Automatic tracking enabled" if enabled else "Manual mode — auto-start disabled"
        )

    def find_matching_book(self, process, title):
        matches = [
            book_id for book_id, _, path in self.store.books()
            if title_matches_document(title, path, process)
        ]
        # Ambiguous titles never automatically charge the wrong book.
        return matches[0] if len(matches) == 1 else None

    def begin(self, book_id):
        if self.segment_id is not None and self.current_book == book_id:
            return
        self.pause()
        self.current_book = book_id
        self.resume_guard = None
        self.pause_needs_leave = False
        self.idle_prompt_suppressed = False
        self.segment_id = self.store.start_segment(book_id)
        self.segment_started = time.monotonic()
        self.last_idle_alert = False
        self.last_nonword_alert = False
        self.continue_outside_word = False
        self.status.setText("Tracking writing time")
        self.refresh_history()

    def pause(self):
        if self.segment_id is not None:
            self.store.save_segment(
                self.segment_id, time.monotonic() - self.segment_started
            )
            self.segment_id = None
            self.segment_started = None
            self.refresh_history()
        self.status.setText("Paused")

    def user_pause(self):
        self.pause()
        self.resume_guard = last_input_tick()
        # Don't instantly resume simply because clicking Pause returns focus to
        # a document. Require a transition away and then back to that document.
        self.pause_needs_leave = True

    def stop(self):
        self.pause()
        self.current_book = None
        self.resume_guard = last_input_tick()
        self.pause_needs_leave = True
        self.status.setText("Stopped — automatic tracking remains available")

    def manual_start(self):
        book_id = self.selected_book() or self.current_book
        if book_id is None:
            QMessageBox.information(self, "Select a book", "Select a book first.")
            return
        self.auto_box.setChecked(True)
        self.resume_guard = None
        self.pause_needs_leave = False
        self.begin(book_id)

    def interrupt(self, reason):
        # The alert itself steals focus; don't poll again until it is dismissed.
        if self.segment_id is None:
            return
        dialog = QMessageBox(None)
        dialog.setWindowFlags(
            dialog.windowFlags() | Qt.WindowType.WindowStaysOnTopHint
        )
        dialog.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose, False)
        dialog.setWindowTitle("Book Work Clock — interruption detected")
        dialog.setText(reason)
        dialog.setInformativeText("Should the writing timer keep running?")
        keep = dialog.addButton("Keep tracking", QMessageBox.ButtonRole.AcceptRole)
        pause = dialog.addButton("Pause", QMessageBox.ButtonRole.ActionRole)
        stop = dialog.addButton("Stop", QMessageBox.ButtonRole.ActionRole)
        disable = dialog.addButton(
            "Pause + don't ask again", QMessageBox.ButtonRole.DestructiveRole
        )
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()
        dialog.exec()
        chosen = dialog.clickedButton()
        if chosen == pause:
            self.user_pause()
        elif chosen == stop:
            self.stop()
        elif chosen == disable:
            self.user_pause()
            self.auto_box.setChecked(False)
        elif chosen == keep:
            self.continue_outside_word = True
            self.idle_prompt_suppressed = True
        else:
            # Dismissing the dialog means no change to running time.
            self.continue_outside_word = True
            self.idle_prompt_suppressed = True

    def tick(self):
        if self.segment_id is not None:
            seconds = time.monotonic() - self.segment_started
            self.store.save_segment(self.segment_id, seconds)
            self.clock.setText(format_duration(seconds))
            self.update_earnings()
        else:
            self.clock.setText("00:00:00")

        process, title = foreground_document()
        if process == "bookworkclock.exe" or (
            user32.GetForegroundWindow() == int(self.winId())
        ):
            self.detection.setText("Detection: Book Work Clock window")
            # The user has been in the tracker window; returning to a
            # registered document is an explicit foreground transition.
            self.pause_needs_leave = False
            return
        book_id = self.find_matching_book(process, title)
        idle = idle_seconds() >= IDLE_SECONDS
        if book_id is not None:
            book_label = next(
                (name for ident, name, _ in self.store.books() if ident == book_id),
                "Document"
            )
            self.detection.setText(f"Detected: {book_label} ({DOCUMENT_APPS.get(process)})")
        elif process in DOCUMENT_APPS:
            self.detection.setText(
                f"Not matched: {process} | {title[:100] or '(no title)'}"
            )
        else:
            self.detection.setText("Detection: no registered document in front")

        if not idle:
            self.last_idle_alert = False
            self.idle_prompt_suppressed = False

        # After Pause, auto-resume is allowed on a *return* to the document,
        # not simply while the same window remains active.
        if self.pause_needs_leave and book_id != self.current_book:
            self.pause_needs_leave = False

        if book_id is not None:
            self.last_nonword_alert = False
            if self.current_book != book_id:
                self.continue_outside_word = False

        if self.auto_enabled and not idle and book_id is not None:
            if not self.pause_needs_leave:
                if book_id != self.current_book or self.segment_id is None:
                    self.begin(book_id)

        if self.segment_id is None:
            return
        if idle and not self.last_idle_alert and not self.idle_prompt_suppressed:
            self.last_idle_alert = True
            self.interrupt("No keyboard or mouse activity for 15 seconds.")
        elif (
            book_id is None and not idle
            and not self.last_nonword_alert and not self.continue_outside_word
        ):
            self.last_nonword_alert = True
            self.interrupt("The tracked Word or PDF document is no longer in front.")

    def exit_app(self):
        self.pause()
        self.tray.hide()
        QApplication.instance().quit()

    def closeEvent(self, event):
        if QSystemTrayIcon.isSystemTrayAvailable():
            self.hide()
            event.ignore()
        else:
            self.pause()
            event.accept()


def main():
    if sys.platform != "win32":
        raise SystemExit("Book Work Clock requires Windows 11.")
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    # Only one process may track work. A desktop-launch request is delivered
    # to the existing background process rather than starting another timer.
    server_name = "BookWorkClock-srozenfe-single-instance"
    client = QLocalSocket()
    client.connectToServer(server_name)
    if client.waitForConnected(450):
        if "--background" not in sys.argv:
            client.write(b"show-and-update")
            client.flush()
            client.waitForBytesWritten(500)
        client.disconnectFromServer()
        return
    QLocalServer.removeServer(server_name)
    server = QLocalServer()
    if not server.listen(server_name):
        raise SystemExit("Book Work Clock is already running or unavailable.")
    window = MainWindow()
    def handle_connection():
        while server.hasPendingConnections():
            incoming = server.nextPendingConnection()
            incoming.waitForReadyRead(200)
            command = bytes(incoming.readAll()).decode("utf-8", "replace")
            incoming.disconnectFromServer()
            if command == "show-and-update":
                window.showNormal()
                window.raise_()
                window.activateWindow()
                QTimer.singleShot(150, window.check_updates)
    server.newConnection.connect(handle_connection)
    if "--background" not in sys.argv or not QSystemTrayIcon.isSystemTrayAvailable():
        window.show()
        QTimer.singleShot(300, window.check_updates)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
