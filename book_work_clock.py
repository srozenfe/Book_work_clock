"""Book Work Clock: local, Windows-only time tracking for Word documents."""
from __future__ import annotations

import ctypes
from ctypes import wintypes
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
    QTabWidget, QRadioButton, QButtonGroup, QGroupBox,
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
if user32 is not None:
    # ctypes defaults to a signed 32-bit return value; HWND is pointer-sized.
    # Without these signatures, a 64-bit foreground HWND may be truncated.
    user32.GetForegroundWindow.restype = wintypes.HWND
    user32.GetWindowThreadProcessId.argtypes = [
        wintypes.HWND, ctypes.POINTER(wintypes.DWORD)
    ]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    user32.GetWindowTextW.argtypes = [
        wintypes.HWND, wintypes.LPWSTR, ctypes.c_int
    ]
    user32.GetWindowTextW.restype = ctypes.c_int



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
    "pdfxedit.exe": "PDF",
    "pdfxcview.exe": "PDF",
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


def resolve_book_override(matched_book, selected_book, origin_hwnd, hwnd, left_window):
    """Honor the chosen book until the user switches windows.

    Only the selected book accumulates time during this temporary override.
    When a new registered document becomes foreground after leaving the
    original window, normal automatic document detection resumes.
    """
    if selected_book is None:
        return matched_book, None, False
    left_window = left_window or hwnd != origin_hwnd
    if left_window and matched_book is not None:
        return matched_book, None, False
    if matched_book is not None:
        return selected_book, selected_book, left_window
    return None, selected_book, left_window


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
        self.ensure_document_schema()
        self.ensure_payment_schema()
        self.db.commit()

    def ensure_document_schema(self):
        """Migrate the single legacy path into an optional Word/PDF collection.

        Idempotent: running this migration again never duplicates a document.
        All existing books, sessions, and historical rates stay untouched.
        """
        self.db.execute("""
            CREATE TABLE IF NOT EXISTS book_documents (
                id INTEGER PRIMARY KEY,
                book_id INTEGER NOT NULL REFERENCES books(id),
                document_path TEXT NOT NULL,
                UNIQUE(book_id, document_path)
            )
        """)
        self.db.execute("""
            INSERT OR IGNORE INTO book_documents(book_id, document_path)
            SELECT id, document_path FROM books
            WHERE document_path IS NOT NULL AND TRIM(document_path) != ''
              AND NOT EXISTS (
                  SELECT 1 FROM book_documents bd WHERE bd.book_id=books.id
              )
        """)
        self.db.commit()

    def document_paths(self, book_id):
        return [row[0] for row in self.db.execute(
            "SELECT document_path FROM book_documents WHERE book_id=? ORDER BY id",
            (book_id,)
        )]

    def set_document_paths(self, book_id, paths):
        if not paths:
            raise ValueError("Select at least one Word or PDF file")
        self.db.execute("DELETE FROM book_documents WHERE book_id=?", (book_id,))
        self.db.executemany(
            "INSERT INTO book_documents(book_id, document_path) VALUES (?,?)",
            [(book_id, path) for path in dict.fromkeys(paths)]
        )
        # Keep old releases/migration compatible, without relying on this path
        # for tracking. New releases use the book_documents table exclusively.
        self.db.execute(
            "UPDATE books SET document_path=? WHERE id=?", (paths[0], book_id)
        )

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

    def archived_books(self):
        return self.db.execute(
            "SELECT id,title,document_path FROM books WHERE archived=1 ORDER BY title"
        ).fetchall()

    def restore_book(self, book_id):
        archived = self.db.execute(
            "SELECT title,archived FROM books WHERE id=?", (book_id,)
        ).fetchone()
        if not archived or not archived[1]:
            raise ValueError("This book is not in the archive.")
        paths = self.document_paths(book_id)
        conflicts = [
            record for record in self._find_conflicts(
                archived[0], paths, exclude_id=book_id
            ) if not record[1]
        ]
        if conflicts:
            raise ValueError(
                "This book or its document already belongs to an active book. "
                "Edit the conflicting active book before restoring."
            )
        with self.db:
            self.db.execute("UPDATE books SET archived=0 WHERE id=?", (book_id,))

    @staticmethod
    def _path_key(path):
        return os.path.normcase(os.path.normpath(path.strip())).casefold()

    def _find_conflicts(self, title, paths, exclude_id=None):
        """Return matching books, matching filenames before matching titles."""
        new_paths = {self._path_key(path) for path in paths}
        matches = []
        for book_id, existing_title, archived in self.db.execute(
            "SELECT id,title,archived FROM books ORDER BY id"
        ).fetchall():
            if book_id == exclude_id:
                continue
            old_paths = {
                self._path_key(path) for path in self.document_paths(book_id)
            }
            same_file = bool(old_paths.intersection(new_paths))
            same_title = existing_title.strip().casefold() == title.strip().casefold()
            if same_file or same_title:
                matches.append((book_id, bool(archived), same_file, same_title))
        return sorted(matches, key=lambda m: (not m[2], m[0]))

    def add_book(self, title, paths):
        """Restore an archived book when name or document matches, keeping history.

        Returns (book_id, restored). Never duplicate an active book; never
        silently merge ambiguous archived records into the wrong history.
        """
        conflicts = self._find_conflicts(title, paths)
        active = [item for item in conflicts if not item[1]]
        if active:
            raise ValueError(
                "This book or document already exists in the book list. "
                "Select the existing book and use Edit to change its files."
            )
        archived = [item for item in conflicts if item[1]]
        if len(archived) > 1:
            raise ValueError(
                "Multiple archived books match this name or document. "
                "Please choose a distinct name or restore one archived record."
            )
        with self.db:
            if archived:
                book_id = archived[0][0]
                self.db.execute(
                    "UPDATE books SET title=?, archived=0 WHERE id=?",
                    (title, book_id),
                )
                self.set_document_paths(book_id, paths)
                return book_id, True
            cur = self.db.execute(
                "INSERT INTO books(title,document_path) VALUES (?,?)",
                (title, paths[0]),
            )
            self.set_document_paths(cur.lastrowid, paths)
            return cur.lastrowid, False

    def edit_book(self, book_id, title, paths):
        if self._find_conflicts(title, paths, exclude_id=book_id):
            raise ValueError(
                "Another book already has this name or one of these documents."
            )
        with self.db:
            self.db.execute("UPDATE books SET title=? WHERE id=?", (title, book_id))
            self.set_document_paths(book_id, paths)

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

    def get_session(self, session_id):
        return self.db.execute(
            """SELECT id,book_id,started_at,ended_at,note,is_manual
               FROM sessions WHERE id=?""", (session_id,)
        ).fetchone()

    def edit_session(self, session_id, book_id, start, end, note):
        """Edit automatic or manual time while preserving capture provenance."""
        duration = (end - start).total_seconds()
        if duration <= 0:
            raise ValueError("End must be after start")
        with self.db:
            result = self.db.execute(
                """UPDATE sessions
                   SET book_id=?,started_at=?,ended_at=?,seconds=?,note=?,
                       edited_at=?
                   WHERE id=?""",
                (book_id, start.isoformat(timespec="seconds"),
                 end.isoformat(timespec="seconds"), duration, note,
                 timestamp(), session_id),
            )
        if result.rowcount == 0:
            raise ValueError("The selected time entry no longer exists.")

    def remove_session(self, session_id):
        """Delete any session type. Confirmation occurs in the UI."""
        with self.db:
            result = self.db.execute(
                "DELETE FROM sessions WHERE id=?", (session_id,)
            )
        return result.rowcount > 0

    def history(self, book_id=None):
        """Show one selected book's history, or all books when none selected."""
        select = """
            SELECT s.id,b.title,s.started_at,s.ended_at,s.seconds,
                   s.is_manual,s.note,s.edited_at
            FROM sessions s JOIN books b ON b.id=s.book_id
        """
        if book_id is not None:
            return self.db.execute(
                select + " WHERE s.book_id=? ORDER BY s.started_at DESC,s.id DESC",
                (book_id,),
            ).fetchall()
        return self.db.execute(
            select + " ORDER BY s.started_at DESC,s.id DESC"
        ).fetchall()

    def totals(self):
        return dict(self.db.execute(
            "SELECT book_id,COALESCE(SUM(seconds),0) FROM sessions GROUP BY book_id"
        ).fetchall())


class BookDialog(QDialog):
    """A book can track one Word file, one PDF file, or both."""

    def __init__(self, parent=None, title="", paths=None):
        super().__init__(parent)
        self.setWindowTitle("Book details")
        from PySide6.QtWidgets import QLineEdit
        self.title_input = QLineEdit(title)
        self.word_input = QLineEdit()
        self.pdf_input = QLineEdit()
        self.word_input.setPlaceholderText("Optional .doc, .docx or .docm file")
        self.pdf_input.setPlaceholderText("Optional .pdf file")
        for path in paths or ():
            if Path(path).suffix.casefold() == ".pdf":
                self.pdf_input.setText(path)
            else:
                self.word_input.setText(path)
        form = QFormLayout()
        form.addRow("Book title:", self.title_input)
        for label, widget, kind in (
            ("Word document:", self.word_input, "word"),
            ("PDF document:", self.pdf_input, "pdf")
        ):
            row = QHBoxLayout()
            row.addWidget(widget)
            browse = QPushButton("Browse…")
            browse.clicked.connect(lambda checked=False, k=kind: self.browse(k))
            row.addWidget(browse)
            form.addRow(label, row)
        form.addRow(
            QLabel("Choose Word, PDF, or both. Time in either file counts toward this book.")
        )
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.validate)
        buttons.rejected.connect(self.reject)
        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(buttons)

    def browse(self, kind):
        filt = (
            "Word documents (*.doc *.docx *.docm)" if kind == "word"
            else "PDF documents (*.pdf)"
        )
        path, _ = QFileDialog.getOpenFileName(
            self, "Choose " + kind.upper() + " document", "", filt
        )
        if path:
            (self.word_input if kind == "word" else self.pdf_input).setText(path)
            if not self.title_input.text().strip():
                self.title_input.setText(Path(path).stem)

    def validate(self):
        if not self.title_input.text().strip():
            QMessageBox.warning(self, "Missing title", "Enter a book title.")
            return
        paths = self.paths()
        if not paths:
            QMessageBox.warning(
                self, "Missing document", "Choose at least a Word or PDF file."
            )
            return
        for path in paths:
            ext = Path(path).suffix.casefold()
            allowed = (".pdf",) if path == self.pdf_input.text().strip() else (
                ".doc", ".docx", ".docm"
            )
            if ext not in allowed:
                QMessageBox.warning(
                    self, "Invalid document",
                    f"The chosen document has an unsupported extension: {path}"
                )
                return
        self.accept()

    def paths(self):
        return [
            path for path in (
                self.word_input.text().strip(), self.pdf_input.text().strip()
            ) if path
        ]

    def values(self):
        return self.title_input.text().strip(), self.paths()


class InterruptionDialog(QDialog):
    """Foreground interruption choice with an explicit active-book selection.

    Exactly one book can receive work time. This prevents accidentally
    double-counting simultaneous Word and PDF windows.
    """
    def __init__(self, books, current_book, reason, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Book Work Clock — interruption detected")
        self.setWindowFlags(
            self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint
        )
        self.setMinimumWidth(430)
        self.action = None
        layout = QVBoxLayout(self)
        reason_label = QLabel(reason)
        reason_label.setWordWrap(True)
        layout.addWidget(reason_label)
        layout.addWidget(QLabel("Choose the book that should continue counting:"))
        group_box = QGroupBox("Active books")
        choices_layout = QVBoxLayout(group_box)
        self.group = QButtonGroup(self)
        self.radio_buttons = {}
        for book_id, title, _ in books:
            radio = QRadioButton(title)
            self.group.addButton(radio, book_id)
            self.radio_buttons[book_id] = radio
            choices_layout.addWidget(radio)
        layout.addWidget(group_box)
        if current_book in self.radio_buttons:
            self.radio_buttons[current_book].setChecked(True)
        elif books:
            self.radio_buttons[books[0][0]].setChecked(True)

        buttons = QHBoxLayout()
        for label, action in (
            ("Keep tracking selected book", "keep"),
            ("Pause", "pause"),
            ("Stop", "stop"),
            ("Pause + don't ask again", "disable"),
        ):
            button = QPushButton(label)
            button.clicked.connect(
                lambda checked=False, choice=action: self.choose(choice)
            )
            buttons.addWidget(button)
        layout.addLayout(buttons)

    def chosen_book(self):
        value = self.group.checkedId()
        return value if value >= 0 else None

    def choose(self, action):
        if action == "keep" and self.chosen_book() is None:
            QMessageBox.warning(self, "Choose a book", "Select an active book.")
            return
        self.action = action
        self.accept()


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
        self.book_override = None  # explicit book picked in Keep tracking popup
        self.override_origin_hwnd = None
        self.override_seen_other_window = False
        self.setWindowTitle(f"Book Work Clock v{APP_VERSION}")
        self.resize(1030, 800)
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

        self.book_tabs = QTabWidget()
        active_tab = QWidget()
        active_layout = QVBoxLayout(active_tab)
        active_layout.addWidget(QLabel("Active books — select a book to see its history:"))
        self.books_list = QListWidget()
        active_layout.addWidget(self.books_list)
        book_buttons = QHBoxLayout()
        for name, handler in (
            ("Add book", self.add_book),
            ("Edit", self.edit_book),
            ("Send to archive", self.remove_book),
            ("Payment / earnings", self.edit_payment),
        ):
            button = QPushButton(name)
            button.clicked.connect(handler)
            book_buttons.addWidget(button)
        active_layout.addLayout(book_buttons)
        archive_tab = QWidget()
        archive_layout = QVBoxLayout(archive_tab)
        archive_layout.addWidget(QLabel("Archived books — work history is preserved:"))
        self.archive_list = QListWidget()
        archive_layout.addWidget(self.archive_list)
        restore_button = QPushButton("Return selected book to Active")
        restore_button.clicked.connect(self.restore_archived_book)
        archive_layout.addWidget(restore_button)
        self.book_tabs.addTab(active_tab, "Active")
        self.book_tabs.addTab(archive_tab, "Archive")
        outer.addWidget(self.book_tabs)
        self.earnings_label = QLabel("Select a book to see earnings")
        self.earnings_label.setStyleSheet(
            "font-size: 16px; font-weight: bold; padding: 10px; color: #1762c0;"
        )
        outer.addWidget(self.earnings_label)
        self.books_list.currentItemChanged.connect(
            lambda current, previous: self.on_book_selected()
        )
        self.history_heading = QLabel("Work history — selected book")
        outer.addWidget(self.history_heading)
        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(
            ["Book", "Start", "End", "Duration", "Type", "Notes"]
        )
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.verticalHeader().setDefaultSectionSize(32)
        self.table.setMinimumHeight(225)  # header plus at least 5 session rows
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        outer.addWidget(self.table)
        manual_row = QHBoxLayout()
        for label, fn in (
            ("+ Add manual time", self.add_manual_entry),
            ("Edit selected time", self.edit_time_entry),
            ("Delete selected time", self.remove_time_entry),
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

    def refresh_books(self, select_book_id=None):
        """Keep selection stable while rebuilding book totals and labels."""
        selected = select_book_id if select_book_id is not None else self.selected_book()
        self.books_list.blockSignals(True)
        try:
            self.books_list.clear()
            from PySide6.QtWidgets import QListWidgetItem
            totals = self.store.totals()
            for book_id, title, path in self.store.books():
                item = QListWidgetItem(
                    f"{title}    —    {format_duration(totals.get(book_id, 0))}"
                )
                item.setData(Qt.ItemDataRole.UserRole, book_id)
                item.setToolTip("\\n".join(self.store.document_paths(book_id)))
                self.books_list.addItem(item)
                if book_id == selected:
                    self.books_list.setCurrentItem(item)
            if self.books_list.currentRow() < 0 and self.books_list.count() > 0:
                self.books_list.setCurrentRow(0)
        finally:
            self.books_list.blockSignals(False)
        # Refresh archived view independently of the active selection.
        archived_selected = self.selected_archived_book()
        self.archive_list.clear()
        from PySide6.QtWidgets import QListWidgetItem
        for book_id, title, path in self.store.archived_books():
            item = QListWidgetItem(
                f"{title}    —    {format_duration(totals.get(book_id, 0))}"
            )
            item.setData(Qt.ItemDataRole.UserRole, book_id)
            item.setToolTip("\\n".join(self.store.document_paths(book_id)))
            self.archive_list.addItem(item)
            if book_id == archived_selected:
                self.archive_list.setCurrentItem(item)
        self.book_tabs.setTabText(0, f"Active ({self.books_list.count()})")
        self.book_tabs.setTabText(1, f"Archive ({self.archive_list.count()})")

    def selected_archived_book(self):
        if not hasattr(self, "archive_list"):
            return None
        item = self.archive_list.currentItem()
        return item.data(Qt.ItemDataRole.UserRole) if item else None

    def restore_archived_book(self):
        book_id = self.selected_archived_book()
        if book_id is None:
            QMessageBox.information(
                self, "Select a book", "Select an archived book to restore."
            )
            return
        try:
            self.store.restore_book(book_id)
        except ValueError as exc:
            QMessageBox.warning(self, "Unable to restore book", str(exc))
            return
        self.book_tabs.setCurrentIndex(0)
        self.refresh_books(select_book_id=book_id)
        self.refresh_history()

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

    def on_book_selected(self):
        # A row selected for Book A must never remain selected when viewing B.
        self.table.clearSelection()
        self.refresh_history()

    def refresh_history(self):
        selected_book = self.selected_book()
        rows = self.store.history(selected_book)
        book_title = next(
            (title for book_id, title, _ in self.store.books()
             if book_id == selected_book), "All books"
        )
        self.history_heading.setText(
            f"Work history — {book_title} ({len(rows)} entries)"
        )
        selected_session = self.selected_session()
        self.table.clearSelection()
        self.table.setRowCount(len(rows))
        for i, (session_id, title, start, end, seconds, manual, note, edited) in enumerate(rows):
            kind = "Manual" if manual else "Automatic"
            if edited:
                kind += " (edited)"
            values = (
                title, start.replace("T", " ")[:19], end.replace("T", " ")[:19],
                format_duration(seconds), kind, note
            )
            for j, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if j == 0:
                    item.setData(Qt.ItemDataRole.UserRole, session_id)
                self.table.setItem(i, j, item)
            if selected_session == session_id:
                self.table.selectRow(i)
        self.refresh_books()
        self.update_earnings()

    def selected_session(self):
        selected = self.table.selectedItems()
        if not selected:
            return None
        row = selected[0].row()
        item = self.table.item(row, 0)
        return item.data(Qt.ItemDataRole.UserRole) if item else None

    def add_manual_entry(self):
        dialog = ManualSessionDialog(self.store.books(), self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.store.add_manual(*dialog.values())
            self.refresh_history()

    def edit_time_entry(self):
        session_id = self.selected_session()
        record = self.store.get_session(session_id) if session_id is not None else None
        if record is None:
            QMessageBox.information(
                self, "Select an entry", "Select a history row to edit."
            )
            return
        if self.segment_id == session_id:
            self.user_pause()  # finalizes elapsed time before making corrections
            record = self.store.get_session(session_id)
        origin = "Manual" if record[5] else "Automatic"
        dialog = ManualSessionDialog(
            self.store.books(), self, record[:5], source=origin
        )
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.store.edit_session(session_id, *dialog.values())
            self.refresh_history()

    def remove_time_entry(self):
        session_id = self.selected_session()
        record = self.store.get_session(session_id) if session_id is not None else None
        if record is None:
            QMessageBox.information(
                self, "Select an entry", "Select a history row to delete."
            )
            return
        origin = "manual" if record[5] else "automatic"
        answer = QMessageBox.question(
            self, "Delete time entry",
            f"Delete this {origin} time entry permanently?\\n"
            "This will also update the book total and hourly earnings."
        )
        if answer == QMessageBox.StandardButton.Yes:
            if self.segment_id == session_id:
                self.user_pause()
            self.store.remove_session(session_id)
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
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        try:
            book_id, restored = self.store.add_book(*dialog.values())
        except (ValueError, sqlite3.IntegrityError) as exc:
            QMessageBox.warning(self, "Book could not be added", str(exc))
            return
        self.refresh_books(select_book_id=book_id)
        self.refresh_history()
        if restored:
            QMessageBox.information(
                self, "Book restored",
                "This book was previously removed from the list. It has been "
                "restored with its original time history and payment details."
            )

    def edit_book(self):
        book_id = self.selected_book()
        if book_id is None:
            return
        book = next((b for b in self.store.books() if b[0] == book_id), None)
        if not book:
            return
        dialog = BookDialog(self, book[1], self.store.document_paths(book_id))
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        try:
            self.store.edit_book(book_id, *dialog.values())
        except (ValueError, sqlite3.IntegrityError) as exc:
            QMessageBox.warning(self, "Book could not be changed", str(exc))
            return
        self.refresh_books(select_book_id=book_id)
        self.refresh_history()

    def remove_book(self):
        book_id = self.selected_book()
        if book_id is None:
            return
        reply = QMessageBox.question(
            self, "Remove book", "Move this book to the Archive? Its time history will be preserved."
        )
        if reply == QMessageBox.StandardButton.Yes:
            if self.current_book == book_id:
                self.user_pause()
                self.current_book = None
            self.store.archive_book(book_id)
            self.refresh_books()
            self.refresh_history()

    def set_auto(self, enabled):
        self.auto_enabled = enabled
        if not enabled and self.segment_id is not None:
            self.user_pause()
        self.status.setText(
            "Automatic tracking enabled" if enabled else "Manual mode — auto-start disabled"
        )

    def find_matching_book(self, process, title):
        matched = []
        for book_id, _, _ in self.store.books():
            if any(
                title_matches_document(title, path, process)
                for path in self.store.document_paths(book_id)
            ):
                matched.append(book_id)
        # Never bill an ambiguous document title to an arbitrary book.
        return matched[0] if len(matched) == 1 else None

    def begin(self, book_id):
        if self.segment_id is not None and self.current_book == book_id:
            return
        self.pause()
        self.current_book = book_id
        self.book_override = None
        self.override_origin_hwnd = None
        self.override_seen_other_window = False
        self.resume_guard = None
        self.pause_needs_leave = False
        self.idle_prompt_suppressed = False
        self.segment_id = self.store.start_segment(book_id)
        self.segment_started = time.monotonic()
        self.last_idle_alert = False
        self.last_nonword_alert = False
        self.continue_outside_word = False
        self.status.setText("Tracking writing time")
        # Follow automatic Word/PDF book switches in the sidebar and history.
        self.refresh_books(select_book_id=book_id)
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
        self.book_override = None
        self.pause()
        self.resume_guard = last_input_tick()
        # Don't instantly resume simply because clicking Pause returns focus to
        # a document. Require a transition away and then back to that document.
        self.pause_needs_leave = True

    def stop(self):
        self.book_override = None
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
        if self.segment_id is None:
            return
        origin_hwnd = int(user32.GetForegroundWindow()) if user32 else None
        dialog = InterruptionDialog(
            self.store.books(), self.current_book, reason
        )
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()
        dialog.exec()
        action = dialog.action
        if action == "pause":
            self.user_pause()
        elif action == "stop":
            self.stop()
        elif action == "disable":
            self.user_pause()
            self.auto_box.setChecked(False)
        elif action == "keep":
            selected = dialog.chosen_book()
            if selected not in {book_id for book_id, _, _ in self.store.books()}:
                self.user_pause()
                return
            if selected != self.current_book:
                self.begin(selected)
            self.book_override = selected
            self.override_origin_hwnd = origin_hwnd
            self.override_seen_other_window = False
            self.continue_outside_word = True
            self.idle_prompt_suppressed = True
            self.status.setText(
                "Tracking chosen book: " + next(
                    name for ident, name, _ in self.store.books()
                    if ident == selected
                )
            )
        else:
            # Close/X is not an affirmative choice to bill idle time.
            self.user_pause()


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
        hwnd = int(user32.GetForegroundWindow()) if user32 else None
        # A user-picked Keep target wins over the document that happened
        # to be in front when the interruption appeared. The override ends
        # only when the user switches windows and returns to a tracked file.
        book_id, self.book_override, self.override_seen_other_window = (
            resolve_book_override(
                book_id, self.book_override, self.override_origin_hwnd,
                hwnd, self.override_seen_other_window,
            )
        )
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
            self.continue_outside_word = False

        if self.auto_enabled and not idle and book_id is not None:
            # A *new* user input while the registered document is foreground
            # resumes a paused session. This covers pauses made from an idle
            # alert while Word remains in front, not just switching windows.
            fresh_input = (
                self.resume_guard is not None
                and last_input_tick() != self.resume_guard
            )
            if not self.pause_needs_leave or fresh_input:
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
