"""Manual session editor for Book Work Clock."""
from datetime import datetime

from PySide6.QtWidgets import (
    QComboBox, QDateTimeEdit, QDialog, QDialogButtonBox,
    QFormLayout, QLineEdit, QMessageBox,
)
from PySide6.QtCore import QDateTime


def ensure_manual_schema(db):
    columns = {row[1] for row in db.execute("PRAGMA table_info(sessions)")}
    if "is_manual" not in columns:
        db.execute("ALTER TABLE sessions ADD COLUMN is_manual INTEGER NOT NULL DEFAULT 0")
    if "note" not in columns:
        db.execute("ALTER TABLE sessions ADD COLUMN note TEXT NOT NULL DEFAULT ''")
    if "edited_at" not in columns:
        db.execute("ALTER TABLE sessions ADD COLUMN edited_at TEXT")
    db.commit()


class ManualSessionDialog(QDialog):
    def __init__(self, books, parent=None, record=None, source="Manual"):
        super().__init__(parent)
        self.setWindowTitle(f"Edit {source.lower()} work time" if record else "Add manual work time")
        self.book = QComboBox()
        for book_id, title, path in books:
            self.book.addItem(title, book_id)
        self.start = QDateTimeEdit()
        self.end = QDateTimeEdit()
        for widget in (self.start, self.end):
            widget.setCalendarPopup(True)
            widget.setDisplayFormat("yyyy-MM-dd HH:mm:ss")
        now = QDateTime.currentDateTime()
        self.end.setDateTime(now)
        self.start.setDateTime(now.addSecs(-3600))
        self.note = QLineEdit()
        self.note.setPlaceholderText("Optional description of the work")
        if record:
            _, book_id, started, ended, note = record
            index = self.book.findData(book_id)
            if index >= 0:
                self.book.setCurrentIndex(index)
            for editor, raw in ((self.start, started), (self.end, ended)):
                dt = datetime.fromisoformat(raw)
                editor.setDateTime(QDateTime.fromString(
                    dt.strftime("%Y-%m-%d %H:%M:%S"), "yyyy-MM-dd HH:mm:ss"
                ))
            self.note.setText(note)
        layout = QFormLayout(self)
        layout.addRow("Book:", self.book)
        layout.addRow("Start:", self.start)
        layout.addRow("End:", self.end)
        layout.addRow("Notes:", self.note)
        actions = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save |
            QDialogButtonBox.StandardButton.Cancel
        )
        actions.accepted.connect(self.validate)
        actions.rejected.connect(self.reject)
        layout.addRow(actions)

    def validate(self):
        if self.book.currentIndex() < 0:
            QMessageBox.warning(self, "Missing book", "Add a book first.")
        elif self.end.dateTime() <= self.start.dateTime():
            QMessageBox.warning(self, "Invalid period", "End must be after start.")
        else:
            self.accept()

    def values(self):
        return (
            self.book.currentData(),
            self.start.dateTime().toPython().astimezone(),
            self.end.dateTime().toPython().astimezone(),
            self.note.text().strip(),
        )
