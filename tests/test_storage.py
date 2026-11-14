"""Portable checks for the SQLite migration and manual-time behaviors.

Run on Windows after installing dependencies; tests can also be used as a reference
for manual session CRUD and backward-compatible schema migration.
"""
import sqlite3
from manual_time import ensure_manual_schema


def test_existing_session_schema_migrates(tmp_path):
    db = sqlite3.connect(tmp_path / "clock.db")
    db.execute("""CREATE TABLE sessions (
        id INTEGER PRIMARY KEY, book_id INTEGER, started_at TEXT,
        ended_at TEXT, seconds REAL
    )""")
    ensure_manual_schema(db)
    columns = {r[1] for r in db.execute("PRAGMA table_info(sessions)")}
    assert "is_manual" in columns
    assert "note" in columns
