"""Regression tests for book restoration, history filtering and editing."""
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path
from unittest.mock import patch

import book_work_clock as app


class StoreRegressionTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.addCleanup(self.tempdir.cleanup)
        folder = Path(self.tempdir.name)
        self.patch_dir = patch.object(app, "APP_DIR", folder)
        self.patch_db = patch.object(app, "DB_FILE", folder / "test.sqlite3")
        self.patch_dir.start()
        self.patch_db.start()
        self.addCleanup(self.patch_dir.stop)
        self.addCleanup(self.patch_db.stop)
        self.store = app.Store()
        self.addCleanup(self.store.db.close)
        self.start = datetime.fromisoformat("2026-11-14T10:00:00+02:00")
        self.end = self.start + timedelta(minutes=30)

    def test_archived_book_restores_same_id_and_history(self):
        book_id, restored = self.store.add_book(
            "Novel One", [r"C:\Books\Novel.docx", r"C:\Books\Novel.pdf"]
        )
        self.assertFalse(restored)
        self.store.add_manual(book_id, self.start, self.end, "Original work")
        self.store.archive_book(book_id)
        self.assertFalse(self.store.books())
        result_id, was_restored = self.store.add_book(
            "Novel One", [r"C:\Books\Novel.docx"]
        )
        self.assertTrue(was_restored)
        self.assertEqual(result_id, book_id)
        self.assertEqual(len(self.store.books()), 1)
        self.assertEqual(len(self.store.history(book_id)), 1)
        self.assertEqual(self.store.totals()[book_id], 1800)

    def test_duplicate_active_book_is_not_created(self):
        self.store.add_book("Project A", [r"C:\Books\A.docx"])
        with self.assertRaisesRegex(ValueError, "already exists"):
            self.store.add_book("Project A", [r"C:\Books\A.docx"])
        self.assertEqual(len(self.store.books()), 1)

    def test_restore_by_document_even_if_book_title_changed(self):
        book_id, _ = self.store.add_book(
            "Older Name", [r"C:\Books\MyFile.docx"]
        )
        self.store.archive_book(book_id)
        restored_id, restored = self.store.add_book(
            "New Name", [r"C:\Books\MyFile.docx"]
        )
        self.assertTrue(restored)
        self.assertEqual(restored_id, book_id)

    def test_edit_and_delete_automatic_capture(self):
        book_id, _ = self.store.add_book("Project A", [r"C:\Books\A.docx"])
        session_id = self.store.start_segment(book_id)
        self.store.save_segment(session_id, 40.0)
        original = self.store.get_session(session_id)
        self.assertEqual(original[5], 0)
        end = self.start + timedelta(hours=2)
        self.store.edit_session(session_id, book_id, self.start, end, "Corrected")
        record = self.store.get_session(session_id)
        self.assertEqual(record[4], "Corrected")
        self.assertEqual(record[5], 0)
        self.assertEqual(int(self.store.totals()[book_id]), 7200)
        entry = self.store.history(book_id)[0]
        self.assertEqual(entry[5], 0)
        self.assertTrue(entry[7])  # edited_at
        self.assertTrue(self.store.remove_session(session_id))
        self.assertEqual(self.store.history(book_id), [])

    def test_history_filters_to_selected_book(self):
        first, _ = self.store.add_book("One", [r"C:\Books\One.docx"])
        second, _ = self.store.add_book("Two", [r"C:\Books\Two.pdf"])
        self.store.add_manual(first, self.start, self.end, "One")
        self.store.add_manual(second, self.start, self.end, "Two")
        self.assertEqual(len(self.store.history(first)), 1)
        self.assertEqual(len(self.store.history(second)), 1)
        self.assertEqual(len(self.store.history()), 2)

    def test_archive_view_and_restore_preserve_total_and_payment(self):
        book_id, _ = self.store.add_book(
            "Second Book", [r"C:\\Books\\Second.docx"]
        )
        self.store.update_payment(book_id, 4600, "ILS")
        self.store.add_manual(book_id, self.start, self.end, "Review")
        self.store.archive_book(book_id)
        self.assertEqual(
            [b[0] for b in self.store.archived_books()], [book_id]
        )
        self.assertEqual(self.store.books(), [])
        self.store.restore_book(book_id)
        self.assertFalse(self.store.archived_books())
        self.assertEqual(self.store.books()[0][0], book_id)
        self.assertEqual(len(self.store.history(book_id)), 1)
        self.assertEqual(self.store.book_payment(book_id), (4600.0, "ILS"))

    def test_archive_restore_conflict_does_not_merge_books(self):
        book_id, _ = self.store.add_book(
            "First", [r"C:\\Books\\Shared.docx"]
        )
        self.store.archive_book(book_id)
        other_id, _ = self.store.add_book(
            "Second", [r"C:\\Books\\Other.docx"]
        )
        self.store.edit_book(other_id, "Second", [r"C:\\Books\\Shared.docx"])
        with self.assertRaisesRegex(ValueError, "belongs to an active book"):
            self.store.restore_book(book_id)
        self.assertEqual([b[0] for b in self.store.archived_books()], [book_id])

    def test_session_edit_can_move_hours_between_books(self):
        first, _ = self.store.add_book("One", [r"C:\Books\One.docx"])
        second, _ = self.store.add_book("Two", [r"C:\Books\Two.docx"])
        segment_id = self.store.start_segment(first)
        self.store.save_segment(segment_id, 50)
        self.store.edit_session(segment_id, second, self.start, self.end, "Moved")
        self.assertEqual(self.store.totals().get(first, 0), 0)
        self.assertEqual(self.store.totals()[second], 1800)


if __name__ == "__main__":
    unittest.main()
