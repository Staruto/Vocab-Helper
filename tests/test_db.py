import tempfile
import unittest
from pathlib import Path

from vocab_helper.db import VocabRepository
from vocab_helper.validators import ValidationError


class VocabRepositoryTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.temp_dir.name) / "test_vocab.db"
        self.repository = VocabRepository(self.db_path)
        self.repository.initialize()

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_add_and_list_entry(self) -> None:
        self.repository.add_entry("食べる", "たべる", "to eat")

        rows = self.repository.list_entries()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].japanese_text, "食べる")
        self.assertEqual(rows[0].kana_text, "たべる")
        self.assertEqual(rows[0].english_text, "to eat")

    def test_optional_kana_saved_as_none(self) -> None:
        self.repository.add_entry("行く", "   ", "to go")

        rows = self.repository.list_entries()
        self.assertEqual(len(rows), 1)
        self.assertIsNone(rows[0].kana_text)

    def test_required_fields_raise_validation_error(self) -> None:
        with self.assertRaises(ValidationError):
            self.repository.add_entry("", "かな", "to do")

        with self.assertRaises(ValidationError):
            self.repository.add_entry("する", "かな", "")

    def test_get_entry_returns_inserted_row(self) -> None:
        inserted = self.repository.add_entry("書く", "かく", "to write")

        loaded = self.repository.get_entry(inserted.id)
        self.assertEqual(loaded.id, inserted.id)
        self.assertEqual(loaded.japanese_text, "書く")
        self.assertEqual(loaded.kana_text, "かく")
        self.assertEqual(loaded.english_text, "to write")

    def test_update_entry_changes_values(self) -> None:
        inserted = self.repository.add_entry("読む", "よむ", "to read")

        updated = self.repository.update_entry(inserted.id, "飲む", "のむ", "to drink")
        self.assertEqual(updated.id, inserted.id)
        self.assertEqual(updated.japanese_text, "飲む")
        self.assertEqual(updated.kana_text, "のむ")
        self.assertEqual(updated.english_text, "to drink")

        loaded = self.repository.get_entry(inserted.id)
        self.assertEqual(loaded.japanese_text, "飲む")

    def test_delete_entry_removes_row(self) -> None:
        inserted = self.repository.add_entry("行く", "いく", "to go")

        self.repository.delete_entry(inserted.id)
        self.assertEqual(self.repository.list_entries(), [])

    def test_add_entries_inserts_multiple_rows(self) -> None:
        created = self.repository.add_entries(
            [
                ("食べる", "たべる", "to eat"),
                ("行く", "", "to go"),
                ("飲む", "のむ", "to drink"),
            ]
        )

        self.assertEqual(len(created), 3)
        rows = self.repository.list_entries()
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[0].japanese_text, "食べる")
        self.assertIsNone(rows[1].kana_text)
        self.assertEqual(rows[2].english_text, "to drink")

    def test_add_entries_is_atomic_when_any_row_invalid(self) -> None:
        with self.assertRaises(ValidationError):
            self.repository.add_entries(
                [
                    ("書く", "かく", "to write"),
                    ("見る", "みる", "   "),
                ]
            )

        self.assertEqual(self.repository.list_entries(), [])

    def test_delete_entries_removes_multiple_rows(self) -> None:
        created = self.repository.add_entries(
            [
                ("赤", "あか", "red"),
                ("青", "あお", "blue"),
                ("白", "しろ", "white"),
            ]
        )

        deleted = self.repository.delete_entries([created[0].id, created[2].id, created[2].id])
        self.assertEqual(deleted, 2)

        rows = self.repository.list_entries()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].japanese_text, "青")

    def test_delete_entries_raises_for_missing_ids_without_partial_delete(self) -> None:
        created = self.repository.add_entries(
            [
                ("山", "やま", "mountain"),
                ("川", "かわ", "river"),
            ]
        )

        with self.assertRaises(LookupError):
            self.repository.delete_entries([created[0].id, 999999])

        rows = self.repository.list_entries()
        self.assertEqual(len(rows), 2)

    def test_missing_entry_operations_raise_lookup_error(self) -> None:
        with self.assertRaises(LookupError):
            self.repository.get_entry(9999)

        with self.assertRaises(LookupError):
            self.repository.update_entry(9999, "食べる", "たべる", "to eat")

        with self.assertRaises(LookupError):
            self.repository.delete_entry(9999)


if __name__ == "__main__":
    unittest.main()
