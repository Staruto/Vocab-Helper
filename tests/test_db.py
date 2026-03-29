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


if __name__ == "__main__":
    unittest.main()
