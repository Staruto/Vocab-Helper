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

    def test_count_entries_reflects_total_rows(self) -> None:
        self.assertEqual(self.repository.count_entries(), 0)

        self.repository.add_entry("食べる", "たべる", "to eat")
        self.repository.add_entry("行く", "いく", "to go")

        self.assertEqual(self.repository.count_entries(), 2)

    def test_get_random_entries_respects_requested_count_bounds(self) -> None:
        created = self.repository.add_entries(
            [
                ("赤", "あか", "red"),
                ("青", "あお", "blue"),
                ("白", "しろ", "white"),
            ]
        )
        created_ids = {entry.id for entry in created}

        picked_two = self.repository.get_random_entries(2)
        self.assertEqual(len(picked_two), 2)
        self.assertTrue(all(entry.id in created_ids for entry in picked_two))

        picked_many = self.repository.get_random_entries(99)
        self.assertEqual(len(picked_many), 3)
        self.assertTrue(all(entry.id in created_ids for entry in picked_many))

        picked_zero = self.repository.get_random_entries(0)
        self.assertEqual(picked_zero, [])

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

    def test_record_test_result_updates_stats_and_tiers(self) -> None:
        entry = self.repository.add_entry("書く", "かく", "to write")

        self.assertEqual(self.repository.get_entry_stats(entry.id), (0, 0, "gray"))

        self.repository.record_test_result(entry.id, is_correct=True)
        self.assertEqual(self.repository.get_entry_stats(entry.id), (1, 0, "green"))

        self.repository.record_test_result(entry.id, is_correct=False)
        self.assertEqual(self.repository.get_entry_stats(entry.id), (2, 1, "yellow"))

        self.repository.record_test_result(entry.id, is_correct=False)
        self.assertEqual(self.repository.get_entry_stats(entry.id), (3, 2, "yellow"))

        self.repository.record_test_result(entry.id, is_correct=False)
        self.assertEqual(self.repository.get_entry_stats(entry.id), (4, 3, "red"))

    def test_list_entries_with_stats_supports_time_and_stats_sort(self) -> None:
        created = self.repository.add_entries(
            [
                ("灰", "はい", "gray"),
                ("緑", "みどり", "green"),
                ("黄", "き", "yellow"),
                ("赤", "あか", "red"),
            ]
        )
        gray_id, green_id, yellow_id, red_id = [entry.id for entry in created]

        self.repository.record_test_result(green_id, is_correct=True)
        self.repository.record_test_result(yellow_id, is_correct=False)
        for _ in range(3):
            self.repository.record_test_result(red_id, is_correct=False)

        newest_rows = self.repository.list_entries_with_stats(sort_mode="time", time_order="newest")
        newest_ids = [entry.id for entry, _, _, _ in newest_rows]
        self.assertEqual(newest_ids, [red_id, yellow_id, green_id, gray_id])

        oldest_rows = self.repository.list_entries_with_stats(sort_mode="time", time_order="oldest")
        oldest_ids = [entry.id for entry, _, _, _ in oldest_rows]
        self.assertEqual(oldest_ids, [gray_id, green_id, yellow_id, red_id])

        stats_rows = self.repository.list_entries_with_stats(sort_mode="stats", time_order="newest")
        stats_ids = [entry.id for entry, _, _, _ in stats_rows]
        self.assertEqual(stats_ids, [gray_id, red_id, yellow_id, green_id])

    def test_get_test_entries_by_preference_strict_and_weighted(self) -> None:
        created = self.repository.add_entries(
            [
                ("灰", "はい", "gray"),
                ("緑", "みどり", "green"),
                ("黄", "き", "yellow"),
                ("赤", "あか", "red"),
            ]
        )
        gray_id, green_id, yellow_id, red_id = [entry.id for entry in created]

        self.repository.record_test_result(green_id, is_correct=True)
        self.repository.record_test_result(yellow_id, is_correct=False)
        for _ in range(3):
            self.repository.record_test_result(red_id, is_correct=False)

        strict_rows = self.repository.get_test_entries_by_preference(4, strategy="strict")
        strict_ids = [entry.id for entry in strict_rows]
        self.assertEqual(strict_ids, [gray_id, red_id, yellow_id, green_id])

        weighted_rows = self.repository.get_test_entries_by_preference(4, strategy="weighted")
        weighted_ids = [entry.id for entry in weighted_rows]
        self.assertEqual(len(weighted_ids), 4)
        self.assertEqual(len(set(weighted_ids)), 4)
        self.assertEqual(set(weighted_ids), {gray_id, green_id, yellow_id, red_id})

    def test_manual_priority_changes_and_gray_restriction(self) -> None:
        gray_entry = self.repository.add_entry("空", "そら", "sky")
        with self.assertRaises(ValueError):
            self.repository.increase_priority(gray_entry.id)

        entry = self.repository.add_entry("水", "みず", "water")
        self.repository.record_test_result(entry.id, is_correct=True)
        self.assertEqual(self.repository.get_entry_stats(entry.id), (1, 0, "green"))

        tier = self.repository.increase_priority(entry.id)
        self.assertEqual(tier, "yellow")
        self.assertEqual(self.repository.get_entry_stats(entry.id), (1, 1, "yellow"))

        tier = self.repository.increase_priority(entry.id)
        self.assertEqual(tier, "red")
        self.assertEqual(self.repository.get_entry_stats(entry.id), (1, 3, "red"))

        tier = self.repository.decrease_priority(entry.id)
        self.assertEqual(tier, "yellow")
        self.assertEqual(self.repository.get_entry_stats(entry.id), (1, 2, "yellow"))

        tier = self.repository.decrease_priority(entry.id)
        self.assertEqual(tier, "green")
        self.assertEqual(self.repository.get_entry_stats(entry.id), (1, 0, "green"))

        tier = self.repository.decrease_priority(entry.id)
        self.assertEqual(tier, "green")
        self.assertEqual(self.repository.get_entry_stats(entry.id), (1, 0, "green"))

    def test_missing_entry_operations_raise_lookup_error(self) -> None:
        with self.assertRaises(LookupError):
            self.repository.get_entry(9999)

        with self.assertRaises(LookupError):
            self.repository.update_entry(9999, "食べる", "たべる", "to eat")

        with self.assertRaises(LookupError):
            self.repository.delete_entry(9999)

        with self.assertRaises(LookupError):
            self.repository.record_test_result(9999, is_correct=False)

        with self.assertRaises(LookupError):
            self.repository.increase_priority(9999)

        with self.assertRaises(LookupError):
            self.repository.decrease_priority(9999)


if __name__ == "__main__":
    unittest.main()
