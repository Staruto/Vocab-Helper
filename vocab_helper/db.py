from __future__ import annotations

import random
import sqlite3
from pathlib import Path
from typing import Iterable

from .models import VocabEntry
from .validators import normalize_optional_text, validate_vocab_fields


class VocabRepository:
    def __init__(self, db_path: Path | str) -> None:
        self.db_path = Path(db_path)

    def initialize(self) -> None:
        self.db_path.parent.mkdir(parents=True, exist_ok=True)

        connection = sqlite3.connect(self.db_path)
        try:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS vocab_entries (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    japanese_text TEXT NOT NULL CHECK (trim(japanese_text) <> ''),
                    kana_text TEXT NULL,
                    english_text TEXT NOT NULL CHECK (trim(english_text) <> ''),
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS vocab_stats (
                    entry_id INTEGER PRIMARY KEY,
                    error_count INTEGER NOT NULL DEFAULT 0,
                    test_count INTEGER NOT NULL DEFAULT 0,
                    last_tested TEXT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (entry_id) REFERENCES vocab_entries(id) ON DELETE CASCADE
                )
                """
            )
            connection.commit()
        finally:
            connection.close()

    def list_entries(self) -> list[VocabEntry]:
        connection = sqlite3.connect(self.db_path)
        try:
            connection.row_factory = sqlite3.Row
            rows: Iterable[sqlite3.Row] = connection.execute(
                """
                SELECT id, japanese_text, kana_text, english_text, created_at
                FROM vocab_entries
                ORDER BY id ASC
                """
            )
            return [self._map_row(row) for row in rows]
        finally:
            connection.close()

    def count_entries(self) -> int:
        connection = sqlite3.connect(self.db_path)
        try:
            row = connection.execute(
                """
                SELECT COUNT(*)
                FROM vocab_entries
                """
            ).fetchone()
            if row is None:
                return 0
            return int(row[0])
        finally:
            connection.close()

    def count_distinct_english_meanings(self) -> int:
        connection = sqlite3.connect(self.db_path)
        try:
            row = connection.execute(
                """
                SELECT COUNT(*)
                FROM (
                    SELECT DISTINCT LOWER(TRIM(english_text)) AS normalized_english
                    FROM vocab_entries
                    WHERE TRIM(english_text) <> ''
                )
                """
            ).fetchone()
            if row is None:
                return 0
            return int(row[0])
        finally:
            connection.close()

    def get_random_entries(self, count: int) -> list[VocabEntry]:
        requested = max(int(count), 0)
        if requested == 0:
            return []

        connection = sqlite3.connect(self.db_path)
        try:
            connection.row_factory = sqlite3.Row
            rows: Iterable[sqlite3.Row] = connection.execute(
                """
                SELECT id, japanese_text, kana_text, english_text, created_at
                FROM vocab_entries
                ORDER BY RANDOM()
                LIMIT ?
                """,
                (requested,),
            )
            return [self._map_row(row) for row in rows]
        finally:
            connection.close()

    def list_entries_with_stats(
        self,
        sort_mode: str = "time",
        time_order: str = "newest",
    ) -> list[tuple[VocabEntry, int, int, str]]:
        order_by = "ORDER BY e.id DESC"
        if sort_mode == "stats":
            order_by = (
                "ORDER BY "
                "CASE "
                "WHEN COALESCE(s.test_count, 0) = 0 THEN 0 "
                "WHEN COALESCE(s.error_count, 0) >= 3 THEN 1 "
                "WHEN COALESCE(s.error_count, 0) >= 1 THEN 2 "
                "ELSE 3 END ASC, "
                "COALESCE(s.error_count, 0) DESC, "
                "e.id DESC"
            )
        elif time_order == "oldest":
            order_by = "ORDER BY e.id ASC"

        connection = sqlite3.connect(self.db_path)
        try:
            connection.row_factory = sqlite3.Row
            rows: Iterable[sqlite3.Row] = connection.execute(
                f"""
                SELECT
                    e.id,
                    e.japanese_text,
                    e.kana_text,
                    e.english_text,
                    e.created_at,
                    COALESCE(s.error_count, 0) AS error_count,
                    COALESCE(s.test_count, 0) AS test_count
                FROM vocab_entries AS e
                LEFT JOIN vocab_stats AS s
                    ON s.entry_id = e.id
                {order_by}
                """
            )

            result: list[tuple[VocabEntry, int, int, str]] = []
            for row in rows:
                entry = self._map_row(row)
                error_count = int(row["error_count"])
                test_count = int(row["test_count"])
                tier = self._tier_from_counts(test_count, error_count)
                result.append((entry, test_count, error_count, tier))
            return result
        finally:
            connection.close()

    def get_test_entries_by_preference(self, count: int, strategy: str = "strict") -> list[VocabEntry]:
        requested = max(int(count), 0)
        if requested == 0:
            return []

        entries_with_stats = self.list_entries_with_stats(sort_mode="time", time_order="newest")
        if not entries_with_stats:
            return []

        if strategy == "weighted":
            weighted_pool = list(entries_with_stats)
            weights_by_tier = {"gray": 4, "red": 3, "yellow": 2, "green": 1}
            ordered: list[tuple[VocabEntry, int, int, str]] = []
            while weighted_pool and len(ordered) < requested:
                weights = [weights_by_tier[item[3]] for item in weighted_pool]
                selected_index = random.choices(range(len(weighted_pool)), weights=weights, k=1)[0]
                ordered.append(weighted_pool.pop(selected_index))
            return [entry for entry, _, _, _ in ordered]

        buckets: dict[str, list[tuple[VocabEntry, int, int, str]]] = {
            "gray": [],
            "red": [],
            "yellow": [],
            "green": [],
        }
        for item in entries_with_stats:
            buckets[item[3]].append(item)

        ordered_strict: list[tuple[VocabEntry, int, int, str]] = []
        for tier in ("gray", "red", "yellow", "green"):
            random.shuffle(buckets[tier])
            ordered_strict.extend(buckets[tier])

        sliced = ordered_strict[:requested]
        return [entry for entry, _, _, _ in sliced]

    def get_english_options_for_entry(self, entry_id: int, max_options: int = 4) -> list[str]:
        max_count = max(int(max_options), 2)

        connection = sqlite3.connect(self.db_path)
        try:
            row = connection.execute(
                """
                SELECT english_text
                FROM vocab_entries
                WHERE id = ?
                """,
                (entry_id,),
            ).fetchone()
            if row is None:
                raise LookupError(f"Vocabulary entry with id {entry_id} was not found.")

            correct_english = str(row[0]).strip()

            distractor_rows = connection.execute(
                """
                SELECT DISTINCT TRIM(english_text) AS english_text
                FROM vocab_entries
                WHERE id <> ?
                  AND TRIM(english_text) <> ''
                  AND LOWER(TRIM(english_text)) <> LOWER(TRIM(?))
                ORDER BY RANDOM()
                LIMIT ?
                """,
                (entry_id, correct_english, max_count - 1),
            ).fetchall()

            options = [correct_english]
            options.extend(str(distractor_row[0]) for distractor_row in distractor_rows)
            random.shuffle(options)
            return options
        finally:
            connection.close()

    def record_test_result(self, entry_id: int, is_correct: bool) -> None:
        connection = sqlite3.connect(self.db_path)
        try:
            entry_exists = connection.execute(
                """
                SELECT 1
                FROM vocab_entries
                WHERE id = ?
                """,
                (entry_id,),
            ).fetchone()
            if entry_exists is None:
                raise LookupError(f"Vocabulary entry with id {entry_id} was not found.")

            connection.execute(
                """
                INSERT INTO vocab_stats (entry_id, error_count, test_count)
                VALUES (?, 0, 0)
                ON CONFLICT(entry_id) DO NOTHING
                """,
                (entry_id,),
            )

            connection.execute(
                """
                UPDATE vocab_stats
                SET
                    test_count = test_count + 1,
                    error_count = error_count + ?,
                    last_tested = CURRENT_TIMESTAMP
                WHERE entry_id = ?
                """,
                (0 if is_correct else 1, entry_id),
            )
            connection.commit()
        finally:
            connection.close()

    def increase_priority(self, entry_id: int) -> str:
        test_count, error_count = self._get_existing_test_stats(entry_id)

        if error_count == 0:
            new_error_count = 1
        elif error_count <= 2:
            new_error_count = 3
        else:
            new_error_count = error_count

        self._set_error_count(entry_id, new_error_count)
        return self._tier_from_counts(test_count, new_error_count)

    def decrease_priority(self, entry_id: int) -> str:
        test_count, error_count = self._get_existing_test_stats(entry_id)

        if error_count >= 3:
            new_error_count = 2
        elif error_count in (1, 2):
            new_error_count = 0
        else:
            new_error_count = error_count

        self._set_error_count(entry_id, new_error_count)
        return self._tier_from_counts(test_count, new_error_count)

    def get_entry_stats(self, entry_id: int) -> tuple[int, int, str]:
        connection = sqlite3.connect(self.db_path)
        try:
            entry_exists = connection.execute(
                """
                SELECT 1
                FROM vocab_entries
                WHERE id = ?
                """,
                (entry_id,),
            ).fetchone()
            if entry_exists is None:
                raise LookupError(f"Vocabulary entry with id {entry_id} was not found.")

            row = connection.execute(
                """
                SELECT test_count, error_count
                FROM vocab_stats
                WHERE entry_id = ?
                """,
                (entry_id,),
            ).fetchone()
            if row is None:
                return 0, 0, "gray"

            test_count = int(row[0])
            error_count = int(row[1])
            return test_count, error_count, self._tier_from_counts(test_count, error_count)
        finally:
            connection.close()

    def add_entry(self, japanese_text: str, kana_text: str, english_text: str) -> VocabEntry:
        created = self.add_entries([(japanese_text, kana_text, english_text)])
        if not created:
            raise RuntimeError("Could not load inserted entry.")
        return created[0]

    def add_entries(self, entries: Iterable[tuple[str, str, str]]) -> list[VocabEntry]:
        normalized_entries: list[tuple[str, str | None, str]] = []
        for japanese_text, kana_text, english_text in entries:
            japanese, english = validate_vocab_fields(japanese_text, english_text)
            kana = normalize_optional_text(kana_text)
            normalized_entries.append((japanese, kana, english))

        if not normalized_entries:
            return []

        connection = sqlite3.connect(self.db_path)
        try:
            connection.row_factory = sqlite3.Row

            inserted_ids: list[int] = []
            for japanese, kana, english in normalized_entries:
                cursor = connection.execute(
                    """
                    INSERT INTO vocab_entries (japanese_text, kana_text, english_text)
                    VALUES (?, ?, ?)
                    """,
                    (japanese, kana, english),
                )
                inserted_ids.append(int(cursor.lastrowid))

            placeholders = ",".join("?" for _ in inserted_ids)
            rows = connection.execute(
                f"""
                SELECT id, japanese_text, kana_text, english_text, created_at
                FROM vocab_entries
                WHERE id IN ({placeholders})
                ORDER BY id ASC
                """,
                tuple(inserted_ids),
            ).fetchall()
            connection.commit()
        finally:
            connection.close()

        if len(rows) != len(inserted_ids):
            raise RuntimeError("Could not load one or more inserted entries.")

        return [self._map_row(row) for row in rows]

    def get_entry(self, entry_id: int) -> VocabEntry:
        connection = sqlite3.connect(self.db_path)
        row: sqlite3.Row | None = None
        try:
            connection.row_factory = sqlite3.Row
            row = connection.execute(
                """
                SELECT id, japanese_text, kana_text, english_text, created_at
                FROM vocab_entries
                WHERE id = ?
                """,
                (entry_id,),
            ).fetchone()
        finally:
            connection.close()

        if row is None:
            raise LookupError(f"Vocabulary entry with id {entry_id} was not found.")

        return self._map_row(row)

    def update_entry(self, entry_id: int, japanese_text: str, kana_text: str, english_text: str) -> VocabEntry:
        japanese, english = validate_vocab_fields(japanese_text, english_text)
        kana = normalize_optional_text(kana_text)

        connection = sqlite3.connect(self.db_path)
        row: sqlite3.Row | None = None
        try:
            connection.row_factory = sqlite3.Row
            cursor = connection.execute(
                """
                UPDATE vocab_entries
                SET japanese_text = ?, kana_text = ?, english_text = ?
                WHERE id = ?
                """,
                (japanese, kana, english, entry_id),
            )
            if cursor.rowcount == 0:
                raise LookupError(f"Vocabulary entry with id {entry_id} was not found.")

            row = connection.execute(
                """
                SELECT id, japanese_text, kana_text, english_text, created_at
                FROM vocab_entries
                WHERE id = ?
                """,
                (entry_id,),
            ).fetchone()
            connection.commit()
        finally:
            connection.close()

        if row is None:
            raise RuntimeError("Could not load updated entry.")

        return self._map_row(row)

    def delete_entry(self, entry_id: int) -> None:
        self.delete_entries([entry_id])

    def delete_entries(self, entry_ids: Iterable[int]) -> int:
        unique_ids = sorted({int(entry_id) for entry_id in entry_ids})
        if not unique_ids:
            return 0

        placeholders = ",".join("?" for _ in unique_ids)
        connection = sqlite3.connect(self.db_path)
        try:
            existing_rows = connection.execute(
                f"""
                SELECT id
                FROM vocab_entries
                WHERE id IN ({placeholders})
                """,
                tuple(unique_ids),
            ).fetchall()
            existing_ids = {int(row[0]) for row in existing_rows}
            missing_ids = [entry_id for entry_id in unique_ids if entry_id not in existing_ids]
            if missing_ids:
                missing_text = ", ".join(str(entry_id) for entry_id in missing_ids)
                raise LookupError(f"Vocabulary entry ids not found: {missing_text}")

            cursor = connection.execute(
                f"""
                DELETE FROM vocab_entries
                WHERE id IN ({placeholders})
                """,
                tuple(unique_ids),
            )
            connection.execute(
                f"""
                DELETE FROM vocab_stats
                WHERE entry_id IN ({placeholders})
                """,
                tuple(unique_ids),
            )
            deleted_count = int(cursor.rowcount)
            connection.commit()
            return deleted_count
        finally:
            connection.close()

    @staticmethod
    def _tier_from_counts(test_count: int, error_count: int) -> str:
        if test_count <= 0:
            return "gray"
        if error_count <= 0:
            return "green"
        if error_count <= 2:
            return "yellow"
        return "red"

    def _get_existing_test_stats(self, entry_id: int) -> tuple[int, int]:
        connection = sqlite3.connect(self.db_path)
        try:
            entry_exists = connection.execute(
                """
                SELECT 1
                FROM vocab_entries
                WHERE id = ?
                """,
                (entry_id,),
            ).fetchone()
            if entry_exists is None:
                raise LookupError(f"Vocabulary entry with id {entry_id} was not found.")

            row = connection.execute(
                """
                SELECT test_count, error_count
                FROM vocab_stats
                WHERE entry_id = ?
                """,
                (entry_id,),
            ).fetchone()
            if row is None or int(row[0]) <= 0:
                raise ValueError("Manual priority change is not supported for gray tier entries.")
            return int(row[0]), int(row[1])
        finally:
            connection.close()

    def _set_error_count(self, entry_id: int, error_count: int) -> None:
        connection = sqlite3.connect(self.db_path)
        try:
            connection.execute(
                """
                UPDATE vocab_stats
                SET error_count = ?
                WHERE entry_id = ?
                """,
                (max(error_count, 0), entry_id),
            )
            connection.commit()
        finally:
            connection.close()

    @staticmethod
    def _map_row(row: sqlite3.Row) -> VocabEntry:
        return VocabEntry(
            id=int(row["id"]),
            japanese_text=str(row["japanese_text"]),
            kana_text=row["kana_text"],
            english_text=str(row["english_text"]),
            created_at=str(row["created_at"]),
        )


def default_db_path() -> Path:
    return Path(__file__).resolve().parent.parent / "vocab.db"
