from __future__ import annotations

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

    def add_entry(self, japanese_text: str, kana_text: str, english_text: str) -> VocabEntry:
        japanese, english = validate_vocab_fields(japanese_text, english_text)
        kana = normalize_optional_text(kana_text)

        connection = sqlite3.connect(self.db_path)
        try:
            connection.row_factory = sqlite3.Row
            cursor = connection.execute(
                """
                INSERT INTO vocab_entries (japanese_text, kana_text, english_text)
                VALUES (?, ?, ?)
                """,
                (japanese, kana, english),
            )
            new_id = int(cursor.lastrowid)

            row = connection.execute(
                """
                SELECT id, japanese_text, kana_text, english_text, created_at
                FROM vocab_entries
                WHERE id = ?
                """,
                (new_id,),
            ).fetchone()
            connection.commit()
        finally:
            connection.close()

        if row is None:
            raise RuntimeError("Could not load inserted entry.")

        return self._map_row(row)

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
