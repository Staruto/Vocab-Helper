from __future__ import annotations

from datetime import date, timedelta
import random
import sqlite3
from pathlib import Path
from typing import Iterable

from .models import VocabEntry, Workbook
from .validators import (
    ValidationError,
    normalize_optional_markdown,
    normalize_optional_text,
    validate_language_code,
    validate_vocab_fields,
)


class VocabRepository:
    ERROR_COUNT_RECOVERY_CHANCE = 0.35
    PREDEFINED_PART_OF_SPEECH_TAGS = (
        "noun",
        "verb",
        "adjective",
        "adverb",
        "expression",
        "particle",
        "auxiliary",
        "other",
    )
    PREDEFINED_DIFFICULTY_TAGS = ("N5", "N4", "N3", "N2", "N1")

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
                    workbook_id INTEGER NOT NULL DEFAULT 1,
                    japanese_text TEXT NOT NULL CHECK (trim(japanese_text) <> ''),
                    kana_text TEXT NULL,
                    english_text TEXT NOT NULL CHECK (trim(english_text) <> ''),
                    part_of_speech TEXT NULL,
                    details_markdown TEXT NULL,
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
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS app_settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS workbooks (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL COLLATE NOCASE,
                    target_language_code TEXT NOT NULL,
                    preset_key TEXT NOT NULL DEFAULT 'generic',
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE (name)
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS practice_daily_unique (
                    entry_id INTEGER NOT NULL,
                    practice_date TEXT NOT NULL,
                    PRIMARY KEY (entry_id, practice_date),
                    FOREIGN KEY (entry_id) REFERENCES vocab_entries(id) ON DELETE CASCADE
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS practice_daily_error_recovery (
                    entry_id INTEGER NOT NULL,
                    practice_date TEXT NOT NULL,
                    has_mistake INTEGER NOT NULL DEFAULT 0,
                    used_decrease INTEGER NOT NULL DEFAULT 0,
                    PRIMARY KEY (entry_id, practice_date),
                    FOREIGN KEY (entry_id) REFERENCES vocab_entries(id) ON DELETE CASCADE
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS tag_types (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    target_language_code TEXT NOT NULL,
                    name TEXT NOT NULL COLLATE NOCASE,
                    is_predefined INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE (target_language_code, name)
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS tags (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tag_type_id INTEGER NOT NULL,
                    name TEXT NOT NULL COLLATE NOCASE,
                    is_predefined INTEGER NOT NULL DEFAULT 0,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE (tag_type_id, name),
                    FOREIGN KEY (tag_type_id) REFERENCES tag_types(id) ON DELETE CASCADE
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS entry_tags (
                    entry_id INTEGER NOT NULL,
                    tag_id INTEGER NOT NULL,
                    added_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    PRIMARY KEY (entry_id, tag_id),
                    FOREIGN KEY (entry_id) REFERENCES vocab_entries(id) ON DELETE CASCADE,
                    FOREIGN KEY (tag_id) REFERENCES tags(id) ON DELETE CASCADE
                )
                """
            )
            connection.execute("CREATE INDEX IF NOT EXISTS idx_entry_tags_entry_id ON entry_tags(entry_id)")
            connection.execute("CREATE INDEX IF NOT EXISTS idx_entry_tags_tag_id ON entry_tags(tag_id)")
            connection.execute("CREATE INDEX IF NOT EXISTS idx_tags_tag_type_id ON tags(tag_type_id)")

            # Backward-compatible migrations for existing databases.
            self._ensure_column(connection, "vocab_entries", "workbook_id", "INTEGER NULL")
            self._ensure_column(connection, "vocab_entries", "part_of_speech", "TEXT NULL")
            self._ensure_column(connection, "vocab_entries", "details_markdown", "TEXT NULL")
            connection.execute("CREATE INDEX IF NOT EXISTS idx_vocab_entries_workbook_id ON vocab_entries(workbook_id)")

            connection.execute(
                """
                INSERT INTO app_settings (key, value)
                VALUES ('target_language', 'JP')
                ON CONFLICT(key) DO NOTHING
                """
            )
            connection.execute(
                """
                INSERT INTO app_settings (key, value)
                VALUES ('assistant_language', 'EN')
                ON CONFLICT(key) DO NOTHING
                """
            )

            default_workbook_id = self._ensure_default_workbook(connection)
            connection.execute(
                """
                UPDATE vocab_entries
                SET workbook_id = ?
                WHERE workbook_id IS NULL
                """,
                (default_workbook_id,),
            )
            connection.execute(
                """
                INSERT INTO app_settings (key, value)
                VALUES ('current_workbook_id', ?)
                ON CONFLICT(key) DO NOTHING
                """,
                (str(default_workbook_id),),
            )

            current_workbook_id = self._read_current_workbook_id_from_connection(connection, default_workbook_id)
            target_language_code = self._read_workbook_target_language_from_connection(connection, current_workbook_id)
            connection.execute(
                """
                INSERT INTO app_settings (key, value)
                VALUES ('target_language', ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (target_language_code,),
            )

            self._ensure_predefined_tags(connection, target_language_code)
            self._migrate_legacy_part_of_speech_tags(connection, target_language_code)
            connection.commit()
        finally:
            connection.close()

    @staticmethod
    def _ensure_column(connection: sqlite3.Connection, table_name: str, column_name: str, definition: str) -> None:
        rows = connection.execute(f"PRAGMA table_info({table_name})").fetchall()
        existing_columns = {str(row[1]) for row in rows}
        if column_name in existing_columns:
            return
        connection.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {definition}")

    @staticmethod
    def _read_target_language_from_connection(connection: sqlite3.Connection) -> str:
        row = connection.execute(
            """
            SELECT value
            FROM app_settings
            WHERE key = 'target_language'
            """
        ).fetchone()
        if row is None or row[0] is None:
            return "JP"
        try:
            return validate_language_code(str(row[0]), "Target language")
        except ValidationError:
            return "JP"

    def _ensure_default_workbook(self, connection: sqlite3.Connection) -> int:
        existing = connection.execute(
            """
            SELECT id
            FROM workbooks
            ORDER BY id ASC
            LIMIT 1
            """
        ).fetchone()
        if existing is not None:
            return int(existing[0])

        target_language_code = self._read_target_language_from_connection(connection)
        preset_key = "japanese" if target_language_code == "JP" else "generic"
        cursor = connection.execute(
            """
            INSERT INTO workbooks (name, target_language_code, preset_key)
            VALUES (?, ?, ?)
            """,
            ("Default", target_language_code, preset_key),
        )
        return int(cursor.lastrowid)

    def _read_current_workbook_id_from_connection(
        self,
        connection: sqlite3.Connection,
        fallback_workbook_id: int,
    ) -> int:
        row = connection.execute(
            """
            SELECT value
            FROM app_settings
            WHERE key = 'current_workbook_id'
            """
        ).fetchone()
        if row is None or row[0] is None:
            return int(fallback_workbook_id)

        try:
            workbook_id = int(str(row[0]))
        except (TypeError, ValueError):
            return int(fallback_workbook_id)

        exists = connection.execute(
            """
            SELECT 1
            FROM workbooks
            WHERE id = ?
            """,
            (workbook_id,),
        ).fetchone()
        if exists is None:
            return int(fallback_workbook_id)
        return workbook_id

    def _read_workbook_target_language_from_connection(
        self,
        connection: sqlite3.Connection,
        workbook_id: int,
    ) -> str:
        row = connection.execute(
            """
            SELECT target_language_code
            FROM workbooks
            WHERE id = ?
            """,
            (workbook_id,),
        ).fetchone()
        if row is None or row[0] is None:
            return self._read_target_language_from_connection(connection)

        try:
            return validate_language_code(str(row[0]), "Target language")
        except ValidationError:
            return self._read_target_language_from_connection(connection)

    def _read_workbook_preset_key_from_connection(self, connection: sqlite3.Connection, workbook_id: int) -> str:
        row = connection.execute(
            """
            SELECT preset_key
            FROM workbooks
            WHERE id = ?
            """,
            (workbook_id,),
        ).fetchone()
        if row is None or row[0] is None:
            return "generic"

        preset_key = str(row[0]).strip().lower()
        return preset_key or "generic"

    def _resolve_workbook_id_from_connection(
        self,
        connection: sqlite3.Connection,
        workbook_id: int | None,
    ) -> int:
        if workbook_id is not None:
            resolved_workbook_id = int(workbook_id)
            exists = connection.execute(
                """
                SELECT 1
                FROM workbooks
                WHERE id = ?
                """,
                (resolved_workbook_id,),
            ).fetchone()
            if exists is None:
                raise LookupError(f"Workbook with id {resolved_workbook_id} was not found.")
            return resolved_workbook_id

        fallback_workbook_id = self._ensure_default_workbook(connection)
        return self._read_current_workbook_id_from_connection(connection, fallback_workbook_id)

    def _resolve_workbook_id(self, workbook_id: int | None = None) -> int:
        connection = sqlite3.connect(self.db_path)
        try:
            return self._resolve_workbook_id_from_connection(connection, workbook_id)
        finally:
            connection.close()

    @staticmethod
    def _normalize_tag_name(value: str, label: str) -> str:
        normalized = value.strip()
        if not normalized:
            raise ValidationError(f"{label} cannot be empty.")
        return normalized

    def _resolve_target_language_code(self, target_language_code: str | None) -> str:
        if target_language_code is None:
            connection = sqlite3.connect(self.db_path)
            try:
                workbook_id = self._resolve_workbook_id_from_connection(connection, None)
                return self._read_workbook_target_language_from_connection(connection, workbook_id)
            finally:
                connection.close()
        return validate_language_code(target_language_code, "Target language")

    def _get_tag_type_id_by_name(
        self,
        connection: sqlite3.Connection,
        target_language_code: str,
        type_name: str,
    ) -> int | None:
        row = connection.execute(
            """
            SELECT id
            FROM tag_types
            WHERE target_language_code = ?
              AND name = ?
            """,
            (target_language_code, type_name),
        ).fetchone()
        if row is None:
            return None
        return int(row[0])

    def _get_or_create_tag_type(
        self,
        connection: sqlite3.Connection,
        target_language_code: str,
        type_name: str,
        is_predefined: bool,
    ) -> int:
        normalized_name = self._normalize_tag_name(type_name, "Tag type")
        existing = connection.execute(
            """
            SELECT id, is_predefined
            FROM tag_types
            WHERE target_language_code = ?
              AND name = ?
            """,
            (target_language_code, normalized_name),
        ).fetchone()
        if existing is not None:
            tag_type_id = int(existing[0])
            if is_predefined and int(existing[1]) == 0:
                connection.execute(
                    """
                    UPDATE tag_types
                    SET is_predefined = 1
                    WHERE id = ?
                    """,
                    (tag_type_id,),
                )
            return tag_type_id

        cursor = connection.execute(
            """
            INSERT INTO tag_types (target_language_code, name, is_predefined)
            VALUES (?, ?, ?)
            """,
            (target_language_code, normalized_name, 1 if is_predefined else 0),
        )
        return int(cursor.lastrowid)

    def _get_or_create_tag(
        self,
        connection: sqlite3.Connection,
        tag_type_id: int,
        tag_name: str,
        is_predefined: bool,
    ) -> int:
        normalized_name = self._normalize_tag_name(tag_name, "Tag")
        existing = connection.execute(
            """
            SELECT id, is_predefined
            FROM tags
            WHERE tag_type_id = ?
              AND name = ?
            """,
            (tag_type_id, normalized_name),
        ).fetchone()
        if existing is not None:
            tag_id = int(existing[0])
            if is_predefined and int(existing[1]) == 0:
                connection.execute(
                    """
                    UPDATE tags
                    SET is_predefined = 1
                    WHERE id = ?
                    """,
                    (tag_id,),
                )
            return tag_id

        cursor = connection.execute(
            """
            INSERT INTO tags (tag_type_id, name, is_predefined)
            VALUES (?, ?, ?)
            """,
            (tag_type_id, normalized_name, 1 if is_predefined else 0),
        )
        return int(cursor.lastrowid)

    def _ensure_predefined_tags(self, connection: sqlite3.Connection, target_language_code: str) -> None:
        part_of_speech_type_id = self._get_or_create_tag_type(
            connection,
            target_language_code,
            "part_of_speech",
            is_predefined=True,
        )
        for tag_name in self.PREDEFINED_PART_OF_SPEECH_TAGS:
            self._get_or_create_tag(connection, part_of_speech_type_id, tag_name, is_predefined=True)

        difficulty_type_id = self._get_or_create_tag_type(
            connection,
            target_language_code,
            "difficulty",
            is_predefined=True,
        )
        for tag_name in self.PREDEFINED_DIFFICULTY_TAGS:
            self._get_or_create_tag(connection, difficulty_type_id, tag_name, is_predefined=True)

    def _migrate_legacy_part_of_speech_tags(self, connection: sqlite3.Connection, target_language_code: str) -> None:
        part_of_speech_type_id = self._get_tag_type_id_by_name(connection, target_language_code, "part_of_speech")
        if part_of_speech_type_id is None:
            return

        rows = connection.execute(
            """
            SELECT id, part_of_speech
            FROM vocab_entries
            WHERE TRIM(COALESCE(part_of_speech, '')) <> ''
            """
        ).fetchall()

        predefined_set = {value.lower() for value in self.PREDEFINED_PART_OF_SPEECH_TAGS}
        for row in rows:
            entry_id = int(row[0])
            normalized_part_of_speech = normalize_optional_text(str(row[1]))
            if normalized_part_of_speech is None:
                continue

            is_predefined = normalized_part_of_speech.lower() in predefined_set
            tag_id = self._get_or_create_tag(
                connection,
                part_of_speech_type_id,
                normalized_part_of_speech,
                is_predefined=is_predefined,
            )
            connection.execute(
                """
                INSERT INTO entry_tags (entry_id, tag_id)
                VALUES (?, ?)
                ON CONFLICT(entry_id, tag_id) DO NOTHING
                """,
                (entry_id, tag_id),
            )

    def _sync_entry_part_of_speech_tag(
        self,
        connection: sqlite3.Connection,
        entry_id: int,
        part_of_speech: str | None,
        target_language_code: str,
    ) -> None:
        part_of_speech_type_id = self._get_or_create_tag_type(
            connection,
            target_language_code,
            "part_of_speech",
            is_predefined=True,
        )

        connection.execute(
            """
            DELETE FROM entry_tags
            WHERE entry_id = ?
              AND tag_id IN (
                  SELECT t.id
                  FROM tags AS t
                  WHERE t.tag_type_id = ?
              )
            """,
            (entry_id, part_of_speech_type_id),
        )

        normalized_part_of_speech = normalize_optional_text(part_of_speech or "")
        if normalized_part_of_speech is None:
            return

        predefined_set = {value.lower() for value in self.PREDEFINED_PART_OF_SPEECH_TAGS}
        is_predefined = normalized_part_of_speech.lower() in predefined_set
        tag_id = self._get_or_create_tag(
            connection,
            part_of_speech_type_id,
            normalized_part_of_speech,
            is_predefined=is_predefined,
        )
        connection.execute(
            """
            INSERT INTO entry_tags (entry_id, tag_id)
            VALUES (?, ?)
            ON CONFLICT(entry_id, tag_id) DO NOTHING
            """,
            (entry_id, tag_id),
        )

    def get_setting(self, key: str, default: str | None = None) -> str | None:
        connection = sqlite3.connect(self.db_path)
        try:
            row = connection.execute(
                """
                SELECT value
                FROM app_settings
                WHERE key = ?
                """,
                (key,),
            ).fetchone()
            if row is None:
                return default
            return str(row[0])
        finally:
            connection.close()

    def set_setting(self, key: str, value: str) -> None:
        connection = sqlite3.connect(self.db_path)
        try:
            connection.execute(
                """
                INSERT INTO app_settings (key, value)
                VALUES (?, ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (key, value),
            )
            connection.commit()
        finally:
            connection.close()

    @staticmethod
    def _normalize_workbook_name(name: str) -> str:
        normalized = name.strip()
        if not normalized:
            raise ValidationError("Workbook name is required.")
        return normalized

    def list_workbooks(self) -> list[Workbook]:
        connection = sqlite3.connect(self.db_path)
        try:
            rows = connection.execute(
                """
                SELECT id, name, target_language_code, preset_key, created_at
                FROM workbooks
                ORDER BY id ASC
                """
            ).fetchall()
            return [
                Workbook(
                    id=int(row[0]),
                    name=str(row[1]),
                    target_language_code=validate_language_code(str(row[2]), "Target language"),
                    preset_key=str(row[3]) if row[3] is not None else "generic",
                    created_at=str(row[4]),
                )
                for row in rows
            ]
        finally:
            connection.close()

    def get_workbook(self, workbook_id: int) -> Workbook:
        connection = sqlite3.connect(self.db_path)
        try:
            row = connection.execute(
                """
                SELECT id, name, target_language_code, preset_key, created_at
                FROM workbooks
                WHERE id = ?
                """,
                (int(workbook_id),),
            ).fetchone()
        finally:
            connection.close()

        if row is None:
            raise LookupError(f"Workbook with id {int(workbook_id)} was not found.")

        return Workbook(
            id=int(row[0]),
            name=str(row[1]),
            target_language_code=validate_language_code(str(row[2]), "Target language"),
            preset_key=str(row[3]) if row[3] is not None else "generic",
            created_at=str(row[4]),
        )

    def create_workbook(self, name: str, target_language_code: str, preset_key: str = "generic") -> Workbook:
        normalized_name = self._normalize_workbook_name(name)
        target_language = validate_language_code(target_language_code, "Target language")
        normalized_preset_key = preset_key.strip().lower() or "generic"
        if normalized_preset_key not in {"generic", "japanese"}:
            raise ValidationError("Preset must be one of: generic, japanese.")
        if normalized_preset_key == "japanese" and target_language != "JP":
            raise ValidationError("Japanese preset is only available for JP workbooks.")

        connection = sqlite3.connect(self.db_path)
        try:
            cursor = connection.execute(
                """
                INSERT INTO workbooks (name, target_language_code, preset_key)
                VALUES (?, ?, ?)
                """,
                (normalized_name, target_language, normalized_preset_key),
            )
            workbook_id = int(cursor.lastrowid)

            if normalized_preset_key == "japanese":
                self._ensure_predefined_tags(connection, "JP")

            connection.commit()
        except sqlite3.IntegrityError as exc:
            raise ValidationError(f"Workbook '{normalized_name}' already exists.") from exc
        finally:
            connection.close()

        return self.get_workbook(workbook_id)

    def get_current_workbook_id(self) -> int:
        connection = sqlite3.connect(self.db_path)
        try:
            fallback_workbook_id = self._ensure_default_workbook(connection)
            workbook_id = self._read_current_workbook_id_from_connection(connection, fallback_workbook_id)
            connection.execute(
                """
                INSERT INTO app_settings (key, value)
                VALUES ('current_workbook_id', ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (str(workbook_id),),
            )
            connection.commit()
            return workbook_id
        finally:
            connection.close()

    def set_current_workbook_id(self, workbook_id: int) -> Workbook:
        resolved_workbook = self.get_workbook(workbook_id)

        connection = sqlite3.connect(self.db_path)
        try:
            connection.execute(
                """
                INSERT INTO app_settings (key, value)
                VALUES ('current_workbook_id', ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (str(resolved_workbook.id),),
            )
            connection.execute(
                """
                INSERT INTO app_settings (key, value)
                VALUES ('target_language', ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (resolved_workbook.target_language_code,),
            )
            connection.commit()
        finally:
            connection.close()

        return resolved_workbook

    def get_language_settings(self) -> tuple[str, str]:
        target_raw = self.get_setting("target_language", "JP") or "JP"
        assistant_raw = self.get_setting("assistant_language", "EN") or "EN"

        try:
            target = validate_language_code(target_raw, "Target language")
        except ValidationError:
            target = "JP"

        try:
            assistant = validate_language_code(assistant_raw, "Assistant language")
        except ValidationError:
            assistant = "EN" if target != "EN" else "JP"

        if target == assistant:
            assistant = "EN" if target == "JP" else "JP"

        return target, assistant

    def set_language_settings(self, target_language: str, assistant_language: str) -> tuple[str, str]:
        target = validate_language_code(target_language, "Target language")
        assistant = validate_language_code(assistant_language, "Assistant language")
        if target == assistant:
            raise ValidationError("Target and assistant languages must be different.")

        connection = sqlite3.connect(self.db_path)
        try:
            connection.execute(
                """
                INSERT INTO app_settings (key, value)
                VALUES ('target_language', ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (target,),
            )
            connection.execute(
                """
                INSERT INTO app_settings (key, value)
                VALUES ('assistant_language', ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (assistant,),
            )
            self._ensure_predefined_tags(connection, target)
            self._migrate_legacy_part_of_speech_tags(connection, target)
            connection.commit()
        finally:
            connection.close()

        return target, assistant

    def list_entries(self, workbook_id: int | None = None) -> list[VocabEntry]:
        resolved_workbook_id = self._resolve_workbook_id(workbook_id)
        connection = sqlite3.connect(self.db_path)
        try:
            connection.row_factory = sqlite3.Row
            rows: Iterable[sqlite3.Row] = connection.execute(
                """
                SELECT id, japanese_text, kana_text, english_text, part_of_speech, details_markdown, created_at
                FROM vocab_entries
                WHERE workbook_id = ?
                ORDER BY id ASC
                """,
                (resolved_workbook_id,),
            )
            return [self._map_row(row) for row in rows]
        finally:
            connection.close()

    def count_entries(self, workbook_id: int | None = None) -> int:
        resolved_workbook_id = self._resolve_workbook_id(workbook_id)
        connection = sqlite3.connect(self.db_path)
        try:
            row = connection.execute(
                """
                SELECT COUNT(*)
                FROM vocab_entries
                WHERE workbook_id = ?
                """,
                (resolved_workbook_id,),
            ).fetchone()
            if row is None:
                return 0
            return int(row[0])
        finally:
            connection.close()

    def count_distinct_english_meanings(self, workbook_id: int | None = None) -> int:
        resolved_workbook_id = self._resolve_workbook_id(workbook_id)
        connection = sqlite3.connect(self.db_path)
        try:
            row = connection.execute(
                """
                SELECT COUNT(*)
                FROM (
                    SELECT DISTINCT LOWER(TRIM(english_text)) AS normalized_english
                    FROM vocab_entries
                    WHERE workbook_id = ?
                      AND TRIM(english_text) <> ''
                )
                """,
                (resolved_workbook_id,),
            ).fetchone()
            if row is None:
                return 0
            return int(row[0])
        finally:
            connection.close()

    def get_random_entries(self, count: int, workbook_id: int | None = None) -> list[VocabEntry]:
        requested = max(int(count), 0)
        if requested == 0:
            return []

        resolved_workbook_id = self._resolve_workbook_id(workbook_id)

        connection = sqlite3.connect(self.db_path)
        try:
            connection.row_factory = sqlite3.Row
            rows: Iterable[sqlite3.Row] = connection.execute(
                """
                SELECT id, japanese_text, kana_text, english_text, part_of_speech, details_markdown, created_at
                FROM vocab_entries
                WHERE workbook_id = ?
                ORDER BY RANDOM()
                LIMIT ?
                """,
                (resolved_workbook_id, requested),
            )
            return [self._map_row(row) for row in rows]
        finally:
            connection.close()

    def list_entries_with_stats(
        self,
        sort_mode: str = "time",
        time_order: str = "newest",
        filter_tag_ids: Iterable[int] | None = None,
        filter_match_mode: str = "all",
        target_language_code: str | None = None,
        workbook_id: int | None = None,
    ) -> list[tuple[VocabEntry, int, int, str]]:
        resolved_workbook_id = self._resolve_workbook_id(workbook_id)
        resolved_target_language_code = self._resolve_target_language_code(target_language_code)
        order_by = "ORDER BY e.id DESC"
        order_params: list[object] = []
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
        elif sort_mode == "tags":
            order_by = (
                "ORDER BY COALESCE(("
                "SELECT GROUP_CONCAT(tt.name || ':' || t.name, '|') "
                "FROM entry_tags AS et "
                "INNER JOIN tags AS t ON t.id = et.tag_id "
                "INNER JOIN tag_types AS tt ON tt.id = t.tag_type_id "
                "WHERE et.entry_id = e.id "
                "AND tt.target_language_code = ?"
                "), '') ASC, "
                "e.id DESC"
            )
            order_params.append(resolved_target_language_code)
        elif time_order == "oldest":
            order_by = "ORDER BY e.id ASC"

        unique_filter_tag_ids = sorted({int(tag_id) for tag_id in (filter_tag_ids or [])})
        normalized_match_mode = filter_match_mode.lower().strip()
        if normalized_match_mode not in {"all", "any"}:
            normalized_match_mode = "all"

        where_clauses = ["e.workbook_id = ?"]
        params: list[object] = [resolved_workbook_id]
        if unique_filter_tag_ids:
            placeholders = ",".join("?" for _ in unique_filter_tag_ids)
            tag_filter_params: list[object] = [resolved_target_language_code, *unique_filter_tag_ids]

            if normalized_match_mode == "all":
                where_clauses.append(
                    "e.id IN ("
                    "SELECT et.entry_id "
                    "FROM entry_tags AS et "
                    "INNER JOIN tags AS t ON t.id = et.tag_id "
                    "INNER JOIN tag_types AS tt ON tt.id = t.tag_type_id "
                    "WHERE tt.target_language_code = ? "
                    f"AND et.tag_id IN ({placeholders}) "
                    "GROUP BY et.entry_id "
                    "HAVING COUNT(DISTINCT et.tag_id) = ?"
                    ")"
                )
                tag_filter_params.append(len(unique_filter_tag_ids))
            else:
                where_clauses.append(
                    "e.id IN ("
                    "SELECT et.entry_id "
                    "FROM entry_tags AS et "
                    "INNER JOIN tags AS t ON t.id = et.tag_id "
                    "INNER JOIN tag_types AS tt ON tt.id = t.tag_type_id "
                    "WHERE tt.target_language_code = ? "
                    f"AND et.tag_id IN ({placeholders})"
                    ")"
                )

            params.extend(tag_filter_params)

        where_clause = f"WHERE {' AND '.join(where_clauses)}"

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
                    e.part_of_speech,
                    e.details_markdown,
                    e.created_at,
                    COALESCE(s.error_count, 0) AS error_count,
                    COALESCE(s.test_count, 0) AS test_count
                FROM vocab_entries AS e
                LEFT JOIN vocab_stats AS s
                    ON s.entry_id = e.id
                {where_clause}
                {order_by}
                """,
                tuple(params + order_params),
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

    def list_tag_types(self, target_language_code: str | None = None) -> list[tuple[int, str, bool]]:
        resolved_target_language_code = self._resolve_target_language_code(target_language_code)

        connection = sqlite3.connect(self.db_path)
        try:
            rows = connection.execute(
                """
                SELECT id, name, is_predefined
                FROM tag_types
                WHERE target_language_code = ?
                ORDER BY is_predefined DESC, name ASC
                """,
                (resolved_target_language_code,),
            ).fetchall()
            return [(int(row[0]), str(row[1]), bool(row[2])) for row in rows]
        finally:
            connection.close()

    def add_tag_type(self, name: str, target_language_code: str | None = None) -> int:
        resolved_target_language_code = self._resolve_target_language_code(target_language_code)
        normalized_name = self._normalize_tag_name(name, "Tag type")

        connection = sqlite3.connect(self.db_path)
        try:
            cursor = connection.execute(
                """
                INSERT INTO tag_types (target_language_code, name, is_predefined)
                VALUES (?, ?, 0)
                """,
                (resolved_target_language_code, normalized_name),
            )
            connection.commit()
            return int(cursor.lastrowid)
        except sqlite3.IntegrityError as exc:
            raise ValidationError(f"Tag type '{normalized_name}' already exists.") from exc
        finally:
            connection.close()

    def delete_tag_type(self, tag_type_id: int) -> None:
        connection = sqlite3.connect(self.db_path)
        try:
            row = connection.execute(
                """
                SELECT is_predefined
                FROM tag_types
                WHERE id = ?
                """,
                (tag_type_id,),
            ).fetchone()
            if row is None:
                raise LookupError(f"Tag type with id {tag_type_id} was not found.")
            if int(row[0]) == 1:
                raise ValueError("Predefined tag types cannot be deleted.")

            connection.execute(
                """
                DELETE FROM tag_types
                WHERE id = ?
                """,
                (tag_type_id,),
            )
            connection.commit()
        finally:
            connection.close()

    def list_tags(
        self,
        target_language_code: str | None = None,
        tag_type_id: int | None = None,
        include_part_of_speech: bool = True,
    ) -> list[tuple[int, int, str, str, bool, bool]]:
        resolved_target_language_code = self._resolve_target_language_code(target_language_code)

        where_clauses = ["tt.target_language_code = ?"]
        params: list[object] = [resolved_target_language_code]

        if tag_type_id is not None:
            where_clauses.append("tt.id = ?")
            params.append(int(tag_type_id))
        if not include_part_of_speech:
            where_clauses.append("LOWER(tt.name) <> 'part_of_speech'")

        where_sql = " AND ".join(where_clauses)

        connection = sqlite3.connect(self.db_path)
        try:
            rows = connection.execute(
                f"""
                SELECT
                    t.id,
                    tt.id,
                    tt.name,
                    t.name,
                    tt.is_predefined,
                    t.is_predefined
                FROM tags AS t
                INNER JOIN tag_types AS tt
                    ON tt.id = t.tag_type_id
                WHERE {where_sql}
                ORDER BY tt.name ASC, t.name ASC
                """,
                tuple(params),
            ).fetchall()
            return [
                (int(row[0]), int(row[1]), str(row[2]), str(row[3]), bool(row[4]), bool(row[5]))
                for row in rows
            ]
        finally:
            connection.close()

    def add_tag(self, tag_type_id: int, name: str) -> int:
        normalized_name = self._normalize_tag_name(name, "Tag")

        connection = sqlite3.connect(self.db_path)
        try:
            type_row = connection.execute(
                """
                SELECT id
                FROM tag_types
                WHERE id = ?
                """,
                (tag_type_id,),
            ).fetchone()
            if type_row is None:
                raise LookupError(f"Tag type with id {tag_type_id} was not found.")

            cursor = connection.execute(
                """
                INSERT INTO tags (tag_type_id, name, is_predefined)
                VALUES (?, ?, 0)
                """,
                (tag_type_id, normalized_name),
            )
            connection.commit()
            return int(cursor.lastrowid)
        except sqlite3.IntegrityError as exc:
            raise ValidationError(f"Tag '{normalized_name}' already exists in this type.") from exc
        finally:
            connection.close()

    def delete_tag(self, tag_id: int) -> None:
        connection = sqlite3.connect(self.db_path)
        try:
            row = connection.execute(
                """
                SELECT is_predefined
                FROM tags
                WHERE id = ?
                """,
                (tag_id,),
            ).fetchone()
            if row is None:
                raise LookupError(f"Tag with id {tag_id} was not found.")
            if int(row[0]) == 1:
                raise ValueError("Predefined tags cannot be deleted.")

            connection.execute(
                """
                DELETE FROM tags
                WHERE id = ?
                """,
                (tag_id,),
            )
            connection.commit()
        finally:
            connection.close()

    def get_entry_tags(
        self,
        entry_id: int,
        target_language_code: str | None = None,
        include_part_of_speech: bool = True,
    ) -> list[tuple[int, int, str, str]]:
        resolved_target_language_code = self._resolve_target_language_code(target_language_code)

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

            where_clauses = ["tt.target_language_code = ?", "et.entry_id = ?"]
            params: list[object] = [resolved_target_language_code, entry_id]
            if not include_part_of_speech:
                where_clauses.append("LOWER(tt.name) <> 'part_of_speech'")

            where_sql = " AND ".join(where_clauses)
            rows = connection.execute(
                f"""
                SELECT t.id, tt.id, tt.name, t.name
                FROM entry_tags AS et
                INNER JOIN tags AS t
                    ON t.id = et.tag_id
                INNER JOIN tag_types AS tt
                    ON tt.id = t.tag_type_id
                WHERE {where_sql}
                ORDER BY tt.name ASC, t.name ASC
                """,
                tuple(params),
            ).fetchall()
            return [(int(row[0]), int(row[1]), str(row[2]), str(row[3])) for row in rows]
        finally:
            connection.close()

    def set_entry_tags(
        self,
        entry_id: int,
        tag_ids: Iterable[int],
        target_language_code: str | None = None,
        include_part_of_speech: bool = False,
    ) -> None:
        resolved_target_language_code = self._resolve_target_language_code(target_language_code)
        unique_tag_ids = sorted({int(tag_id) for tag_id in tag_ids})

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

            valid_tag_ids: set[int] = set()
            if unique_tag_ids:
                placeholders = ",".join("?" for _ in unique_tag_ids)
                rows = connection.execute(
                    f"""
                    SELECT t.id
                    FROM tags AS t
                    INNER JOIN tag_types AS tt
                        ON tt.id = t.tag_type_id
                    WHERE tt.target_language_code = ?
                      AND t.id IN ({placeholders})
                      {'AND LOWER(tt.name) <> \'part_of_speech\'' if not include_part_of_speech else ''}
                    """,
                    (resolved_target_language_code, *unique_tag_ids),
                ).fetchall()
                valid_tag_ids = {int(row[0]) for row in rows}
                if len(valid_tag_ids) != len(unique_tag_ids):
                    raise ValidationError("One or more selected tags are invalid for the current target language.")

            connection.execute(
                f"""
                DELETE FROM entry_tags
                WHERE entry_id = ?
                  AND tag_id IN (
                      SELECT t.id
                      FROM tags AS t
                      INNER JOIN tag_types AS tt
                          ON tt.id = t.tag_type_id
                      WHERE tt.target_language_code = ?
                      {'AND LOWER(tt.name) <> \'part_of_speech\'' if not include_part_of_speech else ''}
                  )
                """,
                (entry_id, resolved_target_language_code),
            )

            for tag_id in sorted(valid_tag_ids):
                connection.execute(
                    """
                    INSERT INTO entry_tags (entry_id, tag_id)
                    VALUES (?, ?)
                    ON CONFLICT(entry_id, tag_id) DO NOTHING
                    """,
                    (entry_id, tag_id),
                )

            connection.commit()
        finally:
            connection.close()

    def get_test_entries_by_preference(
        self,
        count: int,
        strategy: str = "strict",
        workbook_id: int | None = None,
    ) -> list[VocabEntry]:
        requested = max(int(count), 0)
        if requested == 0:
            return []

        resolved_workbook_id = self._resolve_workbook_id(workbook_id)

        if strategy == "weighted":
            entries_with_stats = self.list_entries_with_stats(
                sort_mode="time",
                time_order="newest",
                workbook_id=resolved_workbook_id,
            )
            if not entries_with_stats:
                return []

            weighted_pool = list(entries_with_stats)
            weights_by_tier = {"gray": 4, "red": 3, "yellow": 2, "green": 1}
            ordered: list[tuple[VocabEntry, int, int, str]] = []
            while weighted_pool and len(ordered) < requested:
                weights = [weights_by_tier[item[3]] for item in weighted_pool]
                selected_index = random.choices(range(len(weighted_pool)), weights=weights, k=1)[0]
                ordered.append(weighted_pool.pop(selected_index))
            return [entry for entry, _, _, _ in ordered]

        entries_with_stats = self._list_entries_with_stats_for_selection(workbook_id=resolved_workbook_id)
        if not entries_with_stats:
            return []

        buckets: dict[str, list[tuple[VocabEntry, int, int, str, str | None]]] = {
            "gray": [],
            "red": [],
            "yellow": [],
            "green": [],
        }
        for item in entries_with_stats:
            buckets[item[3]].append(item)

        ordered_strict: list[tuple[VocabEntry, int, int, str, str | None]] = []
        for tier in ("gray", "red", "yellow", "green"):
            buckets[tier].sort(
                key=lambda item: (
                    item[4] is not None,
                    item[4] or "",
                    item[0].id,
                )
            )
            ordered_strict.extend(buckets[tier])

        sliced = ordered_strict[:requested]
        return [entry for entry, _, _, _, _ in sliced]

    def _list_entries_with_stats_for_selection(
        self,
        workbook_id: int | None = None,
    ) -> list[tuple[VocabEntry, int, int, str, str | None]]:
        resolved_workbook_id = self._resolve_workbook_id(workbook_id)
        connection = sqlite3.connect(self.db_path)
        try:
            connection.row_factory = sqlite3.Row
            rows: Iterable[sqlite3.Row] = connection.execute(
                """
                SELECT
                    e.id,
                    e.japanese_text,
                    e.kana_text,
                    e.english_text,
                    e.part_of_speech,
                    e.details_markdown,
                    e.created_at,
                    COALESCE(s.error_count, 0) AS error_count,
                    COALESCE(s.test_count, 0) AS test_count,
                    s.last_tested AS last_tested
                FROM vocab_entries AS e
                LEFT JOIN vocab_stats AS s
                    ON s.entry_id = e.id
                WHERE e.workbook_id = ?
                """,
                (resolved_workbook_id,),
            )

            result: list[tuple[VocabEntry, int, int, str, str | None]] = []
            for row in rows:
                entry = self._map_row(row)
                error_count = int(row["error_count"])
                test_count = int(row["test_count"])
                tier = self._tier_from_counts(test_count, error_count)
                last_tested = str(row["last_tested"]) if row["last_tested"] is not None else None
                result.append((entry, test_count, error_count, tier, last_tested))
            return result
        finally:
            connection.close()

    def get_english_options_for_entry(
        self,
        entry_id: int,
        max_options: int = 4,
        workbook_id: int | None = None,
    ) -> list[str]:
        max_count = max(int(max_options), 2)
        resolved_workbook_id = self._resolve_workbook_id(workbook_id)

        connection = sqlite3.connect(self.db_path)
        try:
            row = connection.execute(
                """
                SELECT english_text
                FROM vocab_entries
                WHERE id = ?
                  AND workbook_id = ?
                """,
                (entry_id, resolved_workbook_id),
            ).fetchone()
            if row is None:
                raise LookupError(f"Vocabulary entry with id {entry_id} was not found.")

            correct_english = str(row[0]).strip()

            distractor_rows = connection.execute(
                """
                SELECT DISTINCT TRIM(english_text) AS english_text
                FROM vocab_entries
                WHERE id <> ?
                                    AND workbook_id = ?
                  AND TRIM(english_text) <> ''
                  AND LOWER(TRIM(english_text)) <> LOWER(TRIM(?))
                ORDER BY RANDOM()
                LIMIT ?
                """,
                                (entry_id, resolved_workbook_id, correct_english, max_count - 1),
            ).fetchall()

            options = [correct_english]
            options.extend(str(distractor_row[0]) for distractor_row in distractor_rows)
            random.shuffle(options)
            return options
        finally:
            connection.close()

    def record_test_result(
        self,
        entry_id: int,
        is_correct: bool,
        recovery_roll: float | None = None,
        practiced_on: date | None = None,
    ) -> None:
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

            practice_date = (practiced_on or date.today()).isoformat()
            connection.execute(
                """
                INSERT INTO practice_daily_unique (entry_id, practice_date)
                VALUES (?, ?)
                ON CONFLICT(entry_id, practice_date) DO NOTHING
                """,
                (entry_id, practice_date),
            )

            connection.execute(
                """
                INSERT INTO practice_daily_error_recovery (entry_id, practice_date)
                VALUES (?, ?)
                ON CONFLICT(entry_id, practice_date) DO NOTHING
                """,
                (entry_id, practice_date),
            )

            if not is_correct:
                connection.execute(
                    """
                    UPDATE practice_daily_error_recovery
                    SET has_mistake = 1
                    WHERE entry_id = ?
                      AND practice_date = ?
                    """,
                    (entry_id, practice_date),
                )
                connection.commit()
                return

            roll = recovery_roll if recovery_roll is not None else random.random()
            if roll >= self.ERROR_COUNT_RECOVERY_CHANCE:
                connection.commit()
                return

            row = connection.execute(
                """
                SELECT
                    s.error_count,
                    r.has_mistake,
                    r.used_decrease
                FROM vocab_stats AS s
                INNER JOIN practice_daily_error_recovery AS r
                    ON r.entry_id = s.entry_id
                WHERE s.entry_id = ?
                  AND r.practice_date = ?
                """,
                (entry_id, practice_date),
            ).fetchone()

            if row is not None:
                error_count = int(row[0])
                has_mistake = int(row[1])
                used_decrease = int(row[2])

                if error_count > 0 and has_mistake == 0 and used_decrease == 0:
                    connection.execute(
                        """
                        UPDATE vocab_stats
                        SET error_count = CASE WHEN error_count > 0 THEN error_count - 1 ELSE 0 END
                        WHERE entry_id = ?
                        """,
                        (entry_id,),
                    )
                    connection.execute(
                        """
                        UPDATE practice_daily_error_recovery
                        SET used_decrease = 1
                        WHERE entry_id = ?
                          AND practice_date = ?
                        """,
                        (entry_id, practice_date),
                    )

            connection.commit()
        finally:
            connection.close()

    def get_entry_last_practiced(self, entry_id: int) -> str | None:
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
                SELECT last_tested
                FROM vocab_stats
                WHERE entry_id = ?
                """,
                (entry_id,),
            ).fetchone()

            if row is None or row[0] is None:
                return None
            return str(row[0])
        finally:
            connection.close()

    def get_daily_unique_practice_counts(
        self,
        days_back: int = 180,
        workbook_id: int | None = None,
    ) -> dict[str, int]:
        range_days = max(int(days_back), 1)
        start_date = (date.today() - timedelta(days=range_days - 1)).isoformat()
        resolved_workbook_id = self._resolve_workbook_id(workbook_id)

        connection = sqlite3.connect(self.db_path)
        try:
            rows = connection.execute(
                """
                SELECT pdu.practice_date, COUNT(*) AS unique_count
                FROM practice_daily_unique AS pdu
                INNER JOIN vocab_entries AS e
                    ON e.id = pdu.entry_id
                WHERE pdu.practice_date >= ?
                  AND e.workbook_id = ?
                GROUP BY pdu.practice_date
                ORDER BY practice_date ASC
                """,
                (start_date, resolved_workbook_id),
            ).fetchall()

            return {str(row[0]): int(row[1]) for row in rows}
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

    def add_entry(
        self,
        japanese_text: str,
        kana_text: str,
        english_text: str,
        part_of_speech: str = "",
        workbook_id: int | None = None,
    ) -> VocabEntry:
        created = self.add_entries(
            [(japanese_text, kana_text, english_text, part_of_speech)],
            workbook_id=workbook_id,
        )
        if not created:
            raise RuntimeError("Could not load inserted entry.")
        return created[0]

    def add_entries(
        self,
        entries: Iterable[tuple[str, str, str] | tuple[str, str, str, str]],
        workbook_id: int | None = None,
    ) -> list[VocabEntry]:
        normalized_entries: list[tuple[str, str | None, str, str | None, str | None]] = []
        for entry in entries:
            if len(entry) == 3:
                japanese_text, kana_text, english_text = entry
                part_of_speech = ""
            elif len(entry) == 4:
                japanese_text, kana_text, english_text, part_of_speech = entry
            else:
                raise ValidationError("Each entry must contain 3 or 4 values.")

            japanese, english = validate_vocab_fields(japanese_text, english_text)
            kana = normalize_optional_text(kana_text)
            normalized_part_of_speech = normalize_optional_text(part_of_speech)
            normalized_entries.append((japanese, kana, english, normalized_part_of_speech, None))

        if not normalized_entries:
            return []

        connection = sqlite3.connect(self.db_path)
        try:
            connection.row_factory = sqlite3.Row
            resolved_workbook_id = self._resolve_workbook_id_from_connection(connection, workbook_id)
            target_language_code = self._read_workbook_target_language_from_connection(connection, resolved_workbook_id)
            preset_key = self._read_workbook_preset_key_from_connection(connection, resolved_workbook_id)
            if preset_key == "japanese":
                self._ensure_predefined_tags(connection, target_language_code)

            inserted_ids: list[int] = []
            for japanese, kana, english, part_of_speech, details_markdown in normalized_entries:
                cursor = connection.execute(
                    """
                    INSERT INTO vocab_entries (
                        workbook_id,
                        japanese_text,
                        kana_text,
                        english_text,
                        part_of_speech,
                        details_markdown
                    )
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (resolved_workbook_id, japanese, kana, english, part_of_speech, details_markdown),
                )
                inserted_ids.append(int(cursor.lastrowid))

            for inserted_id, (_japanese, _kana, _english, part_of_speech, _details_markdown) in zip(
                inserted_ids,
                normalized_entries,
                strict=True,
            ):
                self._sync_entry_part_of_speech_tag(
                    connection,
                    inserted_id,
                    part_of_speech,
                    target_language_code,
                )

            placeholders = ",".join("?" for _ in inserted_ids)
            rows = connection.execute(
                f"""
                SELECT id, japanese_text, kana_text, english_text, part_of_speech, details_markdown, created_at
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
                SELECT id, japanese_text, kana_text, english_text, part_of_speech, details_markdown, created_at
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

    def update_entry(
        self,
        entry_id: int,
        japanese_text: str,
        kana_text: str,
        english_text: str,
        part_of_speech: str = "",
    ) -> VocabEntry:
        japanese, english = validate_vocab_fields(japanese_text, english_text)
        kana = normalize_optional_text(kana_text)
        normalized_part_of_speech = normalize_optional_text(part_of_speech)

        connection = sqlite3.connect(self.db_path)
        row: sqlite3.Row | None = None
        try:
            connection.row_factory = sqlite3.Row
            workbook_row = connection.execute(
                """
                SELECT e.workbook_id, w.target_language_code
                FROM vocab_entries AS e
                INNER JOIN workbooks AS w
                    ON w.id = e.workbook_id
                WHERE e.id = ?
                """,
                (entry_id,),
            ).fetchone()
            if workbook_row is None:
                raise LookupError(f"Vocabulary entry with id {entry_id} was not found.")

            target_language_code = validate_language_code(str(workbook_row[1]), "Target language")

            cursor = connection.execute(
                """
                UPDATE vocab_entries
                SET japanese_text = ?, kana_text = ?, english_text = ?, part_of_speech = ?
                WHERE id = ?
                """,
                (japanese, kana, english, normalized_part_of_speech, entry_id),
            )
            if cursor.rowcount == 0:
                raise LookupError(f"Vocabulary entry with id {entry_id} was not found.")

            self._sync_entry_part_of_speech_tag(
                connection,
                entry_id,
                normalized_part_of_speech,
                target_language_code,
            )

            row = connection.execute(
                """
                SELECT id, japanese_text, kana_text, english_text, part_of_speech, details_markdown, created_at
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

    def update_entry_details(self, entry_id: int, details_markdown: str) -> None:
        normalized_details = normalize_optional_markdown(details_markdown)

        connection = sqlite3.connect(self.db_path)
        try:
            cursor = connection.execute(
                """
                UPDATE vocab_entries
                SET details_markdown = ?
                WHERE id = ?
                """,
                (normalized_details, entry_id),
            )
            if cursor.rowcount == 0:
                raise LookupError(f"Vocabulary entry with id {entry_id} was not found.")
            connection.commit()
        finally:
            connection.close()

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
            connection.execute(
                f"""
                DELETE FROM practice_daily_unique
                WHERE entry_id IN ({placeholders})
                """,
                tuple(unique_ids),
            )
            connection.execute(
                f"""
                DELETE FROM practice_daily_error_recovery
                WHERE entry_id IN ({placeholders})
                """,
                tuple(unique_ids),
            )
            connection.execute(
                f"""
                DELETE FROM entry_tags
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
        keys = set(row.keys())
        part_of_speech = row["part_of_speech"] if "part_of_speech" in keys else None
        details_markdown = row["details_markdown"] if "details_markdown" in keys else None

        return VocabEntry(
            id=int(row["id"]),
            japanese_text=str(row["japanese_text"]),
            kana_text=row["kana_text"],
            english_text=str(row["english_text"]),
            part_of_speech=str(part_of_speech) if part_of_speech is not None else None,
            details_markdown=str(details_markdown) if details_markdown is not None else None,
            created_at=str(row["created_at"]),
        )


def default_db_path() -> Path:
    return Path(__file__).resolve().parent.parent / "vocab.db"
