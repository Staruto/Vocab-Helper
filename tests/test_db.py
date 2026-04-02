import tempfile
import unittest
import sqlite3
from datetime import date, timedelta
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
        self.assertIsNone(rows[0].part_of_speech)
        self.assertIsNone(rows[0].details_markdown)

    def test_initialize_migrates_legacy_schema_and_adds_settings_table(self) -> None:
        legacy_db_path = Path(self.temp_dir.name) / "legacy_vocab.db"

        connection = sqlite3.connect(legacy_db_path)
        try:
            connection.execute(
                """
                CREATE TABLE vocab_entries (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    japanese_text TEXT NOT NULL,
                    kana_text TEXT NULL,
                    english_text TEXT NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE vocab_stats (
                    entry_id INTEGER PRIMARY KEY,
                    error_count INTEGER NOT NULL DEFAULT 0,
                    test_count INTEGER NOT NULL DEFAULT 0,
                    last_tested TEXT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            connection.commit()
        finally:
            connection.close()

    def test_initialize_backfills_part_of_speech_into_tags(self) -> None:
        legacy_db_path = Path(self.temp_dir.name) / "legacy_pos_vocab.db"

        connection = sqlite3.connect(legacy_db_path)
        try:
            connection.execute(
                """
                CREATE TABLE vocab_entries (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    japanese_text TEXT NOT NULL,
                    kana_text TEXT NULL,
                    english_text TEXT NOT NULL,
                    part_of_speech TEXT NULL,
                    details_markdown TEXT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            connection.execute(
                """
                CREATE TABLE vocab_stats (
                    entry_id INTEGER PRIMARY KEY,
                    error_count INTEGER NOT NULL DEFAULT 0,
                    test_count INTEGER NOT NULL DEFAULT 0,
                    last_tested TEXT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                )
                """
            )
            connection.execute(
                """
                INSERT INTO vocab_entries (japanese_text, kana_text, english_text, part_of_speech)
                VALUES ('書く', 'かく', 'to write', 'verb')
                """
            )
            connection.commit()
        finally:
            connection.close()

        legacy_repository = VocabRepository(legacy_db_path)
        legacy_repository.initialize()

        entry = legacy_repository.list_entries()[0]
        tags = legacy_repository.get_entry_tags(entry.id, target_language_code="JP")
        self.assertIn(
            ("part_of_speech", "verb"),
            {(type_name, tag_name) for _tag_id, _type_id, type_name, tag_name in tags},
        )

        legacy_repository = VocabRepository(legacy_db_path)
        legacy_repository.initialize()

        connection = sqlite3.connect(legacy_db_path)
        try:
            columns = {
                row[1]
                for row in connection.execute("PRAGMA table_info(vocab_entries)").fetchall()
            }
            self.assertIn("part_of_speech", columns)
            self.assertIn("details_markdown", columns)

            settings_rows = connection.execute(
                "SELECT key, value FROM app_settings ORDER BY key"
            ).fetchall()
            self.assertEqual(
                settings_rows,
                [
                    ("assistant_language", "EN"),
                    ("current_workbook_id", "1"),
                    ("target_language", "JP"),
                ],
            )
        finally:
            connection.close()

    def test_default_workbook_is_created_and_selected(self) -> None:
        workbooks = self.repository.list_workbooks()
        self.assertEqual(len(workbooks), 1)
        self.assertEqual(workbooks[0].name, "JP")
        self.assertEqual(workbooks[0].target_language_code, "JP")

        current_workbook_id = self.repository.get_current_workbook_id()
        self.assertEqual(current_workbook_id, workbooks[0].id)

    def test_language_properties_are_seeded_for_supported_languages(self) -> None:
        jp_properties = self.repository.list_language_properties("JP")
        jp_keys = {key for _property_id, key, _label, _is_predefined, _is_required in jp_properties}
        self.assertIn("target_text", jp_keys)
        self.assertIn("meaning", jp_keys)
        self.assertIn("kana", jp_keys)

        en_properties = self.repository.list_language_properties("EN")
        en_keys = {key for _property_id, key, _label, _is_predefined, _is_required in en_properties}
        self.assertIn("target_text", en_keys)
        self.assertIn("meaning", en_keys)
        self.assertNotIn("kana", en_keys)

    def test_workbook_visibility_defaults_exist(self) -> None:
        workbook_id = self.repository.get_current_workbook_id()
        self.assertIsNotNone(workbook_id)

        visibility_rows = self.repository.get_workbook_visible_properties(int(workbook_id))
        visibility_by_key = {
            key: is_visible
            for _property_id, key, _label, _is_predefined, _is_required, is_visible, _display_order in visibility_rows
        }
        self.assertTrue(visibility_by_key["target_text"])
        self.assertTrue(visibility_by_key["meaning"])
        self.assertTrue(visibility_by_key["kana"])

    def test_language_property_crud_and_required_property_protection(self) -> None:
        property_id = self.repository.add_language_property("JP", "example_note", "Example note")
        jp_keys = {
            key
            for _property_id, key, _label, _is_predefined, _is_required in self.repository.list_language_properties("JP")
        }
        self.assertIn("example_note", jp_keys)

        self.repository.delete_language_property(property_id)
        jp_keys_after_delete = {
            key
            for _property_id, key, _label, _is_predefined, _is_required in self.repository.list_language_properties("JP")
        }
        self.assertNotIn("example_note", jp_keys_after_delete)

        required_property_id = next(
            property_id
            for property_id, key, _label, _is_predefined, is_required in self.repository.list_language_properties("JP")
            if key == "target_text" and is_required
        )
        with self.assertRaises(ValueError):
            self.repository.delete_language_property(required_property_id)

    def test_entry_property_values_persist_and_sync_predefined_columns(self) -> None:
        entry = self.repository.add_entry("食べる", "たべる", "to eat")
        custom_property_id = self.repository.add_language_property("JP", "register", "Register")

        self.repository.set_entry_property_values(
            entry.id,
            {
                "target_text": "喰べる",
                "meaning": "to consume",
                "kana": "たべる",
                "register": "casual",
            },
        )

        values = self.repository.get_entry_property_values(entry.id)
        self.assertEqual(values["target_text"], "喰べる")
        self.assertEqual(values["meaning"], "to consume")
        self.assertEqual(values["register"], "casual")

        loaded = self.repository.get_entry(entry.id)
        self.assertEqual(loaded.japanese_text, "喰べる")
        self.assertEqual(loaded.english_text, "to consume")

        self.repository.delete_language_property(custom_property_id)
        values_after_property_delete = self.repository.get_entry_property_values(entry.id)
        self.assertNotIn("register", values_after_property_delete)

    def test_set_workbook_visible_properties_keeps_target_text_visible(self) -> None:
        workbook_id = self.repository.get_current_workbook_id()
        self.assertIsNotNone(workbook_id)

        rows = self.repository.get_workbook_visible_properties(int(workbook_id))
        meaning_property_id = next(property_id for property_id, key, *_rest in rows if key == "meaning")
        self.repository.set_workbook_visible_properties(int(workbook_id), [meaning_property_id])

        updated_rows = self.repository.get_workbook_visible_properties(int(workbook_id))
        visible_keys = {
            key
            for _property_id, key, _label, _is_predefined, _is_required, is_visible, _display_order in updated_rows
            if is_visible
        }
        self.assertIn("target_text", visible_keys)
        self.assertIn("meaning", visible_keys)

    def test_deleting_last_workbook_is_allowed(self) -> None:
        workbook_id = self.repository.get_current_workbook_id()
        self.assertIsNotNone(workbook_id)
        entry = self.repository.add_entry("食べる", "たべる", "to eat", workbook_id=int(workbook_id))
        self.assertGreater(entry.id, 0)

        next_current = self.repository.delete_workbook(int(workbook_id))
        self.assertIsNone(next_current)
        self.assertEqual(self.repository.get_current_workbook_id(), None)
        self.assertEqual(self.repository.list_workbooks(), [])

    def test_initialize_renames_legacy_default_workbook_name_to_jp(self) -> None:
        current_workbook_id = self.repository.get_current_workbook_id()

        connection = sqlite3.connect(self.db_path)
        try:
            connection.execute(
                """
                UPDATE workbooks
                SET name = 'Default', target_language_code = 'JP'
                WHERE id = ?
                """,
                (current_workbook_id,),
            )
            connection.commit()
        finally:
            connection.close()

        self.repository.initialize()
        renamed_workbook = self.repository.get_workbook(current_workbook_id)
        self.assertEqual(renamed_workbook.name, "JP")

    def test_entries_are_scoped_by_workbook(self) -> None:
        default_workbook_id = self.repository.get_current_workbook_id()
        second_workbook = self.repository.create_workbook("English workbook", "EN", preset_key="generic")

        self.repository.add_entry("食べる", "たべる", "to eat", workbook_id=default_workbook_id)
        self.repository.add_entry("run", "", "to run", workbook_id=second_workbook.id)

        default_entries = self.repository.list_entries(workbook_id=default_workbook_id)
        second_entries = self.repository.list_entries(workbook_id=second_workbook.id)

        self.assertEqual(len(default_entries), 1)
        self.assertEqual(default_entries[0].japanese_text, "食べる")
        self.assertEqual(len(second_entries), 1)
        self.assertEqual(second_entries[0].japanese_text, "run")

    def test_get_english_options_are_scoped_by_workbook(self) -> None:
        default_workbook_id = self.repository.get_current_workbook_id()
        second_workbook = self.repository.create_workbook("Workbook two", "EN", preset_key="generic")

        first = self.repository.add_entry("食べる", "たべる", "to eat", workbook_id=default_workbook_id)
        self.repository.add_entry("行く", "いく", "to go", workbook_id=default_workbook_id)
        self.repository.add_entry("run", "", "to run", workbook_id=second_workbook.id)

        options = self.repository.get_english_options_for_entry(first.id, max_options=4, workbook_id=default_workbook_id)
        normalized = {option.strip().lower() for option in options}
        self.assertIn("to eat", normalized)
        self.assertIn("to go", normalized)
        self.assertNotIn("to run", normalized)

    def test_count_entries_reflects_total_rows(self) -> None:
        self.assertEqual(self.repository.count_entries(), 0)

        self.repository.add_entry("食べる", "たべる", "to eat")
        self.repository.add_entry("行く", "いく", "to go")

        self.assertEqual(self.repository.count_entries(), 2)

    def test_language_settings_defaults_and_updates(self) -> None:
        target, assistant = self.repository.get_language_settings()
        self.assertEqual((target, assistant), ("JP", "EN"))

        updated_target, updated_assistant = self.repository.set_language_settings("EN", "JP")
        self.assertEqual((updated_target, updated_assistant), ("EN", "JP"))

        target_after, assistant_after = self.repository.get_language_settings()
        self.assertEqual((target_after, assistant_after), ("EN", "JP"))

        with self.assertRaises(ValidationError):
            self.repository.set_language_settings("JP", "JP")

    def test_predefined_tag_types_and_tags_are_seeded_for_target_language(self) -> None:
        tag_types = self.repository.list_tag_types(target_language_code="JP")
        tag_type_names = {name for _tag_type_id, name, _is_predefined in tag_types}
        self.assertIn("part_of_speech", tag_type_names)
        self.assertIn("difficulty", tag_type_names)

        tags = self.repository.list_tags(target_language_code="JP")
        by_type: dict[str, set[str]] = {}
        for _tag_id, _type_id, type_name, tag_name, _type_predefined, _tag_predefined in tags:
            by_type.setdefault(type_name, set()).add(tag_name)

        self.assertTrue(set(VocabRepository.PREDEFINED_PART_OF_SPEECH_TAGS).issubset(by_type.get("part_of_speech", set())))
        self.assertTrue(set(VocabRepository.PREDEFINED_DIFFICULTY_TAGS).issubset(by_type.get("difficulty", set())))

    def test_custom_tag_type_and_tag_crud(self) -> None:
        custom_type_id = self.repository.add_tag_type("topic")
        all_types = self.repository.list_tag_types()
        self.assertIn(custom_type_id, [tag_type_id for tag_type_id, _name, _is_predefined in all_types])

        grammar_tag_id = self.repository.add_tag(custom_type_id, "grammar")
        reading_tag_id = self.repository.add_tag(custom_type_id, "reading")
        tags = self.repository.list_tags(tag_type_id=custom_type_id)
        self.assertEqual(
            {tag_id for tag_id, _type_id, _type_name, _tag_name, _type_predefined, _tag_predefined in tags},
            {grammar_tag_id, reading_tag_id},
        )

        self.repository.delete_tag(reading_tag_id)
        tags_after_delete = self.repository.list_tags(tag_type_id=custom_type_id)
        self.assertEqual(len(tags_after_delete), 1)
        self.assertEqual(tags_after_delete[0][0], grammar_tag_id)

        self.repository.delete_tag_type(custom_type_id)
        remaining_type_ids = [tag_type_id for tag_type_id, _name, _is_predefined in self.repository.list_tag_types()]
        self.assertNotIn(custom_type_id, remaining_type_ids)

    def test_predefined_tag_type_and_tag_cannot_be_deleted(self) -> None:
        tag_types = self.repository.list_tag_types()
        part_of_speech_type_id = next(tag_type_id for tag_type_id, name, _is_predefined in tag_types if name == "part_of_speech")

        with self.assertRaises(ValueError):
            self.repository.delete_tag_type(part_of_speech_type_id)

        part_of_speech_tags = self.repository.list_tags(tag_type_id=part_of_speech_type_id)
        noun_tag_id = next(tag_id for tag_id, _type_id, _type_name, tag_name, _type_predefined, _tag_predefined in part_of_speech_tags if tag_name == "noun")
        with self.assertRaises(ValueError):
            self.repository.delete_tag(noun_tag_id)

    def test_set_and_get_entry_tags(self) -> None:
        entry = self.repository.add_entry("語る", "かたる", "to tell")
        topic_type_id = self.repository.add_tag_type("topic")
        story_tag_id = self.repository.add_tag(topic_type_id, "story")
        media_tag_id = self.repository.add_tag(topic_type_id, "media")

        self.repository.set_entry_tags(entry.id, [story_tag_id, media_tag_id])
        entry_tags = self.repository.get_entry_tags(entry.id, include_part_of_speech=False)
        self.assertEqual({tag_id for tag_id, _type_id, _type_name, _tag_name in entry_tags}, {story_tag_id, media_tag_id})

        self.repository.set_entry_tags(entry.id, [story_tag_id])
        entry_tags_after_replace = self.repository.get_entry_tags(entry.id, include_part_of_speech=False)
        self.assertEqual({tag_id for tag_id, _type_id, _type_name, _tag_name in entry_tags_after_replace}, {story_tag_id})

    def test_list_entries_with_stats_filters_by_all_selected_tags(self) -> None:
        created = self.repository.add_entries(
            [
                ("読む", "よむ", "to read", "verb"),
                ("書く", "かく", "to write", "verb"),
                ("青", "あお", "blue", "noun"),
            ]
        )
        first_id, second_id, third_id = [entry.id for entry in created]

        topic_type_id = self.repository.add_tag_type("topic")
        exam_tag_id = self.repository.add_tag(topic_type_id, "exam")
        daily_tag_id = self.repository.add_tag(topic_type_id, "daily")

        self.repository.set_entry_tags(first_id, [exam_tag_id, daily_tag_id])
        self.repository.set_entry_tags(second_id, [exam_tag_id])
        self.repository.set_entry_tags(third_id, [daily_tag_id])

        filtered_rows = self.repository.list_entries_with_stats(
            sort_mode="time",
            time_order="oldest",
            filter_tag_ids=[exam_tag_id, daily_tag_id],
            filter_match_mode="all",
            target_language_code="JP",
        )
        filtered_ids = [entry.id for entry, _test_count, _error_count, _tier in filtered_rows]
        self.assertEqual(filtered_ids, [first_id])

    def test_list_entries_with_stats_supports_tag_sort_mode(self) -> None:
        created = self.repository.add_entries(
            [
                ("読む", "よむ", "to read", "verb"),
                ("書く", "かく", "to write", "verb"),
                ("青", "あお", "blue", "noun"),
            ]
        )
        first_id, second_id, third_id = [entry.id for entry in created]

        topic_type_id = self.repository.add_tag_type("topic")
        alpha_tag_id = self.repository.add_tag(topic_type_id, "alpha")
        beta_tag_id = self.repository.add_tag(topic_type_id, "beta")

        self.repository.set_entry_tags(first_id, [alpha_tag_id])
        self.repository.set_entry_tags(second_id, [beta_tag_id])

        rows = self.repository.list_entries_with_stats(
            sort_mode="tags",
            time_order="newest",
            target_language_code="JP",
        )
        sorted_ids = [entry.id for entry, _test_count, _error_count, _tier in rows]
        self.assertEqual(sorted_ids, [third_id, first_id, second_id])

    def test_part_of_speech_field_syncs_to_part_of_speech_tags(self) -> None:
        entry = self.repository.add_entry("考える", "かんがえる", "to think", "verb")
        part_of_speech_tags = self.repository.get_entry_tags(entry.id)
        self.assertIn(
            ("part_of_speech", "verb"),
            {(type_name, tag_name) for _tag_id, _type_id, type_name, tag_name in part_of_speech_tags},
        )

        self.repository.update_entry(entry.id, "考える", "かんがえる", "to think", "noun")
        updated_tags = self.repository.get_entry_tags(entry.id)
        self.assertIn(
            ("part_of_speech", "noun"),
            {(type_name, tag_name) for _tag_id, _type_id, type_name, tag_name in updated_tags},
        )
        self.assertNotIn(
            ("part_of_speech", "verb"),
            {(type_name, tag_name) for _tag_id, _type_id, type_name, tag_name in updated_tags},
        )

        self.repository.update_entry(entry.id, "考える", "かんがえる", "to think", "")
        cleared_tags = self.repository.get_entry_tags(entry.id)
        self.assertNotIn(
            "part_of_speech",
            {type_name for _tag_id, _type_id, type_name, _tag_name in cleared_tags},
        )

    def test_tag_types_are_scoped_by_target_language(self) -> None:
        jp_types = {name for _id, name, _is_predefined in self.repository.list_tag_types("JP")}
        self.assertIn("difficulty", jp_types)

        self.repository.set_language_settings("EN", "JP")
        en_types = {name for _id, name, _is_predefined in self.repository.list_tag_types("EN")}
        self.assertIn("difficulty", en_types)

        self.repository.add_tag_type("news", target_language_code="EN")
        en_types_after_add = {name for _id, name, _is_predefined in self.repository.list_tag_types("EN")}
        jp_types_after_add = {name for _id, name, _is_predefined in self.repository.list_tag_types("JP")}

        self.assertIn("news", en_types_after_add)
        self.assertNotIn("news", jp_types_after_add)

    def test_count_distinct_english_meanings_normalizes_case_and_spaces(self) -> None:
        self.repository.add_entries(
            [
                ("食べる", "たべる", "to eat"),
                ("食う", "くう", "  TO EAT  "),
                ("行く", "いく", "to go"),
            ]
        )

        self.assertEqual(self.repository.count_distinct_english_meanings(), 2)

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

    def test_get_english_options_for_entry_returns_distinct_choices_with_correct_answer(self) -> None:
        created = self.repository.add_entries(
            [
                ("食べる", "たべる", "to eat"),
                ("食う", "くう", " TO EAT "),
                ("行く", "いく", "to go"),
                ("飲む", "のむ", "to drink"),
                ("書く", "かく", "to write"),
            ]
        )
        target = created[0]

        options = self.repository.get_english_options_for_entry(target.id, max_options=4)

        normalized = [option.strip().lower() for option in options]
        self.assertGreaterEqual(len(options), 2)
        self.assertLessEqual(len(options), 4)
        self.assertIn(target.english_text.strip().lower(), normalized)
        self.assertEqual(len(normalized), len(set(normalized)))

    def test_get_english_options_for_entry_falls_back_when_pool_is_small(self) -> None:
        created = self.repository.add_entries(
            [
                ("食べる", "たべる", "to eat"),
                ("行く", "いく", "to go"),
            ]
        )

        options = self.repository.get_english_options_for_entry(created[0].id, max_options=4)
        self.assertEqual(len(options), 2)
        self.assertIn("to eat", options)
        self.assertIn("to go", options)

    def test_get_english_options_for_entry_returns_single_option_when_no_alternative_exists(self) -> None:
        created = self.repository.add_entries(
            [
                ("食べる", "たべる", "to eat"),
                ("食う", "くう", "TO EAT"),
            ]
        )

        options = self.repository.get_english_options_for_entry(created[0].id, max_options=4)
        self.assertEqual(len(options), 1)
        self.assertEqual(options[0].strip().lower(), "to eat")

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
        inserted = self.repository.add_entry("書く", "かく", "to write", "verb")

        loaded = self.repository.get_entry(inserted.id)
        self.assertEqual(loaded.id, inserted.id)
        self.assertEqual(loaded.japanese_text, "書く")
        self.assertEqual(loaded.kana_text, "かく")
        self.assertEqual(loaded.english_text, "to write")
        self.assertEqual(loaded.part_of_speech, "verb")
        self.assertIsNone(loaded.details_markdown)

    def test_update_entry_changes_values(self) -> None:
        inserted = self.repository.add_entry("読む", "よむ", "to read")

        updated = self.repository.update_entry(inserted.id, "飲む", "のむ", "to drink", "verb")
        self.assertEqual(updated.id, inserted.id)
        self.assertEqual(updated.japanese_text, "飲む")
        self.assertEqual(updated.kana_text, "のむ")
        self.assertEqual(updated.english_text, "to drink")
        self.assertEqual(updated.part_of_speech, "verb")

        loaded = self.repository.get_entry(inserted.id)
        self.assertEqual(loaded.japanese_text, "飲む")

    def test_update_entry_details_persists_markdown_content(self) -> None:
        inserted = self.repository.add_entry("学ぶ", "まなぶ", "to study")

        self.repository.update_entry_details(inserted.id, "# Notes\n- useful word\n**important**")
        loaded = self.repository.get_entry(inserted.id)
        self.assertEqual(loaded.details_markdown, "# Notes\n- useful word\n**important**")

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

    def test_correct_answer_can_reduce_error_count_once_per_day(self) -> None:
        entry = self.repository.add_entry("書く", "かく", "to write")
        yesterday = date.today() - timedelta(days=1)
        today = date.today()

        self.repository.record_test_result(entry.id, is_correct=False, practiced_on=yesterday)
        self.repository.record_test_result(entry.id, is_correct=False, practiced_on=yesterday)
        self.assertEqual(self.repository.get_entry_stats(entry.id), (2, 2, "yellow"))

        self.repository.record_test_result(entry.id, is_correct=True, recovery_roll=0.0, practiced_on=today)
        self.assertEqual(self.repository.get_entry_stats(entry.id), (3, 1, "yellow"))

        self.repository.record_test_result(entry.id, is_correct=True, recovery_roll=0.0, practiced_on=today)
        self.assertEqual(self.repository.get_entry_stats(entry.id), (4, 1, "yellow"))

    def test_mistake_blocks_same_day_error_count_recovery(self) -> None:
        entry = self.repository.add_entry("読む", "よむ", "to read")
        yesterday = date.today() - timedelta(days=1)
        today = date.today()

        self.repository.record_test_result(entry.id, is_correct=False, practiced_on=yesterday)
        self.repository.record_test_result(entry.id, is_correct=False, practiced_on=yesterday)
        self.assertEqual(self.repository.get_entry_stats(entry.id), (2, 2, "yellow"))

        self.repository.record_test_result(entry.id, is_correct=False, practiced_on=today)
        self.assertEqual(self.repository.get_entry_stats(entry.id), (3, 3, "red"))

        self.repository.record_test_result(entry.id, is_correct=True, recovery_roll=0.0, practiced_on=today)
        self.assertEqual(self.repository.get_entry_stats(entry.id), (4, 3, "red"))

    def test_error_count_recovery_resets_on_next_day(self) -> None:
        entry = self.repository.add_entry("行く", "いく", "to go")
        day_one = date.today() - timedelta(days=2)
        day_two = date.today() - timedelta(days=1)
        day_three = date.today()

        self.repository.record_test_result(entry.id, is_correct=False, practiced_on=day_one)
        self.repository.record_test_result(entry.id, is_correct=False, practiced_on=day_one)
        self.assertEqual(self.repository.get_entry_stats(entry.id), (2, 2, "yellow"))

        self.repository.record_test_result(entry.id, is_correct=True, recovery_roll=0.0, practiced_on=day_two)
        self.assertEqual(self.repository.get_entry_stats(entry.id), (3, 1, "yellow"))

        self.repository.record_test_result(entry.id, is_correct=True, recovery_roll=0.0, practiced_on=day_three)
        self.assertEqual(self.repository.get_entry_stats(entry.id), (4, 0, "green"))

    def test_get_entry_last_practiced_for_untested_and_tested_entry(self) -> None:
        entry = self.repository.add_entry("覚える", "おぼえる", "to memorize")
        self.assertIsNone(self.repository.get_entry_last_practiced(entry.id))

        self.repository.record_test_result(entry.id, is_correct=True)
        latest = self.repository.get_entry_last_practiced(entry.id)
        self.assertIsNotNone(latest)

    def test_daily_unique_practice_counts_tracks_unique_entries_per_day(self) -> None:
        created = self.repository.add_entries(
            [
                ("食べる", "たべる", "to eat"),
                ("行く", "いく", "to go"),
            ]
        )

        self.repository.record_test_result(created[0].id, is_correct=True)
        self.repository.record_test_result(created[0].id, is_correct=False)
        self.repository.record_test_result(created[1].id, is_correct=True)

        counts = self.repository.get_daily_unique_practice_counts(days_back=180)
        today_key = date.today().isoformat()
        self.assertIn(today_key, counts)
        self.assertEqual(counts[today_key], 2)

    def test_daily_unique_practice_counts_respects_days_back(self) -> None:
        entry = self.repository.add_entry("話す", "はなす", "to speak")
        old_date = (date.today() - timedelta(days=365)).isoformat()
        today_key = date.today().isoformat()

        connection = sqlite3.connect(self.db_path)
        try:
            connection.execute(
                """
                INSERT INTO practice_daily_unique (entry_id, practice_date)
                VALUES (?, ?)
                """,
                (entry.id, old_date),
            )
            connection.execute(
                """
                INSERT INTO practice_daily_unique (entry_id, practice_date)
                VALUES (?, ?)
                """,
                (entry.id, today_key),
            )
            connection.commit()
        finally:
            connection.close()

        counts = self.repository.get_daily_unique_practice_counts(days_back=180)
        self.assertIn(today_key, counts)
        self.assertNotIn(old_date, counts)

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

    def test_strict_strategy_uses_oldest_last_practiced_as_tiebreaker(self) -> None:
        created = self.repository.add_entries(
            [
                ("犬", "いぬ", "dog"),
                ("猫", "ねこ", "cat"),
            ]
        )
        first_id, second_id = [entry.id for entry in created]

        # Both entries become yellow tier.
        self.repository.record_test_result(first_id, is_correct=False)
        self.repository.record_test_result(second_id, is_correct=False)

        connection = sqlite3.connect(self.db_path)
        try:
            connection.execute(
                """
                UPDATE vocab_stats
                SET last_tested = ?
                WHERE entry_id = ?
                """,
                ("2020-01-01 00:00:00", first_id),
            )
            connection.execute(
                """
                UPDATE vocab_stats
                SET last_tested = ?
                WHERE entry_id = ?
                """,
                ("2024-01-01 00:00:00", second_id),
            )
            connection.commit()
        finally:
            connection.close()

        strict_rows = self.repository.get_test_entries_by_preference(2, strategy="strict")
        strict_ids = [entry.id for entry in strict_rows]
        self.assertEqual(strict_ids, [first_id, second_id])

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
            self.repository.get_english_options_for_entry(9999, max_options=4)

        with self.assertRaises(LookupError):
            self.repository.record_test_result(9999, is_correct=False)

        with self.assertRaises(LookupError):
            self.repository.get_entry_last_practiced(9999)

        with self.assertRaises(LookupError):
            self.repository.increase_priority(9999)

        with self.assertRaises(LookupError):
            self.repository.decrease_priority(9999)


if __name__ == "__main__":
    unittest.main()
