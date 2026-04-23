import io
import tempfile
import unittest
from pathlib import Path

from rich.console import Console

from vocab_helper.cli_commands import CliCommandDispatcher
from vocab_helper.cli_state import CliState
from vocab_helper.db import VocabRepository


class CliCommandTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.temp_dir.name) / "test_vocab.db"

        self.repository = VocabRepository(self.db_path)
        self.repository.initialize()
        self.workbook = self.repository.create_workbook("JP", "JP", preset_key="japanese")
        self.repository.set_current_workbook_id(self.workbook.id)

        self.state = CliState(current_workbook_id=self.workbook.id)
        self.output = io.StringIO()
        self.console = Console(file=self.output, force_terminal=False, width=120)
        self.answers: list[str] = []

        def ask_input(_prompt: str, default: str | None) -> str:
            if self.answers:
                return self.answers.pop(0)
            return default or ""

        self.dispatcher = CliCommandDispatcher(
            repository=self.repository,
            state=self.state,
            console=self.console,
            ask_input=ask_input,
        )

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_add_and_list_entry(self) -> None:
        self.dispatcher.execute('/add --target "食べる" --kana "たべる" --meaning "to eat"')

        entries = self.repository.list_entries(workbook_id=self.workbook.id)
        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0].japanese_text, "食べる")
        self.assertEqual(entries[0].kana_text, "たべる")
        self.assertEqual(entries[0].english_text, "to eat")

        self.dispatcher.execute("/list")
        rendered = self.output.getvalue()
        self.assertIn("Vocabulary entries", rendered)
        self.assertIn("to eat", rendered)

    def test_workbook_create_switches_current(self) -> None:
        self.dispatcher.execute('/workbook create --name "English" --target-code EN --preset generic')

        current_id = self.state.current_workbook_id
        self.assertIsNotNone(current_id)
        workbook = self.repository.get_workbook(int(current_id))
        self.assertEqual(workbook.name, "English")
        self.assertEqual(workbook.target_language_code, "EN")

    def test_list_with_count_limits_rows(self) -> None:
        self.dispatcher.execute('/add --target "食べる" --kana "たべる" --meaning "to eat"')
        self.dispatcher.execute('/add --target "見る" --kana "みる" --meaning "to see"')

        self.dispatcher.execute("/list --count 1")
        rendered = self.output.getvalue()
        self.assertIn("to see", rendered)
        self.assertNotIn("to eat", rendered)

    def test_list_with_invalid_count_shows_validation_error(self) -> None:
        self.dispatcher.execute("/list --count 0")

        rendered = self.output.getvalue()
        self.assertIn("Validation error", rendered)
        self.assertIn("count must be a positive integer", rendered)

    def test_stats_outputs_practice_graph(self) -> None:
        entry = self.repository.add_entry(
            japanese_text="行く",
            kana_text="いく",
            english_text="to go",
            workbook_id=self.workbook.id,
        )
        self.repository.record_test_result(entry.id, is_correct=True)

        self.dispatcher.execute("/stats")
        rendered = self.output.getvalue()
        self.assertIn("Practice Activity", rendered)
        self.assertIn("Legend:", rendered)

    def test_help_summary_is_concise(self) -> None:
        self.dispatcher.execute("/help")

        rendered = self.output.getvalue()
        self.assertIn("Use /help <command> to see syntax and parameters.", rendered)
        self.assertIn("/stats  Show practice activity graph in terminal.", rendered)

    def test_help_list_shows_detail(self) -> None:
        self.dispatcher.execute("/help list")

        rendered = self.output.getvalue()
        self.assertIn("Help: list", rendered)
        self.assertIn("--count COUNT", rendered)

    def test_gui_action_keeps_cli_running(self) -> None:
        action = self.dispatcher.execute("/gui")

        self.assertTrue(action.continue_running)
        self.assertTrue(action.launch_gui)

    def test_test_mode_records_result(self) -> None:
        entry = self.repository.add_entry(
            japanese_text="食べる",
            kana_text="たべる",
            english_text="to eat",
            workbook_id=self.workbook.id,
        )

        self.answers = ["食べる"]
        self.dispatcher.execute("/test start --mode meaning_to_target --count 1 --strategy strict")

        test_count, error_count, tier = self.repository.get_entry_stats(entry.id)
        self.assertEqual(test_count, 1)
        self.assertEqual(error_count, 0)
        self.assertEqual(tier, "green")


if __name__ == "__main__":
    unittest.main()
