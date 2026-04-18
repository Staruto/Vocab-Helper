import unittest
from unittest.mock import patch

from vocab_helper import __main__ as entrypoint


class EntrypointTests(unittest.TestCase):
    def test_default_mode_is_cli(self) -> None:
        with patch("vocab_helper.cli.main", return_value=7) as cli_main:
            result = entrypoint.main([])

        self.assertEqual(result, 7)
        cli_main.assert_called_once_with(command=None, show_banner=True)

    def test_cli_mode_honors_command_and_banner_flags(self) -> None:
        with patch("vocab_helper.cli.main", return_value=0) as cli_main:
            result = entrypoint.main(["cli", "--command", "/help", "--no-banner"])

        self.assertEqual(result, 0)
        cli_main.assert_called_once_with(command="/help", show_banner=False)

    def test_gui_mode_launches_gui(self) -> None:
        with patch("vocab_helper.app.main") as app_main:
            result = entrypoint.main(["gui"])

        self.assertEqual(result, 0)
        app_main.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
