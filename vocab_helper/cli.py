from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Protocol

from rich.console import Console
from rich.panel import Panel

from .cli_commands import CliAction, CliCommandDispatcher
from .cli_state import CliState
from .db import VocabRepository, default_db_path

try:  # pragma: no cover - optional import guard
    from prompt_toolkit import PromptSession
    from prompt_toolkit.history import FileHistory
    from prompt_toolkit.styles import Style
except ImportError:  # pragma: no cover - runtime fallback
    PromptSession = None
    FileHistory = None
    Style = None


class PromptAdapter(Protocol):
    def prompt_command(self) -> str:
        ...

    def prompt_field(self, label: str, default: str = "") -> str:
        ...


@dataclass(slots=True)
class BasicPromptAdapter:
    PROMPT_COLOR = "\x1b[93m"
    RESET_COLOR = "\x1b[0m"

    def prompt_command(self) -> str:
        raw = input(f"{self.PROMPT_COLOR}> ").strip()
        print(self.RESET_COLOR, end="")
        return raw

    def prompt_field(self, label: str, default: str = "") -> str:
        if default:
            raw = input(f"{label} [{default}]: ").strip()
            return raw or default
        return input(f"{label}: ").strip()


class ToolkitPromptAdapter:
    def __init__(self, history_path: Path) -> None:
        if PromptSession is None:
            raise RuntimeError("prompt_toolkit is not available")

        style = Style.from_dict(
            {
                "": "#ffff87",
                "prompt": "bold #ffff87",
            }
        )

        history = FileHistory(str(history_path))
        self._session = PromptSession(
            history=history,
            style=style,
        )

    def prompt_command(self) -> str:
        return self._session.prompt([("class:prompt", "> ")], default="").strip()

    def prompt_field(self, label: str, default: str = "") -> str:
        raw = self._session.prompt([("class:prompt", f"{label}: ")], default=default)
        cleaned = raw.strip()
        if not cleaned and default:
            return default
        return cleaned


def _build_prompt_adapter(console: Console) -> PromptAdapter:
    history_path = Path.home() / ".vocab_helper_cli_history"
    try:
        return ToolkitPromptAdapter(history_path=history_path)
    except Exception:
        console.print(
            "prompt_toolkit unavailable, using basic input mode.",
            style="yellow",
        )
        return BasicPromptAdapter()


def _launch_gui(console: Console) -> None:
    from .app import main as app_main

    console.print("Opening GUI. Close the window to return to CLI.", style="cyan")
    app_main()
    console.print("Returned to CLI.", style="cyan")


def run_cli(
    repository: VocabRepository,
    command: str | None = None,
    show_banner: bool = True,
) -> CliAction:
    console = Console()
    state = CliState(current_workbook_id=repository.get_current_workbook_id())
    prompt_adapter = _build_prompt_adapter(console)

    def ask_input(label: str, default: str | None) -> str:
        return prompt_adapter.prompt_field(label, default or "")

    dispatcher = CliCommandDispatcher(repository, state, console, ask_input)

    if show_banner:
        banner = (
            "Interactive slash commands for vocabulary practice.\n"
            "Use /help to see commands. Use /gui to launch desktop mode."
        )
        console.print(Panel(banner, title="VocabHelper CLI", border_style="cyan"))

    if command is not None:
        action = dispatcher.execute(command)
        if action.launch_gui:
            _launch_gui(console)
        return action

    while True:
        raw = prompt_adapter.prompt_command()
        action = dispatcher.execute(raw)
        if action.launch_gui:
            _launch_gui(console)
            if action.continue_running:
                continue
            return CliAction(continue_running=False)
        if not action.continue_running:
            return action


def main(command: str | None = None, show_banner: bool = True) -> int:
    repository = VocabRepository(default_db_path())
    repository.initialize()

    run_cli(repository=repository, command=command, show_banner=show_banner)
    return 0
