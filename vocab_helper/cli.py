from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Protocol

from rich.console import Console
from rich.panel import Panel

from .cli_commands import CliAction, CliCommandDispatcher, short_language_name
from .cli_state import CliState
from .db import VocabRepository, default_db_path

try:  # pragma: no cover - optional import guard
    from prompt_toolkit import PromptSession
    from prompt_toolkit.auto_suggest import AutoSuggestFromHistory
    from prompt_toolkit.completion import WordCompleter
    from prompt_toolkit.history import FileHistory
except ImportError:  # pragma: no cover - runtime fallback
    PromptSession = None
    AutoSuggestFromHistory = None
    WordCompleter = None
    FileHistory = None


class PromptAdapter(Protocol):
    def prompt(self, text: str, default: str = "") -> str:
        ...


@dataclass(slots=True)
class BasicPromptAdapter:
    def prompt(self, text: str, default: str = "") -> str:
        if default:
            raw = input(f"{text} [{default}]: ").strip()
            return raw or default
        return input(f"{text}: ").strip()


class ToolkitPromptAdapter:
    def __init__(self, history_path: Path, completer_words: list[str]) -> None:
        if PromptSession is None:
            raise RuntimeError("prompt_toolkit is not available")
        completer = WordCompleter(completer_words, ignore_case=True)
        history = FileHistory(str(history_path))
        self._session = PromptSession(
            completer=completer,
            complete_while_typing=True,
            auto_suggest=AutoSuggestFromHistory(),
            history=history,
        )

    def prompt(self, text: str, default: str = "") -> str:
        return self._session.prompt(f"{text}> ", default=default)


def _command_words() -> list[str]:
    return [
        "/help",
        "/exit",
        "/quit",
        "/gui",
        "/workbook",
        "/workbooks",
        "/list",
        "/filters",
        "/clear-filters",
        "/add",
        "/view",
        "/edit",
        "/delete",
        "/tags",
        "/test",
    ]


def _build_prompt_adapter(console: Console) -> PromptAdapter:
    history_path = Path.home() / ".vocab_helper_cli_history"
    try:
        return ToolkitPromptAdapter(history_path=history_path, completer_words=_command_words())
    except Exception:
        console.print(
            "prompt_toolkit unavailable, using basic input mode. Install dependencies for richer CLI features.",
            style="yellow",
        )
        return BasicPromptAdapter()


def _build_shell_prompt(state: CliState, repository: VocabRepository) -> str:
    workbook_id = state.current_workbook_id
    if workbook_id is None:
        workbook_id = repository.get_current_workbook_id()
        state.current_workbook_id = workbook_id

    if workbook_id is None:
        return "vocab/no-workbook"

    try:
        workbook = repository.get_workbook(workbook_id)
    except LookupError:
        state.current_workbook_id = repository.get_current_workbook_id()
        return "vocab/no-workbook"

    language_name = short_language_name(workbook)
    return f"vocab/{workbook.name}/{language_name}"


def run_cli(
    repository: VocabRepository,
    command: str | None = None,
    show_banner: bool = True,
) -> CliAction:
    console = Console()
    state = CliState(current_workbook_id=repository.get_current_workbook_id())
    prompt_adapter = _build_prompt_adapter(console)

    def ask_input(label: str, default: str | None) -> str:
        return prompt_adapter.prompt(label, default or "")

    dispatcher = CliCommandDispatcher(repository, state, console, ask_input)

    if show_banner:
        banner = (
            "Interactive slash commands for vocabulary practice.\n"
            "Use /help to see commands. Use /gui to launch desktop mode."
        )
        console.print(Panel(banner, title="VocabHelper CLI", border_style="cyan"))

    if command is not None:
        return dispatcher.execute(command)

    while True:
        prompt_label = _build_shell_prompt(state, repository)
        raw = prompt_adapter.prompt(prompt_label, "")
        action = dispatcher.execute(raw)
        if not action.continue_running or action.launch_gui:
            return action


def main(command: str | None = None, show_banner: bool = True) -> int:
    repository = VocabRepository(default_db_path())
    repository.initialize()

    action = run_cli(repository=repository, command=command, show_banner=show_banner)
    if action.launch_gui:
        from .app import main as app_main

        app_main()
    return 0
