from __future__ import annotations

from dataclasses import dataclass
import shlex
from typing import Callable

from rich.console import Console
from rich.panel import Panel

from .cli_formatters import (
    build_entries_table,
    build_entry_detail_table,
    build_tags_table,
    build_workbooks_table,
)
from .cli_state import CliState
from .db import VocabRepository
from .kana import suggest_hiragana
from .languages import PREDEFINED_LANGUAGE_NAMES
from .models import VocabEntry, Workbook
from .validators import ValidationError


AskInput = Callable[[str, str | None], str]


@dataclass(slots=True)
class CliAction:
    continue_running: bool = True
    launch_gui: bool = False


class CliCommandDispatcher:
    def __init__(self, repository: VocabRepository, state: CliState, console: Console, ask_input: AskInput) -> None:
        self.repository = repository
        self.state = state
        self.console = console
        self.ask_input = ask_input

    def execute(self, line: str) -> CliAction:
        text = line.strip()
        if not text:
            return CliAction()

        if text.startswith("/"):
            text = text[1:]

        try:
            tokens = shlex.split(text)
        except ValueError as exc:
            self.console.print(f"Input error: {exc}", style="bold red")
            return CliAction()

        if not tokens:
            return CliAction()

        command = tokens[0].lower()
        args = tokens[1:]

        try:
            if command in {"help", "?"}:
                self._print_help()
            elif command in {"exit", "quit"}:
                return CliAction(continue_running=False)
            elif command == "gui":
                return CliAction(continue_running=True, launch_gui=True)
            elif command in {"workbook", "wb"}:
                self._handle_workbook(args)
            elif command == "workbooks":
                self._list_workbooks()
            elif command in {"list", "ls"}:
                self._handle_list(args)
            elif command == "filters":
                self._show_filters()
            elif command == "clear-filters":
                self.state.filter_tag_ids = []
                self.state.search_query = ""
                self.state.filter_match_mode = "all"
                self.console.print("Filters cleared.", style="green")
            elif command == "add":
                self._handle_add(args)
            elif command in {"view", "show"}:
                self._handle_view(args)
            elif command == "edit":
                self._handle_edit(args)
            elif command in {"delete", "del", "rm"}:
                self._handle_delete(args)
            elif command == "tags":
                self._handle_tags(args)
            elif command == "test":
                self._handle_test(args)
            else:
                self.console.print("Unknown command. Use /help.", style="bold yellow")
        except (ValidationError, ValueError) as exc:
            self.console.print(f"Validation error: {exc}", style="bold red")
        except LookupError as exc:
            self.console.print(str(exc), style="bold red")

        return CliAction()

    def _print_help(self) -> None:
        lines = [
            "Slash commands:",
            "/help",
            "/exit",
            "/gui",
            "/workbook list|current|switch <id>|create --name NAME --target-code CODE [--preset generic|japanese] [--target-label TEXT] [--meaning-label TEXT]|delete <id>",
            "/list [--sort time|stats|tags] [--time newest|oldest] [--search TEXT] [--tags 1,2] [--match all|any] [--count COUNT]",
            "/filters",
            "/clear-filters",
            "/add [--target TEXT] [--meaning TEXT] [--kana TEXT] [--pos TEXT] [--tags 1,2]",
            "/view <entry_id>",
            "/edit <entry_id> [--target TEXT] [--meaning TEXT] [--kana TEXT] [--pos TEXT] [--tags 1,2]",
            "/delete <entry_id[,entry_id...]>",
            "/tags list|types|add-type --name NAME|delete-type --id ID|add --type-id ID --name NAME|delete --id ID|set --entry-id ID --tag-ids 1,2",
            "/test start [--mode meaning_to_target|target_to_kana|target_to_meaning] [--count N] [--strategy strict|weighted]",
        ]
        self.console.print(Panel("\n".join(lines), title="Commands Menu", border_style="cyan"))

    def _active_workbook(self) -> Workbook:
        if self.state.current_workbook_id is None:
            self.state.current_workbook_id = self.repository.get_current_workbook_id()
        if self.state.current_workbook_id is None:
            raise ValidationError("No workbook selected. Use /workbook create or /workbook switch.")
        workbook = self.repository.get_workbook(self.state.current_workbook_id)
        return workbook

    def _list_workbooks(self) -> None:
        workbooks = self.repository.list_workbooks()
        if not workbooks:
            self.console.print("No workbooks yet. Create one with /workbook create ...", style="yellow")
            return
        self.console.print(build_workbooks_table(workbooks, self.state.current_workbook_id))

    def _handle_workbook(self, args: list[str]) -> None:
        positional, options = self._parse_options(args)
        action = positional[0].lower() if positional else "list"

        if action == "list":
            self._list_workbooks()
            return
        if action == "current":
            workbook = self._active_workbook()
            self.console.print(f"Current workbook: #{workbook.id} {workbook.name}", style="green")
            return
        if action == "switch":
            workbook_id_text = positional[1] if len(positional) > 1 else options.get("id")
            if not workbook_id_text:
                raise ValidationError("Provide workbook id: /workbook switch <id>")
            workbook = self.repository.set_current_workbook_id(int(workbook_id_text))
            self.state.current_workbook_id = workbook.id
            self.console.print(f"Switched to workbook #{workbook.id} {workbook.name}", style="green")
            return
        if action == "create":
            name = options.get("name") or self.ask_input("Workbook name", None)
            target_code = options.get("target_code") or options.get("target") or self.ask_input(
                "Target schema code (for example JP, EN, CUSTOM_01)",
                "JP",
            )
            preset = (options.get("preset") or "generic").strip().lower()
            target_label = options.get("target_label")
            meaning_label = options.get("meaning_label")

            workbook = self.repository.create_workbook(
                name=name,
                target_language_code=target_code,
                preset_key=preset,
                target_label=target_label,
                meaning_label=meaning_label,
            )
            self.repository.set_current_workbook_id(workbook.id)
            self.state.current_workbook_id = workbook.id
            self.console.print(f"Created workbook #{workbook.id} {workbook.name}", style="green")
            return
        if action == "delete":
            workbook_id_text = positional[1] if len(positional) > 1 else options.get("id")
            if not workbook_id_text:
                raise ValidationError("Provide workbook id: /workbook delete <id>")
            new_current = self.repository.delete_workbook(int(workbook_id_text))
            self.state.current_workbook_id = new_current
            if new_current is None:
                self.console.print("Workbook deleted. No current workbook selected.", style="yellow")
            else:
                workbook = self.repository.get_workbook(new_current)
                self.console.print(
                    f"Workbook deleted. Current workbook is now #{workbook.id} {workbook.name}",
                    style="green",
                )
            return

        raise ValidationError(f"Unknown workbook action: {action}")

    def _show_filters(self) -> None:
        filters = [
            f"sort={self.state.sort_mode}",
            f"time={self.state.time_order}",
            f"search={self.state.search_query or '<none>'}",
            f"tags={','.join(str(tag_id) for tag_id in self.state.filter_tag_ids) or '<none>'}",
            f"match={self.state.filter_match_mode}",
        ]
        self.console.print("Active filters: " + " | ".join(filters), style="cyan")

    def _handle_list(self, args: list[str]) -> None:
        workbook = self._active_workbook()
        _positional, options = self._parse_options(args)

        sort_mode = options.get("sort", self.state.sort_mode).strip().lower()
        if sort_mode not in {"time", "stats", "tags"}:
            raise ValidationError("sort must be one of: time, stats, tags")

        time_order = options.get("time", self.state.time_order).strip().lower()
        if time_order not in {"newest", "oldest"}:
            raise ValidationError("time must be one of: newest, oldest")

        search_query = options.get("search", self.state.search_query)
        tags_value = options.get("tags")
        filter_tag_ids = self.state.filter_tag_ids
        if tags_value is not None:
            filter_tag_ids = self._parse_id_list(tags_value)

        filter_match_mode = options.get("match", self.state.filter_match_mode).strip().lower()
        if filter_match_mode not in {"all", "any"}:
            raise ValidationError("match must be one of: all, any")

        count_text = options.get("count")
        count_limit: int | None = None
        if count_text is not None:
            count_limit = int(count_text)
            if count_limit <= 0:
                raise ValidationError("count must be a positive integer")

        self.state.sort_mode = sort_mode
        self.state.time_order = time_order
        self.state.search_query = search_query
        self.state.filter_tag_ids = filter_tag_ids
        self.state.filter_match_mode = filter_match_mode

        rows = self.repository.list_entries_with_stats(
            sort_mode=self.state.sort_mode,
            time_order=self.state.time_order,
            filter_tag_ids=self.state.filter_tag_ids,
            filter_match_mode=self.state.filter_match_mode,
            search_query=self.state.search_query,
            target_language_code=workbook.target_language_code,
            workbook_id=workbook.id,
        )
        if count_limit is not None:
            rows = rows[:count_limit]
        self._show_filters()
        if not rows:
            self.console.print("No entries found.", style="yellow")
            return
        self.console.print(build_entries_table(rows))

    def _handle_add(self, args: list[str]) -> None:
        workbook = self._active_workbook()
        _positional, options = self._parse_options(args)

        target = options.get("target")
        meaning = options.get("meaning")
        kana = options.get("kana", "")
        part_of_speech = options.get("pos", "")

        if not target:
            target = self.ask_input(workbook.target_label, None)
        if not meaning:
            meaning = self.ask_input(workbook.meaning_label, None)

        if not kana and workbook.preset_key == "japanese":
            suggested, reliable, status_message = suggest_hiragana(target)
            if suggested and reliable:
                kana = self.ask_input("Kana", suggested)
            elif status_message:
                self.console.print(status_message, style="yellow")

        entry = self.repository.add_entry(
            japanese_text=target,
            kana_text=kana,
            english_text=meaning,
            part_of_speech=part_of_speech,
            workbook_id=workbook.id,
        )

        if "tags" in options:
            tag_ids = self._parse_id_list(options.get("tags", ""))
            self.repository.set_entry_tags(
                entry.id,
                tag_ids,
                target_language_code=workbook.target_language_code,
                include_part_of_speech=False,
            )

        self.console.print(f"Added entry #{entry.id}", style="green")

    def _handle_view(self, args: list[str]) -> None:
        workbook = self._active_workbook()
        positional, _options = self._parse_options(args)
        if not positional:
            raise ValidationError("Usage: /view <entry_id>")

        entry_id = int(positional[0])
        entry = self.repository.get_entry(entry_id)
        stats = self.repository.get_entry_stats(entry_id)
        tags = self.repository.get_entry_tags(
            entry_id,
            target_language_code=workbook.target_language_code,
            include_part_of_speech=False,
        )
        self.console.print(build_entry_detail_table(entry, stats, tags))

    def _handle_edit(self, args: list[str]) -> None:
        workbook = self._active_workbook()
        positional, options = self._parse_options(args)
        if not positional:
            raise ValidationError("Usage: /edit <entry_id> [--target ... --meaning ...]")

        entry_id = int(positional[0])
        existing = self.repository.get_entry(entry_id)

        target = options.get("target")
        meaning = options.get("meaning")
        kana = options.get("kana")
        part_of_speech = options.get("pos")

        if target is None:
            target = self.ask_input(workbook.target_label, existing.japanese_text)
        if meaning is None:
            meaning = self.ask_input(workbook.meaning_label, existing.english_text)
        if kana is None:
            kana = self.ask_input("Kana", existing.kana_text or "")
        if part_of_speech is None:
            part_of_speech = self.ask_input("Part of speech", existing.part_of_speech or "")

        self.repository.update_entry(
            entry_id=entry_id,
            japanese_text=target,
            kana_text=kana,
            english_text=meaning,
            part_of_speech=part_of_speech,
        )

        if "tags" in options:
            tag_ids = self._parse_id_list(options.get("tags", ""))
            self.repository.set_entry_tags(
                entry_id,
                tag_ids,
                target_language_code=workbook.target_language_code,
                include_part_of_speech=False,
            )

        self.console.print(f"Updated entry #{entry_id}", style="green")

    def _handle_delete(self, args: list[str]) -> None:
        positional, options = self._parse_options(args)
        target = positional[0] if positional else options.get("ids")
        if not target:
            raise ValidationError("Usage: /delete <entry_id[,entry_id...]>")

        entry_ids = self._parse_id_list(target)
        deleted_count = self.repository.delete_entries(entry_ids)
        self.console.print(f"Deleted {deleted_count} entries.", style="green")

    def _handle_tags(self, args: list[str]) -> None:
        workbook = self._active_workbook()
        positional, options = self._parse_options(args)
        action = positional[0].lower() if positional else "list"

        if action == "list":
            tag_type_id = int(options["type_id"]) if "type_id" in options else None
            rows = self.repository.list_tags(
                target_language_code=workbook.target_language_code,
                tag_type_id=tag_type_id,
                include_part_of_speech=False,
            )
            if not rows:
                self.console.print("No tags found.", style="yellow")
                return
            self.console.print(build_tags_table(rows))
            return

        if action == "types":
            rows = self.repository.list_tag_types(target_language_code=workbook.target_language_code)
            if not rows:
                self.console.print("No tag types found.", style="yellow")
                return
            for tag_type_id, name, is_predefined in rows:
                marker = " [predefined]" if is_predefined else ""
                self.console.print(f"{tag_type_id}: {name}{marker}")
            return

        if action == "add-type":
            name = options.get("name") or self.ask_input("Tag type name", None)
            tag_type_id = self.repository.add_tag_type(name, target_language_code=workbook.target_language_code)
            self.console.print(f"Created tag type #{tag_type_id}", style="green")
            return

        if action == "delete-type":
            tag_type_id_text = options.get("id")
            if not tag_type_id_text:
                raise ValidationError("Usage: /tags delete-type --id <type_id>")
            self.repository.delete_tag_type(int(tag_type_id_text))
            self.console.print("Tag type deleted.", style="green")
            return

        if action == "add":
            tag_type_id_text = options.get("type_id")
            if not tag_type_id_text:
                raise ValidationError("Usage: /tags add --type-id <type_id> --name <tag>")
            name = options.get("name") or self.ask_input("Tag name", None)
            tag_id = self.repository.add_tag(int(tag_type_id_text), name)
            self.console.print(f"Created tag #{tag_id}", style="green")
            return

        if action == "delete":
            tag_id_text = options.get("id")
            if not tag_id_text:
                raise ValidationError("Usage: /tags delete --id <tag_id>")
            self.repository.delete_tag(int(tag_id_text))
            self.console.print("Tag deleted.", style="green")
            return

        if action == "set":
            entry_id_text = options.get("entry_id")
            if not entry_id_text:
                raise ValidationError("Usage: /tags set --entry-id <entry_id> --tag-ids 1,2")
            tag_ids = self._parse_id_list(options.get("tag_ids", ""))
            self.repository.set_entry_tags(
                int(entry_id_text),
                tag_ids,
                target_language_code=workbook.target_language_code,
                include_part_of_speech=False,
            )
            self.console.print("Entry tags updated.", style="green")
            return

        raise ValidationError(f"Unknown tags action: {action}")

    def _handle_test(self, args: list[str]) -> None:
        workbook = self._active_workbook()
        positional, options = self._parse_options(args)
        action = positional[0].lower() if positional else "start"
        if action != "start":
            raise ValidationError("Usage: /test start [--mode ... --count ... --strategy ...]")

        mode = options.get("mode", "meaning_to_target").strip().lower()
        strategy = options.get("strategy", "strict").strip().lower()
        count = int(options.get("count", "15"))

        if mode not in {"meaning_to_target", "target_to_kana", "target_to_meaning"}:
            raise ValidationError("mode must be one of: meaning_to_target, target_to_kana, target_to_meaning")
        if strategy not in {"strict", "weighted"}:
            raise ValidationError("strategy must be one of: strict, weighted")
        if count <= 0:
            raise ValidationError("count must be a positive integer")

        candidates = self.repository.get_test_entries_by_preference(count, strategy=strategy, workbook_id=workbook.id)
        if mode == "target_to_kana":
            candidates = [entry for entry in candidates if (entry.kana_text or "").strip()]

        if not candidates:
            self.console.print("No eligible entries for this test.", style="yellow")
            return

        initial_questions = list(candidates)
        retry_questions = list(initial_questions)
        initial_correct = 0
        retry_cycle = 0

        while retry_questions:
            failed_in_cycle: list[VocabEntry] = []
            retry_cycle += 1
            cycle_name = "Initial" if retry_cycle == 1 else f"Retry {retry_cycle - 1}"
            self.console.print(f"{cycle_name} cycle: {len(retry_questions)} questions", style="cyan")

            for index, entry in enumerate(retry_questions, start=1):
                if mode == "meaning_to_target":
                    prompt = f"[{index}/{len(retry_questions)}] Meaning: {entry.english_text}"
                    answer = self.ask_input(prompt, None).strip()
                    expected = entry.japanese_text.strip()
                    is_correct = answer == expected
                    if is_correct:
                        self.console.print("Correct", style="green")
                    else:
                        kana_hint = f" | Kana: {entry.kana_text}" if entry.kana_text else ""
                        self.console.print(f"Incorrect. Answer: {expected}{kana_hint}", style="yellow")
                elif mode == "target_to_kana":
                    prompt = f"[{index}/{len(retry_questions)}] Target: {entry.japanese_text}"
                    answer = self.ask_input(prompt, None).strip()
                    expected = (entry.kana_text or "").strip()
                    is_correct = answer == expected
                    if is_correct:
                        self.console.print("Correct", style="green")
                    else:
                        self.console.print(f"Incorrect. Answer: {expected}", style="yellow")
                else:
                    options_list = self.repository.get_english_options_for_entry(
                        entry.id,
                        max_options=4,
                        workbook_id=workbook.id,
                    )
                    if len(options_list) < 2:
                        self.console.print(
                            f"Skipping entry #{entry.id}; not enough options.",
                            style="yellow",
                        )
                        continue

                    self.console.print(f"[{index}/{len(retry_questions)}] Target: {entry.japanese_text}")
                    for option_index, option in enumerate(options_list, start=1):
                        self.console.print(f"  {option_index}. {option}")
                    selected_text = self.ask_input("Choose option number", None).strip()
                    selected_index = int(selected_text) if selected_text else 0
                    if selected_index < 1 or selected_index > len(options_list):
                        is_correct = False
                    else:
                        is_correct = options_list[selected_index - 1].strip() == entry.english_text.strip()
                    if is_correct:
                        self.console.print("Correct", style="green")
                    else:
                        self.console.print(f"Incorrect. Answer: {entry.english_text}", style="yellow")

                self.repository.record_test_result(entry.id, is_correct)
                if retry_cycle == 1 and is_correct:
                    initial_correct += 1
                if not is_correct:
                    failed_in_cycle.append(entry)

            retry_questions = failed_in_cycle

        self.console.print(
            f"Test complete. Initial score: {initial_correct}/{len(initial_questions)}",
            style="bold green",
        )

    @staticmethod
    def _parse_options(args: list[str]) -> tuple[list[str], dict[str, str]]:
        positional: list[str] = []
        options: dict[str, str] = {}

        index = 0
        while index < len(args):
            token = args[index]
            if token.startswith("--"):
                key = token[2:].strip().replace("-", "_")
                if not key:
                    index += 1
                    continue
                if index + 1 < len(args) and not args[index + 1].startswith("--"):
                    options[key] = args[index + 1]
                    index += 2
                    continue
                options[key] = ""
                index += 1
                continue

            positional.append(token)
            index += 1

        return positional, options

    @staticmethod
    def _parse_id_list(value: str) -> list[int]:
        cleaned = value.strip()
        if not cleaned:
            return []
        parts = [part.strip() for part in cleaned.split(",")]
        result: list[int] = []
        for part in parts:
            if not part:
                continue
            result.append(int(part))
        return sorted(set(result))


def short_language_name(workbook: Workbook) -> str:
    key = workbook.target_language_code.upper().strip()
    if key in PREDEFINED_LANGUAGE_NAMES:
        return PREDEFINED_LANGUAGE_NAMES[key]
    return key
