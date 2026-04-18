from __future__ import annotations

from rich.table import Table

from .models import VocabEntry, Workbook


def build_workbooks_table(workbooks: list[Workbook], current_workbook_id: int | None) -> Table:
    table = Table(title="Workbooks")
    table.add_column("Current", justify="center")
    table.add_column("ID", justify="right")
    table.add_column("Name")
    table.add_column("Target schema")
    table.add_column("Preset")
    table.add_column("Target label")
    table.add_column("Meaning label")

    for workbook in workbooks:
        is_current = "*" if workbook.id == current_workbook_id else ""
        table.add_row(
            is_current,
            str(workbook.id),
            workbook.name,
            workbook.target_language_code,
            workbook.preset_key,
            workbook.target_label,
            workbook.meaning_label,
        )
    return table


def build_entries_table(rows: list[tuple[VocabEntry, int, int, str]]) -> Table:
    table = Table(title="Vocabulary entries")
    table.add_column("ID", justify="right")
    table.add_column("Target")
    table.add_column("Kana")
    table.add_column("Meaning")
    table.add_column("Part of speech")
    table.add_column("Tests", justify="right")
    table.add_column("Errors", justify="right")
    table.add_column("Tier")

    for entry, test_count, error_count, tier in rows:
        table.add_row(
            str(entry.id),
            entry.japanese_text,
            entry.kana_text or "",
            entry.english_text,
            entry.part_of_speech or "",
            str(test_count),
            str(error_count),
            tier,
        )

    return table


def build_tags_table(rows: list[tuple[int, int, str, str, bool, bool]]) -> Table:
    table = Table(title="Tags")
    table.add_column("Tag ID", justify="right")
    table.add_column("Type ID", justify="right")
    table.add_column("Type")
    table.add_column("Tag")
    table.add_column("Predefined type", justify="center")
    table.add_column("Predefined tag", justify="center")

    for tag_id, tag_type_id, type_name, tag_name, type_predefined, tag_predefined in rows:
        table.add_row(
            str(tag_id),
            str(tag_type_id),
            type_name,
            tag_name,
            "yes" if type_predefined else "",
            "yes" if tag_predefined else "",
        )
    return table


def build_entry_detail_table(entry: VocabEntry, stats: tuple[int, int, str], tags: list[tuple[int, int, str, str]]) -> Table:
    table = Table(title=f"Entry #{entry.id}")
    table.add_column("Field")
    table.add_column("Value")

    table.add_row("Target", entry.japanese_text)
    table.add_row("Kana", entry.kana_text or "")
    table.add_row("Meaning", entry.english_text)
    table.add_row("Part of speech", entry.part_of_speech or "")
    table.add_row("Created", entry.created_at)
    table.add_row("Tier", stats[2])
    table.add_row("Tests", str(stats[0]))
    table.add_row("Errors", str(stats[1]))
    if tags:
        tag_text = ", ".join(f"{type_name}:{tag_name}" for _tag_id, _type_id, type_name, tag_name in tags)
    else:
        tag_text = ""
    table.add_row("Tags", tag_text)
    if entry.details_markdown:
        table.add_row("Details", entry.details_markdown)
    return table
