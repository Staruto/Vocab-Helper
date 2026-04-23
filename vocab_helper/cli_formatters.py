from __future__ import annotations

from datetime import date, timedelta

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


def _activity_symbol(count: int) -> str:
    if count <= 0:
        return "·"
    if count <= 10:
        return "░"
    if count <= 20:
        return "▒"
    if count <= 30:
        return "▓"
    return "█"


def build_practice_graph_lines(counts_by_date: dict[str, int], days_back: int = 180) -> list[str]:
    range_days = max(int(days_back), 1)
    today = date.today()
    start_date = today - timedelta(days=range_days - 1)
    grid_start = start_date - timedelta(days=start_date.weekday())
    grid_end = today
    weeks = ((grid_end - grid_start).days // 7) + 1

    grid: list[list[str]] = [["·" for _ in range(weeks)] for _ in range(7)]

    current = start_date
    while current <= today:
        week_index = (current - grid_start).days // 7
        weekday_index = current.weekday()
        count = counts_by_date.get(current.isoformat(), 0)
        grid[weekday_index][week_index] = _activity_symbol(count)
        current += timedelta(days=1)

    month_labels = ["  " for _ in range(weeks)]
    current = start_date
    seen_months: set[tuple[int, int]] = set()
    while current <= today:
        month_key = (current.year, current.month)
        if current.day == 1 or current == start_date:
            week_index = (current - grid_start).days // 7
            if month_key not in seen_months and 0 <= week_index < weeks:
                month_labels[week_index] = current.strftime("%b")
                seen_months.add(month_key)
        current += timedelta(days=1)

    month_header = "    " + " ".join(f"{label:>3}" for label in month_labels)
    day_labels = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    day_rows = [f"{day_labels[index]} " + "  ".join(grid[index]) for index in range(7)]

    active_days = sum(1 for value in counts_by_date.values() if value > 0)
    total_practiced = sum(counts_by_date.values())

    summary = f"Unique vocab practiced: {total_practiced} across {active_days} active days"
    legend = "Legend: · 0  ░ 1-10  ▒ 11-20  ▓ 21-30  █ 31+"

    return [month_header, *day_rows, "", summary, legend]
