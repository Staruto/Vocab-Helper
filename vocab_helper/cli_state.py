from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(slots=True)
class CliState:
    current_workbook_id: int | None = None
    sort_mode: str = "time"
    time_order: str = "newest"
    filter_tag_ids: list[int] = field(default_factory=list)
    filter_match_mode: str = "all"
    search_query: str = ""
