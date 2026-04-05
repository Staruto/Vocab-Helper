from dataclasses import dataclass
from typing import Optional


@dataclass(frozen=True, slots=True)
class VocabEntry:
    id: int
    japanese_text: str
    kana_text: Optional[str]
    english_text: str
    part_of_speech: Optional[str]
    details_markdown: Optional[str]
    created_at: str


@dataclass(frozen=True, slots=True)
class Workbook:
    id: int
    name: str
    target_language_code: str
    preset_key: str
    target_label: str
    meaning_label: str
    created_at: str
