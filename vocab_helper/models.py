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
