from __future__ import annotations

import re
from typing import Optional

try:
    from pykakasi import kakasi
except ImportError:  # pragma: no cover - depends on optional runtime package
    kakasi = None


_JAPANESE_CHARS_PATTERN = re.compile(r"[\u3040-\u30ff\u4e00-\u9fff]")
_HIRAGANA_PATTERN = re.compile(r"[\u3041-\u3096]")


def suggest_hiragana(japanese_text: str) -> tuple[Optional[str], bool, str]:
    cleaned = japanese_text.strip()
    if not cleaned:
        return None, False, "Enter Japanese writing to suggest kana."

    if kakasi is None:
        return None, False, "Kana suggestion unavailable. Install pykakasi to enable it."

    try:
        suggestion = _convert_to_hiragana(cleaned)
    except Exception:
        return None, False, "Kana suggestion failed. Please enter kana manually."

    if not suggestion:
        return None, False, "No kana suggestion was produced. Please enter kana manually."

    reliable = _looks_reliable(cleaned, suggestion)
    if not reliable:
        return None, False, "Kana suggestion is uncertain. Please enter kana manually."

    return suggestion, True, "Kana suggested. Please confirm or edit before saving."


def _convert_to_hiragana(text: str) -> str:
    converter = kakasi()

    if hasattr(converter, "convert"):
        parts = converter.convert(text)
        joined = "".join(_part_to_hiragana(part) for part in parts)
        return _katakana_to_hiragana(joined)

    converter.setMode("J", "H")
    converter.setMode("K", "H")
    converter.setMode("H", "H")
    legacy_converter = converter.getConverter()
    return legacy_converter.do(text)


def _part_to_hiragana(part: dict[str, str]) -> str:
    hira = part.get("hira")
    if hira:
        return hira

    kana = part.get("kana")
    if kana:
        return _katakana_to_hiragana(kana)

    return part.get("orig", "")


def _katakana_to_hiragana(text: str) -> str:
    chars = []
    for char in text:
        code = ord(char)
        if 0x30A1 <= code <= 0x30F6:
            chars.append(chr(code - 0x60))
        else:
            chars.append(char)
    return "".join(chars)


def _looks_reliable(source_text: str, suggested_hiragana: str) -> bool:
    has_japanese = bool(_JAPANESE_CHARS_PATTERN.search(source_text))
    has_hiragana = bool(_HIRAGANA_PATTERN.search(suggested_hiragana))
    return has_japanese and has_hiragana
