from __future__ import annotations

PREDEFINED_LANGUAGE_NAMES: dict[str, str] = {
    "JP": "Japanese",
    "EN": "English",
    "ZH": "Chinese",
    "KO": "Korean",
    "ES": "Spanish",
    "FR": "French",
    "DE": "German",
}

PREDEFINED_LANGUAGE_CODES: tuple[str, ...] = tuple(PREDEFINED_LANGUAGE_NAMES.keys())
PREDEFINED_LANGUAGE_CODE_SET: frozenset[str] = frozenset(PREDEFINED_LANGUAGE_CODES)

_COMMON_PREDEFINED_PROPERTIES: tuple[tuple[str, str, bool], ...] = (
    ("target_text", "Target text", True),
    ("meaning", "Meaning", True),
)

PREDEFINED_LANGUAGE_PROPERTY_PROFILES: dict[str, tuple[tuple[str, str, bool], ...]] = {
    "JP": (
        ("target_text", "Target text", True),
        ("meaning", "Meaning", True),
        ("kana", "Kana", False),
    ),
    "EN": _COMMON_PREDEFINED_PROPERTIES,
    "ZH": _COMMON_PREDEFINED_PROPERTIES,
    "KO": _COMMON_PREDEFINED_PROPERTIES,
    "ES": _COMMON_PREDEFINED_PROPERTIES,
    "FR": _COMMON_PREDEFINED_PROPERTIES,
    "DE": _COMMON_PREDEFINED_PROPERTIES,
}