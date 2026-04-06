from typing import Optional

from .languages import PREDEFINED_LANGUAGE_CODE_SET

SUPPORTED_LANGUAGE_CODES = set(PREDEFINED_LANGUAGE_CODE_SET)


class ValidationError(ValueError):
    """Raised when one or more input fields are invalid."""


def validate_required_text(field_name: str, value: str) -> str:
    cleaned = value.strip()
    if not cleaned:
        raise ValidationError(f"{field_name} is required.")
    return cleaned


def normalize_optional_text(value: str) -> Optional[str]:
    cleaned = value.strip()
    return cleaned or None


def normalize_optional_markdown(value: str) -> Optional[str]:
    if value.strip() == "":
        return None
    return value.rstrip()


def validate_vocab_fields(japanese_text: str, english_text: str) -> tuple[str, str]:
    target_text = validate_required_text("Target text", japanese_text)
    meaning_text = validate_required_text("Meaning", english_text)
    return target_text, meaning_text


def validate_language_code(value: str, field_name: str = "Language") -> str:
    code = value.strip().upper()
    if code not in SUPPORTED_LANGUAGE_CODES:
        allowed = ", ".join(sorted(SUPPORTED_LANGUAGE_CODES))
        raise ValidationError(f"{field_name} must be one of: {allowed}.")
    return code


def validate_target_schema_code(value: str, field_name: str = "Target schema") -> str:
    code = value.strip().upper()
    if not code:
        raise ValidationError(f"{field_name} is required.")
    return code
