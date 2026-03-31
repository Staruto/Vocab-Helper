from typing import Optional


SUPPORTED_LANGUAGE_CODES = {"JP", "EN"}


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
    japanese = validate_required_text("Japanese writing", japanese_text)
    english = validate_required_text("English meaning", english_text)
    return japanese, english


def validate_language_code(value: str, field_name: str = "Language") -> str:
    code = value.strip().upper()
    if code not in SUPPORTED_LANGUAGE_CODES:
        allowed = ", ".join(sorted(SUPPORTED_LANGUAGE_CODES))
        raise ValidationError(f"{field_name} must be one of: {allowed}.")
    return code
