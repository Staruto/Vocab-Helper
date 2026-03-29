from typing import Optional


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


def validate_vocab_fields(japanese_text: str, english_text: str) -> tuple[str, str]:
    japanese = validate_required_text("Japanese writing", japanese_text)
    english = validate_required_text("English meaning", english_text)
    return japanese, english
