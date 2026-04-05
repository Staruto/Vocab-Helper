import unittest

from vocab_helper.validators import (
    ValidationError,
    normalize_optional_text,
    validate_target_schema_code,
    validate_vocab_fields,
)


class ValidatorsTests(unittest.TestCase):
    def test_required_fields_are_trimmed(self) -> None:
        japanese, english = validate_vocab_fields("  日本語  ", "  Japanese language  ")
        self.assertEqual(japanese, "日本語")
        self.assertEqual(english, "Japanese language")

    def test_required_japanese_rejected_when_empty(self) -> None:
        with self.assertRaises(ValidationError):
            validate_vocab_fields("   ", "Valid")

    def test_required_english_rejected_when_empty(self) -> None:
        with self.assertRaises(ValidationError):
            validate_vocab_fields("単語", "   ")

    def test_optional_text_normalizes_blank_to_none(self) -> None:
        self.assertIsNone(normalize_optional_text("   "))
        self.assertEqual(normalize_optional_text("  かな  "), "かな")

    def test_validate_target_schema_code_accepts_non_language_keys(self) -> None:
        self.assertEqual(validate_target_schema_code(" custom_hash_001 "), "CUSTOM_HASH_001")

    def test_validate_target_schema_code_rejects_blank(self) -> None:
        with self.assertRaises(ValidationError):
            validate_target_schema_code("   ")


if __name__ == "__main__":
    unittest.main()
