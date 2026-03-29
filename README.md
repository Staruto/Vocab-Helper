# VocabHelper

Simple JP <-> EN vocabulary memorization helper with a desktop GUI.

## Features in v2

- Lists vocabulary in three columns: Japanese writing, kana (hiragana), English meaning
- Adds new entries through a bottom "+" button
- Supports bulk add in a dedicated 3-column line editor (Japanese, kana, English)
- Edits and deletes entries from a right-click context menu
- Supports multi-select delete from the vocabulary list
- Supports keyboard shortcuts: Ctrl+N (add), Ctrl+Shift+N (bulk add), Ctrl+T (EN->JP test), Enter (edit selected row), Delete (remove selected rows)
- Includes test mode (English -> Japanese) with immediate per-question judgment
- Requires Japanese writing and English meaning
- Treats kana as optional
- Suggests kana offline using pykakasi and lets users edit before save
- Uses larger UI typography for readability (base size 12)
- Prefers Yu Gothic for Japanese text when available
- Shows a global count at the bottom: total number of vocabularies
- Stores data locally in SQLite

## Requirements

- Python 3.10+
- Windows, macOS, or Linux with Tk support

## Setup

1. Create and activate a virtual environment.
2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Run

```bash
python -m vocab_helper
```

## Test

```bash
python -m unittest discover -s tests -p "test_*.py"
```

## Bulk add input format

- Three columns from left to right: Japanese writing, kana, English meaning
- One entry per line (same row index across the three columns)
- Japanese and English are required, kana is optional
- A row where all three fields are empty is ignored

## Test mode (English -> Japanese)

- Click `Test EN->JP` (or press `Ctrl+T`) to open the test dialog
- Default questions per test: 15
- You can set a custom positive integer for questions per test
- Questions are picked randomly from the database
- If requested count is larger than available vocabularies, the test uses all available entries
- Given an English meaning, type the Japanese writing and submit
- Judgement rule: exact Japanese-writing match after trimming surrounding spaces

## Data location

The app stores entries in a SQLite file at project root:

- vocab.db
