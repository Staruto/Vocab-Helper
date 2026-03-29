# VocabHelper

Simple JP <-> EN vocabulary memorization helper with a desktop GUI.

## Features in v2

- Lists vocabulary in three columns: Japanese writing, kana (hiragana), English meaning
- Adds new entries through a bottom "+" button
- Edits and deletes entries from a right-click context menu
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

## Data location

The app stores entries in a SQLite file at project root:

- vocab.db
