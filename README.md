# VocabHelper

Simple JP <-> EN vocabulary memorization helper with a desktop GUI.

## Features

- Lists vocabulary in three columns: target text, kana (optional), assistant meaning
- Supports global language settings for target and assistant languages (default: target JP, assistant EN)
- Adds new entries through a bottom "+" button
- Supports bulk add in a dedicated 3-column line editor (Japanese, kana, English)
- Edits and deletes entries from a right-click context menu
- Supports opening a full vocabulary detail page via double-click or right-click -> View details
- Supports multi-select delete from the vocabulary list
- Supports keyboard shortcuts: Ctrl+N (add), Ctrl+Shift+N (bulk add), Ctrl+T (EN->JP test), Enter (edit selected row), Delete (remove selected rows)
- Includes three test modes with immediate per-question judgment:
	- English -> Japanese (fill-in)
	- Japanese -> Kana (fill-in)
	- Japanese -> English (single-choice)
- Tracks per-entry test stats (test count and error count)
- Classifies entries into priority tiers: gray, green, yellow, red
- Supports tier color highlighting in the list view (enabled by default, toggleable)
- Supports list sorting by creation time or by tier-based stats priority
- Supports creation-time order selection (newest first or oldest first)
- Supports test pick preference strategy selection: strict or weighted
- Supports manual priority increase/decrease from the context menu (except gray tier)
- Refreshes the main list immediately after closing a test dialog
- Supports optional part of speech metadata for each vocabulary
- Supports editable markdown details for each vocabulary with in-app display mode
- Requires target text and assistant meaning
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

## Language settings

- Click `Languages` in the settings row to configure global language roles
- Default values:
	- Target language: `JP`
	- Assistant language: `EN`
- Target and assistant languages must be different
- List column headers and entry/detail forms update to match current language settings

## Entry metadata and detail page

- Add/Edit dialogs support optional `Part of speech`
- Open detail page by:
	- Double-clicking a vocabulary row, or
	- Right-click row -> `View details`
- Detail page includes:
	- Editable top fields (target text, kana, assistant meaning, part of speech)
	- Stats summary (tests, errors, tier, created time)
	- Markdown details section with `Edit markdown` and display mode
- Markdown display mode supports a basic subset for readability:
	- Headings (`#`, `##`, `###`)
	- Bullet lists (`-` / `*`)
	- Bold/italic/code inline styling
	- Code blocks fenced with triple backticks

## Test mode (English -> Japanese)

- Click `Test EN->JP` (or press `Ctrl+T`) to open the test dialog
- Default questions per test: 15
- You can set a custom positive integer for questions per test
- You can choose pick preference strategy:
	- `strict` (default): pick by tier order gray -> red -> yellow -> green
	- `weighted`: probabilistic pick favoring higher-priority tiers while still sampling all tiers
- If requested count is larger than available vocabularies, the test uses all available entries
- Given an English meaning, type the Japanese writing and submit
- Judgement rule: exact Japanese-writing match after trimming surrounding spaces
- For incorrect answers, feedback shows the correct Japanese writing and kana when available
- Each submitted answer updates stats:
	- `test_count` always increases by 1
	- `error_count` increases by 1 only when the answer is incorrect

## Test mode (Japanese -> Kana)

- Click `Test JP->Kana` to open the test dialog
- Uses the same settings and behavior style as EN->JP:
	- Default questions per test: 15
	- Positive-integer question count validation
	- Same pick preference strategy (`strict` or `weighted`)
	- Same score/result flow
- Questions only use entries that have kana
- Prompt is Japanese writing; answer is kana
- Judgement rule: exact kana match after trimming surrounding spaces
- Each submitted answer updates stats with the same rule as EN->JP

## Test mode (Japanese -> English)

- Click `Test JP->EN` to open the test dialog
- Prompt is Japanese writing; answer format is single choice
- For each question, the app builds options as:
	- Correct English meaning
	- Up to 3 random distractor meanings
- Option fallback behavior:
	- Shows as many options as available when fewer than 4 distinct options exist
	- Requires at least 2 options; otherwise the test cannot start
- Uses the same pick preference strategy (`strict` or `weighted`) and same stats update rule

## Priority tiers

- `gray`: never tested yet (`test_count = 0`)
- `green`: tested and currently no errors (`test_count > 0` and `error_count = 0`)
- `yellow`: medium error level (`error_count` is 1 or 2)
- `red`: high error level (`error_count >= 3`)

Manual priority actions adjust `error_count` thresholds (for non-gray entries):

- Increase priority: green -> yellow, yellow -> red, red stays red
- Decrease priority: red -> yellow, yellow -> green, green stays green

## Data location

The app stores entries in a SQLite file at project root:

- vocab.db
