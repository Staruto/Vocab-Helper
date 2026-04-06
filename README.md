# VocabHelper

Simple vocabulary memorization helper with a desktop GUI.

## Features

- Lists vocabulary in two columns: target text, assistant meaning
- Supports workbook-specific target/meaning labels so each workbook can represent any text-to-text mapping
- Adds new entries through a bottom "+" button
- Supports bulk add in a dedicated line editor with dynamic columns (2 columns by default, 3 when kana is enabled)
- Edits and deletes entries from a right-click context menu
- Supports opening a full vocabulary detail page via double-click or right-click -> View details
- Supports multi-select delete from the vocabulary list
- Splits the UI into tabs:
	- `Home` for vocabularies
	- `Profile` for the practice activity grid
- Supports keyboard shortcuts: Ctrl+N (add), Ctrl+Shift+N (bulk add), Ctrl+T (meaning->target test), Enter (edit selected row), Delete (remove selected rows)
- Includes three test modes with immediate per-question judgment:
	- Meaning -> Target (fill-in)
	- Target -> Kana (fill-in, when kana is enabled)
	- Target -> Meaning (single-choice)
- Tracks per-entry test stats (test count and error count)
- Classifies entries into priority tiers: gray, green, yellow, red
- Supports tier color highlighting in the list view (enabled by default, toggleable)
- Supports list sorting by creation time or by tier-based stats priority
- Supports list sorting by tags for the current target-language tag scope
- Supports creation-time order selection (newest first or oldest first)
- Supports test pick preference strategy selection: strict or weighted
- Supports manual priority increase/decrease from the context menu (except gray tier)
- Refreshes the main list immediately after closing a test dialog
- Supports optional part of speech metadata for each vocabulary
- Supports typed tags for vocabularies (user-defined types and tags)
- Includes predefined tag types: `part_of_speech` and `difficulty` (`N5`, `N4`, `N3`, `N2`, `N1`)
- Supports per-target-language tag scopes (each target language keeps its own tag catalog)
- Supports tag filtering in Home with ALL-match semantics across selected tags
- Supports live search in Home across target text and meaning (case-insensitive contains)
- Supports editable markdown details for each vocabulary with in-app display mode
- Tracks latest practice date per vocabulary
- Shows a GitHub-style contributions grid for daily practice activity (last 180 days)
- Requires target text and meaning text for every entry
- Treats kana as optional
- Suggests kana offline using pykakasi when the workbook preset enables kana
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

- Two columns by default: target label and meaning label
- Three columns when kana is enabled: target label, kana, meaning label
- One entry per line (same row index across the three columns)
- Target and meaning columns are required, kana is optional
- A row where all three fields are empty is ignored

## Workbook creation and presets

- Workbook creation supports target-side configuration as either:
	- A supported language (`JP`, `EN`, `ZH`, `KO`, `ES`, `FR`, `DE`) with automatic label, or
	- A custom target label text
- Meaning label supports:
	- Default `Meaning`
	- Supported language label (`JP`, `EN`, `ZH`, `KO`, `ES`, `FR`, `DE`)
	- Custom text label
- Preset selection is a boolean toggle (`Enable preset`) instead of choosing a preset name
- Preset controls are shown only when a preset is available for the selected supported target language
- Current preset support:
	- `JP` preset adds kana behavior and predefined difficulty tags
	- Other supported languages currently use generic behavior (no additional language-specific preset)
- Workbook labels can be edited later from `Settings` -> `Workbook columns` -> `Edit labels`

## Language settings

- Click `Languages` in the settings row to configure global language roles
- Default values:
	- Target language: `JP`
	- Assistant language: `EN`
- Target and assistant languages must be different
- List column headers and entry/detail forms update to match current language settings
- Tag catalogs are scoped by target language

## Tags and filtering

- Click `Tags` in the Home settings row to open Tag Manager
- Tag Manager supports:
	- Creating/deleting custom tag types
	- Creating/deleting tags under each type
	- Protecting predefined types/tags from deletion
- Predefined types:
	- `part_of_speech`: noun, verb, adjective, adverb, expression, particle, auxiliary, other
	- `difficulty`: N5, N4, N3, N2, N1
- Entry Add/Edit and Detail pages support selecting multiple tags
- `Part of speech` field stays available and synchronizes to `part_of_speech` tags automatically
- Use `Filter tags` in Home to select filter tags
- Filter behavior is `ALL`: a vocabulary must contain every selected tag to be shown
- `Clear filter` removes all active tag filtering
- Use `Search` in Home to filter by target text or meaning with case-insensitive partial matching
- Search is applied live while typing and is combined with active tag filters

## Entry metadata and detail page

- Add/Edit dialogs support optional `Part of speech`
- Add/Edit dialogs support optional multi-tag assignment
- Open detail page by:
	- Double-clicking a vocabulary row, or
	- Right-click row -> `View details`
- Detail page includes:
	- Editable top fields (target text, kana, assistant meaning, part of speech, tags)
	- Stats summary (tests, errors, tier, created time)
	- Latest practice date
	- Markdown details section with `Edit markdown` and display mode
- Markdown display mode supports a basic subset for readability:
	- Headings (`#`, `##`, `###`)
	- Bullet lists (`-` / `*`)
	- Bold/italic/code inline styling
	- Code blocks fenced with triple backticks

## Test mode (Meaning -> Target)

- Click the meaning->target test button (or press `Ctrl+T`) to open the test dialog
- Default questions per test: 15
- You can set a custom positive integer for questions per test
- You can choose pick preference strategy:
	- `strict` (default): pick by tier order gray -> red -> yellow -> green, then by oldest latest-practice date within the same tier
	- `weighted`: probabilistic pick favoring higher-priority tiers while still sampling all tiers
- If requested count is larger than available vocabularies, the test uses all available entries
- Given a meaning prompt, type the target text and submit
- Judgement rule: exact target-text match after trimming surrounding spaces
- For incorrect answers, feedback shows the correct target text and kana when available
- For incorrect answers, a `View details` button appears to jump directly to that vocabulary detail page
- After the initial cycle, any incorrectly answered vocabularies are retried until each is answered correctly
- Final score/success-rate remains based on the initial cycle only (`initial correct / initial count`)
- Each submitted answer updates stats:
	- `test_count` always increases by 1
	- `error_count` increases by 1 only when the answer is incorrect
	- On correct answers, there is a chance to decrease `error_count` by 1
	- Recovery decrease rules:
		- At most one `error_count` decrease per vocabulary per day
		- If that vocabulary has any wrong answer on that day, no decrease is allowed for the rest of that day

## Test mode (Target -> Kana)

- Click the target->kana test button to open the test dialog
- Uses the same settings and behavior style as EN->JP:
	- Default questions per test: 15
	- Positive-integer question count validation
	- Same pick preference strategy (`strict` or `weighted`)
	- Same score/result flow
- Questions only use entries that have kana
- Prompt is target text; answer is kana
- Judgement rule: exact kana match after trimming surrounding spaces
- For incorrect answers, a `View details` button appears to jump directly to that vocabulary detail page
- After the initial cycle, any incorrectly answered vocabularies are retried until each is answered correctly
- Final score/success-rate remains based on the initial cycle only (`initial correct / initial count`)
- Each submitted answer updates stats with the same rule as EN->JP
	- Includes the same chance-based `error_count` recovery rule and daily restrictions as EN->JP

## Test mode (Target -> Meaning)

- Click the target->meaning test button to open the test dialog
- Prompt is target text; answer format is single choice
- For each question, the app builds options as:
	- Correct meaning
	- Up to 3 random distractor meanings
- Option fallback behavior:
	- Shows as many options as available when fewer than 4 distinct options exist
	- Requires at least 2 options; otherwise the test cannot start
- Uses the same pick preference strategy (`strict` or `weighted`) and same stats update rule
- Includes the same chance-based `error_count` recovery rule and daily restrictions as EN->JP
- For incorrect answers, a `View details` button appears to jump directly to that vocabulary detail page
- After the initial cycle, any incorrectly answered vocabularies are retried until each is answered correctly
- Final score/success-rate remains based on the initial cycle only (`initial correct / initial count`)

## Priority tiers

- `gray`: never tested yet (`test_count = 0`)
- `green`: tested and currently no errors (`test_count > 0` and `error_count = 0`)
- `yellow`: medium error level (`error_count` is 1 or 2)
- `red`: high error level (`error_count >= 3`)

## Daily activity grid

- The `Profile` tab shows a contributions-style grid for the last 180 days
- Each cell counts unique vocabularies practiced on that date
- Repeating the same vocabulary multiple times in one day counts once for that day
- Color thresholds:
	- 0: no color
	- 1-10: tier 1
	- 11-20: tier 2
	- 21-30: tier 3
	- 31+: tier 4
- The grid refreshes automatically after test dialogs close

Manual priority actions adjust `error_count` thresholds (for non-gray entries):

- Increase priority: green -> yellow, yellow -> red, red stays red
- Decrease priority: red -> yellow, yellow -> green, green stays green

## Data location

The app stores entries in a SQLite file at project root:

- vocab.db
