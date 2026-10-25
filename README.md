# VocabHelper 3.4.1

A modern TypeScript TUI for vocabulary memorization.

## Features

- Multiple configurable workbooks
- List, add, view, edit, and delete vocabulary entries
- Custom meaning and optional fields
- Workbook tag types with optional list and practice badges
- Prioritized practice with learning statistics
- Normalized SQLite storage
- Previewed UTF-8 labeled-text imports into the open workbook

## Requirements

- Node 24 LTS or newer
- npm

## Run

```powershell
cd tui
npm install
npm run dev
```

Build:

```powershell
npm run build
npm run start
```

## Import

Run `/import` inside an open workbook. If that workbook has no default import file, enter the path to a UTF-8 `.txt` file. In the preview, use Up/Down to select the file path, status filters, or record list. Press Enter on the file path to load a different file, or press Enter on the filters or records to import all ready entries. Esc cancels.

Set, replace, or clear the workbook's default file under `/setting` > `Import`. Saved paths are validated and stored as absolute paths. A valid default opens directly in the import preview; if it can no longer be loaded, the app reports the error and returns to path entry. After importing from a different file, the app asks whether to remember it as the workbook default.

Each line uses `Workbook label: value`, and blank lines separate entries. Labels match the workbook's displayed attribute and tag-type names exactly; their order does not matter. Multiple tags use commas. Unknown fields and tag values are reported and ignored, while records missing the vocabulary or primary meaning are skipped. New vocabularies are added; valid duplicates synchronize changed meanings, optional fields, and tags, and identical duplicates are left untouched.

```text
Japanese: 連休
English: consecutive holidays / long weekend
Part of Speech: 名詞

Japanese: 休暇
English: vacation
```

Updating an existing entry through `/edit` or `/import` resets its test count, error count, last-tested time, and next-test deadline, returning it to the gray, untested tier. Changes to vocabulary, meanings, optional fields, or assigned tags trigger the reset. Identical edits and imports preserve the test record.

## Practice

Run `/test` to begin a practice session. For workbooks with multiple meanings, press Tab to show the next meaning, wrapping through the workbook's meaning order. Empty meanings display as `-`. Switching preserves your typed answer. Every new question starts on the primary meaning, including retry questions.

Learning results are saved together when you advance past the final correct answer, after all retries finish. Test Count includes every answer; correct initial answers lower Error Count, wrong answers raise it (up to 3), and correct retries leave it unchanged. Esc or Ctrl+C before completion discards the unfinished session. If saving fails, Enter retries saving and Esc discards the results.

Last tested is the local calendar day of completion. The next deadline is that day plus 15, 7, 4, or 1 days for final Error Counts of 0, 1, 2, or 3. New entries come first, followed by entries due today or earlier, then future entries to fill the session; higher Error Count takes priority within due and future groups, with random ties. `/test` defaults to 15 entries; `/test [count]` changes the session size. Entry details show Last tested and Next deadline in Review.

## Database

- Default database path: repo-root `vocab.db`
- Override with `VOCAB_HELPER_DB_PATH`
- Numeric IDs are stable and are intentionally not recycled after deletion. Gaps are normal.
- Normalized databases created by VocabHelper 3.0.0 or newer are upgraded automatically.
- Legacy Python and hybrid database formats are no longer supported.

