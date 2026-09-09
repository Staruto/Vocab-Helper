# VocabHelper 3.1.0

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

Run `/import` inside an open workbook, then enter the path to a UTF-8 `.txt` file. Review the preview and press Enter to import or Esc to cancel.

Each line uses `Workbook label: value`, and blank lines separate entries. Labels match the workbook's displayed attribute and tag-type names exactly; their order does not matter. Multiple tags use commas. Unknown fields and tag values are reported and ignored, while records missing the vocabulary or primary meaning are skipped.

```text
Japanese: 連休
English: consecutive holidays / long weekend
Part of Speech: 名詞

Japanese: 休暇
English: vacation
```

## Database

- Default database path: repo-root `vocab.db`
- Override with `VOCAB_HELPER_DB_PATH`
- Numeric IDs are stable and are intentionally not recycled after deletion. Gaps are normal.
- Normalized databases created by VocabHelper 3.0.0 or newer are upgraded automatically.
- Legacy Python and hybrid database formats are no longer supported.

