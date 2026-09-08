# VocabHelper 3.1.0

A modern TypeScript TUI for vocabulary memorization.

## Features

- Multiple configurable workbooks
- List, add, view, edit, and delete vocabulary entries
- Custom meaning and optional fields
- Workbook tag types with optional list and practice badges
- Prioritized practice with learning statistics
- Normalized SQLite storage

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

## Database

- Default database path: repo-root `vocab.db`
- Override with `VOCAB_HELPER_DB_PATH`
- Numeric IDs are stable and are intentionally not recycled after deletion. Gaps are normal.
- Normalized databases created by VocabHelper 3.0.0 or newer are upgraded automatically.
- Legacy Python and hybrid database formats are no longer supported.

