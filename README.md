# VocabHelper

TypeScript TUI MVP for vocabulary memorization.

## MVP

- One flat vocabulary list
- List, add, edit, delete
- SQLite-backed
- Imports existing rows from the legacy `vocab.db` on first run
- No Python runtime required for the TUI

## Requirements

- Node 23+
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

## Legacy Code

The Python CLI and GUI code remains in the repository as legacy source, but the supported experience is the TypeScript TUI under `tui/`.

