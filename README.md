# VocabHelper 0.1.0

TypeScript TUI MVP for vocabulary memorization.

## MVP

- One flat vocabulary list
- List, add, edit, delete
- SQLite-backed
- Uses a normalized, TypeScript-owned SQLite database
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
- Numeric IDs are stable and are intentionally not recycled after deletion. Gaps are normal.

### Convert a pre-0.1 database

The application does not modify a legacy or hybrid database during startup. From `tui/`, validate a conversion first:

```powershell
npm run db:convert
```

Then apply it:

```powershell
npm run db:convert:apply
```

The converter builds and validates a clean database before changing the source. It renames the original to a dated `vocab.db.backup-before-v1-*` file and writes a JSON validation report beside the backup. Active TypeScript/MVP vocabulary, fields, statistics, and POS data are preserved; obsolete legacy-only tags and tables are not copied.

## Legacy Code

The Python CLI and GUI code remains in the repository as reference source. It is not part of the supported 0.1 runtime or release artifact; the supported experience is the TypeScript TUI under `tui/`.

