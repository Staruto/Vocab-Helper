# VocabHelper 3.5.0

A modern TypeScript TUI for vocabulary memorization.

## Features

- Multiple configurable workbooks
- List, add, view, edit, and delete vocabulary entries
- Custom meaning and optional fields
- Permanent language preset fields with independent activation and list visibility
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

## Language preset attributes

Japanese, English, German, Chinese, Korean, Spanish, and French workbooks automatically include Example Sentence 1–5, Common Collocations, and Similar Words. Japanese also includes 読み方 (kana). All preset fields start active and hidden from list columns.

During creation, choose which exclusive attributes are active. Later, open `/setting` → **Attributes** → **Language Preset Attributes**, between Meanings and Optional Attributes. Use Up/Down to select, **T** to toggle activation, **Space** to toggle list visibility, and **Ctrl+R** to rename. The selected field uses the selected color; other fields are white. Labels show active/inactive status. While renaming, T types normally. Preset fields cannot be added or removed. Settings changes use the existing save/discard prompt.

Inactive fields are omitted from entry prompts, details, list columns, search, and import destinations. Their stored values and visibility preferences survive edits and imports and become available again when reactivated. Active fields still follow their separate visibility setting for list columns and search.

Existing supported-language workbooks are upgraded automatically, including workbooks created with presets disabled. Existing preset identities and matching optional-field names are promoted without losing values or learning history. Missing fields are added active and hidden. The default Kana label becomes 読み方 (kana); customized names remain unchanged. Import labels use the current displayed names.

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

The home workbook list shows **Today / Goal** for each workbook. Set or clear a workbook's daily goal through `/setting` → **Daily Goal**. Goals are optional positive whole numbers; an empty value means no goal. Tracking works even without a goal, and totals can exceed the goal.

Daily practice counts each word once per completed session, excluding retries. Practicing the same word in another completed session counts again. Cancelled or failed sessions add nothing; a successful save retry counts the session once. Sessions count on their local completion day, including sessions that cross midnight. The home page refreshes when you return and checks for a new day every 30 seconds while open.

Daily totals persist across restarts and entry edits, imports, and deletions. Tracking starts with this feature; past totals cannot be reconstructed from existing learning statistics. New and existing workbooks start with no goal.

Run `/test` to begin a practice session. For workbooks with multiple meanings, press Tab to show the next meaning, wrapping through the workbook's meaning order. Empty meanings display as `-`. Switching preserves your typed answer. Every new question starts on the primary meaning, including retry questions.

Learning results are saved together when you advance past the final correct answer, after all retries finish. Test Count includes every answer; correct initial answers lower Error Count, wrong answers raise it (up to 3), and correct retries leave it unchanged. Esc or Ctrl+C before completion discards the unfinished session. If saving fails, Enter retries saving and Esc discards the results.

Last tested is the local calendar day of completion. The next deadline is that day plus 15, 7, 4, or 1 days for final Error Counts of 0, 1, 2, or 3. New entries come first, followed by entries due today or earlier, then future entries to fill the session; higher Error Count takes priority within due and future groups, with random ties. `/test` defaults to 15 entries; `/test [count]` changes the session size. Entry details show Last tested and Next deadline in Review.

## Practice profile

Run `/profile` inside a workbook to see its practice calendar and streaks. The Monday–Sunday grid covers up to 53 weeks ending this week; narrower terminals show only the newest weeks that fit, without scrolling. Month labels appear above the grid, with the visible date range and color legend underneath. Esc returns to the workbook.

A grey `·` means no recorded practice, yellow `█` means positive practice below the current goal, and green `█` means the goal was reached or exceeded. Without a goal, any positive practice is green. All historical colors use the current daily goal, so changing it recolors history. Future days are blank. History exists only since daily tracking began; missing dates show no recorded practice.

Current and longest streaks count consecutive local calendar days with any practice, regardless of goal completion. If today has no practice yet, the current streak may end yesterday. Longest streak uses all recorded history, including dates outside the visible grid.

## Database

- Default database path: repo-root `vocab.db`
- Override with `VOCAB_HELPER_DB_PATH`
- Numeric IDs are stable and are intentionally not recycled after deletion. Gaps are normal.
- Normalized databases created by VocabHelper 3.0.0 or newer are upgraded automatically.
- On first launch of v3.4.2, tested entries with missing deadlines are repaired automatically using the local calendar date of Last tested and the current Error Count. Existing deadlines and other learning data are preserved; never-tested entries remain unscheduled. Historical deadlines may already be overdue.
- Legacy Python and hybrid database formats are no longer supported.

