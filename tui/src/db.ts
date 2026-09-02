import { DatabaseSync } from "node:sqlite";
import { resolve } from "node:path";

export type EntryRow = {
  id: number;
  vocabulary: string;
  meaning: string;
  kanaText: string | null;
  createdAt: string;
  updatedAt: string;
};

class ValidationError extends Error {}

function trimRequired(value: string, label: string): string {
  const cleaned = value.trim();
  if (!cleaned) {
    throw new ValidationError(`${label} is required.`);
  }
  return cleaned;
}

function trimOptional(value: string | null | undefined): string | null {
  if (value == null) {
    return null;
  }
  const cleaned = value.trim();
  return cleaned || null;
}

function rowToEntry(row: Record<string, unknown>): EntryRow {
  return {
    id: Number(row.id),
    vocabulary: String(row.vocabulary ?? ""),
    meaning: String(row.meaning ?? ""),
    kanaText: row.kana_text == null ? null : String(row.kana_text),
    createdAt: String(row.created_at ?? ""),
    updatedAt: String(row.updated_at ?? ""),
  };
}

export function defaultDbPath(): string {
  if (process.env.VOCAB_HELPER_DB_PATH?.trim()) {
    return process.env.VOCAB_HELPER_DB_PATH.trim();
  }
  return resolve(process.cwd(), "..", "vocab.db");
}

export class VocabularyRepository {
  private readonly db: DatabaseSync;
  private closed = false;

  constructor(private readonly dbPath: string = defaultDbPath()) {
    this.db = new DatabaseSync(dbPath);
    this.db.exec("PRAGMA foreign_keys = ON;");
    this.initialize();
  }

  close(): void {
    if (this.closed) {
      return;
    }
    this.db.close();
    this.closed = true;
  }

  private transaction<T>(work: () => T): T {
    this.db.exec("BEGIN IMMEDIATE");
    try {
      const result = work();
      this.db.exec("COMMIT");
      return result;
    } catch (error) {
      try {
        this.db.exec("ROLLBACK");
      } catch {
        // Ignore rollback errors so the original failure is preserved.
      }
      throw error;
    }
  }

  initialize(): void {
    this.transaction(() => {
      this.db.exec(`
        CREATE TABLE IF NOT EXISTS mvp_meta (
          key TEXT PRIMARY KEY,
          value TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS mvp_entries (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          vocabulary TEXT NOT NULL CHECK (trim(vocabulary) <> ''),
          meaning TEXT NOT NULL CHECK (trim(meaning) <> ''),
          kana_text TEXT NULL,
          created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
          updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        CREATE INDEX IF NOT EXISTS idx_mvp_entries_created_at
          ON mvp_entries(created_at);
      `);

      const legacyImportComplete = this.getMeta("legacy_import_complete");
      if (legacyImportComplete !== "1") {
        const currentCount = this.countEntries();
        if (currentCount === 0) {
          this.importLegacyRows();
        }
        this.setMeta("legacy_import_complete", "1");
      }
    });
  }

  listEntries(): EntryRow[] {
    const rows = this.db
      .prepare(
        `
        SELECT id, vocabulary, meaning, kana_text, created_at, updated_at
        FROM mvp_entries
        ORDER BY id ASC
        `,
      )
      .all() as Record<string, unknown>[];
    return rows.map(rowToEntry);
  }

  countEntries(): number {
    const row = this.db.prepare("SELECT COUNT(*) AS count FROM mvp_entries").get() as Record<string, unknown>;
    return Number(row.count ?? 0);
  }

  getEntry(entryId: number): EntryRow | null {
    const row = this.db
      .prepare(
        `
        SELECT id, vocabulary, meaning, kana_text, created_at, updated_at
        FROM mvp_entries
        WHERE id = ?
        `,
      )
      .get(entryId) as Record<string, unknown> | undefined;
    return row ? rowToEntry(row) : null;
  }

  addEntry(vocabulary: string, meaning: string): EntryRow {
    const cleanedVocabulary = trimRequired(vocabulary, "Vocabulary");
    const cleanedMeaning = trimRequired(meaning, "Meaning");
    const now = new Date().toISOString();

    const result = this.db
      .prepare(
        `
        INSERT INTO mvp_entries (vocabulary, meaning, created_at, updated_at)
        VALUES (?, ?, ?, ?)
        `,
      )
      .run(cleanedVocabulary, cleanedMeaning, now, now);

    const inserted = this.getEntry(Number(result.lastInsertRowid));
    if (!inserted) {
      throw new Error("Could not load inserted entry.");
    }
    return inserted;
  }

  updateEntry(entryId: number, vocabulary: string, meaning: string): EntryRow {
    const cleanedVocabulary = trimRequired(vocabulary, "Vocabulary");
    const cleanedMeaning = trimRequired(meaning, "Meaning");
    const existing = this.getEntry(entryId);
    if (!existing) {
      throw new Error(`Entry with id ${entryId} was not found.`);
    }

    this.db
      .prepare(
        `
        UPDATE mvp_entries
        SET vocabulary = ?, meaning = ?, updated_at = ?
        WHERE id = ?
        `,
      )
      .run(cleanedVocabulary, cleanedMeaning, new Date().toISOString(), entryId);

    const updated = this.getEntry(entryId);
    if (!updated) {
      throw new Error(`Entry with id ${entryId} was not found.`);
    }
    return updated;
  }

  deleteEntry(entryId: number): void {
    const existing = this.getEntry(entryId);
    if (!existing) {
      throw new Error(`Entry with id ${entryId} was not found.`);
    }

    this.db.prepare("DELETE FROM mvp_entries WHERE id = ?").run(entryId);
  }

  private getMeta(key: string): string | null {
    const row = this.db.prepare("SELECT value FROM mvp_meta WHERE key = ?").get(key) as Record<string, unknown> | undefined;
    return row ? String(row.value ?? "") : null;
  }

  private setMeta(key: string, value: string): void {
    this.db
      .prepare(
        `
        INSERT INTO mvp_meta (key, value)
        VALUES (?, ?)
        ON CONFLICT(key) DO UPDATE SET value = excluded.value
        `,
      )
      .run(key, value);
  }

  private importLegacyRows(): void {
    const legacyTable = this.db
      .prepare(
        `
        SELECT 1 AS present
        FROM sqlite_master
        WHERE type = 'table'
          AND name = 'vocab_entries'
        `,
      )
      .get() as Record<string, unknown> | undefined;

    if (!legacyTable) {
      return;
    }

    const rows = this.db
      .prepare(
        `
        SELECT id, japanese_text, kana_text, english_text, created_at
        FROM vocab_entries
        ORDER BY id ASC
        `,
      )
      .all() as Record<string, unknown>[];

    for (const row of rows) {
      const vocabulary = trimRequired(String(row.japanese_text ?? ""), "Vocabulary");
      const meaning = trimRequired(String(row.english_text ?? ""), "Meaning");
      const kanaText = trimOptional(row.kana_text == null ? null : String(row.kana_text));
      const createdAt = String(row.created_at ?? new Date().toISOString());

      this.db
        .prepare(
          `
          INSERT INTO mvp_entries (vocabulary, meaning, kana_text, created_at, updated_at)
          VALUES (?, ?, ?, ?, ?)
          `,
        )
        .run(vocabulary, meaning, kanaText, createdAt, createdAt);
    }
  }
}
