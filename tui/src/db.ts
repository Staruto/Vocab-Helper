import { resolve } from "node:path";
import { DatabaseSync } from "node:sqlite";

export type WorkbookRow = {
  id: number;
  name: string;
  wordCount: number;
  createdAt: string;
  vocabularyLabel: string;
  vocabularyLanguageCode: string | null;
  meaningAttributes: MeaningAttribute[];
};

export type MeaningAttribute = {
  position: number;
  label: string;
  languageCode: string | null;
};

export type EntryRow = {
  id: number;
  workbookId: number;
  vocabulary: string;
  meaning: string;
  meanings: string[];
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
    workbookId: Number(row.workbook_id),
    vocabulary: String(row.vocabulary ?? ""),
    meaning: String(row.meaning ?? ""),
    meanings: [],
    kanaText: row.kana_text == null ? null : String(row.kana_text),
    createdAt: String(row.created_at ?? ""),
    updatedAt: String(row.updated_at ?? ""),
  };
}

function rowToWorkbook(row: Record<string, unknown>): WorkbookRow {
  return {
    id: Number(row.id),
    name: String(row.name ?? ""),
    wordCount: Number(row.word_count ?? 0),
    createdAt: String(row.created_at ?? ""),
    vocabularyLabel: String(row.vocabulary_label ?? "Vocabulary"),
    vocabularyLanguageCode: row.vocabulary_language_code == null ? null : String(row.vocabulary_language_code),
    meaningAttributes: [],
  };
}

const LANGUAGE_PRESETS = [
  { code: "JP", label: "Japanese" },
  { code: "EN", label: "English" },
  { code: "ZH", label: "Chinese" },
  { code: "KO", label: "Korean" },
  { code: "ES", label: "Spanish" },
  { code: "FR", label: "French" },
  { code: "DE", label: "German" },
] as const;

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

        CREATE TABLE IF NOT EXISTS mvp_workbooks (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          name TEXT NOT NULL COLLATE NOCASE UNIQUE,
          vocabulary_label TEXT NOT NULL DEFAULT 'Vocabulary',
          vocabulary_language_code TEXT NULL,
          created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS mvp_workbook_meaning_attributes (
          workbook_id INTEGER NOT NULL,
          position INTEGER NOT NULL,
          label TEXT NOT NULL,
          language_code TEXT NULL,
          PRIMARY KEY (workbook_id, position),
          FOREIGN KEY (workbook_id) REFERENCES mvp_workbooks(id) ON DELETE CASCADE,
          CHECK (position >= 1),
          CHECK (trim(label) <> '')
        );

        CREATE TABLE IF NOT EXISTS mvp_entries (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          workbook_id INTEGER NULL,
          vocabulary TEXT NOT NULL CHECK (trim(vocabulary) <> ''),
          meaning TEXT NOT NULL CHECK (trim(meaning) <> ''),
          kana_text TEXT NULL,
          created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
          updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
          FOREIGN KEY (workbook_id) REFERENCES mvp_workbooks(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS mvp_entry_meanings (
          entry_id INTEGER NOT NULL,
          position INTEGER NOT NULL,
          value TEXT NOT NULL DEFAULT '',
          PRIMARY KEY (entry_id, position),
          FOREIGN KEY (entry_id) REFERENCES mvp_entries(id) ON DELETE CASCADE,
          CHECK (position >= 1)
        );

        CREATE INDEX IF NOT EXISTS idx_mvp_entries_created_at
          ON mvp_entries(created_at);

      `);

      this.ensureColumn("mvp_entries", "workbook_id", "INTEGER NULL");
      this.ensureColumn("mvp_workbooks", "vocabulary_label", "TEXT NOT NULL DEFAULT 'Vocabulary'");
      this.ensureColumn("mvp_workbooks", "vocabulary_language_code", "TEXT NULL");
      this.db.exec(`
        CREATE INDEX IF NOT EXISTS idx_mvp_entries_workbook_id
          ON mvp_entries(workbook_id);
      `);
      this.importLegacyDataIfNeeded();
      this.ensureWorkbookBackfill();
      this.ensureWorkbookSchemaBackfill();
      this.ensureCurrentWorkbookSetting();
    });
  }

  listWorkbooks(): WorkbookRow[] {
    const rows = this.db
      .prepare(
        `
        SELECT
          w.id,
          w.name,
          w.vocabulary_label,
          w.vocabulary_language_code,
          w.created_at,
          COUNT(e.id) AS word_count
        FROM mvp_workbooks AS w
        LEFT JOIN mvp_entries AS e
          ON e.workbook_id = w.id
        GROUP BY w.id, w.name, w.created_at
        ORDER BY w.id ASC
        `,
      )
      .all() as Record<string, unknown>[];
    return rows.map((row) => this.withWorkbookAttributes(rowToWorkbook(row)));
  }

  getWorkbook(workbookId: number): WorkbookRow | null {
    const row = this.db
      .prepare(
        `
        SELECT
          w.id,
          w.name,
          w.vocabulary_label,
          w.vocabulary_language_code,
          w.created_at,
          COUNT(e.id) AS word_count
        FROM mvp_workbooks AS w
        LEFT JOIN mvp_entries AS e
          ON e.workbook_id = w.id
        WHERE w.id = ?
        GROUP BY w.id, w.name, w.created_at
        `,
      )
      .get(workbookId) as Record<string, unknown> | undefined;
    return row ? this.withWorkbookAttributes(rowToWorkbook(row)) : null;
  }

  listMeaningAttributes(workbookId: number): MeaningAttribute[] {
    const workbook = this.getWorkbook(workbookId);
    if (!workbook) {
      throw new Error(`Workbook with id ${workbookId} was not found.`);
    }
    return workbook.meaningAttributes;
  }

  createWorkbook(
    name: string,
    vocabularyLabel = "Vocabulary",
    vocabularyLanguageCode: string | null = null,
    meaningAttributes: MeaningAttribute[] = [{ position: 1, label: "Meaning 1", languageCode: null }],
  ): WorkbookRow {
    const cleanedName = trimRequired(name, "Workbook name");
    const cleanedVocabularyLabel = trimOptional(vocabularyLabel) ?? "Vocabulary";
    const attributes = this.normalizeMeaningAttributes(meaningAttributes);
    const workbook = this.transaction(() => {
      const result = this.db.prepare(
        "INSERT INTO mvp_workbooks (name, vocabulary_label, vocabulary_language_code) VALUES (?, ?, ?)",
      ).run(cleanedName, cleanedVocabularyLabel, vocabularyLanguageCode);
      const workbookId = Number(result.lastInsertRowid);
      const insertAttribute = this.db.prepare(
        "INSERT INTO mvp_workbook_meaning_attributes (workbook_id, position, label, language_code) VALUES (?, ?, ?, ?)",
      );
      for (const attribute of attributes) {
        insertAttribute.run(workbookId, attribute.position, attribute.label, attribute.languageCode);
      }
      return this.getWorkbook(workbookId);
    });
    if (!workbook) {
      throw new Error("Could not load inserted workbook.");
    }
    if (this.getCurrentWorkbookId() === null) {
      this.setCurrentWorkbookId(workbook.id);
    }
    return workbook;
  }

  deleteWorkbook(workbookId: number): number | null {
    const workbook = this.getWorkbook(workbookId);
    if (!workbook) {
      throw new Error(`Workbook with id ${workbookId} was not found.`);
    }

    return this.transaction(() => {
      const currentWorkbookId = this.readCurrentWorkbookId();
      this.db.prepare("DELETE FROM mvp_entries WHERE workbook_id = ?").run(workbookId);
      this.db.prepare("DELETE FROM mvp_workbooks WHERE id = ?").run(workbookId);

      const remainingWorkbookId = this.firstWorkbookId();
      const nextCurrentWorkbookId = currentWorkbookId === workbookId ? remainingWorkbookId : currentWorkbookId ?? remainingWorkbookId;
      this.writeCurrentWorkbookId(nextCurrentWorkbookId);
      return nextCurrentWorkbookId;
    });
  }

  getCurrentWorkbookId(): number | null {
    const currentWorkbookId = this.readCurrentWorkbookId();
    if (currentWorkbookId !== null && this.getWorkbook(currentWorkbookId)) {
      return currentWorkbookId;
    }

    const fallbackWorkbookId = this.firstWorkbookId();
    this.writeCurrentWorkbookId(fallbackWorkbookId);
    return fallbackWorkbookId;
  }

  setCurrentWorkbookId(workbookId: number): WorkbookRow {
    const workbook = this.getWorkbook(workbookId);
    if (!workbook) {
      throw new Error(`Workbook with id ${workbookId} was not found.`);
    }
    this.writeCurrentWorkbookId(workbook.id);
    return workbook;
  }

  listEntries(workbookId?: number): EntryRow[] {
    const resolvedWorkbookId = this.resolveWorkbookId(workbookId);
    if (resolvedWorkbookId === null) {
      return [];
    }

    const rows = this.db
      .prepare(
        `
        SELECT id, workbook_id, vocabulary, meaning, kana_text, created_at, updated_at
        FROM mvp_entries
        WHERE workbook_id = ?
        ORDER BY id ASC
        `,
      )
      .all(resolvedWorkbookId) as Record<string, unknown>[];
    return rows.map((row) => this.withEntryMeanings(rowToEntry(row)));
  }

  countEntries(workbookId?: number): number {
    if (workbookId === undefined) {
      const row = this.db.prepare("SELECT COUNT(*) AS count FROM mvp_entries").get() as Record<string, unknown>;
      return Number(row.count ?? 0);
    }

    const row = this.db.prepare("SELECT COUNT(*) AS count FROM mvp_entries WHERE workbook_id = ?").get(workbookId) as Record<string, unknown>;
    return Number(row.count ?? 0);
  }

  getEntry(entryId: number): EntryRow | null {
    const row = this.db
      .prepare(
        `
        SELECT id, workbook_id, vocabulary, meaning, kana_text, created_at, updated_at
        FROM mvp_entries
        WHERE id = ?
        `,
      )
      .get(entryId) as Record<string, unknown> | undefined;
    return row ? this.withEntryMeanings(rowToEntry(row)) : null;
  }

  addEntry(workbookId: number, vocabulary: string, meaning: string, meanings?: string[]): EntryRow {
    if (!this.getWorkbook(workbookId)) {
      throw new Error(`Workbook with id ${workbookId} was not found.`);
    }

    const cleanedVocabulary = trimRequired(vocabulary, "Vocabulary");
    const values = meanings && meanings.length > 0 ? meanings : [meaning];
    const cleanedMeanings = values.map((value) => trimOptional(value) ?? "");
    cleanedMeanings[0] = trimRequired(cleanedMeanings[0], "Meaning");
    const now = new Date().toISOString();

    const result = this.transaction(() => {
      const inserted = this.db.prepare(
        `
        INSERT INTO mvp_entries (workbook_id, vocabulary, meaning, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?)
        `,
      ).run(workbookId, cleanedVocabulary, cleanedMeanings[0], now, now);
      const entryId = Number(inserted.lastInsertRowid);
      const insertMeaning = this.db.prepare("INSERT INTO mvp_entry_meanings (entry_id, position, value) VALUES (?, ?, ?)");
      cleanedMeanings.forEach((value, index) => insertMeaning.run(entryId, index + 1, value));
      return entryId;
    });

    const inserted = this.getEntry(result);
    if (!inserted) {
      throw new Error("Could not load inserted entry.");
    }
    return inserted;
  }

  updateEntry(entryId: number, vocabulary: string, meaning: string, meanings?: string[]): EntryRow {
    const cleanedVocabulary = trimRequired(vocabulary, "Vocabulary");
    const values = meanings && meanings.length > 0 ? meanings : [meaning];
    const cleanedMeanings = values.map((value) => trimOptional(value) ?? "");
    cleanedMeanings[0] = trimRequired(cleanedMeanings[0], "Meaning");
    const existing = this.getEntry(entryId);
    if (!existing) {
      throw new Error(`Entry with id ${entryId} was not found.`);
    }

    this.transaction(() => {
      this.db
        .prepare(
        `
        UPDATE mvp_entries
        SET vocabulary = ?, meaning = ?, updated_at = ?
        WHERE id = ?
        `,
        )
        .run(cleanedVocabulary, cleanedMeanings[0], new Date().toISOString(), entryId);
      this.db.prepare("DELETE FROM mvp_entry_meanings WHERE entry_id = ?").run(entryId);
      const insertMeaning = this.db.prepare("INSERT INTO mvp_entry_meanings (entry_id, position, value) VALUES (?, ?, ?)");
      cleanedMeanings.forEach((value, index) => insertMeaning.run(entryId, index + 1, value));
    });

    const updated = this.getEntry(entryId);
    if (!updated) {
      throw new Error(`Entry with id ${entryId} was not found.`);
    }
    return updated;
  }

  private withWorkbookAttributes(workbook: WorkbookRow): WorkbookRow {
    const rows = this.db.prepare(
      "SELECT position, label, language_code FROM mvp_workbook_meaning_attributes WHERE workbook_id = ? ORDER BY position ASC",
    ).all(workbook.id) as Record<string, unknown>[];
    return {
      ...workbook,
      meaningAttributes: rows.map((row) => ({
        position: Number(row.position),
        label: String(row.label),
        languageCode: row.language_code == null ? null : String(row.language_code),
      })),
    };
  }

  private withEntryMeanings(entry: EntryRow): EntryRow {
    const rows = this.db.prepare("SELECT position, value FROM mvp_entry_meanings WHERE entry_id = ? ORDER BY position ASC").all(entry.id) as Record<string, unknown>[];
    const meanings = rows.length > 0 ? rows.map((row) => String(row.value ?? "")) : [entry.meaning];
    return { ...entry, meanings, meaning: meanings[0] ?? entry.meaning };
  }

  private normalizeMeaningAttributes(attributes: MeaningAttribute[]): MeaningAttribute[] {
    if (attributes.length < 1 || attributes.length > 5) {
      throw new ValidationError("Meaning attributes must contain between 1 and 5 items.");
    }
    const normalized = attributes.map((attribute, index) => ({
      position: index + 1,
      label: trimRequired(attribute.label, `Meaning ${index + 1} label`),
      languageCode: attribute.languageCode?.trim().toUpperCase() || null,
    }));
    if (new Set(normalized.map((attribute) => attribute.label.toLocaleLowerCase())).size !== normalized.length) {
      throw new ValidationError("Meaning attribute labels must be unique.");
    }
    return normalized;
  }

  deleteEntry(entryId: number): void {
    const existing = this.getEntry(entryId);
    if (!existing) {
      throw new Error(`Entry with id ${entryId} was not found.`);
    }

    this.db.prepare("DELETE FROM mvp_entries WHERE id = ?").run(entryId);
  }

  private ensureColumn(tableName: string, columnName: string, definition: string): void {
    const rows = this.db.prepare(`PRAGMA table_info(${tableName})`).all() as Record<string, unknown>[];
    const existingColumns = new Set(rows.map((row) => String(row.name)));
    if (!existingColumns.has(columnName)) {
      this.db.exec(`ALTER TABLE ${tableName} ADD COLUMN ${columnName} ${definition}`);
    }
  }

  private tableExists(tableName: string): boolean {
    const row = this.db
      .prepare(
        `
        SELECT 1 AS present
        FROM sqlite_master
        WHERE type = 'table'
          AND name = ?
        `,
      )
      .get(tableName) as Record<string, unknown> | undefined;
    return Boolean(row);
  }

  private importLegacyDataIfNeeded(): void {
    if (this.getMeta("legacy_import_complete") === "1") {
      return;
    }

    if (this.countEntries() > 0 || this.countWorkbooks() > 0) {
      this.setMeta("legacy_import_complete", "1");
      return;
    }

    const legacyEntriesExist = this.tableExists("vocab_entries");
    const legacyWorkbooksExist = this.tableExists("workbooks");
    if (!legacyEntriesExist && !legacyWorkbooksExist) {
      this.setMeta("legacy_import_complete", "1");
      return;
    }

    const workbookIdMap = legacyWorkbooksExist ? this.importLegacyWorkbooks() : new Map<number, number>();
    const fallbackWorkbookId = this.ensureDefaultWorkbookIfNeeded(workbookIdMap.size > 0 ? null : "Default");

    if (legacyEntriesExist) {
      this.importLegacyEntries(workbookIdMap, fallbackWorkbookId);
    }

    const legacyCurrentWorkbookId = this.readLegacyCurrentWorkbookId();
    const currentWorkbookId =
      legacyCurrentWorkbookId !== null ? workbookIdMap.get(legacyCurrentWorkbookId) ?? fallbackWorkbookId : fallbackWorkbookId ?? this.firstWorkbookId();
    this.writeCurrentWorkbookId(currentWorkbookId);
    this.setMeta("legacy_import_complete", "1");
  }

  private importLegacyWorkbooks(): Map<number, number> {
    const workbookIdMap = new Map<number, number>();
    const rows = this.db
      .prepare(
        `
        SELECT id, name, created_at
        FROM workbooks
        ORDER BY id ASC
        `,
      )
      .all() as Record<string, unknown>[];

    for (const row of rows) {
      const legacyId = Number(row.id);
      const name = trimRequired(String(row.name ?? `Workbook ${legacyId}`), "Workbook name");
      const createdAt = String(row.created_at ?? new Date().toISOString());
      const result = this.db.prepare("INSERT INTO mvp_workbooks (name, created_at) VALUES (?, ?)").run(name, createdAt);
      workbookIdMap.set(legacyId, Number(result.lastInsertRowid));
    }

    return workbookIdMap;
  }

  private importLegacyEntries(workbookIdMap: Map<number, number>, fallbackWorkbookId: number | null): void {
    const columns = this.tableColumns("vocab_entries");
    const hasWorkbookId = columns.has("workbook_id");
    const workbookColumn = hasWorkbookId ? "workbook_id" : "NULL AS workbook_id";
    const rows = this.db
      .prepare(
        `
        SELECT ${workbookColumn}, japanese_text, kana_text, english_text, created_at
        FROM vocab_entries
        ORDER BY id ASC
        `,
      )
      .all() as Record<string, unknown>[];

    for (const row of rows) {
      const legacyWorkbookId = row.workbook_id == null ? null : Number(row.workbook_id);
      const workbookId = (legacyWorkbookId === null ? null : workbookIdMap.get(legacyWorkbookId)) ?? fallbackWorkbookId ?? this.firstWorkbookId();
      if (workbookId === null) {
        continue;
      }

      const vocabulary = trimRequired(String(row.japanese_text ?? ""), "Vocabulary");
      const meaning = trimRequired(String(row.english_text ?? ""), "Meaning");
      const kanaText = trimOptional(row.kana_text == null ? null : String(row.kana_text));
      const createdAt = String(row.created_at ?? new Date().toISOString());

      this.db
        .prepare(
          `
          INSERT INTO mvp_entries (workbook_id, vocabulary, meaning, kana_text, created_at, updated_at)
          VALUES (?, ?, ?, ?, ?, ?)
          `,
        )
        .run(workbookId, vocabulary, meaning, kanaText, createdAt, createdAt);
    }
  }

  private tableColumns(tableName: string): Set<string> {
    const rows = this.db.prepare(`PRAGMA table_info(${tableName})`).all() as Record<string, unknown>[];
    return new Set(rows.map((row) => String(row.name)));
  }

  private ensureWorkbookBackfill(): void {
    if (this.countEntries() === 0) {
      return;
    }

    const defaultWorkbookId = this.ensureDefaultWorkbookIfNeeded("Default");
    if (defaultWorkbookId !== null) {
      this.db.prepare("UPDATE mvp_entries SET workbook_id = ? WHERE workbook_id IS NULL").run(defaultWorkbookId);
    }
  }

  private ensureWorkbookSchemaBackfill(): void {
    const workbooks = this.db.prepare("SELECT id FROM mvp_workbooks ORDER BY id ASC").all() as Record<string, unknown>[];
    const insertAttribute = this.db.prepare(
      "INSERT OR IGNORE INTO mvp_workbook_meaning_attributes (workbook_id, position, label, language_code) VALUES (?, 1, 'Meaning 1', NULL)",
    );
    for (const row of workbooks) {
      const workbookId = Number(row.id);
      insertAttribute.run(workbookId);
    }

    const entries = this.db.prepare("SELECT id, meaning FROM mvp_entries").all() as Record<string, unknown>[];
    const insertMeaning = this.db.prepare(
      "INSERT OR IGNORE INTO mvp_entry_meanings (entry_id, position, value) VALUES (?, 1, ?)",
    );
    for (const row of entries) {
      insertMeaning.run(Number(row.id), String(row.meaning ?? ""));
    }
  }

  private ensureDefaultWorkbookIfNeeded(name: string | null): number | null {
    const existingWorkbookId = this.firstWorkbookId();
    if (existingWorkbookId !== null) {
      return existingWorkbookId;
    }
    if (name === null) {
      return null;
    }
    const result = this.db.prepare("INSERT INTO mvp_workbooks (name) VALUES (?)").run(name);
    return Number(result.lastInsertRowid);
  }

  private ensureCurrentWorkbookSetting(): void {
    const currentWorkbookId = this.getCurrentWorkbookId();
    this.writeCurrentWorkbookId(currentWorkbookId);
  }

  private countWorkbooks(): number {
    const row = this.db.prepare("SELECT COUNT(*) AS count FROM mvp_workbooks").get() as Record<string, unknown>;
    return Number(row.count ?? 0);
  }

  private resolveWorkbookId(workbookId?: number): number | null {
    if (workbookId !== undefined) {
      return this.getWorkbook(workbookId) ? workbookId : null;
    }
    return this.getCurrentWorkbookId();
  }

  private firstWorkbookId(): number | null {
    const row = this.db.prepare("SELECT id FROM mvp_workbooks ORDER BY id ASC LIMIT 1").get() as Record<string, unknown> | undefined;
    return row ? Number(row.id) : null;
  }

  private readCurrentWorkbookId(): number | null {
    const value = this.getMeta("current_workbook_id");
    if (value === null || value.trim() === "") {
      return null;
    }
    const parsed = Number(value);
    return Number.isInteger(parsed) ? parsed : null;
  }

  private writeCurrentWorkbookId(workbookId: number | null): void {
    this.setMeta("current_workbook_id", workbookId === null ? "" : String(workbookId));
  }

  private readLegacyCurrentWorkbookId(): number | null {
    if (!this.tableExists("app_settings")) {
      return null;
    }
    const row = this.db
      .prepare(
        `
        SELECT value
        FROM app_settings
        WHERE key = 'current_workbook_id'
        `,
      )
      .get() as Record<string, unknown> | undefined;
    if (!row || row.value == null || String(row.value).trim() === "") {
      return null;
    }
    const parsed = Number(row.value);
    return Number.isInteger(parsed) ? parsed : null;
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
}
