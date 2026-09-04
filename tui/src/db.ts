import { resolve } from "node:path";
import { DatabaseSync } from "node:sqlite";

export type WorkbookRow = {
  id: number;
  name: string;
  wordCount: number;
  createdAt: string;
  vocabularyLabel: string;
  vocabularyLanguageCode: string | null;
  presetEnabled: boolean;
  meaningAttributes: MeaningAttribute[];
  metadataAttributes: MetadataAttribute[];
};

export type MeaningAttribute = {
  position: number;
  label: string;
  languageCode: string | null;
};

export type MetadataAttribute = {
  key: string;
  label: string;
  languageCode: string | null;
  required: boolean;
  visible: boolean;
  displayOrder: number;
};

export type PosTag = { id: number; name: string; predefined: boolean };

export type EntryRow = {
  id: number;
  workbookId: number;
  vocabulary: string;
  meaning: string;
  meanings: string[];
  kanaText: string | null;
  attributes: Record<string, string>;
  posTags: PosTag[];
  createdAt: string;
  updatedAt: string;
  testCount: number;
  errorCount: number;
  tier: "gray" | "green" | "yellow" | "red";
  lastTested: string | null;
};

class ValidationError extends Error {}

function tierFor(testCount: number, errorCount: number): EntryRow["tier"] {
  if (testCount <= 0) return "gray";
  if (errorCount <= 0) return "green";
  if (errorCount <= 2) return "yellow";
  return "red";
}

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
    attributes: {},
    posTags: [],
    createdAt: String(row.created_at ?? ""),
    updatedAt: String(row.updated_at ?? ""),
    testCount: 0,
    errorCount: 0,
    tier: "gray",
    lastTested: null,
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
    presetEnabled: Number(row.preset_enabled ?? 0) === 1,
    meaningAttributes: [],
    metadataAttributes: [],
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
const JAPANESE_POS_TAGS = ["名詞", "固有名詞", "イ形容詞", "ナ形容詞", "動詞 (自動詞)", "動詞 (他動詞)", "副詞", "連体詞", "接続詞", "連語", "その他"];

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
          preset_enabled INTEGER NOT NULL DEFAULT 0,
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

        CREATE TABLE IF NOT EXISTS mvp_workbook_attributes (
          workbook_id INTEGER NOT NULL,
          attribute_key TEXT NOT NULL,
          label TEXT NOT NULL,
          language_code TEXT NULL,
          is_required INTEGER NOT NULL DEFAULT 0,
          is_visible INTEGER NOT NULL DEFAULT 0,
          display_order INTEGER NOT NULL DEFAULT 0,
          PRIMARY KEY (workbook_id, attribute_key),
          FOREIGN KEY (workbook_id) REFERENCES mvp_workbooks(id) ON DELETE CASCADE,
          CHECK (trim(attribute_key) <> ''), CHECK (trim(label) <> '')
        );

        CREATE TABLE IF NOT EXISTS mvp_entry_attributes (
          entry_id INTEGER NOT NULL,
          attribute_key TEXT NOT NULL,
          value TEXT NOT NULL DEFAULT '',
          PRIMARY KEY (entry_id, attribute_key),
          FOREIGN KEY (entry_id) REFERENCES mvp_entries(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS mvp_pos_tags (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          workbook_id INTEGER NOT NULL,
          name TEXT NOT NULL COLLATE NOCASE,
          is_predefined INTEGER NOT NULL DEFAULT 0,
          UNIQUE (workbook_id, name),
          FOREIGN KEY (workbook_id) REFERENCES mvp_workbooks(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS mvp_entry_pos_tags (
          entry_id INTEGER NOT NULL,
          tag_id INTEGER NOT NULL,
          PRIMARY KEY (entry_id, tag_id),
          FOREIGN KEY (entry_id) REFERENCES mvp_entries(id) ON DELETE CASCADE,
          FOREIGN KEY (tag_id) REFERENCES mvp_pos_tags(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS mvp_entry_stats (
          entry_id INTEGER PRIMARY KEY,
          test_count INTEGER NOT NULL DEFAULT 0,
          error_count INTEGER NOT NULL DEFAULT 0,
          last_tested TEXT NULL,
          FOREIGN KEY (entry_id) REFERENCES mvp_entries(id) ON DELETE CASCADE,
          CHECK (test_count >= 0), CHECK (error_count >= 0 AND error_count <= 3)
        );

        CREATE INDEX IF NOT EXISTS idx_mvp_entries_created_at
          ON mvp_entries(created_at);

      `);

      this.ensureColumn("mvp_entries", "workbook_id", "INTEGER NULL");
      this.ensureColumn("mvp_workbooks", "vocabulary_label", "TEXT NOT NULL DEFAULT 'Vocabulary'");
      this.ensureColumn("mvp_workbooks", "vocabulary_language_code", "TEXT NULL");
      this.ensureColumn("mvp_workbooks", "preset_enabled", "INTEGER NOT NULL DEFAULT 0");
      this.db.exec(`
        CREATE INDEX IF NOT EXISTS idx_mvp_entries_workbook_id
          ON mvp_entries(workbook_id);
      `);
      this.importLegacyDataIfNeeded();
      this.ensureWorkbookBackfill();
      this.ensureWorkbookSchemaBackfill();
      this.ensureMetadataBackfill();
      this.repairLegacyWorkbookSplit();
      this.ensureStatsBackfill();
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
          w.preset_enabled,
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
          w.preset_enabled,
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
    presetEnabled = vocabularyLanguageCode === "JP",
  ): WorkbookRow {
    const cleanedName = trimRequired(name, "Workbook name");
    const cleanedVocabularyLabel = trimOptional(vocabularyLabel) ?? "Vocabulary";
    const attributes = this.normalizeMeaningAttributes(meaningAttributes);
    const workbook = this.transaction(() => {
      const result = this.db.prepare(
        "INSERT INTO mvp_workbooks (name, vocabulary_label, vocabulary_language_code, preset_enabled) VALUES (?, ?, ?, ?)",
      ).run(cleanedName, cleanedVocabularyLabel, vocabularyLanguageCode, presetEnabled ? 1 : 0);
      const workbookId = Number(result.lastInsertRowid);
      const insertAttribute = this.db.prepare(
        "INSERT INTO mvp_workbook_meaning_attributes (workbook_id, position, label, language_code) VALUES (?, ?, ?, ?)",
      );
      for (const attribute of attributes) {
        insertAttribute.run(workbookId, attribute.position, attribute.label, attribute.languageCode);
      }
      this.ensureMetadataBackfill();
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

  updateWorkbookSettings(
    workbookId: number,
    name: string,
    vocabularyLabel: string,
    vocabularyLanguageCode: string | null,
    meaningAttributes: MeaningAttribute[],
    presetEnabled = vocabularyLanguageCode === "JP",
  ): WorkbookRow {
    if (!this.getWorkbook(workbookId)) {
      throw new Error(`Workbook with id ${workbookId} was not found.`);
    }
    const attributes = this.normalizeMeaningAttributes(meaningAttributes);
    const cleanedName = trimRequired(name, "Workbook name");
    this.transaction(() => {
      this.db.prepare("UPDATE mvp_workbooks SET name = ?, vocabulary_label = ?, vocabulary_language_code = ?, preset_enabled = ? WHERE id = ?")
        .run(cleanedName, trimOptional(vocabularyLabel) ?? "Vocabulary", vocabularyLanguageCode, presetEnabled ? 1 : 0, workbookId);
      this.db.prepare("DELETE FROM mvp_workbook_meaning_attributes WHERE workbook_id = ?").run(workbookId);
      const insert = this.db.prepare("INSERT INTO mvp_workbook_meaning_attributes (workbook_id, position, label, language_code) VALUES (?, ?, ?, ?)");
      for (const attribute of attributes) insert.run(workbookId, attribute.position, attribute.label, attribute.languageCode);
      this.ensureMetadataBackfill();
    });
    const updated = this.getWorkbook(workbookId);
    if (!updated) throw new Error(`Workbook with id ${workbookId} was not found.`);
    return updated;
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

  getTierColorsEnabled(): boolean {
    return this.getMeta("tier_colors_enabled") !== "0";
  }

  setTierColorsEnabled(enabled: boolean): boolean {
    this.setMeta("tier_colors_enabled", enabled ? "1" : "0");
    return enabled;
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

  addEntry(workbookId: number, vocabulary: string, meaning: string, meanings?: string[], attributes: Record<string, string> = {}, posTagIds: number[] = []): EntryRow {
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
      this.saveEntryAttributes(entryId, attributes);
      const addTag = this.db.prepare("INSERT OR IGNORE INTO mvp_entry_pos_tags (entry_id, tag_id) VALUES (?, ?)");
      for (const tagId of [...new Set(posTagIds)]) addTag.run(entryId, tagId);
      return entryId;
    });

    const inserted = this.getEntry(result);
    if (!inserted) {
      throw new Error("Could not load inserted entry.");
    }
    return inserted;
  }

  listMetadataAttributes(workbookId: number): MetadataAttribute[] {
    const workbook = this.getWorkbook(workbookId);
    if (!workbook) throw new Error(`Workbook with id ${workbookId} was not found.`);
    return workbook.metadataAttributes;
  }

  updateMetadataAttributes(workbookId: number, attributes: MetadataAttribute[]): WorkbookRow {
    const workbook = this.getWorkbook(workbookId);
    if (!workbook) throw new Error(`Workbook with id ${workbookId} was not found.`);
    const normalized = attributes.map((attribute, index) => ({ ...attribute, key: attribute.key.trim(), label: trimRequired(attribute.label, "Attribute label"), displayOrder: index }));
    if (!normalized.some((a) => a.key === "vocab") || !normalized.some((a) => a.key === "meaning_1")) throw new ValidationError("Vocab and Meaning 1 cannot be removed.");
    if (!normalized.find((a) => a.key === "vocab")?.required || !normalized.find((a) => a.key === "meaning_1")?.required) throw new ValidationError("Vocab and Meaning 1 are required.");
    if (new Set(normalized.map((a) => a.key)).size !== normalized.length) throw new ValidationError("Attribute keys must be unique.");
    this.transaction(() => {
      const save = this.db.prepare("INSERT INTO mvp_workbook_attributes (workbook_id, attribute_key, label, language_code, is_required, is_visible, display_order) VALUES (?, ?, ?, ?, ?, ?, ?) ON CONFLICT(workbook_id, attribute_key) DO UPDATE SET label=excluded.label, language_code=excluded.language_code, is_required=excluded.is_required, is_visible=excluded.is_visible, display_order=excluded.display_order");
      for (const a of normalized) save.run(workbookId, a.key, a.label, a.languageCode, a.required ? 1 : 0, a.visible ? 1 : 0, a.displayOrder);
      const keys = normalized.map((a) => a.key);
      const existing = this.db.prepare("SELECT attribute_key FROM mvp_workbook_attributes WHERE workbook_id = ?").all(workbookId) as Record<string, unknown>[];
      for (const row of existing) if (!keys.includes(String(row.attribute_key))) this.db.prepare("DELETE FROM mvp_workbook_attributes WHERE workbook_id = ? AND attribute_key = ?").run(workbookId, String(row.attribute_key));
    });
    return this.getWorkbook(workbookId)!;
  }

  listPosTags(workbookId: number): PosTag[] {
    const workbook = this.getWorkbook(workbookId);
    if (!workbook) throw new Error(`Workbook with id ${workbookId} was not found.`);
    if (!["JP", "EN"].includes(workbook.vocabularyLanguageCode ?? "")) return [];
    const rows = this.db.prepare("SELECT id, name, is_predefined FROM mvp_pos_tags WHERE workbook_id = ? ORDER BY name").all(workbookId) as Record<string, unknown>[];
    return rows.map((row) => ({ id: Number(row.id), name: String(row.name), predefined: Number(row.is_predefined) === 1 }));
  }

  addPosTag(workbookId: number, name: string): PosTag {
    const clean = trimRequired(name, "POS tag");
    this.ensurePosSupported(workbookId);
    const result = this.db.prepare("INSERT INTO mvp_pos_tags (workbook_id, name, is_predefined) VALUES (?, ?, 0)").run(workbookId, clean);
    return { id: Number(result.lastInsertRowid), name: clean, predefined: false };
  }

  renamePosTag(tagId: number, name: string): void { this.db.prepare("UPDATE mvp_pos_tags SET name = ? WHERE id = ?").run(trimRequired(name, "POS tag"), tagId); }
  deletePosTag(tagId: number): void { this.db.prepare("DELETE FROM mvp_pos_tags WHERE id = ?").run(tagId); }
  setEntryPosTags(entryId: number, tagIds: number[]): void {
    this.db.prepare("DELETE FROM mvp_entry_pos_tags WHERE entry_id = ?").run(entryId);
    const add = this.db.prepare("INSERT OR IGNORE INTO mvp_entry_pos_tags (entry_id, tag_id) VALUES (?, ?)");
    for (const id of [...new Set(tagIds)]) add.run(entryId, id);
  }

  private saveEntryAttributes(entryId: number, attributes: Record<string, string>): void {
    const save = this.db.prepare("INSERT INTO mvp_entry_attributes (entry_id, attribute_key, value) VALUES (?, ?, ?) ON CONFLICT(entry_id, attribute_key) DO UPDATE SET value=excluded.value");
    const remove = this.db.prepare("DELETE FROM mvp_entry_attributes WHERE entry_id = ? AND attribute_key = ?");
    for (const [key, value] of Object.entries(attributes)) {
      const clean = value.trim();
      if (clean) save.run(entryId, key, clean); else remove.run(entryId, key);
    }
  }

  private ensurePosSupported(workbookId: number): void {
    const workbook = this.getWorkbook(workbookId);
    if (!workbook || !["JP", "EN"].includes(workbook.vocabularyLanguageCode ?? "")) throw new ValidationError("POS tags are unavailable for this language.");
  }

  updateEntry(entryId: number, vocabulary: string, meaning: string, meanings?: string[], attributes?: Record<string, string>, posTagIds?: number[]): EntryRow {
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
      if (attributes) this.saveEntryAttributes(entryId, attributes);
      if (posTagIds) {
        this.db.prepare("DELETE FROM mvp_entry_pos_tags WHERE entry_id = ?").run(entryId);
        const addTag = this.db.prepare("INSERT OR IGNORE INTO mvp_entry_pos_tags (entry_id, tag_id) VALUES (?, ?)");
        for (const tagId of [...new Set(posTagIds)]) addTag.run(entryId, tagId);
      }
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
    const definitions = this.db.prepare(
      "SELECT attribute_key, label, language_code, is_required, is_visible, display_order FROM mvp_workbook_attributes WHERE workbook_id = ? ORDER BY display_order ASC, attribute_key ASC",
    ).all(workbook.id) as Record<string, unknown>[];
    if (definitions.length === 0) {
      const defaults = [
        { key: "vocab", label: workbook.vocabularyLabel, languageCode: workbook.vocabularyLanguageCode, required: true, visible: true, order: 0 },
        { key: "meaning_1", label: workbook.meaningAttributes[0]?.label ?? "Meaning 1", languageCode: workbook.meaningAttributes[0]?.languageCode ?? null, required: true, visible: true, order: 1 },
      ];
      if (workbook.presetEnabled && workbook.vocabularyLanguageCode === "JP") {
        defaults.push({ key: "kana", label: "Kana", languageCode: "JP", required: false, visible: false, order: 2 });
        defaults.push({ key: "example_1", label: "Example 1", languageCode: "JP", required: false, visible: false, order: 3 });
        defaults.push({ key: "example_2", label: "Example 2", languageCode: "JP", required: false, visible: false, order: 4 });
      }
      const insert = this.db.prepare("INSERT OR IGNORE INTO mvp_workbook_attributes (workbook_id, attribute_key, label, language_code, is_required, is_visible, display_order) VALUES (?, ?, ?, ?, ?, ?, ?)");
      for (const item of defaults) insert.run(workbook.id, item.key, item.label, item.languageCode, item.required ? 1 : 0, item.visible ? 1 : 0, item.order);
    }
    if (workbook.presetEnabled && workbook.vocabularyLanguageCode === "JP") {
      const count = Number((this.db.prepare("SELECT COUNT(*) AS count FROM mvp_pos_tags WHERE workbook_id = ?").get(workbook.id) as Record<string, unknown>).count ?? 0);
      if (count === 0) {
        const insertTag = this.db.prepare("INSERT OR IGNORE INTO mvp_pos_tags (workbook_id, name, is_predefined) VALUES (?, ?, 1)");
        for (const tag of JAPANESE_POS_TAGS) insertTag.run(workbook.id, tag);
      }
    }
    const metadataRows = this.db.prepare(
      "SELECT attribute_key, label, language_code, is_required, is_visible, display_order FROM mvp_workbook_attributes WHERE workbook_id = ? ORDER BY display_order ASC, attribute_key ASC",
    ).all(workbook.id) as Record<string, unknown>[];
    return {
      ...workbook,
      meaningAttributes: rows.map((row) => ({
        position: Number(row.position),
        label: String(row.label),
        languageCode: row.language_code == null ? null : String(row.language_code),
          })),
      metadataAttributes: metadataRows.map((row) => ({
        key: String(row.attribute_key), label: String(row.label), languageCode: row.language_code == null ? null : String(row.language_code),
        required: Number(row.is_required) === 1, visible: Number(row.is_visible) === 1, displayOrder: Number(row.display_order),
      })),
    };
  }

  private withEntryMeanings(entry: EntryRow): EntryRow {
    const rows = this.db.prepare("SELECT position, value FROM mvp_entry_meanings WHERE entry_id = ? ORDER BY position ASC").all(entry.id) as Record<string, unknown>[];
    const meanings = rows.length > 0 ? rows.map((row) => String(row.value ?? "")) : [entry.meaning];
    const attributes = this.db.prepare("SELECT attribute_key, value FROM mvp_entry_attributes WHERE entry_id = ?").all(entry.id) as Record<string, unknown>[];
    const attributeMap: Record<string, string> = {};
    for (const row of attributes) attributeMap[String(row.attribute_key)] = String(row.value ?? "");
    if (entry.kanaText) attributeMap.kana ??= entry.kanaText;
    const tags = this.db.prepare("SELECT t.id, t.name, t.is_predefined FROM mvp_entry_pos_tags et JOIN mvp_pos_tags t ON t.id = et.tag_id WHERE et.entry_id = ? ORDER BY t.name").all(entry.id) as Record<string, unknown>[];
    const stats = this.db.prepare("SELECT test_count, error_count, last_tested FROM mvp_entry_stats WHERE entry_id = ?").get(entry.id) as Record<string, unknown> | undefined;
    const testCount = Math.max(0, Number(stats?.test_count ?? 0));
    const errorCount = Math.min(3, Math.max(0, Number(stats?.error_count ?? 0)));
    return { ...entry, meanings, meaning: meanings[0] ?? entry.meaning, attributes: attributeMap, posTags: tags.map((row) => ({ id: Number(row.id), name: String(row.name), predefined: Number(row.is_predefined) === 1 })), testCount, errorCount, tier: tierFor(testCount, errorCount), lastTested: stats?.last_tested == null ? null : String(stats.last_tested) };
  }

  getEntryStats(entryId: number): { testCount: number; errorCount: number; tier: EntryRow["tier"]; lastTested: string | null } {
    const entry = this.getEntry(entryId);
    if (!entry) throw new Error(`Entry with id ${entryId} was not found.`);
    return { testCount: entry.testCount, errorCount: entry.errorCount, tier: entry.tier, lastTested: entry.lastTested };
  }

  recordTestResult(entryId: number, isCorrect: boolean): EntryRow {
    if (!this.getEntry(entryId)) throw new Error(`Entry with id ${entryId} was not found.`);
    this.db.prepare("INSERT INTO mvp_entry_stats (entry_id, test_count, error_count, last_tested) VALUES (?, 0, 0, CURRENT_TIMESTAMP) ON CONFLICT(entry_id) DO NOTHING").run(entryId);
    this.db.prepare("UPDATE mvp_entry_stats SET test_count = test_count + 1, error_count = MIN(error_count + ?, 3), last_tested = CURRENT_TIMESTAMP WHERE entry_id = ?").run(isCorrect ? 0 : 1, entryId);
    return this.getEntry(entryId)!;
  }

  increasePriority(entryId: number): EntryRow { return this.adjustPriority(entryId, true); }
  decreasePriority(entryId: number): EntryRow { return this.adjustPriority(entryId, false); }

  private adjustPriority(entryId: number, increase: boolean): EntryRow {
    const entry = this.getEntry(entryId);
    if (!entry) throw new Error(`Entry with id ${entryId} was not found.`);
    if (entry.testCount === 0) throw new ValidationError("Untested entries cannot have their priority adjusted.");
    const next = increase ? (entry.errorCount === 0 ? 1 : 3) : (entry.errorCount >= 3 ? 2 : 0);
    this.db.prepare("INSERT INTO mvp_entry_stats (entry_id, test_count, error_count, last_tested) VALUES (?, ?, ?, ?) ON CONFLICT(entry_id) DO UPDATE SET error_count = excluded.error_count").run(entryId, entry.testCount, next, entry.lastTested);
    return this.getEntry(entryId)!;
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

  private ensureMetadataBackfill(): void {
    const workbooks = this.db.prepare("SELECT id, vocabulary_language_code, preset_enabled FROM mvp_workbooks ORDER BY id ASC").all() as Record<string, unknown>[];
    const add = this.db.prepare("INSERT OR IGNORE INTO mvp_workbook_attributes (workbook_id, attribute_key, label, language_code, is_required, is_visible, display_order) VALUES (?, ?, ?, ?, ?, ?, ?)");
    for (const row of workbooks) {
      const workbookId = Number(row.id);
      const meaningRows = this.db.prepare("SELECT position, label, language_code FROM mvp_workbook_meaning_attributes WHERE workbook_id = ? ORDER BY position ASC").all(workbookId) as Record<string, unknown>[];
      add.run(workbookId, "vocab", "Vocabulary", String(row.vocabulary_language_code ?? "") || null, 1, 1, 0);
      const meanings = meaningRows.length > 0 ? meaningRows : [{ position: 1, label: "Meaning 1", language_code: null }];
      for (const meaning of meanings) {
        const position = Number(meaning.position);
        add.run(workbookId, `meaning_${position}`, String(meaning.label ?? `Meaning ${position}`), meaning.language_code == null ? null : String(meaning.language_code), position === 1 ? 1 : 0, position === 1 ? 1 : 0, position);
        this.db.prepare("UPDATE mvp_workbook_attributes SET is_required = ?, is_visible = CASE WHEN attribute_key = 'meaning_1' THEN 1 ELSE is_visible END WHERE workbook_id = ? AND attribute_key = ?").run(position === 1 ? 1 : 0, workbookId, `meaning_${position}`);
      }
      const language = String(row.vocabulary_language_code ?? "").toUpperCase();
      const optional = [
        ["kana", "Kana", "JP", 2],
        ["example_1", "Example 1", "JP", 3],
        ["example_2", "Example 2", "JP", 4],
      ] as const;
      if (language === "JP" && (Number(row.preset_enabled) === 1 || !this.hasAttribute(workbookId, "kana"))) {
        this.db.prepare("UPDATE mvp_workbooks SET preset_enabled = 1 WHERE id = ?").run(workbookId);
        for (const [key, label, code, order] of optional) add.run(workbookId, key, label, code, 0, 0, order);
      }
      this.db.prepare("UPDATE mvp_entries SET kana_text = kana_text WHERE workbook_id = ?").run(workbookId);
    }
  }

  private ensureStatsBackfill(): void {
    if (this.getMeta("stats_import_v2_complete") === "1") return;
    if (!this.tableExists("vocab_stats") || !this.tableExists("vocab_entries")) {
      this.setMeta("stats_import_v2_complete", "1");
      return;
    }

    // Legacy and MVP entry IDs are not guaranteed to match because the MVP
    // importer creates new rows. Match by the stable vocabulary/meaning pair
    // so a newly-created MVP entry cannot inherit another entry's stats.
    const legacyEntries = this.db.prepare("SELECT id, japanese_text, english_text FROM vocab_entries").all() as Record<string, unknown>[];
    const legacyStats = this.db.prepare("SELECT entry_id, test_count, error_count, last_tested FROM vocab_stats").all() as Record<string, unknown>[];
    const statsByLegacyId = new Map<number, Record<string, unknown>>();
    for (const row of legacyStats) statsByLegacyId.set(Number(row.entry_id), row);
    const legacyByKey = new Map<string, Record<string, unknown>>();
    for (const row of legacyEntries) {
      const key = `${String(row.japanese_text ?? "").trim().toLocaleLowerCase()}\u0000${String(row.english_text ?? "").trim().toLocaleLowerCase()}`;
      if (!legacyByKey.has(key)) legacyByKey.set(key, row);
    }

    this.db.prepare("DELETE FROM mvp_entry_stats").run();
    const mvpEntries = this.db.prepare("SELECT id, vocabulary, meaning FROM mvp_entries").all() as Record<string, unknown>[];
    const insert = this.db.prepare("INSERT INTO mvp_entry_stats (entry_id, test_count, error_count, last_tested) VALUES (?, ?, ?, ?)");
    for (const entry of mvpEntries) {
      const key = `${String(entry.vocabulary ?? "").trim().toLocaleLowerCase()}\u0000${String(entry.meaning ?? "").trim().toLocaleLowerCase()}`;
      const legacy = legacyByKey.get(key);
      const stats = legacy ? statsByLegacyId.get(Number(legacy.id)) : undefined;
      if (!stats) continue;
      insert.run(Number(entry.id), Math.max(0, Number(stats.test_count ?? 0)), Math.min(3, Math.max(0, Number(stats.error_count ?? 0))), stats.last_tested == null ? null : String(stats.last_tested));
    }
    this.setMeta("stats_import_v2_complete", "1");
  }

  private hasAttribute(workbookId: number, key: string): boolean {
    return Boolean(this.db.prepare("SELECT 1 FROM mvp_workbook_attributes WHERE workbook_id = ? AND attribute_key = ?").get(workbookId, key));
  }

  private repairLegacyWorkbookSplit(): void {
    if (this.getMeta("legacy_workbook_split_complete") === "1" || !this.tableExists("workbooks") || !this.tableExists("vocab_entries")) return;
    const mvp = this.db.prepare("SELECT id FROM mvp_workbooks ORDER BY id ASC").all() as Record<string, unknown>[];
    const legacy = this.db.prepare("SELECT id, name, target_language_code, target_label, meaning_label, created_at FROM workbooks ORDER BY id ASC").all() as Record<string, unknown>[];
    if (mvp.length !== 1 || legacy.length < 2) {
      this.setMeta("legacy_workbook_split_complete", "1");
      return;
    }
    const firstId = Number(mvp[0].id);
    const ids: number[] = [];
    for (let index = 0; index < legacy.length; index += 1) {
      const row = legacy[index];
      const name = trimRequired(String(row.name ?? `Workbook ${index + 1}`), "Workbook name");
      const vocabularyLabel = trimOptional(String(row.target_label ?? "")) ?? "Vocabulary";
      const languageCode = trimOptional(String(row.target_language_code ?? ""))?.toUpperCase() ?? null;
      const meaningLabel = trimOptional(String(row.meaning_label ?? "")) ?? "Meaning 1";
      const id = index === 0 ? firstId : Number(this.db.prepare("INSERT INTO mvp_workbooks (name, vocabulary_label, vocabulary_language_code, created_at) VALUES (?, ?, ?, ?)").run(name, vocabularyLabel, languageCode, String(row.created_at ?? new Date().toISOString())).lastInsertRowid);
      if (index === 0) this.db.prepare("UPDATE mvp_workbooks SET name = ?, vocabulary_label = ?, vocabulary_language_code = ?, created_at = ? WHERE id = ?").run(name, vocabularyLabel, languageCode, String(row.created_at ?? new Date().toISOString()), id);
      this.db.prepare("DELETE FROM mvp_workbook_meaning_attributes WHERE workbook_id = ?").run(id);
      this.db.prepare("INSERT INTO mvp_workbook_meaning_attributes (workbook_id, position, label, language_code) VALUES (?, 1, ?, NULL)").run(id, meaningLabel);
      ids.push(id);
    }
    const entries = this.db.prepare("SELECT id, workbook_id FROM vocab_entries ORDER BY id ASC").all() as Record<string, unknown>[];
    const mvpEntries = this.db.prepare("SELECT id FROM mvp_entries ORDER BY id ASC").all() as Record<string, unknown>[];
    const update = this.db.prepare("UPDATE mvp_entries SET workbook_id = ? WHERE id = ?");
    for (let index = 0; index < entries.length && index < mvpEntries.length; index += 1) {
      const sourceIndex = legacy.findIndex((row) => Number(row.id) === Number(entries[index].workbook_id));
      update.run(ids[sourceIndex >= 0 ? sourceIndex : 0], Number(mvpEntries[index].id));
    }
    this.setMeta("legacy_workbook_split_complete", "1");
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
