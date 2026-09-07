import { resolve } from "node:path";
import { DatabaseSync } from "node:sqlite";
import { runSchemaMigrations } from "./schema.js";

export type VocabularyKind = "preset_language" | "other_language" | "non_language";
export type MeaningAttribute = { id?: number; key?: string; position: number; label: string; languageCode: string | null };
export type MetadataAttribute = { id?: number; key: string; role?: "vocabulary" | "meaning" | "optional"; label: string; languageCode: string | null; required: boolean; visible: boolean; displayOrder: number; provenance?: "preset" | "custom" };
export type PosTag = { id: number; name: string; predefined: boolean };
export type WorkbookConfigurationInput = {
  name: string;
  vocabularyKind: VocabularyKind;
  vocabularyLabel: string;
  vocabularyLanguageCode: string | null;
  presetEnabled: boolean;
  posEnabled: boolean;
  meaningAttributes: MeaningAttribute[];
  optionalAttributes: MetadataAttribute[];
  posTags: Array<{ id?: number; name: string; predefined: boolean }>;
};
export type CreateWorkbookInput = WorkbookConfigurationInput;
export type WorkbookUpdateImpact = { populatedFields: Array<{ key: string; label: string; valueCount: number }> };
export type WorkbookAttributesDraft = { vocabularyLabel: string; fields: MetadataAttribute[] };
export type WorkbookRow = {
  id: number; name: string; wordCount: number; createdAt: string;
  vocabularyLabel: string; vocabularyLanguageCode: string | null;
  presetEnabled: boolean; vocabularyKind: VocabularyKind; posEnabled: boolean;
  meaningAttributes: MeaningAttribute[]; metadataAttributes: MetadataAttribute[];
};
export type EntryRow = {
  id: number; workbookId: number; vocabulary: string; meaning: string; meanings: string[];
  kanaText: string | null; attributes: Record<string, string>; posTags: PosTag[];
  createdAt: string; updatedAt: string; testCount: number; errorCount: number;
  tier: "gray" | "green" | "yellow" | "red"; lastTested: string | null; nextTestDeadline: string | null;
};
export type LanguagePresetDefinition = { optionalAttributes: Array<{ key: string; label: string; languageCode: string | null }>; posTags: string[] };

function exampleFields(languageCode: string): LanguagePresetDefinition["optionalAttributes"] {
  return [
    { key: "example_sentence_1", label: "Example Sentence 1", languageCode },
    { key: "example_sentence_2", label: "Example Sentence 2", languageCode },
  ];
}

export const LANGUAGE_PRESET_DEFINITIONS: Record<string, LanguagePresetDefinition> = {
  JP: {
    optionalAttributes: [{ key: "kana", label: "Kana", languageCode: "JP" }, ...exampleFields("JP")],
    posTags: ["名詞", "固有名詞", "イ形容詞", "ナ形容詞", "動詞 (自動詞)", "動詞 (他動詞)", "副詞", "連体詞", "接続詞", "連語", "その他"],
  },
  EN: { optionalAttributes: exampleFields("EN"), posTags: ["n.", "v.", "adj.", "adv.", "pron.", "prep.", "conj.", "phrase."] },
  DE: {
    optionalAttributes: exampleFields("DE"),
    posTags: ["m. noun - maskulines Substantiv", "f. noun - feminines Substantiv", "n. noun - neutrales Substantiv", "art. - Artikel", "adj. - Adjektiv", "pron. - Pronomen", "num. - Numerale", "adv. - Adverb", "prep. - Präposition", "conj. - Konjunktion", "interj. - Interjektion"],
  },
  ZH: { optionalAttributes: exampleFields("ZH"), posTags: [] },
  KO: { optionalAttributes: exampleFields("KO"), posTags: [] },
  ES: { optionalAttributes: exampleFields("ES"), posTags: [] },
  FR: { optionalAttributes: exampleFields("FR"), posTags: [] },
};

export class WorkbookDataLossError extends Error {
  constructor(readonly impact: WorkbookUpdateImpact) {
    super("This change removes populated workbook fields and requires confirmation.");
  }
}
export class ValidationError extends Error {}

export function defaultDbPath(): string {
  return process.env.VOCAB_HELPER_DB_PATH?.trim() || resolve(process.cwd(), "..", "vocab.db");
}

function trimRequired(value: string, label: string): string {
  const cleaned = value.trim();
  if (!cleaned) throw new ValidationError(`${label} is required.`);
  return cleaned;
}
function trimOptional(value: string | null | undefined): string | null { return value?.trim() || null; }
function tierFor(testCount: number, errorCount: number): EntryRow["tier"] {
  if (testCount <= 0) return "gray";
  if (errorCount <= 0) return "green";
  return errorCount <= 2 ? "yellow" : "red";
}
function transaction<T>(db: DatabaseSync, work: () => T): T {
  db.exec("BEGIN IMMEDIATE");
  try { const result = work(); db.exec("COMMIT"); return result; }
  catch (error) { try { db.exec("ROLLBACK"); } catch { /* Preserve the original error. */ } throw error; }
}

export class VocabularyRepository {
  private readonly db: DatabaseSync;
  private closed = false;

  constructor(private readonly dbPath: string = defaultDbPath()) {
    this.db = new DatabaseSync(dbPath);
    runSchemaMigrations(this.db);
  }
  close(): void { if (!this.closed) this.db.close(); this.closed = true; }
  initialize(): void { runSchemaMigrations(this.db); }

  listWorkbooks(): WorkbookRow[] {
    const rows = this.db.prepare("SELECT w.*, COUNT(e.id) AS word_count FROM workbooks w LEFT JOIN entries e ON e.workbook_id = w.id GROUP BY w.id ORDER BY w.id").all() as Record<string, unknown>[];
    return rows.map((row) => this.hydrateWorkbook(row));
  }
  getWorkbook(workbookId: number): WorkbookRow | null {
    const row = this.db.prepare("SELECT w.*, COUNT(e.id) AS word_count FROM workbooks w LEFT JOIN entries e ON e.workbook_id = w.id WHERE w.id = ? GROUP BY w.id").get(workbookId) as Record<string, unknown> | undefined;
    return row ? this.hydrateWorkbook(row) : null;
  }
  listMeaningAttributes(workbookId: number): MeaningAttribute[] { return this.requireWorkbook(workbookId).meaningAttributes; }

  createConfiguredWorkbook(input: WorkbookConfigurationInput): WorkbookRow {
    const config = this.normalizeConfiguration(input);
    const now = new Date().toISOString();
    const workbookId = transaction(this.db, () => {
      const result = this.db.prepare(`INSERT INTO workbooks
        (name, vocabulary_kind, vocabulary_label, vocabulary_language_code, preset_enabled, pos_enabled, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)`)
        .run(config.name, config.vocabularyKind, config.vocabularyLabel, config.vocabularyLanguageCode, config.presetEnabled ? 1 : 0, config.posEnabled ? 1 : 0, now, now);
      const id = Number(result.lastInsertRowid);
      this.writeFields(id, config);
      this.writeInitialTags(id, config.posTags);
      this.db.prepare("UPDATE app_settings SET current_workbook_id = COALESCE(current_workbook_id, ?) WHERE singleton_id = 1").run(id);
      return id;
    });
    return this.requireWorkbook(workbookId);
  }
  createWorkbook(name: string, vocabularyLabel = "Vocabulary", vocabularyLanguageCode: string | null = null, meaningAttributes: MeaningAttribute[] = [{ position: 1, label: "Meaning 1", languageCode: null }], presetEnabled = vocabularyLanguageCode === "JP"): WorkbookRow {
    const kind: VocabularyKind = vocabularyLanguageCode ? "preset_language" : "non_language";
    const preset = vocabularyLanguageCode ? LANGUAGE_PRESET_DEFINITIONS[vocabularyLanguageCode] : undefined;
    return this.createConfiguredWorkbook({
      name, vocabularyKind: kind, vocabularyLabel, vocabularyLanguageCode, meaningAttributes, presetEnabled,
      posEnabled: kind !== "non_language",
      optionalAttributes: presetEnabled ? (preset?.optionalAttributes ?? []).map((field, index) => ({ ...field, required: false, visible: false, displayOrder: index + 1, provenance: "preset" })) : [],
      posTags: (preset?.posTags ?? []).map((tagName) => ({ name: tagName, predefined: true })),
    });
  }

  previewWorkbookUpdate(workbookId: number, input: WorkbookConfigurationInput): WorkbookUpdateImpact {
    this.requireWorkbook(workbookId);
    const config = this.normalizeConfiguration(input);
    const retainedIds = new Set([...config.meaningAttributes, ...config.optionalAttributes].flatMap((field) => field.id === undefined ? [] : [field.id]));
    const retainedKeys = new Set([...config.meaningAttributes.map((field) => field.key ?? `meaning_${field.position}`), ...config.optionalAttributes.map((field) => field.key)]);
    const rows = this.db.prepare(`SELECT f.id, f.field_key, f.label, COUNT(CASE WHEN trim(v.value) <> '' THEN 1 END) AS value_count
      FROM workbook_fields f LEFT JOIN entry_field_values v ON v.field_id = f.id
      WHERE f.workbook_id = ? GROUP BY f.id ORDER BY f.role, f.position`).all(workbookId) as Record<string, unknown>[];
    return { populatedFields: rows.filter((row) => !retainedIds.has(Number(row.id)) && !retainedKeys.has(String(row.field_key)) && Number(row.value_count) > 0)
      .map((row) => ({ key: String(row.field_key), label: String(row.label), valueCount: Number(row.value_count) })) };
  }

  updateConfiguredWorkbook(workbookId: number, input: WorkbookConfigurationInput, confirmDataLoss = false): WorkbookRow {
    const config = this.normalizeConfiguration(input);
    const current = this.requireWorkbook(workbookId);
    if (config.vocabularyKind !== current.vocabularyKind || config.vocabularyLanguageCode !== current.vocabularyLanguageCode) {
      throw new ValidationError("Workbook type cannot be changed after creation.");
    }
    if (config.presetEnabled !== current.presetEnabled) throw new ValidationError("Preset selection cannot be changed after creation.");
    transaction(this.db, () => {
      const impact = this.previewWorkbookUpdate(workbookId, config);
      if (impact.populatedFields.length > 0 && !confirmDataLoss) throw new WorkbookDataLossError(impact);
      this.db.prepare(`UPDATE workbooks SET name = ?, vocabulary_kind = ?, vocabulary_label = ?, vocabulary_language_code = ?,
        preset_enabled = ?, pos_enabled = ?, updated_at = ? WHERE id = ?`)
        .run(config.name, config.vocabularyKind, config.vocabularyLabel, config.vocabularyLanguageCode, config.presetEnabled ? 1 : 0, config.posEnabled ? 1 : 0, new Date().toISOString(), workbookId);
      this.syncFields(workbookId, config);
      this.syncTags(workbookId, config.posTags);
    });
    return this.requireWorkbook(workbookId);
  }
  updateWorkbookSettings(workbookId: number, name: string, vocabularyLabel: string, vocabularyLanguageCode: string | null, meaningAttributes: MeaningAttribute[], presetEnabled = vocabularyLanguageCode === "JP"): WorkbookRow {
    const current = this.requireWorkbook(workbookId);
    return this.updateConfiguredWorkbook(workbookId, {
      name, vocabularyLabel, vocabularyLanguageCode, meaningAttributes, presetEnabled,
      vocabularyKind: vocabularyLanguageCode ? "preset_language" : current.vocabularyKind, posEnabled: current.posEnabled,
      optionalAttributes: current.metadataAttributes.filter((field) => field.role === "optional"),
      posTags: this.listStoredPosTags(workbookId),
    });
  }

  deleteWorkbook(workbookId: number): number | null {
    this.requireWorkbook(workbookId);
    return transaction(this.db, () => {
      this.db.prepare("DELETE FROM workbooks WHERE id = ?").run(workbookId);
      const current = this.readCurrentWorkbookId();
      const next = current ?? this.firstWorkbookId();
      this.db.prepare("UPDATE app_settings SET current_workbook_id = ? WHERE singleton_id = 1").run(next);
      return next;
    });
  }
  getCurrentWorkbookId(): number | null { return this.readCurrentWorkbookId() ?? this.firstWorkbookId(); }
  setCurrentWorkbookId(workbookId: number): WorkbookRow {
    const workbook = this.requireWorkbook(workbookId);
    this.db.prepare("UPDATE app_settings SET current_workbook_id = ? WHERE singleton_id = 1").run(workbookId);
    return workbook;
  }

  listEntries(workbookId?: number): EntryRow[] {
    const resolved = workbookId ?? this.getCurrentWorkbookId();
    if (resolved === null) return [];
    this.requireWorkbook(resolved);
    const ids = this.db.prepare("SELECT id FROM entries WHERE workbook_id = ? ORDER BY id").all(resolved) as Array<{ id: number }>;
    return ids.map((row) => this.getEntry(Number(row.id))!);
  }
  countEntries(workbookId?: number): number {
    const row = workbookId === undefined ? this.db.prepare("SELECT COUNT(*) AS count FROM entries").get() : this.db.prepare("SELECT COUNT(*) AS count FROM entries WHERE workbook_id = ?").get(workbookId);
    return Number((row as { count?: number } | undefined)?.count ?? 0);
  }
  getTierColorsEnabled(): boolean {
    return Number((this.db.prepare("SELECT tier_colors_enabled FROM app_settings WHERE singleton_id = 1").get() as { tier_colors_enabled: number }).tier_colors_enabled) === 1;
  }
  setTierColorsEnabled(enabled: boolean): boolean {
    this.db.prepare("UPDATE app_settings SET tier_colors_enabled = ? WHERE singleton_id = 1").run(enabled ? 1 : 0);
    return enabled;
  }

  getEntry(entryId: number): EntryRow | null {
    const row = this.db.prepare("SELECT e.*, s.test_count, s.error_count, s.last_tested, s.next_test_deadline FROM entries e JOIN entry_stats s ON s.entry_id = e.id WHERE e.id = ?").get(entryId) as Record<string, unknown> | undefined;
    if (!row) return null;
    const workbookId = Number(row.workbook_id);
    const fields = this.db.prepare(`SELECT f.field_key, f.role, f.position, v.value FROM workbook_fields f
      LEFT JOIN entry_field_values v ON v.field_id = f.id AND v.entry_id = ? WHERE f.workbook_id = ? ORDER BY f.role, f.position`).all(entryId, workbookId) as Record<string, unknown>[];
    const meanings = fields.filter((field) => field.role === "meaning").map((field) => String(field.value ?? ""));
    const attributes: Record<string, string> = {};
    for (const field of fields.filter((item) => item.role === "optional")) attributes[String(field.field_key)] = String(field.value ?? "");
    const tags = this.db.prepare("SELECT t.id, t.name, t.is_predefined FROM entry_pos_tags et JOIN pos_tags t ON t.id = et.tag_id WHERE et.entry_id = ? ORDER BY t.name").all(entryId) as Record<string, unknown>[];
    const testCount = Number(row.test_count); const errorCount = Number(row.error_count);
    return {
      id: Number(row.id), workbookId, vocabulary: String(row.vocabulary), meaning: meanings[0] ?? "", meanings,
      kanaText: attributes.kana || null, attributes,
      posTags: tags.map((tag) => ({ id: Number(tag.id), name: String(tag.name), predefined: Number(tag.is_predefined) === 1 })),
      createdAt: String(row.created_at), updatedAt: String(row.updated_at), testCount, errorCount, tier: tierFor(testCount, errorCount),
      lastTested: row.last_tested == null ? null : String(row.last_tested), nextTestDeadline: row.next_test_deadline == null ? null : String(row.next_test_deadline),
    };
  }

  addEntry(workbookId: number, vocabulary: string, meaning: string, meanings?: string[], attributes: Record<string, string> = {}, posTagIds: number[] = []): EntryRow {
    const workbook = this.requireWorkbook(workbookId);
    const values = this.normalizeEntryMeanings(workbook, meanings?.length ? meanings : [meaning]);
    this.validateEntryAssociations(workbookId, attributes, posTagIds);
    const now = new Date().toISOString();
    const id = transaction(this.db, () => {
      const result = this.db.prepare("INSERT INTO entries (workbook_id, vocabulary, created_at, updated_at) VALUES (?, ?, ?, ?)").run(workbookId, trimRequired(vocabulary, "Vocabulary"), now, now);
      const entryId = Number(result.lastInsertRowid);
      this.saveEntryValues(entryId, workbookId, values, attributes); this.saveEntryTags(entryId, workbookId, posTagIds);
      this.db.prepare("INSERT INTO entry_stats (entry_id) VALUES (?)").run(entryId);
      return entryId;
    });
    return this.getEntry(id)!;
  }
  updateEntry(entryId: number, vocabulary: string, meaning: string, meanings?: string[], attributes: Record<string, string> = {}, posTagIds: number[] = []): EntryRow {
    const existing = this.requireEntry(entryId); const workbook = this.requireWorkbook(existing.workbookId);
    const values = this.normalizeEntryMeanings(workbook, meanings?.length ? meanings : [meaning]);
    this.validateEntryAssociations(existing.workbookId, attributes, posTagIds);
    transaction(this.db, () => {
      this.db.prepare("UPDATE entries SET vocabulary = ?, updated_at = ? WHERE id = ?").run(trimRequired(vocabulary, "Vocabulary"), new Date().toISOString(), entryId);
      this.db.prepare("DELETE FROM entry_field_values WHERE entry_id = ?").run(entryId);
      this.saveEntryValues(entryId, existing.workbookId, values, attributes); this.saveEntryTags(entryId, existing.workbookId, posTagIds);
    });
    return this.getEntry(entryId)!;
  }
  deleteEntry(entryId: number): void { this.requireEntry(entryId); this.db.prepare("DELETE FROM entries WHERE id = ?").run(entryId); }

  listMetadataAttributes(workbookId: number): MetadataAttribute[] { return this.requireWorkbook(workbookId).metadataAttributes; }
  previewWorkbookAttributesUpdate(workbookId: number, draft: WorkbookAttributesDraft): WorkbookUpdateImpact {
    this.requireWorkbook(workbookId);
    const normalized = this.normalizeAttributesDraft(workbookId, draft);
    const retainedIds = new Set(normalized.fields.flatMap((field) => field.id === undefined ? [] : [field.id]));
    const rows = this.db.prepare(`SELECT f.id, f.field_key, f.label, COUNT(CASE WHEN trim(v.value) <> '' THEN 1 END) AS value_count
      FROM workbook_fields f LEFT JOIN entry_field_values v ON v.field_id = f.id
      WHERE f.workbook_id = ? GROUP BY f.id ORDER BY f.role, f.position`).all(workbookId) as Record<string, unknown>[];
    return { populatedFields: rows.filter((row) => !retainedIds.has(Number(row.id)) && Number(row.value_count) > 0)
      .map((row) => ({ key: String(row.field_key), label: String(row.label), valueCount: Number(row.value_count) })) };
  }
  updateWorkbookAttributes(workbookId: number, draft: WorkbookAttributesDraft, confirmDataLoss = false): WorkbookRow {
    const normalized = this.normalizeAttributesDraft(workbookId, draft);
    transaction(this.db, () => {
      const impact = this.previewWorkbookAttributesUpdate(workbookId, normalized);
      if (impact.populatedFields.length > 0 && !confirmDataLoss) throw new WorkbookDataLossError(impact);
      const retainedIds = new Set(normalized.fields.flatMap((field) => field.id === undefined ? [] : [field.id]));
      const existing = this.db.prepare("SELECT id FROM workbook_fields WHERE workbook_id = ?").all(workbookId) as Array<{ id: number }>;
      for (const row of existing) if (!retainedIds.has(Number(row.id))) this.db.prepare("DELETE FROM workbook_fields WHERE id = ? AND workbook_id = ?").run(row.id, workbookId);
      const maxPosition = Number((this.db.prepare("SELECT COALESCE(MAX(position), 0) AS value FROM workbook_fields WHERE workbook_id = ?").get(workbookId) as { value: number }).value);
      this.db.prepare("UPDATE workbook_fields SET position = position + ? WHERE workbook_id = ?").run(maxPosition + 100, workbookId);
      const update = this.db.prepare("UPDATE workbook_fields SET position=?, label=?, language_code=?, is_required=?, is_visible=? WHERE id=? AND workbook_id=?");
      const insert = this.db.prepare("INSERT INTO workbook_fields (workbook_id,field_key,role,position,label,language_code,is_required,is_visible,provenance) VALUES (?,?,?,?,?,?,?,?,'custom')");
      for (const field of normalized.fields) {
        if (field.id !== undefined) update.run(field.displayOrder, field.label, field.languageCode, field.required ? 1 : 0, field.visible ? 1 : 0, field.id, workbookId);
        else {
          const result = insert.run(workbookId, field.key, field.role, field.displayOrder, field.label, field.languageCode, field.required ? 1 : 0, field.visible ? 1 : 0);
          const fieldId = Number(result.lastInsertRowid);
          this.db.prepare("INSERT INTO entry_field_values (entry_id,field_id,workbook_id,value) SELECT id,?,?,'' FROM entries WHERE workbook_id=?").run(fieldId, workbookId, workbookId);
        }
      }
      this.db.prepare("UPDATE workbooks SET vocabulary_label=?, updated_at=? WHERE id=?").run(normalized.vocabularyLabel, new Date().toISOString(), workbookId);
    });
    return this.requireWorkbook(workbookId);
  }

  listPosTags(workbookId: number): PosTag[] { return this.requireWorkbook(workbookId).posEnabled ? this.listStoredPosTags(workbookId) : []; }
  listStoredPosTags(workbookId: number): PosTag[] {
    this.requireWorkbook(workbookId);
    const rows = this.db.prepare("SELECT id, name, is_predefined FROM pos_tags WHERE workbook_id = ? ORDER BY name").all(workbookId) as Record<string, unknown>[];
    return rows.map((row) => ({ id: Number(row.id), name: String(row.name), predefined: Number(row.is_predefined) === 1 }));
  }
  addPosTag(workbookId: number, name: string): PosTag {
    this.ensurePosSupported(workbookId); const clean = trimRequired(name, "POS tag");
    const result = this.db.prepare("INSERT INTO pos_tags (workbook_id, name, is_predefined) VALUES (?, ?, 0)").run(workbookId, clean);
    return { id: Number(result.lastInsertRowid), name: clean, predefined: false };
  }
  renamePosTag(tagId: number, name: string): void {
    const result = this.db.prepare("UPDATE pos_tags SET name = ? WHERE id = ?").run(trimRequired(name, "POS tag"), tagId);
    if (Number(result.changes) === 0) throw new Error(`Part-of-speech tag ${tagId} was not found.`);
  }
  deletePosTag(tagId: number): void {
    const result = this.db.prepare("DELETE FROM pos_tags WHERE id = ?").run(tagId);
    if (Number(result.changes) === 0) throw new Error(`Part-of-speech tag ${tagId} was not found.`);
  }
  setEntryPosTags(entryId: number, tagIds: number[]): void {
    const entry = this.requireEntry(entryId); this.validateEntryAssociations(entry.workbookId, {}, tagIds);
    transaction(this.db, () => this.saveEntryTags(entryId, entry.workbookId, tagIds));
  }
  setPosEnabled(workbookId: number, enabled: boolean): WorkbookRow {
    this.requireWorkbook(workbookId); this.db.prepare("UPDATE workbooks SET pos_enabled = ?, updated_at = ? WHERE id = ?").run(enabled ? 1 : 0, new Date().toISOString(), workbookId);
    return this.requireWorkbook(workbookId);
  }

  getEntryStats(entryId: number) {
    const entry = this.requireEntry(entryId);
    return { testCount: entry.testCount, errorCount: entry.errorCount, tier: entry.tier, lastTested: entry.lastTested, nextTestDeadline: entry.nextTestDeadline };
  }
  recordTestResult(entryId: number, isCorrect: boolean, decreaseError = true): EntryRow {
    this.requireEntry(entryId);
    transaction(this.db, () => {
      const errors = Number((this.db.prepare("SELECT error_count FROM entry_stats WHERE entry_id = ?").get(entryId) as { error_count: number }).error_count);
      const now = new Date().toISOString();
      if (isCorrect && decreaseError) {
        const deadline = new Date(Date.now() + [15, 7, 4, 1][errors] * 86400000).toISOString();
        this.db.prepare("UPDATE entry_stats SET test_count = test_count + 1, error_count = MAX(error_count - 1, 0), last_tested = ?, next_test_deadline = ? WHERE entry_id = ?").run(now, deadline, entryId);
      } else if (isCorrect) this.db.prepare("UPDATE entry_stats SET test_count = test_count + 1, last_tested = ? WHERE entry_id = ?").run(now, entryId);
      else this.db.prepare("UPDATE entry_stats SET test_count = test_count + 1, error_count = MIN(error_count + 1, 3), last_tested = ? WHERE entry_id = ?").run(now, entryId);
    });
    return this.requireEntry(entryId);
  }
  increasePriority(entryId: number): EntryRow { return this.adjustPriority(entryId, true); }
  decreasePriority(entryId: number): EntryRow { return this.adjustPriority(entryId, false); }
  selectPracticeCandidates(workbookId: number, count: number): EntryRow[] {
    this.requireWorkbook(workbookId);
    const rows = this.db.prepare(`SELECT e.id FROM entries e JOIN entry_stats s ON s.entry_id = e.id WHERE e.workbook_id = ?
      ORDER BY CASE WHEN s.test_count = 0 THEN 0 WHEN s.next_test_deadline IS NOT NULL AND datetime(s.next_test_deadline) <= datetime('now') THEN 1 ELSE 2 END,
      s.error_count DESC, RANDOM() LIMIT ?`).all(workbookId, Math.max(0, Math.floor(count))) as Array<{ id: number }>;
    return rows.map((row) => this.getEntry(Number(row.id))!);
  }

  private adjustPriority(entryId: number, increase: boolean): EntryRow {
    const entry = this.requireEntry(entryId);
    if (entry.testCount === 0) throw new ValidationError("Untested entries cannot have their priority adjusted.");
    const next = increase ? (entry.errorCount === 0 ? 1 : 3) : (entry.errorCount >= 3 ? 2 : 0);
    this.db.prepare("UPDATE entry_stats SET error_count = ? WHERE entry_id = ?").run(next, entryId);
    return this.requireEntry(entryId);
  }
  private hydrateWorkbook(row: Record<string, unknown>): WorkbookRow {
    const id = Number(row.id);
    const fields = this.db.prepare("SELECT * FROM workbook_fields WHERE workbook_id = ? ORDER BY role, position").all(id) as Record<string, unknown>[];
    const meanings = fields.filter((field) => field.role === "meaning").map((field) => ({ id: Number(field.id), key: String(field.field_key), position: Number(field.position), label: String(field.label), languageCode: field.language_code == null ? null : String(field.language_code) }));
    const vocabularyLabel = String(row.vocabulary_label); const vocabularyLanguageCode = row.vocabulary_language_code == null ? null : String(row.vocabulary_language_code);
    const metadata: MetadataAttribute[] = [
      { key: "vocab", role: "vocabulary", label: vocabularyLabel, languageCode: vocabularyLanguageCode, required: true, visible: true, displayOrder: 0, provenance: "custom" },
      ...fields.map((field) => ({
        id: Number(field.id), key: String(field.field_key), role: String(field.role) as "meaning" | "optional", label: String(field.label), languageCode: field.language_code == null ? null : String(field.language_code),
        required: Number(field.is_required) === 1, visible: Number(field.is_visible) === 1,
        displayOrder: field.role === "meaning" ? Number(field.position) : meanings.length + Number(field.position), provenance: String(field.provenance) as "preset" | "custom",
      })),
    ];
    return {
      id, name: String(row.name), wordCount: Number(row.word_count ?? 0), createdAt: String(row.created_at), vocabularyLabel, vocabularyLanguageCode,
      presetEnabled: Number(row.preset_enabled) === 1, vocabularyKind: String(row.vocabulary_kind) as VocabularyKind, posEnabled: Number(row.pos_enabled) === 1,
      meaningAttributes: meanings, metadataAttributes: metadata,
    };
  }
  private normalizeConfiguration(input: WorkbookConfigurationInput): WorkbookConfigurationInput {
    const kinds: VocabularyKind[] = ["preset_language", "other_language", "non_language"];
    if (!kinds.includes(input.vocabularyKind)) throw new ValidationError("Choose a valid vocabulary type.");
    const code = input.vocabularyKind === "preset_language" ? trimRequired(input.vocabularyLanguageCode ?? "", "Preset language").toUpperCase() : null;
    if (input.vocabularyKind === "preset_language" && !LANGUAGE_PRESET_DEFINITIONS[code!]) throw new ValidationError("Choose a supported language preset.");
    const meanings = this.normalizeMeaningAttributes(input.meaningAttributes); const keys = new Set<string>();
    const optional = input.optionalAttributes.map((field, index) => {
      const key = trimRequired(field.key, "Attribute key").toLowerCase();
      if (key === "vocab" || key.startsWith("meaning_") || keys.has(key)) throw new ValidationError("Optional attribute keys must be unique.");
      keys.add(key);
      return { ...field, key, label: trimRequired(field.label, "Attribute label"), languageCode: trimOptional(field.languageCode)?.toUpperCase() ?? null, required: false, displayOrder: index + 1, provenance: field.provenance ?? "custom" };
    });
    const tagNames = new Set<string>();
    const tags = input.posTags.map((tag) => ({ ...tag, name: trimRequired(tag.name, "POS tag") })).filter((tag) => { const key = tag.name.toLocaleLowerCase(); if (tagNames.has(key)) return false; tagNames.add(key); return true; });
    return {
      ...input, name: trimRequired(input.name, "Workbook name"), vocabularyLabel: trimRequired(input.vocabularyLabel, "Vocabulary label"), vocabularyLanguageCode: code,
      presetEnabled: input.vocabularyKind === "preset_language" && input.presetEnabled, posEnabled: input.vocabularyKind === "non_language" ? false : input.posEnabled,
      meaningAttributes: meanings, optionalAttributes: optional, posTags: tags,
    };
  }
  private normalizeMeaningAttributes(attributes: MeaningAttribute[]): MeaningAttribute[] {
    if (attributes.length < 1 || attributes.length > 5) throw new ValidationError("Meaning attributes must contain between 1 and 5 items.");
    const fields = attributes.map((field, index) => ({ ...field, position: index + 1, label: trimRequired(field.label, `Meaning ${index + 1} label`), languageCode: trimOptional(field.languageCode)?.toUpperCase() ?? null }));
    if (new Set(fields.map((field) => field.label.toLocaleLowerCase())).size !== fields.length) throw new ValidationError("Meaning attribute labels must be unique.");
    return fields;
  }
  private normalizeAttributesDraft(workbookId: number, draft: WorkbookAttributesDraft): WorkbookAttributesDraft {
    const vocabularyLabel = trimRequired(draft.vocabularyLabel, "Vocabulary label");
    const existingRows = this.db.prepare("SELECT id, field_key, role FROM workbook_fields WHERE workbook_id = ?").all(workbookId) as Array<{ id: number; field_key: string; role: "meaning" | "optional" }>;
    const existingById = new Map(existingRows.map((row) => [Number(row.id), row]));
    const existingByKey = new Map(existingRows.map((row) => [String(row.field_key), row]));
    const usedIds = new Set<number>(); const usedKeys = new Set<string>();
    const meanings = draft.fields.filter((field) => field.role === "meaning");
    const optional = draft.fields.filter((field) => field.role === "optional");
    if (meanings.length < 1 || meanings.length > 5) throw new ValidationError("Meaning attributes must contain between 1 and 5 items.");
    for (const [section, fields] of [["Meaning", meanings], ["Optional", optional]] as const) {
      const labels = fields.map((field) => trimRequired(field.label, `${section} attribute label`).toLocaleLowerCase());
      if (new Set(labels).size !== labels.length) throw new ValidationError(`${section} attribute labels must be unique.`);
    }
    const normalizeFields = (fields: MetadataAttribute[], role: "meaning" | "optional") => fields.map((field, index) => {
      const id = field.id;
      if (id !== undefined) {
        const existing = existingById.get(id);
        if (!existing || existing.role !== role || existing.field_key !== field.key) throw new ValidationError(`Attribute '${field.label}' does not belong to this workbook.`);
        if (usedIds.has(id)) throw new ValidationError("Attribute IDs must be unique.");
        usedIds.add(id);
      }
      const key = trimRequired(field.key, "Attribute key");
      if (id === undefined && existingByKey.has(key)) throw new ValidationError(`Attribute key '${key}' already belongs to an existing field.`);
      if (usedKeys.has(key)) throw new ValidationError("Attribute keys must be unique.");
      usedKeys.add(key);
      const firstMeaning = role === "meaning" && index === 0;
      return { ...field, id, key, role, label: trimRequired(field.label, `${role === "meaning" ? "Meaning" : "Optional"} attribute label`), languageCode: trimOptional(field.languageCode)?.toUpperCase() ?? null,
        required: firstMeaning, visible: firstMeaning ? true : Boolean(field.visible), displayOrder: index + 1 };
    });
    return { vocabularyLabel, fields: [...normalizeFields(meanings, "meaning"), ...normalizeFields(optional, "optional")] };
  }
  private writeFields(workbookId: number, config: WorkbookConfigurationInput): void {
    const insert = this.db.prepare("INSERT INTO workbook_fields (workbook_id, field_key, role, position, label, language_code, is_required, is_visible, provenance) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)");
    for (const field of config.meaningAttributes) insert.run(workbookId, field.key ?? `meaning_${field.position}`, "meaning", field.position, field.label, field.languageCode, field.position === 1 ? 1 : 0, field.position === 1 ? 1 : 0, "custom");
    config.optionalAttributes.forEach((field, index) => insert.run(workbookId, field.key, "optional", index + 1, field.label, field.languageCode, 0, field.visible ? 1 : 0, field.provenance ?? "custom"));
  }
  private syncFields(workbookId: number, config: WorkbookConfigurationInput): void {
    const existing = this.db.prepare("SELECT id, field_key, role FROM workbook_fields WHERE workbook_id = ?").all(workbookId) as Array<{ id: number; field_key: string; role: "meaning" | "optional" }>;
    const existingById = new Map(existing.map((row) => [Number(row.id), row]));
    const existingByKey = new Map(existing.map((row) => [String(row.field_key), row]));
    const desired = [
      ...config.meaningAttributes.map((field) => ({ field, role: "meaning" as const, key: field.key ?? `meaning_${field.position}` })),
      ...config.optionalAttributes.map((field) => ({ field, role: "optional" as const, key: field.key })),
    ];
    const resolved = desired.map((item) => {
      const row = item.field.id === undefined ? existingByKey.get(item.key) : existingById.get(item.field.id);
      if (row && row.role !== item.role) throw new ValidationError(`Attribute '${item.field.label}' cannot change sections.`);
      if (item.field.id !== undefined && (!row || row.field_key !== item.key)) throw new ValidationError(`Attribute '${item.field.label}' does not belong to this workbook.`);
      return { ...item, row };
    });
    const retainedIds = new Set(resolved.flatMap((item) => item.row ? [Number(item.row.id)] : []));
    for (const row of existing) if (!retainedIds.has(Number(row.id))) this.db.prepare("DELETE FROM workbook_fields WHERE id = ? AND workbook_id = ?").run(row.id, workbookId);
    const maxPosition = Number((this.db.prepare("SELECT COALESCE(MAX(position), 0) AS value FROM workbook_fields WHERE workbook_id = ?").get(workbookId) as { value: number }).value);
    this.db.prepare("UPDATE workbook_fields SET position = position + ? WHERE workbook_id = ?").run(maxPosition + 100, workbookId);
    const update = this.db.prepare("UPDATE workbook_fields SET position=?, label=?, language_code=?, is_required=?, is_visible=? WHERE id=? AND workbook_id=?");
    const insert = this.db.prepare("INSERT INTO workbook_fields (workbook_id,field_key,role,position,label,language_code,is_required,is_visible,provenance) VALUES (?,?,?,?,?,?,?,?,?)");
    const save = (field: MetadataAttribute, role: "meaning" | "optional", position: number, key: string, required: boolean, visible: boolean, existingId?: number) => {
      if (existingId !== undefined) update.run(position, field.label, field.languageCode, required ? 1 : 0, visible ? 1 : 0, existingId, workbookId);
      else {
        const result = insert.run(workbookId, key, role, position, field.label, field.languageCode, required ? 1 : 0, visible ? 1 : 0, field.provenance ?? "custom");
        this.db.prepare("INSERT INTO entry_field_values (entry_id,field_id,workbook_id,value) SELECT id,?,?,'' FROM entries WHERE workbook_id=?").run(Number(result.lastInsertRowid), workbookId, workbookId);
      }
    };
    resolved.filter((item) => item.role === "meaning").forEach(({ field, key, row }, index) => {
      const resolvedKey = row?.field_key ?? (field.key || this.nextFieldKey(workbookId, "meaning"));
      save({ ...field, key: resolvedKey, role: "meaning", required: index === 0, visible: index === 0, displayOrder: index + 1 }, "meaning", index + 1, resolvedKey, index === 0, index === 0, row?.id);
    });
    resolved.filter((item) => item.role === "optional").forEach(({ field, key, row }, index) => save(field, "optional", index + 1, row?.field_key ?? key, false, field.visible, row?.id));
  }
  private nextFieldKey(workbookId: number, base: string): string {
    const used = new Set((this.db.prepare("SELECT field_key FROM workbook_fields WHERE workbook_id = ?").all(workbookId) as Array<{ field_key: string }>).map((row) => String(row.field_key)));
    let suffix = 1; let key = `${base}_${suffix}`;
    while (used.has(key)) key = `${base}_${++suffix}`;
    return key;
  }
  private writeInitialTags(workbookId: number, tags: WorkbookConfigurationInput["posTags"]): void {
    const insert = this.db.prepare("INSERT INTO pos_tags (workbook_id, name, is_predefined) VALUES (?, ?, ?)");
    for (const tag of tags) insert.run(workbookId, tag.name, tag.predefined ? 1 : 0);
  }
  private syncTags(workbookId: number, tags: WorkbookConfigurationInput["posTags"]): void {
    const retained = new Set(tags.flatMap((tag) => tag.id === undefined ? [] : [tag.id]));
    const existing = this.db.prepare("SELECT id FROM pos_tags WHERE workbook_id = ?").all(workbookId) as Array<{ id: number }>;
    for (const row of existing) if (!retained.has(Number(row.id))) this.db.prepare("DELETE FROM pos_tags WHERE workbook_id = ? AND id = ?").run(workbookId, row.id);
    for (const tag of tags) {
      if (tag.id === undefined) this.db.prepare("INSERT INTO pos_tags (workbook_id, name, is_predefined) VALUES (?, ?, ?)").run(workbookId, tag.name, tag.predefined ? 1 : 0);
      else {
        const result = this.db.prepare("UPDATE pos_tags SET name = ?, is_predefined = ? WHERE id = ? AND workbook_id = ?").run(tag.name, tag.predefined ? 1 : 0, tag.id, workbookId);
        if (Number(result.changes) === 0) throw new ValidationError(`Part-of-speech tag ${tag.id} does not belong to this workbook.`);
      }
    }
  }
  private normalizeEntryMeanings(workbook: WorkbookRow, input: string[]): string[] {
    const values = workbook.meaningAttributes.map((_, index) => trimOptional(input[index]) ?? "");
    values[0] = trimRequired(values[0] ?? "", workbook.meaningAttributes[0]?.label ?? "Meaning 1"); return values;
  }
  private validateEntryAssociations(workbookId: number, attributes: Record<string, string>, tagIds: number[]): void {
    const fields = new Set((this.db.prepare("SELECT field_key FROM workbook_fields WHERE workbook_id = ? AND role = 'optional'").all(workbookId) as Array<{ field_key: string }>).map((row) => String(row.field_key)));
    for (const key of Object.keys(attributes)) if (!fields.has(key)) throw new ValidationError(`Attribute '${key}' does not belong to this workbook.`);
    for (const tagId of new Set(tagIds)) if (!this.db.prepare("SELECT 1 FROM pos_tags WHERE id = ? AND workbook_id = ?").get(tagId, workbookId)) throw new ValidationError(`Part-of-speech tag ${tagId} does not belong to this workbook.`);
  }
  private saveEntryValues(entryId: number, workbookId: number, meanings: string[], attributes: Record<string, string>): void {
    const fields = this.db.prepare("SELECT id, field_key, role, position FROM workbook_fields WHERE workbook_id = ?").all(workbookId) as Record<string, unknown>[];
    const insert = this.db.prepare("INSERT INTO entry_field_values (entry_id, field_id, workbook_id, value) VALUES (?, ?, ?, ?)");
    for (const field of fields) insert.run(entryId, Number(field.id), workbookId, field.role === "meaning" ? meanings[Number(field.position) - 1] ?? "" : attributes[String(field.field_key)] ?? "");
  }
  private saveEntryTags(entryId: number, workbookId: number, tagIds: number[]): void {
    this.db.prepare("DELETE FROM entry_pos_tags WHERE entry_id = ?").run(entryId);
    const insert = this.db.prepare("INSERT INTO entry_pos_tags (entry_id, tag_id, workbook_id) VALUES (?, ?, ?)");
    for (const tagId of new Set(tagIds)) insert.run(entryId, tagId, workbookId);
  }
  private ensurePosSupported(workbookId: number): void { if (!this.requireWorkbook(workbookId).posEnabled) throw new ValidationError("Part of speech is disabled for this workbook."); }
  private requireWorkbook(workbookId: number): WorkbookRow { const workbook = this.getWorkbook(workbookId); if (!workbook) throw new Error(`Workbook with id ${workbookId} was not found.`); return workbook; }
  private requireEntry(entryId: number): EntryRow { const entry = this.getEntry(entryId); if (!entry) throw new Error(`Entry with id ${entryId} was not found.`); return entry; }
  private readCurrentWorkbookId(): number | null { const row = this.db.prepare("SELECT current_workbook_id FROM app_settings WHERE singleton_id = 1").get() as { current_workbook_id: number | null }; return row.current_workbook_id == null ? null : Number(row.current_workbook_id); }
  private firstWorkbookId(): number | null { const row = this.db.prepare("SELECT id FROM workbooks ORDER BY id LIMIT 1").get() as { id: number } | undefined; return row ? Number(row.id) : null; }
}
