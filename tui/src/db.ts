import { resolve } from "node:path";
import { DatabaseSync } from "node:sqlite";
import { runSchemaMigrations } from "./schema.js";

export type VocabularyKind = "preset_language" | "other_language" | "non_language";
export type MeaningAttribute = { id?: number; key?: string; position: number; label: string; languageCode: string | null };
export type MetadataAttribute = { id?: number; key: string; role?: "vocabulary" | "meaning" | "optional"; label: string; languageCode: string | null; required: boolean; visible: boolean; displayOrder: number; provenance?: "preset" | "custom" };
export type Tag = { id: number; tagTypeId: number; name: string };
export type TagType = { id: number; workbookId: number; name: string; position: number; visible: boolean; tags: Tag[] };
export type TagDraft = { id?: number; name: string };
export type TagTypeDraft = { id?: number; name: string; visible: boolean; tags: TagDraft[] };
export type WorkbookConfigurationInput = {
  name: string;
  vocabularyKind: VocabularyKind;
  vocabularyLabel: string;
  vocabularyLanguageCode: string | null;
  presetEnabled: boolean;
  meaningAttributes: MeaningAttribute[];
  optionalAttributes: MetadataAttribute[];
  tagTypes: TagTypeDraft[];
};
export type CreateWorkbookInput = WorkbookConfigurationInput;
export type WorkbookUpdateImpact = { populatedFields: Array<{ key: string; label: string; valueCount: number }> };
export type MeaningPromotionImpact = { emptyEntryCount: number };
export type WorkbookAttributesDraft = { vocabularyLabel: string; fields: MetadataAttribute[] };
export type WorkbookTagsDraft = { types: TagTypeDraft[] };
export type TagUpdateImpact = { removals: Array<{ typeName: string; tagName: string | null; tagCount: number; assignmentCount: number; entryCount: number }> };
export type WorkbookRow = {
  id: number; name: string; wordCount: number; createdAt: string;
  vocabularyLabel: string; vocabularyLanguageCode: string | null;
  presetEnabled: boolean; vocabularyKind: VocabularyKind;
  meaningAttributes: MeaningAttribute[]; metadataAttributes: MetadataAttribute[];
};
export type EntryRow = {
  id: number; workbookId: number; vocabulary: string; meaning: string; meanings: string[];
  kanaText: string | null; attributes: Record<string, string>; tags: Tag[];
  createdAt: string; updatedAt: string; testCount: number; errorCount: number;
  tier: "gray" | "green" | "yellow" | "red"; lastTested: string | null; nextTestDeadline: string | null;
};
export type LanguagePresetDefinition = { optionalAttributes: Array<{ key: string; label: string; languageCode: string | null }>; partOfSpeechTags: string[] };

function exampleFields(languageCode: string): LanguagePresetDefinition["optionalAttributes"] {
  return [
    { key: "example_sentence_1", label: "Example Sentence 1", languageCode },
    { key: "example_sentence_2", label: "Example Sentence 2", languageCode },
  ];
}

export const LANGUAGE_PRESET_DEFINITIONS: Record<string, LanguagePresetDefinition> = {
  JP: {
    optionalAttributes: [{ key: "kana", label: "Kana", languageCode: "JP" }, ...exampleFields("JP")],
    partOfSpeechTags: ["名詞", "固有名詞", "イ形容詞", "ナ形容詞", "動詞 (自動詞)", "動詞 (他動詞)", "副詞", "連体詞", "接続詞", "連語", "その他"],
  },
  EN: { optionalAttributes: exampleFields("EN"), partOfSpeechTags: ["n.", "v.", "adj.", "adv.", "pron.", "prep.", "conj.", "phrase."] },
  DE: {
    optionalAttributes: exampleFields("DE"),
    partOfSpeechTags: ["m. noun - maskulines Substantiv", "f. noun - feminines Substantiv", "n. noun - neutrales Substantiv", "art. - Artikel", "adj. - Adjektiv", "pron. - Pronomen", "num. - Numerale", "adv. - Adverb", "prep. - Präposition", "conj. - Konjunktion", "interj. - Interjektion"],
  },
  ZH: { optionalAttributes: exampleFields("ZH"), partOfSpeechTags: [] },
  KO: { optionalAttributes: exampleFields("KO"), partOfSpeechTags: [] },
  ES: { optionalAttributes: exampleFields("ES"), partOfSpeechTags: [] },
  FR: { optionalAttributes: exampleFields("FR"), partOfSpeechTags: [] },
};

export class WorkbookDataLossError extends Error {
  constructor(readonly impact: WorkbookUpdateImpact) {
    super("This change removes populated workbook fields and requires confirmation.");
  }
}
export class TagDataLossError extends Error {
  constructor(readonly impact: TagUpdateImpact) {
    super("This change removes assigned tags and requires confirmation.");
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
        (name, vocabulary_kind, vocabulary_label, vocabulary_language_code, preset_enabled, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)`)
        .run(config.name, config.vocabularyKind, config.vocabularyLabel, config.vocabularyLanguageCode, config.presetEnabled ? 1 : 0, now, now);
      const id = Number(result.lastInsertRowid);
      this.writeFields(id, config);
      this.writeInitialTagTypes(id, config.tagTypes);
      this.db.prepare("UPDATE app_settings SET current_workbook_id = COALESCE(current_workbook_id, ?) WHERE singleton_id = 1").run(id);
      return id;
    });
    return this.requireWorkbook(workbookId);
  }
  createWorkbook(name: string, vocabularyLabel = "Vocabulary", vocabularyLanguageCode: string | null = null, meaningAttributes: MeaningAttribute[] = [{ position: 1, label: "Primary Meaning", languageCode: null }], presetEnabled = vocabularyLanguageCode === "JP"): WorkbookRow {
    const kind: VocabularyKind = vocabularyLanguageCode ? "preset_language" : "non_language";
    const preset = vocabularyLanguageCode ? LANGUAGE_PRESET_DEFINITIONS[vocabularyLanguageCode] : undefined;
    return this.createConfiguredWorkbook({
      name, vocabularyKind: kind, vocabularyLabel, vocabularyLanguageCode, meaningAttributes, presetEnabled,
      optionalAttributes: presetEnabled ? (preset?.optionalAttributes ?? []).map((field, index) => ({ ...field, required: false, visible: false, displayOrder: index + 1, provenance: "preset" })) : [],
      tagTypes: kind === "non_language" ? [] : [{ name: "Part of Speech", visible: false, tags: (preset?.partOfSpeechTags ?? []).map((tagName) => ({ name: tagName })) }],
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
        preset_enabled = ?, updated_at = ? WHERE id = ?`)
        .run(config.name, config.vocabularyKind, config.vocabularyLabel, config.vocabularyLanguageCode, config.presetEnabled ? 1 : 0, new Date().toISOString(), workbookId);
      this.syncFields(workbookId, config);
    });
    return this.requireWorkbook(workbookId);
  }
  updateWorkbookSettings(workbookId: number, name: string, vocabularyLabel: string, vocabularyLanguageCode: string | null, meaningAttributes: MeaningAttribute[], presetEnabled = vocabularyLanguageCode === "JP"): WorkbookRow {
    const current = this.requireWorkbook(workbookId);
    return this.updateConfiguredWorkbook(workbookId, {
      name, vocabularyLabel, vocabularyLanguageCode, meaningAttributes, presetEnabled,
      vocabularyKind: vocabularyLanguageCode ? "preset_language" : current.vocabularyKind,
      optionalAttributes: current.metadataAttributes.filter((field) => field.role === "optional"),
      tagTypes: this.listTagTypes(workbookId).map((type) => ({ id: type.id, name: type.name, visible: type.visible, tags: type.tags.map((tag) => ({ id: tag.id, name: tag.name })) })),
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
    const tags = this.db.prepare(`SELECT t.id, t.tag_type_id, t.name FROM entry_tags et
      JOIN tags t ON t.id = et.tag_id JOIN tag_types tt ON tt.id = t.tag_type_id
      WHERE et.entry_id = ? ORDER BY tt.position, t.name`).all(entryId) as Record<string, unknown>[];
    const testCount = Number(row.test_count); const errorCount = Number(row.error_count);
    return {
      id: Number(row.id), workbookId, vocabulary: String(row.vocabulary), meaning: meanings[0] ?? "", meanings,
      kanaText: attributes.kana || null, attributes,
      tags: tags.map((tag) => ({ id: Number(tag.id), tagTypeId: Number(tag.tag_type_id), name: String(tag.name) })),
      createdAt: String(row.created_at), updatedAt: String(row.updated_at), testCount, errorCount, tier: tierFor(testCount, errorCount),
      lastTested: row.last_tested == null ? null : String(row.last_tested), nextTestDeadline: row.next_test_deadline == null ? null : String(row.next_test_deadline),
    };
  }

  addEntry(workbookId: number, vocabulary: string, meaning: string, meanings?: string[], attributes: Record<string, string> = {}, tagIds: number[] = []): EntryRow {
    const workbook = this.requireWorkbook(workbookId);
    const values = this.normalizeEntryMeanings(workbook, meanings?.length ? meanings : [meaning]);
    this.validateEntryAssociations(workbookId, attributes, tagIds);
    const now = new Date().toISOString();
    const id = transaction(this.db, () => {
      const result = this.db.prepare("INSERT INTO entries (workbook_id, vocabulary, created_at, updated_at) VALUES (?, ?, ?, ?)").run(workbookId, trimRequired(vocabulary, "Vocabulary"), now, now);
      const entryId = Number(result.lastInsertRowid);
      this.saveEntryValues(entryId, workbookId, values, attributes); this.saveEntryTags(entryId, workbookId, tagIds);
      this.db.prepare("INSERT INTO entry_stats (entry_id) VALUES (?)").run(entryId);
      return entryId;
    });
    return this.getEntry(id)!;
  }
  updateEntry(entryId: number, vocabulary: string, meaning: string, meanings?: string[], attributes: Record<string, string> = {}, tagIds: number[] = []): EntryRow {
    const existing = this.requireEntry(entryId); const workbook = this.requireWorkbook(existing.workbookId);
    const values = this.normalizeEntryMeanings(workbook, meanings?.length ? meanings : [meaning]);
    this.validateEntryAssociations(existing.workbookId, attributes, tagIds);
    transaction(this.db, () => {
      this.db.prepare("UPDATE entries SET vocabulary = ?, updated_at = ? WHERE id = ?").run(trimRequired(vocabulary, "Vocabulary"), new Date().toISOString(), entryId);
      this.db.prepare("DELETE FROM entry_field_values WHERE entry_id = ?").run(entryId);
      this.saveEntryValues(entryId, existing.workbookId, values, attributes); this.saveEntryTags(entryId, existing.workbookId, tagIds);
    });
    return this.getEntry(entryId)!;
  }
  deleteEntry(entryId: number): void { this.requireEntry(entryId); this.db.prepare("DELETE FROM entries WHERE id = ?").run(entryId); }

  listMetadataAttributes(workbookId: number): MetadataAttribute[] { return this.requireWorkbook(workbookId).metadataAttributes; }
  getMeaningPromotionImpact(workbookId: number, fieldId: number): MeaningPromotionImpact {
    const field = this.db.prepare("SELECT role FROM workbook_fields WHERE id = ? AND workbook_id = ?").get(fieldId, workbookId) as { role: string } | undefined;
    if (!field || field.role !== "meaning") throw new ValidationError(`Meaning attribute ${fieldId} does not belong to this workbook.`);
    const row = this.db.prepare(`SELECT COUNT(*) AS count FROM entries e
      LEFT JOIN entry_field_values v ON v.entry_id = e.id AND v.field_id = ?
      WHERE e.workbook_id = ? AND trim(COALESCE(v.value, '')) = ''`).get(fieldId, workbookId) as { count: number };
    return { emptyEntryCount: Number(row.count) };
  }
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
    transaction(this.db, () => {
      const normalized = this.normalizeAttributesDraft(workbookId, draft);
      this.assertPrimaryMeaningComplete(workbookId, normalized);
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

  listTagTypes(workbookId: number): TagType[] {
    this.requireWorkbook(workbookId);
    const types = this.db.prepare("SELECT id, name, position, is_visible FROM tag_types WHERE workbook_id = ? ORDER BY position").all(workbookId) as Record<string, unknown>[];
    const tags = this.db.prepare("SELECT id, tag_type_id, name FROM tags WHERE workbook_id = ? ORDER BY name").all(workbookId) as Record<string, unknown>[];
    return types.map((type) => ({
      id: Number(type.id), workbookId, name: String(type.name), position: Number(type.position), visible: Number(type.is_visible) === 1,
      tags: tags.filter((tag) => Number(tag.tag_type_id) === Number(type.id)).map((tag) => ({ id: Number(tag.id), tagTypeId: Number(type.id), name: String(tag.name) })),
    }));
  }
  previewWorkbookTagsUpdate(workbookId: number, draft: WorkbookTagsDraft): TagUpdateImpact {
    const normalized = this.normalizeTagsDraft(workbookId, draft);
    const retainedTypeIds = new Set(normalized.types.flatMap((type) => type.id === undefined ? [] : [type.id]));
    const retainedTagIds = new Set(normalized.types.flatMap((type) => type.tags.flatMap((tag) => tag.id === undefined ? [] : [tag.id])));
    const existingTypes = this.listTagTypes(workbookId);
    const assignmentRows = this.db.prepare(`SELECT t.id, COUNT(et.entry_id) AS assignment_count, COUNT(DISTINCT et.entry_id) AS entry_count
      FROM tags t LEFT JOIN entry_tags et ON et.tag_id = t.id WHERE t.workbook_id = ? GROUP BY t.id`).all(workbookId) as Array<{ id: number; assignment_count: number; entry_count: number }>;
    const assignments = new Map(assignmentRows.map((row) => [Number(row.id), { assignmentCount: Number(row.assignment_count), entryCount: Number(row.entry_count) }]));
    const removals: TagUpdateImpact["removals"] = [];
    for (const type of existingTypes) {
      if (!retainedTypeIds.has(type.id)) {
        const assignmentCount = type.tags.reduce((sum, tag) => sum + (assignments.get(tag.id)?.assignmentCount ?? 0), 0);
        const entryRow = this.db.prepare(`SELECT COUNT(DISTINCT et.entry_id) AS count FROM entry_tags et JOIN tags t ON t.id = et.tag_id
          WHERE t.tag_type_id = ?`).get(type.id) as { count: number };
        if (assignmentCount > 0) removals.push({ typeName: type.name, tagName: null, tagCount: type.tags.length, assignmentCount, entryCount: Number(entryRow.count) });
        continue;
      }
      for (const tag of type.tags) {
        if (retainedTagIds.has(tag.id)) continue;
        const counts = assignments.get(tag.id) ?? { assignmentCount: 0, entryCount: 0 };
        if (counts.assignmentCount > 0) removals.push({ typeName: type.name, tagName: tag.name, tagCount: 1, ...counts });
      }
    }
    return { removals };
  }
  updateWorkbookTags(workbookId: number, draft: WorkbookTagsDraft, confirmDataLoss = false): TagType[] {
    transaction(this.db, () => {
      const normalized = this.normalizeTagsDraft(workbookId, draft);
      const impact = this.previewWorkbookTagsUpdate(workbookId, normalized);
      if (impact.removals.length > 0 && !confirmDataLoss) throw new TagDataLossError(impact);
      const retainedTypeIds = new Set(normalized.types.flatMap((type) => type.id === undefined ? [] : [type.id]));
      const retainedTagIds = new Set(normalized.types.flatMap((type) => type.tags.flatMap((tag) => tag.id === undefined ? [] : [tag.id])));
      for (const row of this.db.prepare("SELECT id FROM tag_types WHERE workbook_id = ?").all(workbookId) as Array<{ id: number }>) {
        if (!retainedTypeIds.has(Number(row.id))) this.db.prepare("DELETE FROM tag_types WHERE id = ? AND workbook_id = ?").run(row.id, workbookId);
      }
      for (const row of this.db.prepare("SELECT id FROM tags WHERE workbook_id = ?").all(workbookId) as Array<{ id: number }>) {
        if (!retainedTagIds.has(Number(row.id))) this.db.prepare("DELETE FROM tags WHERE id = ? AND workbook_id = ?").run(row.id, workbookId);
      }
      // Move retained names out of their UNIQUE namespaces so swaps remain atomic.
      for (const typeId of retainedTypeIds) {
        const result = this.db.prepare("UPDATE tag_types SET name = ? WHERE id = ? AND workbook_id = ?").run(`__vocabhelper_stage_${Date.now()}_${Math.random()}_${typeId}__`, typeId, workbookId);
        if (Number(result.changes) !== 1) throw new ValidationError("A tag type changed while it was being saved.");
      }
      for (const tagId of retainedTagIds) {
        const result = this.db.prepare("UPDATE tags SET name = ? WHERE id = ? AND workbook_id = ?").run(`__vocabhelper_stage_${Date.now()}_${Math.random()}_${tagId}__`, tagId, workbookId);
        if (Number(result.changes) !== 1) throw new ValidationError("A tag changed while it was being saved.");
      }
      const offset = Number((this.db.prepare("SELECT COALESCE(MAX(position), 0) AS value FROM tag_types WHERE workbook_id = ?").get(workbookId) as { value: number }).value) + 100;
      this.db.prepare("UPDATE tag_types SET position = position + ? WHERE workbook_id = ?").run(offset, workbookId);
      for (const [index, type] of normalized.types.entries()) {
        let typeId = type.id;
        if (typeId === undefined) typeId = Number(this.db.prepare("INSERT INTO tag_types (workbook_id, name, position, is_visible) VALUES (?, ?, ?, ?)").run(workbookId, type.name, index + 1, type.visible ? 1 : 0).lastInsertRowid);
        else {
          const result = this.db.prepare("UPDATE tag_types SET name = ?, position = ?, is_visible = ? WHERE id = ? AND workbook_id = ?").run(type.name, index + 1, type.visible ? 1 : 0, typeId, workbookId);
          if (Number(result.changes) !== 1) throw new ValidationError("A tag type changed while it was being saved.");
        }
        for (const tag of type.tags) {
          if (tag.id === undefined) this.db.prepare("INSERT INTO tags (tag_type_id, workbook_id, name) VALUES (?, ?, ?)").run(typeId, workbookId, tag.name);
          else {
            const result = this.db.prepare("UPDATE tags SET name = ? WHERE id = ? AND tag_type_id = ? AND workbook_id = ?").run(tag.name, tag.id, typeId, workbookId);
            if (Number(result.changes) !== 1) throw new ValidationError("A tag changed while it was being saved.");
          }
        }
      }
    });
    return this.listTagTypes(workbookId);
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
      presetEnabled: Number(row.preset_enabled) === 1, vocabularyKind: String(row.vocabulary_kind) as VocabularyKind,
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
    const typeNames = new Set<string>();
    const tagTypes = input.tagTypes.map((type) => {
      const name = trimRequired(type.name, "Tag type name"); const typeKey = name.toLocaleLowerCase();
      if (typeNames.has(typeKey)) throw new ValidationError("Tag type names must be unique.");
      typeNames.add(typeKey);
      const tagNames = new Set<string>();
      const tags = type.tags.map((tag) => {
        const tagName = trimRequired(tag.name, "Tag name"); const tagKey = tagName.toLocaleLowerCase();
        if (tagNames.has(tagKey)) throw new ValidationError(`Tag names in ${name} must be unique.`);
        tagNames.add(tagKey); return { ...tag, name: tagName };
      });
      return { ...type, name, visible: Boolean(type.visible), tags };
    });
    return {
      ...input, name: trimRequired(input.name, "Workbook name"), vocabularyLabel: trimRequired(input.vocabularyLabel, "Vocabulary label"), vocabularyLanguageCode: code,
      presetEnabled: input.vocabularyKind === "preset_language" && input.presetEnabled,
      meaningAttributes: meanings, optionalAttributes: optional, tagTypes,
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
  private assertPrimaryMeaningComplete(workbookId: number, draft: WorkbookAttributesDraft): void {
    const proposed = draft.fields.find((field) => field.role === "meaning");
    const current = this.db.prepare("SELECT id FROM workbook_fields WHERE workbook_id = ? AND role = 'meaning' AND position = 1").get(workbookId) as { id: number } | undefined;
    if (proposed?.id === current?.id) return;
    if (proposed?.id === undefined) throw new ValidationError("Save a new meaning before making it the Primary Meaning.");
    const emptyEntryCount = this.getMeaningPromotionImpact(workbookId, proposed.id).emptyEntryCount;
    if (emptyEntryCount > 0) {
      const noun = emptyEntryCount === 1 ? "entry is" : "entries are";
      throw new ValidationError(`Cannot make ${proposed?.label ?? "this field"} the Primary Meaning: ${emptyEntryCount} ${noun} empty.`);
    }
  }
  private normalizeTagsDraft(workbookId: number, draft: WorkbookTagsDraft): WorkbookTagsDraft {
    this.requireWorkbook(workbookId);
    const existingTypes = this.db.prepare("SELECT id FROM tag_types WHERE workbook_id = ?").all(workbookId) as Array<{ id: number }>;
    const existingTypeIds = new Set(existingTypes.map((row) => Number(row.id)));
    const existingTags = this.db.prepare("SELECT id, tag_type_id FROM tags WHERE workbook_id = ?").all(workbookId) as Array<{ id: number; tag_type_id: number }>;
    const existingTagsById = new Map(existingTags.map((row) => [Number(row.id), Number(row.tag_type_id)]));
    const usedTypeIds = new Set<number>(); const typeNames = new Set<string>(); const usedTagIds = new Set<number>();
    const types = draft.types.map((type) => {
      const name = trimRequired(type.name, "Tag type name"); const nameKey = name.toLocaleLowerCase();
      if (typeNames.has(nameKey)) throw new ValidationError("Tag type names must be unique.");
      typeNames.add(nameKey);
      if (type.id !== undefined) {
        if (!existingTypeIds.has(type.id)) throw new ValidationError(`Tag type '${name}' does not belong to this workbook.`);
        if (usedTypeIds.has(type.id)) throw new ValidationError("Tag type IDs must be unique.");
        usedTypeIds.add(type.id);
      }
      const tagNames = new Set<string>();
      const tags = type.tags.map((tag) => {
        const tagName = trimRequired(tag.name, "Tag name"); const tagKey = tagName.toLocaleLowerCase();
        if (tagNames.has(tagKey)) throw new ValidationError(`Tag names in ${name} must be unique.`);
        tagNames.add(tagKey);
        if (tag.id !== undefined) {
          if (type.id === undefined || existingTagsById.get(tag.id) !== type.id) throw new ValidationError(`Tag '${tagName}' does not belong to ${name}.`);
          if (usedTagIds.has(tag.id)) throw new ValidationError("Tag IDs must be unique.");
          usedTagIds.add(tag.id);
        }
        return { ...tag, name: tagName };
      });
      return { ...type, name, tags };
    });
    return { types };
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
  private writeInitialTagTypes(workbookId: number, types: TagTypeDraft[]): void {
    const addType = this.db.prepare("INSERT INTO tag_types (workbook_id, name, position, is_visible) VALUES (?, ?, ?, ?)");
    const addTag = this.db.prepare("INSERT INTO tags (tag_type_id, workbook_id, name) VALUES (?, ?, ?)");
    for (const [index, type] of types.entries()) {
      const typeId = Number(addType.run(workbookId, type.name, index + 1, type.visible ? 1 : 0).lastInsertRowid);
      for (const tag of type.tags) addTag.run(typeId, workbookId, tag.name);
    }
  }
  private normalizeEntryMeanings(workbook: WorkbookRow, input: string[]): string[] {
    const values = workbook.meaningAttributes.map((_, index) => trimOptional(input[index]) ?? "");
    values[0] = trimRequired(values[0] ?? "", workbook.meaningAttributes[0]?.label ?? "Primary Meaning"); return values;
  }
  private validateEntryAssociations(workbookId: number, attributes: Record<string, string>, tagIds: number[]): void {
    const fields = new Set((this.db.prepare("SELECT field_key FROM workbook_fields WHERE workbook_id = ? AND role = 'optional'").all(workbookId) as Array<{ field_key: string }>).map((row) => String(row.field_key)));
    for (const key of Object.keys(attributes)) if (!fields.has(key)) throw new ValidationError(`Attribute '${key}' does not belong to this workbook.`);
    for (const tagId of new Set(tagIds)) if (!this.db.prepare("SELECT 1 FROM tags WHERE id = ? AND workbook_id = ?").get(tagId, workbookId)) throw new ValidationError(`Tag ${tagId} does not belong to this workbook.`);
  }
  private saveEntryValues(entryId: number, workbookId: number, meanings: string[], attributes: Record<string, string>): void {
    const fields = this.db.prepare("SELECT id, field_key, role, position FROM workbook_fields WHERE workbook_id = ?").all(workbookId) as Record<string, unknown>[];
    const insert = this.db.prepare("INSERT INTO entry_field_values (entry_id, field_id, workbook_id, value) VALUES (?, ?, ?, ?)");
    for (const field of fields) insert.run(entryId, Number(field.id), workbookId, field.role === "meaning" ? meanings[Number(field.position) - 1] ?? "" : attributes[String(field.field_key)] ?? "");
  }
  private saveEntryTags(entryId: number, workbookId: number, tagIds: number[]): void {
    this.db.prepare("DELETE FROM entry_tags WHERE entry_id = ?").run(entryId);
    const insert = this.db.prepare("INSERT INTO entry_tags (entry_id, tag_id, workbook_id) VALUES (?, ?, ?)");
    for (const tagId of new Set(tagIds)) insert.run(entryId, tagId, workbookId);
  }
  private requireWorkbook(workbookId: number): WorkbookRow { const workbook = this.getWorkbook(workbookId); if (!workbook) throw new Error(`Workbook with id ${workbookId} was not found.`); return workbook; }
  private requireEntry(entryId: number): EntryRow { const entry = this.getEntry(entryId); if (!entry) throw new Error(`Entry with id ${entryId} was not found.`); return entry; }
  private readCurrentWorkbookId(): number | null { const row = this.db.prepare("SELECT current_workbook_id FROM app_settings WHERE singleton_id = 1").get() as { current_workbook_id: number | null }; return row.current_workbook_id == null ? null : Number(row.current_workbook_id); }
  private firstWorkbookId(): number | null { const row = this.db.prepare("SELECT id FROM workbooks ORDER BY id LIMIT 1").get() as { id: number } | undefined; return row ? Number(row.id) : null; }
}
