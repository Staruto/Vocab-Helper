import { createHash } from "node:crypto";
import { existsSync, renameSync, unlinkSync, writeFileSync } from "node:fs";
import { basename, dirname, join, resolve } from "node:path";
import { DatabaseSync } from "node:sqlite";
import { defaultDbPath, LANGUAGE_PRESET_DEFINITIONS } from "./db.js";
import { assertDatabaseIntegrity, isHybridOrLegacyDatabase, runSchemaMigrations } from "./schema.js";

type Row = Record<string, unknown>;
export type ConversionReport = {
  source: string;
  applied: boolean;
  backup: string | null;
  counts: Record<string, number>;
  discarded: Record<string, number>;
  checksums: Record<string, string>;
  integrity: "ok";
};

function tableExists(db: DatabaseSync, name: string): boolean {
  return Boolean(db.prepare("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?").get(name));
}
function rows(db: DatabaseSync, sql: string, ...params: unknown[]): Row[] { return db.prepare(sql).all(...params) as Row[]; }
function scalar(db: DatabaseSync, sql: string, ...params: unknown[]): number {
  const row = db.prepare(sql).get(...params) as Row | undefined;
  return Number(row ? Object.values(row)[0] : 0);
}
function checksum(value: unknown): string { return createHash("sha256").update(JSON.stringify(value)).digest("hex"); }
function stamp(): string { return new Date().toISOString().replace(/[-:]/g, "").replace(/\.\d{3}Z$/, "Z").replace("T", "-"); }

export function convertHybridDatabase(sourcePath: string, apply = false): ConversionReport {
  const source = resolve(sourcePath);
  if (!existsSync(source)) throw new Error(`Database not found: ${source}`);
  const sourceDb = new DatabaseSync(source, { readOnly: true });
  if (!isHybridOrLegacyDatabase(sourceDb) || !tableExists(sourceDb, "mvp_workbooks")) {
    sourceDb.close();
    throw new Error("The source is not a TypeScript MVP hybrid database, so no conversion was performed.");
  }

  const temp = join(dirname(source), `${basename(source)}.v1.tmp`);
  if (existsSync(temp)) { sourceDb.close(); throw new Error(`Temporary conversion file already exists: ${temp}`); }
  const targetDb = new DatabaseSync(temp);
  let backup: string | null = null;
  try {
    runSchemaMigrations(targetDb);
    const workbooks = rows(sourceDb, "SELECT * FROM mvp_workbooks ORDER BY id");
    const definitions = rows(sourceDb, "SELECT * FROM mvp_workbook_meaning_attributes ORDER BY workbook_id, position");
    const metadata = rows(sourceDb, "SELECT * FROM mvp_workbook_attributes ORDER BY workbook_id, display_order, attribute_key");
    const sourceEntries = rows(sourceDb, "SELECT * FROM mvp_entries WHERE workbook_id IS NOT NULL ORDER BY id");
    const sourceMeanings = rows(sourceDb, "SELECT * FROM mvp_entry_meanings ORDER BY entry_id, position");
    const sourceValues = rows(sourceDb, "SELECT * FROM mvp_entry_attributes ORDER BY entry_id, attribute_key");
    const sourceTags = rows(sourceDb, "SELECT * FROM mvp_pos_tags ORDER BY id");
    const sourceAssignments = rows(sourceDb, "SELECT ep.entry_id, ep.tag_id, e.workbook_id AS entry_workbook_id, t.workbook_id AS tag_workbook_id FROM mvp_entry_pos_tags ep JOIN mvp_entries e ON e.id=ep.entry_id JOIN mvp_pos_tags t ON t.id=ep.tag_id ORDER BY ep.entry_id, ep.tag_id");
    const sourceStats = tableExists(sourceDb, "mvp_entry_stats") ? rows(sourceDb, "SELECT * FROM mvp_entry_stats ORDER BY entry_id") : [];
    const legacyTagCount = tableExists(sourceDb, "entry_tags") ? scalar(sourceDb, "SELECT COUNT(*) FROM entry_tags") : 0;
    const entryById = new Map(sourceEntries.map((row) => [Number(row.id), row]));
    const valueByEntryKey = new Map(sourceValues.map((row) => [`${row.entry_id}:${row.attribute_key}`, String(row.value ?? "")]));
    const meaningByEntryPosition = new Map(sourceMeanings.map((row) => [`${row.entry_id}:${row.position}`, String(row.value ?? "")]));
    const statsByEntry = new Map(sourceStats.map((row) => [Number(row.entry_id), row]));
    const fieldMap = new Map<string, number>();
    const sourceKeyMap = new Map<string, string>();
    const expectedValues: unknown[] = [];

    targetDb.exec("BEGIN IMMEDIATE");
    try {
      const addWorkbook = targetDb.prepare("INSERT INTO workbooks (id,name,vocabulary_kind,vocabulary_label,vocabulary_language_code,preset_enabled,created_at,updated_at) VALUES (?,?,?,?,?,?,?,?)");
      const addField = targetDb.prepare("INSERT INTO workbook_fields (workbook_id,field_key,role,position,label,language_code,is_required,is_visible,provenance) VALUES (?,?,?,?,?,?,?,?,?)");
      for (const workbook of workbooks) {
        const id = Number(workbook.id); const code = workbook.vocabulary_language_code == null ? null : String(workbook.vocabulary_language_code);
        const kind = ["preset_language", "other_language", "non_language"].includes(String(workbook.vocabulary_kind)) ? String(workbook.vocabulary_kind) : code ? "preset_language" : "non_language";
        const created = String(workbook.created_at ?? new Date().toISOString());
        addWorkbook.run(id, String(workbook.name), kind, String(workbook.vocabulary_label ?? "Vocabulary"), kind === "preset_language" ? code : null, Number(workbook.preset_enabled) ? 1 : 0, created, created);
        const workbookMeanings = definitions.filter((row) => Number(row.workbook_id) === id);
        const meanings = workbookMeanings.length ? workbookMeanings : [{ position: 1, label: "Primary Meaning", language_code: null }];
        for (const definition of meanings) {
          const position = Number(definition.position); const result = addField.run(id, `meaning_${position}`, "meaning", position, String(definition.label), definition.language_code ?? null, position === 1 ? 1 : 0, position === 1 ? 1 : 0, "custom");
          fieldMap.set(`${id}:meaning_${position}`, Number(result.lastInsertRowid));
        }
        const candidates = metadata.filter((row) => Number(row.workbook_id) === id && String(row.attribute_key) !== "vocab" && !String(row.attribute_key).startsWith("meaning_"));
        const hasKanaData = sourceEntries.some((entry) => Number(entry.workbook_id) === id && String(entry.kana_text ?? "").trim() !== "");
        if (hasKanaData && !candidates.some((row) => String(row.attribute_key) === "kana")) candidates.unshift({ workbook_id: id, attribute_key: "kana", label: "Kana", language_code: "JP", is_visible: 0 });
        const convertedKeys = new Set<string>();
        for (let index = 0; index < candidates.length; index += 1) {
          const field = candidates[index]; const oldKey = String(field.attribute_key); const oldLabel = String(field.label);
          const isDefaultJapaneseExample = code === "JP" && ((oldKey === "example_1" && oldLabel === "Example 1") || (oldKey === "example_2" && oldLabel === "Example 2"));
          const key = isDefaultJapaneseExample ? `example_sentence_${oldKey.endsWith("1") ? "1" : "2"}` : oldKey;
          const label = isDefaultJapaneseExample ? `Example Sentence ${oldKey.endsWith("1") ? "1" : "2"}` : oldLabel;
          if (convertedKeys.has(key)) throw new Error(`Workbook ${id} has conflicting optional field key '${key}'.`);
          convertedKeys.add(key); sourceKeyMap.set(`${id}:${key}`, oldKey);
          const preset = code ? LANGUAGE_PRESET_DEFINITIONS[code] : undefined;
          const provenance = preset?.optionalAttributes.some((item) => item.key === key && item.label === label) ? "preset" : "custom";
          const result = addField.run(id, key, "optional", index + 1, label, field.language_code ?? null, 0, Number(field.is_visible) ? 1 : 0, provenance);
          fieldMap.set(`${id}:${key}`, Number(result.lastInsertRowid));
        }
      }

      const addEntry = targetDb.prepare("INSERT INTO entries (id,workbook_id,vocabulary,created_at,updated_at) VALUES (?,?,?,?,?)");
      const addValue = targetDb.prepare("INSERT INTO entry_field_values (entry_id,field_id,workbook_id,value) VALUES (?,?,?,?)");
      const addStat = targetDb.prepare("INSERT INTO entry_stats (entry_id,test_count,error_count,last_tested,next_test_deadline) VALUES (?,?,?,?,?)");
      for (const entry of sourceEntries) {
        const id = Number(entry.id); const workbookId = Number(entry.workbook_id); const created = String(entry.created_at ?? new Date().toISOString());
        addEntry.run(id, workbookId, String(entry.vocabulary), created, String(entry.updated_at ?? created));
        const targetFields = rows(targetDb, "SELECT id,field_key,role,position FROM workbook_fields WHERE workbook_id=? ORDER BY role,position", workbookId);
        for (const field of targetFields) {
          const key = String(field.field_key); const oldKey = sourceKeyMap.get(`${workbookId}:${key}`) ?? key;
          let value = field.role === "meaning" ? meaningByEntryPosition.get(`${id}:${field.position}`) : valueByEntryKey.get(`${id}:${oldKey}`);
          if (value === undefined && field.role === "meaning" && Number(field.position) === 1) value = String(entry.meaning ?? "");
          if (value === undefined && key === "kana") value = String(entry.kana_text ?? "");
          value ??= ""; addValue.run(id, Number(field.id), workbookId, value); expectedValues.push([id, key, value]);
        }
        const stat = statsByEntry.get(id);
        addStat.run(id, Math.max(0, Number(stat?.test_count ?? 0)), Math.min(3, Math.max(0, Number(stat?.error_count ?? 0))), stat?.last_tested ?? null, stat?.next_test_deadline ?? null);
      }

      const addTagType = targetDb.prepare("INSERT INTO tag_types (workbook_id,name,position,is_visible) VALUES (?,'Part of Speech',1,0)");
      const addTag = targetDb.prepare("INSERT INTO tags (id,tag_type_id,workbook_id,name) VALUES (?,?,?,?)");
      for (const workbook of workbooks) {
        const workbookId = Number(workbook.id);
        const workbookTags = sourceTags.filter((tag) => Number(tag.workbook_id) === workbookId);
        if (!Number(workbook.pos_enabled) && workbookTags.length === 0) continue;
        const tagTypeId = Number(addTagType.run(workbookId).lastInsertRowid);
        for (const tag of workbookTags) addTag.run(Number(tag.id), tagTypeId, workbookId, String(tag.name));
      }
      const validAssignments = sourceAssignments.filter((row) => Number(row.entry_workbook_id) === Number(row.tag_workbook_id) && entryById.has(Number(row.entry_id)));
      const addAssignment = targetDb.prepare("INSERT INTO entry_tags (entry_id,tag_id,workbook_id) VALUES (?,?,?)");
      for (const assignment of validAssignments) addAssignment.run(Number(assignment.entry_id), Number(assignment.tag_id), Number(assignment.entry_workbook_id));
      const settings = tableExists(sourceDb, "mvp_meta") ? new Map(rows(sourceDb, "SELECT key,value FROM mvp_meta").map((row) => [String(row.key), String(row.value)])) : new Map<string, string>();
      const currentId = Number(settings.get("current_workbook_id"));
      const current = Number.isFinite(currentId) && workbooks.some((row) => Number(row.id) === currentId) ? currentId : (workbooks[0] ? Number(workbooks[0].id) : null);
      targetDb.prepare("UPDATE app_settings SET current_workbook_id=?, tier_colors_enabled=? WHERE singleton_id=1").run(current, settings.get("tier_colors_enabled") === "0" ? 0 : 1);
      targetDb.exec("COMMIT");

      assertDatabaseIntegrity(targetDb);
      const counts = {
        workbooks: scalar(targetDb, "SELECT COUNT(*) FROM workbooks"), entries: scalar(targetDb, "SELECT COUNT(*) FROM entries"),
        fields: scalar(targetDb, "SELECT COUNT(*) FROM workbook_fields"), values: scalar(targetDb, "SELECT COUNT(*) FROM entry_field_values"),
        tags: scalar(targetDb, "SELECT COUNT(*) FROM tags"), tagAssignments: scalar(targetDb, "SELECT COUNT(*) FROM entry_tags"), stats: scalar(targetDb, "SELECT COUNT(*) FROM entry_stats"),
      };
      if (counts.workbooks !== workbooks.length || counts.entries !== sourceEntries.length || counts.tags !== sourceTags.length || counts.stats !== sourceEntries.length) throw new Error("Converted row counts do not match the MVP source.");
      const targetValues = rows(targetDb, "SELECT v.entry_id,f.field_key,v.value FROM entry_field_values v JOIN workbook_fields f ON f.id=v.field_id ORDER BY v.entry_id,f.field_key").map((row) => [Number(row.entry_id), String(row.field_key), String(row.value)]).sort();
      const checksums = { valuesExpected: checksum([...expectedValues].sort()), valuesActual: checksum(targetValues), entries: checksum(rows(targetDb, "SELECT id,workbook_id,vocabulary,created_at,updated_at FROM entries ORDER BY id")), stats: checksum(rows(targetDb, "SELECT * FROM entry_stats ORDER BY entry_id")) };
      if (checksums.valuesExpected !== checksums.valuesActual) throw new Error("Converted field-value checksum does not match the source projection.");
      const report: ConversionReport = { source, applied: apply, backup: null, counts, discarded: { legacyTags: legacyTagCount, invalidMvpPosAssignments: sourceAssignments.length - validAssignments.length }, checksums, integrity: "ok" };
      targetDb.close(); sourceDb.close();
      if (!apply) { unlinkSync(temp); return report; }
      backup = join(dirname(source), `${basename(source)}.backup-before-v1-${stamp()}`);
      if (existsSync(backup)) throw new Error(`Backup path already exists: ${backup}`);
      renameSync(source, backup);
      try { renameSync(temp, source); }
      catch (error) { renameSync(backup, source); throw error; }
      report.backup = backup;
      writeFileSync(`${backup}.conversion-report.json`, `${JSON.stringify(report, null, 2)}\n`, "utf8");
      return report;
    } catch (error) {
      try { targetDb.exec("ROLLBACK"); } catch { /* The transaction may already be closed. */ }
      throw error;
    }
  } finally {
    try { targetDb.close(); } catch { /* Already closed. */ }
    try { sourceDb.close(); } catch { /* Already closed. */ }
    if (existsSync(temp) && !backup) try { unlinkSync(temp); } catch { /* Keep the original failure. */ }
  }
}

function main(): void {
  const args = process.argv.slice(2); const apply = args.includes("--apply"); const sourceIndex = args.indexOf("--source");
  const source = sourceIndex >= 0 && args[sourceIndex + 1] ? args[sourceIndex + 1] : defaultDbPath();
  const report = convertHybridDatabase(source, apply);
  console.log(JSON.stringify(report, null, 2));
  if (!apply) console.log("Dry run complete. Re-run with --apply to back up and replace the source database.");
}

if (process.argv[1]?.endsWith("convert-db.ts") || process.argv[1]?.endsWith("convert-db.js")) {
  try { main(); } catch (error) { console.error(error instanceof Error ? error.message : String(error)); process.exit(1); }
}
