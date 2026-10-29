import assert from "node:assert/strict";
import test from "node:test";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { DatabaseSync } from "node:sqlite";
import { VocabularyRepository, WorkbookConfigurationInput, WorkbookRow } from "./db.js";
import { createPresetAttributes, entryAttributes, isPresetActivationInput } from "./attributes.js";
import { LANGUAGE_PRESET_DEFINITIONS } from "./language-presets.js";
import { assertDatabaseIntegrity, runSchemaMigrations, SCHEMA_MIGRATIONS } from "./schema.js";
import { buildDetailSections } from "./detail-display.js";
import { filterEntriesForSearch } from "./search.js";
import { parseLabeledTextImport } from "./import.js";
import { createInitialState, update } from "./orchestration.js";

function configuration(workbook: WorkbookRow): WorkbookConfigurationInput {
  return {
    name: workbook.name, vocabularyKind: workbook.vocabularyKind, vocabularyLabel: workbook.vocabularyLabel,
    vocabularyLanguageCode: workbook.vocabularyLanguageCode, meaningAttributes: workbook.meaningAttributes,
    presetAttributes: workbook.metadataAttributes.filter((field) => Boolean(field.presetKey)),
    optionalAttributes: workbook.metadataAttributes.filter((field) => field.role === "optional" && !field.presetKey), tagTypes: [],
  };
}

test("every supported language automatically receives its complete, ordered, active preset", () => {
  const repository = new VocabularyRepository(":memory:");
  try {
    for (const code of Object.keys(LANGUAGE_PRESET_DEFINITIONS)) {
      const workbook = repository.createWorkbook(code, code, code.toLowerCase());
      const presets = workbook.metadataAttributes.filter((field) => field.presetKey);
      assert.deepEqual(presets.map((field) => field.label), [
        ...(code === "JP" ? ["読み方 (kana)"] : []),
        ...Array.from({ length: 5 }, (_, index) => `Example Sentence ${index + 1}`), "Common Collocations", "Similar Words",
      ]);
      assert.ok(presets.every((field) => field.active && !field.visible && !field.required));
      assert.deepEqual(presets.map((field) => field.presetKey), createPresetAttributes(code).map((field) => field.key));
    }
    const workbook = repository.createWorkbook("No language");
    assert.equal(workbook.metadataAttributes.filter((field) => field.presetKey).length, 0);
    const custom = repository.createConfiguredWorkbook({ ...configuration(workbook), name: "Other", vocabularyKind: "other_language" });
    assert.equal(custom.metadataAttributes.filter((field) => field.presetKey).length, 0);
  } finally { repository.close(); }
});

test("both settings APIs protect preset membership and persist rename, visibility and activation", () => {
  const repository = new VocabularyRepository(":memory:");
  try {
    const workbook = repository.createWorkbook("Japanese", "Japanese", "JP");
    const draft = { vocabularyLabel: workbook.vocabularyLabel, fields: workbook.metadataAttributes.filter((field) => field.role !== "vocabulary") };
    const kana = draft.fields.find((field) => field.presetKey === "kana")!;
    for (const fields of [
      draft.fields.filter((field) => field !== kana),
      draft.fields.map((field) => field === kana ? { ...field, presetKey: null } : field),
      draft.fields.map((field) => field === kana ? { ...field, role: "meaning" as const } : field),
      [...draft.fields, { ...kana, id: undefined, key: "extra_kana" }],
      draft.fields.map((field) => field === kana ? { ...field, key: "changed_key" } : field),
    ]) assert.throws(() => repository.updateWorkbookAttributes(workbook.id, { ...draft, fields }, true), /preset/i);
    const config = configuration(workbook);
    assert.throws(() => repository.updateConfiguredWorkbook(workbook.id, { ...config, presetAttributes: config.presetAttributes.slice(1) }, true), /preset/i);
    assert.throws(() => repository.updateConfiguredWorkbook(workbook.id, { ...config, presetAttributes: config.presetAttributes.map((field) => ({ ...field, key: `changed_${field.key}` })) }, true), /preset/i);
    const updated = repository.updateWorkbookAttributes(workbook.id, { ...draft, fields: draft.fields.map((field) => field === kana ? { ...field, label: "Pronunciation", active: false, visible: true } : field) });
    const renamed = updated.metadataAttributes.find((field) => field.id === kana.id)!;
    assert.deepEqual([renamed.label, renamed.active, renamed.visible, renamed.presetKey], ["Pronunciation", false, true, "kana"]);
    const fullConfig = configuration(updated);
    const reactivated = repository.updateConfiguredWorkbook(workbook.id, { ...fullConfig, presetAttributes: fullConfig.presetAttributes.map((field) => ({ ...field, active: true })) });
    assert.equal(reactivated.metadataAttributes.find((field) => field.id === kana.id)?.visible, true);
  } finally { repository.close(); }
});

test("inactive values survive edits and imports without false changes or learning resets", () => {
  const repository = new VocabularyRepository(":memory:");
  try {
    let workbook = repository.createWorkbook("Japanese", "Japanese", "JP");
    const entry = repository.addEntry(workbook.id, "猫", "cat", ["cat"], { kana: "ねこ", example_sentence_1: "A cat." });
    repository.completePracticeSession(workbook.id, [{ entryId: entry.id, answers: [true] }], "2026-10-28");
    const tested = repository.getEntry(entry.id)!;
    workbook = repository.updateWorkbookAttributes(workbook.id, { vocabularyLabel: workbook.vocabularyLabel, fields: workbook.metadataAttributes.filter((field) => field.role !== "vocabulary").map((field) => field.presetKey ? { ...field, active: false, visible: true } : field) });
    assert.deepEqual(repository.getEntry(entry.id), tested);
    assert.deepEqual(repository.updateEntry(entry.id, "猫", "cat", ["cat"], { kana: "ignored" }), tested);
    assert.equal(repository.importEntries(workbook.id, [{ vocabulary: "猫", meanings: ["cat"], attributes: {}, tagIds: [] }]).unchanged, 1);
    assert.deepEqual(repository.getEntry(entry.id), tested);
    const changed = repository.updateEntry(entry.id, "猫", "feline", ["feline"]);
    assert.equal(changed.attributes.kana, "ねこ");
    assert.equal(changed.attributes.example_sentence_1, "A cat.");
    assert.equal(changed.testCount, 0);
    repository.importEntries(workbook.id, [{ vocabulary: "猫", meanings: ["kitty"], attributes: {}, tagIds: [] }]);
    assert.equal(repository.getEntry(entry.id)?.attributes.kana, "ねこ");
    const added = repository.addEntry(workbook.id, "犬", "dog", ["dog"], { kana: "ignored" });
    assert.equal(added.attributes.kana, "");
    const config = configuration(workbook);
    repository.updateConfiguredWorkbook(workbook.id, { ...config, presetAttributes: config.presetAttributes.map((field) => ({ ...field, active: true })) });
    assert.equal(repository.getEntry(entry.id)?.kanaText, "ねこ");
  } finally { repository.close(); }
});

test("inactive presets leave prompts, detail, visible search, and import while active hidden presets remain usable", () => {
  const repository = new VocabularyRepository(":memory:");
  try {
    let workbook = repository.createWorkbook("Japanese", "Japanese", "JP");
    const entry = repository.addEntry(workbook.id, "猫", "cat", ["cat"], { kana: "ねこ", example_sentence_1: "hidden example" });
    workbook = repository.updateWorkbookAttributes(workbook.id, { vocabularyLabel: workbook.vocabularyLabel, fields: workbook.metadataAttributes.filter((field) => field.role !== "vocabulary").map((field) => ({ ...field, active: field.presetKey === "example_sentence_1", visible: field.presetKey === "kana" || field.visible })) });
    assert.deepEqual(entryAttributes(workbook.metadataAttributes).map((field) => field.key), ["example_sentence_1"]);
    assert.deepEqual(buildDetailSections(workbook, entry, []).attributes, [{ key: "Example Sentence 1", value: "hidden example" }]);
    assert.deepEqual(filterEntriesForSearch([entry], workbook, [], "ねこ"), []);
    assert.deepEqual(filterEntriesForSearch([entry], workbook, [], "hidden example"), []);
    const preview = parseLabeledTextImport("Japanese: 猫\nPrimary Meaning: cat\n読み方 (kana): ignored\nExample Sentence 1: imported", workbook, [], []);
    assert.equal(preview.entries[0].attributes.kana, undefined);
    assert.equal(preview.entries[0].attributes.example_sentence_1, "imported");
    assert.equal(preview.diagnostics.filter((diagnostic) => diagnostic.kind === "ignored-field").length, 1);
    let state = update(createInitialState(workbook), { type: "commandSubmitted", raw: "/add" }).state;
    state = update(state, { type: "addFieldSubmitted", value: "猫" }).state;
    state = update(state, { type: "addFieldSubmitted", value: "cat" }).state;
    assert.equal(state.mode.kind, "add");
    assert.match(state.statusLines.join(" "), /Example Sentence 1/);
    const transition = update(state, { type: "addFieldSubmitted", value: "example" });
    assert.equal(transition.effects[0]?.type, "addEntry");
  } finally { repository.close(); }
});

test("preset activation uses plain T and rejects modifier combinations and legacy inputs", () => {
  assert.equal(isPresetActivationInput("t", { ctrl: false }), true);
  assert.equal(isPresetActivationInput("T", { ctrl: false }), true);
  for (const input of [" ", "\x00", "`", "\x14"]) {
    assert.equal(isPresetActivationInput(input, { ctrl: false }), false);
    assert.equal(isPresetActivationInput(input, { ctrl: true }), false);
  }
  assert.equal(isPresetActivationInput("t", { ctrl: true }), false);
  assert.equal(isPresetActivationInput("t", { ctrl: false, meta: true }), false);
  assert.equal(isPresetActivationInput("T", { ctrl: false, meta: true }), false);
});

test("v6 migration promotes by identity then name, preserves all values and stats, and runs once", () => {
  const directory = mkdtempSync(join(tmpdir(), "vocab-presets-"));
  const path = join(directory, "migration.db");
  const handles: Array<{ close(): void }> = [];
  try {
    const db = new DatabaseSync(path);
    handles.push(db);
    runSchemaMigrations(db, SCHEMA_MIGRATIONS.filter((migration) => migration.version <= 6));
    db.exec(`INSERT INTO workbooks (id,name,vocabulary_kind,vocabulary_label,vocabulary_language_code,preset_enabled,created_at,updated_at) VALUES
      (1,'JP','preset_language','Japanese','JP',1,'old','old'),
      (2,'EN disabled','preset_language','English','EN',0,'old','old'),
      (3,'Other','other_language','Other',NULL,0,'old','old');
      INSERT INTO workbook_fields (id,workbook_id,field_key,role,position,label,language_code,provenance,is_visible) VALUES
      (1,1,'meaning_1','meaning',1,'Meaning','EN','custom',1),
      (2,1,'custom_reading','optional',1,'Kana','JP','custom',1),
      (3,1,'custom_example','optional',2,'  EXAMPLE SENTENCE 2  ',NULL,'custom',1),
      (4,1,'example_sentence_1','optional',3,'My example','JP','preset',0),
      (5,1,'duplicate','optional',4,'Example Sentence 2',NULL,'custom',0),
      (6,1,'similar_words','optional',5,'Unrelated note',NULL,'custom',1),
      (7,2,'meaning_1','meaning',1,'Meaning','EN','custom',1),
      (8,2,'collocations','optional',1,'Common Collocations',NULL,'custom',1),
      (9,3,'meaning_1','meaning',1,'Meaning',NULL,'custom',1),
      (10,3,'examples','optional',1,'Example Sentence 1',NULL,'custom',1),
      (11,1,'example_sentence_5','optional',6,'Example Sentence 3','JP','preset',0);
      INSERT INTO entries (id,workbook_id,vocabulary,created_at,updated_at) VALUES (1,1,'猫','old','old'),(2,2,'cat','old','old');
      INSERT INTO entry_field_values (entry_id,field_id,workbook_id,value) SELECT 1,id,1,'value-' || id FROM workbook_fields WHERE workbook_id=1;
      INSERT INTO entry_field_values (entry_id,field_id,workbook_id,value) SELECT 2,id,2,'value-' || id FROM workbook_fields WHERE workbook_id=2;
      INSERT INTO entry_stats (entry_id,test_count,error_count,last_tested,next_test_deadline) VALUES (1,5,1,'2026-10-28','2026-11-04'),(2,2,0,'2026-10-28','2026-11-12');`);
    const before = db.prepare("SELECT * FROM entry_field_values ORDER BY field_id").all();
    const stats = db.prepare("SELECT * FROM entry_stats ORDER BY entry_id").all();
    db.close();
    const repository = new VocabularyRepository(path);
    handles.push(repository);
    const jp = repository.getWorkbook(1)!;
    assert.equal(jp.metadataAttributes.filter((field) => field.presetKey).length, 8);
    assert.deepEqual(jp.metadataAttributes.filter((field) => field.presetKey).map((field) => field.presetKey), createPresetAttributes("JP").map((field) => field.presetKey));
    assert.equal(jp.metadataAttributes.find((field) => field.id === 2)?.label, "読み方 (kana)");
    assert.equal(jp.metadataAttributes.find((field) => field.id === 2)?.key, "custom_reading");
    assert.equal(jp.metadataAttributes.find((field) => field.id === 3)?.presetKey, "example_sentence_2");
    assert.equal(jp.metadataAttributes.find((field) => field.id === 4)?.label, "My example");
    assert.equal(jp.metadataAttributes.find((field) => field.id === 5)?.presetKey, null);
    assert.equal(jp.metadataAttributes.find((field) => field.id === 6)?.presetKey, null);
    assert.equal(jp.metadataAttributes.find((field) => field.id === 11)?.presetKey, "example_sentence_5");
    assert.notEqual(jp.metadataAttributes.find((field) => field.presetKey === "example_sentence_3")?.id, 11);
    assert.equal(jp.metadataAttributes.find((field) => field.presetKey === "similar_words")?.key, "similar_words_2");
    assert.equal(repository.getWorkbook(2)?.metadataAttributes.find((field) => field.id === 8)?.presetKey, "common_collocations");
    assert.equal(repository.getWorkbook(3)?.metadataAttributes.some((field) => field.presetKey), false);
    assert.ok(jp.metadataAttributes.filter((field) => field.presetKey).every((field) => field.active));
    assert.equal(repository.getEntry(1)?.attributes.custom_example, "value-3");
    assert.equal(repository.getEntry(1)?.kanaText, "value-2");
    // Duplicate labels left in different groups must not prevent activation changes.
    repository.updateWorkbookAttributes(1, { vocabularyLabel: jp.vocabularyLabel, fields: jp.metadataAttributes.filter((field) => field.role !== "vocabulary").map((field) => field.presetKey ? { ...field, active: false } : field) });
    const snapshot = repository.listWorkbooks();
    repository.close();
    const reopened = new VocabularyRepository(path);
    handles.push(reopened);
    assert.deepEqual(reopened.listWorkbooks(), snapshot);
    reopened.close();
    const check = new DatabaseSync(path);
    handles.push(check);
    const originalIds = new Set(before.map((row) => Number(row.field_id)));
    assert.deepEqual(check.prepare("SELECT * FROM entry_field_values ORDER BY field_id").all().filter((row) => originalIds.has(Number(row.field_id))), before);
    assert.deepEqual(check.prepare("SELECT * FROM entry_stats ORDER BY entry_id").all(), stats);
    assertDatabaseIntegrity(check);
    check.close();
  } finally {
    for (const handle of handles) { try { handle.close(); } catch { /* Already closed. */ } }
    rmSync(directory, { recursive: true, force: true });
  }
});

test("a failed preset migration rolls back field changes and its ledger entry", () => {
  const db = new DatabaseSync(":memory:");
  try {
    runSchemaMigrations(db, SCHEMA_MIGRATIONS.filter((migration) => migration.version <= 6));
    db.exec(`INSERT INTO workbooks (name,vocabulary_kind,vocabulary_label,vocabulary_language_code,created_at,updated_at) VALUES ('English','preset_language','English','EN','old','old');
      CREATE TRIGGER reject_preset BEFORE INSERT ON workbook_fields WHEN NEW.field_key = 'similar_words' BEGIN SELECT RAISE(ABORT, 'test migration failure'); END;`);
    assert.throws(() => runSchemaMigrations(db), /migration 7 failed.*test migration failure/i);
    assert.equal(db.prepare("SELECT 1 FROM pragma_table_info('workbook_fields') WHERE name = 'preset_key'").get(), undefined);
    assert.equal(db.prepare("SELECT 1 FROM schema_migrations WHERE version = 7").get(), undefined);
    assert.equal(db.prepare("SELECT 1 FROM workbook_fields").get(), undefined);
  } finally { db.close(); }
});
