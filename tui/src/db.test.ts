import assert from "node:assert/strict";
import { existsSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { DatabaseSync } from "node:sqlite";
import test from "node:test";
import { convertHybridDatabase } from "./convert-db.js";
import { LANGUAGE_PRESET_DEFINITIONS, ValidationError, VocabularyRepository, WorkbookDataLossError, WorkbookConfigurationInput } from "./db.js";
import { assertDatabaseIntegrity, runSchemaMigrations } from "./schema.js";

function temporaryDatabase(): { directory: string; path: string; cleanup: () => void } {
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-"));
  return { directory, path: join(directory, "vocab.db"), cleanup: () => rmSync(directory, { recursive: true, force: true }) };
}

function basicWorkbook(name = "Test"): WorkbookConfigurationInput {
  return {
    name, vocabularyKind: "preset_language", vocabularyLabel: "Japanese", vocabularyLanguageCode: "JP",
    presetEnabled: true, posEnabled: true,
    meaningAttributes: [{ position: 1, label: "English", languageCode: "EN" }],
    optionalAttributes: LANGUAGE_PRESET_DEFINITIONS.JP.optionalAttributes.map((field, index) => ({ ...field, required: false, visible: false, displayOrder: index + 1, provenance: "preset" })),
    posTags: [{ name: "名詞", predefined: true }],
  };
}

test("fresh databases use the v1 schema and migrations are idempotent", () => {
  const temp = temporaryDatabase();
  try {
    const repository = new VocabularyRepository(temp.path); repository.close();
    const db = new DatabaseSync(temp.path);
    assert.equal((db.prepare("SELECT COUNT(*) AS count FROM schema_migrations").get() as { count: number }).count, 1);
    assert.equal((db.prepare("SELECT COUNT(*) AS count FROM sqlite_master WHERE type='table' AND name LIKE 'mvp_%'").get() as { count: number }).count, 0);
    assertDatabaseIntegrity(db); db.close();
    const reopened = new VocabularyRepository(temp.path); reopened.close();
  } finally { temp.cleanup(); }
});

test("a failed migration rolls back its schema and ledger row", () => {
  const temp = temporaryDatabase();
  try {
    const db = new DatabaseSync(temp.path); runSchemaMigrations(db);
    assert.throws(() => runSchemaMigrations(db, [{ version: 2, apply(inner) { inner.exec("CREATE TABLE rollback_probe (id INTEGER); INSERT INTO missing_table VALUES (1)"); } }]), /migration 2 failed/i);
    assert.equal((db.prepare("SELECT COUNT(*) AS count FROM sqlite_master WHERE name='rollback_probe'").get() as { count: number }).count, 0);
    assert.equal((db.prepare("SELECT COUNT(*) AS count FROM schema_migrations WHERE version=2").get() as { count: number }).count, 0);
    db.close();
  } finally { temp.cleanup(); }
});

test("entry stats, priority, ownership, and cascades are enforced", () => {
  const temp = temporaryDatabase();
  try {
    const repository = new VocabularyRepository(temp.path);
    const first = repository.createConfiguredWorkbook(basicWorkbook("First"));
    const second = repository.createConfiguredWorkbook(basicWorkbook("Second"));
    const firstTag = repository.listPosTags(first.id)[0]; const secondTag = repository.listPosTags(second.id)[0];
    assert.throws(() => repository.addEntry(first.id, "猫", "cat", ["cat"], {}, [secondTag.id]), /does not belong/);
    const entry = repository.addEntry(first.id, "猫", "cat", ["cat"], { kana: "ねこ" }, [firstTag.id]);
    assert.equal(entry.tier, "gray"); assert.equal(entry.testCount, 0); assert.equal(entry.errorCount, 0);
    assert.equal(repository.recordTestResult(entry.id, false).errorCount, 1);
    assert.equal(repository.recordTestResult(entry.id, true).errorCount, 0);
    repository.deleteEntry(entry.id); repository.close();
    const db = new DatabaseSync(temp.path);
    assert.equal((db.prepare("SELECT COUNT(*) AS count FROM entry_stats").get() as { count: number }).count, 0);
    assert.equal((db.prepare("SELECT COUNT(*) AS count FROM entry_field_values").get() as { count: number }).count, 0);
    assertDatabaseIntegrity(db); db.close();
  } finally { temp.cleanup(); }
});

test("workbook updates require confirmation before deleting populated fields", () => {
  const temp = temporaryDatabase();
  try {
    const repository = new VocabularyRepository(temp.path); const input = basicWorkbook(); const workbook = repository.createConfiguredWorkbook(input);
    repository.addEntry(workbook.id, "猫", "cat", ["cat"], { kana: "ねこ", example_sentence_1: "A cat." });
    const changed = { ...input, optionalAttributes: input.optionalAttributes.filter((field) => field.key !== "example_sentence_1") };
    const impact = repository.previewWorkbookUpdate(workbook.id, changed);
    assert.deepEqual(impact.populatedFields.map((field) => field.label), ["Example Sentence 1"]);
    assert.throws(() => repository.updateConfiguredWorkbook(workbook.id, changed), (error) => error instanceof WorkbookDataLossError);
    repository.updateConfiguredWorkbook(workbook.id, changed, true);
    assert.equal(repository.getEntry(1)?.attributes.example_sentence_1, undefined);
    repository.close();
  } finally { temp.cleanup(); }
});

test("attribute updates preserve stable meaning identity and values when a middle meaning is removed", () => {
  const temp = temporaryDatabase(); let repository: VocabularyRepository | undefined;
  try {
    repository = new VocabularyRepository(temp.path);
    const input = basicWorkbook();
    input.meaningAttributes = [
      { position: 1, label: "English", languageCode: "EN" },
      { position: 2, label: "Definition", languageCode: "EN" },
      { position: 3, label: "Alias", languageCode: "EN" },
    ];
    const workbook = repository.createConfiguredWorkbook(input);
    const entry = repository.addEntry(workbook.id, "cat", "cat", ["cat", "small feline", "kitty"]);
    const before = repository.getWorkbook(workbook.id)!;
    const alias = before.metadataAttributes.find((field) => field.role === "meaning" && field.label === "Alias")!;
    const fields = before.metadataAttributes.filter((field) => field.key !== "vocab" && field.label !== "Definition");

    assert.deepEqual(repository.previewWorkbookAttributesUpdate(workbook.id, { vocabularyLabel: before.vocabularyLabel, fields }).populatedFields.map((field) => field.label), ["Definition"]);
    const updated = repository.updateWorkbookAttributes(workbook.id, { vocabularyLabel: before.vocabularyLabel, fields }, true);
    const movedAlias = updated.metadataAttributes.find((field) => field.label === "Alias")!;
    assert.equal(movedAlias.id, alias.id);
    assert.equal(movedAlias.displayOrder, 2);
    assert.deepEqual(repository.getEntry(entry.id)?.meanings, ["cat", "kitty"]);
  } finally { repository?.close(); temp.cleanup(); }
});

test("new attributes receive blank values and Meaning 1 remains required and visible", () => {
  const temp = temporaryDatabase(); let repository: VocabularyRepository | undefined;
  try {
    repository = new VocabularyRepository(temp.path);
    const workbook = repository.createConfiguredWorkbook(basicWorkbook());
    const entry = repository.addEntry(workbook.id, "cat", "cat");
    const fields = workbook.metadataAttributes.filter((field) => field.key !== "vocab").map((field) => field.role === "meaning" ? { ...field, required: false, visible: false } : field);
    fields.push({ key: "meaning_note", role: "meaning", label: "Meaning Note", languageCode: null, required: false, visible: false, displayOrder: 2 });
    fields.push({ key: "source", role: "optional", label: "Source", languageCode: null, required: false, visible: true, displayOrder: 99 });

    const updated = repository.updateWorkbookAttributes(workbook.id, { vocabularyLabel: "Term", fields });
    assert.equal(updated.vocabularyLabel, "Term");
    assert.equal(updated.meaningAttributes[0].position, 1);
    const firstMeaning = updated.metadataAttributes.find((field) => field.role === "meaning" && field.displayOrder === 1)!;
    assert.equal(firstMeaning.required, true);
    assert.equal(firstMeaning.visible, true);
    assert.deepEqual(repository.getEntry(entry.id)?.meanings, ["cat", ""]);
    assert.equal(repository.getEntry(entry.id)?.attributes.source, "");
  } finally { repository?.close(); temp.cleanup(); }
});

test("attribute drafts reject invalid identities, section changes, and duplicate labels", () => {
  const temp = temporaryDatabase(); let repository: VocabularyRepository | undefined;
  try {
    repository = new VocabularyRepository(temp.path);
    const workbook = repository.createConfiguredWorkbook(basicWorkbook());
    const fields = workbook.metadataAttributes.filter((field) => field.key !== "vocab");
    const meaning = fields.find((field) => field.role === "meaning")!;
    const optional = fields.find((field) => field.role === "optional")!;
    assert.throws(() => repository!.updateWorkbookAttributes(workbook.id, { vocabularyLabel: workbook.vocabularyLabel, fields: [{ ...meaning, role: "optional" }, ...fields.filter((field) => field !== meaning)] }), ValidationError);
    assert.throws(() => repository!.updateWorkbookAttributes(workbook.id, { vocabularyLabel: workbook.vocabularyLabel, fields: fields.map((field) => field === optional ? { ...field, id: undefined } : field) }), /already belongs/i);
    assert.throws(() => repository!.updateWorkbookAttributes(workbook.id, { vocabularyLabel: workbook.vocabularyLabel, fields: fields.map((field) => field === optional ? { ...field, label: fields.find((candidate) => candidate.role === "optional" && candidate !== optional)!.label } : field) }), /must be unique/i);

    const changedType = { ...basicWorkbook(), vocabularyKind: "non_language" as const, vocabularyLanguageCode: null };
    assert.throws(() => repository!.updateConfiguredWorkbook(workbook.id, changedType), /type cannot be changed/i);
    assert.throws(() => repository!.updateConfiguredWorkbook(workbook.id, { ...basicWorkbook(), presetEnabled: false }), /preset selection cannot be changed/i);
  } finally { repository?.close(); temp.cleanup(); }
});

test("only populated attribute removal requires destructive confirmation", () => {
  const temp = temporaryDatabase(); let repository: VocabularyRepository | undefined;
  try {
    repository = new VocabularyRepository(temp.path);
    const workbook = repository.createConfiguredWorkbook(basicWorkbook());
    repository.addEntry(workbook.id, "cat", "cat", ["cat"], { kana: "", example_sentence_1: "A cat." });
    const baseFields = workbook.metadataAttributes.filter((field) => field.key !== "vocab");
    const withoutEmptyKana = baseFields.filter((field) => field.key !== "kana");
    assert.deepEqual(repository.previewWorkbookAttributesUpdate(workbook.id, { vocabularyLabel: workbook.vocabularyLabel, fields: withoutEmptyKana }).populatedFields, []);
    repository.updateWorkbookAttributes(workbook.id, { vocabularyLabel: workbook.vocabularyLabel, fields: withoutEmptyKana });

    const refreshed = repository.getWorkbook(workbook.id)!;
    const withoutExample = refreshed.metadataAttributes.filter((field) => field.key !== "vocab" && field.key !== "example_sentence_1");
    const impact = repository.previewWorkbookAttributesUpdate(workbook.id, { vocabularyLabel: refreshed.vocabularyLabel, fields: withoutExample });
    assert.deepEqual(impact.populatedFields.map((field) => [field.label, field.valueCount]), [["Example Sentence 1", 1]]);
    assert.throws(() => repository!.updateWorkbookAttributes(workbook.id, { vocabularyLabel: refreshed.vocabularyLabel, fields: withoutExample }), WorkbookDataLossError);
  } finally { repository?.close(); temp.cleanup(); }
});

test("never-tested entries rank before tested entries", () => {
  const temp = temporaryDatabase();
  try {
    const repository = new VocabularyRepository(temp.path); const workbook = repository.createConfiguredWorkbook(basicWorkbook());
    const tested = repository.addEntry(workbook.id, "犬", "dog"); repository.recordTestResult(tested.id, true);
    const fresh = repository.addEntry(workbook.id, "猫", "cat");
    assert.equal(repository.selectPracticeCandidates(workbook.id, 1)[0].id, fresh.id);
    repository.close();
  } finally { temp.cleanup(); }
});

test("the converter preserves MVP data and upgrades default Japanese example labels", () => {
  const temp = temporaryDatabase();
  try {
    const db = new DatabaseSync(temp.path);
    db.exec(`
      CREATE TABLE mvp_workbooks(id INTEGER PRIMARY KEY,name TEXT,created_at TEXT,vocabulary_label TEXT,vocabulary_language_code TEXT,preset_enabled INTEGER,vocabulary_kind TEXT,pos_enabled INTEGER);
      CREATE TABLE mvp_workbook_meaning_attributes(workbook_id INTEGER,position INTEGER,label TEXT,language_code TEXT);
      CREATE TABLE mvp_workbook_attributes(workbook_id INTEGER,attribute_key TEXT,label TEXT,language_code TEXT,is_required INTEGER,is_visible INTEGER,display_order INTEGER);
      CREATE TABLE mvp_entries(id INTEGER PRIMARY KEY,vocabulary TEXT,meaning TEXT,kana_text TEXT,created_at TEXT,updated_at TEXT,workbook_id INTEGER);
      CREATE TABLE mvp_entry_meanings(entry_id INTEGER,position INTEGER,value TEXT);
      CREATE TABLE mvp_entry_attributes(entry_id INTEGER,attribute_key TEXT,value TEXT);
      CREATE TABLE mvp_pos_tags(id INTEGER PRIMARY KEY,workbook_id INTEGER,name TEXT,is_predefined INTEGER);
      CREATE TABLE mvp_entry_pos_tags(entry_id INTEGER,tag_id INTEGER);
      CREATE TABLE mvp_entry_stats(entry_id INTEGER,test_count INTEGER,error_count INTEGER,last_tested TEXT,next_test_deadline TEXT);
      CREATE TABLE mvp_meta(key TEXT,value TEXT);
      INSERT INTO mvp_workbooks VALUES(7,'Japanese','2026-01-01','Japanese','JP',1,'preset_language',1);
      INSERT INTO mvp_workbook_meaning_attributes VALUES(7,1,'English','EN');
      INSERT INTO mvp_workbook_attributes VALUES(7,'vocab','Japanese','JP',1,1,0),(7,'meaning_1','Wrong metadata label','EN',1,1,1),(7,'kana','Kana','JP',0,0,2),(7,'example_1','Example 1','JP',0,0,3);
      INSERT INTO mvp_entries VALUES(42,'猫','cat','ねこ','2026-01-01','2026-01-02',7);
      INSERT INTO mvp_entry_meanings VALUES(42,1,'cat'); INSERT INTO mvp_entry_attributes VALUES(42,'example_1','A cat.');
      INSERT INTO mvp_pos_tags VALUES(9,7,'名詞',1); INSERT INTO mvp_entry_pos_tags VALUES(42,9);
      INSERT INTO mvp_entry_stats VALUES(42,5,2,'2026-01-03','2026-01-07'); INSERT INTO mvp_meta VALUES('current_workbook_id','7');
    `); db.close();
    const report = convertHybridDatabase(temp.path, true); assert.equal(report.integrity, "ok"); assert.ok(report.backup && existsSync(report.backup));
    const converted = new VocabularyRepository(temp.path); const workbook = converted.getWorkbook(7)!; const entry = converted.getEntry(42)!;
    assert.equal(workbook.meaningAttributes[0].label, "English");
    assert.ok(workbook.metadataAttributes.some((field) => field.key === "example_sentence_1" && field.label === "Example Sentence 1"));
    assert.equal(entry.attributes.example_sentence_1, "A cat."); assert.equal(entry.attributes.kana, "ねこ"); assert.equal(entry.testCount, 5); assert.equal(entry.posTags[0].id, 9);
    converted.close();
  } finally { temp.cleanup(); }
});
