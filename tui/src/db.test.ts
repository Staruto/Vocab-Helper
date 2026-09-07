import assert from "node:assert/strict";
import { existsSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { DatabaseSync } from "node:sqlite";
import test from "node:test";
import { convertHybridDatabase } from "./convert-db.js";
import { LANGUAGE_PRESET_DEFINITIONS, TagDataLossError, ValidationError, VocabularyRepository, WorkbookDataLossError, WorkbookConfigurationInput } from "./db.js";
import { assertDatabaseIntegrity, runSchemaMigrations, SCHEMA_MIGRATIONS } from "./schema.js";

function temporaryDatabase(): { directory: string; path: string; cleanup: () => void } {
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-"));
  return { directory, path: join(directory, "vocab.db"), cleanup: () => rmSync(directory, { recursive: true, force: true }) };
}

function basicWorkbook(name = "Test"): WorkbookConfigurationInput {
  return {
    name, vocabularyKind: "preset_language", vocabularyLabel: "Japanese", vocabularyLanguageCode: "JP",
    presetEnabled: true,
    meaningAttributes: [{ position: 1, label: "English", languageCode: "EN" }],
    optionalAttributes: LANGUAGE_PRESET_DEFINITIONS.JP.optionalAttributes.map((field, index) => ({ ...field, required: false, visible: false, displayOrder: index + 1, provenance: "preset" })),
    tagTypes: [{ name: "Part of Speech", tags: [{ name: "名詞" }] }],
  };
}

test("fresh databases use the v2 schema and migrations are idempotent", () => {
  const temp = temporaryDatabase();
  try {
    const repository = new VocabularyRepository(temp.path); repository.close();
    const db = new DatabaseSync(temp.path);
    assert.equal((db.prepare("SELECT COUNT(*) AS count FROM schema_migrations").get() as { count: number }).count, 2);
    assert.equal((db.prepare("SELECT COUNT(*) AS count FROM sqlite_master WHERE type='table' AND name IN ('pos_tags','entry_pos_tags')").get() as { count: number }).count, 0);
    assert.equal((db.prepare("SELECT COUNT(*) AS count FROM sqlite_master WHERE type='table' AND name LIKE 'mvp_%'").get() as { count: number }).count, 0);
    assertDatabaseIntegrity(db); db.close();
    const reopened = new VocabularyRepository(temp.path); reopened.close();
  } finally { temp.cleanup(); }
});

test("a failed migration rolls back its schema and ledger row", () => {
  const temp = temporaryDatabase();
  try {
    const db = new DatabaseSync(temp.path); runSchemaMigrations(db);
    assert.throws(() => runSchemaMigrations(db, [{ version: 3, apply(inner) { inner.exec("CREATE TABLE rollback_probe (id INTEGER); INSERT INTO missing_table VALUES (1)"); } }]), /migration 3 failed/i);
    assert.equal((db.prepare("SELECT COUNT(*) AS count FROM sqlite_master WHERE name='rollback_probe'").get() as { count: number }).count, 0);
    assert.equal((db.prepare("SELECT COUNT(*) AS count FROM schema_migrations WHERE version=3").get() as { count: number }).count, 0);
    db.close();
  } finally { temp.cleanup(); }
});

test("entry stats, priority, ownership, and cascades are enforced", () => {
  const temp = temporaryDatabase();
  try {
    const repository = new VocabularyRepository(temp.path);
    const first = repository.createConfiguredWorkbook(basicWorkbook("First"));
    const second = repository.createConfiguredWorkbook(basicWorkbook("Second"));
    const firstTag = repository.listTagTypes(first.id)[0].tags[0]; const secondTag = repository.listTagTypes(second.id)[0].tags[0];
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

test("new attributes receive blank values and the primary meaning remains required and visible", () => {
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

test("a fully populated meaning can become primary without changing field identity or values", () => {
  const temp = temporaryDatabase(); let repository: VocabularyRepository | undefined;
  try {
    repository = new VocabularyRepository(temp.path);
    const input = basicWorkbook();
    input.meaningAttributes = [
      { position: 1, label: "English", languageCode: "EN" },
      { position: 2, label: "Japanese", languageCode: "JP" },
      { position: 3, label: "Chinese", languageCode: "ZH" },
    ];
    const workbook = repository.createConfiguredWorkbook(input);
    const entry = repository.addEntry(workbook.id, "cat", "cat", ["cat", "猫", "猫"]);
    const before = repository.getWorkbook(workbook.id)!;
    const meaningFields = before.metadataAttributes.filter((field) => field.role === "meaning");
    const japanese = meaningFields[1];
    assert.deepEqual(repository.getMeaningPromotionImpact(workbook.id, japanese.id!).emptyEntryCount, 0);

    const fields = [japanese, meaningFields[0], meaningFields[2], ...before.metadataAttributes.filter((field) => field.role === "optional")];
    const updated = repository.updateWorkbookAttributes(workbook.id, { vocabularyLabel: before.vocabularyLabel, fields });
    assert.deepEqual(updated.meaningAttributes.map((field) => [field.id, field.label]), [[japanese.id, "Japanese"], [meaningFields[0].id, "English"], [meaningFields[2].id, "Chinese"]]);
    const updatedFields = updated.metadataAttributes.filter((field) => field.role === "meaning");
    assert.equal(updatedFields[0].required, true);
    assert.equal(updatedFields[0].visible, true);
    assert.equal(updatedFields[1].required, false);
    assert.equal(updatedFields[1].visible, true);
    assert.deepEqual(repository.getEntry(entry.id)?.meanings, ["猫", "cat", "猫"]);
    assert.equal(repository.getEntry(entry.id)?.meaning, "猫");
  } finally { repository?.close(); temp.cleanup(); }
});

test("an incomplete meaning cannot become primary and the failed save is atomic", () => {
  const temp = temporaryDatabase(); let repository: VocabularyRepository | undefined;
  try {
    repository = new VocabularyRepository(temp.path);
    const input = basicWorkbook();
    input.meaningAttributes = [
      { position: 1, label: "English", languageCode: "EN" },
      { position: 2, label: "Japanese", languageCode: "JP" },
    ];
    const workbook = repository.createConfiguredWorkbook(input);
    repository.addEntry(workbook.id, "cat", "cat", ["cat", "猫"]);
    repository.addEntry(workbook.id, "dog", "dog", ["dog", "  "]);
    const before = repository.getWorkbook(workbook.id)!;
    const meaningFields = before.metadataAttributes.filter((field) => field.role === "meaning");
    assert.equal(repository.getMeaningPromotionImpact(workbook.id, meaningFields[1].id!).emptyEntryCount, 1);

    const fields = [meaningFields[1], meaningFields[0], ...before.metadataAttributes.filter((field) => field.role === "optional")];
    assert.throws(() => repository!.updateWorkbookAttributes(workbook.id, { vocabularyLabel: before.vocabularyLabel, fields }), /1 entry is empty/);
    assert.deepEqual(repository.getWorkbook(workbook.id)?.meaningAttributes.map((field) => field.id), meaningFields.map((field) => field.id));
    assert.deepEqual(repository.listEntries(workbook.id).map((entry) => entry.meanings), [["cat", "猫"], ["dog", ""]]);
  } finally { repository?.close(); temp.cleanup(); }
});

test("empty workbooks allow primary promotion and enforce it for later entries", () => {
  const temp = temporaryDatabase(); let repository: VocabularyRepository | undefined;
  try {
    repository = new VocabularyRepository(temp.path);
    const input = basicWorkbook();
    input.meaningAttributes = [
      { position: 1, label: "English", languageCode: "EN" },
      { position: 2, label: "Japanese", languageCode: "JP" },
    ];
    const workbook = repository.createConfiguredWorkbook(input);
    const meaningFields = workbook.metadataAttributes.filter((field) => field.role === "meaning");
    assert.equal(repository.getMeaningPromotionImpact(workbook.id, meaningFields[1].id!).emptyEntryCount, 0);
    const fields = [meaningFields[1], meaningFields[0], ...workbook.metadataAttributes.filter((field) => field.role === "optional")];
    repository.updateWorkbookAttributes(workbook.id, { vocabularyLabel: workbook.vocabularyLabel, fields });
    assert.throws(() => repository!.addEntry(workbook.id, "cat", "", ["", "cat"]), /Japanese is required/);
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
    const newPrimary = { key: "meaning_new", role: "meaning" as const, label: "New Primary", languageCode: null, required: true, visible: true, displayOrder: 1 };
    assert.throws(() => repository!.updateWorkbookAttributes(workbook.id, { vocabularyLabel: workbook.vocabularyLabel, fields: [newPrimary, ...fields] }), /save a new meaning/i);
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

test("v1 POS data migrates to generic tag types without losing disabled data", () => {
  const temp = temporaryDatabase();
  let db: DatabaseSync | undefined;
  try {
    db = new DatabaseSync(temp.path);
    runSchemaMigrations(db, [SCHEMA_MIGRATIONS[0]]);
    const addWorkbook = db.prepare(`INSERT INTO workbooks
      (id,name,vocabulary_kind,vocabulary_label,vocabulary_language_code,preset_enabled,pos_enabled,created_at,updated_at)
      VALUES (?,?,?,?,?,?,?,?,?)`);
    addWorkbook.run(1, "Disabled", "preset_language", "Japanese", "JP", 1, 0, "2026-01-01", "2026-01-01");
    addWorkbook.run(2, "Empty enabled", "other_language", "Vocabulary", null, 0, 1, "2026-01-01", "2026-01-01");
    db.prepare("INSERT INTO workbook_fields (workbook_id,field_key,role,position,label,is_required,is_visible,provenance) VALUES (1,'meaning_1','meaning',1,'English',1,1,'custom')").run();
    db.prepare("INSERT INTO entries (id,workbook_id,vocabulary,created_at,updated_at) VALUES (11,1,'猫','2026-01-01','2026-01-01')").run();
    db.prepare("INSERT INTO entry_field_values (entry_id,field_id,workbook_id,value) SELECT 11,id,1,'cat' FROM workbook_fields WHERE workbook_id=1").run();
    db.prepare("INSERT INTO entry_stats (entry_id) VALUES (11)").run();
    db.prepare("INSERT INTO pos_tags (id,workbook_id,name,is_predefined) VALUES (91,1,'名詞',1)").run();
    db.prepare("INSERT INTO entry_pos_tags (entry_id,tag_id,workbook_id) VALUES (11,91,1)").run();
    runSchemaMigrations(db);
    assert.equal((db.prepare("SELECT COUNT(*) AS count FROM tag_types").get() as { count: number }).count, 2);
    assert.deepEqual((db.prepare("SELECT id,name FROM tags").all() as Array<{ id: number; name: string }>).map((row) => [row.id, row.name]), [[91, "名詞"]]);
    assert.deepEqual((db.prepare("SELECT entry_id,tag_id FROM entry_tags").all() as Array<{ entry_id: number; tag_id: number }>).map((row) => [row.entry_id, row.tag_id]), [[11, 91]]);
    assert.equal((db.prepare("SELECT COUNT(*) AS count FROM pragma_table_info('workbooks') WHERE name='pos_enabled'").get() as { count: number }).count, 0);
    assertDatabaseIntegrity(db);
  } finally { db?.close(); temp.cleanup(); }
});

test("generic tag updates preserve identities, support multiple assignments, and protect assigned data", () => {
  const temp = temporaryDatabase(); let repository: VocabularyRepository | undefined;
  try {
    repository = new VocabularyRepository(temp.path);
    const workbook = repository.createConfiguredWorkbook(basicWorkbook());
    let types = repository.updateWorkbookTags(workbook.id, { types: [
      { id: repository.listTagTypes(workbook.id)[0].id, name: "Part of Speech", tags: [{ name: "noun" }, { name: "verb" }, { name: "unused" }] },
      { name: "Topic", tags: [{ name: "Travel" }, { name: "noun" }] },
    ] });
    const pos = types[0]; const noun = pos.tags.find((tag) => tag.name === "noun")!; const verb = pos.tags.find((tag) => tag.name === "verb")!;
    const travel = types[1].tags.find((tag) => tag.name === "Travel")!;
    const entry = repository.addEntry(workbook.id, "run", "move quickly", ["move quickly"], {}, [noun.id, verb.id, travel.id]);
    assert.deepEqual(entry.tags.map((tag) => tag.name).sort(), ["Travel", "noun", "verb"]);
    assert.throws(() => repository!.updateWorkbookTags(workbook.id, { types: types.map((type) => type.id === pos.id ? { ...type, tags: [...type.tags, { name: "noun" }] } : type) }), /must be unique/);
    types = repository.updateWorkbookTags(workbook.id, { types: types.map((type) => type.id === pos.id ? { ...type, tags: type.tags.map((tag) => ({ ...tag, name: tag.name === "noun" ? "verb" : tag.name === "verb" ? "noun" : tag.name })) } : type) });
    assert.equal(types[0].tags.find((tag) => tag.id === noun.id)?.name, "verb");
    const withoutUnused = { types: types.map((type) => ({ ...type, tags: type.tags.filter((tag) => tag.name !== "unused") })) };
    repository.updateWorkbookTags(workbook.id, withoutUnused);
    types = repository.listTagTypes(workbook.id);
    const withoutAssigned = { types: types.map((type) => ({ ...type, tags: type.tags.filter((tag) => tag.id !== noun.id) })) };
    assert.throws(() => repository!.updateWorkbookTags(workbook.id, withoutAssigned), TagDataLossError);
    repository.updateWorkbookTags(workbook.id, withoutAssigned, true);
    assert.equal(repository.getEntry(entry.id)?.tags.some((tag) => tag.id === noun.id), false);
    types = repository.listTagTypes(workbook.id);
    const withoutTopic = { types: types.filter((type) => type.name !== "Topic") };
    assert.throws(() => repository!.updateWorkbookTags(workbook.id, withoutTopic), TagDataLossError);
    repository.updateWorkbookTags(workbook.id, withoutTopic, true);
    assert.equal(repository.listTagTypes(workbook.id).some((type) => type.name === "Topic"), false);
    assert.equal(repository.getEntry(entry.id)?.tags.some((tag) => tag.id === travel.id), false);
  } finally { repository?.close(); temp.cleanup(); }
});

test("tag drafts reject cross-workbook IDs and moving a tag between types", () => {
  const temp = temporaryDatabase(); let repository: VocabularyRepository | undefined;
  try {
    repository = new VocabularyRepository(temp.path);
    const first = repository.createConfiguredWorkbook(basicWorkbook("First"));
    const second = repository.createConfiguredWorkbook(basicWorkbook("Second"));
    const firstType = repository.listTagTypes(first.id)[0]; const secondType = repository.listTagTypes(second.id)[0];
    assert.throws(() => repository!.updateWorkbookTags(first.id, { types: [{ id: secondType.id, name: secondType.name, tags: [] }] }), /does not belong/);
    assert.throws(() => repository!.updateWorkbookTags(first.id, { types: [{ id: firstType.id, name: firstType.name, tags: [] }, { name: "Other", tags: [{ id: firstType.tags[0].id, name: firstType.tags[0].name }] }] }), /does not belong/);
  } finally { repository?.close(); temp.cleanup(); }
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
    assert.equal(entry.attributes.example_sentence_1, "A cat."); assert.equal(entry.attributes.kana, "ねこ"); assert.equal(entry.testCount, 5); assert.equal(entry.tags[0].id, 9);
    converted.close();
  } finally { temp.cleanup(); }
});
