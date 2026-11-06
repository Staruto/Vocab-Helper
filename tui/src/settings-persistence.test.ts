import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import test from "node:test";
import { localCalendarDay } from "./calendar.js";
import { TagDataLossError, VocabularyRepository, WorkbookDataLossError } from "./db.js";

function fixture() {
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-settings-"));
  const repository = new VocabularyRepository(join(directory, "vocab.db"));
  const workbook = repository.createConfiguredWorkbook({
    name: "Settings", vocabularyKind: "non_language", vocabularyLabel: "Word", vocabularyLanguageCode: null,
    meaningAttributes: [{ position: 1, label: "Meaning", languageCode: null }], presetAttributes: [],
    optionalAttributes: [{ key: "note", role: "optional", label: "Note", languageCode: null, required: false, visible: false, displayOrder: 1 }],
    tagTypes: [{ name: "Topic", visible: true, tags: [{ name: "Science" }] }],
  });
  const types = repository.listTagTypes(workbook.id);
  const entry = repository.addEntry(workbook.id, "atom", "small", ["small"], { note: "retain me" }, [types[0].tags[0].id]);
  repository.setWorkbookDailyGoal(workbook.id, 10);
  repository.completePracticeSession(workbook.id, [{ entryId: entry.id, answers: [true] }], "2020-01-01");
  repository.completePracticeSession(workbook.id, [{ entryId: entry.id, answers: [true] }], localCalendarDay());
  return { repository, workbook, types, entry, cleanup() { repository.close(); rmSync(directory, { recursive: true, force: true }); } };
}

test("combined settings save retains identities, values, assignments and practice history", () => {
  const f = fixture();
  const { repository: r, workbook: w } = f;
  try {
    const before = r.getEntry(f.entry.id)!;
    const stats = r.getEntryStats(f.entry.id);
    const updated = r.updateWorkbookSettingsSections(w.id, {
      attributes: { vocabularyLabel: "Term", fields: w.metadataAttributes.filter((field) => field.role !== "vocabulary").map((field) => ({ ...field, visible: true })) },
      tags: { types: f.types.map((type) => ({ ...type, name: "Subject" })) },
      importFilePath: "words.txt", dailyGoal: 25,
    });
    assert.equal(updated.vocabularyLabel, "Term");
    assert.equal(updated.dailyGoal, 25);
    assert.equal(updated.importFilePath, resolve("words.txt"));
    assert.deepEqual(updated.metadataAttributes.map((field) => field.id), w.metadataAttributes.map((field) => field.id));
    assert.equal(r.listTagTypes(w.id)[0].id, f.types[0].id);
    const after = r.getEntry(f.entry.id)!;
    assert.deepEqual(after.attributes, before.attributes);
    assert.deepEqual(after.tags.map((tag) => tag.id), before.tags.map((tag) => tag.id));
    assert.deepEqual(r.getEntryStats(after.id), stats);
    const history = r.listDailyPractice(w.id);
    assert.equal(history.find((day) => day.day === "2020-01-01")?.goalSnapshot, 10);
    assert.equal(history.find((day) => day.day === localCalendarDay())?.goalSnapshot, 25);
  } finally { f.cleanup(); }
});

test("a later section failure rolls back attribute deletion and preserves all saved state", () => {
  const f = fixture();
  const { repository: r, workbook: w } = f;
  try {
    const before = { workbook: r.getWorkbook(w.id), entry: r.getEntry(f.entry.id), types: r.listTagTypes(w.id), history: r.listDailyPractice(w.id) };
    assert.throws(() => r.updateWorkbookSettingsSections(w.id, {
      attributes: { vocabularyLabel: "Changed", fields: w.metadataAttributes.filter((field) => field.role === "meaning") },
      tags: { types: [{ ...f.types[0], name: "" }] }, importFilePath: "changed.txt", dailyGoal: 99,
    }, true), /Tag type name is required/);
    assert.deepEqual({ workbook: r.getWorkbook(w.id), entry: r.getEntry(f.entry.id), types: r.listTagTypes(w.id), history: r.listDailyPractice(w.id) }, before);
  } finally { f.cleanup(); }
});

test("combined destructive edits require confirmation for fields and tags", () => {
  const f = fixture();
  const { repository: r, workbook: w } = f;
  try {
    const attributes = { vocabularyLabel: "Term", fields: w.metadataAttributes.filter((field) => field.role === "meaning") };
    assert.throws(() => r.updateWorkbookSettingsSections(w.id, { attributes, tags: { types: [] } }), WorkbookDataLossError);
    const nonDestructiveAttributes = { ...attributes, fields: w.metadataAttributes.filter((field) => field.role !== "vocabulary") };
    assert.throws(() => r.updateWorkbookSettingsSections(w.id, { attributes: nonDestructiveAttributes, tags: { types: [] } }), TagDataLossError);
    assert.equal(r.getWorkbook(w.id)?.vocabularyLabel, "Word");
    assert.equal(r.getEntry(f.entry.id)?.attributes.note, "retain me");
    assert.equal(r.getEntry(f.entry.id)?.tags.length, 1);
    r.updateWorkbookSettingsSections(w.id, { attributes, tags: { types: [] }, dailyGoal: null }, true);
    assert.equal(r.getWorkbook(w.id)?.dailyGoal, null);
    assert.equal(r.getEntry(f.entry.id)?.attributes.note, undefined);
    assert.deepEqual(r.getEntry(f.entry.id)?.tags, []);
  } finally { f.cleanup(); }
});
