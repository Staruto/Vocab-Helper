import assert from "node:assert/strict";
import test from "node:test";
import { parseSimilarWords, similarWordLines } from "./similar-words.js";
import { VocabularyRepository } from "./db.js";
import { buildDetailSections } from "./detail-display.js";
import { parseLabeledTextImport } from "./import.js";
import { createInitialState, update } from "./orchestration.js";

test("similar words support optional explanations, separators, and punctuation within items", () => {
  assert.deepEqual(parseSimilarWords(" { near by } ( short distance, usually (physical) ) , {close} {隣} () "), [
    { word: "near by", explanation: "short distance, usually (physical)" },
    { word: "close" }, { word: "隣" },
  ]);
  assert.deepEqual(parseSimilarWords("{a,b}(one),{c} (two) {d}"), [
    { word: "a,b", explanation: "one" }, { word: "c", explanation: "two" }, { word: "d" },
  ]);
});

test("malformed and mixed values fall back as a whole", () => {
  for (const value of ["", "near, close", "{}", "{ } (why)", "{near", "{{near}}", "{near} (unclosed", "{near} (extra))", "{near} trailing", "prefix {near}", "{near}{close}", "{near},", "{near},, {close}"]) {
    assert.equal(parseSimilarWords(value), null, value);
  }
});

test("similar word lines retain styling boundaries and start each item on a new line", () => {
  const items = [{ word: "near", explanation: "short distance" }, { word: "close" }];
  assert.deepEqual(similarWordLines(items, 80, (text) => text.length), [
    [{ text: "near", word: true }, { text: "  short distance", word: false }],
    [{ text: "close", word: true }],
  ]);
  const lines = similarWordLines(items, 3, (text) => text.length);
  assert.deepEqual(lines.slice(0, 2), [
    [{ text: "nea", word: true }],
    [{ text: "r", word: true }, { text: "  ", word: false }],
  ]);
  assert.equal(lines.flat().map((span) => span.text).join(""), "near  short distanceclose");
  assert.ok(lines.every((line) => line.reduce((sum, span) => sum + span.text.length, 0) <= 3));
  const wideLines = similarWordLines([{ word: "隣近", explanation: "近い" }], 4, (text) => Array.from(text).reduce((sum, char) => sum + (char.charCodeAt(0) > 255 ? 2 : 1), 0));
  assert.deepEqual(wideLines, [[{ text: "隣近", word: true }], [{ text: "  近", word: false }], [{ text: "い", word: false }]]);
});

test("only the active preset identity receives structured details", () => {
  const repository = new VocabularyRepository(":memory:");
  try {
    const workbook = repository.createWorkbook("English", "Word", "EN");
    const value = "{near} (short distance) {close}";
    const entry = repository.addEntry(workbook.id, "nearby", "not far", ["not far"], { similar_words: value });
    const preset = workbook.metadataAttributes.find((field) => field.presetKey === "similar_words")!;
    const renamed = { ...preset, key: "similar_words_2", label: "Related words" };
    const fields = [renamed, { ...preset, presetKey: null, key: "custom", label: "Similar Words" }];
    const customEntry = { ...entry, attributes: { similar_words_2: value, custom: value } };
    assert.deepEqual(buildDetailSections({ ...workbook, metadataAttributes: fields }, customEntry, []).attributes, [
      { key: "Related words", value, similarWords: [{ word: "near", explanation: "short distance" }, { word: "close" }] },
      { key: "Similar Words", value },
    ]);
    assert.deepEqual(buildDetailSections({ ...workbook, metadataAttributes: [{ ...renamed, active: false }] }, customEntry, []).attributes, []);
    for (const raw of ["legacy text", "{near} (why) trailing", ""]) {
      const field = buildDetailSections({ ...workbook, metadataAttributes: [preset] }, { ...entry, attributes: { similar_words: raw } }, []).attributes[0];
      assert.deepEqual(field, { key: "Similar Words", value: raw || "—" });
    }
  } finally { repository.close(); }
});

test("add, edit, and labeled imports preserve the source and share formatted details", () => {
  const repository = new VocabularyRepository(":memory:");
  try {
    let workbook = repository.createWorkbook("English", "Word", "EN");
    workbook = repository.updateWorkbookAttributes(workbook.id, {
      vocabularyLabel: workbook.vocabularyLabel,
      fields: workbook.metadataAttributes.filter((field) => field.role !== "vocabulary").map((field) => field.presetKey ? { ...field, active: field.presetKey === "similar_words", label: field.presetKey === "similar_words" ? "Related" : field.label } : field),
    });
    const value = "{near} (short distance), {close}";
    let state = update(createInitialState(workbook), { type: "commandSubmitted", raw: "/add" }).state;
    for (const input of ["nearby", "not far"]) state = update(state, { type: "addFieldSubmitted", value: input }).state;
    const add = update(state, { type: "addFieldSubmitted", value }).effects[0];
    assert.equal(add.type, "addEntry");
    if (add.type !== "addEntry") throw new Error("Expected add effect");
    const entry = repository.addEntry(add.workbookId, add.vocabulary, add.meanings[0], add.meanings, add.attributes, add.tagIds);
    assert.equal(entry.attributes.similar_words, value);
    const expected = buildDetailSections(workbook, entry, []).attributes;
    state = update(createInitialState(workbook, [entry]), { type: "commandSubmitted", raw: `/edit ${entry.id}` }).state;
    for (const input of ["nearby", "not far"]) state = update(state, { type: "editFieldSubmitted", value: input }).state;
    assert.equal(state.mode.kind, "edit");
    if (state.mode.kind !== "edit") throw new Error("Expected edit form");
    assert.equal(state.mode.metadata.similar_words, value);
    const edit = update(state, { type: "editFieldSubmitted", value }).effects[0];
    assert.equal(edit.type, "updateEntry");
    if (edit.type !== "updateEntry") throw new Error("Expected edit effect");
    const edited = repository.updateEntry(edit.entryId, edit.vocabulary, edit.meanings[0], edit.meanings, edit.attributes, edit.tagIds);
    assert.deepEqual(buildDetailSections(workbook, edited, []).attributes, expected);
    const preview = parseLabeledTextImport(`Word: nearby\nPrimary Meaning: not far\nRelated: ${value}`, workbook, [], [entry.vocabulary]);
    assert.deepEqual(preview.diagnostics.map((item) => item.kind), ["duplicate"]);
    assert.equal(preview.entries[0].attributes.similar_words, value);
    assert.equal(repository.importEntries(workbook.id, preview.entries).unchanged, 1);
    assert.deepEqual(buildDetailSections(workbook, repository.getEntry(entry.id)!, []).attributes, expected);
    const newPreview = parseLabeledTextImport(`Word: adjacent\nPrimary Meaning: next to\nRelated: ${value}`, workbook, [], []);
    const imported = repository.importEntries(workbook.id, newPreview.entries);
    assert.equal(imported.added, 1);
    assert.deepEqual(buildDetailSections(workbook, imported.entries[0], []).attributes, expected);
  } finally { repository.close(); }
});
