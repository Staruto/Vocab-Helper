import assert from "node:assert/strict";
import test from "node:test";
import { parseExampleSentence, wrapStyledText } from "./styled-text.js";
import { VocabularyRepository } from "./db.js";
import { buildDetailSections } from "./detail-display.js";
import { parseLabeledTextImport } from "./import.js";
import { createDetailEditor, detailEditRows, saveDetailEditor, updateDetailEditor } from "./detail-editor.js";
import { createInitialState, update } from "./orchestration.js";

test("sentence highlighting preserves surrounding text, whitespace, and adjacent pairs", () => {
  assert.deepEqual(parseExampleSentence("She lives {near the station}."), [
    { text: "She lives ", highlight: false }, { text: "near the station", highlight: true }, { text: ".", highlight: false },
  ]);
  assert.deepEqual(parseExampleSentence("{ 猫 }{と犬}, {} too"), [
    { text: " 猫 ", highlight: true }, { text: "と犬", highlight: true },
    { text: ", ", highlight: false }, { text: " too", highlight: false },
  ]);
  assert.deepEqual(parseExampleSentence("{}"), []);
  for (const plain of ["plain text", "", "{unclosed", "unopened}", "{{nested}}", "{outer {inner} text}"]) {
    assert.equal(parseExampleSentence(plain), null, plain);
  }
});

test("independent complete pairs highlight alongside unmatched braces and nested groups", () => {
  assert.deepEqual(parseExampleSentence("} {yes} {{no}} {also} {broken"), [
    { text: "} ", highlight: false }, { text: "yes", highlight: true },
    { text: " {{no}} ", highlight: false }, { text: "also", highlight: true },
    { text: " {broken", highlight: false },
  ]);
  assert.deepEqual(parseExampleSentence("{broken {yes}"), [
    { text: "{broken ", highlight: false }, { text: "yes", highlight: true },
  ]);
});

test("styled wrapping preserves highlighting across line breaks and handles wide text", () => {
  const spans = parseExampleSentence("a {bcdef}!")!;
  assert.deepEqual(wrapStyledText(spans, 4, (text) => text.length), [
    [{ text: "a ", highlight: false }, { text: "bc", highlight: true }],
    [{ text: "def", highlight: true }, { text: "!", highlight: false }],
  ]);
  const wideMeasure = (text: string) => Array.from(text).reduce((sum, char) => sum + (char.charCodeAt(0) > 255 ? 2 : 1), 0);
  assert.deepEqual(wrapStyledText(parseExampleSentence("猫{と犬}。")!, 4, wideMeasure), [
    [{ text: "猫", highlight: false }, { text: "と", highlight: true }],
    [{ text: "犬", highlight: true }, { text: "。", highlight: false }],
  ]);
  assert.deepEqual(wrapStyledText([], 4, wideMeasure), [[]]);
  assert.deepEqual(spans, parseExampleSentence("a {bcdef}!"));
});

test("all five sentence presets use their identity, including renamed and collision-adjusted fields", () => {
  const repository = new VocabularyRepository(":memory:");
  try {
    const workbook = repository.createWorkbook("English", "Word", "EN");
    const value = "A {sentence}.";
    const presets = workbook.metadataAttributes.filter((field) => /^example_sentence_[1-5]$/.test(field.presetKey ?? ""));
    assert.equal(presets.length, 5);
    const fields = presets.map((field, index) => ({ ...field, key: `collision_${index}`, label: `Renamed ${index}` }));
    const attributes = Object.fromEntries(fields.map((field) => [field.key, value]));
    const entry = repository.addEntry(workbook.id, "sentence", "a phrase");
    const sections = buildDetailSections({ ...workbook, metadataAttributes: fields }, { ...entry, attributes }, []);
    assert.deepEqual(sections.attributes, fields.map((field) => ({ key: field.label, value, spans: parseExampleSentence(value) })));
    assert.deepEqual(buildDetailSections({ ...workbook, metadataAttributes: fields.map((field) => ({ ...field, active: false })) }, { ...entry, attributes }, []).attributes, []);
    const custom = { ...presets[0], presetKey: null, key: "custom" };
    assert.deepEqual(buildDetailSections({ ...workbook, metadataAttributes: [custom] }, { ...entry, attributes: { custom: value } }, []).attributes, [{ key: "Example Sentence 1", value }]);
    for (const raw of ["plain", "{broken", ""]) {
      assert.deepEqual(buildDetailSections({ ...workbook, metadataAttributes: [presets[0]] }, { ...entry, attributes: { [presets[0].key]: raw } }, []).attributes, [{ key: "Example Sentence 1", value: raw || "—" }]);
    }
  } finally { repository.close(); }
});

test("add, edit, and import keep sentence braces in storage and produce the same styled details", () => {
  const repository = new VocabularyRepository(":memory:");
  try {
    let workbook = repository.createWorkbook("English", "Word", "EN");
    workbook = repository.updateWorkbookAttributes(workbook.id, {
      vocabularyLabel: workbook.vocabularyLabel,
      fields: workbook.metadataAttributes.filter((field) => field.role !== "vocabulary").map((field) => field.presetKey ? { ...field, active: /^example_sentence_[1-5]$/.test(field.presetKey), label: field.presetKey === "example_sentence_1" ? "Usage" : field.label } : field),
    });
    let state = update(createInitialState(workbook), { type: "commandSubmitted", raw: "/add" }).state;
    const original = "She lives {near the station}.";
    for (const value of ["near", "not far", ...Array(4).fill(original)]) state = update(state, { type: "addFieldSubmitted", value }).state;
    const add = update(state, { type: "addFieldSubmitted", value: original }).effects[0];
    assert.equal(add.type, "addEntry");
    if (add.type !== "addEntry") throw new Error("Expected add effect");
    const entry = repository.addEntry(add.workbookId, add.vocabulary, add.meanings[0], add.meanings, add.attributes, add.tagIds);
    for (let index = 1; index <= 5; index++) assert.equal(entry.attributes[`example_sentence_${index}`], original);
    let editor = createDetailEditor(entry, true);
    assert.equal(editor.draft.attributes.example_sentence_1, original);
    const changed = "She lives {nearby}, {too}.";
    for (const [index, row] of detailEditRows(workbook, []).entries()) {
      if (row.kind !== "attribute") continue;
      editor = updateDetailEditor(editor, { type: "move", delta: index - editor.selected }, workbook, []);
      editor = updateDetailEditor(editor, { type: "change", value: changed }, workbook, []);
    }
    editor = updateDetailEditor(editor, { type: "exit" }, workbook, []);
    editor = saveDetailEditor(editor, workbook, (draft, options) => repository.updateEntry(entry.id, draft.vocabulary, draft.meanings[0], draft.meanings, draft.attributes, draft.tagIds, options));
    assert.equal(editor.mode, "view");
    const edited = editor.original;
    for (let index = 1; index <= 5; index++) assert.equal(edited.attributes[`example_sentence_${index}`], changed);
    const expected = buildDetailSections(workbook, edited, []).attributes;
    assert.ok(expected.every((field) => field.spans?.some((span) => span.highlight)));
    const labels = workbook.metadataAttributes.filter((field) => field.role === "optional" && field.active).map((field) => `${field.label}: ${changed}`).join("\n");
    const preview = parseLabeledTextImport(`Word: near\nPrimary Meaning: not far\n${labels}`, workbook, [], [entry.vocabulary]);
    assert.equal(repository.importEntries(workbook.id, preview.entries).unchanged, 1);
    assert.deepEqual(buildDetailSections(workbook, repository.getEntry(entry.id)!, []).attributes, expected);
    const imported = parseLabeledTextImport(`Word: close\nPrimary Meaning: not far\n${labels}`, workbook, [], []);
    const result = repository.importEntries(workbook.id, imported.entries);
    assert.equal(result.added, 1);
    assert.deepEqual(buildDetailSections(workbook, result.entries[0], []).attributes, expected);
  } finally { repository.close(); }
});
