import assert from "node:assert/strict";
import test from "node:test";
import type { EntryRow, WorkbookRow } from "./db.js";
import { createInitialState, update } from "./orchestration.js";

const workbook: WorkbookRow = { id: 1, name: "Demo", wordCount: 0, createdAt: "", vocabularyLabel: "Vocabulary", vocabularyLanguageCode: "EN", presetEnabled: true, vocabularyKind: "preset_language", importFilePath: null, meaningAttributes: [{ position: 1, label: "Primary Meaning", languageCode: null }, { position: 2, label: "Secondary Meaning", languageCode: null }], metadataAttributes: [{ key: "note", role: "optional", label: "Note", languageCode: null, required: false, visible: true, displayOrder: 1 }] };
const entry = (id: number, vocabulary: string): EntryRow => ({ id, workbookId: 1, vocabulary, meaning: "meaning", meanings: ["meaning", ""], kanaText: null, attributes: {}, tags: [], createdAt: "", updatedAt: "", testCount: 0, errorCount: 0, tier: "gray", lastTested: null, nextTestDeadline: null });

test("command transitions enter add mode and search resets page", () => {
  let state = createInitialState(workbook, [entry(1, "alpha")]);
  let transition = update(state, { type: "commandSubmitted", raw: "/add" });
  assert.equal(transition.state.mode.kind, "add");
  state = { ...transition.state, pageIndex: 4 };
  transition = update(state, { type: "searchChanged", query: "alp" });
  assert.equal(transition.state.activeSearchQuery, "alp");
  assert.equal(transition.state.pageIndex, 0);
});

test("add flow advances through meaning and optional metadata", () => {
  let state = update(createInitialState(workbook), { type: "commandSubmitted", raw: "/add" }).state;
  state = update(state, { type: "addFieldSubmitted", value: "alpha" }).state;
  state = update(state, { type: "addFieldSubmitted", value: "first" }).state;
  assert.equal(state.mode.kind, "add");
  assert.equal(state.mode.stage, "meaning");
  state = update(state, { type: "addFieldSubmitted", value: "second" }).state;
  assert.equal(state.mode.kind, "add");
  assert.equal(state.mode.stage, "metadata");
  const transition = update(state, { type: "addFieldSubmitted", value: "memo" });
  assert.equal(transition.effects[0]?.type, "addEntry");
});

test("delete requires yes confirmation", () => {
  let state = update(createInitialState(workbook, [entry(3, "gamma")]), { type: "commandSubmitted", raw: "/delete 3" }).state;
  assert.equal(state.mode.kind, "delete");
  state = update(state, { type: "deleteConfirmationSubmitted", value: "no" }).state;
  assert.equal(state.mode.kind, "command");
  state = update(createInitialState(workbook, [entry(3, "gamma")]), { type: "commandSubmitted", raw: "/delete 3" }).state;
  const transition = update(state, { type: "deleteConfirmationSubmitted", value: "yes" });
  assert.equal(transition.effects[0]?.type, "deleteEntry");
});

test("import request ids ignore stale results and commit refreshes", () => {
  let state = createInitialState(workbook);
  let transition = update(state, { type: "importPathLoadRequested", path: "one.txt" });
  state = transition.state;
  const first = state.importRequestId;
  state = update(state, { type: "importPathLoadRequested", path: "two.txt" }).state;
  const stale = update(state, { type: "importPreviewLoaded", requestId: first, path: "one.txt", preview: { totalRecords: 0, entries: [], skippedInvalid: 0, skippedDuplicates: 0, diagnostics: [], records: [] } });
  assert.equal(stale.state.mode.kind, "importLoading");
  const current = update(state, { type: "importPreviewLoaded", requestId: state.importRequestId, path: "two.txt", preview: { totalRecords: 1, entries: [{ recordNumber: 1, vocabulary: "two", meanings: ["x"], attributes: {}, tagIds: [] }], skippedInvalid: 0, skippedDuplicates: 0, diagnostics: [], records: [] } });
  assert.equal(current.state.mode.kind, "importPreview");
  const commit = update(current.state, { type: "importConfirmed" });
  assert.equal(commit.effects[0]?.type, "commitImport");
});

test("practice initial scoring and empty candidates are explicit", () => {
  let state = createInitialState(workbook);
  state = update(state, { type: "practiceStarted", candidates: [] }).state;
  assert.equal(state.practice.phase, "empty");
  state = update(state, { type: "practiceStarted", candidates: [entry(1, "answer")] }).state;
  let transition = update(state, { type: "practiceAnswerSubmitted", answer: "answer" });
  assert.equal(transition.effects[0]?.type, "recordPracticeResult");
  state = update(transition.state, { type: "practiceResultRecorded", entry: entry(1, "answer"), isCorrect: true, initialRound: true }).state;
  assert.equal(state.practice.score, 1);
});
