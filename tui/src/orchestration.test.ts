import assert from "node:assert/strict";
import test from "node:test";
import type { EntryRow, WorkbookRow } from "./db.js";
import { createInitialState, update } from "./orchestration.js";

const workbook: WorkbookRow = { id: 1, name: "Demo", wordCount: 0, createdAt: "", vocabularyLabel: "Vocabulary", vocabularyLanguageCode: "EN", vocabularyKind: "preset_language", importFilePath: null, listOrder: "add_asc", dailyGoal: null, todayPracticeCount: 0, meaningAttributes: [{ position: 1, label: "Primary Meaning", languageCode: null }, { position: 2, label: "Secondary Meaning", languageCode: null }], metadataAttributes: [{ key: "note", role: "optional", label: "Note", languageCode: null, required: false, visible: true, displayOrder: 1 }] };
const entry = (id: number, vocabulary: string): EntryRow => ({ id, workbookId: 1, vocabulary, meaning: "meaning", meanings: ["meaning", ""], kanaText: null, attributes: {}, tags: [], createdAt: "", updatedAt: "", testCount: 0, errorCount: 0, tier: "gray", lastTested: null, nextTestDeadline: null });

test("search resets page", () => {
  const state = { ...createInitialState({ ...workbook, metadataAttributes: [{ key: "vocab", role: "vocabulary", label: "Vocabulary", languageCode: null, required: true, visible: true, displayOrder: 0 }] }, [entry(1, "alpha")]), pageIndex: 4 };
  const transition = update(state, { type: "searchChanged", query: "alp" });
  assert.equal(transition.state.activeSearchQuery, "alp");
  assert.equal(transition.state.pageIndex, 0);
  assert.equal(transition.state.selectedEntryId, 1);
});

test("list navigation and resets use filtered entries and preserve command drafts", () => {
  const entries = Array.from({ length: 50 }, (_, index) => entry(index + 1, index % 2 === 0 ? `match${index}` : `other${index}`));
  const searchableWorkbook = { ...workbook, metadataAttributes: [{ key: "vocab", role: "vocabulary" as const, label: "Vocabulary", languageCode: null, required: true, visible: true, displayOrder: 0 }] };
  let state = { ...createInitialState(searchableWorkbook, entries), commandBuffer: "draft" };
  state = update(state, { type: "listFocusMoved", delta: -1 }).state;
  assert.equal(state.inputFocus, "search");
  assert.equal(state.commandBuffer, "draft");
  state = update(state, { type: "searchChanged", query: "match" }).state;
  assert.equal(state.selectedEntryId, 1);
  state = update(state, { type: "listFocusMoved", delta: -1 }).state;
  assert.equal(state.inputFocus, "table");
  assert.equal(state.selectedEntryId, 39);
  state = update(state, { type: "pageChanged", delta: 1 }).state;
  assert.equal(state.pageIndex, 1);
  assert.equal(state.selectedEntryId, 49);
  assert.equal(update(state, { type: "pageChanged", delta: 1 }).state.pageIndex, 1);
  state = update(state, { type: "searchChanged", query: "missing" }).state;
  assert.equal(state.pageIndex, 0);
  assert.equal(state.selectedEntryId, null);
  assert.equal(state.inputFocus, "search");
  state = update(state, { type: "searchChanged", query: "match" }).state;
  state = update(state, { type: "listOrderSaved", workbook: { ...searchableWorkbook, listOrder: "add_desc" }, entries: [...entries].reverse() }).state;
  assert.equal(state.pageIndex, 0);
  assert.equal(state.selectedEntryId, 49);
  state = update(state, { type: "entriesLoaded", entries: [] }).state;
  assert.equal(state.selectedEntryId, null);
  assert.equal(state.pageIndex, 0);
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
  assert.equal(commit.effects[0]?.type, "checkImportUpdates");
  const checked = update(commit.state, { type: "importUpdatesChecked", updated: 0 });
  assert.equal(checked.effects[0]?.type, "commitImport");
});

test("practice initial scoring and empty candidates are explicit", () => {
  let state = createInitialState(workbook);
  state = update(state, { type: "practiceStarted", candidates: [] }).state;
  assert.equal(state.practice.phase, "empty");
  state = update(state, { type: "practiceStarted", candidates: [entry(1, "answer")] }).state;
  let transition = update(state, { type: "practiceAnswerSubmitted", answer: "answer" });
  assert.deepEqual(transition.effects, []);
  state = transition.state;
  assert.equal(state.practice.score, 1);
});

test("practice candidate loading is an effect and commit failures stay visible", () => {
  let state = createInitialState(workbook);
  let transition = update(state, { type: "practiceCandidatesRequested", workbookId: workbook.id, count: 3 });
  assert.equal(transition.effects[0]?.type, "loadPracticeCandidates");
  state = update(state, { type: "commandSubmitted", raw: "/import" }).state;
  state = update(state, { type: "importPreviewLoaded", requestId: state.importRequestId, path: "words.txt", preview: { totalRecords: 1, entries: [{ recordNumber: 1, vocabulary: "word", meanings: ["meaning"], attributes: {}, tagIds: [] }], skippedInvalid: 0, skippedDuplicates: 0, diagnostics: [], records: [] } }).state;
  transition = update(state, { type: "importCommitFailed", error: "Import failed; no entries were written." });
  assert.equal(transition.state.mode.kind, "importPreview");
  assert.equal(transition.state.mode.error, "Import failed; no entries were written.");
});

test("import checking and saving block duplicate inputs and failures restore the preview", () => {
  const preview = { totalRecords: 1, entries: [{ recordNumber: 1, vocabulary: "word", meanings: ["new"], attributes: {}, tagIds: [] }], skippedInvalid: 0, skippedDuplicates: 0, diagnostics: [], records: [] };
  const initial = { ...createInitialState(workbook), mode: {
    kind: "importPreview" as const, path: "words.txt", pathBuffer: "draft.txt", preview, filter: "records" as const, pageIndex: 2, focus: "records" as const,
  } };
  let state = update(initial, { type: "importConfirmed" }).state;
  for (const intent of [{ type: "importConfirmed" }, { type: "modeCancelled" }, { type: "commandSubmitted", raw: "/import" }] as const) {
    assert.deepEqual(update(state, intent), { state, effects: [] });
  }
  const failedCheck = update(state, { type: "importCommitFailed", error: "Check failed" }).state;
  assert.deepEqual(failedCheck.mode, { ...initial.mode, error: "Check failed" });
  state = update(state, { type: "importUpdatesChecked", updated: 1 }).state;
  state = update(state, { type: "importChoiceMoved", delta: -1 }).state;
  assert.ok(state.mode.kind === "importConfirm");
  assert.equal(state.mode.choice, 3);
  state = update(state, { type: "importChoiceMoved", delta: 1 }).state;
  const saving = update(state, { type: "importChoiceConfirmed" });
  assert.equal(saving.effects[0]?.type, "commitImport");
  assert.ok(saving.effects[0]?.type === "commitImport");
  assert.deepEqual(saving.effects[0].options, { preserveTestData: true });
  assert.deepEqual(update(saving.state, { type: "importChoiceConfirmed" }).effects, []);
  assert.deepEqual(update(saving.state, { type: "importConfirmed" }).effects, []);
  const failedSave = update(saving.state, { type: "importCommitFailed", error: "Write failed" }).state;
  assert.deepEqual(failedSave.mode, { ...initial.mode, error: "Write failed" });
  assert.equal(update(failedSave, { type: "importConfirmed" }).effects[0]?.type, "checkImportUpdates");
});
