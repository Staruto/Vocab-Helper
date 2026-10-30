import assert from "node:assert/strict";
import test from "node:test";
import { VocabularyRepository } from "./db.js";
import type { EntryRow, EntryUpdateOptions, TagType, WorkbookRow } from "./db.js";
import { createDetailEditor, detailEditRows, detailEditorKeyAction, detailHasChanges, detailRowChanged, detailRowValue, saveDetailEditor, updateDetailEditor } from "./detail-editor.js";
import type { DetailEditorAction, DetailEditorState } from "./detail-editor.js";

const workbook: WorkbookRow = {
  id: 1, name: "Demo", wordCount: 1, createdAt: "", vocabularyLabel: "Word", vocabularyLanguageCode: null, vocabularyKind: "non_language", importFilePath: null, dailyGoal: null, todayPracticeCount: 0,
  meaningAttributes: [{ position: 1, label: "Meaning", languageCode: null }, { position: 2, label: "Other meaning", languageCode: null }],
  metadataAttributes: [
    { key: "usage", role: "optional", label: "Usage", languageCode: null, required: false, visible: false, displayOrder: 1, presetKey: "example_sentence_1", active: true },
    { key: "inactive", role: "optional", label: "Inactive", languageCode: null, required: false, visible: false, displayOrder: 2, presetKey: "example_sentence_2", active: false },
  ],
};
const tagTypes: TagType[] = [
  { id: 1, workbookId: 1, name: "Visible", position: 1, visible: true, tags: [{ id: 10, tagTypeId: 1, name: "one" }, { id: 11, tagTypeId: 1, name: "two" }] },
  { id: 2, workbookId: 1, name: "Hidden", position: 2, visible: false, tags: [{ id: 20, tagTypeId: 2, name: "secret" }] },
  { id: 3, workbookId: 1, name: "Empty", position: 3, visible: false, tags: [] },
];
const entry: EntryRow = {
  id: 1, workbookId: 1, vocabulary: "near", meaning: "not far", meanings: ["not far", ""], kanaText: null, attributes: { usage: "Lives {nearby}.", inactive: " preserved " }, tags: [tagTypes[0].tags[0], tagTypes[1].tags[0]],
  createdAt: "", updatedAt: "", testCount: 5, errorCount: 1, tier: "yellow", lastTested: "2026-10-30", nextTestDeadline: "2026-11-06",
};
const rows = detailEditRows(workbook, tagTypes);
const act = (state: DetailEditorState, action: DetailEditorAction) => updateDetailEditor(state, action, workbook, tagTypes);

test("editing starts on vocabulary, includes hidden and empty tag types, and never selects Review or inactive attributes", () => {
  const initial = createDetailEditor(entry);
  const editing = act(initial, { type: "begin" });
  assert.equal(editing.mode, "edit");
  assert.equal(editing.selected, 0);
  assert.deepEqual(rows.map((row) => row.label), ["Word", "Meaning", "Other meaning", "Usage", "Visible", "Hidden", "Empty"]);
  assert.equal(act(editing, { type: "move", delta: -1 }).selected, 0);
  assert.equal(act(editing, { type: "move", delta: 100 }).selected, rows.length - 1);
  assert.equal(act(editing, { type: "exit" }).mode, "view");
  assert.equal(initial.mode, "view");
});

test("text changes are staged across row movement and restoring values removes dirty markers", () => {
  let state = createDetailEditor(entry, true);
  state = act(state, { type: "change", value: "nearer" });
  assert.ok(detailRowChanged(state, rows[0], tagTypes));
  state = act(state, { type: "move", delta: 3 });
  assert.equal(detailRowValue(state, rows[3]), "Lives {nearby}.");
  state = act(state, { type: "change", value: "" });
  assert.equal(state.draft.attributes.inactive, " preserved ");
  assert.ok(detailHasChanges(state, workbook, tagTypes));
  state = act(state, { type: "move", delta: -3 });
  assert.equal(detailRowValue(state, rows[0]), "nearer");
  state = act(state, { type: "change", value: "near" });
  assert.equal(detailRowChanged(state, rows[0], tagTypes), false);
  state = act(state, { type: "move", delta: 3 });
  state = act(state, { type: "change", value: entry.attributes.usage });
  assert.equal(detailHasChanges(state, workbook, tagTypes), false);
  assert.deepEqual(entry.attributes, { usage: "Lives {nearby}.", inactive: " preserved " });
});

test("tag picker stages only its type, supports cancellation, and ignores assignment order", () => {
  let state = act(createDetailEditor(entry, true), { type: "move", delta: 4 });
  const original = state;
  state = act(state, { type: "openTags" });
  state = act(state, { type: "move", delta: 1 });
  state = act(state, { type: "toggleTag" });
  assert.deepEqual(state.draft.tagIds, [10, 20]);
  assert.deepEqual(act(state, { type: "cancelTags" }), original);
  state = act(state, { type: "acceptTags" });
  assert.deepEqual(new Set(state.draft.tagIds), new Set([10, 11, 20]));
  assert.ok(detailRowChanged(state, rows[4], tagTypes));
  assert.equal(detailRowChanged(state, rows[5], tagTypes), false);
  state = act(state, { type: "openTags" });
  state = act(state, { type: "move", delta: 1 });
  state = act(state, { type: "toggleTag" });
  state = act(state, { type: "acceptTags" });
  assert.equal(detailHasChanges(state, workbook, tagTypes), false);
  state = act(state, { type: "move", delta: 2 });
  state = act(state, { type: "openTags" });
  assert.equal(act(state, { type: "toggleTag" }), state);
  assert.equal(act(state, { type: "exit" }).mode, "view");
});

test("Ctrl+S in the picker stages pending selections before showing exit confirmation", () => {
  let state = act(createDetailEditor(entry, true), { type: "move", delta: 5 });
  state = act(state, { type: "openTags" });
  state = act(state, { type: "toggleTag" });
  state = act(state, { type: "exit" });
  assert.equal(state.mode, "confirm");
  assert.deepEqual(state.draft.tagIds, [10]);
  assert.ok(detailRowChanged(state, rows[5], tagTypes));
});

test("exit confirmation supports continue, discard, and unchanged exits without writing", () => {
  let state = act(createDetailEditor(entry, true), { type: "change", value: "changed" });
  let writes = 0;
  const write = () => { writes++; return entry; };
  assert.equal(saveDetailEditor(state, workbook, write), state);
  state = act(state, { type: "exit" });
  assert.equal(state.mode, "confirm");
  assert.equal(state.exitChoice, 0);
  state = act(state, { type: "moveChoice", delta: -1 });
  assert.equal(state.exitChoice, 3);
  assert.equal(saveDetailEditor(state, workbook, write), state);
  assert.equal(act(state, { type: "moveChoice", delta: 1 }).exitChoice, 0);
  const discardChoice = act(state, { type: "moveChoice", delta: -1 });
  assert.equal(discardChoice.exitChoice, 2);
  assert.equal(saveDetailEditor(discardChoice, workbook, write), discardChoice);
  state = act(state, { type: "continue" });
  assert.equal(state.mode, "edit");
  assert.equal(state.draft.vocabulary, "changed");
  state = act(state, { type: "exit" });
  state = act(state, { type: "discard" });
  assert.equal(state.mode, "view");
  assert.equal(state.draft.vocabulary, entry.vocabulary);
  assert.equal(writes, 0);
});

for (const preserveTestData of [true, false]) {
  test(`validation and write failures preserve drafts with preserveTestData=${preserveTestData}`, () => {
    const confirm = (state: DetailEditorState) => {
      const exiting = act(state, { type: "exit" });
      return preserveTestData ? exiting : act(exiting, { type: "moveChoice", delta: 1 });
    };
    const fail = () => { throw new Error("Database unavailable"); };
    let state = act(createDetailEditor(entry, true), { type: "change", value: " " });
    state = saveDetailEditor(confirm(state), workbook, fail);
    assert.equal(state.mode, "edit");
    assert.equal(state.selected, 0);
    assert.match(state.error, /Vocabulary is required/);
    state = act(state, { type: "change", value: "fixed" });
    state = act(state, { type: "move", delta: 1 });
    state = act(state, { type: "change", value: "" });
    state = saveDetailEditor(confirm(state), workbook, fail);
    assert.equal(state.selected, 1);
    assert.match(state.error, /Meaning is required/);
    state = act(state, { type: "change", value: "restored" });
    state = saveDetailEditor(confirm(state), workbook, fail);
    assert.equal(state.mode, "edit");
    assert.equal(state.error, "Database unavailable");
    assert.equal(state.draft.vocabulary, "fixed");
    assert.equal(state.draft.meanings[0], "restored");
  });
}

test("keyboard ownership keeps entry navigation out of edit, picker, and confirmation modes", () => {
  assert.deepEqual(detailEditorKeyAction("view", "e", { ctrl: true }), { type: "begin" });
  for (const key of [{ escape: true }, { ctrl: true }]) assert.deepEqual(detailEditorKeyAction("edit", "s", key), { type: "exit" });
  assert.deepEqual(detailEditorKeyAction("tags", "", { escape: true }), { type: "cancelTags" });
  assert.deepEqual(detailEditorKeyAction("tags", "s", { ctrl: true }), { type: "exit" });
  assert.deepEqual(detailEditorKeyAction("confirm", "", { escape: true }), { type: "continue" });
  assert.equal(detailEditorKeyAction("confirm", "", { return: true }), "confirmChoice");
  for (const mode of ["edit", "tags", "confirm"] as const) {
    assert.equal(detailEditorKeyAction(mode, "", { leftArrow: true }), null);
    assert.equal(detailEditorKeyAction(mode, "", { rightArrow: true }), null);
  }
});

for (const preserveTestData of [true, false]) {
  test(`confirmed saving ${preserveTestData ? "preserves" : "resets"} learning state; discard leaves the database untouched`, () => {
    const repository = new VocabularyRepository(":memory:");
    try {
      const workbook = repository.createWorkbook("Editor");
      const added = repository.addEntry(workbook.id, "word", "meaning");
      const [original] = repository.completePracticeSession(workbook.id, [{ entryId: added.id, answers: [false, true] }], "2026-10-30");
      const update = (state: DetailEditorState, action: DetailEditorAction) => updateDetailEditor(state, action, workbook, []);
      const write = (draft: DetailEditorState["draft"], options: EntryUpdateOptions) => {
        assert.equal(options.preserveTestData, preserveTestData);
        return repository.updateEntry(added.id, draft.vocabulary, draft.meanings[0], draft.meanings, draft.attributes, draft.tagIds, options);
      };
      let state = update(createDetailEditor(original, true), { type: "change", value: " changed " });
      assert.deepEqual(repository.getEntry(added.id), original);
      assert.deepEqual(update(update(state, { type: "exit" }), { type: "discard" }).original, original);
      state = update(state, { type: "exit" });
      if (!preserveTestData) state = update(state, { type: "moveChoice", delta: 1 });
      state = saveDetailEditor(state, workbook, write);
      assert.equal(state.mode, "view");
      assert.equal(state.original.vocabulary, "changed");
      assert.equal(state.original.tier, preserveTestData ? original.tier : "gray");
      assert.equal(state.original.testCount, preserveTestData ? original.testCount : 0);
      assert.equal(state.original.lastTested, preserveTestData ? original.lastTested : null);
      assert.deepEqual(repository.getEntry(added.id), state.original);
      assert.equal(detailHasChanges(state, workbook, []), false);
      const [tested] = repository.completePracticeSession(workbook.id, [{ entryId: added.id, answers: [true] }], "2026-10-31");
      state = update(createDetailEditor(tested, true), { type: "change", value: " changed " });
      state = update(state, { type: "exit" });
      if (!preserveTestData) state = update(state, { type: "moveChoice", delta: 1 });
      state = saveDetailEditor(state, workbook, write);
      assert.deepEqual(state.original, tested);
    } finally { repository.close(); }
  });
}
