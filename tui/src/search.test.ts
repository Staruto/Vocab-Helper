import assert from "node:assert/strict";
import test from "node:test";
import { EntryRow, TagType, WorkbookRow } from "./db.js";
import { filterEntriesForSearch } from "./search.js";

const workbook = { id: 1, metadataAttributes: [
  { key: "vocab", label: "Word", visible: true }, { key: "meaning_1", label: "Meaning", visible: true },
  { key: "note", label: "Note", visible: true }, { key: "internal", label: "Internal", visible: false },
] } as WorkbookRow;
const tags: TagType[] = [
  { id: 1, workbookId: 1, name: "Part of Speech", position: 1, visible: true, tags: [{ id: 10, tagTypeId: 1, name: "Verb" }] },
  { id: 2, workbookId: 1, name: "Private", position: 2, visible: false, tags: [{ id: 20, tagTypeId: 2, name: "Archived" }] },
];
const entries = [
  { id: 1, workbookId: 1, vocabulary: "Travel", meaning: "to journey", meanings: ["to journey"], kanaText: null, attributes: { note: "Common phrase", internal: "secret" }, tags: [{ id: 10, tagTypeId: 1, name: "Verb" }], createdAt: "", updatedAt: "", testCount: 0, errorCount: 0, tier: "gray", lastTested: null, nextTestDeadline: null },
  { id: 2, workbookId: 1, vocabulary: "Home", meaning: "residence", meanings: ["residence"], kanaText: null, attributes: { note: "", internal: "hidden-value" }, tags: [{ id: 20, tagTypeId: 2, name: "Archived" }], createdAt: "", updatedAt: "", testCount: 0, errorCount: 0, tier: "green", lastTested: null, nextTestDeadline: null },
] satisfies EntryRow[];

test("search matches visible values and tags case-insensitively", () => {
  assert.deepEqual(filterEntriesForSearch(entries, workbook, tags, "TRAV").map((entry) => entry.id), [1]);
  assert.deepEqual(filterEntriesForSearch(entries, workbook, tags, "jour").map((entry) => entry.id), [1]);
  assert.deepEqual(filterEntriesForSearch(entries, workbook, tags, "phrase").map((entry) => entry.id), [1]);
  assert.deepEqual(filterEntriesForSearch(entries, workbook, tags, "verb").map((entry) => entry.id), [1]);
});

test("search excludes hidden fields and tags and preserves empty-list identity", () => {
  assert.deepEqual(filterEntriesForSearch(entries, workbook, tags, "hidden-value"), []);
  assert.deepEqual(filterEntriesForSearch(entries, workbook, tags, "archived"), []);
  assert.equal(filterEntriesForSearch(entries, workbook, tags, "   "), entries);
});
