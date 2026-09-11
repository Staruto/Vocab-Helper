import assert from "node:assert/strict";
import test from "node:test";
import { adjacentEntryId, buildDetailSections } from "./detail-display.js";
import { EntryRow, WorkbookRow } from "./db.js";

const workbook = {
  id: 1,
  vocabularyLabel: "Word",
  meaningAttributes: [{ key: "meaning_1", label: "Meaning", role: "meaning", position: 1, languageCode: null, required: true, visible: true }],
  metadataAttributes: [{ key: "note", label: "Note", role: "optional", position: 1, languageCode: null, required: false, visible: true }],
} as unknown as WorkbookRow;
const entry = { id: 2, workbookId: 1, meanings: ["to travel"], attributes: { note: "" }, tags: [{ id: 10, tagTypeId: 1, name: "verb" }] } as EntryRow;

test("detail sections preserve fields, use an em dash for empty optionals, and filter tags by visibility", () => {
  const sections = buildDetailSections(workbook, entry, [
    { id: 1, workbookId: 1, name: "Part of Speech", position: 1, visible: true, tags: [{ id: 10, tagTypeId: 1, name: "verb" }] },
    { id: 2, workbookId: 1, name: "Hidden", position: 2, visible: false, tags: [{ id: 11, tagTypeId: 2, name: "internal" }] },
  ]);
  assert.deepEqual(sections.meanings, [{ key: "Meaning", value: "to travel" }]);
  assert.deepEqual(sections.attributes, [{ key: "Note", value: "—" }]);
  assert.deepEqual(sections.tags, [{ typeId: 1, typeName: "Part of Speech", tagNames: ["verb"] }]);
});

test("adjacent entry lookup follows list order and stops at boundaries", () => {
  const entries = [{ id: 10 }, { id: 20 }, { id: 30 }] as EntryRow[];
  assert.equal(adjacentEntryId(entries, 20, "previous"), 10);
  assert.equal(adjacentEntryId(entries, 20, "next"), 30);
  assert.equal(adjacentEntryId(entries, 10, "previous"), null);
  assert.equal(adjacentEntryId(entries, 30, "next"), null);
});

