import assert from "node:assert/strict";
import test from "node:test";
import type { EntryRow } from "./db.js";
import { changeListPage, moveListFocus, restoreListPosition, type VocabularyListPosition } from "./list-navigation.js";

const entries = Array.from({ length: 43 }, (_, index) => ({ id: (index + 1) * 3 }) as EntryRow);
const initial: VocabularyListPosition = { pageIndex: 0, selectedEntryId: entries[0].id, inputFocus: "command" };

test("arrows follow screen order through rows and inputs without wrapping", () => {
  assert.deepEqual(moveListFocus(entries, initial, 1), initial);
  let position = moveListFocus(entries, initial, -1);
  assert.equal(position.inputFocus, "search");
  position = moveListFocus(entries, position, -1);
  assert.equal(position.inputFocus, "table");
  assert.equal(position.selectedEntryId, entries[19].id);
  position = moveListFocus(entries, position, -1);
  assert.equal(position.selectedEntryId, entries[18].id);
  position = moveListFocus(entries, position, 1);
  assert.equal(position.selectedEntryId, entries[19].id);
  position = moveListFocus(entries, position, 1);
  assert.equal(position.inputFocus, "search");
  assert.equal(moveListFocus(entries, position, 1).inputFocus, "command");
  const firstRow = { ...initial, inputFocus: "table" as const };
  assert.deepEqual(moveListFocus(entries, firstRow, -1), firstRow);
});

test("empty lists skip table focus and page changes remain safe", () => {
  const search = moveListFocus([], initial, -1);
  assert.equal(search.inputFocus, "search");
  assert.equal(search.selectedEntryId, null);
  assert.deepEqual(moveListFocus([], search, -1), search);
  assert.equal(moveListFocus([], search, 1).inputFocus, "command");
  assert.deepEqual(changeListPage([], { ...initial, inputFocus: "table" }, 1), { pageIndex: 0, selectedEntryId: null, inputFocus: "search" });
});

test("paging preserves row position, clamps short pages, and uses result count", () => {
  const position = { ...initial, selectedEntryId: entries[18].id, inputFocus: "table" as const };
  const second = changeListPage(entries, position, 1);
  assert.equal(second.pageIndex, 1);
  assert.equal(second.selectedEntryId, entries[38].id);
  const last = changeListPage(entries, second, 1);
  assert.equal(last.pageIndex, 2);
  assert.equal(last.selectedEntryId, entries[42].id);
  assert.deepEqual(changeListPage(entries, last, 1), last);
  assert.equal(changeListPage(entries, last, -1).selectedEntryId, entries[22].id);
  const filtered = entries.filter((_, index) => index % 2 === 0);
  const filteredLast = changeListPage(filtered, { ...initial, selectedEntryId: filtered[18].id }, 20);
  assert.equal(filteredLast.pageIndex, 1);
  assert.equal(filteredLast.selectedEntryId, filtered[21].id);
});

test("restoration follows entry identity after reorder and falls back when it disappears", () => {
  const position = { pageIndex: 1, selectedEntryId: entries[25].id, inputFocus: "table" as const };
  const reordered = [entries[25], ...entries.filter((entry) => entry.id !== position.selectedEntryId)];
  assert.deepEqual(restoreListPosition(reordered, position), { ...position, pageIndex: 0 });
  const removed = entries.filter((entry) => entry.id !== position.selectedEntryId);
  assert.deepEqual(restoreListPosition(removed, position), { ...position, selectedEntryId: removed[20].id });
  assert.deepEqual(restoreListPosition(entries.slice(0, 3), position), { ...position, pageIndex: 0, selectedEntryId: entries[0].id });
});
