import assert from "node:assert/strict";
import { mkdtempSync, unlinkSync, rmdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { DatabaseSync } from "node:sqlite";
import test from "node:test";
import { VocabularyRepository } from "./db.js";
import { LIST_ORDER_OPTIONS, parseListOrder } from "./list-order.js";
import { runSchemaMigrations, SCHEMA_MIGRATIONS } from "./schema.js";
import { createInitialState, update } from "./orchestration.js";
import { createOrchestrationCapabilities, executeEffect } from "./orchestration-runner.js";

test("list parameters accept precisely the four orders, ignoring case", () => {
  for (const option of LIST_ORDER_OPTIONS) assert.equal(parseListOrder(option.parameter.toUpperCase().split(" ")), option.value);
  for (const args of [[], ["add"], ["asc"], ["error", "desc", "extra"], ["id", "asc"], ["error", "wrong"]]) assert.equal(parseListOrder(args), null);
});

test("list order upgrades from schema 9, sorts deterministically, and persists per workbook", () => {
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-list-order-"));
  const path = join(directory, "vocab.db");
  const db = new DatabaseSync(path);
  runSchemaMigrations(db, SCHEMA_MIGRATIONS.filter((migration) => migration.version <= 9));
  db.prepare(`INSERT INTO workbooks (id, name, vocabulary_kind, vocabulary_label, created_at, updated_at)
    VALUES (1, 'Legacy', 'non_language', 'Vocabulary', '', '')`).run();
  db.prepare(`INSERT INTO workbook_fields (workbook_id, field_key, role, position, label, is_required, is_visible)
    VALUES (1, 'meaning_1', 'meaning', 1, 'Meaning', 1, 1)`).run();
  db.close();
  const repository = new VocabularyRepository(path);
  try {
    assert.equal(repository.getWorkbook(1)!.listOrder, "add_asc");
    const other = repository.createWorkbook("Other");
    assert.equal(other.listOrder, "add_asc");
    assert.deepEqual(repository.listEntries(other.id), []);
    const entries = Array.from({ length: 7 }, (_, index) => repository.addEntry(1, `word${index}`, "meaning"));
    // IDs: unpracticed, error 2, error 0, error 3, error 2, unpracticed, error 1.
    const stats = new DatabaseSync(path);
    for (const [index, count] of [[1, 2], [2, 0], [3, 3], [4, 2], [6, 1]]) {
      stats.prepare("UPDATE entry_stats SET test_count = 5, error_count = ? WHERE entry_id = ?").run(count, entries[index].id);
    }
    stats.close();
    const ids = (indices: number[]) => indices.map((index) => entries[index].id);
    const expected = {
      add_asc: [0, 1, 2, 3, 4, 5, 6], add_desc: [6, 5, 4, 3, 2, 1, 0],
      error_asc: [2, 6, 1, 4, 3, 0, 5], error_desc: [0, 5, 3, 1, 4, 6, 2],
    };
    for (const { value } of LIST_ORDER_OPTIONS) {
      repository.setWorkbookListOrder(1, value);
      assert.deepEqual(repository.listEntries(1).map((entry) => entry.id), ids(expected[value]));
      assert.equal(repository.getWorkbook(other.id)!.listOrder, "add_asc");
    }
    repository.initialize();
    repository.close();
    const reopened = new VocabularyRepository(path);
    try {
      assert.equal(reopened.getWorkbook(1)!.listOrder, "error_desc");
      assert.deepEqual(reopened.listEntries(1).map((entry) => entry.id), ids(expected.error_desc));
    } finally { reopened.close(); }
  } finally {
    repository.close(); unlinkSync(path); rmdirSync(directory);
  }
});

test("picker cancels safely, wraps, refreshes current order, and handles save failures", async () => {
  const repository = new VocabularyRepository(":memory:");
  try {
    const workbook = repository.createWorkbook("Picker");
    repository.addEntry(workbook.id, "first", "meaning");
    repository.addEntry(workbook.id, "second", "meaning");
    const capabilities = createOrchestrationCapabilities(repository);
    const initial = { ...createInitialState(workbook, repository.listEntries(workbook.id)), activeSearchQuery: "first", pageIndex: 3 };
    let transition = update(initial, { type: "commandSubmitted", raw: "/list" });
    assert.deepEqual(transition.effects, []);
    assert.deepEqual(transition.state.mode, { kind: "listOrder", choice: 0 });
    const cancelled = update(transition.state, { type: "modeCancelled" }).state;
    assert.equal(cancelled.workbook.listOrder, "add_asc");
    assert.equal(cancelled.pageIndex, 3);
    let state = update(transition.state, { type: "listOrderMoved", delta: -1 }).state;
    assert.deepEqual(state.mode, { kind: "listOrder", choice: 3 });
    state = update(state, { type: "listOrderMoved", delta: 1 }).state;
    transition = update(state, { type: "listOrderConfirmed" });
    assert.deepEqual(transition.effects, [{ type: "saveListOrder", workbookId: workbook.id, order: "add_asc" }]);
    assert.deepEqual(update(transition.state, { type: "listOrderConfirmed" }).effects, []);
    const failed = await executeEffect(transition.effects[0], transition.state, { ...capabilities, list: { saveOrder() { throw new Error("Disk full"); } } });
    state = update(transition.state, failed).state;
    assert.deepEqual(state.mode, { kind: "listOrder", choice: 0, saving: false, error: "Disk full" });
    assert.equal(state.workbook.listOrder, "add_asc");
    transition = update(state, { type: "listOrderConfirmed" });
    const result = await executeEffect(transition.effects[0], transition.state, capabilities);
    state = update(transition.state, result).state;
    assert.equal(state.mode.kind, "command");
    assert.equal(state.pageIndex, 0);
    assert.equal(state.activeSearchQuery, "first");
    assert.equal(state.entries.length, 2);
    for (const option of LIST_ORDER_OPTIONS) {
      transition = update(state, { type: "commandSubmitted", raw: `/LIST ${option.parameter.toUpperCase()}` });
      assert.deepEqual(transition.effects, [{ type: "saveListOrder", workbookId: workbook.id, order: option.value }]);
      state = update(transition.state, await executeEffect(transition.effects[0], transition.state, capabilities)).state;
      assert.equal(state.workbook.listOrder, option.value);
      assert.deepEqual(update(state, { type: "commandSubmitted", raw: "/list" }).state.mode, { kind: "listOrder", choice: LIST_ORDER_OPTIONS.indexOf(option) });
    }
    for (const raw of ["/list add", "/list error desc extra", "/list invalid"]) {
      const invalid = update(state, { type: "commandSubmitted", raw });
      assert.deepEqual(invalid.effects, []);
      assert.equal(invalid.state.workbook.listOrder, state.workbook.listOrder);
      assert.match(invalid.state.statusLines[0], /Usage: \/list/);
    }
  } finally { repository.close(); }
});
