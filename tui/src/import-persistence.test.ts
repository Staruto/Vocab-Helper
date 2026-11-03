import assert from "node:assert/strict";
import test from "node:test";
import { resolve } from "node:path";
import { VocabularyRepository } from "./db.js";
import { createOrchestrationController } from "./ink-orchestration.js";
import { createOrchestrationCapabilities } from "./orchestration-runner.js";
import { createInitialState } from "./orchestration.js";
import { parseLabeledTextImport } from "./import.js";

for (const choice of [0, 1, 2]) {
  test(`mixed import choice ${choice} applies once to all updates`, async () => {
    const repository = new VocabularyRepository(":memory:");
    try {
      const workbook = repository.createWorkbook("Import");
      const first = repository.addEntry(workbook.id, "first", "old");
      const second = repository.addEntry(workbook.id, "second", "old");
      const unchanged = repository.addEntry(workbook.id, "unchanged", "same");
      repository.completePracticeSession(workbook.id, [first, second, unchanged].map(({ id }) => ({ entryId: id, answers: [false, true] })), "2026-10-25");
      const before = repository.listEntries(workbook.id);
      const history = repository.listDailyPractice(workbook.id);
      const text = ["first:new", "second:new", "unchanged:same", "added:new"].map((value) => {
        const [word, meaning] = value.split(":");
        return `${workbook.vocabularyLabel}: ${word}\n${workbook.meaningAttributes[0].label}: ${meaning}`;
      }).join("\n\n");
      const preview = parseLabeledTextImport(text, workbook, [], before.map(({ vocabulary }) => vocabulary));
      const controller = createOrchestrationController({ ...createInitialState(workbook, before), mode: {
        kind: "importPreview", path: resolve("words.txt"), pathBuffer: "draft.txt", preview, filter: "duplicates", pageIndex: 2, focus: "records", error: undefined,
      } }, createOrchestrationCapabilities(repository), () => {});
      const previous = controller.getState().mode;
      await controller.dispatch({ type: "importConfirmed" });
      assert.equal(controller.getState().mode.kind, "importConfirm");
      const confirm = controller.getState().mode;
      assert.ok(confirm.kind === "importConfirm");
      assert.equal(confirm.updated, 2);
      assert.equal(confirm.choice, 0);
      assert.deepEqual(repository.listEntries(workbook.id), before);
      await controller.dispatch({ type: "importPreviewReturned" });
      assert.deepEqual(controller.getState().mode, previous);
      await controller.dispatch({ type: "importConfirmed" });
      await controller.dispatch({ type: "importChoiceMoved", delta: 3 });
      await controller.dispatch({ type: "importChoiceConfirmed" });
      assert.deepEqual(controller.getState().mode, previous);
      await controller.dispatch({ type: "importConfirmed" });
      await controller.dispatch({ type: "importChoiceMoved", delta: choice });
      await controller.dispatch({ type: "importChoiceConfirmed" });
      if (choice === 2) {
        assert.equal(controller.getState().mode.kind, "command");
        assert.deepEqual(repository.listEntries(workbook.id), before);
      } else {
        assert.equal(controller.getState().mode.kind, "importSaveDefault");
        for (const { id } of [first, second]) {
          assert.equal(repository.getEntry(id)!.meaning, "new");
          assert.deepEqual(repository.getEntryStats(id), choice === 0 ? repository.getEntryStats(unchanged.id) : {
            testCount: 0, errorCount: 0, lastTested: null, nextTestDeadline: null, tier: "gray",
          });
        }
        assert.deepEqual(repository.getEntry(unchanged.id), before[2]);
        assert.equal(repository.countEntries(workbook.id), 4);
        await controller.dispatch({ type: "importSavePathDecision", save: true });
        assert.equal(controller.getState().entries.length, 4);
        assert.equal(repository.getWorkbook(workbook.id)!.importFilePath, resolve("words.txt"));
      }
      assert.deepEqual(repository.listDailyPractice(workbook.id), history);
    } finally { repository.close(); }
  });
}

test("update check normalizes content, simulates repeated records, and never writes", () => {
  const repository = new VocabularyRepository(":memory:");
  try {
    const workbook = repository.createWorkbook("Check");
    const entry = repository.addEntry(workbook.id, "word", "meaning");
    const same = { vocabulary: " word ", meanings: [" meaning "], attributes: {}, tagIds: [] };
    const changed = { ...same, meanings: ["changed"] };
    assert.equal(repository.previewImportUpdates(workbook.id, [same, same]), 0);
    assert.equal(repository.previewImportUpdates(workbook.id, [changed, same, changed]), 1);
    assert.equal(repository.previewImportUpdates(workbook.id, [
      { ...same, vocabulary: "new" }, { ...changed, vocabulary: "new" },
    ]), 0);
    assert.deepEqual(repository.listEntries(workbook.id), [entry]);
  } finally { repository.close(); }
});

test("unchanged imports bypass confirmation and preserve timestamps", async () => {
  const repository = new VocabularyRepository(":memory:");
  try {
    const created = repository.createWorkbook("Unchanged");
    const workbook = repository.setWorkbookImportFilePath(created.id, "words.txt");
    const entry = repository.addEntry(workbook.id, "word", "meaning");
    const preview = parseLabeledTextImport(`${workbook.vocabularyLabel}: word\n${workbook.meaningAttributes[0].label}: meaning`, workbook, [], ["word"]);
    const controller = createOrchestrationController({ ...createInitialState(workbook), mode: {
      kind: "importPreview", path: resolve("words.txt"), pathBuffer: "words.txt", preview, filter: "records", pageIndex: 0, focus: "filters",
    } }, createOrchestrationCapabilities(repository), () => {});
    await controller.dispatch({ type: "importConfirmed" });
    assert.equal(controller.getState().mode.kind, "command");
    assert.deepEqual(controller.getState().entries, [entry]);
  } finally { repository.close(); }
});
