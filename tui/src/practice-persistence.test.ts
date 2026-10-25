import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { DatabaseSync } from "node:sqlite";
import test from "node:test";
import { VocabularyRepository } from "./db.js";
import { createOrchestrationController } from "./ink-orchestration.js";
import { createOrchestrationCapabilities } from "./orchestration-runner.js";
import { createInitialState } from "./orchestration.js";

test("session persistence schedules from final errors and counts all answers", () => {
  const repository = new VocabularyRepository(":memory:");
  try {
    const workbook = repository.createWorkbook("Intervals");
    const expectedErrors = [[0, 1, 3], [0, 2, 3], [1, 3, 3], [2, 3, 3]];
    for (let before = 0; before <= 3; before++) {
      for (const [index, answers] of [[true], [false, true], [false, false, false, true]].entries()) {
        const entry = repository.addEntry(workbook.id, `${before}-${index}`, "meaning");
        const seed = [...Array<boolean>(before).fill(false), true];
        repository.completePracticeSession(workbook.id, [{ entryId: entry.id, answers: seed }], "2026-10-01");
        const [saved] = repository.completePracticeSession(workbook.id, [{ entryId: entry.id, answers }], "2026-10-25");
        const expected = expectedErrors[before][index];
        assert.equal(saved.errorCount, expected);
        assert.equal(saved.testCount, seed.length + answers.length);
        assert.equal(saved.lastTested, "2026-10-25");
        assert.equal(saved.nextTestDeadline, ["2026-11-09", "2026-11-01", "2026-10-29", "2026-10-26"][expected]);
      }
    }
  } finally { repository.close(); }
});

test("real practice stages answers, discards cancellation, and saves on the completion day", async (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: new Date(2026, 9, 25, 23, 59) });
  const repository = new VocabularyRepository(":memory:");
  try {
    const workbook = repository.createWorkbook("Practice");
    const first = repository.addEntry(workbook.id, "first", "meaning");
    const second = repository.addEntry(workbook.id, "second", "meaning");
    const controller = createOrchestrationController(createInitialState(workbook), createOrchestrationCapabilities(repository), () => {});
    const dispatch = controller.dispatch;
    await dispatch({ type: "practiceStarted", candidates: [first, second] });
    await dispatch({ type: "practiceAnswerSubmitted", answer: "wrong" });
    await dispatch({ type: "practiceCancelled" });
    assert.deepEqual(repository.getEntry(first.id), first);
    await dispatch({ type: "practiceStarted", candidates: [first, second] });
    await dispatch({ type: "practiceAnswerSubmitted", answer: "wrong" });
    await dispatch({ type: "practiceDetailContinued" });
    await dispatch({ type: "practiceAnswerSubmitted", answer: "second" });
    await dispatch({ type: "practiceRetryAdvanced" });
    await dispatch({ type: "practiceAnswerSubmitted", answer: "wrong" });
    await dispatch({ type: "practiceDetailContinued" });
    await dispatch({ type: "practiceAnswerSubmitted", answer: "first" });
    assert.deepEqual(repository.getEntry(first.id), first);
    assert.deepEqual(repository.getEntry(second.id), second);
    t.mock.timers.tick(120_000);
    await dispatch({ type: "practiceRetryAdvanced" });
    assert.equal(controller.getState().practice.phase, "done");
    assert.equal(controller.getState().practice.score, 1);
    const savedFirst = repository.getEntry(first.id)!;
    const savedSecond = repository.getEntry(second.id)!;
    assert.equal(savedFirst.errorCount, 2);
    assert.equal(savedFirst.testCount, 3);
    assert.equal(savedFirst.lastTested, "2026-10-26");
    assert.equal(savedFirst.nextTestDeadline, "2026-10-30");
    assert.equal(savedSecond.errorCount, 0);
    assert.equal(savedSecond.testCount, 1);
    assert.equal(savedSecond.lastTested, "2026-10-26");
    assert.equal(savedSecond.nextTestDeadline, "2026-11-10");
    controller.dispose();
  } finally { repository.close(); }
});

test("failed session save rolls back every entry and retry writes exactly once", () => {
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-session-"));
  const path = join(directory, "test.db");
  const repository = new VocabularyRepository(path);
  const db = new DatabaseSync(path);
  try {
    const workbook = repository.createWorkbook("Atomic save");
    const first = repository.addEntry(workbook.id, "first", "meaning");
    const second = repository.addEntry(workbook.id, "second", "meaning");
    const results = [{ entryId: first.id, answers: [true] }, { entryId: second.id, answers: [false, true] }];
    db.exec(`CREATE TRIGGER reject_session BEFORE UPDATE ON entry_stats WHEN OLD.entry_id = ${second.id}
      BEGIN SELECT RAISE(ABORT, 'save failed'); END`);
    assert.throws(() => repository.completePracticeSession(workbook.id, results, "2026-10-25"), /save failed/);
    assert.deepEqual(repository.getEntry(first.id), first);
    assert.deepEqual(repository.getEntry(second.id), second);
    db.exec("DROP TRIGGER reject_session");
    const saved = repository.completePracticeSession(workbook.id, results, "2026-10-25");
    assert.deepEqual(saved.map((entry) => entry.testCount), [1, 2]);
    assert.deepEqual(saved.map((entry) => entry.nextTestDeadline), ["2026-11-09", "2026-11-01"]);
    const other = repository.createWorkbook("Other");
    const foreign = repository.addEntry(other.id, "foreign", "meaning");
    assert.throws(() => repository.completePracticeSession(workbook.id, [{ entryId: first.id, answers: [true] }, { entryId: foreign.id, answers: [true] }], "2026-10-25"), /belong/);
    assert.deepEqual(repository.getEntry(first.id), saved[0]);
    assert.throws(() => repository.completePracticeSession(workbook.id, [{ entryId: first.id, answers: [false] }], "2026-10-25"), /correct answer/);
    assert.throws(() => repository.completePracticeSession(workbook.id, results, "2026-02-30"), /valid completion day/);
  } finally { db.close(); repository.close(); rmSync(directory, { recursive: true, force: true }); }
});

test("selection compares calendar days, supports legacy timestamps, and keeps priorities", (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: new Date(2026, 9, 25, 0, 1) });
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-selection-"));
  const path = join(directory, "test.db");
  const repository = new VocabularyRepository(path);
  const db = new DatabaseSync(path);
  try {
    const workbook = repository.createWorkbook("Selection");
    const create = (name: string, errors: number, deadline: string | null, lastTested: string | null = "2026-10-01", testCount = 1) => {
      const entry = repository.addEntry(workbook.id, name, "meaning");
      db.prepare("UPDATE entry_stats SET test_count = ?, error_count = ?, last_tested = ?, next_test_deadline = ? WHERE entry_id = ?")
        .run(testCount, errors, lastTested, deadline, entry.id);
      return entry;
    };
    create("future-high", 3, "2026-10-26");
    create("due-today-low", 0, "2026-10-25");
    const legacyDeadline = new Date(2026, 9, 25, 23, 59, 59).toISOString();
    const legacy = create("due-today-legacy", 2, legacyDeadline);
    create("due-high", 3, "2026-10-24");
    create("future-low", 0, "2026-10-26");
    create("new", 0, null, null, 1);
    create("tested-zero-count", 1, "2026-10-24", "2026-10-01", 0);
    assert.deepEqual(repository.selectPracticeCandidates(workbook.id, 15).map((entry) => entry.vocabulary), [
      "new", "due-high", "due-today-legacy", "tested-zero-count", "due-today-low", "future-high", "future-low",
    ]);
    assert.equal(repository.selectPracticeCandidates(workbook.id, 3).length, 3);
    assert.deepEqual(repository.selectPracticeCandidates(workbook.id, 0), []);
    assert.equal(repository.getEntry(legacy.id)?.nextTestDeadline, legacyDeadline);
  } finally { db.close(); repository.close(); rmSync(directory, { recursive: true, force: true }); }
});
