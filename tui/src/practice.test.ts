import assert from "node:assert/strict";
import test from "node:test";
import type { EntryRow, WorkbookRow } from "./db.js";
import { createOrchestrationController, runTransitionEffects } from "./ink-orchestration.js";
import { createInitialState, update } from "./orchestration.js";
import type { Intent, OrchestrationCapabilities } from "./orchestration.js";

const workbook: WorkbookRow = {
  id: 1, name: "Practice", wordCount: 0, createdAt: "",
  vocabularyLabel: "Vocabulary", vocabularyLanguageCode: null,
  presetEnabled: false, vocabularyKind: "non_language", importFilePath: null,
  meaningAttributes: [{ position: 1, label: "Meaning", languageCode: null }],
  metadataAttributes: [],
};
function entry(id: number): EntryRow {
  return {
    id, workbookId: 1, vocabulary: `word${id}`, meaning: `meaning${id}`,
    meanings: [`meaning${id}`], attributes: {}, tags: [], kanaText: null,
    createdAt: "", updatedAt: "", testCount: 0, errorCount: 2,
    tier: "yellow", lastTested: null, nextTestDeadline: null,
  };
}
const unexpected = (): never => { throw new Error("Unexpected capability call"); };

function session(count: number) {
  const entries = Array.from({ length: count }, (_, i) => entry(i + 1));
  const records: Array<{ id: number; correct: boolean; decreaseError: boolean }> = [];
  const capabilities: OrchestrationCapabilities = {
    reads: { listEntries: unexpected, getEntry: unexpected, listTagTypes: unexpected, getWorkbook: unexpected },
    writes: { addEntry: unexpected, updateEntry: unexpected, deleteEntry: unexpected },
    import: { readLabeledTextFile: unexpected, importEntries: unexpected, readWorkbookImportPath: unexpected, saveWorkbookImportPath: unexpected },
    practice: {
      selectPracticeCandidates: (workbookId, requested) => {
        assert.equal(workbookId, workbook.id);
        return entries.slice(0, requested);
      },
      recordTestResult: (id, correct, decreaseError = true) => {
        records.push({ id, correct, decreaseError });
        const index = entries.findIndex((item) => item.id === id);
        const previous = entries[index];
        const errorCount = correct
          ? Math.max(0, previous.errorCount - (decreaseError ? 1 : 0))
          : Math.min(3, previous.errorCount + 1);
        entries[index] = { ...previous, testCount: previous.testCount + 1, errorCount };
        return entries[index];
      },
    },
  };
  let state = createInitialState(workbook);
  async function dispatch(intent: Intent) {
    const transition = update(state, intent);
    state = transition.state;
    const results = await runTransitionEffects(transition.effects, state, capabilities);
    for (const result of results) await dispatch(result);
  }
  return {
    get state() { return state; },
    records, entries, capabilities, dispatch,
    async start() { await dispatch({ type: "practiceCandidatesRequested", workbookId: 1, count }); },
    async answer(value: string) {
      await dispatch({ type: "practiceAnswerSubmitted", answer: value });
    },
    async continue() {
      await dispatch({ type: state.practice.phase === "detail" ? "practiceDetailContinued" : "practiceRetryAdvanced" });
    },
  };
}

for (const count of [1, 3]) {
  test(`all ${count} correct initial answers finish without any retry`, async () => {
    const s = session(count);
    await s.start();
    for (let i = 1; i <= count; i++) {
      assert.equal(s.state.practice.currentEntry?.id, i);
      await s.answer(`word${i}`);
      assert.equal(s.state.practice.feedback, "Correct!");
      assert.deepEqual(s.state.practice.nextRetryRound, []);
      await s.continue();
    }
    assert.equal(s.state.practice.phase, "done");
    assert.equal(s.state.practice.score, count);
    assert.equal(s.state.practice.currentEntry, null);
    assert.deepEqual(s.state.practice.nextRetryRound, []);
    assert.deepEqual(s.state.practice.retryRound, []);
    assert.equal(s.records.length, count);
    assert.ok(s.records.every((record) => record.correct && record.decreaseError));
  });
}

test("mixed initial answers retry only mistakes in their original order", async () => {
  const s = session(3);
  await s.start();
  for (const answer of ["wrong", "word2", "wrong"]) {
    await s.answer(answer);
    await s.continue();
  }
  assert.equal(s.state.practice.score, 1);
  assert.equal(s.state.practice.phase, "retry");
  assert.deepEqual(s.state.practice.retryRound.map((item) => item.id), [1, 3]);
  assert.equal(s.state.practice.currentEntry?.testCount, 1);
  for (const id of [1, 3]) {
    assert.equal(s.state.practice.currentEntry?.id, id);
    await s.answer(`word${id}`);
    await s.continue();
  }
  assert.equal(s.state.practice.phase, "done");
  assert.equal(s.state.practice.score, 1);
  assert.deepEqual(s.records, [
    { id: 1, correct: false, decreaseError: true },
    { id: 2, correct: true, decreaseError: true },
    { id: 3, correct: false, decreaseError: true },
    { id: 1, correct: true, decreaseError: false },
    { id: 3, correct: true, decreaseError: false },
  ]);
  assert.deepEqual(s.entries.map((item) => item.errorCount), [3, 1, 3]);
});

test("a mistake on the final initial question starts a retry that terminates", async () => {
  const s = session(2);
  await s.start();
  await s.answer("word1");
  await s.continue();
  await s.answer("wrong");
  assert.equal(s.state.practice.phase, "detail");
  assert.equal(s.state.practice.currentEntry?.id, 2);
  await s.continue();
  assert.deepEqual(s.state.practice.retryRound.map((item) => item.id), [2]);
  assert.equal(s.state.practice.retryNumber, 1);
  await s.answer("word2");
  await s.continue();
  assert.equal(s.state.practice.phase, "done");
  assert.equal(s.state.practice.score, 1);
  assert.equal(s.records.length, 3);
});

test("later retry rounds contain only the preceding round's mistakes", async () => {
  const s = session(2);
  await s.start();
  for (let i = 0; i < 2; i++) { await s.answer("wrong"); await s.continue(); }
  await s.answer("word1");
  await s.continue();
  await s.answer("wrong");
  await s.continue();
  assert.equal(s.state.practice.retryNumber, 2);
  assert.deepEqual(s.state.practice.retryRound.map((item) => item.id), [2]);
  await s.answer("word2");
  await s.continue();
  assert.equal(s.state.practice.phase, "done");
  assert.equal(s.state.practice.score, 0);
  assert.equal(s.records.length, 5);
});

test("answers are exact after trimming surrounding whitespace", async () => {
  const s = session(2);
  await s.start();
  await s.answer(" \tword1\n");
  assert.equal(s.state.practice.feedback, "Correct!");
  await s.continue();
  await s.answer("Word2");
  assert.equal(s.state.practice.phase, "detail");
  assert.equal(s.state.practice.score, 1);
});

test("empty candidates stay empty and ignore answer/advance intents", async () => {
  const s = session(0);
  await s.start();
  await s.answer("anything");
  await s.continue();
  assert.equal(s.state.practice.phase, "empty");
  assert.equal(s.records.length, 0);
});

test("pending answers block submissions, edits, and both advancement intents", () => {
  const initial = update(createInitialState(workbook), { type: "practiceStarted", candidates: [entry(1)] }).state;
  const pending = update(initial, { type: "practiceAnswerSubmitted", answer: "word1" }).state;
  assert.notEqual(pending.practice.pendingAnswer, null);
  assert.equal(pending.practice.score, 0);
  assert.equal(pending.practice.feedback, null);
  for (const intent of [
    { type: "practiceAnswerSubmitted", answer: "wrong" },
    { type: "practiceAnswerChanged", answer: "changed" },
    { type: "practiceRetryAdvanced" },
    { type: "practiceDetailContinued" },
  ] satisfies Intent[]) {
    const transition = update(pending, intent);
    assert.equal(transition.state, pending);
    assert.deepEqual(transition.effects, []);
  }
});

test("duplicate and stale recording results cannot score twice or replace another question", async () => {
  const s = session(2);
  const initial = update(createInitialState(workbook), { type: "practiceStarted", candidates: s.entries }).state;
  const submitted = update(initial, { type: "practiceAnswerSubmitted", answer: "word1" });
  const [result] = await runTransitionEffects(submitted.effects, submitted.state, s.capabilities);
  const recorded = update(submitted.state, result).state;
  assert.equal(recorded.practice.score, 1);
  assert.equal(update(recorded, result).state, recorded);
  assert.equal(update(recorded, { type: "practiceDetailContinued" }).state, recorded);
  const next = update(recorded, { type: "practiceRetryAdvanced" }).state;
  assert.equal(next.practice.currentEntry?.id, 2);
  assert.equal(update(next, { type: "practiceRetryAdvanced" }).state, next);
  const nextPending = update(next, { type: "practiceAnswerSubmitted", answer: "word2" }).state;
  assert.equal(update(nextPending, result).state, nextPending);
  assert.equal(s.records.length, 1);
});

test("cancellation and replacement invalidate in-flight results and failures", async () => {
  const s = session(1);
  const initial = update(createInitialState(workbook), { type: "practiceStarted", candidates: s.entries }).state;
  const submitted = update(initial, { type: "practiceAnswerSubmitted", answer: "word1" });
  const [result] = await runTransitionEffects(submitted.effects, submitted.state, s.capabilities);
  const cancelled = update(submitted.state, { type: "practiceCancelled" }).state;
  assert.equal(update(cancelled, result).state, cancelled);
  const restarted = update(cancelled, { type: "practiceStarted", candidates: [entry(1)] }).state;
  const newPending = update(restarted, { type: "practiceAnswerSubmitted", answer: "word1" }).state;
  assert.equal(update(newPending, result).state, newPending);
  assert.equal(update(newPending, { type: "practiceResultFailed", requestId: submitted.state.practice.requestId, error: "old error" }).state, newPending);
  assert.equal(newPending.practice.score, 0);
});

test("candidate loading ignores older results, including after cancellation", () => {
  const intent = { type: "practiceCandidatesRequested", workbookId: 1, count: 1 } as const;
  const first = update(createInitialState(workbook), intent).state;
  const second = update(first, intent).state;
  const result: Intent = { type: "practiceStarted", candidates: [entry(1)], requestId: first.practice.requestId };
  assert.equal(update(second, result).state, second);
  assert.equal(update(second, { type: "practiceCandidatesFailed", requestId: first.practice.requestId, error: "old" }).state, second);
  const cancelled = update(first, { type: "practiceCancelled" }).state;
  assert.equal(update(cancelled, result).state, cancelled);
});

test("a recording failure retains the answer and question and allows resubmission", async () => {
  const s = session(1);
  await s.start();
  const record = s.capabilities.practice.recordTestResult;
  s.capabilities.practice.recordTestResult = () => { throw new Error("Database unavailable"); };
  await s.answer("word1");
  assert.equal(s.state.practice.error, "Database unavailable");
  assert.equal(s.state.practice.pendingAnswer, null);
  assert.equal(s.state.practice.phase, "initial");
  assert.equal(s.state.practice.score, 0);
  assert.equal(s.state.commandBuffer, "word1");
  await s.continue();
  assert.equal(s.state.practice.currentEntry?.id, 1);
  s.capabilities.practice.recordTestResult = record;
  await s.answer("word1");
  await s.continue();
  assert.equal(s.state.practice.phase, "done");
  assert.equal(s.state.practice.score, 1);
  assert.equal(s.records.length, 1);
});

test("candidate failures become typed visible error state", async () => {
  const s = session(1);
  s.capabilities.practice.selectPracticeCandidates = () => { throw new Error("Could not select entries"); };
  await s.start();
  assert.equal(s.state.practice.phase, "empty");
  assert.equal(s.state.practice.error, "Could not select entries");
  assert.equal(s.records.length, 0);
});

test("the screen controller records once for rapid repeated input", async () => {
  const s = session(1);
  const controller = createOrchestrationController(createInitialState(workbook), s.capabilities, () => {});
  await controller.dispatch({ type: "practiceCandidatesRequested", workbookId: 1, count: 1 });
  await controller.dispatch({ type: "practiceAnswerChanged", answer: "word1" });
  const recording = controller.dispatch({ type: "practiceAnswerSubmitted", answer: controller.getState().commandBuffer });
  assert.notEqual(controller.getState().practice.pendingAnswer, null);
  await controller.dispatch({ type: "practiceAnswerSubmitted", answer: "word1" });
  await recording;
  assert.equal(s.records.length, 1);
  assert.equal(controller.getState().practice.score, 1);
  await controller.dispatch({ type: "practiceRetryAdvanced" });
  assert.equal(controller.getState().practice.phase, "done");
  await controller.dispatch({ type: "practiceAnswerSubmitted", answer: "word1" });
  assert.equal(s.records.length, 1);
  controller.dispose();
});

test("disposed controllers suppress late state notifications and further effects", async () => {
  const s = session(1);
  let notifications = 0;
  const controller = createOrchestrationController(createInitialState(workbook), s.capabilities, () => { notifications++; });
  await controller.dispatch({ type: "practiceStarted", candidates: s.entries });
  const recording = controller.dispatch({ type: "practiceAnswerSubmitted", answer: "word1" });
  controller.dispose();
  const before = notifications;
  await recording;
  await controller.dispatch({ type: "practiceAnswerSubmitted", answer: "word1" });
  assert.equal(notifications, before);
  assert.equal(s.records.length, 1);
});

test("bounded sessions terminate after every combination of initial and first-retry mistakes", async () => {
  for (let initialMask = 0; initialMask < 8; initialMask++) {
    for (let retryMask = 0; retryMask < 8; retryMask++) {
      const s = session(3);
      await s.start();
      const initialMistakes: number[] = [];
      for (let id = 1; id <= 3; id++) {
        const wrong = Boolean(initialMask & (1 << (id - 1)));
        if (wrong) initialMistakes.push(id);
        await s.answer(wrong ? "wrong" : `word${id}`);
        await s.continue();
      }
      assert.deepEqual(s.state.practice.retryRound.map((item) => item.id), initialMistakes);
      const remaining: number[] = [];
      for (const id of initialMistakes) {
        assert.equal(s.state.practice.currentEntry?.id, id);
        const wrong = Boolean(retryMask & (1 << (id - 1)));
        if (wrong) remaining.push(id);
        await s.answer(wrong ? "wrong" : `word${id}`);
        await s.continue();
      }
      assert.deepEqual(s.state.practice.retryRound.map((item) => item.id), remaining);
      for (const id of remaining) {
        assert.equal(s.state.practice.currentEntry?.id, id);
        await s.answer(`word${id}`);
        await s.continue();
      }
      assert.equal(s.state.practice.phase, "done", `initial mask ${initialMask}, retry mask ${retryMask}`);
      assert.equal(s.state.practice.score, 3 - initialMistakes.length);
      assert.equal(s.records.length, 3 + initialMistakes.length + remaining.length);
    }
  }
});
