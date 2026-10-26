import assert from "node:assert/strict";
import test from "node:test";
import type { EntryRow, WorkbookRow } from "./db.js";
import { createOrchestrationController, runTransitionEffects } from "./ink-orchestration.js";
import { createInitialState, update } from "./orchestration.js";
import type { Intent, OrchestrationCapabilities } from "./orchestration.js";
import { textActionFromRawInput } from "./text-input.js";

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

function session(count: number, sessionWorkbook = workbook) {
  const entries = Array.from({ length: count }, (_, i) => entry(i + 1));
  let saves = 0;
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
      completePracticeSession: (_workbookId, results, completedDay) => {
        saves++;
        return results.map(({ entryId: id, answers }) => {
          const index = entries.findIndex((item) => item.id === id);
          let current = entries[index];
          for (const [attempt, correct] of answers.entries()) {
            const decreaseError = attempt === 0;
            records.push({ id, correct, decreaseError });
            current = { ...current, testCount: current.testCount + 1,
              errorCount: correct ? Math.max(0, current.errorCount - (decreaseError ? 1 : 0)) : Math.min(3, current.errorCount + 1) };
          }
          entries[index] = { ...current, lastTested: completedDay };
          return entries[index];
        });
      },
    },
  };
  let state = createInitialState(sessionWorkbook);
  async function dispatch(intent: Intent) {
    const transition = update(state, intent);
    state = transition.state;
    const results = await runTransitionEffects(transition.effects, state, capabilities);
    for (const result of results) await dispatch(result);
  }
  return {
    get state() { return state; },
    records, entries, capabilities, dispatch,
    get saves() { return saves; },
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
  assert.equal(s.state.practice.currentEntry?.testCount, 0);
  assert.equal(s.records.length, 0);
  for (const id of [1, 3]) {
    assert.equal(s.state.practice.currentEntry?.id, id);
    await s.answer(`word${id}`);
    await s.continue();
  }
  assert.equal(s.state.practice.phase, "done");
  assert.equal(s.state.practice.score, 1);
  assert.deepEqual(s.records, [
    { id: 1, correct: false, decreaseError: true },
    { id: 1, correct: true, decreaseError: false },
    { id: 2, correct: true, decreaseError: true },
    { id: 3, correct: false, decreaseError: true },
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

for (const round of ["initial", "retry"] as const) {
  test(`correct ${round} answer stays visible until Enter advances`, async () => {
    const s = session(1);
    await s.start();
    if (round === "retry") {
      await s.answer("wrong");
      await s.continue();
    }
    const answer = "  word1  ";
    await s.answer(answer);
    assert.equal(s.state.practice.feedback, "Correct!");
    assert.equal(s.state.commandBuffer, answer);
    await s.dispatch({ type: "practiceAnswerChanged", answer: "edited" });
    assert.equal(s.state.commandBuffer, answer);
    await s.continue();
    assert.equal(s.state.commandBuffer, "");
    assert.equal(s.state.practice.phase, "done");
  });
}

test("empty candidates stay empty and ignore answer/advance intents", async () => {
  const s = session(0);
  await s.start();
  await s.answer("anything");
  await s.continue();
  assert.equal(s.state.practice.phase, "empty");
  assert.equal(s.records.length, 0);
});

test("answers stay staged and saved Review data stays unchanged until completion", async () => {
  const s = session(1);
  await s.start();
  await s.answer("wrong");
  assert.equal(s.state.practice.currentEntry?.errorCount, 2);
  assert.equal(s.entries[0].testCount, 0);
  assert.equal(s.saves, 0);
  await s.continue();
  await s.answer("word1");
  assert.equal(s.saves, 0);
  assert.deepEqual(s.state.practice.results, [{ entryId: 1, answers: [false, true] }]);
  await s.continue();
  assert.equal(s.saves, 1);
  assert.equal(s.entries[0].testCount, 2);
  assert.equal(s.entries[0].errorCount, 3);
});

test("cancellation discards staged answers and restarting starts fresh", async () => {
  const s = session(1);
  await s.start();
  await s.answer("wrong");
  await s.continue();
  await s.answer("word1");
  await s.dispatch({ type: "practiceCancelled" });
  assert.deepEqual(s.state.practice.results, []);
  assert.equal(s.records.length, 0);
  assert.equal(s.entries[0].testCount, 0);
  await s.start();
  await s.answer("word1");
  await s.continue();
  assert.equal(s.entries[0].testCount, 1);
  assert.equal(s.entries[0].errorCount, 1);
});

test("saving blocks duplicate submissions and advancement; stale and duplicate completions are ignored", async () => {
  const s = session(1);
  let state = update(createInitialState(workbook), { type: "practiceStarted", candidates: s.entries }).state;
  state = update(state, { type: "practiceAnswerSubmitted", answer: "word1" }).state;
  const saving = update(state, { type: "practiceRetryAdvanced" });
  assert.equal(saving.state.practice.phase, "saving");
  for (const intent of [
    { type: "practiceAnswerSubmitted", answer: "wrong" },
    { type: "practiceAnswerChanged", answer: "changed" },
    { type: "practiceRetryAdvanced" }, { type: "practiceDetailContinued" }, { type: "practiceSaveRetried" },
  ] satisfies Intent[]) {
    assert.equal(update(saving.state, intent).state, saving.state);
  }
  const [result] = await runTransitionEffects(saving.effects, saving.state, s.capabilities);
  const completed = update(saving.state, result).state;
  assert.equal(completed.practice.phase, "done");
  assert.equal(update(completed, result).state, completed);
  const cancelled = update(saving.state, { type: "practiceCancelled" }).state;
  assert.equal(update(cancelled, result).state, cancelled);
  const restarted = update(cancelled, { type: "practiceStarted", candidates: s.entries }).state;
  assert.equal(update(restarted, result).state, restarted);
  assert.equal(update(restarted, { type: "practiceSaveFailed", requestId: saving.state.practice.requestId, error: "old" }).state, restarted);
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

test("a save failure preserves results and completion day for retry after midnight", async (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: new Date(2026, 9, 25, 23, 59) });
  const s = session(1);
  const save = s.capabilities.practice.completePracticeSession;
  s.capabilities.practice.completePracticeSession = () => { throw new Error("Database unavailable"); };
  await s.start();
  await s.answer("word1");
  await s.continue();
  assert.equal(s.state.practice.phase, "saveFailed");
  assert.equal(s.state.practice.error, "Database unavailable");
  assert.equal(s.entries[0].testCount, 0);
  const completedDay = s.state.practice.completedDay;
  const results = s.state.practice.results;
  assert.equal(completedDay, "2026-10-25");
  t.mock.timers.tick(120_000);
  s.capabilities.practice.completePracticeSession = (id, savedResults, day) => {
    assert.equal(day, completedDay);
    assert.equal(savedResults, results);
    return save(id, savedResults, day);
  };
  await s.dispatch({ type: "practiceSaveRetried" });
  assert.equal(s.state.practice.phase, "done");
  assert.equal(s.saves, 1);
  assert.equal(s.entries[0].testCount, 1);
});

test("candidate failures become typed visible error state", async () => {
  const s = session(1);
  s.capabilities.practice.selectPracticeCandidates = () => { throw new Error("Could not select entries"); };
  await s.start();
  assert.equal(s.state.practice.phase, "empty");
  assert.equal(s.state.practice.error, "Could not select entries");
  assert.equal(s.records.length, 0);
});

test("the screen controller saves once for rapid repeated input", async () => {
  const s = session(1);
  const controller = createOrchestrationController(createInitialState(workbook), s.capabilities, () => {});
  await controller.dispatch({ type: "practiceStarted", candidates: s.entries });
  await controller.dispatch({ type: "practiceAnswerSubmitted", answer: "word1" });
  await controller.dispatch({ type: "practiceAnswerSubmitted", answer: "wrong" });
  assert.equal(controller.getState().practice.score, 1);
  assert.equal(s.saves, 0);
  const saving = controller.dispatch({ type: "practiceRetryAdvanced" });
  assert.equal(controller.getState().practice.phase, "saving");
  await controller.dispatch({ type: "practiceRetryAdvanced" });
  await saving;
  assert.equal(controller.getState().practice.phase, "done");
  assert.equal(s.saves, 1);
  controller.dispose();
});

test("practice loading and saving stay internal while results and failures remain visible", async () => {
  const s = session(1);
  const visible: string[] = [];
  const controller = createOrchestrationController(createInitialState(workbook), s.capabilities, (state) => {
    visible.push(state.practice.phase);
  });
  const loading = controller.dispatch({ type: "practiceCandidatesRequested", workbookId: 1, count: 1 });
  assert.equal(controller.getState().practice.phase, "loading");
  assert.equal(visible.length, 0);
  await loading;
  assert.deepEqual(visible, ["initial"]);
  await controller.dispatch({ type: "practiceAnswerSubmitted", answer: "word1" });
  s.capabilities.practice.completePracticeSession = () => { throw new Error("Save failed"); };
  const saving = controller.dispatch({ type: "practiceRetryAdvanced" });
  assert.equal(controller.getState().practice.phase, "saving");
  assert.equal(visible.at(-1), "initial");
  await saving;
  assert.equal(visible.at(-1), "saveFailed");
  s.capabilities.practice.completePracticeSession = () => s.entries;
  await controller.dispatch({ type: "practiceSaveRetried" });
  assert.equal(visible.at(-1), "done");
  assert.ok(!visible.includes("loading") && !visible.includes("saving"));
  controller.dispose();
});

test("disposed controllers suppress late state notifications and further effects", async () => {
  const s = session(1);
  let notifications = 0;
  const controller = createOrchestrationController(createInitialState(workbook), s.capabilities, () => { notifications++; });
  await controller.dispatch({ type: "practiceStarted", candidates: s.entries });
  await controller.dispatch({ type: "practiceAnswerSubmitted", answer: "word1" });
  const saving = controller.dispatch({ type: "practiceRetryAdvanced" });
  controller.dispose();
  const before = notifications;
  await saving;
  await controller.dispatch({ type: "practiceSaveRetried" });
  assert.equal(notifications, before);
  assert.equal(s.saves, 1);
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

const multipleMeaningsWorkbook: WorkbookRow = {
  ...workbook,
  meaningAttributes: ["English", "Japanese", "Chinese"].map((label, index) => ({ position: index + 1, label, languageCode: null })),
};

test("meaning cycling wraps in workbook order, includes blanks, and preserves the answer without effects", () => {
  const candidate = { ...entry(1), meanings: ["first", "", "third"] };
  let state = update(createInitialState(multipleMeaningsWorkbook), { type: "practiceStarted", candidates: [candidate] }).state;
  state = update(state, { type: "practiceAnswerChanged", answer: "partial answer" }).state;
  for (const expected of [1, 2, 0, 1, 2, 0]) {
    const previous = state;
    const transition = update(state, { type: "practiceMeaningCycled" });
    state = transition.state;
    assert.deepEqual(transition.effects, []);
    assert.deepEqual(state, { ...previous, practice: { ...previous.practice, meaningIndex: expected } });
    assert.equal(state.commandBuffer, "partial answer");
    assert.equal(state.workbook.meaningAttributes[expected].label, ["English", "Japanese", "Chinese"][expected]);
    assert.equal(state.practice.currentEntry?.meanings[expected], candidate.meanings[expected]);
  }
});

test("meaning cycling is ignored with one meaning or outside an unanswered question", () => {
  const single = update(createInitialState(workbook), { type: "practiceStarted", candidates: [entry(1)] }).state;
  const initial = update(createInitialState(multipleMeaningsWorkbook), { type: "practiceStarted", candidates: [entry(1)] }).state;
  const feedback = update(initial, { type: "practiceAnswerSubmitted", answer: "word1" }).state;
  for (const state of [single, feedback, ...(["loading", "empty", "detail", "saving", "saveFailed", "done"] as const).map((phase) => ({ ...initial, practice: { ...initial.practice, phase } })),
    { ...initial, practice: { ...initial.practice, feedback: "Correct!" } },
    { ...initial, practice: { ...initial.practice, currentEntry: null } },
  ]) {
    const transition = update(state, { type: "practiceMeaningCycled" });
    assert.equal(transition.state, state);
    assert.deepEqual(transition.effects, []);
  }
});

test("new questions and sessions reset the meaning, including correct, incorrect, and retry transitions", async () => {
  const s = session(3, multipleMeaningsWorkbook);
  await s.start();
  const cycle = async () => {
    await s.dispatch({ type: "practiceMeaningCycled" });
    assert.equal(s.state.practice.meaningIndex, 1);
  };
  await cycle();
  await s.answer("word1");
  await s.continue();
  assert.equal(s.state.practice.meaningIndex, 0);
  await cycle();
  await s.answer("wrong");
  await s.continue();
  assert.equal(s.state.practice.meaningIndex, 0);
  await cycle();
  await s.answer("wrong");
  await s.continue();
  assert.equal(s.state.practice.phase, "retry");
  assert.equal(s.state.practice.meaningIndex, 0);
  await cycle();
  await s.answer("wrong");
  await s.continue();
  assert.equal(s.state.practice.meaningIndex, 0);
  await cycle();
  await s.answer("word3");
  await s.continue();
  assert.equal(s.state.practice.retryNumber, 2);
  assert.equal(s.state.practice.meaningIndex, 0);
  await cycle();
  await s.answer("word2");
  await s.continue();
  assert.equal(s.state.practice.phase, "done");
  assert.equal(s.state.practice.score, 1);
  assert.equal(s.records.length, 6);
  await s.start();
  await cycle();
  await s.dispatch({ type: "practiceCancelled" });
  assert.equal(s.state.practice.meaningIndex, 0);
  await s.start();
  assert.equal(s.state.practice.meaningIndex, 0);
});

test("Tab and Shift+Tab never edit the practice answer or caret", () => {
  for (const input of ["\t", "\x1b[Z"]) {
    assert.equal(textActionFromRawInput(input), null);
  }
});
