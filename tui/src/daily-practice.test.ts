import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { DatabaseSync } from "node:sqlite";
import test from "node:test";
import { VocabularyBackend } from "./backend.js";
import { localCalendarDay } from "./calendar.js";
import { VocabularyRepository } from "./db.js";
import { createOrchestrationController } from "./ink-orchestration.js";
import { createOrchestrationCapabilities } from "./orchestration-runner.js";
import { createInitialState } from "./orchestration.js";
import { assertDatabaseIntegrity, runSchemaMigrations, SCHEMA_MIGRATIONS } from "./schema.js";
import { buildWorkbookMenuLines, parseDailyGoalInput } from "./workbook-display.js";

function fixture() {
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-daily-"));
  return { path: join(directory, "test.db"), cleanup: () => rmSync(directory, { recursive: true, force: true }) };
}

test("schema 5 upgrades without reconstructing daily history and remains idempotent", () => {
  const temp = fixture();
  const db = new DatabaseSync(temp.path);
  try {
    runSchemaMigrations(db, SCHEMA_MIGRATIONS.filter((migration) => migration.version <= 5));
    db.exec(`INSERT INTO workbooks (name, vocabulary_kind, vocabulary_label, created_at, updated_at)
      VALUES ('Existing', 'non_language', 'Vocabulary', '2026-01-01', '2026-01-01');
      INSERT INTO entries (workbook_id, vocabulary, created_at, updated_at) VALUES (1, 'word', '2026-01-01', '2026-01-01');
      INSERT INTO entry_stats (entry_id, test_count, error_count, last_tested, next_test_deadline)
      VALUES (1, 12, 2, '2026-10-25', '2026-10-29');`);
    const before = db.prepare("SELECT * FROM entry_stats").all();
    const repository = new VocabularyRepository(temp.path);
    try {
      repository.initialize();
      assert.equal(repository.getWorkbook(1)?.dailyGoal, null);
      assert.equal(repository.getWorkbook(1)?.todayPracticeCount, 0);
      assert.deepEqual(db.prepare("SELECT * FROM entry_stats").all(), before);
      assert.equal(db.prepare("SELECT COUNT(*) AS n FROM workbook_daily_practice").get()?.n, 0);
      assertDatabaseIntegrity(db);
    } finally { repository.close(); }
  } finally { db.close(); temp.cleanup(); }
});

test("goals validate, stay workbook-specific, persist, and clear through the backend", () => {
  const temp = fixture();
  let backend = new VocabularyBackend(temp.path);
  try {
    const first = backend.createWorkbook("First");
    const second = backend.createWorkbook("Second");
    assert.equal(first.dailyGoal, null);
    assert.equal(first.todayPracticeCount, 0);
    assert.equal(backend.setWorkbookDailyGoal(first.id, 15).dailyGoal, 15);
    for (const goal of [0, -1, 1.5, NaN, Infinity, Number.MAX_SAFE_INTEGER + 1]) {
      assert.throws(() => backend.setWorkbookDailyGoal(first.id, goal), /positive whole number/);
    }
    assert.equal(backend.getWorkbook(first.id)?.dailyGoal, 15);
    assert.equal(backend.getWorkbook(second.id)?.dailyGoal, null);
    assert.throws(() => backend.setWorkbookDailyGoal(9999, 15), /not found/);
    backend.close();
    backend = new VocabularyBackend(temp.path);
    assert.equal(backend.getWorkbook(first.id)?.dailyGoal, 15);
    assert.equal(backend.setWorkbookDailyGoal(first.id, 25).dailyGoal, 25);
    assert.equal(backend.setWorkbookDailyGoal(first.id, null).dailyGoal, null);
    backend.close();
    backend = new VocabularyBackend(temp.path);
    assert.equal(backend.getWorkbook(first.id)?.dailyGoal, null);
  } finally { backend.close(); temp.cleanup(); }
});

test("daily totals count session words, survive entry changes, and reset by local day", (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: new Date(2026, 9, 25, 12) });
  const temp = fixture();
  let repository = new VocabularyRepository(temp.path);
  const db = new DatabaseSync(temp.path);
  try {
    const workbook = repository.createWorkbook("Daily");
    const other = repository.createWorkbook("Other");
    const first = repository.addEntry(workbook.id, "first", "meaning");
    const second = repository.addEntry(workbook.id, "second", "meaning");
    repository.completePracticeSession(workbook.id, [{ entryId: first.id, answers: [false, false, true] }, { entryId: second.id, answers: [true] }], "2026-10-25");
    repository.completePracticeSession(workbook.id, [{ entryId: first.id, answers: [true] }], "2026-10-25");
    repository.completePracticeSession(workbook.id, [], "2026-10-25");
    assert.equal(repository.getWorkbook(workbook.id)?.todayPracticeCount, 3);
    assert.equal(repository.listWorkbooks()[0].todayPracticeCount, 3);
    assert.equal(repository.getWorkbook(other.id)?.todayPracticeCount, 0);
    repository.increasePriority(first.id);
    repository.decreasePriority(first.id);
    repository.updateEntry(first.id, "updated", "meaning");
    repository.importEntries(workbook.id, [{ vocabulary: "second", meanings: ["changed"], attributes: {}, tagIds: [] }]);
    repository.deleteEntry(first.id);
    repository.deleteEntry(second.id);
    assert.equal(repository.getWorkbook(workbook.id)?.todayPracticeCount, 3);
    assert.equal(repository.getWorkbook(workbook.id)?.wordCount, 0);
    repository.close();
    repository = new VocabularyRepository(temp.path);
    assert.equal(repository.getWorkbook(workbook.id)?.todayPracticeCount, 3);
    t.mock.timers.tick(24 * 60 * 60 * 1000);
    assert.equal(repository.getWorkbook(workbook.id)?.todayPracticeCount, 0);
    assert.equal(repository.listWorkbooks()[0].todayPracticeCount, 0);
    const next = repository.addEntry(workbook.id, "next", "meaning");
    repository.completePracticeSession(workbook.id, [{ entryId: next.id, answers: [true] }], localCalendarDay());
    assert.equal(repository.getWorkbook(workbook.id)?.todayPracticeCount, 1);
    assert.deepEqual(db.prepare("SELECT practice_count FROM workbook_daily_practice ORDER BY practice_day").all().map((row) => row.practice_count), [3, 1]);
    repository.deleteWorkbook(workbook.id);
    assert.equal(db.prepare("SELECT COUNT(*) AS n FROM workbook_daily_practice").get()?.n, 0);
    assertDatabaseIntegrity(db);
  } finally { db.close(); repository.close(); temp.cleanup(); }
});

test("cancelled practice adds nothing; sessions crossing midnight count on completion", async (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: new Date(2026, 9, 25, 23, 59) });
  const repository = new VocabularyRepository(":memory:");
  const workbook = repository.createWorkbook("Midnight");
  const entry = repository.addEntry(workbook.id, "word", "meaning");
  const controller = createOrchestrationController(createInitialState(workbook), createOrchestrationCapabilities(repository), () => {});
  try {
    await controller.dispatch({ type: "practiceStarted", candidates: [entry] });
    await controller.dispatch({ type: "practiceAnswerSubmitted", answer: "wrong" });
    await controller.dispatch({ type: "practiceCancelled" });
    assert.equal(repository.getWorkbook(workbook.id)?.todayPracticeCount, 0);
    await controller.dispatch({ type: "practiceStarted", candidates: [entry] });
    await controller.dispatch({ type: "practiceAnswerSubmitted", answer: "word" });
    t.mock.timers.tick(120_000);
    await controller.dispatch({ type: "practiceRetryAdvanced" });
    assert.equal(controller.getState().practice.phase, "done");
    assert.equal(repository.getWorkbook(workbook.id)?.todayPracticeCount, 1);
    t.mock.timers.setTime(new Date(2026, 9, 25, 23, 59).getTime());
    assert.equal(repository.getWorkbook(workbook.id)?.todayPracticeCount, 0);
  } finally { controller.dispose(); repository.close(); }
});

test("a failure writing daily totals rolls back entry stats and retry counts once", (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: new Date(2026, 9, 25, 12) });
  const temp = fixture();
  const repository = new VocabularyRepository(temp.path);
  const db = new DatabaseSync(temp.path);
  try {
    const workbook = repository.createWorkbook("Atomic");
    const entry = repository.addEntry(workbook.id, "word", "meaning");
    const results = [{ entryId: entry.id, answers: [false, true] }];
    db.exec("CREATE TRIGGER reject_daily BEFORE INSERT ON workbook_daily_practice BEGIN SELECT RAISE(ABORT, 'daily save failed'); END");
    assert.throws(() => repository.completePracticeSession(workbook.id, results, "2026-10-25"), /daily save failed/);
    assert.deepEqual(repository.getEntry(entry.id), entry);
    assert.equal(repository.getWorkbook(workbook.id)?.todayPracticeCount, 0);
    db.exec("DROP TRIGGER reject_daily");
    repository.completePracticeSession(workbook.id, results, "2026-10-25");
    assert.equal(repository.getWorkbook(workbook.id)?.todayPracticeCount, 1);
    assert.equal(repository.getEntry(entry.id)?.testCount, 2);
    assert.throws(() => repository.completePracticeSession(workbook.id, [{ entryId: entry.id, answers: [false] }], "2026-10-25"), /correct answer/);
    assert.equal(repository.getWorkbook(workbook.id)?.todayPracticeCount, 1);
  } finally { db.close(); repository.close(); temp.cleanup(); }
});

test("goal editor accepts decimal digits or blank and rejects invalid input", () => {
  assert.equal(parseDailyGoalInput(""), null);
  assert.equal(parseDailyGoalInput("   "), null);
  assert.equal(parseDailyGoalInput(" 15 "), 15);
  assert.equal(parseDailyGoalInput("001"), 1);
  assert.equal(parseDailyGoalInput(String(Number.MAX_SAFE_INTEGER)), Number.MAX_SAFE_INTEGER);
  for (const value of ["0", "-1", "1.5", "1e2", "0x10", "NaN", "Infinity", "words", "9007199254740992"]) {
    assert.throws(() => parseDailyGoalInput(value), /positive whole number/);
  }
});

test("home table keeps goals readable and excludes creation from workbook rows", () => {
  const repository = new VocabularyRepository(":memory:");
  try {
    const workbook = repository.createWorkbook("A very long workbook name that must be clipped");
    const workbooks = [
      { ...workbook, todayPracticeCount: 8, dailyGoal: null },
      { ...workbook, todayPracticeCount: 0, dailyGoal: 15 },
      { ...workbook, todayPracticeCount: 8, dailyGoal: 15 },
      { ...workbook, todayPracticeCount: 15, dailyGoal: 15 },
      { ...workbook, todayPracticeCount: 20, dailyGoal: 15 },
      { ...workbook, todayPracticeCount: 0, dailyGoal: null },
    ];
    const pad = (value: string, width: number) => value.slice(0, width).padEnd(width);
    for (const width of [56, 80, 120]) {
      const table = buildWorkbookMenuLines(workbooks, width, workbooks.length, pad);
      assert.match(table.header, /Today \/ Goal/);
      assert.ok(table.header.indexOf("Today / Goal") < table.header.indexOf("Words"));
      const colors = ["green", "#777777", "#e5c65f", "green", "green", "#777777"];
      const bars = ["          ", "░░░░░░░░░░", "█████░░░░░", "██████████", "██████████"];
      for (const [index, progress] of ["8 / No goal", "0 / 15", "8 / 15", "15 / 15", "20 / 15"].entries()) {
        assert.equal(table.rows[index].progress, ` ${bars[index]}  ${progress.padStart(16)} `);
        assert.equal(table.rows[index].selected, false);
      }
      for (const [index, color] of colors.entries()) {
        assert.equal(table.rows[index].progressColor, color);
        const selected = buildWorkbookMenuLines(workbooks, width, index, pad).rows[index];
        assert.equal(selected.selected, true);
        assert.equal(selected.progressColor, color);
        assert.equal(selected.progress, table.rows[index].progress);
      }
      assert.equal(table.rows.length, workbooks.length);
      assert.ok(table.rows.every((row) => !row.selected && !row.line.includes("+ Create new workbook")));
      for (const line of [table.border, table.header, ...table.rows.map((row) => row.line)]) assert.equal(line.length, width);
    }
    const empty = buildWorkbookMenuLines([], 80, 0, pad);
    assert.equal(empty.rows.length, 0);
    assert.equal(empty.border.length, 80);
    assert.equal(empty.header.length, 80);
    const ratios = [
      { today: 0, goal: 1, bar: "░░░░░░░░░░" },
      { today: 1, goal: 3, bar: "███░░░░░░░" },
      { today: 2, goal: 3, bar: "██████░░░░" },
      { today: 9, goal: 10, bar: "█████████░" },
      { today: 1, goal: 1, bar: "██████████" },
      { today: 12, goal: 10, bar: "██████████" },
      { today: 999, goal: 1000, bar: "█████████░" },
    ];
    const ratioTable = buildWorkbookMenuLines(ratios.map(({ today, goal }) => ({ ...workbook, todayPracticeCount: today, dailyGoal: goal })), 80, 0, pad);
    for (const [index, ratio] of ratios.entries()) {
      assert.equal(ratioTable.rows[index].progress.slice(1, 11), ratio.bar);
      assert.equal(ratioTable.rows[index].progress.length, 30);
      assert.equal(ratioTable.rows[index].progress.slice(13), `${`${ratio.today} / ${ratio.goal}`.padStart(16)} `);
    }
    const large = buildWorkbookMenuLines([{ ...workbook, dailyGoal: Number.MAX_SAFE_INTEGER, todayPracticeCount: Number.MAX_SAFE_INTEGER }], 120, 0, pad);
    assert.ok(large.rows[0].line.includes(`${Number.MAX_SAFE_INTEGER} / ${Number.MAX_SAFE_INTEGER}`));
    assert.equal(large.rows[0].line.length, 120);
  } finally { repository.close(); }
});
