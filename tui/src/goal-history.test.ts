import assert from "node:assert/strict";
import test from "node:test";
import { DatabaseSync } from "node:sqlite";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { VocabularyRepository } from "./db.js";
import { assertDatabaseIntegrity, runSchemaMigrations, SCHEMA_MIGRATIONS } from "./schema.js";
import { buildPracticeProfile } from "./profile-display.js";

function fixture() {
  const directory = mkdtempSync(join(tmpdir(), "vocab-goal-history-"));
  return { path: join(directory, "test.db"), cleanup: () => rmSync(directory, { recursive: true, force: true }) };
}

test("a failed snapshot backfill rolls back schema and ledger changes", () => {
  const db = new DatabaseSync(":memory:");
  try {
    runSchemaMigrations(db, SCHEMA_MIGRATIONS.filter((migration) => migration.version <= 7));
    db.exec(`INSERT INTO workbooks (id,name,vocabulary_kind,vocabulary_label,daily_goal,created_at,updated_at)
      VALUES (1,'Daily','non_language','Vocabulary',15,'old','old');
      INSERT INTO workbook_daily_practice VALUES (1,'2026-10-25',5);
      CREATE TRIGGER reject_backfill BEFORE UPDATE ON workbook_daily_practice BEGIN SELECT RAISE(ABORT, 'backfill failed'); END;`);
    assert.throws(() => runSchemaMigrations(db), /migration 8 failed.*backfill failed/);
    assert.equal(db.prepare("SELECT 1 FROM pragma_table_info('workbook_daily_practice') WHERE name = 'goal_snapshot'").get(), undefined);
    assert.equal(db.prepare("SELECT 1 FROM schema_migrations WHERE version = 8").get(), undefined);
    assert.equal(db.prepare("SELECT practice_count FROM workbook_daily_practice").get()?.practice_count, 5);
    db.exec("DROP TRIGGER reject_backfill");
    runSchemaMigrations(db);
    assert.equal(db.prepare("SELECT goal_snapshot FROM workbook_daily_practice").get()?.goal_snapshot, 15);
  } finally { db.close(); }
});

test("schema 8 snapshots legacy goals once without changing counts or learning state", () => {
  const db = new DatabaseSync(":memory:");
  try {
    runSchemaMigrations(db, SCHEMA_MIGRATIONS.filter((migration) => migration.version <= 7));
    db.exec(`INSERT INTO workbooks (id, name, vocabulary_kind, vocabulary_label, daily_goal, created_at, updated_at) VALUES
      (1,'Goal 15','non_language','Vocabulary',15,'old','old'),
      (2,'No goal','non_language','Vocabulary',NULL,'old','old'),
      (3,'Goal 30','non_language','Vocabulary',30,'old','old');
      INSERT INTO workbook_daily_practice VALUES
      (1,'2026-10-25',5),(1,'2026-10-26',15),(1,'2026-10-27',20),
      (2,'2026-10-25',1),(2,'2026-10-26',0),(3,'2026-10-25',20);
      INSERT INTO entries (id,workbook_id,vocabulary,created_at,updated_at) VALUES (1,1,'word','old','old');
      INSERT INTO entry_stats (entry_id,test_count,error_count,last_tested,next_test_deadline) VALUES (1,10,1,'2026-10-27','2026-11-03');`);
    const counts = db.prepare("SELECT workbook_id, practice_day, practice_count FROM workbook_daily_practice ORDER BY workbook_id, practice_day").all();
    const stats = db.prepare("SELECT * FROM entry_stats").all();
    runSchemaMigrations(db);
    const snapshots = db.prepare("SELECT goal_snapshot FROM workbook_daily_practice ORDER BY workbook_id, practice_day").all().map((row) => row.goal_snapshot);
    assert.deepEqual(snapshots, [15, 15, 15, null, null, 30]);
    assert.deepEqual(db.prepare("SELECT workbook_id, practice_day, practice_count FROM workbook_daily_practice ORDER BY workbook_id, practice_day").all(), counts);
    assert.deepEqual(db.prepare("SELECT * FROM entry_stats").all(), stats);
    db.exec("UPDATE workbooks SET daily_goal = 999");
    runSchemaMigrations(db);
    assert.deepEqual(db.prepare("SELECT goal_snapshot FROM workbook_daily_practice ORDER BY workbook_id, practice_day").all().map((row) => row.goal_snapshot), snapshots);
    for (const value of [0, -1, 1.5, "'invalid'"]) assert.throws(() => db.exec(`UPDATE workbook_daily_practice SET goal_snapshot = ${value}`), /CHECK/);
    assertDatabaseIntegrity(db);
  } finally { db.close(); }
});

test("today follows goal edits while past snapshots survive clearing, midnight, sessions and restart", (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: new Date(2026, 9, 27, 12) });
  const temp = fixture();
  let repository = new VocabularyRepository(temp.path);
  const db = new DatabaseSync(temp.path);
  try {
    const workbook = repository.createWorkbook("Daily");
    const other = repository.createWorkbook("Other");
    const entry = repository.addEntry(workbook.id, "word", "meaning");
    const results = [{ entryId: entry.id, answers: [true] }];
    repository.setWorkbookDailyGoal(workbook.id, 2);
    assert.deepEqual(repository.listDailyPractice(workbook.id), []);
    repository.completePracticeSession(workbook.id, results, "2026-10-26");
    repository.completePracticeSession(workbook.id, results, "2026-10-27");
    const colors = () => new Map(buildPracticeProfile(repository.listDailyPractice(workbook.id), 80, "2026-10-28").rows.flatMap((row) => row.cells).map((cell) => [cell.day, cell.color]));
    assert.equal(colors().get("2026-10-26"), "#e5c65f");
    repository.setWorkbookDailyGoal(workbook.id, 1);
    assert.equal(colors().get("2026-10-27"), "green");
    assert.equal(colors().get("2026-10-26"), "#e5c65f");
    repository.setWorkbookDailyGoal(workbook.id, 5);
    assert.equal(colors().get("2026-10-27"), "#e5c65f");
    repository.setWorkbookDailyGoal(workbook.id, null);
    assert.equal(colors().get("2026-10-27"), "green");
    repository.setWorkbookDailyGoal(workbook.id, 2);
    repository.completePracticeSession(workbook.id, results, "2026-10-27");
    assert.equal(colors().get("2026-10-27"), "green");
    t.mock.timers.tick(24 * 60 * 60 * 1000);
    repository.setWorkbookDailyGoal(workbook.id, 10);
    repository.completePracticeSession(workbook.id, results, "2026-10-28");
    // A late save for yesterday must retain yesterday's goal.
    repository.completePracticeSession(workbook.id, results, "2026-10-27");
    assert.deepEqual(repository.listDailyPractice(workbook.id), [
      { day: "2026-10-26", practiceCount: 1, goalSnapshot: 2 },
      { day: "2026-10-27", practiceCount: 3, goalSnapshot: 2 },
      { day: "2026-10-28", practiceCount: 1, goalSnapshot: 10 },
    ]);
    assert.deepEqual(repository.listDailyPractice(other.id), []);
    repository.close(); repository = new VocabularyRepository(temp.path);
    assert.equal(repository.listDailyPractice(workbook.id)[1].goalSnapshot, 2);
    assert.equal(colors().get("2026-10-27"), "green");
    assert.equal(colors().get("2026-10-28"), "#e5c65f");
    assertDatabaseIntegrity(db);
  } finally { repository.close(); db.close(); temp.cleanup(); }
});

test("failed goal changes and session saves roll back snapshots, counts, and learning state", (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: new Date(2026, 9, 27, 12) });
  const temp = fixture();
  const repository = new VocabularyRepository(temp.path);
  const db = new DatabaseSync(temp.path);
  try {
    const workbook = repository.createWorkbook("Daily");
    repository.setWorkbookDailyGoal(workbook.id, 15);
    const entry = repository.addEntry(workbook.id, "word", "meaning");
    const results = [{ entryId: entry.id, answers: [true] }];
    repository.completePracticeSession(workbook.id, results, "2026-10-27");
    const before = repository.listDailyPractice(workbook.id);
    const practiced = repository.getEntry(entry.id);
    db.exec("CREATE TRIGGER reject_snapshot BEFORE UPDATE OF goal_snapshot ON workbook_daily_practice BEGIN SELECT RAISE(ABORT, 'snapshot failed'); END");
    assert.throws(() => repository.setWorkbookDailyGoal(workbook.id, 30), /snapshot failed/);
    assert.equal(repository.getWorkbook(workbook.id)?.dailyGoal, 15);
    assert.deepEqual(repository.listDailyPractice(workbook.id), before);
    db.exec("CREATE TRIGGER reject_count BEFORE UPDATE OF practice_count ON workbook_daily_practice BEGIN SELECT RAISE(ABORT, 'count failed'); END");
    assert.throws(() => repository.completePracticeSession(workbook.id, results, "2026-10-27"), /count failed/);
    assert.deepEqual(repository.listDailyPractice(workbook.id), before);
    assert.deepEqual(repository.getEntry(entry.id), practiced);
    db.exec("DROP TRIGGER reject_count; DROP TRIGGER reject_snapshot");
    repository.completePracticeSession(workbook.id, results, "2026-10-27");
    assert.deepEqual(repository.listDailyPractice(workbook.id), [{ day: "2026-10-27", practiceCount: 2, goalSnapshot: 15 }]);
    assertDatabaseIntegrity(db);
  } finally { repository.close(); db.close(); temp.cleanup(); }
});
