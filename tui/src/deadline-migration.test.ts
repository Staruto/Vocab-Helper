import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { DatabaseSync } from "node:sqlite";
import test from "node:test";
import { VocabularyRepository } from "./db.js";
import { buildReviewFields } from "./detail-display.js";
import { assertDatabaseIntegrity, runSchemaMigrations, SCHEMA_MIGRATIONS } from "./schema.js";

type Stats = {
  entry_id: number; test_count: number; error_count: number;
  last_tested: string | null; next_test_deadline: string | null;
};

function legacyDatabase() {
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-deadlines-"));
  const path = join(directory, "vocab.db");
  const db = new DatabaseSync(path);
  runSchemaMigrations(db, SCHEMA_MIGRATIONS.filter((migration) => migration.version <= 4));
  for (const name of ["First", "Second"]) {
    const result = db.prepare(`INSERT INTO workbooks
      (name, vocabulary_kind, vocabulary_label, preset_enabled, created_at, updated_at)
      VALUES (?, 'non_language', 'Vocabulary', 0, '2026-01-01', '2026-01-01')`).run(name);
    db.prepare(`INSERT INTO workbook_fields (workbook_id, field_key, role, position, label, is_required)
      VALUES (?, 'meaning_1', 'meaning', 1, 'Meaning', 1)`).run(Number(result.lastInsertRowid));
  }
  let sequence = 0;
  function add(workbookId: number, errors: number, lastTested: string | null, deadline: string | null, count = 9): number {
    const result = db.prepare(`INSERT INTO entries (workbook_id, vocabulary, created_at, updated_at)
      VALUES (?, ?, '2026-01-01', '2026-01-01')`).run(workbookId, `word-${++sequence}`);
    const id = Number(result.lastInsertRowid);
    db.prepare("INSERT INTO entry_stats (entry_id, test_count, error_count, last_tested, next_test_deadline) VALUES (?, ?, ?, ?, ?)")
      .run(id, count, errors, lastTested, deadline);
    return id;
  }
  const stats = () => (db.prepare("SELECT * FROM entry_stats ORDER BY entry_id").all() as Stats[]).map((row) => ({ ...row }));
  return { path, db, add, stats, cleanup() { db.close(); rmSync(directory, { recursive: true, force: true }); } };
}

test("startup repairs only missing deadlines across workbooks and runs once", (t) => {
  t.mock.timers.enable({ apis: ["Date"], now: new Date(2026, 9, 26, 12) });
  const fixture = legacyDatabase();
  let repository: VocabularyRepository | undefined;
  try {
    const expected = new Map<number, string>();
    const deadlines = ["2026-10-16", "2026-10-08", "2026-10-05", "2026-10-02"];
    for (const lastTested of ["2026-10-01", new Date(2026, 9, 1, 0, 30).toISOString()]) {
      for (let errors = 0; errors <= 3; errors++) {
        expected.set(fixture.add(errors % 2 + 1, errors, lastTested, null), deadlines[errors]);
      }
    }
    const freshId = fixture.add(1, 0, null, null, 0);
    const scheduledId = fixture.add(1, 2, "2026-10-01T00:00:00.000Z", "2999-01-01T00:00:00.000Z");
    fixture.add(2, 0, "invalid-but-already-scheduled", "2026-11-01");
    const before = fixture.stats();
    const expectedStats = before.map((row) => ({ ...row, next_test_deadline: expected.get(row.entry_id) ?? row.next_test_deadline }));

    repository = new VocabularyRepository(fixture.path);
    assert.deepEqual(fixture.stats(), expectedStats);
    assert.equal((fixture.db.prepare("SELECT COUNT(*) AS count FROM schema_migrations WHERE version = 5").get() as { count: number }).count, 1);
    assertDatabaseIntegrity(fixture.db);
    const entry = repository.getEntry([...expected.keys()][0])!;
    assert.equal(buildReviewFields(entry).at(-1)?.value, "2026-10-16");
    assert.equal(repository.selectPracticeCandidates(1, 1)[0].id, freshId);
    const testedSelection = repository.selectPracticeCandidates(1, 3);
    assert.ok(testedSelection.slice(1).every((item) => expected.has(item.id)));
    assert.ok(testedSelection.every((item) => item.id !== scheduledId));
    repository.close();
    repository = new VocabularyRepository(fixture.path);
    repository.initialize();
    assert.deepEqual(fixture.stats(), expectedStats);
    assert.equal((fixture.db.prepare("SELECT COUNT(*) AS count FROM schema_migrations WHERE version = 5").get() as { count: number }).count, 1);
  } finally { repository?.close(); fixture.cleanup(); }
});

for (const invalidDate of ["not-a-date", "2026-02-30"]) {
  test(`invalid Last tested (${invalidDate}) rolls back repairs and migration ledger`, () => {
    const fixture = legacyDatabase();
    try {
      fixture.add(1, 0, "2026-10-01", null);
      const invalidId = fixture.add(2, 3, invalidDate, null);
      const before = fixture.stats();
      assert.throws(() => new VocabularyRepository(fixture.path), new RegExp(`migration 5 failed: Cannot repair deadline for entry #${invalidId}: invalid last-tested date`));
      assert.deepEqual(fixture.stats(), before);
      assert.equal((fixture.db.prepare("SELECT COUNT(*) AS count FROM schema_migrations WHERE version = 5").get() as { count: number }).count, 0);
      fixture.db.prepare("UPDATE entry_stats SET last_tested = '2026-10-01' WHERE entry_id = ?").run(invalidId);
      const repository = new VocabularyRepository(fixture.path);
      try {
        assert.deepEqual(fixture.stats().map((row) => row.next_test_deadline), ["2026-10-16", "2026-10-02"]);
        assertDatabaseIntegrity(fixture.db);
      } finally { repository.close(); }
    } finally { fixture.cleanup(); }
  });
}
