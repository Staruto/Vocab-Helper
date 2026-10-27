import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { VocabularyBackend } from "./backend.js";
import { addCalendarDays } from "./calendar.js";
import type { DailyPracticeRow } from "./db.js";
import { buildPracticeProfile, practiceStreaks } from "./profile-display.js";

const history = (...days: string[]): DailyPracticeRow[] => days.map((day) => ({ day, practiceCount: 1 }));

test("daily history is ordered, workbook-specific, and persists through the backend", () => {
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-profile-"));
  const path = join(directory, "test.db");
  let backend = new VocabularyBackend(path);
  try {
    const first = backend.createWorkbook("First");
    const second = backend.createWorkbook("Second");
    const entry = backend.addEntry(first.id, "word", "meaning");
    assert.deepEqual(backend.listDailyPractice(first.id), []);
    const results = [{ entryId: entry.id, answers: [false, true] }];
    backend.completePracticeSession(first.id, results, "2026-10-27");
    backend.completePracticeSession(first.id, results, "2026-10-25");
    backend.completePracticeSession(first.id, results, "2026-10-27");
    const expected = [{ day: "2026-10-25", practiceCount: 1 }, { day: "2026-10-27", practiceCount: 2 }];
    assert.deepEqual(backend.listDailyPractice(first.id), expected);
    assert.deepEqual(backend.listDailyPractice(second.id), []);
    assert.throws(() => backend.listDailyPractice(99999), /not found/);
    backend.close();
    backend = new VocabularyBackend(path);
    assert.deepEqual(backend.listDailyPractice(first.id), expected);
  } finally { backend.close(); rmSync(directory, { recursive: true, force: true }); }
});

test("profile colors use the current goal, keep future days blank, and fill missing dates", () => {
  const days = [
    { day: "2026-10-19", practiceCount: 0 },
    { day: "2026-10-20", practiceCount: 1 },
    { day: "2026-10-21", practiceCount: 15 },
    { day: "2026-10-22", practiceCount: 20 },
    { day: "2026-10-28", practiceCount: 25 },
  ];
  const profile = buildPracticeProfile(days, 15, 80, "2026-10-27");
  const cells = new Map(profile.rows.flatMap((row) => row.cells).map((cell) => [cell.day, cell]));
  assert.deepEqual([cells.get("2026-10-19")!.symbol, cells.get("2026-10-19")!.color], ["·", "#777777"]);
  assert.deepEqual([cells.get("2026-10-20")!.symbol, cells.get("2026-10-20")!.color], ["█", "#e5c65f"]);
  assert.equal(cells.get("2026-10-21")!.color, "green");
  assert.equal(cells.get("2026-10-22")!.color, "green");
  assert.equal(cells.get("2026-10-23")!.symbol, "·");
  assert.equal(cells.get("2026-10-28")!.symbol, " ");
  const recolored = buildPracticeProfile(days, 30, 80, "2026-10-27");
  assert.equal(recolored.rows.flatMap((row) => row.cells).find((cell) => cell.day === "2026-10-21")!.color, "#e5c65f");
  const noGoal = buildPracticeProfile(days, null, 80, "2026-10-27");
  assert.equal(noGoal.rows.flatMap((row) => row.cells).find((cell) => cell.day === "2026-10-20")!.color, "green");
  assert.equal(noGoal.rows.flatMap((row) => row.cells).find((cell) => cell.day === "2026-10-19")!.color, "#777777");
});

test("calendar fits newest Monday-first weeks at terminal widths and labels months", () => {
  for (const width of [20, 56, 80, 120]) {
    const profile = buildPracticeProfile([], null, width, "2026-10-27");
    assert.equal(profile.weekCount, Math.min(53, Math.floor((width - 3) / 2)));
    assert.deepEqual(profile.rows.map((row) => row.label), ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]);
    assert.equal(profile.rows[0].cells.at(-1)!.day, "2026-10-26");
    assert.equal(profile.rows[6].cells.at(-1)!.day, "2026-11-01");
    assert.equal(profile.endDay, "2026-10-27");
    assert.ok(profile.monthLabels.includes("Oct"));
    assert.ok(profile.monthLabels.length <= width);
    for (const [dayIndex, row] of profile.rows.entries()) {
      const rendered = `${row.label} ${row.cells.map((cell) => cell.symbol).join(" ")}`;
      assert.ok(rendered.length <= width);
      assert.equal(row.cells.length, profile.weekCount);
      row.cells.forEach((cell, index) => assert.equal(cell.day, addCalendarDays(profile.startDay, index * 7 + dayIndex)));
    }
  }
  assert.equal(buildPracticeProfile([], null, 120, "2026-10-27").weekCount, 53);
  const leap = buildPracticeProfile([], null, 80, "2028-03-01");
  assert.equal(leap.rows[1].cells.at(-1)!.day, "2028-02-29");
  const year = buildPracticeProfile([], null, 80, "2027-01-01");
  assert.equal(year.rows[0].cells.at(-1)!.day, "2026-12-28");
  assert.equal(year.rows[4].cells.at(-1)!.day, "2027-01-01");
  const empty = buildPracticeProfile([], 15, 80, "2026-10-27");
  assert.deepEqual(empty.streaks, { current: 0, longest: 0 });
  assert.ok(empty.rows.flatMap((row) => row.cells).filter((cell) => cell.day <= "2026-10-27").every((cell) => cell.symbol === "·"));
});

test("streaks count any practice through today or yesterday and ignore gaps and future rows", () => {
  assert.deepEqual(practiceStreaks([], "2026-10-27"), { current: 0, longest: 0 });
  assert.deepEqual(practiceStreaks(history("2026-10-25", "2026-10-26", "2026-10-27"), "2026-10-27"), { current: 3, longest: 3 });
  assert.deepEqual(practiceStreaks(history("2026-10-25", "2026-10-26"), "2026-10-27"), { current: 2, longest: 2 });
  assert.deepEqual(practiceStreaks(history("2026-10-24", "2026-10-25"), "2026-10-27"), { current: 0, longest: 2 });
  assert.deepEqual(practiceStreaks([
    ...history("2026-10-27", "2026-10-25", "2026-10-28", "2026-10-29", "2026-10-30"),
    { day: "2026-10-26", practiceCount: 0 },
  ], "2026-10-27"), { current: 1, longest: 1 });
  assert.deepEqual(practiceStreaks(history("2028-02-28", "2028-02-29", "2028-03-01"), "2028-03-01"), { current: 3, longest: 3 });
  assert.deepEqual(practiceStreaks(history("2026-12-31", "2027-01-01"), "2027-01-01"), { current: 2, longest: 2 });
});

test("longest streak uses history outside the visible calendar", () => {
  const oldHistory = history("2024-01-01", "2024-01-02", "2024-01-03", "2024-01-04", "2026-10-27");
  const profile = buildPracticeProfile(oldHistory, 999, 56, "2026-10-27");
  assert.deepEqual(profile.streaks, { current: 1, longest: 4 });
  assert.ok(profile.rows.every((row) => row.cells.every((cell) => cell.day > "2024-01-04")));
});
