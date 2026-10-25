import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import test from "node:test";
import { addCalendarDays, calendarDay, localCalendarDay } from "./calendar.js";

test("calendar intervals cross months, years, and leap days", () => {
  assert.equal(addCalendarDays("2026-12-31", 1), "2027-01-01");
  assert.equal(addCalendarDays("2028-02-28", 1), "2028-02-29");
  assert.equal(addCalendarDays("2028-02-28", 4), "2028-03-03");
  assert.equal(addCalendarDays("2026-02-28", 1), "2026-03-01");
  assert.equal(addCalendarDays("2026-10-25", 15), "2026-11-09");
  assert.equal(calendarDay("2026-10-25"), "2026-10-25");
  assert.equal(localCalendarDay(new Date(2026, 9, 25, 23, 59)), "2026-10-25");
});

for (const timezone of ["Asia/Tokyo", "America/New_York"]) {
  test(`calendar dates use ${timezone} and intervals survive daylight-saving changes`, () => {
    const source = `
      import assert from "node:assert/strict";
      import { addCalendarDays, calendarDay, localCalendarDay } from "./src/calendar.ts";
      assert.equal(localCalendarDay(new Date(2026, 9, 25, 23, 59)), "2026-10-25");
      assert.equal(calendarDay("2026-10-25"), "2026-10-25");
      assert.equal(calendarDay("2026-10-25T01:00:00.000Z"), ${JSON.stringify(timezone === "Asia/Tokyo" ? "2026-10-25" : "2026-10-24")});
      assert.equal(addCalendarDays("2026-03-08", 1), "2026-03-09");
      assert.equal(addCalendarDays("2026-11-01", 1), "2026-11-02");
    `;
    const result = spawnSync(process.execPath, ["--import", "tsx", "--input-type=module", "-e", source], {
      cwd: new URL("../", import.meta.url), encoding: "utf8", env: { ...process.env, TZ: timezone },
    });
    assert.equal(result.status, 0, result.stderr);
  });
}
