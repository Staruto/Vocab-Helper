import type { WorkbookRow } from "./db.js";

export function parseDailyGoalInput(input: string): number | null {
  const value = input.trim();
  if (!value) return null;
  const goal = Number(value);
  if (!/^\d+$/.test(value) || !Number.isSafeInteger(goal) || goal <= 0) {
    throw new Error("Daily goal must be a positive whole number, or empty for no goal.");
  }
  return goal;
}

export function buildWorkbookMenuLines(
  workbooks: WorkbookRow[],
  width: number,
  selectedIndex: number,
  pad: (text: string, width: number) => string,
): { border: string; header: string; rows: Array<{ line: string; selected: boolean }> } {
  const progress = workbooks.map((workbook) => `${workbook.todayPracticeCount} / ${workbook.dailyGoal ?? "No goal"}`);
  const progressWidth = Math.max("Today / Goal".length, ...progress.map((value) => value.length));
  const selectWidth = 3;
  const countWidth = 8;
  const nameWidth = Math.max(1, Math.max(width, 56) - selectWidth - countWidth - progressWidth - 5);
  const line = (marker: string, name: string, words: string, today: string) =>
    `|${pad(marker, selectWidth)}|${pad(name, nameWidth)}|${pad(words, countWidth)}|${pad(today, progressWidth)}|`;
  const border = `+${"-".repeat(selectWidth)}+${"-".repeat(nameWidth)}+${"-".repeat(countWidth)}+${"-".repeat(progressWidth)}+`;
  const rows = workbooks.map((workbook, index) => ({
    selected: index === selectedIndex,
    line: line(index === selectedIndex ? ">" : " ", workbook.name, String(workbook.wordCount), progress[index]),
  }));
  rows.push({ selected: selectedIndex === workbooks.length, line: line(selectedIndex === workbooks.length ? ">" : " ", "+ Create new workbook", "", "") });
  return { border, header: line("", "Name", "Words", "Today / Goal"), rows };
}
