import type { WorkbookRow } from "./db.js";

type WorkbookMenuRow = {
  line: string;
  selected: boolean;
  beforeProgress: string;
  progress: string;
  afterProgress: string;
  progressColor: "#777777" | "#e5c65f" | "green" | undefined;
};

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
): { border: string; header: string; rows: WorkbookMenuRow[] } {
  const progressValues = workbooks.map((workbook) => `${workbook.todayPracticeCount} / ${workbook.dailyGoal ?? "No goal"}`);
  const valueWidth = Math.max(16, ...progressValues.map((value) => value.length));
  const progressWidth = valueWidth + 14;
  const selectWidth = 3;
  const countWidth = 8;
  const nameWidth = Math.max(1, Math.max(width, 56) - selectWidth - countWidth - progressWidth - 5);
  const row = (selected: boolean, name: string, words: string, progress: string, progressColor: WorkbookMenuRow["progressColor"]): WorkbookMenuRow => {
    const beforeProgress = `|${pad(selected ? ">" : " ", selectWidth)}|${pad(name, nameWidth)}|`;
    const afterProgress = `|${pad(words, countWidth)}|`;
    return { selected, beforeProgress, progress, afterProgress, progressColor, line: beforeProgress + progress + afterProgress };
  };
  const border = `+${"─".repeat(selectWidth)}+${"─".repeat(nameWidth)}+${"─".repeat(progressWidth)}+${"─".repeat(countWidth)}+`;
  const rows = workbooks.map((workbook, index) => {
    const filled = workbook.dailyGoal === null ? 0 : Math.floor(Math.min(workbook.todayPracticeCount / workbook.dailyGoal, 1) * 10);
    const bar = workbook.dailyGoal === null ? " ".repeat(10) : "■".repeat(filled) + "□".repeat(10 - filled);
    return row(
      index === selectedIndex,
      workbook.name,
      String(workbook.wordCount),
      ` ${bar}  ${progressValues[index].padStart(valueWidth)} `,
      workbook.todayPracticeCount === 0 ? "#777777"
        : workbook.dailyGoal !== null && workbook.todayPracticeCount < workbook.dailyGoal ? "#e5c65f" : "green",
    );
  });
  const headerProgress = pad(`${" ".repeat(Math.floor((progressWidth - "Today / Goal".length) / 2))}Today / Goal`, progressWidth);
  return { border, header: `|${pad("", selectWidth)}|${pad("Name", nameWidth)}|${headerProgress}|${pad("Words", countWidth)}|`, rows };
}
