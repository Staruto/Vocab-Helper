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
  const progressWidth = Math.max(18, ...progressValues.map((value) => value.length + 2));
  const selectWidth = 3;
  const countWidth = 8;
  const nameWidth = Math.max(1, Math.max(width, 56) - selectWidth - countWidth - progressWidth - 5);
  const row = (selected: boolean, name: string, words: string, progress: string, progressColor: WorkbookMenuRow["progressColor"]): WorkbookMenuRow => {
    const beforeProgress = `|${pad(selected ? ">" : " ", selectWidth)}|${pad(name, nameWidth)}|`;
    const afterProgress = `|${pad(words, countWidth)}|`;
    return { selected, beforeProgress, progress, afterProgress, progressColor, line: beforeProgress + progress + afterProgress };
  };
  const border = `+${"-".repeat(selectWidth)}+${"-".repeat(nameWidth)}+${"-".repeat(progressWidth)}+${"-".repeat(countWidth)}+`;
  const rows = workbooks.map((workbook, index) => row(
    index === selectedIndex,
    workbook.name,
    String(workbook.wordCount),
    `${progressValues[index].padStart(progressWidth - 1)} `,
    workbook.todayPracticeCount === 0 ? "#777777"
      : workbook.dailyGoal !== null && workbook.todayPracticeCount < workbook.dailyGoal ? "#e5c65f" : "green",
  ));
  rows.push(row(selectedIndex === workbooks.length, "+ Create new workbook", "", " ".repeat(progressWidth), undefined));
  const headerProgress = pad(`${" ".repeat(Math.floor((progressWidth - "Today / Goal".length) / 2))}Today / Goal`, progressWidth);
  return { border, header: `|${pad("", selectWidth)}|${pad("Name", nameWidth)}|${headerProgress}|${pad("Words", countWidth)}|`, rows };
}
