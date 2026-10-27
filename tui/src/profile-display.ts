import { addCalendarDays, localCalendarDay } from "./calendar.js";
import type { DailyPracticeRow } from "./db.js";

export type PracticeDayCell = {
  day: string;
  practiceCount: number;
  symbol: "·" | "█" | " ";
  color: "#777777" | "#e5c65f" | "green";
};

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

export function practiceStreaks(history: DailyPracticeRow[], today = localCalendarDay()): { current: number; longest: number } {
  const days = [...new Set(history.filter((row) => row.practiceCount > 0 && row.day <= today).map((row) => row.day))].sort();
  let longest = 0;
  let run = 0;
  let previous: string | undefined;
  for (const day of days) {
    run = previous !== undefined && addCalendarDays(previous, 1) === day ? run + 1 : 1;
    longest = Math.max(longest, run);
    previous = day;
  }
  const active = new Set(days);
  let day = active.has(today) ? today : addCalendarDays(today, -1);
  let current = 0;
  while (active.has(day)) {
    current++;
    day = addCalendarDays(day, -1);
  }
  return { current, longest };
}

export function buildPracticeProfile(history: DailyPracticeRow[], goal: number | null, width: number, today = localCalendarDay()) {
  const weekday = (new Date(`${today}T00:00:00.000Z`).getUTCDay() + 6) % 7;
  const currentMonday = addCalendarDays(today, -weekday);
  // Four characters for weekday labels, then two columns per week (no trailing gap).
  const weekCount = Math.min(53, Math.max(1, Math.floor((width - 3) / 2)));
  const startDay = addCalendarDays(currentMonday, -(weekCount - 1) * 7);
  const counts = new Map(history.map((row) => [row.day, row.practiceCount]));
  const rows = WEEKDAYS.map((label, dayIndex) => ({
    label,
    cells: Array.from({ length: weekCount }, (_, week): PracticeDayCell => {
      const day = addCalendarDays(startDay, week * 7 + dayIndex);
      const practiceCount = counts.get(day) ?? 0;
      return {
        day, practiceCount,
        symbol: day > today ? " " : practiceCount === 0 ? "·" : "█",
        color: practiceCount === 0 ? "#777777" : goal !== null && practiceCount < goal ? "#e5c65f" : "green",
      };
    }),
  }));
  const monthChars = Array<string>(weekCount * 2 - 1).fill(" ");
  let previousMonth = "";
  let nextLabelPosition = 0;
  for (let week = 0; week < weekCount; week++) {
    const monday = addCalendarDays(startDay, week * 7);
    const month = monday.slice(0, 7);
    if (month !== previousMonth) {
      const position = week * 2;
      const label = MONTHS[Number(monday.slice(5, 7)) - 1];
      if (position >= nextLabelPosition && position + label.length <= monthChars.length) {
        [...label].forEach((char, index) => { monthChars[position + index] = char; });
        nextLabelPosition = position + label.length + 1;
      }
      previousMonth = month;
    }
  }
  return { startDay, endDay: today, weekCount, monthLabels: `    ${monthChars.join("")}`, rows, streaks: practiceStreaks(history, today) };
}
