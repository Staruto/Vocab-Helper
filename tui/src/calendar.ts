/** Local calendar dates for review scheduling; date-only values never shift timezone. */
export function localCalendarDay(date = new Date()): string {
  return `${String(date.getFullYear()).padStart(4, "0")}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

export function calendarDay(value: string): string {
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value.slice(0, 10) : localCalendarDay(date);
}

export function addCalendarDays(day: string, days: number): string {
  const date = new Date(`${day}T00:00:00.000Z`);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

export function reviewDeadline(lastTested: string, errorCount: number): string {
  const day = calendarDay(lastTested);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(day) || Number.isNaN(Date.parse(lastTested)) || addCalendarDays(day, 0) !== day) {
    throw new Error("Invalid last-tested date.");
  }
  return addCalendarDays(day, [15, 7, 4, 1][errorCount]);
}
