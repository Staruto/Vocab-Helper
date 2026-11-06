export const SETTINGS_SECTIONS = ["Attributes", "Tags", "Import", "Appearance", "Daily Goal"] as const;
export type SettingsSection = typeof SETTINGS_SECTIONS[number];

/** Keep the heading fixed and fit a contiguous window of whole tabs. */
export function settingsTabWindow(width: number, selected: number): number[] {
  const available = Math.max(0, width - "Settings   ".length);
  let first = selected;
  let last = selected;
  let used = SETTINGS_SECTIONS[selected].length + 2;
  while (first > 0 && used + SETTINGS_SECTIONS[first - 1].length + 3 <= available) {
    used += SETTINGS_SECTIONS[--first].length + 3;
  }
  while (last + 1 < SETTINGS_SECTIONS.length && used + SETTINGS_SECTIONS[last + 1].length + 3 <= available) {
    used += SETTINGS_SECTIONS[++last].length + 3;
  }
  return Array.from({ length: last - first + 1 }, (_, index) => first + index);
}
