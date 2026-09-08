import { EntryRow, TagType } from "./db.js";

export type VisibleTagGroup = { typeId: number; typeName: string; tagNames: string[] };

export function visibleAssignedTagGroups(entry: EntryRow, tagTypes: TagType[]): VisibleTagGroup[] {
  const assignedIds = new Set(entry.tags.map((tag) => tag.id));
  return tagTypes.flatMap((type) => {
    if (!type.visible) return [];
    const tagNames = type.tags.filter((tag) => assignedIds.has(tag.id)).map((tag) => tag.name);
    return tagNames.length ? [{ typeId: type.id, typeName: type.name, tagNames }] : [];
  });
}

export function fitTagBadges(tagNames: string[], maxWidth: number, measure: (text: string) => number): string[] {
  const badges: string[] = [];
  let used = 0;
  for (const name of tagNames) {
    const badge = `[${name}]`;
    const nextWidth = measure(badge) + (badges.length ? 1 : 0);
    if (used + nextWidth > maxWidth) continue;
    badges.push(badge);
    used += nextWidth;
  }
  return badges;
}
