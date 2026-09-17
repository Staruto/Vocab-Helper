import { EntryRow, TagType, WorkbookRow } from "./db.js";
import { visibleAssignedTagGroups } from "./tag-display.js";

export function searchableEntryValues(entry: EntryRow, workbook: WorkbookRow, tagTypes: TagType[]): string[] {
  const values = workbook.metadataAttributes.filter((attribute) => attribute.visible).map((attribute) => {
    if (attribute.key === "vocab") return entry.vocabulary;
    if (attribute.key.startsWith("meaning_")) return entry.meanings[Number(attribute.key.slice(8)) - 1] ?? "";
    return entry.attributes[attribute.key] ?? "";
  });
  return [...values, ...visibleAssignedTagGroups(entry, tagTypes).flatMap((group) => group.tagNames)];
}

export function filterEntriesForSearch(entries: EntryRow[], workbook: WorkbookRow, tagTypes: TagType[], query: string): EntryRow[] {
  const normalized = query.trim().toLocaleLowerCase();
  if (!normalized) return entries;
  return entries.filter((entry) => searchableEntryValues(entry, workbook, tagTypes).some((value) => value.toLocaleLowerCase().includes(normalized)));
}
