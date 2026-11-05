import { EntryRow, TagType, WorkbookRow } from "./db.js";
import { entryAttributeValue, isActiveAttribute } from "./attributes.js";
import { visibleAssignedTagGroups } from "./tag-display.js";

export function searchableEntryValues(entry: EntryRow, workbook: WorkbookRow, tagTypes: TagType[]): string[] {
  const values = workbook.metadataAttributes.filter((attribute) => attribute.visible && isActiveAttribute(attribute))
    .map((attribute) => entryAttributeValue(entry, attribute));
  return [...values, ...visibleAssignedTagGroups(entry, tagTypes).flatMap((group) => group.tagNames)];
}

export function filterEntriesForSearch(entries: EntryRow[], workbook: WorkbookRow, tagTypes: TagType[], query: string): EntryRow[] {
  const normalized = query.trim().toLocaleLowerCase();
  if (!normalized) return entries;
  return entries.filter((entry) => searchableEntryValues(entry, workbook, tagTypes).some((value) => value.toLocaleLowerCase().includes(normalized)));
}
