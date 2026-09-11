import { EntryRow, TagType, WorkbookRow } from "./db.js";
import { visibleAssignedTagGroups, VisibleTagGroup } from "./tag-display.js";

export type DetailField = { key: string; value: string };

export type DetailSections = {
  meanings: DetailField[];
  attributes: DetailField[];
  tags: VisibleTagGroup[];
};

export function buildDetailSections(workbook: WorkbookRow, entry: EntryRow, tagTypes: TagType[]): DetailSections {
  const meanings = entry.meanings.map((value, index) => ({
    key: workbook.meaningAttributes[index]?.label ?? `Meaning ${index + 1}`,
    value: value.trim() || "—",
  }));
  const attributes = workbook.metadataAttributes
    .filter((attribute) => attribute.role === "optional")
    .map((attribute) => ({
      key: attribute.label,
      value: entry.attributes[attribute.key]?.trim() || "—",
    }));
  return { meanings, attributes, tags: visibleAssignedTagGroups(entry, tagTypes) };
}

export function adjacentEntryId(entries: EntryRow[], currentId: number, direction: "previous" | "next"): number | null {
  const index = entries.findIndex((entry) => entry.id === currentId);
  if (index < 0) return null;
  const adjacent = direction === "previous" ? entries[index - 1] : entries[index + 1];
  return adjacent?.id ?? null;
}

