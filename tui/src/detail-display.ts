import { EntryRow, TagType, WorkbookRow } from "./db.js";
import { entryAttributes } from "./attributes.js";
import { calendarDay } from "./calendar.js";
import { visibleAssignedTagGroups, VisibleTagGroup } from "./tag-display.js";
import { parseSimilarWords, SimilarWord } from "./similar-words.js";
import { parseExampleSentence, StyledTextSpan } from "./styled-text.js";

export type DetailField = { key: string; value: string; similarWords?: SimilarWord[]; spans?: StyledTextSpan[] };

export function buildReviewFields(savedEntry: Pick<EntryRow, "testCount" | "errorCount" | "tier" | "lastTested" | "nextTestDeadline"> | null): DetailField[] {
  const entry = savedEntry ?? { testCount: 0, errorCount: 0, tier: "gray", lastTested: null, nextTestDeadline: null };
  return [
    { key: "Tests", value: String(entry.testCount) },
    { key: "Errors", value: String(entry.errorCount) },
    { key: "Tier", value: `${entry.tier[0].toUpperCase()}${entry.tier.slice(1)}` },
    { key: "Last tested", value: entry.lastTested ? calendarDay(entry.lastTested) : "Never" },
    { key: "Next deadline", value: entry.nextTestDeadline ? calendarDay(entry.nextTestDeadline) : "—" },
  ];
}

export type DetailSections = {
  meanings: DetailField[];
  attributes: DetailField[];
  tags: VisibleTagGroup[];
};

export function buildDetailSections(workbook: WorkbookRow, entry: Pick<EntryRow, "meanings" | "attributes" | "tags">, tagTypes: TagType[]): DetailSections {
  const meanings = entry.meanings.map((value, index) => ({
    key: workbook.meaningAttributes[index]?.label ?? `Meaning ${index + 1}`,
    value: value.trim() || "—",
  }));
  const attributes = entryAttributes(workbook.metadataAttributes)
    .map((attribute): DetailField => {
      const value = entry.attributes[attribute.key]?.trim() || "—";
      const similarWords = attribute.presetKey === "similar_words" ? parseSimilarWords(value) : null;
      const spans = /^example_sentence_[1-5]$/.test(attribute.presetKey ?? "") ? parseExampleSentence(value) : null;
      return { key: attribute.label, value, ...(similarWords ? { similarWords } : {}), ...(spans ? { spans } : {}) };
    });
  return { meanings, attributes, tags: visibleAssignedTagGroups(entry, tagTypes) };
}

export function adjacentEntryId(entries: EntryRow[], currentId: number, direction: "previous" | "next"): number | null {
  const index = entries.findIndex((entry) => entry.id === currentId);
  if (index < 0) return null;
  const adjacent = direction === "previous" ? entries[index - 1] : entries[index + 1];
  return adjacent?.id ?? null;
}

export function detailNavigationLabel(entry: EntryRow | null, direction: "previous" | "next"): string {
  if (!entry) return direction === "previous" ? "←" : "→";
  return direction === "previous" ? `← ${entry.vocabulary}` : `${entry.vocabulary} →`;
}
