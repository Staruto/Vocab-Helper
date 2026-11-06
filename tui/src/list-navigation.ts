import type { EntryRow } from "./db.js";

export const VOCABULARY_PAGE_SIZE = 20;
export type VocabularyListFocus = "table" | "search" | "command";
export type VocabularyListPosition = {
  pageIndex: number;
  selectedEntryId: number | null;
  inputFocus: VocabularyListFocus;
};
export type VocabularyListContext = VocabularyListPosition & {
  activeSearchQuery: string;
  searchDraft: string;
  commandBuffer: string;
};

export function restoreListPosition(entries: EntryRow[], position: VocabularyListPosition): VocabularyListPosition {
  const selectedIndex = entries.findIndex((entry) => entry.id === position.selectedEntryId);
  const lastPage = Math.max(0, Math.ceil(entries.length / VOCABULARY_PAGE_SIZE) - 1);
  const pageIndex = selectedIndex >= 0 ? Math.floor(selectedIndex / VOCABULARY_PAGE_SIZE) : Math.max(0, Math.min(position.pageIndex, lastPage));
  return {
    pageIndex,
    selectedEntryId: selectedIndex >= 0 ? position.selectedEntryId : entries[pageIndex * VOCABULARY_PAGE_SIZE]?.id ?? null,
    inputFocus: entries.length === 0 && position.inputFocus === "table" ? "search" : position.inputFocus,
  };
}

export function moveListFocus(entries: EntryRow[], position: VocabularyListPosition, delta: -1 | 1): VocabularyListPosition {
  const current = restoreListPosition(entries, position);
  const page = entries.slice(current.pageIndex * VOCABULARY_PAGE_SIZE, (current.pageIndex + 1) * VOCABULARY_PAGE_SIZE);
  if (current.inputFocus === "command") return { ...current, inputFocus: delta < 0 ? "search" : "command" };
  if (current.inputFocus === "search") return delta > 0
    ? { ...current, inputFocus: "command" }
    : page.length ? { ...current, inputFocus: "table", selectedEntryId: page[page.length - 1].id } : current;
  const index = page.findIndex((entry) => entry.id === current.selectedEntryId);
  if (delta > 0 && index === page.length - 1) return { ...current, inputFocus: "search" };
  return { ...current, selectedEntryId: page[Math.max(0, index + delta)].id };
}

export function changeListPage(entries: EntryRow[], position: VocabularyListPosition, delta: number): VocabularyListPosition {
  const current = restoreListPosition(entries, position);
  const selectedIndex = entries.findIndex((entry) => entry.id === current.selectedEntryId);
  const rowIndex = selectedIndex < 0 ? 0 : selectedIndex % VOCABULARY_PAGE_SIZE;
  const lastPage = Math.max(0, Math.ceil(entries.length / VOCABULARY_PAGE_SIZE) - 1);
  const pageIndex = Math.max(0, Math.min(current.pageIndex + delta, lastPage));
  const nextIndex = Math.min(pageIndex * VOCABULARY_PAGE_SIZE + rowIndex, entries.length - 1);
  return { ...current, pageIndex, selectedEntryId: entries[nextIndex]?.id ?? null };
}
