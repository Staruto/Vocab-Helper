import type { PracticeSessionResult, EntryRow, EntryUpdateOptions, WorkbookRow } from "./db.js";
import { loadLabeledTextFile, parseLabeledTextImport } from "./import.js";
import type { ListOrder } from "./list-order.js";
import type { Effect, Intent, OrchestrationCapabilities, OrchestrationState } from "./orchestration.js";

export async function executeEffect(effect: Effect, state: OrchestrationState, capabilities: OrchestrationCapabilities): Promise<Intent> {
  switch (effect.type) {
    case "saveListOrder":
      try {
        const workbook = capabilities.list.saveOrder(effect.workbookId, effect.order);
        return { type: "listOrderSaved", workbook, entries: capabilities.reads.listEntries(effect.workbookId) };
      } catch (error) {
        return { type: "listOrderFailed", error: error instanceof Error ? error.message : "Could not save or reload list order." };
      }
    case "loadEntries":
      return { type: "entriesLoaded", entries: capabilities.reads.listEntries(effect.workbookId) };
    case "loadPracticeCandidates":
      try {
        return { type: "practiceStarted", requestId: effect.requestId, candidates: capabilities.practice.selectPracticeCandidates(effect.workbookId, effect.count) };
      } catch (error) {
        return { type: "practiceCandidatesFailed", requestId: effect.requestId, error: error instanceof Error ? error.message : "Could not load practice entries." };
      }
    case "loadImportPreview": {
      try {
        const source = await capabilities.import.readLabeledTextFile(effect.path);
        const preview = parseLabeledTextImport(source.text, state.workbook, state.tagTypes, state.entries.map((entry) => entry.vocabulary));
        return { type: "importPreviewLoaded", requestId: effect.requestId, path: source.path, preview };
      } catch (error) {
        return { type: "importLoadFailed", requestId: effect.requestId, path: effect.path, error: error instanceof Error ? error.message : "Could not read the import file." };
      }
    }
    case "checkImportUpdates": {
      try {
        return { type: "importUpdatesChecked", updated: capabilities.import.previewImportUpdates(effect.workbookId, effect.entries) };
      } catch (error) {
        return { type: "importCommitFailed", error: error instanceof Error ? error.message : "Could not check import updates." };
      }
    }
    case "commitImport": {
      try {
        const result = capabilities.import.importEntries(effect.workbookId, effect.entries, effect.options);
        const previewPath = effect.path;
        const preview = state.mode.kind === "importPending" ? state.mode.previous.preview : null;
        const ignoredFields = preview?.diagnostics.filter((item) => item.kind === "ignored-field").length ?? 0;
        const ignoredTags = preview?.diagnostics.filter((item) => item.kind === "ignored-tag").length ?? 0;
        const resultMessage = `Added ${result.added}; Updated ${result.updated}; Unchanged ${result.unchanged}; Invalid ${preview?.skippedInvalid ?? 0}; ignored fields ${ignoredFields}; ignored tags ${ignoredTags}.`;
        return { type: "importCommitted", result, resultMessage, path: previewPath };
      } catch (error) {
        return { type: "importCommitFailed", error: error instanceof Error ? error.message : "Import failed; no entries were written." };
      }
    }
    case "saveImportPath": {
      const workbook = capabilities.import.saveWorkbookImportPath(effect.workbookId, effect.path);
      return { type: "importPathSaved", path: workbook.importFilePath };
    }
    case "deleteEntry":
      capabilities.writes.deleteEntry(effect.entryId);
      return { type: "entriesLoaded", entries: state.entries.filter((entry) => entry.id !== effect.entryId), message: "Entry deleted." };
    case "completePracticeSession": {
      try {
        const entries = capabilities.practice.completePracticeSession(effect.workbookId, effect.results, effect.completedDay);
        return { type: "practiceSessionSaved", requestId: effect.requestId, entries };
      } catch (error) {
        return { type: "practiceSaveFailed", requestId: effect.requestId, error: error instanceof Error ? error.message : "Could not save practice results." };
      }
    }
  }
}

export function createOrchestrationCapabilities(backend: {
  setWorkbookListOrder(workbookId: number, order: ListOrder): WorkbookRow;
  listEntries(workbookId: number): EntryRow[]; getEntry(entryId: number): EntryRow | null; listTagTypes(workbookId: number): import("./db.js").TagType[]; getWorkbook(workbookId: number): WorkbookRow | null;
  addEntry(workbookId: number, vocabulary: string, meaning: string, meanings?: string[], attributes?: Record<string, string>, tagIds?: number[]): EntryRow; updateEntry(entryId: number, vocabulary: string, meaning: string, meanings?: string[], attributes?: Record<string, string>, tagIds?: number[], options?: EntryUpdateOptions): EntryRow; deleteEntry(entryId: number): void;
  previewImportUpdates(workbookId: number, entries: import("./db.js").ImportEntryInput[]): number;
  importEntries(workbookId: number, entries: import("./db.js").ImportEntryInput[], options?: EntryUpdateOptions): import("./db.js").ImportResult; setWorkbookImportFilePath(workbookId: number, path: string | null): WorkbookRow; selectPracticeCandidates(workbookId: number, count: number): EntryRow[]; completePracticeSession(workbookId: number, results: PracticeSessionResult[], completedDay: string): EntryRow[];
}): OrchestrationCapabilities {
  return { list: { saveOrder: (id, order) => backend.setWorkbookListOrder(id, order) }, reads: { listEntries: (id) => backend.listEntries(id), getEntry: (id) => backend.getEntry(id), listTagTypes: (id) => backend.listTagTypes(id), getWorkbook: (id) => backend.getWorkbook(id) }, writes: { addEntry: (...args) => backend.addEntry(...args), updateEntry: (...args) => backend.updateEntry(...args), deleteEntry: (id) => backend.deleteEntry(id) }, import: { readLabeledTextFile: loadLabeledTextFile, previewImportUpdates: (id, entries) => backend.previewImportUpdates(id, entries), importEntries: (id, entries, options) => backend.importEntries(id, entries, options), readWorkbookImportPath: (id) => backend.getWorkbook(id)?.importFilePath ?? null, saveWorkbookImportPath: (id, path) => backend.setWorkbookImportFilePath(id, path) }, practice: { selectPracticeCandidates: (id, count) => backend.selectPracticeCandidates(id, count), completePracticeSession: (id, results, day) => backend.completePracticeSession(id, results, day) } };
}
