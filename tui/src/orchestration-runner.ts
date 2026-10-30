import type { PracticeSessionResult, EntryRow, EntryUpdateOptions, WorkbookRow } from "./db.js";
import { loadLabeledTextFile, parseLabeledTextImport } from "./import.js";
import type { Effect, Intent, OrchestrationCapabilities, OrchestrationState } from "./orchestration.js";

export async function executeEffect(effect: Effect, state: OrchestrationState, capabilities: OrchestrationCapabilities): Promise<Intent> {
  switch (effect.type) {
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
    case "commitImport": {
      try {
        const result = capabilities.import.importEntries(effect.workbookId, effect.entries);
        const previewPath = effect.path;
        const preview = state.mode.kind === "importPreview" ? state.mode.preview : null;
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
    case "addEntry": {
      const entry = capabilities.writes.addEntry(effect.workbookId, effect.vocabulary, effect.meanings[0] ?? "", effect.meanings, effect.attributes, effect.tagIds);
      return { type: "entriesLoaded", entries: [...state.entries, entry], message: `Added #${entry.id}.` };
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
  listEntries(workbookId: number): EntryRow[]; getEntry(entryId: number): EntryRow | null; listTagTypes(workbookId: number): import("./db.js").TagType[]; getWorkbook(workbookId: number): WorkbookRow | null;
  addEntry(workbookId: number, vocabulary: string, meaning: string, meanings?: string[], attributes?: Record<string, string>, tagIds?: number[]): EntryRow; updateEntry(entryId: number, vocabulary: string, meaning: string, meanings?: string[], attributes?: Record<string, string>, tagIds?: number[], options?: EntryUpdateOptions): EntryRow; deleteEntry(entryId: number): void;
  importEntries(workbookId: number, entries: import("./db.js").ImportEntryInput[]): import("./db.js").ImportResult; setWorkbookImportFilePath(workbookId: number, path: string | null): WorkbookRow; selectPracticeCandidates(workbookId: number, count: number): EntryRow[]; completePracticeSession(workbookId: number, results: PracticeSessionResult[], completedDay: string): EntryRow[];
}): OrchestrationCapabilities {
  return { reads: { listEntries: (id) => backend.listEntries(id), getEntry: (id) => backend.getEntry(id), listTagTypes: (id) => backend.listTagTypes(id), getWorkbook: (id) => backend.getWorkbook(id) }, writes: { addEntry: (...args) => backend.addEntry(...args), updateEntry: (...args) => backend.updateEntry(...args), deleteEntry: (id) => backend.deleteEntry(id) }, import: { readLabeledTextFile: loadLabeledTextFile, importEntries: (id, entries) => backend.importEntries(id, entries), readWorkbookImportPath: (id) => backend.getWorkbook(id)?.importFilePath ?? null, saveWorkbookImportPath: (id, path) => backend.setWorkbookImportFilePath(id, path) }, practice: { selectPracticeCandidates: (id, count) => backend.selectPracticeCandidates(id, count), completePracticeSession: (id, results, day) => backend.completePracticeSession(id, results, day) } };
}
