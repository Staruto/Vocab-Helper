import type { EntryRow, WorkbookRow } from "./db.js";
import { loadLabeledTextFile, parseLabeledTextImport, shouldPromptToSaveImportPath } from "./import.js";
import type { Effect, Intent, OrchestrationCapabilities, OrchestrationState } from "./orchestration.js";

export async function executeEffect(effect: Effect, state: OrchestrationState, capabilities: OrchestrationCapabilities): Promise<Intent> {
  switch (effect.type) {
    case "loadEntries":
      return { type: "entriesLoaded", entries: capabilities.reads.listEntries(effect.workbookId) };
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
        const resultMessage = `Added ${result.added}; Updated ${result.updated}; Unchanged ${result.unchanged}.`;
        return { type: "importCommitted", result, resultMessage, path: previewPath };
      } catch (error) {
        return { type: "importLoadFailed", requestId: state.importRequestId, path: effect.path, error: error instanceof Error ? error.message : "Import failed; no entries were written." };
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
    case "updateEntry": {
      const entry = capabilities.writes.updateEntry(effect.entryId, effect.vocabulary, effect.meanings[0] ?? "", effect.meanings, effect.attributes, effect.tagIds);
      return { type: "entriesLoaded", entries: state.entries.map((item) => item.id === entry.id ? entry : item), message: `Updated #${entry.id}.` };
    }
    case "deleteEntry":
      capabilities.writes.deleteEntry(effect.entryId);
      return { type: "entriesLoaded", entries: state.entries.filter((entry) => entry.id !== effect.entryId), message: "Entry deleted." };
    case "recordPracticeResult": {
      const entry = capabilities.practice.recordTestResult(effect.entryId, effect.isCorrect, effect.initialRound);
      return { type: "practiceResultRecorded", entry, isCorrect: effect.isCorrect, initialRound: effect.initialRound };
    }
  }
}

export function createOrchestrationCapabilities(backend: {
  listEntries(workbookId: number): EntryRow[]; getEntry(entryId: number): EntryRow | null; listTagTypes(workbookId: number): import("./db.js").TagType[]; getWorkbook(workbookId: number): WorkbookRow | null;
  addEntry(workbookId: number, vocabulary: string, meaning: string, meanings?: string[], attributes?: Record<string, string>, tagIds?: number[]): EntryRow; updateEntry(entryId: number, vocabulary: string, meaning: string, meanings?: string[], attributes?: Record<string, string>, tagIds?: number[]): EntryRow; deleteEntry(entryId: number): void;
  importEntries(workbookId: number, entries: import("./db.js").ImportEntryInput[]): import("./db.js").ImportResult; setWorkbookImportFilePath(workbookId: number, path: string | null): WorkbookRow; selectPracticeCandidates(workbookId: number, count: number): EntryRow[]; recordTestResult(entryId: number, isCorrect: boolean, decreaseError?: boolean): EntryRow;
}): OrchestrationCapabilities {
  return { reads: { listEntries: (id) => backend.listEntries(id), getEntry: (id) => backend.getEntry(id), listTagTypes: (id) => backend.listTagTypes(id), getWorkbook: (id) => backend.getWorkbook(id) }, writes: { addEntry: (...args) => backend.addEntry(...args), updateEntry: (...args) => backend.updateEntry(...args), deleteEntry: (id) => backend.deleteEntry(id) }, import: { readLabeledTextFile: loadLabeledTextFile, importEntries: (id, entries) => backend.importEntries(id, entries), readWorkbookImportPath: (id) => backend.getWorkbook(id)?.importFilePath ?? null, saveWorkbookImportPath: (id, path) => backend.setWorkbookImportFilePath(id, path) }, practice: { selectPracticeCandidates: (id, count) => backend.selectPracticeCandidates(id, count), recordTestResult: (id, correct, decreaseError) => backend.recordTestResult(id, correct, decreaseError) } };
}
