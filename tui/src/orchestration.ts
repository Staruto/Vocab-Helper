import type { EntryRow, ImportEntryInput, ImportResult, TagType, WorkbookRow } from "./db.js";
import type { ImportPreview, ImportPreviewFilter, ImportDiagnostic } from "./import.js";

export type { ImportPreview, ImportPreviewFilter, ImportDiagnostic };
export type ImportPreviewFocus = "path" | "filters" | "records";
export type ParameterizedCommand = "edit" | "delete";

export type UiMode =
  | { kind: "command" }
  | { kind: "commandArg"; command: ParameterizedCommand }
  | { kind: "add"; stage: "vocabulary" | "meaning" | "metadata" | "tags"; vocabulary: string; meanings: string[]; meaningIndex: number; metadata: Record<string, string>; metadataIndex: number; selectedTagIds: number[]; tagIndex: number }
  | { kind: "edit"; stage: "vocabulary" | "meaning" | "metadata" | "tags"; entryId: number; vocabulary: string; meanings: string[]; meaningIndex: number; metadata: Record<string, string>; metadataIndex: number; selectedTagIds: number[]; tagIndex: number }
  | { kind: "delete"; entryId: number; label: string }
  | { kind: "importPath"; alternate: boolean }
  | { kind: "importLoading"; path: string; alternate: boolean; requestId?: number }
  | { kind: "importSaveDefault"; path: string; save: boolean; resultMessage: string }
  | { kind: "importPreview"; path: string; pathBuffer: string; preview: ImportPreview; pageIndex: number; filter: ImportPreviewFilter; focus: ImportPreviewFocus; loading?: boolean; error?: string };

export type PracticePhase = "empty" | "initial" | "retry" | "detail" | "done";
export type PracticeState = { phase: PracticePhase; candidates: EntryRow[]; index: number; retryRound: EntryRow[]; nextRetryRound: EntryRow[]; retryNumber: number; currentEntry: EntryRow | null; detailSourcePhase: "initial" | "retry"; score: number; feedback: string | null };

export type OrchestrationState = {
  workbook: WorkbookRow;
  entries: EntryRow[];
  tagTypes: TagType[];
  activeSearchQuery: string;
  pageIndex: number;
  commandBuffer: string;
  inputFocus: "command" | "search";
  statusLines: string[];
  mode: UiMode;
  importRequestId: number;
  importError?: string;
  practice: PracticeState;
};

export type OrchestrationCapabilities = {
  reads: { listEntries(workbookId: number): EntryRow[]; getEntry(entryId: number): EntryRow | null; listTagTypes(workbookId: number): TagType[]; getWorkbook(workbookId: number): WorkbookRow | null };
  writes: { addEntry(workbookId: number, vocabulary: string, meaning: string, meanings?: string[], attributes?: Record<string, string>, tagIds?: number[]): EntryRow; updateEntry(entryId: number, vocabulary: string, meaning: string, meanings?: string[], attributes?: Record<string, string>, tagIds?: number[]): EntryRow; deleteEntry(entryId: number): void };
  import: { readLabeledTextFile(path: string): Promise<{ path: string; text: string }>; importEntries(workbookId: number, entries: ImportEntryInput[]): ImportResult; readWorkbookImportPath(workbookId: number): string | null; saveWorkbookImportPath(workbookId: number, path: string | null): WorkbookRow };
  practice: { selectPracticeCandidates(workbookId: number, count: number): EntryRow[]; recordTestResult(entryId: number, isCorrect: boolean, decreaseError?: boolean): EntryRow };
};

export type Intent =
  | { type: "commandSubmitted"; raw: string }
  | { type: "searchChanged"; query: string }
  | { type: "pageChanged"; delta: number }
  | { type: "addFieldSubmitted"; value: string }
  | { type: "editFieldSubmitted"; value: string }
  | { type: "deleteConfirmationSubmitted"; value: string }
  | { type: "tagSelectionToggled"; tagId: number }
  | { type: "tagCursorMoved"; delta: number }
  | { type: "importPathLoadRequested"; path: string; alternate?: boolean }
  | { type: "importReloadRequested" }
  | { type: "importPreviewFilterChanged"; filter: ImportPreviewFilter }
  | { type: "importPreviewPageChanged"; delta: number }
  | { type: "importPathBufferChanged"; path: string }
  | { type: "importConfirmed" }
  | { type: "importSavePathDecision"; save: boolean }
  | { type: "entriesLoaded"; entries: EntryRow[]; message?: string }
  | { type: "importPreviewLoaded"; requestId: number; path: string; preview: ImportPreview }
  | { type: "importLoadFailed"; requestId: number; path: string; error: string }
  | { type: "importCommitted"; result: ImportResult; resultMessage: string; path: string }
  | { type: "importPathSaved"; path: string | null }
  | { type: "practiceStarted"; candidates: EntryRow[] }
  | { type: "practiceAnswerSubmitted"; answer: string }
  | { type: "practiceRetryAdvanced" }
  | { type: "practiceDetailContinued" }
  | { type: "practiceResultRecorded"; entry: EntryRow; isCorrect: boolean; initialRound: boolean };

export type Effect =
  | { type: "loadEntries"; workbookId: number }
  | { type: "loadImportPreview"; workbookId: number; path: string; requestId: number }
  | { type: "commitImport"; workbookId: number; path: string; entries: ImportEntryInput[] }
  | { type: "saveImportPath"; workbookId: number; path: string | null }
  | { type: "addEntry"; workbookId: number; vocabulary: string; meanings: string[]; attributes: Record<string, string>; tagIds: number[] }
  | { type: "updateEntry"; entryId: number; vocabulary: string; meanings: string[]; attributes: Record<string, string>; tagIds: number[] }
  | { type: "deleteEntry"; entryId: number }
  | { type: "recordPracticeResult"; entryId: number; isCorrect: boolean; initialRound: boolean };

export type Transition = { state: OrchestrationState; effects: Effect[] };

const status = (message: string, count = 5): string[] => [...message.split("\n").slice(-count), ...Array(Math.max(0, count - message.split("\n").length)).fill("")];
const pageCount = (count: number, size = 20): number => Math.max(1, Math.ceil(count / size));
const clampPage = (page: number, count: number): number => Math.max(0, Math.min(page, pageCount(count) - 1));
const normalize = (raw: string): string[] => raw.trim().replace(/^\/+/, "").split(/\s+/).filter(Boolean);
const entryLabel = (entry: EntryRow) => `#${entry.id} ${entry.vocabulary}`;

export function createInitialState(workbook: WorkbookRow, entries: EntryRow[] = [], tagTypes: TagType[] = []): OrchestrationState {
  return { workbook, entries, tagTypes, activeSearchQuery: "", pageIndex: 0, commandBuffer: "", inputFocus: "command", statusLines: status("Ready."), mode: { kind: "command" }, importRequestId: 0, practice: { phase: "empty", candidates: [], index: 0, retryRound: [], nextRetryRound: [], retryNumber: 1, currentEntry: null, detailSourcePhase: "initial", score: 0, feedback: null } };
}

function formComplete(state: OrchestrationState, form: Extract<UiMode, { kind: "add" | "edit" }>): Transition {
  const tags = state.workbook.metadataAttributes.length ? form.selectedTagIds : form.selectedTagIds;
  const effect: Effect = form.kind === "add" ? { type: "addEntry", workbookId: state.workbook.id, vocabulary: form.vocabulary, meanings: form.meanings, attributes: form.metadata, tagIds: tags } : { type: "updateEntry", entryId: form.entryId, vocabulary: form.vocabulary, meanings: form.meanings, attributes: form.metadata, tagIds: tags };
  return { state: { ...state, mode: { kind: "command" }, commandBuffer: "", statusLines: status(form.kind === "add" ? "Saving entry..." : "Updating entry...") }, effects: [effect] };
}

export function update(state: OrchestrationState, intent: Intent): Transition {
  let next = state;
  const effects: Effect[] = [];
  if (intent.type === "entriesLoaded") return { state: { ...state, entries: intent.entries, pageIndex: clampPage(state.pageIndex, intent.entries.length), statusLines: status(intent.message ?? `Loaded ${intent.entries.length} entr${intent.entries.length === 1 ? "y" : "ies"}.`) }, effects: [] };
  if (intent.type === "searchChanged") return { state: { ...state, activeSearchQuery: intent.query.trim(), pageIndex: 0, statusLines: status(intent.query.trim() ? "Search updated." : "Search cleared.") }, effects: [] };
  if (intent.type === "pageChanged") return { state: { ...state, pageIndex: clampPage(state.pageIndex + intent.delta, state.entries.length) }, effects: [] };
  if (intent.type === "commandSubmitted") {
    const [command, ...args] = normalize(intent.raw); const lower = command?.toLowerCase();
    if (lower === "list") return { state: { ...state, commandBuffer: "" }, effects: [{ type: "loadEntries", workbookId: state.workbook.id }] };
    if (lower === "search") return update(state, { type: "searchChanged", query: args.join(" " ) });
    if (lower === "add") return { state: { ...state, mode: { kind: "add", stage: "vocabulary", vocabulary: "", meanings: [], meaningIndex: 0, metadata: {}, metadataIndex: 0, selectedTagIds: [], tagIndex: 0 }, commandBuffer: "", statusLines: status(`Adding a new entry.\nEnter ${state.workbook.vocabularyLabel}.`) }, effects: [] };
    if (lower === "import") { const saved = state.workbook.importFilePath; return saved ? update(state, { type: "importPathLoadRequested", path: saved, alternate: false }) : { state: { ...state, mode: { kind: "importPath", alternate: true }, commandBuffer: "", statusLines: status("Enter the path to a UTF-8 .txt import file.") }, effects: [] }; }
    if (lower === "edit" || lower === "delete") { const id = Number(args[0]); if (!Number.isInteger(id)) return { state: { ...state, mode: { kind: "commandArg", command: lower }, commandBuffer: `/${lower} `, statusLines: status(`Enter id for /${lower}.`) }, effects: [] }; const entry = state.entries.find((item) => item.id === id); if (!entry) return { state: { ...state, statusLines: status(`Entry #${id} was not found.`) }, effects: [] }; if (lower === "delete") return { state: { ...state, mode: { kind: "delete", entryId: id, label: entryLabel(entry) }, commandBuffer: "", statusLines: status(`Type yes to delete ${entryLabel(entry)}.`) }, effects: [] }; return { state: { ...state, mode: { kind: "edit", stage: "vocabulary", entryId: id, vocabulary: entry.vocabulary, meanings: [...entry.meanings], meaningIndex: 0, metadata: { ...entry.attributes }, metadataIndex: 0, selectedTagIds: entry.tags.map((tag) => tag.id), tagIndex: 0 }, commandBuffer: entry.vocabulary, statusLines: status(`Editing #${id}.\nEdit vocabulary.`) }, effects: [] }; }
    return { state: { ...state, statusLines: status(command ? `Unknown command: ${command}` : "") }, effects: [] };
  }
  if (intent.type === "addFieldSubmitted" || intent.type === "editFieldSubmitted") {
    const mode = state.mode; if (mode.kind !== (intent.type === "addFieldSubmitted" ? "add" : "edit")) return { state, effects: [] }; const text = intent.value.trim();
    if (mode.stage === "vocabulary") { if (!text) return { state: { ...state, statusLines: status("Vocabulary is required.") }, effects: [] }; return { state: { ...state, mode: { ...mode, stage: "meaning", vocabulary: text }, commandBuffer: "", statusLines: status(`Enter ${state.workbook.meaningAttributes[0]?.label ?? "Primary Meaning"}.`) }, effects: [] }; }
    if (mode.stage === "meaning") { if (mode.meaningIndex === 0 && !text) return { state: { ...state, statusLines: status("Meaning is required.") }, effects: [] }; const meanings = [...mode.meanings]; meanings[mode.meaningIndex] = text; if (mode.meaningIndex + 1 < state.workbook.meaningAttributes.length) return { state: { ...state, mode: { ...mode, meanings, meaningIndex: mode.meaningIndex + 1 }, commandBuffer: "", statusLines: status(`Enter ${state.workbook.meaningAttributes[mode.meaningIndex + 1].label}. Optional.`) }, effects: [] }; const optional = state.workbook.metadataAttributes.filter((field) => field.role === "optional"); if (optional.length) return { state: { ...state, mode: { ...mode, meanings, stage: "metadata", metadataIndex: 0 }, commandBuffer: "", statusLines: status(`Enter ${optional[0].label}. Optional.`) }, effects: [] }; if (state.tagTypes.flatMap((type) => type.tags).length) return { state: { ...state, mode: { ...mode, meanings, stage: "tags", tagIndex: 0 }, commandBuffer: "", statusLines: status("Select tags with Space, then press Enter.") }, effects: [] }; return formComplete(state, { ...mode, meanings }); }
    if (mode.stage === "metadata") { const optional = state.workbook.metadataAttributes.filter((field) => field.role === "optional"); const field = optional[mode.metadataIndex]; const metadata = { ...mode.metadata, ...(field ? { [field.key]: text } : {}) }; if (mode.metadataIndex + 1 < optional.length) return { state: { ...state, mode: { ...mode, metadata, metadataIndex: mode.metadataIndex + 1 }, commandBuffer: "", statusLines: status(`Enter ${optional[mode.metadataIndex + 1].label}. Optional.`) }, effects: [] }; if (state.tagTypes.flatMap((type) => type.tags).length) return { state: { ...state, mode: { ...mode, metadata, stage: "tags", tagIndex: 0 }, commandBuffer: "", statusLines: status("Select tags with Space, then press Enter.") }, effects: [] }; return formComplete(state, { ...mode, metadata }); }
    if (mode.stage === "tags") return formComplete(state, mode);
  }
  if (intent.type === "deleteConfirmationSubmitted") { if (state.mode.kind !== "delete") return { state, effects: [] }; if (intent.value.trim().toLowerCase() !== "yes") return { state: { ...state, mode: { kind: "command" }, commandBuffer: "", statusLines: status("Delete cancelled.") }, effects: [] }; return { state: { ...state, mode: { kind: "command" }, commandBuffer: "", statusLines: status("Deleting entry...") }, effects: [{ type: "deleteEntry", entryId: state.mode.entryId }, { type: "loadEntries", workbookId: state.workbook.id }] }; }
  if (intent.type === "tagSelectionToggled") { const mode = state.mode; if ((mode.kind !== "add" && mode.kind !== "edit") || mode.stage !== "tags") return { state, effects: [] }; const selectedTagIds = mode.selectedTagIds.includes(intent.tagId) ? mode.selectedTagIds.filter((id) => id !== intent.tagId) : [...mode.selectedTagIds, intent.tagId]; return { state: { ...state, mode: { ...mode, selectedTagIds } }, effects: [] }; }
  if (intent.type === "tagCursorMoved") { const mode = state.mode; if ((mode.kind !== "add" && mode.kind !== "edit") || mode.stage !== "tags") return { state, effects: [] }; return { state: { ...state, mode: { ...mode, tagIndex: Math.max(0, mode.tagIndex + intent.delta) } }, effects: [] }; }
  if (intent.type === "importPathLoadRequested") { const requestId = state.importRequestId + 1; const path = intent.path.trim(); if (!path) return { state: { ...state, statusLines: status("A file path is required.") }, effects: [] }; return { state: { ...state, importRequestId: requestId, mode: { kind: "importLoading", path, alternate: Boolean(intent.alternate), requestId }, statusLines: status("Reading and validating import file...") }, effects: [{ type: "loadImportPreview", workbookId: state.workbook.id, path, requestId }] }; }
  if (intent.type === "importPreviewLoaded") { if (intent.requestId !== state.importRequestId) return { state, effects: [] }; return { state: { ...state, mode: { kind: "importPreview", path: intent.path, pathBuffer: intent.path, preview: intent.preview, pageIndex: 0, filter: "records", focus: "filters" }, statusLines: status("Import preview ready.") }, effects: [] }; }
  if (intent.type === "importLoadFailed") { if (intent.requestId !== state.importRequestId) return { state, effects: [] }; return { state: { ...state, mode: { kind: "importPath", alternate: true }, commandBuffer: intent.path, statusLines: status(`Could not load '${intent.path}': ${intent.error}`), importError: intent.error }, effects: [] }; }
  if (intent.type === "importPathBufferChanged") { const mode = state.mode; return mode.kind === "importPreview" ? { state: { ...state, mode: { ...mode, pathBuffer: intent.path, error: undefined } }, effects: [] } : { state: { ...state, commandBuffer: intent.path }, effects: [] }; }
  if (intent.type === "importPreviewFilterChanged") { const mode = state.mode; return mode.kind === "importPreview" ? { state: { ...state, mode: { ...mode, filter: intent.filter, pageIndex: 0 } }, effects: [] } : { state, effects: [] }; }
  if (intent.type === "importPreviewPageChanged") { const mode = state.mode; return mode.kind === "importPreview" ? { state: { ...state, mode: { ...mode, pageIndex: Math.max(0, mode.pageIndex + intent.delta) } }, effects: [] } : { state, effects: [] }; }
  if (intent.type === "importConfirmed") { const mode = state.mode; if (mode.kind !== "importPreview" || mode.preview.entries.length === 0) return { state, effects: [] }; return { state: { ...state, statusLines: status("Importing entries...") }, effects: [{ type: "commitImport", workbookId: state.workbook.id, path: mode.path, entries: mode.preview.entries.map(({ recordNumber: _recordNumber, ...entry }) => entry) }] }; }
  if (intent.type === "importCommitted") { const save = state.workbook.importFilePath !== intent.path; return { state: { ...state, mode: save ? { kind: "importSaveDefault", path: intent.path, save: false, resultMessage: intent.resultMessage } : { kind: "command" }, commandBuffer: "", statusLines: status(intent.resultMessage) }, effects: save ? [] : [{ type: "loadEntries", workbookId: state.workbook.id }] }; }
  if (intent.type === "importSavePathDecision") { const mode = state.mode; if (mode.kind !== "importSaveDefault") return { state, effects: [] }; return { state: { ...state, mode: { kind: "command" }, commandBuffer: "" }, effects: intent.save ? [{ type: "saveImportPath", workbookId: state.workbook.id, path: mode.path }, { type: "loadEntries", workbookId: state.workbook.id }] : [{ type: "loadEntries", workbookId: state.workbook.id }] }; }
  if (intent.type === "importPathSaved") return { state: { ...state, workbook: { ...state.workbook, importFilePath: intent.path }, statusLines: status("Import path saved.") }, effects: [] };
  if (intent.type === "practiceStarted") { const phase: PracticePhase = intent.candidates.length ? "initial" : "empty"; return { state: { ...state, practice: { ...state.practice, phase, candidates: intent.candidates, index: 0, retryRound: [], nextRetryRound: [], retryNumber: 1, currentEntry: intent.candidates[0] ?? null, score: 0, feedback: null } }, effects: [] }; }
  if (intent.type === "practiceAnswerSubmitted") { const p = state.practice; const current = p.phase === "initial" ? p.candidates[p.index] : p.phase === "retry" ? p.retryRound[p.index] : null; if (!current) return { state: { ...state, commandBuffer: "", practice: { ...p, phase: "done", currentEntry: null } }, effects: [] }; const isCorrect = intent.answer.trim() === current.vocabulary; return { state: { ...state, commandBuffer: "", practice: { ...p, feedback: isCorrect ? "Correct!" : null, currentEntry: current, phase: isCorrect ? p.phase : "detail", detailSourcePhase: p.phase === "retry" ? "retry" : "initial" } }, effects: [{ type: "recordPracticeResult", entryId: current.id, isCorrect, initialRound: p.phase === "initial" }] }; }
  if (intent.type === "practiceResultRecorded") { const p = state.practice; const score = intent.initialRound && intent.isCorrect ? p.score + 1 : p.score; return { state: { ...state, practice: { ...p, score, currentEntry: intent.entry } }, effects: [] }; }
  if (intent.type === "practiceDetailContinued" || intent.type === "practiceRetryAdvanced") { const p = state.practice; const source = p.detailSourcePhase; const queued = p.currentEntry ? [...p.nextRetryRound, p.currentEntry] : p.nextRetryRound; const round = source === "initial" ? p.candidates : p.retryRound; const nextIndex = p.index + 1; if (nextIndex < round.length) return { state: { ...state, practice: { ...p, phase: source, index: nextIndex, nextRetryRound: queued, currentEntry: round[nextIndex], feedback: null } }, effects: [] }; if (queued.length) return { state: { ...state, practice: { ...p, phase: "retry", retryRound: queued, nextRetryRound: [], index: 0, retryNumber: source === "retry" ? p.retryNumber + 1 : 1, currentEntry: queued[0], feedback: null } }, effects: [] }; return { state: { ...state, practice: { ...p, phase: "done", currentEntry: null, nextRetryRound: queued, feedback: null } }, effects: [] }; }
  return { state: next, effects };
}

