import { entryAttributes } from "./attributes.js";
import type { EntryRow, EntryUpdateOptions, ImportEntryInput, TagType, WorkbookRow } from "./db.js";

export const DETAIL_EXIT_CHOICES = [
  { label: "Save Changes (Preserve Test Data)", action: "save", preserveTestData: true },
  { label: "Save Changes (Reset Test Data)", action: "save", preserveTestData: false },
  { label: "Discard Changes", action: "discard" },
  { label: "Continue Editing", action: "continue" },
] as const;

export type DetailEditRow =
  | { id: "vocabulary"; kind: "vocabulary"; label: string }
  | { id: string; kind: "meaning"; label: string; index: number }
  | { id: string; kind: "attribute"; label: string; key: string }
  | { id: string; kind: "tags"; label: string; typeId: number };

export type DetailEditorState = {
  mode: "view" | "edit" | "tags" | "confirm";
  original: EntryRow;
  draft: ImportEntryInput;
  selected: number;
  exitChoice: number;
  picker: { typeId: number; cursor: number; tagIds: number[] } | null;
  error: string;
};

export type DetailEditorAction =
  | { type: "begin" }
  | { type: "move"; delta: number }
  | { type: "change"; value: string }
  | { type: "openTags" }
  | { type: "toggleTag" }
  | { type: "acceptTags" }
  | { type: "cancelTags" }
  | { type: "exit" }
  | { type: "moveChoice"; delta: number }
  | { type: "continue" }
  | { type: "discard" };

export function detailEditRows(workbook: WorkbookRow, tagTypes: TagType[]): DetailEditRow[] {
  return [
    { id: "vocabulary", kind: "vocabulary", label: workbook.vocabularyLabel },
    ...workbook.meaningAttributes.map((field, index): DetailEditRow => ({ id: `meaning-${index}`, kind: "meaning", label: field.label, index })),
    ...entryAttributes(workbook.metadataAttributes).map((field): DetailEditRow => ({ id: `attribute-${field.key}`, kind: "attribute", label: field.label, key: field.key })),
    ...tagTypes.map((type): DetailEditRow => ({ id: `tags-${type.id}`, kind: "tags", label: type.name, typeId: type.id })),
  ];
}

export function createDetailEditor(entry: EntryRow, editing = false): DetailEditorState {
  return {
    mode: editing ? "edit" : "view", original: entry,
    draft: { vocabulary: entry.vocabulary, meanings: [...entry.meanings], attributes: { ...entry.attributes }, tagIds: entry.tags.map((tag) => tag.id) },
    selected: 0, exitChoice: 0, picker: null, error: "",
  };
}

export function detailRowValue(state: DetailEditorState, row: DetailEditRow): string {
  if (row.kind === "vocabulary") return state.draft.vocabulary;
  if (row.kind === "meaning") return state.draft.meanings[row.index] ?? "";
  if (row.kind === "attribute") return state.draft.attributes[row.key] ?? "";
  return "";
}

function sameIds(first: number[], second: number[]): boolean {
  const a = new Set(first); const b = new Set(second);
  return a.size === b.size && [...a].every((id) => b.has(id));
}

export function detailRowChanged(state: DetailEditorState, row: DetailEditRow, tagTypes: TagType[]): boolean {
  if (row.kind === "vocabulary") return state.draft.vocabulary !== state.original.vocabulary;
  if (row.kind === "meaning") return detailRowValue(state, row) !== (state.original.meanings[row.index] ?? "");
  if (row.kind === "attribute") return detailRowValue(state, row) !== (state.original.attributes[row.key] ?? "");
  const ids = new Set(tagTypes.find((type) => type.id === row.typeId)?.tags.map((tag) => tag.id));
  return !sameIds(state.draft.tagIds.filter((id) => ids.has(id)), state.original.tags.filter((tag) => ids.has(tag.id)).map((tag) => tag.id));
}

export function detailHasChanges(state: DetailEditorState, workbook: WorkbookRow, tagTypes: TagType[]): boolean {
  return detailEditRows(workbook, tagTypes).some((row) => detailRowChanged(state, row, tagTypes));
}

function acceptTags(state: DetailEditorState, tagTypes: TagType[]): DetailEditorState {
  const picker = state.picker;
  if (!picker) return state;
  const typeIds = new Set(tagTypes.find((type) => type.id === picker.typeId)?.tags.map((tag) => tag.id));
  return { ...state, mode: "edit", picker: null, draft: { ...state.draft, tagIds: [...state.draft.tagIds.filter((id) => !typeIds.has(id)), ...picker.tagIds] } };
}

export function updateDetailEditor(state: DetailEditorState, action: DetailEditorAction, workbook: WorkbookRow, tagTypes: TagType[]): DetailEditorState {
  const rows = detailEditRows(workbook, tagTypes);
  if (action.type === "begin" && state.mode === "view") return createDetailEditor(state.original, true);
  if (action.type === "continue" && state.mode === "confirm") return { ...state, mode: "edit", error: "" };
  if (action.type === "discard" && state.mode === "confirm") return createDetailEditor(state.original);
  if (action.type === "moveChoice" && state.mode === "confirm") return { ...state, exitChoice: (state.exitChoice + action.delta + DETAIL_EXIT_CHOICES.length) % DETAIL_EXIT_CHOICES.length };
  if (action.type === "exit" && (state.mode === "edit" || state.mode === "tags")) {
    const staged = state.mode === "tags" ? acceptTags(state, tagTypes) : state;
    return detailHasChanges(staged, workbook, tagTypes) ? { ...staged, mode: "confirm", exitChoice: 0, error: "" } : createDetailEditor(state.original);
  }
  if (state.mode === "tags" && state.picker) {
    const picker = state.picker;
    const tags = tagTypes.find((type) => type.id === picker.typeId)?.tags ?? [];
    if (action.type === "move") return { ...state, picker: { ...state.picker, cursor: Math.max(0, Math.min(Math.max(0, tags.length - 1), state.picker.cursor + action.delta)) } };
    if (action.type === "toggleTag") {
      const tag = tags[state.picker.cursor];
      if (!tag) return state;
      const tagIds = state.picker.tagIds.includes(tag.id) ? state.picker.tagIds.filter((id) => id !== tag.id) : [...state.picker.tagIds, tag.id];
      return { ...state, picker: { ...state.picker, tagIds } };
    }
    if (action.type === "acceptTags") return acceptTags(state, tagTypes);
    if (action.type === "cancelTags") return { ...state, mode: "edit", picker: null };
    return state;
  }
  if (state.mode !== "edit") return state;
  if (action.type === "move") return { ...state, selected: Math.max(0, Math.min(rows.length - 1, state.selected + action.delta)), error: "" };
  const row = rows[state.selected];
  if (action.type === "openTags" && row?.kind === "tags") {
    const ids = new Set(tagTypes.find((type) => type.id === row.typeId)?.tags.map((tag) => tag.id));
    return { ...state, mode: "tags", error: "", picker: { typeId: row.typeId, cursor: 0, tagIds: state.draft.tagIds.filter((id) => ids.has(id)) } };
  }
  if (action.type !== "change" || !row || row.kind === "tags") return state;
  const draft = { ...state.draft };
  if (row.kind === "vocabulary") draft.vocabulary = action.value;
  else if (row.kind === "meaning") { draft.meanings = [...draft.meanings]; draft.meanings[row.index] = action.value; }
  else draft.attributes = { ...draft.attributes, [row.key]: action.value };
  return { ...state, draft, error: "" };
}

// The only persistence boundary: a draft cannot be written before exit confirmation.
export function saveDetailEditor(state: DetailEditorState, workbook: WorkbookRow, write: (draft: ImportEntryInput, options: EntryUpdateOptions) => EntryRow): DetailEditorState {
  const choice = DETAIL_EXIT_CHOICES[state.exitChoice];
  if (state.mode !== "confirm" || choice?.action !== "save") return state;
  try {
    const vocabulary = state.draft.vocabulary.trim();
    const meanings = state.draft.meanings.map((value) => value.trim());
    if (!vocabulary) return { ...state, mode: "edit", selected: 0, error: "Vocabulary is required." };
    if (!meanings[0]) return { ...state, mode: "edit", selected: 1, error: `${workbook.meaningAttributes[0]?.label ?? "Primary Meaning"} is required.` };
    const attributes = { ...state.draft.attributes };
    for (const field of entryAttributes(workbook.metadataAttributes)) attributes[field.key] = (attributes[field.key] ?? "").trim();
    return createDetailEditor(write({ vocabulary, meanings, attributes, tagIds: [...state.draft.tagIds] }, { preserveTestData: choice.preserveTestData }));
  } catch (caught) {
    return { ...state, mode: "edit", error: caught instanceof Error ? caught.message : "Could not save vocabulary changes." };
  }
}

type DetailKey = { ctrl?: boolean; escape?: boolean; upArrow?: boolean; downArrow?: boolean; leftArrow?: boolean; rightArrow?: boolean; return?: boolean };

export function detailEditorKeyAction(mode: DetailEditorState["mode"], input: string, key: DetailKey): DetailEditorAction | "confirmChoice" | null {
  if (mode === "view") return key.ctrl && input.toLowerCase() === "e" ? { type: "begin" } : null;
  if (mode === "confirm") {
    if (key.escape) return { type: "continue" };
    if (key.upArrow || key.downArrow) return { type: "moveChoice", delta: key.upArrow ? -1 : 1 };
    return key.return ? "confirmChoice" : null;
  }
  if (key.ctrl && input.toLowerCase() === "s") return { type: "exit" };
  if (key.escape) return { type: mode === "tags" ? "cancelTags" : "exit" };
  if (key.upArrow || key.downArrow) return { type: "move", delta: key.upArrow ? -1 : 1 };
  if (key.return) return { type: mode === "tags" ? "acceptTags" : "openTags" };
  if (mode === "tags" && input === " " && !key.ctrl) return { type: "toggleTag" };
  return null;
}
