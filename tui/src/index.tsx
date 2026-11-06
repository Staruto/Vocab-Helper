import { entryAttributeValue, isActiveAttribute, createPresetAttributes } from "./attributes.js";
import React, { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { Box, Text, render, useApp, useInput, useStdout } from "ink";
import { CreateWorkbookInput, EntryRow, LANGUAGE_PRESET_DEFINITIONS, MeaningAttribute, MetadataAttribute, TagDataLossError, TagType, TagTypeDraft, VocabularyKind, WorkbookConfigurationInput, WorkbookDataLossError, WorkbookRow, WorkbookTagsDraft } from "./db.js";
import { VocabularyBackend } from "./backend.js";
import { fitTagBadges, visibleAssignedTagGroups } from "./tag-display.js";
import { adjacentEntryId, buildDetailSections, buildReviewFields, detailNavigationLabel, DetailField } from "./detail-display.js";
import { similarWordLines } from "./similar-words.js";
import { wrapStyledText } from "./styled-text.js";
import { formatImportPreviewFilter, formatImportPreviewRecord, importPreviewPageSize, importPreviewWordCount, IMPORT_EXIT_CHOICES, ImportPreview, ImportPreviewFilter, loadLabeledTextFile, paginateImportPreview } from "./import.js";
import { createAddDetailEditor, createDetailEditor, detailExitChoices, detailEditRows, detailEditorKeyAction, detailRowChanged, detailRowValue, saveDetailEditor, updateDetailEditor, DetailEditRow, DetailEditorState } from "./detail-editor.js";
import { SettingsScreen, type SettingsPanelControls, type AttributeDraftControls, type TagDraftControls } from "./settings-screen.js";
import { CaretInputLine } from "./text-input.js";
import { filterEntriesForSearch } from "./search.js";
import { LIST_ORDER_OPTIONS } from "./list-order.js";
import { restoreListPosition, VOCABULARY_PAGE_SIZE, type VocabularyListContext } from "./list-navigation.js";
import { localCalendarDay } from "./calendar.js";
import { buildPracticeProfile } from "./profile-display.js";
import { buildWorkbookMenuLines, buildWorkbookProgress } from "./workbook-display.js";
import { UiMode } from "./orchestration.js";
import { createOrchestrationCapabilities } from "./orchestration-runner.js";
import { createInitialState, update as updateOrchestration } from "./orchestration.js";
import { createOrchestrationController, runTransitionEffects } from "./ink-orchestration.js";

type ImportPreviewFocus = "path" | "filters" | "records";

type AppScreen =
  | { kind: "menu" }
  | { kind: "create-workbook" }
  | { kind: "edit-workbook"; workbook: WorkbookRow }
  | { kind: "delete-workbook"; workbook: WorkbookRow; confirm: string }
  | { kind: "settings"; workbook: WorkbookRow }
  | { kind: "tags"; workbook: WorkbookRow }
  | { kind: "add"; workbook: WorkbookRow }
  | { kind: "view"; workbook: WorkbookRow; entryId: number; editing?: boolean; listContext?: VocabularyListContext }
  | { kind: "profile"; workbook: WorkbookRow }
  | { kind: "practice"; workbook: WorkbookRow; count: number }
  | { kind: "vocab"; workbook: WorkbookRow; listContext?: VocabularyListContext };

type VocabularyScreenProps = {
  workbook: WorkbookRow;
  initialListContext?: VocabularyListContext;
  onBackToMenu: () => void;
  onQuit: () => void;
  onOpenSettings: () => void;
  onOpenTags: () => void;
  onOpenProfile: () => void;
  onAddEntry: () => void;
  onViewEntry: (entryId: number, editing: boolean, listContext: VocabularyListContext) => void;
  onStartPractice: (count?: number) => void;
};

type CommandSpec = {
  name: string;
  hint: string;
};

type ParameterizedCommand = "edit" | "delete";

type LanguagePreset = { code: string; label: string };

const PAGE_SIZE = VOCABULARY_PAGE_SIZE;
const TITLE = "VocabHelper 3.8.0";
const AUXILIARY_TEXT_COLOR = "#8a8a8a"; // light gray
const GRAY_TIER_COLOR = "#777777"; // medium gray
const SELECTED_TEXT_COLOR = "#c5a8ff";
const TAG_BADGE_COLOR = "#5fd7d7";
const YELLOW_TIER_COLOR = "#e5c65f";
const COMMAND_SUGGESTION_ROWS = 6;
const LANGUAGE_PRESETS: LanguagePreset[] = [
  { code: "JP", label: "Japanese" },
  { code: "EN", label: "English" },
  { code: "ZH", label: "Chinese" },
  { code: "KO", label: "Korean" },
  { code: "ES", label: "Spanish" },
  { code: "FR", label: "French" },
  { code: "DE", label: "German" },
];
const VOCABULARY_TYPES: Array<{ kind: VocabularyKind; code: string | null; label: string }> = [
  ...LANGUAGE_PRESETS.map((item) => ({ kind: "preset_language" as const, code: item.code, label: `${item.label} (${item.code})` })),
  { kind: "other_language", code: null, label: "Other Language" },
  { kind: "non_language", code: null, label: "Not a Language" },
];
const WORKBOOK_MENU_HINT = "↑↓ select | Enter open | Ctrl+E edit | Del delete | Esc exit";
const WORKBOOK_DELETE_HINT = "Type yes to confirm. Enter deletes. Esc cancels.";
const COMMANDS: CommandSpec[] = [
  { name: "list", hint: "Order and refresh entries" },
  { name: "search", hint: "Search visible vocabulary properties" },
  { name: "add", hint: "Add a new entry" },
  { name: "import", hint: "Import entries from labeled text" },
  { name: "edit", hint: "Open entry details in edit mode" },
  { name: "delete", hint: "Delete an entry by id" },
  { name: "menu", hint: "Return to the workbook menu" },
  { name: "help", hint: "Show command help" },
  { name: "setting", hint: "Configure workbook fields" },
  { name: "tag", hint: "Manage workbook tags" },
  { name: "view", hint: "View entry details and status" },
  { name: "profile", hint: "Show practice calendar and streaks" },
  { name: "test", hint: "Practice vocabulary (15 questions)" },
  { name: "exit", hint: "Exit the app" },
];
const backend = new VocabularyBackend();
const orchestrationCapabilities = createOrchestrationCapabilities(backend);

function writeToStdout(text: string): void {
  process.stdout?.write?.(text);
}

function enterAlternateScreen(): void {
  writeToStdout("\x1b[?1049h\x1b[?25l\x1b[2J\x1b[H");
}

function leaveAlternateScreen(): void {
  writeToStdout("\x1b[?1049l\x1b[?25h");
}

function normalizeCommand(input: string): string[] {
  const cleaned = input.trim().replace(/^\/+/, "");
  if (!cleaned) {
    return [];
  }
  return cleaned.split(/\s+/);
}

function isCombiningMark(codePoint: number): boolean {
  return (
    (codePoint >= 0x0300 && codePoint <= 0x036f) ||
    (codePoint >= 0x1ab0 && codePoint <= 0x1aff) ||
    (codePoint >= 0x1dc0 && codePoint <= 0x1dff) ||
    (codePoint >= 0x20d0 && codePoint <= 0x20ff) ||
    (codePoint >= 0xfe20 && codePoint <= 0xfe2f)
  );
}

function isWide(codePoint: number): boolean {
  return (
    codePoint >= 0x1100 &&
    (
      codePoint <= 0x115f ||
      codePoint === 0x2329 ||
      codePoint === 0x232a ||
      (codePoint >= 0x2e80 && codePoint <= 0xa4cf && codePoint !== 0x303f) ||
      (codePoint >= 0xac00 && codePoint <= 0xd7a3) ||
      (codePoint >= 0xf900 && codePoint <= 0xfaff) ||
      (codePoint >= 0xfe10 && codePoint <= 0xfe19) ||
      (codePoint >= 0xfe30 && codePoint <= 0xfe6f) ||
      (codePoint >= 0xff00 && codePoint <= 0xff60) ||
      (codePoint >= 0xffe0 && codePoint <= 0xffe6) ||
      (codePoint >= 0x1f300 && codePoint <= 0x1f64f) ||
      (codePoint >= 0x1f900 && codePoint <= 0x1f9ff) ||
      (codePoint >= 0x20000 && codePoint <= 0x3fffd)
    )
  );
}

function displayWidth(text: string): number {
  let width = 0;
  for (const char of Array.from(text)) {
    const codePoint = char.codePointAt(0) ?? 0;
    if (isCombiningMark(codePoint)) {
      continue;
    }
    width += isWide(codePoint) ? 2 : 1;
  }
  return width;
}

function sliceToWidth(text: string, maxWidth: number): string {
  if (maxWidth <= 0) {
    return "";
  }

  let width = 0;
  let result = "";
  for (const char of Array.from(text)) {
    const codePoint = char.codePointAt(0) ?? 0;
    const charWidth = isCombiningMark(codePoint) ? 0 : isWide(codePoint) ? 2 : 1;
    if (width + charWidth > maxWidth) {
      break;
    }
    result += char;
    width += charWidth;
  }
  return result;
}

function truncate(text: string, width: number): string {
  if (width <= 0) {
    return "";
  }
  if (displayWidth(text) <= width) {
    return text;
  }
  if (width === 1) {
    return ".";
  }
  return `${sliceToWidth(text, width - 1)}.`;
}

function padLine(text: string, width: number): string {
  const clipped = truncate(text, width);
  return `${clipped}${" ".repeat(Math.max(0, width - displayWidth(clipped)))}`;
}

function AttributeRowText({ text, inactiveLabel }: { text: string; inactiveLabel?: string }): JSX.Element {
  if (inactiveLabel === undefined) return <>{text}</>;
  const labelEnd = Math.min(text.length, 2 + inactiveLabel.length);
  return <>{text.slice(0, 2)}<Text strikethrough dimColor>{text.slice(2, labelEnd)}</Text>{text.slice(labelEnd)}</>;
}

function centerLine(text: string, width: number): string {
  if (width <= 0) {
    return "";
  }

  const clipped = truncate(text, width);
  const clippedWidth = displayWidth(clipped);
  const leftPadding = Math.max(0, Math.floor((width - clippedWidth) / 2));
  const rightPadding = Math.max(0, width - clippedWidth - leftPadding);
  return `${" ".repeat(leftPadding)}${clipped}${" ".repeat(rightPadding)}`;
}

function rightLine(text: string, width: number): string {
  if (width <= 0) {
    return "";
  }

  const clipped = truncate(text, width);
  return `${" ".repeat(Math.max(0, width - displayWidth(clipped)))}${clipped}`;
}

function buildEntryLabel(entry: EntryRow): string {
  return `#${entry.id} ${entry.vocabulary}`;
}

function getPageCount(totalEntries: number): number {
  return Math.max(1, Math.ceil(totalEntries / PAGE_SIZE));
}

function clampPageIndex(pageIndex: number, totalEntries: number): number {
  return Math.max(0, Math.min(pageIndex, getPageCount(totalEntries) - 1));
}

type VocabularyTableLayout = { border: string; header: string; columns: Array<Pick<MetadataAttribute, "key" | "label" | "role" | "displayOrder">>; widths: number[]; entries: EntryRow[] };

function buildTableLayout(entries: EntryRow[], pageIndex: number, width: number, availableRows: number, vocabularyLabel: string, meaningLabel: string, attributes?: MetadataAttribute[]): VocabularyTableLayout {
  const totalWidth = Math.max(width, 40);
  const innerWidth = Math.max(20, totalWidth - 2);
  const columns = (attributes?.filter((a) => a.visible && isActiveAttribute(a)) ?? [
    { key: "vocab", label: vocabularyLabel, role: "vocabulary" as const, displayOrder: 0 },
    { key: "meaning_1", label: meaningLabel, role: "meaning" as const, displayOrder: 1 },
  ]).map((a) => ({ key: a.key, label: a.label, role: a.role, displayOrder: a.displayOrder }));
  const columnWidth = Math.max(12, Math.floor((innerWidth - columns.length + 1) / columns.length));
  const widths = columns.map((_c, i) => i === columns.length - 1 ? innerWidth - (columnWidth + 1) * (columns.length - 1) : columnWidth);
  const border = `+${widths.map((w) => "─".repeat(w)).join("+")}+`;
  const visibleRows = Math.max(1, Math.min(PAGE_SIZE, availableRows));
  const pageEntries = entries.slice(pageIndex * PAGE_SIZE, pageIndex * PAGE_SIZE + visibleRows);
  const header = `|${columns.map((c, i) => padLine(c.label, widths[i])).join("|")}|`;
  return { border, header, columns, widths, entries: pageEntries };
}

function VocabularyTableCell({ entry, width, tagTypes, selected = false }: { entry: EntryRow; width: number; tagTypes: TagType[]; selected?: boolean }): JSX.Element {
  const tagNames = visibleAssignedTagGroups(entry, tagTypes).flatMap((group) => group.tagNames);
  const minimumVocabularyWidth = Math.min(width, 8);
  const badges = fitTagBadges(tagNames, Math.max(0, width - minimumVocabularyWidth - 1), displayWidth);
  if (badges.length === 0) return <>{padLine(buildEntryLabel(entry), width)}</>;
  const badgeWidth = badges.reduce((total, badge) => total + displayWidth(badge), 0) + badges.length - 1;
  const vocabularyWidth = Math.max(1, width - badgeWidth - 1);
  const vocabulary = truncate(buildEntryLabel(entry), vocabularyWidth);
  const gap = Math.max(1, width - displayWidth(vocabulary) - badgeWidth);
  return <><Text>{vocabulary}{" ".repeat(gap)}</Text>{badges.map((badge, index) => <React.Fragment key={`${index}-${badge}`}>{index > 0 ? " " : ""}<Text color={selected ? undefined : TAG_BADGE_COLOR}>{badge}</Text></React.Fragment>)}</>;
}

function buildHelpText(): string {
  return ["Commands:", ...COMMANDS.map((command) => `/${command.name}  ${command.hint}`), "/edit <id> opens detail editing; Ctrl+E edits while viewing.", "↑/↓ navigates rows, search, and command; Enter opens the selected entry.", "Esc cancels forms.", "←/→ to change pages."].join("\n");
}

function tierColor(tier: EntryRow["tier"]): string {
  return tier === "gray" ? GRAY_TIER_COLOR : tier === "green" ? "green" : tier === "yellow" ? YELLOW_TIER_COLOR : "red";
}

function buildStatusLines(message: string, lineCount = 5): string[] {
  const visible = message.split("\n").map((line) => line.trimEnd()).slice(-lineCount);
  while (visible.length < lineCount) {
    visible.push("");
  }
  return visible;
}

function getCommandPrefix(buffer: string): string | null {
  if (!buffer.startsWith("/")) {
    return null;
  }

  const afterSlash = buffer.slice(1);
  if (afterSlash.includes(" ")) {
    return null;
  }

  return afterSlash.toLowerCase();
}

function buildCommandSuggestions(buffer: string): CommandSpec[] {
  const prefix = getCommandPrefix(buffer);
  if (prefix === null) {
    return [];
  }

  if (prefix === "") {
    return COMMANDS;
  }

  return COMMANDS.filter((command) => command.name.startsWith(prefix));
}

function buildSuggestionLine(command: CommandSpec, width: number): { commandText: string; hintText: string } {
  const totalWidth = Math.max(width, 40);
  const leftWidth = Math.max(12, Math.min(16, Math.floor(totalWidth * 0.22)));
  const rightWidth = Math.max(1, totalWidth - 3 - leftWidth);
  return {
    commandText: padLine(`/${command.name}`, leftWidth),
    hintText: padLine(command.hint, rightWidth),
  };
}

function buildSuggestionLines(suggestions: CommandSpec[], selectedIndex: number, width: number, scrollOffset = 0): Array<{ commandText: string; hintText: string }> {
  if (suggestions.length === 0) {
    return [
      { commandText: padLine("No matching commands.", width), hintText: "" },
      ...Array.from({ length: COMMAND_SUGGESTION_ROWS - 1 }, () => ({ commandText: "", hintText: "" })),
    ];
  }

  const rows = suggestions.slice(scrollOffset, scrollOffset + COMMAND_SUGGESTION_ROWS).map((command) =>
    buildSuggestionLine(command, width),
  );
  while (rows.length < COMMAND_SUGGESTION_ROWS) {
    rows.push({ commandText: "", hintText: "" });
  }
  return rows;
}
function VocabularyScreen({ workbook, initialListContext, onBackToMenu, onQuit, onOpenSettings, onOpenTags, onOpenProfile, onAddEntry, onViewEntry, onStartPractice }: VocabularyScreenProps): JSX.Element {
  const { stdout } = useStdout();
  const [width, setWidth] = useState(() => stdout?.columns ?? 80);
  const [rows, setRows] = useState(() => stdout?.rows ?? 24);
  const [orchestration, setOrchestration] = useState(() => {
    const initial = createInitialState(backend.getWorkbook(workbook.id) ?? workbook, orchestrationCapabilities.reads.listEntries(workbook.id), orchestrationCapabilities.reads.listTagTypes(workbook.id));
    initial.statusLines = buildStatusLines("/ for commands, Esc to exit.");
    if (!initialListContext) return initial;
    const visible = filterEntriesForSearch(initial.entries, initial.workbook, initial.tagTypes, initialListContext.activeSearchQuery);
    return { ...initial, activeSearchQuery: initialListContext.activeSearchQuery, commandBuffer: initialListContext.commandBuffer, ...restoreListPosition(visible, initialListContext) };
  });
  const { entries, pageIndex, mode, commandBuffer: buffer, activeSearchQuery, inputFocus, selectedEntryId, statusLines } = orchestration;
  const orchestrationRef = useRef(orchestration);
  const [searchDraft, setSearchDraft] = useState(initialListContext?.searchDraft ?? "");
  const [suggestionIndex, setSuggestionIndex] = useState(0);
  const [tierColorsEnabled, setTierColorsEnabled] = useState(() => backend.getTierColorsEnabled());
  const [workbookProgressEnabled] = useState(() => backend.getWorkbookProgressEnabled());
  const [savedImportFilePath, setSavedImportFilePath] = useState(workbook.importFilePath);
  const tagTypes = orchestration.tagTypes;
  const visibleEntries = useMemo(() => filterEntriesForSearch(entries, workbook, tagTypes, activeSearchQuery), [entries, workbook, tagTypes, activeSearchQuery]);
  const textInputActive = (mode.kind === "command" && inputFocus === "command") || mode.kind === "commandArg" || mode.kind === "delete" || mode.kind === "importPath";
  const textInputKey = mode.kind;

  useEffect(() => {
    if (!stdout) {
      return;
    }

    const onResize = () => {
      setWidth(stdout.columns ?? 80);
      setRows(stdout.rows ?? 24);
    };

    stdout.on("resize", onResize);
    return () => {
      stdout.off("resize", onResize);
    };
  }, [stdout]);

  function syncOrchestration(next: ReturnType<typeof createInitialState>): void {
    orchestrationRef.current = next;
    setOrchestration(next);
  }

  function setBuffer(value: string): void {
    syncOrchestration({ ...orchestrationRef.current, commandBuffer: value });
  }

  function setStatusLines(lines: string[]): void {
    syncOrchestration({ ...orchestrationRef.current, statusLines: lines });
  }

  function openEntry(entryId: number, editing = false, consumeCommand = false): void {
    const current = orchestrationRef.current;
    onViewEntry(entryId, editing, {
      activeSearchQuery: current.activeSearchQuery, searchDraft,
      commandBuffer: consumeCommand ? "" : current.commandBuffer,
      pageIndex: current.pageIndex, selectedEntryId: current.selectedEntryId, inputFocus: current.inputFocus,
    });
  }

  function dispatchIntent(intent: import("./orchestration.js").Intent): void {
    const transition = updateOrchestration(orchestrationRef.current, intent);
    syncOrchestration(transition.state);
    void runTransitionEffects(transition.effects, transition.state, orchestrationCapabilities).then((results) => {
      for (const result of results) {
        dispatchIntent(result);
        if (result.type === "importPathSaved") setSavedImportFilePath(result.path);
      }
    });
  }

  const commandSuggestions = useMemo(() => buildCommandSuggestions(buffer), [buffer]);
  const commandSuggestionIndex = commandSuggestions.length === 0 ? 0 : Math.min(suggestionIndex, commandSuggestions.length - 1);
  const commandSuggestionScrollOffset = commandSuggestions.length <= COMMAND_SUGGESTION_ROWS
    ? 0
    : Math.min(
        Math.max(0, commandSuggestionIndex - COMMAND_SUGGESTION_ROWS + 1),
        commandSuggestions.length - COMMAND_SUGGESTION_ROWS,
      );
  const commandPaletteActive = mode.kind === "command" && inputFocus === "command" && getCommandPrefix(buffer) !== null;
  const suggestionLines = useMemo(
    () => buildSuggestionLines(commandSuggestions, commandSuggestionIndex, width, commandSuggestionScrollOffset),
    [commandSuggestions, commandSuggestionIndex, commandSuggestionScrollOffset, width],
  );

  useEffect(() => {
    if (mode.kind !== "command") {
      setSuggestionIndex(0);
      return;
    }

    setSuggestionIndex(0);
  }, [buffer, mode.kind, commandSuggestions.length]);

  function applySearch(query: string): void {
    const normalized = query.trim();
    setSearchDraft(normalized);
    dispatchIntent({ type: "searchChanged", query: normalized });
  }

  async function reloadImportPreview(mode: Extract<UiMode, { kind: "importPreview" }>): Promise<void> {
    dispatchIntent({ type: "importPathLoadRequested", path: mode.pathBuffer, alternate: true });
  }

  function continueImportAfterSaveChoice(mode: Extract<UiMode, { kind: "importSaveDefault" }>): void {
    dispatchIntent({ type: "importSavePathDecision", save: mode.save });
  }

  function confirmImport(mode: Extract<UiMode, { kind: "importPreview" }>): void {
    if (mode.preview.entries.length > 0) dispatchIntent({ type: "importConfirmed" });
  }

  function beginPendingCommand(command: ParameterizedCommand): void {
    dispatchIntent({ type: "commandSubmitted", raw: `/${command}` });
  }

  function cancelActiveMode(message = "Cancelled."): void {
    dispatchIntent({ type: "modeCancelled", message });
    setSuggestionIndex(0);
  }

  function submitHighlightedCommand(): void {
    const highlighted = commandSuggestions[commandSuggestionIndex];
    if (!highlighted) {
      return;
    }

    setBuffer("");
    setSuggestionIndex(0);
    submitCommand(`/${highlighted.name}`);
  }

  function autocompleteHighlightedCommand(): void {
    const highlighted = commandSuggestions[commandSuggestionIndex];
    if (!highlighted) {
      return;
    }

    setBuffer(`/${highlighted.name} `);
  }

  function submitCommand(raw: string): void {
    const parts = normalizeCommand(raw);
    if (parts.length === 0) {
      return;
    }

    const [command, ...args] = parts;
    const lower = command.toLowerCase();

    if (lower === "setting" || lower === "settings") { onOpenSettings(); return; }
    if (lower === "tag" || lower === "tags") { onOpenTags(); return; }
    if (lower === "profile") {
      if (args.length > 0) { setStatusLines(buildStatusLines("Usage: /profile")); return; }
      onOpenProfile();
      return;
    }
    if (lower === "help") {
      setStatusLines(buildStatusLines(buildHelpText()));
      return;
    }

    if (lower === "view") {
      const entryId = Number(args[0]);
      if (!args[0] || !Number.isInteger(entryId)) { setStatusLines(buildStatusLines("Usage: /view <id>")); return; }
      const entry = orchestrationRef.current.entries.find((item) => item.id === entryId) ?? null;
      if (!entry || entry.workbookId !== workbook.id) { setStatusLines(buildStatusLines(`Entry #${entryId} was not found.`)); return; }
      openEntry(entry.id, false, true);
      return;
    }

    if (lower === "test") {
      if (args.length > 1 || (args[0] && (!/^\d+$/.test(args[0]) || Number(args[0]) <= 0))) { setStatusLines(buildStatusLines("Usage: /test [count]")); return; }
      onStartPractice(args[0] ? Number(args[0]) : undefined);
      return;
    }

    if (lower === "list") {
      dispatchIntent({ type: "commandSubmitted", raw });
      return;
    }

    if (lower === "search") {
      setBuffer("");
      if (args.length === 0) dispatchIntent({ type: "inputFocusChanged", focus: "search" });
      else applySearch(args.join(" "));
      return;
    }

    if (lower === "menu") {
      onBackToMenu();
      return;
    }

    if (lower === "add") {
      onAddEntry();
      return;
    }

    if (lower === "import") {
      if (args.length > 0) { setStatusLines(buildStatusLines("Usage: /import")); return; }
      dispatchIntent({ type: "commandSubmitted", raw: "/import" });
      return;
    }

    if (lower === "edit") {
      const entryId = Number(args[0]);
      if (!args[0] || Number.isNaN(entryId)) {
        beginPendingCommand("edit");
        return;
      }
      if (!Number.isInteger(entryId)) { setStatusLines(buildStatusLines("Usage: /edit <id>")); return; }
      const entry = orchestrationRef.current.entries.find((item) => item.id === entryId);
      if (!entry || entry.workbookId !== workbook.id) { setStatusLines(buildStatusLines(`Entry #${entryId} was not found.`)); return; }
      openEntry(entryId, true, true);
      return;
    }

    if (lower === "delete") {
      const entryId = Number(args[0]);
      if (!args[0] || Number.isNaN(entryId)) {
        beginPendingCommand("delete");
        return;
      }
      dispatchIntent({ type: "commandSubmitted", raw: `/delete ${entryId}` });
      return;
    }

    if (lower === "exit") {
      onQuit();
      return;
    }

    setStatusLines(buildStatusLines(`Unknown command: ${command}`));
  }

  function submitForm(value: string): void {
    const text = value.trim();
    if (mode.kind === "importPath") { dispatchIntent({ type: "importPathLoadRequested", path: value, alternate: mode.alternate }); return; }
    if (mode.kind === "importPreview") { dispatchIntent({ type: "importConfirmed" }); return; }
    if (mode.kind === "importLoading") return;
    if (mode.kind === "importSaveDefault") { dispatchIntent({ type: "importSavePathDecision", save: mode.save }); return; }
    if (mode.kind === "delete") { dispatchIntent({ type: "deleteConfirmationSubmitted", value: text }); }
  }
  useInput((input, key) => {
    if (key.ctrl && input === "c") {
      onQuit();
      return;
    }

    if (mode.kind === "listOrder") {
      if (mode.saving) return;
      if (key.escape) cancelActiveMode();
      else if (key.upArrow || key.downArrow) dispatchIntent({ type: "listOrderMoved", delta: key.upArrow ? -1 : 1 });
      else if (key.return) dispatchIntent({ type: "listOrderConfirmed" });
      return;
    }
    if (mode.kind === "importPending") return;
    if (mode.kind === "importConfirm") {
      if (key.escape) dispatchIntent({ type: "importPreviewReturned" });
      else if (key.upArrow || key.downArrow) dispatchIntent({ type: "importChoiceMoved", delta: key.upArrow ? -1 : 1 });
      else if (key.return) dispatchIntent({ type: "importChoiceConfirmed" });
      return;
    }

    if (key.escape) {
      if (mode.kind === "importSaveDefault") {
        continueImportAfterSaveChoice({ ...mode, save: false });
      } else if (mode.kind !== "command") {
        cancelActiveMode("Cancelled.");
      } else if (activeSearchQuery) {
        applySearch("");
      } else {
        onBackToMenu();
      }
      return;
    }

    if (mode.kind === "importSaveDefault") {
      if (key.leftArrow || key.upArrow || key.rightArrow || key.downArrow) dispatchIntent({ type: "importSavePathToggled" });
      else if (key.return) continueImportAfterSaveChoice(mode);
      return;
    }

    if (mode.kind === "importPath" && key.ctrl && input === "u") {
      setBuffer("");
      return;
    }

    if (mode.kind === "importPreview") {
      const focuses: ImportPreviewFocus[] = ["path", "filters", "records"];
      const pageSize = importPreviewPageSize(rows, Boolean(mode.error));
      const { pageIndex: safePreviewPage, pageCount } = paginateImportPreview(mode.preview, mode.filter, mode.pageIndex, pageSize);
      if (key.upArrow || key.downArrow) {
        const current = focuses.indexOf(mode.focus);
        const next = key.upArrow ? Math.max(0, current - 1) : Math.min(focuses.length - 1, current + 1);
        dispatchIntent({ type: "importPreviewFocusChanged", focus: focuses[next] });
      } else if (mode.focus === "path") {
        if (key.ctrl && input === "u") dispatchIntent({ type: "importPathBufferChanged", path: "" });
        else if (key.return && !mode.loading) void reloadImportPreview(mode);
      } else if (mode.focus === "filters" && (key.leftArrow || key.rightArrow)) {
        const filters: ImportPreviewFilter[] = ["records", "ready", "invalid", "duplicates"];
        const current = filters.indexOf(mode.filter);
        const next = key.leftArrow ? (current <= 0 ? filters.length - 1 : current - 1) : (current + 1) % filters.length;
        dispatchIntent({ type: "importPreviewFilterChanged", filter: filters[next] });
      } else if (mode.focus === "records" && key.leftArrow) dispatchIntent({ type: "importPreviewPageChanged", delta: -1 });
      else if (mode.focus === "records" && key.rightArrow) dispatchIntent({ type: "importPreviewPageChanged", delta: 1 });
      else if (key.return) confirmImport(mode);
      return;
    }

    if (mode.kind === "importLoading") return;

    if (mode.kind === "command" && (key.upArrow || key.downArrow) && !commandPaletteActive) {
      dispatchIntent({ type: "listFocusMoved", delta: key.upArrow ? -1 : 1 });
      return;
    }

    if (key.upArrow && commandPaletteActive) {
      setSuggestionIndex((current) => (current <= 0 ? commandSuggestions.length - 1 : current - 1));
      return;
    }

    if (key.downArrow && commandPaletteActive) {
      setSuggestionIndex((current) => (current >= commandSuggestions.length - 1 ? 0 : current + 1));
      return;
    }

    if ((key.leftArrow || key.rightArrow) && textInputActive && buffer.length > 0) {
      return;
    }

    if ((key.leftArrow || key.rightArrow) && mode.kind === "command" && inputFocus === "search" && searchDraft.length > 0) {
      return;
    }

    if (key.leftArrow && mode.kind === "command" && !commandPaletteActive) {
      dispatchIntent({ type: "pageChanged", delta: -1 });
      return;
    }

    if (key.rightArrow && mode.kind === "command" && !commandPaletteActive) {
      dispatchIntent({ type: "pageChanged", delta: 1 });
      return;
    }

    if (key.leftArrow || key.rightArrow || key.backspace || key.delete) {
      return;
    }

    if (key.tab) {
      if (commandPaletteActive && commandSuggestions.length > 0) {
        autocompleteHighlightedCommand();
      }
      return;
    }

    if (key.return) {
      const current = buffer;
      if (mode.kind === "command" && inputFocus === "table") {
        if (selectedEntryId !== null) openEntry(selectedEntryId);
      } else if (mode.kind === "command" && inputFocus === "search") {
        applySearch(searchDraft);
      } else if (commandPaletteActive && commandSuggestions.length > 0) {
        submitHighlightedCommand();
      } else if (mode.kind === "command" || mode.kind === "commandArg") {
        submitCommand(current);
      } else {
        submitForm(current);
      }
      return;
    }

  });

  const pageCount = getPageCount(visibleEntries.length);
  const safePageIndex = clampPageIndex(pageIndex, visibleEntries.length);
  const pageText = activeSearchQuery
    ? visibleEntries.length === 0 ? "0 matches" : `Page ${safePageIndex + 1}/${pageCount} | ${visibleEntries.length} matches`
    : `Page ${safePageIndex + 1}/${pageCount}`;
  const tableLayout = useMemo(
    () => buildTableLayout(visibleEntries, safePageIndex, width, PAGE_SIZE, workbook.vocabularyLabel, workbook.meaningAttributes[0]?.label ?? "Primary Meaning", workbook.metadataAttributes),
    [visibleEntries, safePageIndex, width, workbook.vocabularyLabel, workbook.meaningAttributes, workbook.metadataAttributes],
  );
  const screenTitle = `${TITLE} - ${workbook.name}`;
  const dailyProgress = buildWorkbookProgress(workbook);
  const searchBoxWidth = Math.min(34, Math.max(12, width - 56));
  const footerLeftWidth = Math.floor((width - searchBoxWidth) / 2);
  const footerRightWidth = Math.max(0, width - searchBoxWidth - footerLeftWidth);
  const footerHint = inputFocus === "table" ? "↑↓ rows | Enter view | ←/→ pages" : activeSearchQuery ? "↑↓ rows/inputs | ←/→ pages | Esc clears" : "↑↓ rows/inputs | ←/→ pages | Esc menu";

  if (mode.kind === "importPending") {
    return <Text color={AUXILIARY_TEXT_COLOR}>{mode.phase === "checking" ? "Checking import updates..." : "Importing entries..."}</Text>;
  }

  if (mode.kind === "importConfirm") {
    return <Box flexDirection="column">
      <Text color="cyan" bold>{centerLine("Unsaved import changes", width)}</Text>
      <Text color={AUXILIARY_TEXT_COLOR}>{padLine(`${mode.updated} existing ${mode.updated === 1 ? "vocabulary will" : "vocabularies will"} change. This choice applies to all updates.`, width)}</Text>
      <Text>{padLine("", width)}</Text>
      {IMPORT_EXIT_CHOICES.map((choice, index) => <Text key={choice} color={index === mode.choice ? "black" : "white"} backgroundColor={index === mode.choice ? "white" : undefined}>{padLine(`${index === mode.choice ? ">" : " "} ${choice}`, width)}</Text>)}
      <Text>{padLine("", width)}</Text>
      <Text color={AUXILIARY_TEXT_COLOR}>{padLine("↑↓ select | Enter confirm | Esc return to preview", width)}</Text>
    </Box>;
  }

  if (mode.kind === "importSaveDefault") {
    return (
      <Box flexDirection="column">
        <Text color="cyan" bold>{centerLine(`Import - ${workbook.name}`, width)}</Text>
        <Text color="green">{padLine(mode.resultMessage, width)}</Text>
        <Text color={AUXILIARY_TEXT_COLOR}>{padLine(`Imported from: ${truncate(mode.path, Math.max(1, width - 15))}`, width)}</Text>
        <Text>{padLine("", width)}</Text>
        <Text>{padLine("Save this file as the workbook default before returning?", width)}</Text>
        <Text color={!mode.save ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>{padLine(`${!mode.save ? ">" : " "} Don't save`, width)}</Text>
        <Text color={mode.save ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>{padLine(`${mode.save ? ">" : " "} Save`, width)}</Text>
        <Text>{padLine("", width)}</Text>
        <Text color={AUXILIARY_TEXT_COLOR}>{rightLine("↑↓ select | Enter confirms | Esc doesn't save", width)}</Text>
      </Box>
    );
  }

  if (mode.kind === "importPreview") {
    const visibleRows = importPreviewPageSize(rows, Boolean(mode.error));
    const filters: Array<{ key: ImportPreviewFilter; label: string }> = [
      { key: "records", label: "All" },
      { key: "ready", label: "Ready" },
      { key: "invalid", label: "Invalid" },
      { key: "duplicates", label: "Duplicates" },
    ];
    const statusColor = (status: "ready" | "invalid" | "duplicate") => status === "ready" ? "green" : status === "invalid" ? "red" : GRAY_TIER_COLOR;
    const { records: pageRecords, pageIndex: safePageIndex, pageCount } = paginateImportPreview(mode.preview, mode.filter, mode.pageIndex, visibleRows);
    const footer = mode.focus === "path"
      ? "↑↓ focus | Enter loads file | Ctrl+U clears | Esc cancels"
      : mode.focus === "filters"
        ? `${mode.preview.entries.length > 0 ? "Enter imports | " : ""}↑↓ focus | ←/→ filter | Esc cancels`
        : `${mode.preview.entries.length > 0 ? "Enter imports | " : ""}↑↓ focus | ←/→ pages | Esc cancels`;
    return (
      <Box flexDirection="column">
        <Text color="cyan" bold>{centerLine(`Import - ${workbook.name}`, width)}</Text>
        <Box flexDirection="column" borderStyle="single" borderColor={mode.focus === "path" ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR} paddingX={1}>
          <CaretInputLine value={mode.pathBuffer} onChange={(pathBuffer) => dispatchIntent({ type: "importPathBufferChanged", path: pathBuffer })} prefix={mode.loading ? "Loading: " : "File: "} width={Math.max(1, width - 4)} color={mode.focus === "path" ? "white" : AUXILIARY_TEXT_COLOR} focus={mode.focus === "path" && !mode.loading} inputKey={`import-preview-path-${mode.path}`} />
        </Box>
        {mode.error ? <Text color="red">{padLine(mode.error, width)}</Text> : null}
        <Text>{padLine("", width)}</Text>
        <Text color={SELECTED_TEXT_COLOR}>{padLine("▔".repeat(width), width)}</Text>
        <Box flexDirection="row" gap={3}>
          <Text color={SELECTED_TEXT_COLOR} bold={true}>Filters</Text>
          {filters.map((item) => <Text key={item.key} color={item.key === mode.filter && mode.focus === "filters" ? SELECTED_TEXT_COLOR : "white"} inverse={item.key === mode.filter}>{formatImportPreviewFilter(item.label, item.key === mode.filter)}</Text>)}
        </Box>
        <Text color={AUXILIARY_TEXT_COLOR}>{padLine(`Ignored fields: ${mode.preview.diagnostics.filter((item) => item.kind === "ignored-field").length} | Ignored tags: ${mode.preview.diagnostics.filter((item) => item.kind === "ignored-tag").length}`, width)}</Text>
        <Text>{padLine("", width)}</Text>
        <Box flexDirection="column" borderStyle="single" borderColor={mode.focus === "records" ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR} paddingX={1}>
          <Box justifyContent="space-between">
            <Text color={mode.focus === "records" ? SELECTED_TEXT_COLOR : "white"}>Preview</Text>
            <Text color={mode.focus === "records" ? SELECTED_TEXT_COLOR : "white"}>{importPreviewWordCount(mode.preview, mode.filter)}</Text>
          </Box>
          {pageRecords.length === 0 ? <Text color={AUXILIARY_TEXT_COLOR}>{padLine("No records in this category.", Math.max(1, width - 4))}</Text> : pageRecords.map((record) => <Text key={record.recordNumber} color={statusColor(record.status)}>{padLine(formatImportPreviewRecord(record), Math.max(1, width - 4))}</Text>)}
          {Array.from({ length: Math.max(0, visibleRows - Math.max(1, pageRecords.length)) }, (_, index) => <Text key={`blank-${index}`}>{padLine("", Math.max(1, width - 4))}</Text>)}
          <Text color={mode.focus === "records" ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>{rightLine(`Page ${safePageIndex + 1}/${pageCount}`, Math.max(1, width - 4))}</Text>
        </Box>
        <Text>{padLine("", width)}</Text>
        <Text color={AUXILIARY_TEXT_COLOR}>{rightLine(footer, width)}</Text>
      </Box>
    );
  }

  return (
    <Box flexDirection="column">
      <Text color="cyan" bold>
        {centerLine(screenTitle, width)}
      </Text>
      <Text>{tableLayout.border}</Text>
      <Text>{tableLayout.header}</Text>
      <Text>{tableLayout.border}</Text>
      {tableLayout.entries.map((entry) => {
        const selected = mode.kind === "command" && inputFocus === "table" && entry.id === selectedEntryId;
        return <Text key={entry.id} inverse={selected} color={!selected && tierColorsEnabled ? tierColor(entry.tier) : undefined}>
          |{tableLayout.columns.map((column, index) => <React.Fragment key={column.key}>{column.key === "vocab" ? <VocabularyTableCell entry={entry} width={tableLayout.widths[index]} tagTypes={tagTypes} selected={selected} /> : padLine(entryAttributeValue(entry, column), tableLayout.widths[index])}|</React.Fragment>)}
        </Text>
      })}
      {activeSearchQuery && tableLayout.entries.length === 0 ? (
        <Text color={AUXILIARY_TEXT_COLOR}>|{padLine(`No vocabulary matches "${activeSearchQuery}".`, Math.max(1, tableLayout.border.length - 2))}|</Text>
      ) : null}
      <Text>{tableLayout.border}</Text>
      <Text color={dailyProgress.color}>{centerLine(workbookProgressEnabled ? dailyProgress.label : "", width)}</Text>
      <Text color={dailyProgress.color}>{centerLine(workbookProgressEnabled ? dailyProgress.bar : "", width)}</Text>
      <Box flexDirection="row" width={width} alignItems="center">
        <Box width={footerLeftWidth}><Text color={AUXILIARY_TEXT_COLOR}>{padLine(pageText, footerLeftWidth)}</Text></Box>
        <Box width={searchBoxWidth} borderStyle="single" borderColor={inputFocus === "search" ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR} paddingX={1}>
          <CaretInputLine prefix="⌕ Search: " value={searchDraft} onChange={setSearchDraft} width={Math.max(1, searchBoxWidth - 4)} color="white" focus={mode.kind === "command" && inputFocus === "search"} inputKey="workbook-search" />
        </Box>
        <Box width={footerRightWidth}><Text color={AUXILIARY_TEXT_COLOR} italic>{rightLine(footerHint, footerRightWidth)}</Text></Box>
      </Box>
      <Text>{padLine("", width)}</Text>
      <Text color={inputFocus === "command" ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>{"─".repeat(width)}</Text>
      <CaretInputLine key={textInputKey} prefix="> " value={buffer} onChange={setBuffer} width={width} focus={textInputActive} inputKey={textInputKey} />
      <Text color={inputFocus === "command" ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>{"─".repeat(width)}</Text>
      {mode.kind === "listOrder" ? <>
        <Text bold color={SELECTED_TEXT_COLOR}>{padLine("Select a list order", width)}</Text>
        <Text>{padLine("", width)}</Text>
        {LIST_ORDER_OPTIONS.map((option, index) => <Text key={option.value} inverse={index === mode.choice} color={index === mode.choice ? undefined : "white"}>{padLine(`${index === mode.choice ? ">" : " "} ${option.label}`, width)}</Text>)}
        <Text>{padLine("", width)}</Text>
        <Text color={mode.error ? "red" : AUXILIARY_TEXT_COLOR} italic>{padLine(mode.error ?? (mode.saving ? "Saving list order..." : "↑↓ select | Enter apply | Esc cancel"), width)}</Text>
        <Text>{padLine("", width)}</Text>
      </> : null}
      {commandPaletteActive
        ? suggestionLines.map((line, index) => {
          const hasCommands = commandSuggestions.length > 0;
          const isActive = hasCommands && index + commandSuggestionScrollOffset === commandSuggestionIndex;
          return (
            <Text key={`suggestion-${index}-${line.commandText}-${line.hintText}`} inverse={isActive}>
              <Text bold={hasCommands} color={!hasCommands ? AUXILIARY_TEXT_COLOR : isActive ? undefined : "white"}>
                {hasCommands ? (isActive ? "> " : "  ") : ""}
                {line.commandText}
              </Text>
              <Text color={isActive ? undefined : AUXILIARY_TEXT_COLOR}>
                {" "}
                {line.hintText}
              </Text>
            </Text>
          );
        })
        : null}
      {commandPaletteActive ? <Text>{padLine("", width)}</Text> : null}
      {!commandPaletteActive && mode.kind !== "listOrder"
        ? statusLines.map((line, index) => (
            <Text key={`status-${index}-${line}`} color={AUXILIARY_TEXT_COLOR}>
              {padLine(line, width)}
            </Text>
          ))
        : null}
    </Box>
  );
}

function WorkbookMenuScreen({
  workbooks,
  selectedIndex,
  onSelectedIndexChange,
  onOpenWorkbook,
  onCreateWorkbook,
  onEditWorkbook,
  onDeleteWorkbook,
  onQuit,
}: {
  workbooks: WorkbookRow[];
  selectedIndex: number;
  onSelectedIndexChange: (index: number) => void;
  onOpenWorkbook: (workbook: WorkbookRow) => void;
  onCreateWorkbook: () => void;
  onEditWorkbook: (workbook: WorkbookRow) => void;
  onDeleteWorkbook: (workbook: WorkbookRow) => void;
  onQuit: () => void;
}): JSX.Element {
  const { stdout } = useStdout();
  const [width, setWidth] = useState(() => stdout?.columns ?? 80);

  useEffect(() => {
    if (!stdout) {
      return;
    }

    const onResize = () => {
      setWidth(stdout.columns ?? 80);
    };

    stdout.on("resize", onResize);
    return () => {
      stdout.off("resize", onResize);
    };
  }, [stdout]);

  const displayRows = useMemo(() => buildWorkbookMenuLines(workbooks, width, selectedIndex, padLine), [workbooks, selectedIndex, width]);

  useInput((input, key) => {
    if (key.ctrl && input === "c") {
      onQuit();
      return;
    }

    if (key.escape) {
      onQuit();
      return;
    }

    if (key.upArrow) {
      onSelectedIndexChange(selectedIndex <= 0 ? Math.max(0, workbooks.length) : selectedIndex - 1);
      return;
    }

    if (key.downArrow) {
      onSelectedIndexChange(selectedIndex >= Math.max(0, workbooks.length) ? 0 : selectedIndex + 1);
      return;
    }

    if (key.delete) {
      const selected = selectedIndex < workbooks.length ? workbooks[selectedIndex] : null;
      if (selected) {
        onDeleteWorkbook(selected);
      }
      return;
    }

    if (key.ctrl && input.toLowerCase() === "e") {
      const selected = selectedIndex < workbooks.length ? workbooks[selectedIndex] : null;
      if (selected) onEditWorkbook(selected);
      return;
    }

    if (key.return) {
      const selected = selectedIndex < workbooks.length ? workbooks[selectedIndex] : null;
      if (selected) {
        onOpenWorkbook(selected);
        return;
      }
      onCreateWorkbook();
    }
  });

  return (
    <Box flexDirection="column">
      <Text color="cyan" bold>
        {centerLine(TITLE, width)}
      </Text>
      <Text color={AUXILIARY_TEXT_COLOR}>{padLine("Select a workbook.", width)}</Text>
      <Text>{padLine("", width)}</Text>
      <Text>{displayRows.border}</Text>
      <Text>{displayRows.header}</Text>
      <Text>{displayRows.border}</Text>
      {displayRows.rows.map((row, index) => {
        const color = row.selected ? SELECTED_TEXT_COLOR : "white";
        const beforeProgress = sliceToWidth(row.beforeProgress, width);
        const progress = sliceToWidth(row.progress, Math.max(0, width - displayWidth(beforeProgress)));
        const afterProgress = sliceToWidth(row.afterProgress, Math.max(0, width - displayWidth(beforeProgress) - displayWidth(progress)));
        return (
          <Text key={`${index}-${row.line}`} color={color}>
            {beforeProgress}<Text color={row.progressColor}>{progress}</Text>{afterProgress}
          </Text>
        );
      })}
      <Text>{displayRows.border}</Text>
      <Box width={width} borderStyle="single" borderColor={selectedIndex === workbooks.length ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR} justifyContent="center">
        <Text color={selectedIndex === workbooks.length ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>
          {centerLine("+ Create new workbook", Math.max(0, width - 2))}
        </Text>
      </Box>
      <Text>{padLine("", width)}</Text>
      <Text color={AUXILIARY_TEXT_COLOR}>{rightLine(WORKBOOK_MENU_HINT, width)}</Text>
      <Text>{padLine("", width)}</Text>
      <Text color={AUXILIARY_TEXT_COLOR}>
        {padLine(
          selectedIndex < workbooks.length
            ? `Selected: ${workbooks[selectedIndex].name}`
            : "Create new workbook.",
          width,
        )}
      </Text>
    </Box>
  );
}

type CreateStage = "name" | "type" | "label" | "preset" | "pos" | "tags" | "meaning-count" | "meaning" | "attributes" | "confirm" | "destructive-confirm";

function WorkbookWizard({ existingWorkbook, onSave, onCancel, onQuit }: { existingWorkbook?: WorkbookRow; onSave: (input: WorkbookConfigurationInput, confirmDataLoss: boolean) => void; onCancel: () => void; onQuit: () => void }): JSX.Element {
  const { stdout } = useStdout();
  const [width, setWidth] = useState(() => stdout?.columns ?? 80);
  const [stage, setStage] = useState<CreateStage>("name");
  const [history, setHistory] = useState<CreateStage[]>([]);
  const initialTypeIndex = existingWorkbook ? Math.max(0, VOCABULARY_TYPES.findIndex((item) => item.kind === existingWorkbook.vocabularyKind && item.code === existingWorkbook.vocabularyLanguageCode)) : 0;
  const [name, setName] = useState(existingWorkbook?.name ?? "");
  const [typeIndex, setTypeIndex] = useState(initialTypeIndex);
  const [vocabularyLabel, setVocabularyLabel] = useState(existingWorkbook?.vocabularyLabel ?? "");
  const [presetAttributes, setPresetAttributes] = useState<MetadataAttribute[]>(() => existingWorkbook?.metadataAttributes.filter((field) => Boolean(field.presetKey)) ?? createPresetAttributes(VOCABULARY_TYPES[initialTypeIndex].code));
  const [presetLanguage, setPresetLanguage] = useState(VOCABULARY_TYPES[initialTypeIndex].code);
  const existingTagTypes = useMemo(() => existingWorkbook ? backend.listTagTypes(existingWorkbook.id) : [], [existingWorkbook]);
  const [addPosType, setAddPosType] = useState(() => existingWorkbook ? existingTagTypes.some((type) => type.name.toLocaleLowerCase() === "part of speech") : true);
  const [tags, setTags] = useState<Array<{ id?: number; name: string }>>(() => existingTagTypes.find((type) => type.name.toLocaleLowerCase() === "part of speech")?.tags ?? []);
  const [meaningCount, setMeaningCount] = useState(existingWorkbook?.meaningAttributes.length ?? 1);
  const [meanings, setMeanings] = useState<MeaningAttribute[]>(existingWorkbook?.meaningAttributes ?? [{ position: 1, label: "Primary Meaning", languageCode: null }]);
  const [meaningIndex, setMeaningIndex] = useState(0);
  const [meaningPalette, setMeaningPalette] = useState<number | null>(null);
  const [attributes, setAttributes] = useState<MetadataAttribute[]>(() => existingWorkbook?.metadataAttributes.filter((item) => item.role === "optional" && !item.presetKey) ?? []);
  const [selected, setSelected] = useState(0);
  const [editAction, setEditAction] = useState<"none" | "add" | "rename">("none");
  const [editBuffer, setEditBuffer] = useState("");
  const [error, setError] = useState("");
  const [confirmBuffer, setConfirmBuffer] = useState("");
  const [pendingConfiguration, setPendingConfiguration] = useState<WorkbookConfigurationInput | null>(null);
  const [destructiveFields, setDestructiveFields] = useState<Array<{ label: string; valueCount: number }>>([]);
  const selectedType = VOCABULARY_TYPES[typeIndex];
  const presetDefinition = selectedType.code ? LANGUAGE_PRESET_DEFINITIONS[selectedType.code] : undefined;

  const activeTextValue = stage === "destructive-confirm"
    ? confirmBuffer
    : editAction !== "none"
      ? editBuffer
      : stage === "name"
        ? name
        : stage === "label"
          ? vocabularyLabel
          : stage === "meaning"
            ? meanings[meaningIndex]?.label ?? ""
            : null;
  useEffect(() => { if (!stdout) return; const f = () => setWidth(stdout.columns ?? 80); stdout.on("resize", f); return () => { stdout.off("resize", f); }; }, [stdout]);

  function go(next: CreateStage): void { setHistory((items) => [...items, stage]); setStage(next); setError(""); setSelected(0); setEditAction("none"); setEditBuffer(""); }
  function back(): void { const previous = history.at(-1); if (!previous) return; setHistory((items) => items.slice(0, -1)); setStage(previous); setError(""); setEditAction("none"); }
  function optionalKey(label: string): string {
    const base = label.trim().toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "") || "field";
    const used = new Set([...presetAttributes, ...attributes].map((item) => item.key)); let key = base; let suffix = 2;
    while (used.has(key)) key = `${base}_${suffix++}`;
    return key;
  }
  function commitListEdit(): void {
    const label = editBuffer.trim(); if (!label) { setError("A name is required."); return; }
    if (editAction === "add") {
      if (stage === "tags") setTags((items) => items.some((item) => item.name.toLocaleLowerCase() === label.toLocaleLowerCase()) ? items : [...items, { name: label }]);
      else setAttributes((items) => [...items, { key: optionalKey(label), label, languageCode: null, required: false, visible: false, displayOrder: items.length }]);
    } else if (editAction === "rename") {
      if (stage === "tags") setTags((items) => items.map((item, index) => index === selected ? { ...item, name: label } : item));
      else if (stage === "preset") setPresetAttributes((items) => items.map((item, index) => index === selected ? { ...item, label } : item));
      else setAttributes((items) => items.map((item, index) => index === selected ? { ...item, label } : item));
    }
    setEditAction("none"); setEditBuffer(""); setError("");
  }
  function configuration(): WorkbookConfigurationInput {
    return { name: name.trim(), vocabularyKind: selectedType.kind, vocabularyLabel: vocabularyLabel.trim(), vocabularyLanguageCode: selectedType.code,
      presetAttributes, meaningAttributes: meanings.slice(0, meaningCount),
      optionalAttributes: attributes, tagTypes: existingWorkbook ? existingTagTypes : addPosType ? [{ name: "Part of Speech", visible: false, tags }] : [] };
  }
  function submit(config: WorkbookConfigurationInput, confirmed: boolean): void {
    try { onSave(config, confirmed); }
    catch (caught) {
      if (caught instanceof WorkbookDataLossError) {
        setPendingConfiguration(config); setDestructiveFields(caught.impact.populatedFields); setConfirmBuffer(""); setStage("destructive-confirm"); setError("");
      } else setError(caught instanceof Error ? caught.message : "Could not save workbook.");
    }
  }
  function advance(): void {
    if (stage === "name") { if (!name.trim()) { setError("Workbook name is required."); return; } go("type"); return; }
    if (stage === "type") {
      const typeChanged = !existingWorkbook || selectedType.kind !== existingWorkbook.vocabularyKind || selectedType.code !== existingWorkbook.vocabularyLanguageCode;
      if (typeChanged) {
        setVocabularyLabel(selectedType.kind === "preset_language" ? selectedType.label : "Vocabulary");
        if (!existingWorkbook) { setAddPosType(selectedType.kind !== "non_language");
          if (presetLanguage !== selectedType.code) { setPresetAttributes(createPresetAttributes(selectedType.code)); setPresetLanguage(selectedType.code); setTags([]); } }
      }
      if (selectedType.kind === "other_language") setAddPosType(true);
      if (selectedType.kind === "non_language") setAddPosType(false);
      go("label"); return;
    }
    if (stage === "label") {
      if (!vocabularyLabel.trim()) { setError("Vocabulary label is required."); return; }
      go(existingWorkbook ? "meaning-count" : selectedType.kind === "preset_language" ? "preset" : "meaning-count"); return;
    }
    if (stage === "preset") {
      go("pos"); return;
    }
    if (stage === "pos") {
      if (!existingWorkbook && addPosType && selectedType.kind === "preset_language") setTags((current) => [...current, ...(presetDefinition?.partOfSpeechTags ?? []).filter((name) => !current.some((tag) => tag.name.toLocaleLowerCase() === name.toLocaleLowerCase())).map((tagName) => ({ name: tagName }))]);
      else if (!existingWorkbook && !addPosType) setTags([]);
      go(addPosType ? "tags" : "meaning-count"); return;
    }
    if (stage === "tags") { go("meaning-count"); return; }
    if (stage === "meaning-count") {
      const next = Array.from({ length: meaningCount }, (_, index) => meanings[index] ?? { position: index + 1, label: `Meaning ${index + 1}`, languageCode: null });
      setMeanings(next); setMeaningIndex(0); setMeaningPalette(null); go("meaning"); return;
    }
    if (stage === "meaning") {
      const label = meanings[meaningIndex]?.label.trim(); if (!label) { setError("Meaning label is required."); return; }
      if (new Set(meanings.map((item) => item.label.trim().toLocaleLowerCase())).size !== meanings.length) { setError("Meaning labels must be unique."); return; }
      if (meaningIndex + 1 < meaningCount) { setMeaningIndex((index) => index + 1); setMeaningPalette(null); setError(""); return; }
      go("attributes"); return;
    }
    if (stage === "attributes") { go("confirm"); return; }
    if (stage === "confirm") {
      submit(configuration(), false);
    }
  }

  useInput((input, key) => {
    if (key.ctrl && input === "c") return onQuit();
    if (key.escape) return onCancel();
    if (stage === "destructive-confirm") {
      if (key.return) {
        if (confirmBuffer.trim().toLowerCase() !== "yes") setError("Type yes to confirm removal of populated fields.");
        else if (pendingConfiguration) submit(pendingConfiguration, true);
      }
      return;
    }
    if ((key.leftArrow || key.rightArrow) && activeTextValue !== null && activeTextValue.length > 0) return;
    if (key.leftArrow) { if (editAction !== "none") { setEditAction("none"); setEditBuffer(""); } else if (stage === "meaning" && meaningIndex > 0) { setMeaningIndex((index) => index - 1); setMeaningPalette(null); } else back(); return; }
    if (key.rightArrow) { if (editAction === "none") advance(); return; }
    if (editAction !== "none") {
      if (key.return) commitListEdit();
      return;
    }
    if (!existingWorkbook && stage === "type" && (key.upArrow || key.downArrow)) { setTypeIndex((value) => key.upArrow ? (value <= 0 ? VOCABULARY_TYPES.length - 1 : value - 1) : (value + 1) % VOCABULARY_TYPES.length); return; }
    if (stage === "preset") {
      if (key.upArrow || key.downArrow) { setSelected((value) => Math.max(0, Math.min(presetAttributes.length - 1, value + (key.upArrow ? -1 : 1)))); return; }
      if (key.tab || input === " ") {
        const activation = key.tab;
        setPresetAttributes((items) => items.map((field, index) => index !== selected ? field : activation ? { ...field, active: !isActiveAttribute(field) } : { ...field, visible: !field.visible })); return;
      }
      if (key.ctrl && input.toLowerCase() === "r" && presetAttributes[selected]) { setEditAction("rename"); setEditBuffer(presetAttributes[selected].label); return; }
      if ((key.ctrl && input.toLowerCase() === "a") || key.delete) { setError("Preset attributes are fixed. Tab toggles activation."); return; }
    }
    if (stage === "pos" && (key.upArrow || key.downArrow || input === " ")) { setAddPosType((value) => !value); return; }
    if (stage === "meaning-count" && (key.upArrow || key.downArrow)) { setMeaningCount((value) => key.upArrow ? Math.min(5, value + 1) : Math.max(1, value - 1)); return; }
    if (stage === "meaning" && (key.upArrow || key.downArrow)) {
      const next = meaningPalette === null ? 0 : key.upArrow ? (meaningPalette <= 0 ? LANGUAGE_PRESETS.length - 1 : meaningPalette - 1) : (meaningPalette + 1) % LANGUAGE_PRESETS.length;
      setMeaningPalette(next); const preset = LANGUAGE_PRESETS[next]; setMeanings((items) => items.map((item, index) => index === meaningIndex ? { ...item, label: `${preset.label} (${preset.code})`, languageCode: preset.code } : item)); return;
    }
    if (stage === "tags" || stage === "attributes") {
      const items = stage === "tags" ? tags : attributes;
      if (key.upArrow) { setSelected((value) => Math.max(0, value - 1)); return; }
      if (key.downArrow) { setSelected((value) => Math.min(Math.max(0, items.length - 1), value + 1)); return; }
      if (key.ctrl && input.toLowerCase() === "a") { setEditAction("add"); setEditBuffer(""); return; }
      if (key.ctrl && input.toLowerCase() === "r" && items.length) { setEditAction("rename"); setEditBuffer(stage === "tags" ? tags[selected].name : attributes[selected].label); return; }
      if (key.delete && items.length) { stage === "tags" ? setTags((values) => values.filter((_, i) => i !== selected)) : setAttributes((values) => values.filter((_, i) => i !== selected)); setSelected((value) => Math.max(0, value - 1)); return; }
    }
    if (key.return) { advance(); return; }
  });

  const question = ["name"].includes(stage) ? 1 : ["type", "label", "preset", "pos", "tags"].includes(stage) ? 2 : ["meaning-count", "meaning"].includes(stage) ? 3 : stage === "attributes" ? 4 : 4;
  const title = stage === "destructive-confirm" ? "Confirm data removal" : stage === "confirm" ? `Confirm ${existingWorkbook ? "changes" : "creation"}` : `Question ${question}/4`;
  const listItems = stage === "tags" ? tags.map((item) => item.name) : attributes.map((item) => item.label);
  const tagSummary = existingWorkbook ? existingTagTypes.map((type) => `${type.name} (${type.tags.length} tags)`).join(", ") || "None" : addPosType ? `Part of Speech (${tags.length} tags)` : "None";
  const summary = [`Name: ${name}`, `Vocabulary: ${vocabularyLabel} - ${selectedType.label}`, `Active preset attributes: ${presetAttributes.filter(isActiveAttribute).map((item) => item.label).join(", ") || "None"}`, `Tag types: ${tagSummary}`, `Meanings: ${meanings.map((item) => item.label).join(", ")}`, `Optional attributes: ${attributes.map((item) => item.label).join(", ") || "None"}`];
  let description = `Enter a name for the ${existingWorkbook ? "workbook" : "new workbook"}.`;
  let footer = "Enter next | Esc cancel";
  if (stage === "type") { description = existingWorkbook ? "The vocabulary type is fixed after workbook creation." : "Choose a vocabulary type with ↑↓. This selection is required."; footer = existingWorkbook ? "←/→ navigate questions | Esc cancel" : "↑↓ choose type | ←/→ navigate questions | Esc cancel"; }
  if (stage === "label") { description = "Choose the label users will see for vocabulary entries. You can edit the default."; footer = "←/→ to navigate questions | Enter next | Esc cancel"; }
  if (stage === "preset") { description = "Select exclusive attributes to be active or inactive. You can change them later in Settings."; footer = "↑↓ select | Tab active/inactive | Space show/hide | Ctrl+R rename | → next"; }
  if (stage === "pos") { description = "Choose whether to add the Part of Speech tag type and its preset tags. After creation, manage it in Settings."; footer = "↑↓/Space toggle | ←/→ to navigate questions | Esc cancel"; }
  if (stage === "tags") { description = "Review and customize the Part of Speech tags for this workbook."; footer = "↑↓ select | Ctrl+A add | Ctrl+R rename | Del delete | ←/→ navigate"; }
  if (stage === "meaning-count") { description = "Choose how many meaning fields each vocabulary entry will have."; footer = "↑↓ choose number | ←/→ to navigate questions | Esc cancel"; }
  if (stage === "meaning") { description = "Name each meaning field. Use a language preset or type your own label."; footer = "↑↓ choose preset | ←/→ to navigate questions | Esc cancel"; }
  if (stage === "attributes") { description = "Add any custom fields you want to store. Language preset fields are managed separately."; footer = "↑↓ select | Ctrl+A add | Ctrl+R rename | Del delete | ←/→ navigate"; }
  if (stage === "confirm") { description = `Review the workbook configuration before ${existingWorkbook ? "saving" : "creating"} it.`; footer = `← back | Enter ${existingWorkbook ? "save" : "create"} | Esc cancel`; }
  if (stage === "destructive-confirm") { description = "This change will permanently remove stored field values."; footer = "Type yes and press Enter | Esc cancel"; }
  return <Box flexDirection="column"><Text color="cyan" bold>{centerLine(title, width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine(description, width)}</Text><Text>{padLine("", width)}</Text>
    {stage === "name" ? <CaretInputLine key="workbook-name" prefix="Workbook name: " value={name} onChange={(value) => { setName(value); setError(""); }} width={width} inputKey="workbook-name" /> : null}
    {stage === "type" ? existingWorkbook ? <Text color={SELECTED_TEXT_COLOR}>{padLine(selectedType.label, width)}</Text> : VOCABULARY_TYPES.map((item, index) => <Text key={item.label} color={index === typeIndex ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>{padLine(`${index === typeIndex ? ">" : " "} ${item.label}`, width)}</Text>) : null}
    {stage === "label" ? <CaretInputLine key="vocabulary-label" prefix="Vocabulary label: " value={vocabularyLabel} onChange={(value) => { setVocabularyLabel(value); setError(""); }} width={width} inputKey="vocabulary-label" /> : null}
    {stage === "preset" ? <>{presetAttributes.map((field, index) => <Text key={field.key} color={index === selected ? SELECTED_TEXT_COLOR : "white"}><AttributeRowText text={padLine(`${index === selected ? ">" : " "} ${field.label} [${field.visible ? "shown" : "hidden"}]`, width)} inactiveLabel={isActiveAttribute(field) ? undefined : field.label} /></Text>)}{editAction === "rename" ? <CaretInputLine prefix="Rename: " value={editBuffer} onChange={setEditBuffer} width={width} inputKey={`preset-rename-${selected}`} /> : null}</> : null}
    {stage === "pos" ? <><Text color={addPosType ? "green" : AUXILIARY_TEXT_COLOR}>{padLine(`Add Part of Speech: ${addPosType ? "Yes" : "No"}`, width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine((presetDefinition?.partOfSpeechTags ?? []).length ? `Preset tags: ${(presetDefinition?.partOfSpeechTags ?? []).join(", ")}` : "No preset tags are currently defined for this language.", width)}</Text></> : null}
    {stage === "meaning-count" ? <Text>{padLine(`Meaning attributes: ${meaningCount}`, width)}</Text> : null}
    {stage === "meaning" ? <><CaretInputLine key={`meaning-${meaningIndex}`} prefix={`Meaning ${meaningIndex + 1}/${meaningCount}: `} value={meanings[meaningIndex]?.label ?? ""} onChange={(value) => { setMeaningPalette(null); setMeanings((items) => items.map((item, index) => index === meaningIndex ? { ...item, label: value, languageCode: null } : item)); setError(""); }} width={width} inputKey={`meaning-${meaningIndex}`} />{meaningPalette !== null ? LANGUAGE_PRESETS.map((item, index) => <Text key={item.code} color={index === meaningPalette ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>{padLine(`${index === meaningPalette ? ">" : " "} ${item.label} (${item.code})`, width)}</Text>) : null}</> : null}
    {(stage === "tags" || stage === "attributes") ? <>{listItems.length ? listItems.map((item, index) => <Text key={`${index}-${typeof item === "string" ? item : ""}`} color={index === selected ? SELECTED_TEXT_COLOR : "white"}>{padLine(`${index === selected ? ">" : " "} ${item}`, width)}</Text>) : <Text color={AUXILIARY_TEXT_COLOR}>{padLine(stage === "tags" ? "No tags." : "No optional attributes.", width)}</Text>}{editAction !== "none" ? <CaretInputLine key={`${stage}-${editAction}-${selected}`} prefix={`${editAction === "add" ? "Add" : "Rename"}: `} value={editBuffer} onChange={(value) => { setEditBuffer(value); setError(""); }} width={width} inputKey={`${stage}-${editAction}-${selected}`} /> : <Text>{padLine("", width)}</Text>}</> : null}
    {stage === "confirm" ? <>{summary.map((line) => <Text key={line}>{padLine(line, width)}</Text>)}<Text>{padLine("", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine(`Press Enter to ${existingWorkbook ? "save changes" : "create the workbook"}.`, width)}</Text></> : null}
    {stage === "destructive-confirm" ? <>{destructiveFields.map((field) => <Text key={field.label} color="red">{padLine(`${field.label}: ${field.valueCount} populated value(s)`, width)}</Text>)}<Text>{padLine("", width)}</Text><CaretInputLine key="workbook-destructive-confirm" prefix="Confirmation: " value={confirmBuffer} onChange={(value) => { setConfirmBuffer(value); setError(""); }} width={width} color="cyan" inputKey="workbook-destructive-confirm" /></> : null}
    <Text>{padLine("", width)}</Text><Text color="red">{padLine(error, width)}</Text><Text>{padLine("", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{rightLine(footer, width)}</Text></Box>;
}

function WorkbookDeleteConfirmScreen({
  workbook,
  onConfirm,
  onCancel,
  onQuit,
}: {
  workbook: WorkbookRow;
  onConfirm: () => void;
  onCancel: () => void;
  onQuit: () => void;
}): JSX.Element {
  const { stdout } = useStdout();
  const [width, setWidth] = useState(() => stdout?.columns ?? 80);
  const [confirm, setConfirm] = useState("");
  const [error, setError] = useState("");

  useEffect(() => {
    if (!stdout) {
      return;
    }

    const onResize = () => {
      setWidth(stdout.columns ?? 80);
    };

    stdout.on("resize", onResize);
    return () => {
      stdout.off("resize", onResize);
    };
  }, [stdout]);

  useInput((input, key) => {
    if (key.ctrl && input === "c") {
      onQuit();
      return;
    }

    if (key.escape) {
      onCancel();
      return;
    }

    if (key.return) {
      if (confirm.trim().toLowerCase() === "yes") {
        onConfirm();
        return;
      }
      setError("Type yes to delete.");
      return;
    }

  });

  return (
    <Box flexDirection="column">
      <Text color="cyan" bold>
        {centerLine("Delete workbook", width)}
      </Text>
      <Text color={AUXILIARY_TEXT_COLOR}>{padLine(`Workbook: ${workbook.name}`, width)}</Text>
      <Text>{padLine("", width)}</Text>
      <Text color={AUXILIARY_TEXT_COLOR}>{padLine(`Delete ${workbook.name}?`, width)}</Text>
      <Text>{padLine("", width)}</Text>
      <CaretInputLine prefix="> " value={confirm} onChange={(value) => { setConfirm(value); setError(""); }} width={width} color="cyan" inputKey="delete-workbook-confirm" />
      <Text>{padLine("", width)}</Text>
      <Text color={AUXILIARY_TEXT_COLOR}>{padLine(error || WORKBOOK_DELETE_HINT, width)}</Text>
    </Box>
  );
}

type AttributeSection = "vocabulary" | "meaning" | "preset" | "optional";
type AttributeSelection = { section: AttributeSection; fieldIndex: number | null };

function MetadataSettingsScreen({ onCancel, panel, draft }: { panel: SettingsPanelControls; draft: AttributeDraftControls; onCancel: () => void }): JSX.Element {
  const { stdout } = useStdout();
  const [width, setWidth] = useState(() => stdout?.columns ?? 80);
  const [height, setHeight] = useState(() => stdout?.rows ?? 24);
  const { vocabularyLabel, setVocabularyLabel, fields, setFields } = draft;
  const [selectedIndex, setSelectedIndex] = useState(0);
  const [editAction, setEditAction] = useState<"none" | "add" | "rename">("none");
  const [buffer, setBuffer] = useState("");
  const [message, setMessage] = useState("");
  const meanings = fields.filter((field) => field.role === "meaning");
  const presets = fields.filter((field) => Boolean(field.presetKey));
  const optional = fields.filter((field) => field.role === "optional" && !field.presetKey);
  const selections: AttributeSelection[] = [
    { section: "vocabulary", fieldIndex: null },
    ...fields.map((field, fieldIndex) => ({ section: field.presetKey ? "preset" as const : field.role as "meaning" | "optional", fieldIndex })),
    ...(optional.length === 0 ? [{ section: "optional" as const, fieldIndex: null }] : []),
  ];
  const safeSelectedIndex = Math.min(selectedIndex, Math.max(0, selections.length - 1));
  const selected = selections[safeSelectedIndex];
  const selectedField = selected?.fieldIndex == null ? null : fields[selected.fieldIndex];

  useEffect(() => { if (!stdout) return; const f = () => { setWidth(stdout.columns ?? 80); setHeight(stdout.rows ?? 24); }; stdout.on("resize", f); return () => { stdout.off("resize", f); }; }, [stdout]);
  useEffect(() => { if (selectedIndex !== safeSelectedIndex) setSelectedIndex(safeSelectedIndex); }, [selectedIndex, safeSelectedIndex]);

  function nextKey(section: "meaning" | "optional", label: string): string {
    const used = new Set(fields.map((field) => field.key));
    if (section === "meaning") {
      let suffix = 1; let key = `meaning_${suffix}`;
      while (used.has(key)) key = `meaning_${++suffix}`;
      return key;
    }
    const raw = label.trim().toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "") || "field";
    let key = raw; let suffix = 1;
    while (used.has(key) || key === "vocab") key = `${raw}_${++suffix}`;
    return key;
  }
  function labelsAreUnique(section: AttributeSection, label: string, ignoredField?: MetadataAttribute): boolean {
    return !fields.some((field) => field !== ignoredField && (section === "meaning" ? field.role === "meaning" : field.role === "optional") && field.label.trim().toLocaleLowerCase() === label.trim().toLocaleLowerCase());
  }
  function togglePreset(): void {
    if (selected.section !== "preset" || !selectedField) return;
    setFields((current) => current.map((field) => field === selectedField ? { ...field, active: !isActiveAttribute(field) } : field));
    setMessage("");
  }
  function beginAdd(): void {
    if (selected.section === "preset") { setMessage("Preset attributes are fixed. T toggles activation."); return; }
    if (selected.section === "vocabulary") { setMessage("Vocabulary cannot be added or removed."); return; }
    if (selected.section === "meaning" && meanings.length >= 5) { setMessage("A workbook can have at most five meaning attributes."); return; }
    setEditAction("add"); setBuffer(""); setMessage(`Enter a new ${selected.section === "meaning" ? "meaning" : "optional"} attribute name.`);
  }
  function beginRename(): void {
    setEditAction("rename"); setBuffer(selected.section === "vocabulary" ? vocabularyLabel : selectedField?.label ?? ""); setMessage("Enter a new display name.");
  }
  function promoteSelectedMeaning(): void {
    if (selected.section !== "meaning" || !selectedField || meanings[0] === selectedField) return;
    if (selectedField.id === undefined) { setMessage("Save this new meaning before making it the Primary Meaning."); return; }
    try {
      setFields((current) => {
        const promoted = current.find((field) => field.id === selectedField.id)!;
        const reorderedMeanings = [
          { ...promoted, required: false },
          ...current.filter((field) => field.role === "meaning" && field.id !== promoted.id).map((field) => ({ ...field, required: false })),
        ];
        return [...reorderedMeanings, ...current.filter((field) => field.role === "optional")];
      });
      setSelectedIndex(1);
      setMessage(`${selectedField.label} will become the Primary Meaning when changes are saved.`);
    } catch (caught) { setMessage(caught instanceof Error ? caught.message : "Could not check this meaning."); }
  }
  function commitEdit(): void {
    const label = buffer.trim();
    if (!label) { setMessage("Attribute name is required."); return; }
    if (editAction === "rename") {
      if (selected.section === "vocabulary") setVocabularyLabel(label);
      else if (selectedField) {
        if (!labelsAreUnique(selected.section, label, selectedField)) { setMessage("Attribute names must be unique within this section."); return; }
        setFields((current) => current.map((field) => field === selectedField ? { ...field, label } : field));
      }
    } else {
      const section = selected.section as "meaning" | "optional";
      if (!labelsAreUnique(section, label)) { setMessage("Attribute names must be unique within this section."); return; }
      const newField: MetadataAttribute = { key: nextKey(section, label), role: section, label, languageCode: null, required: false, visible: false, displayOrder: section === "meaning" ? meanings.length + 1 : optional.length + 1 };
      setFields((current) => section === "meaning" ? [...current.filter((field) => field.role === "meaning"), newField, ...current.filter((field) => field.role === "optional")] : [...current, newField]);
      setSelectedIndex(section === "meaning" ? 1 + meanings.length : 1 + meanings.length + presets.length + optional.length);
    }
    setEditAction("none"); setBuffer(""); setMessage("");
  }
  const focused = panel.focused;
  useInput((input, key) => {
    if (!focused) return;
    if (key.ctrl && input === "c") return;
    if (editAction !== "none") {
      if (key.escape) { setEditAction("none"); setBuffer(""); setMessage(""); return; }
      if (key.return) { commitEdit(); return; }
      return;
    }
    if (key.escape) { onCancel(); return; }
    if (key.upArrow) { if (safeSelectedIndex === 0) panel.onTop(); else setSelectedIndex((value) => Math.max(0, value - 1)); setMessage(""); return; }
    if (key.downArrow) { setSelectedIndex((value) => Math.min(selections.length - 1, value + 1)); setMessage(""); return; }
    if (key.ctrl && input.toLowerCase() === "p") { promoteSelectedMeaning(); return; }
    if (key.ctrl && input.toLowerCase() === "a") { beginAdd(); return; }
    if (key.ctrl && input.toLowerCase() === "r") { beginRename(); return; }
    if (key.tab) { togglePreset(); return; }
    if (input === " ") {
      if (selected.section === "vocabulary") { setMessage("Vocabulary is always shown."); return; }
      if (selectedField) setFields((current) => current.map((field) => field === selectedField ? { ...field, visible: !field.visible } : field));
      return;
    }
    if (key.delete) {
      if (selected.section === "preset") { setMessage("Preset attributes cannot be removed. Tab toggles activation."); return; }
      if (selected.section === "vocabulary") { setMessage("Vocabulary cannot be deleted."); return; }
      if (!selectedField) return;
      if (selected.section === "meaning" && meanings[0] === selectedField) { setMessage("The Primary Meaning cannot be deleted."); return; }
      setFields((current) => current.filter((field) => field !== selectedField)); setSelectedIndex((value) => Math.max(0, value - 1)); setMessage(`${selectedField.label} will be removed if changes are saved.`);
    }
  });

  const meaningFooter = meanings[0] === selectedField ? "↑↓ navigate | Ctrl+A add | Ctrl+R rename | Space show/hide | Esc leave" : `↑↓ navigate | Ctrl+A add | Ctrl+R rename${selectedField?.id === undefined ? "" : " | Ctrl+P make primary"} | Space show/hide | Del remove | Esc leave`;
  const footer = editAction !== "none" ? "Enter confirm | Esc cancel edit" : selected.section === "vocabulary" ? "↑↓ navigate | Ctrl+R rename | Esc leave" : selected.section === "meaning" ? meaningFooter : selected.section === "preset" ? "↑↓ select | Tab active/inactive | Ctrl+R rename | Space show/hide | Esc leave" : "↑↓ navigate | Ctrl+A add | Ctrl+R rename | Space show/hide | Del remove | Esc leave";
  useEffect(() => {
    panel.setNavigationHint(footer);
    return () => panel.setNavigationHint("");
  }, [panel.setNavigationHint, footer]);
  const lines: Array<{ text: string; color?: string; borderColor?: string; bold?: boolean; selected?: boolean; inactiveLabel?: string }> = [];
  function group(section: AttributeSection, title: string, description: string, members: Array<{ field: MetadataAttribute; index: number }>): void {
    const borderColor = focused && selected.section === section ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR;
    const innerWidth = Math.max(1, width - 4);
    const line = (text: string, color?: string, bold?: boolean, focused = false, inactiveLabel?: string) => lines.push({ text: padLine(text, innerWidth), color, borderColor, bold, selected: focused, inactiveLabel });
    lines.push({ text: `┌${"─".repeat(Math.max(1, width - 2))}┐`, color: borderColor });
    line(title, borderColor, true);
    line(description, AUXILIARY_TEXT_COLOR);
    if (section === "vocabulary") line(`${focused && selected.section === "vocabulary" ? ">" : " "} ${vocabularyLabel} [shown] (required)`, focused && selected.section === "vocabulary" ? SELECTED_TEXT_COLOR : "white", false, focused && selected.section === "vocabulary");
    else if (!members.length) line(`${focused && selected.section === section ? ">" : " "} No optional attributes`, AUXILIARY_TEXT_COLOR, false, focused && selected.section === section);
    else for (const { field, index } of members) {
      const rowFocused = focused && selected.fieldIndex === index;
      const color = rowFocused ? SELECTED_TEXT_COLOR : "white";
      line(`${rowFocused ? ">" : " "} ${field.label} [${field.visible ? "shown" : "hidden"}]${meanings[0] === field ? " (primary)" : ""}`, color, false, rowFocused, isActiveAttribute(field) ? undefined : field.label);
    }
    lines.push({ text: `└${"─".repeat(Math.max(1, width - 2))}┘`, color: borderColor });
  }
  const indexed = fields.map((field, index) => ({ field, index }));
  group("vocabulary", "Vocabulary", "The workbook type is fixed. This display name can be changed.", []);
  group("meaning", "Meanings", "Meaning values are optional. Space shows or hides a column. Up to five meanings.", indexed.filter(({ field }) => field.role === "meaning"));
  if (presets.length) group("preset", "Language Preset Attributes", "Rename, activate, or change visibility. Preset membership is fixed.", indexed.filter(({ field }) => Boolean(field.presetKey)));
  group("optional", "Optional Attributes", "Custom fields can be added, renamed, shown, hidden, or removed.", indexed.filter(({ field }) => field.role === "optional" && !field.presetKey));
  // Leave one terminal row free so Ink keeps its incremental-render cache current.
  const viewportSize = Math.max(1, height - 10 - (editAction !== "none" ? 1 : 0) - (!focused ? 1 : 0));
  const focusedLine = focused ? Math.max(0, lines.findIndex((line) => line.selected)) : 0;
  const firstLine = Math.max(0, Math.min(lines.length - viewportSize, focusedLine - Math.floor(viewportSize / 2)));
  return <Box flexDirection="column"><Text color={AUXILIARY_TEXT_COLOR}>{padLine("Configure field names, activation, and visibility. Changes are staged.", width)}</Text><Text color={message ? "red" : AUXILIARY_TEXT_COLOR}>{padLine(message, width)}</Text>
    {lines.slice(firstLine, firstLine + viewportSize).map((line, index) => <Text key={firstLine + index} color={line.borderColor ?? line.color}>{line.borderColor ? <>{"│ "}<Text color={line.color} bold={line.bold}><AttributeRowText text={line.text} inactiveLabel={line.inactiveLabel} /></Text>{" │"}</> : line.text}</Text>)}
    {editAction !== "none" ? <CaretInputLine key={`${editAction}-${selectedIndex}`} prefix={`${editAction === "add" ? "New attribute" : "Display name"}: `} value={buffer} onChange={(value) => { setBuffer(value); setMessage(""); }} width={width} focus={focused} color="cyan" inputKey={`${editAction}-${selectedIndex}`} /> : null}
    <Text color={AUXILIARY_TEXT_COLOR}>{padLine(`${firstLine > 0 ? "↑ more above" : ""}${firstLine + viewportSize < lines.length ? "  ↓ more below" : ""}`, width)}</Text></Box>;
}

type TagSelection = { kind: "type"; typeIndex: number } | { kind: "tag"; typeIndex: number; tagIndex: number } | { kind: "add-type" };

function tagDraftSignature(types: TagTypeDraft[]): string {
  return JSON.stringify(types.map((type) => ({ id: type.id, name: type.name, visible: type.visible, tags: type.tags.map((tag) => ({ id: tag.id, name: tag.name })) })));
}

function TagSettingsScreen({ workbook, onSave, onCancel, onQuit, panel, draft }: { panel?: SettingsPanelControls; draft?: TagDraftControls; workbook: WorkbookRow; onSave: (draft: WorkbookTagsDraft, confirmDataLoss: boolean) => void; onCancel: () => void; onQuit: () => void }): JSX.Element {
  const { stdout } = useStdout();
  const [width, setWidth] = useState(() => stdout?.columns ?? 80);
  const [height, setHeight] = useState(() => stdout?.rows ?? 24);
  const initialTypes = useMemo<TagTypeDraft[]>(() => backend.listTagTypes(workbook.id).map((type) => ({ id: type.id, name: type.name, visible: type.visible, tags: type.tags.map((tag) => ({ id: tag.id, name: tag.name })) })), [workbook.id]);
  const [localTypes, setLocalTypes] = useState<TagTypeDraft[]>(initialTypes);
  const { types, setTypes } = draft ?? { types: localTypes, setTypes: setLocalTypes };
  const [selectedIndex, setSelectedIndex] = useState(0);
  const [action, setAction] = useState<"none" | "add-type" | "add-tag" | "rename-type" | "rename-tag">("none");
  const [buffer, setBuffer] = useState("");
  const [screenMode, setScreenMode] = useState<"edit" | "exit-confirm" | "destructive-confirm">("edit");
  const [exitChoice, setExitChoice] = useState(0);
  const [confirmBuffer, setConfirmBuffer] = useState("");
  const [removals, setRemovals] = useState<TagDataLossError["impact"]["removals"]>([]);
  const [message, setMessage] = useState("");
  const selections: TagSelection[] = [...types.flatMap((type, typeIndex): TagSelection[] => [{ kind: "type", typeIndex }, ...type.tags.map((_tag, tagIndex): TagSelection => ({ kind: "tag", typeIndex, tagIndex }))]), { kind: "add-type" }];
  const safeSelectedIndex = Math.min(selectedIndex, selections.length - 1);
  const selected = selections[safeSelectedIndex];
  const selectedType = selected.kind === "add-type" ? undefined : types[selected.typeIndex];
  const selectedTag = selected.kind === "tag" ? selectedType?.tags[selected.tagIndex] : undefined;
  const dirty = tagDraftSignature(types) !== tagDraftSignature(initialTypes);
  useEffect(() => { if (!stdout) return; const f = () => { setWidth(stdout.columns ?? 80); setHeight(stdout.rows ?? 24); }; stdout.on("resize", f); return () => { stdout.off("resize", f); }; }, [stdout]);
  useEffect(() => { if (selectedIndex !== safeSelectedIndex) setSelectedIndex(safeSelectedIndex); }, [selectedIndex, safeSelectedIndex]);

  function begin(next: typeof action): void {
    setAction(next); setBuffer(next === "rename-type" ? selectedType?.name ?? "" : next === "rename-tag" ? selectedTag?.name ?? "" : ""); setMessage("");
  }
  function commitEdit(): void {
    const name = buffer.trim();
    if (!name) { setMessage(action.includes("type") ? "Tag type name is required." : "Tag name is required."); return; }
    if (action === "add-type") {
      if (types.some((type) => type.name.toLocaleLowerCase() === name.toLocaleLowerCase())) { setMessage("Tag type names must be unique."); return; }
      setTypes((current) => [...current, { name, visible: false, tags: [] }]);
    } else if (action === "add-tag" && selectedType) {
      if (selectedType.tags.some((tag) => tag.name.toLocaleLowerCase() === name.toLocaleLowerCase())) { setMessage(`Tag names in ${selectedType.name} must be unique.`); return; }
      setTypes((current) => current.map((type) => type === selectedType ? { ...type, tags: [...type.tags, { name }] } : type));
    } else if (action === "rename-type" && selectedType) {
      if (types.some((type) => type !== selectedType && type.name.toLocaleLowerCase() === name.toLocaleLowerCase())) { setMessage("Tag type names must be unique."); return; }
      setTypes((current) => current.map((type) => type === selectedType ? { ...type, name } : type));
    } else if (action === "rename-tag" && selectedType && selectedTag) {
      if (selectedType.tags.some((tag) => tag !== selectedTag && tag.name.toLocaleLowerCase() === name.toLocaleLowerCase())) { setMessage(`Tag names in ${selectedType.name} must be unique.`); return; }
      setTypes((current) => current.map((type) => type === selectedType ? { ...type, tags: type.tags.map((tag) => tag === selectedTag ? { ...tag, name } : tag) } : type));
    }
    setAction("none"); setBuffer(""); setMessage("");
  }
  function save(confirmDataLoss = false): void {
    try { onSave({ types }, confirmDataLoss); }
    catch (caught) {
      if (caught instanceof TagDataLossError) { setRemovals(caught.impact.removals); setConfirmBuffer(""); setScreenMode("destructive-confirm"); setMessage(""); }
      else { setScreenMode("edit"); setMessage(caught instanceof Error ? caught.message : "Could not save tags."); }
    }
  }

  const focused = panel?.focused ?? true;
  useInput((input, key) => {
    if (!focused) return;
    if (key.ctrl && input === "c") return onQuit();
    if (screenMode === "destructive-confirm") {
      if (key.escape) { setScreenMode("edit"); setConfirmBuffer(""); return; }
      if (key.return) { if (confirmBuffer.trim().toLowerCase() === "yes") save(true); else setMessage("Type yes to permanently remove assigned tag data."); return; }
      return;
    }
    if (screenMode === "exit-confirm") {
      if (key.escape) { setScreenMode("edit"); return; }
      if (key.upArrow) { setExitChoice((value) => value <= 0 ? 2 : value - 1); return; }
      if (key.downArrow) { setExitChoice((value) => (value + 1) % 3); return; }
      if (key.return) { if (exitChoice === 0) save(); else if (exitChoice === 1) onCancel(); else setScreenMode("edit"); } return;
    }
    if (action !== "none") {
      if (key.escape) { setAction("none"); setBuffer(""); setMessage(""); return; }
      if (key.return) { commitEdit(); return; }
      return;
    }
    if (key.escape) { if (panel) { onCancel(); return; } if (dirty) { setExitChoice(0); setScreenMode("exit-confirm"); } else onCancel(); return; }
    if (key.upArrow) { if (safeSelectedIndex === 0 && panel) panel.onTop(); else setSelectedIndex((value) => Math.max(0, value - 1)); setMessage(""); return; }
    if (key.downArrow) { setSelectedIndex((value) => Math.min(selections.length - 1, value + 1)); setMessage(""); return; }
    if (key.return && selected.kind === "add-type") { begin("add-type"); return; }
    if (input === " " && selected.kind === "type" && selectedType) {
      setTypes((current) => current.map((type) => type === selectedType ? { ...type, visible: !type.visible } : type));
      setMessage(`${selectedType.name} will be ${selectedType.visible ? "hidden" : "shown"} when changes are saved.`);
      return;
    }
    if (key.ctrl && input.toLowerCase() === "a" && selectedType) { begin("add-tag"); return; }
    if (key.ctrl && input.toLowerCase() === "r") { if (selected.kind === "type") begin("rename-type"); else if (selected.kind === "tag") begin("rename-tag"); return; }
    if (key.delete && selectedType) {
      if (selected.kind === "type") setTypes((current) => current.filter((type) => type !== selectedType));
      else if (selectedTag) setTypes((current) => current.map((type) => type === selectedType ? { ...type, tags: type.tags.filter((tag) => tag !== selectedTag) } : type));
      setSelectedIndex((value) => Math.max(0, value - 1)); setMessage("The selected item will be removed when changes are saved.");
    }
  });

  if (screenMode === "exit-confirm") {
    const choices = ["Save changes", "Discard changes", "Continue editing"];
    return <Box flexDirection="column"><Text color="cyan" bold>{centerLine("Unsaved tag changes", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine("Choose what to do with your changes.", width)}</Text><Text>{padLine("", width)}</Text>{choices.map((choice, index) => <Text key={choice} inverse={index === exitChoice} color={index === exitChoice ? undefined : "white"}>{padLine(`${index === exitChoice ? ">" : " "} ${choice}`, width)}</Text>)}<Text>{padLine("", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{rightLine("↑↓ select | Enter confirm | Esc continue editing", width)}</Text></Box>;
  }
  if (screenMode === "destructive-confirm") {
    return <Box flexDirection="column"><Text color="cyan" bold>{centerLine("Confirm tag data removal", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine("Saving will permanently remove these entry assignments:", width)}</Text><Text>{padLine("", width)}</Text>{removals.map((item, index) => <Text key={`${item.typeName}-${item.tagName}-${index}`} color="red">{padLine(`${item.typeName}${item.tagName ? ` / ${item.tagName}` : ""}: ${item.assignmentCount} assignment(s) across ${item.entryCount} entry/entries`, width)}</Text>)}<Text>{padLine("", width)}</Text><CaretInputLine prefix="Type yes: " value={confirmBuffer} onChange={(value) => { setConfirmBuffer(value); setMessage(""); }} width={width} color="cyan" inputKey="tag-destructive-confirm" /><Text color="red">{padLine(message, width)}</Text><Text>{padLine("", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{rightLine("Enter confirm | Esc continue editing", width)}</Text></Box>;
  }
  const footer = action !== "none" ? "Enter confirm | Esc cancel edit" : selected.kind === "add-type" ? "↑↓ navigate | Enter add tag type | Esc leave" : selected.kind === "type" ? "↑↓ navigate | Space show/hide | Ctrl+A add tag | Ctrl+R rename | Del remove | Esc leave" : "↑↓ navigate | Ctrl+A add tag | Ctrl+R rename | Del remove | Esc leave";
  useEffect(() => {
    if (!panel) return;
    panel.setNavigationHint(footer);
    return () => panel.setNavigationHint("");
  }, [panel?.setNavigationHint, footer]);
  const lines: Array<{ text: string; color?: string; selected?: boolean }> = [];
  const inner = Math.max(1, width - 4);
  function row(text: string, active = false, color = AUXILIARY_TEXT_COLOR): void {
    lines.push({ text: `│ ${padLine(text, inner)} │`, color: active ? SELECTED_TEXT_COLOR : color, selected: active });
  }
  types.forEach((type, typeIndex) => {
    const active = focused && selected.kind !== "add-type" && selected.typeIndex === typeIndex;
    const border = active ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR;
    lines.push({ text: `┌${"─".repeat(Math.max(1, width - 2))}┐`, color: border });
    row(`${active && selected.kind === "type" ? ">" : " "} ${type.name} [${type.visible ? "shown" : "hidden"}]`, active && selected.kind === "type", "white");
    if (!type.tags.length) row("No tags");
    type.tags.forEach((tag, tagIndex) => {
      const tagActive = focused && selected.kind === "tag" && selected.typeIndex === typeIndex && selected.tagIndex === tagIndex;
      row(`${tagActive ? ">" : " "} ${tag.name}`, tagActive);
    });
    lines.push({ text: `└${"─".repeat(Math.max(1, width - 2))}┘`, color: border });
  });
  lines.push({ text: `+${"-".repeat(Math.max(1, width - 2))}+`, color: AUXILIARY_TEXT_COLOR });
  row(`${focused && selected.kind === "add-type" ? ">" : " "} + Add tag type`, focused && selected.kind === "add-type");
  lines.push({ text: `+${"-".repeat(Math.max(1, width - 2))}+`, color: AUXILIARY_TEXT_COLOR });
  const viewport = Math.max(1, height - (panel ? 9 : 6) - (action !== "none" ? 1 : 0) - (panel && !focused ? 1 : 0));
  const focusedLine = Math.max(0, lines.findIndex((line) => line.selected));
  const firstLine = Math.max(0, Math.min(lines.length - viewport, focusedLine - Math.floor(viewport / 2)));
  return <Box flexDirection="column">{!panel ? <Text color="cyan" bold>{centerLine(`Tags - ${workbook.name}`, width)}</Text> : null}<Text color={AUXILIARY_TEXT_COLOR}>{padLine("Organize tags into types. Changes are staged until you leave Settings.", width)}</Text><Text color={message ? "red" : AUXILIARY_TEXT_COLOR}>{padLine(message, width)}</Text>
    {lines.slice(firstLine, firstLine + viewport).map((line, index) => <Text key={firstLine + index} color={line.color}>{line.text}</Text>)}
    {action !== "none" ? (<CaretInputLine key={`${action}-${safeSelectedIndex}`} prefix={`${action.includes("type") ? "Tag type" : "Tag"} name: `} value={buffer} onChange={(value) => { setBuffer(value); setMessage(""); }}
      width={width}
      focus={focused}
      color="cyan"
      inputKey={`${action}-${safeSelectedIndex}`}
    />
    ) : null}
    {!panel ? (
      <>
        <Text>{padLine("", width)}</Text>
        <Text color={AUXILIARY_TEXT_COLOR}>
          {rightLine(footer, width)}
        </Text>
      </>
    ) : null}
  </Box>;
}

function wrapDetailText(value: string, width: number): string[] {
  if (width <= 0) return [""];
  const chunks: string[] = [];
  let current = "";
  for (const char of Array.from(value)) {
    if (displayWidth(current + char) > width && current) {
      chunks.push(current);
      current = "";
    }
    current += char;
  }
  if (current || chunks.length === 0) chunks.push(current);
  return chunks;
}

function detailFieldLines(field: DetailField, width: number, keyWidth: number, valueColor = "white", keyColor = AUXILIARY_TEXT_COLOR): JSX.Element[] {
  const valueWidth = Math.max(1, width - keyWidth - 1);
  const styledLines = field.similarWords ? similarWordLines(field.similarWords, valueWidth, displayWidth)
    : field.spans ? wrapStyledText(field.spans, valueWidth, displayWidth) : null;
  if (styledLines) {
    return styledLines.map((spans, index) => {
      const keyPart = padLine(index === 0 ? `${field.key}:` : "", keyWidth);
      const content = spans.map((span) => span.text).join("");
      const remaining = Math.max(0, width - displayWidth(keyPart) - 1 - displayWidth(content));
      return <Text key={`${field.key}-${index}`}><Text color={keyColor}>{keyPart}</Text>{" "}{spans.map((span, spanIndex) => <Text key={spanIndex} color={span.highlight ? "cyan" : valueColor} bold={span.highlight}>{span.text}</Text>)}{" ".repeat(remaining)}</Text>;
    });
  }
  return wrapDetailText(field.value, valueWidth).map((chunk, index) => {
    const key = index === 0 ? `${field.key}:` : "";
    const keyPart = padLine(key, keyWidth);
    const remaining = Math.max(0, width - displayWidth(keyPart) - 1 - displayWidth(chunk));
    return <Text key={`${field.key}-${index}`}><Text color={keyColor}>{keyPart}</Text>{" "}<Text color={valueColor}>{chunk}</Text>{" ".repeat(remaining)}</Text>;
  });
}

function EntryDetails({ workbook, entry, tagTypes, width, showVocabulary = false, editor, onChange, viewportHeight }: {
  workbook: WorkbookRow; entry: EntryRow | null; tagTypes: TagType[]; width: number;
  showVocabulary?: boolean; editor?: DetailEditorState; onChange?: (value: string) => void; viewportHeight?: number;
}): JSX.Element {
  const displayedEntry = editor ? {
    ...entry, vocabulary: editor.draft.vocabulary, meanings: editor.draft.meanings, attributes: editor.draft.attributes,
    tags: tagTypes.flatMap((type) => type.tags).filter((tag) => editor.draft.tagIds.includes(tag.id)),
  } : entry;
  if (!displayedEntry) return <Box />;
  const sections = buildDetailSections(workbook, displayedEntry, tagTypes);
  const statusFields = buildReviewFields(entry);
  const editRows = detailEditRows(workbook, tagTypes);
  const selectedRow = editor ? editRows[editor.selected] : null;
  const fields = [...sections.meanings, ...sections.attributes, ...statusFields];
  const labels = [...fields.map((field) => field.key), ...(showVocabulary ? [workbook.vocabularyLabel] : []), ...(editor ? tagTypes.map((type) => type.name) : [])];
  const reviewKeyWidth = Math.max(...statusFields.map((field) => displayWidth(`${field.key}:`)));
  const keyWidth = Math.min(Math.max(12, ...labels.map((label) => displayWidth(`${editor ? "> *" : ""}${label}:`))), Math.max(1, Math.min(width - 2, Math.max(reviewKeyWidth, Math.floor(width * 0.4)))));
  const lines: Array<{ id: string; element: JSX.Element; selected?: boolean }> = [];
  function line(id: string, element: JSX.Element, selected = false): void { lines.push({ id, element, selected }); }
  function blank(id: string): void { line(id, <Text>{padLine("", width)}</Text>); }
  function heading(id: string, label: string): void { blank(`${id}-space`); line(id, <Text color="cyan" bold>{padLine(label, width)}</Text>); }
  function fieldLines(id: string, field: DetailField, row?: DetailEditRow, valueColor = "white"): void {
    const changed = Boolean(editor && row && detailRowChanged(editor, row, tagTypes));
    const selected = Boolean(editor && row && row.id === selectedRow?.id);
    const keyColor = selected ? SELECTED_TEXT_COLOR : changed ? YELLOW_TIER_COLOR : AUXILIARY_TEXT_COLOR;
    const label = `${editor && row ? `${selected ? "> " : "  "}${changed ? "*" : " "}` : ""}${field.key}`;
    if (selected && row && row.kind !== "tags" && editor && onChange) {
      line(id, <CaretInputLine prefix={`${padLine(`${label}:`, keyWidth)} `} value={detailRowValue(editor, row)} onChange={onChange} width={width} color={SELECTED_TEXT_COLOR} prefixColor={SELECTED_TEXT_COLOR} inputKey={`${entry?.id ?? "new"}-${row.id}`} />, true);
    } else if (selected) {
      line(id, <Text color={SELECTED_TEXT_COLOR}>{padLine(`${label}:`, keyWidth)}{" "}{padLine(truncate(field.value, Math.max(1, width - keyWidth - 1)), Math.max(1, width - keyWidth - 1))}</Text>, true);
    } else {
      detailFieldLines({ ...field, key: label }, width, keyWidth, valueColor, keyColor).forEach((element, index) => line(`${id}-${index}`, element));
    }
  }
  wrapDetailText(entry?.vocabulary ?? "New Vocabulary", Math.max(1, width - 2)).forEach((text, index) => line(`title-${index}`, <Text color="cyan" bold>{centerLine(text, width)}</Text>));
  blank("title-space");
  if (showVocabulary) {
    line("vocabulary-heading", <Text color="cyan" bold>{padLine("Vocabulary", width)}</Text>);
    fieldLines("vocabulary", { key: workbook.vocabularyLabel, value: displayedEntry.vocabulary || "—" }, editRows[0]);
    blank("vocabulary-space");
  }
  if (sections.meanings.length) {
    line("meanings", <Text color="cyan" bold>{padLine("Meanings", width)}</Text>);
    sections.meanings.forEach((field, index) => fieldLines(`meaning-${index}`, field, editRows.find((row) => row.kind === "meaning" && row.index === index)));
  }
  if (sections.attributes.length) {
    heading("attributes", "Attributes");
    const attributeRows = editRows.filter((row) => row.kind === "attribute");
    sections.attributes.forEach((field, index) => fieldLines(`attribute-${index}`, field, attributeRows[index]));
  }
  const tagGroups = editor ? tagTypes.map((type) => ({ typeId: type.id, typeName: type.name, tagNames: type.tags.filter((tag) => editor.draft.tagIds.includes(tag.id)).map((tag) => tag.name) })) : sections.tags;
  if (tagGroups.length) {
    heading("tags", "Tags");
    tagGroups.forEach((group) => fieldLines(`tags-${group.typeId}`, { key: group.typeName, value: group.tagNames.map((name) => `[${name}]`).join(" ") || "—" }, editRows.find((row) => row.kind === "tags" && row.typeId === group.typeId), TAG_BADGE_COLOR));
  }
  heading("review", "Review");
  statusFields.forEach((field, index) => fieldLines(`review-${index}`, field, undefined, field.key === "Tier" ? tierColor(entry?.tier ?? "gray") : "white"));
  blank("end-space");
  const viewport = Math.max(1, viewportHeight ?? lines.length);
  const focusedLine = Math.max(0, lines.findIndex((item) => item.selected));
  const firstLine = Math.max(0, Math.min(lines.length - viewport, focusedLine - Math.floor(viewport / 2)));
  return <Box flexDirection="column">
    {lines.slice(firstLine, firstLine + viewport).map((item) => <React.Fragment key={item.id}>{item.element}</React.Fragment>)}
    {viewportHeight !== undefined ? <Text color={AUXILIARY_TEXT_COLOR}>{padLine(`${firstLine > 0 ? "↑ more above" : ""}${firstLine + viewport < lines.length ? "  ↓ more below" : ""}`, width)}</Text> : null}
  </Box>;
}

function EntryViewScreen({ workbook, entryId, editing = false, onCreated, onNavigate, onCancel, onQuit }: { workbook: WorkbookRow; entryId: number | null; editing?: boolean; onCreated: (entryId: number) => void; onNavigate: (entryId: number) => void; onCancel: () => void; onQuit: () => void }): JSX.Element {
  const { stdout } = useStdout();
  const [size, setSize] = useState(() => ({ width: stdout?.columns ?? 80, height: stdout?.rows ?? 24 }));
  const { width, height } = size;
  const [boundaryMessage, setBoundaryMessage] = useState("");
  const [entries, setEntries] = useState(() => backend.listEntries(workbook.id));
  const [tagTypes] = useState(() => backend.listTagTypes(workbook.id));
  const [editor, setEditor] = useState(() => {
    if (entryId === null) return createAddDetailEditor(workbook);
    const entry = backend.getEntry(entryId);
    return entry && entry.workbookId === workbook.id ? createDetailEditor(entry, editing) : null;
  });
  const editorRef = useRef(editor);
  function replaceEditor(next: DetailEditorState): void { editorRef.current = next; setEditor(next); if (next.mode === "cancelled") onCancel(); }
  function dispatch(action: import("./detail-editor.js").DetailEditorAction): void {
    if (editorRef.current) replaceEditor(updateDetailEditor(editorRef.current, action, workbook, tagTypes));
  }
  useEffect(() => {
    if (!stdout) return;
    const resize = () => setSize({ width: stdout.columns ?? 80, height: stdout.rows ?? 24 });
    stdout.on("resize", resize); return () => { stdout.off("resize", resize); };
  }, [stdout]);
  useInput((input, key) => {
    if (key.ctrl && input === "c") { onQuit(); return; }
    const current = editorRef.current;
    if (!current) { if (key.escape) onCancel(); return; }
    const action = detailEditorKeyAction(current.mode, input, key);
    if (action === "confirmChoice") {
      const choice = detailExitChoices(current)[current.exitChoice];
      if (choice.action === "save") {
        const saved = saveDetailEditor(current, workbook, (draft, options) => current.original
          ? backend.updateEntry(current.original.id, draft.vocabulary, draft.meanings[0], draft.meanings, draft.attributes, draft.tagIds, options)
          : backend.addEntry(workbook.id, draft.vocabulary, draft.meanings[0], draft.meanings, draft.attributes, draft.tagIds));
        replaceEditor(saved);
        if (saved.mode === "view" && saved.original) {
          setEntries(backend.listEntries(workbook.id));
          if (!current.original) onCreated(saved.original.id);
        }
      } else dispatch({ type: choice.action });
      return;
    }
    if (action) { setBoundaryMessage(""); dispatch(action); return; }
    if (current.mode !== "view" || entryId === null) return;
    if (key.escape) { onCancel(); return; }
    if (key.leftArrow || key.rightArrow) {
      const direction = key.leftArrow ? "previous" : "next";
      const adjacent = adjacentEntryId(entries, entryId, direction);
      if (adjacent === null) setBoundaryMessage(direction === "previous" ? "Already at the first vocabulary." : "Already at the last vocabulary.");
      else { setBoundaryMessage(""); onNavigate(adjacent); }
    }
  });
  if (!editor) return <Box flexDirection="column"><Text color="red">{padLine("This vocabulary is no longer available.", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine("Esc returns to the vocabulary list.", width)}</Text></Box>;
  const entry = editor.original;
  if (editor.mode === "confirm") {
    return <Box flexDirection="column">
      <Text color="cyan" bold>{centerLine("Unsaved vocabulary changes", width)}</Text>
      <Text color={AUXILIARY_TEXT_COLOR}>{padLine("Choose what to do with your staged changes.", width)}</Text>
      <Text>{padLine("", width)}</Text>
      {detailExitChoices(editor).map((choice, index) => <Text key={choice.label} color={index === editor.exitChoice ? "black" : "white"} backgroundColor={index === editor.exitChoice ? "white" : undefined}>{padLine(`${index === editor.exitChoice ? ">" : " "} ${choice.label}`, width)}</Text>)}
      <Text>{padLine("", width)}</Text>
      <Text color={AUXILIARY_TEXT_COLOR}>{padLine("↑↓ select | Enter confirm | Esc continue editing", width)}</Text>
    </Box>;
  }
  if (editor.mode === "tags" && editor.picker) {
    const picker = editor.picker;
    const type = tagTypes.find((type) => type.id === picker.typeId)!;
    const viewport = Math.max(1, height - 6);
    const first = Math.max(0, Math.min(type.tags.length - viewport, picker.cursor - Math.floor(viewport / 2)));
    return <Box flexDirection="column">
      <Text color="cyan" bold>{centerLine(`${entry?.vocabulary ?? "New Vocabulary"} — ${type.name}`, width)}</Text>
      <Text color={AUXILIARY_TEXT_COLOR}>{padLine("Select tags. Changes remain staged until saved.", width)}</Text>
      <Text>{padLine("", width)}</Text>
      {type.tags.length ? type.tags.slice(first, first + viewport).map((tag, index) => <Text key={tag.id} color={first + index === picker.cursor ? "black" : "white"} backgroundColor={first + index === picker.cursor ? "white" : undefined}>{padLine(`${first + index === picker.cursor ? ">" : " "} [${picker.tagIds.includes(tag.id) ? "●" : " "}] ${tag.name}`, width)}</Text>) : <Text color={AUXILIARY_TEXT_COLOR}>{padLine("No tags configured for this type.", width)}</Text>}
      <Text color={AUXILIARY_TEXT_COLOR}>{padLine(`${first > 0 ? "↑ more above" : ""}${first + viewport < type.tags.length ? "  ↓ more below" : ""}`, width)}</Text>
      <Text color={AUXILIARY_TEXT_COLOR} italic>{padLine("↑↓ select | Space toggle | Enter stage and back | Esc cancel | Ctrl+S exit", width)}</Text>
    </Box>;
  }
  if (editor.mode === "edit") return <Box flexDirection="column">
    <EntryDetails workbook={workbook} entry={entry} tagTypes={tagTypes} width={width} showVocabulary editor={editor} onChange={(value) => dispatch({ type: "change", value })} viewportHeight={Math.max(1, height - 4)} />
    <Text color="red">{padLine(editor.error, width)}</Text>
    <Text color={AUXILIARY_TEXT_COLOR} italic>{padLine("↑↓ rows | Enter edit tags | Esc / Ctrl+S exit", width)}</Text>
  </Box>;
  if (!entry) return <Box />;
  const currentIndex = entries.findIndex((item) => item.id === entry.id);
  const previousEntry = currentIndex > 0 ? entries[currentIndex - 1] : null;
  const nextEntry = currentIndex >= 0 && currentIndex + 1 < entries.length ? entries[currentIndex + 1] : null;
  const navigationHalfWidth = Math.max(1, Math.floor(width / 2) - 2);
  const previousLabel = previousEntry ? `← ${truncate(previousEntry.vocabulary, Math.max(1, navigationHalfWidth - 2))}` : detailNavigationLabel(null, "previous");
  const nextLabel = nextEntry ? `${truncate(nextEntry.vocabulary, Math.max(1, navigationHalfWidth - 2))} →` : detailNavigationLabel(null, "next");
  return <Box flexDirection="column">
    <EntryDetails workbook={workbook} entry={entry} tagTypes={tagTypes} width={width} showVocabulary />
    {boundaryMessage ? <Text color={AUXILIARY_TEXT_COLOR}>{padLine(boundaryMessage, width)}</Text> : null}
    <Text color={AUXILIARY_TEXT_COLOR} italic>{centerLine("Ctrl+E edit | Esc returns to the vocabulary list", width)}</Text>
    <Box width={width} justifyContent="space-between"><Text>{previousLabel}</Text><Text>{nextLabel}</Text></Box>
  </Box>;
}

function PracticeTagLine({ typeName, tagNames, width }: { typeName: string; tagNames: string[]; width: number }): JSX.Element | null {
  const label = truncate(`${typeName}:`, Math.max(1, Math.floor(width * 0.4)));
  const badges = fitTagBadges(tagNames, Math.max(0, width - displayWidth(label) - 1), displayWidth);
  if (badges.length === 0) return null;
  return <Text><Text color={AUXILIARY_TEXT_COLOR}>{label} </Text>{badges.map((badge, index) => <React.Fragment key={`${index}-${badge}`}>{index > 0 ? " " : ""}<Text color={TAG_BADGE_COLOR}>{badge}</Text></React.Fragment>)}</Text>;
}

function PracticeScreen({ workbook, count, onCancel, onQuit, onDone }: { workbook: WorkbookRow; count: number; onCancel: () => void; onQuit: () => void; onDone: (score: number, total: number) => void }): JSX.Element {
  const { stdout } = useStdout();
  const [width, setWidth] = useState(() => stdout?.columns ?? 80);
  const [orchestration, setOrchestration] = useState(() => createInitialState(workbook, [], orchestrationCapabilities.reads.listTagTypes(workbook.id)));
  const controller = useRef<ReturnType<typeof createOrchestrationController> | null>(null);
  const practice = orchestration.practice;
  const tagTypes = orchestration.tagTypes;
  const current = practice.currentEntry;
  useEffect(() => { if (!stdout) return; const f = () => setWidth(stdout.columns ?? 80); stdout.on("resize", f); return () => { stdout.off("resize", f); }; }, [stdout]);
  useLayoutEffect(() => {
    const nextController = createOrchestrationController(
      createInitialState(workbook, [], orchestrationCapabilities.reads.listTagTypes(workbook.id)),
      orchestrationCapabilities, setOrchestration,
    );
    controller.current = nextController;
    void nextController.dispatch({ type: "practiceCandidatesRequested", workbookId: workbook.id, count });
    return () => { nextController.dispose(); };
  }, [workbook.id, count]);

  function dispatch(intent: import("./orchestration.js").Intent): void {
    void controller.current?.dispatch(intent);
  }

  function leave(callback: () => void): void {
    dispatch({ type: "practiceCancelled" });
    controller.current?.dispose();
    callback();
  }

  useInput((input, key) => {
    const latest = controller.current?.getState();
    if (latest?.practice.phase === "saving") return;
    if (key.ctrl && input === "c") return leave(onQuit);
    if (latest?.practice.phase === "detail") {
      if (key.return || key.escape) dispatch({ type: "practiceDetailContinued" });
      return;
    }
    if (key.escape) return leave(onCancel);
    if (!latest) return;
    const p = latest.practice;
    if (p.phase === "loading") return;
    if (p.phase === "saveFailed") { if (key.return) dispatch({ type: "practiceSaveRetried" }); return; }
    if (p.phase === "empty") { if (key.return) leave(onCancel); return; }
    if (p.phase === "done") { if (key.return) leave(() => onDone(p.score, p.candidates.length)); return; }
    if (p.feedback !== null) {
      if (key.return) dispatch({ type: p.feedback === "Incorrect!" ? "practiceDetailRequested" : "practiceRetryAdvanced" });
      else if (p.feedback === "Incorrect!" && (key.backspace || key.delete)) dispatch({ type: "practiceAnswerRetried" });
      else if (key.tab && !key.shift) dispatch({ type: p.feedback === "Incorrect!" ? "practiceAnswerRetried" : "practiceDetailRequested" });
      return;
    }
    if (key.tab) { if (!key.shift) dispatch({ type: "practiceMeaningCycled" }); return; }
    if (key.return) dispatch({ type: "practiceAnswerSubmitted", answer: latest.commandBuffer });
  });

  if (practice.phase === "saveFailed") return <Box flexDirection="column"><Text color="red">{practice.error}</Text><Text>Enter retries saving | Esc discards results</Text></Box>;
  if (practice.phase === "empty") return <Box flexDirection="column"><Text>{practice.error ?? `No entries available in ${workbook.name}.`}</Text><Text color={AUXILIARY_TEXT_COLOR}>Press Enter or Esc to return.</Text></Box>;
  if (practice.phase === "done") return <Box flexDirection="column"><Text color="cyan" bold>{centerLine(`Practice - ${workbook.name}`, width)}</Text><Text>{padLine("", width)}</Text><Text color="green">{padLine(`Final initial-round score: ${practice.score}/${practice.candidates.length}`, width)}</Text><Text>{padLine("Press Enter to return.", width)}</Text></Box>;
  if (practice.phase === "detail" && current) return <Box flexDirection="column">
    <EntryDetails workbook={workbook} entry={current} tagTypes={tagTypes} width={width} />
    <Text color={AUXILIARY_TEXT_COLOR}>{centerLine("Enter / Esc returns to practice", width)}</Text>
  </Box>;
  const roundLabel = practice.phase === "retry" ? `Retry round ${practice.retryNumber} - Question ${practice.index + 1}/${practice.retryRound.length}` : `Question ${practice.index + 1}/${practice.candidates.length}`;
  const visibleTagGroups = current ? visibleAssignedTagGroups(current, tagTypes) : [];
  const practiceHint = practice.feedback === "Incorrect!" ? "Tab / Backspace retry | Enter view details | Esc cancels" : practice.feedback === "Correct!" ? "Tab view details | Enter continues | Esc cancels" : "Tab cycle meanings | Enter submits | Esc cancels";
  return <Box flexDirection="column"><Text color="cyan" bold>{centerLine(`Practice - ${workbook.name}`, width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine(roundLabel, width)}</Text><Text>{padLine("", width)}</Text><Text>{padLine(`${workbook.meaningAttributes[practice.meaningIndex]?.label ?? "Primary Meaning"}: ${current?.meanings[practice.meaningIndex] || "-"}`, width)}</Text>{visibleTagGroups.map((group) => <PracticeTagLine key={group.typeId} typeName={group.typeName} tagNames={group.tagNames} width={width} />)}<Text>{padLine("", width)}</Text><CaretInputLine key={`${practice.phase}-${practice.index}-${practice.retryNumber}`} prefix="Answer: " value={orchestration.commandBuffer} onChange={(answer) => dispatch({ type: "practiceAnswerChanged", answer })} width={width} focus={practice.feedback === null} inputKey={`${practice.phase}-${practice.index}-${practice.retryNumber}`} /><Text>{padLine("", width)}</Text><Text color={practice.error || practice.feedback === "Incorrect!" ? "red" : "green"}>{padLine(practice.error ?? practice.feedback ?? "", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{rightLine(practiceHint, width)}</Text></Box>;
}

function PracticeProfileScreen({ workbook, onCancel, onQuit }: { workbook: WorkbookRow; onCancel: () => void; onQuit: () => void }): JSX.Element {
  const { stdout } = useStdout();
  const [width, setWidth] = useState(() => stdout?.columns ?? 80);
  const [today, setToday] = useState(() => localCalendarDay());
  const history = useMemo(() => backend.listDailyPractice(workbook.id), [workbook.id, today]);
  const profile = useMemo(() => buildPracticeProfile(history, width, today), [history, width, today]);
  useEffect(() => {
    const resize = () => setWidth(stdout?.columns ?? 80);
    stdout?.on("resize", resize);
    const timer = setInterval(() => setToday(localCalendarDay()), 30_000);
    return () => { stdout?.off("resize", resize); clearInterval(timer); };
  }, [stdout]);
  useInput((input, key) => {
    if (key.ctrl && input === "c") onQuit();
    else if (key.escape) onCancel();
  });
  const notes = [
    `Current streak: ${profile.streaks.current} days | Longest streak: ${profile.streaks.longest} days`,
    `Current daily goal: ${workbook.dailyGoal ?? "No goal"}${workbook.dailyGoal != null ? " words" : ""}`,
    `Any practice counts for the streak.`,
  ];
  return <Box flexDirection="column">
    <Text color="cyan" bold>{centerLine(`Profile - ${workbook.name}`, width)}</Text>
    <Text>{padLine("", width)}</Text>
    <Text color={AUXILIARY_TEXT_COLOR}>{padLine(profile.monthLabels, width)}</Text>
    {profile.rows.map((row) => <Text key={row.label}>
      <Text color={AUXILIARY_TEXT_COLOR}>{row.label}{" "}</Text>
      {row.cells.map((cell, index) => <React.Fragment key={cell.day}>{index > 0 ? " " : ""}<Text color={cell.color}>{cell.symbol}</Text></React.Fragment>)}
    </Text>)}
    <Text>{padLine("", width)}</Text>
    <Text><Text color={GRAY_TIER_COLOR}>· No practice</Text>{"  "}<Text color={YELLOW_TIER_COLOR}>█ Below goal</Text>{"  "}<Text color="green">█ Goal met</Text></Text>
    <Text>{padLine("", width)}</Text>
    {notes.map((note, index) => <Text key={note} color={index === 2 ? AUXILIARY_TEXT_COLOR : undefined}>{note}</Text>)}
    <Text>{padLine("", width)}</Text>
    <Text color={AUXILIARY_TEXT_COLOR} italic>{centerLine("Esc return", width)}</Text>
  </Box>;
}

function App(): JSX.Element {
  const { exit } = useApp();
  const [screen, setScreen] = useState<AppScreen>(() => ({ kind: "menu" }));
  const [workbooks, setWorkbooks] = useState<WorkbookRow[]>(() => backend.listWorkbooks());
  const [selectedIndex, setSelectedIndex] = useState(0);

  useEffect(() => {
    const next = backend.listWorkbooks();
    setWorkbooks(next);
    const currentWorkbookId = backend.getCurrentWorkbookId();
    if (currentWorkbookId !== null) {
      const currentIndex = next.findIndex((workbook) => workbook.id === currentWorkbookId);
      setSelectedIndex(currentIndex >= 0 ? currentIndex : (next.length > 0 ? 0 : 0));
    } else {
      setSelectedIndex(next.length > 0 ? 0 : 0);
    }

    return () => {
      leaveAlternateScreen();
      backend.close();
    };
  }, []);

  useEffect(() => {
    if (screen.kind !== "menu" && screen.kind !== "vocab") return;
    let displayedDay = localCalendarDay();
    const timer = setInterval(() => {
      const today = localCalendarDay();
      if (today !== displayedDay) {
        setWorkbooks(backend.listWorkbooks());
        setScreen((current) => {
          if (current.kind !== "vocab") return current;
          const workbook = backend.getWorkbook(current.workbook.id);
          return workbook ? { ...current, workbook } : current;
        });
        displayedDay = today;
      }
    }, 30_000);
    return () => clearInterval(timer);
  }, [screen.kind]);

  function openWorkbook(workbook: WorkbookRow): void {
    backend.setCurrentWorkbookId(workbook.id);
    setScreen({ kind: "vocab", workbook });
  }

  function backToMenu(): void {
    const nextWorkbooks = backend.listWorkbooks();
    setWorkbooks(nextWorkbooks);
    const currentWorkbookId = backend.getCurrentWorkbookId();
    if (currentWorkbookId !== null) {
      const currentIndex = nextWorkbooks.findIndex((workbook) => workbook.id === currentWorkbookId);
      setSelectedIndex(currentIndex >= 0 ? currentIndex : (nextWorkbooks.length > 0 ? 0 : 0));
    } else {
      setSelectedIndex(nextWorkbooks.length > 0 ? 0 : 0);
    }
    setScreen({ kind: "menu" });
  }

  function createWorkbook(): void {
    setScreen({ kind: "create-workbook" });
  }

  function editWorkbook(workbook: WorkbookRow): void {
    setScreen({ kind: "edit-workbook", workbook });
  }

  function handleCreateWorkbook(input: CreateWorkbookInput): void {
    const workbook = backend.createConfiguredWorkbook(input);
    backend.setCurrentWorkbookId(workbook.id);
    const nextWorkbooks = backend.listWorkbooks();
    setWorkbooks(nextWorkbooks);
    const nextIndex = nextWorkbooks.findIndex((item) => item.id === workbook.id);
    setSelectedIndex(nextIndex >= 0 ? nextIndex : (nextWorkbooks.length > 0 ? 0 : 0));
    setScreen({ kind: "menu" });
  }

  function handleUpdateWorkbook(workbookId: number, input: WorkbookConfigurationInput, confirmDataLoss: boolean): void {
    const updated = backend.updateConfiguredWorkbook(workbookId, input, confirmDataLoss);
    const nextWorkbooks = backend.listWorkbooks();
    setWorkbooks(nextWorkbooks);
    const nextIndex = nextWorkbooks.findIndex((item) => item.id === updated.id);
    setSelectedIndex(nextIndex >= 0 ? nextIndex : 0);
    setScreen({ kind: "menu" });
  }

  function handleDeleteWorkbook(workbook: WorkbookRow): void {
    const nextCurrentId = backend.deleteWorkbook(workbook.id);
    const nextWorkbooks = backend.listWorkbooks();
    setWorkbooks(nextWorkbooks);
    if (nextCurrentId === null) {
      setSelectedIndex(nextWorkbooks.length > 0 ? 0 : 0);
    } else {
      const nextIndex = nextWorkbooks.findIndex((item) => item.id === nextCurrentId);
      setSelectedIndex(nextIndex >= 0 ? nextIndex : (nextWorkbooks.length > 0 ? 0 : 0));
    }
    setScreen({ kind: "menu" });
  }

  function quit(): void {
    leaveAlternateScreen();
    backend.close();
    exit();
  }

  function refreshWorkbook(workbookId: number): WorkbookRow {
    return backend.getWorkbook(workbookId) ?? workbooks.find((w) => w.id === workbookId)!;
  }

  if (screen.kind === "vocab") {
    return <VocabularyScreen workbook={screen.workbook} initialListContext={screen.listContext} onBackToMenu={backToMenu} onQuit={quit} onOpenSettings={() => setScreen({ kind: "settings", workbook: refreshWorkbook(screen.workbook.id) })} onOpenTags={() => setScreen({ kind: "tags", workbook: refreshWorkbook(screen.workbook.id) })} onOpenProfile={() => setScreen({ kind: "profile", workbook: refreshWorkbook(screen.workbook.id) })} onAddEntry={() => setScreen({ kind: "add", workbook: screen.workbook })} onViewEntry={(entryId, editing, listContext) => setScreen({ kind: "view", workbook: screen.workbook, entryId, editing, listContext })} onStartPractice={(count) => { const n = count ?? 15; setScreen({ kind: "practice", workbook: refreshWorkbook(screen.workbook.id), count: Math.min(n, backend.countEntries(screen.workbook.id)) }); }} />;
  }

  if (screen.kind === "profile") return <PracticeProfileScreen workbook={screen.workbook} onCancel={() => setScreen({ kind: "vocab", workbook: refreshWorkbook(screen.workbook.id) })} onQuit={quit} />;

  if (screen.kind === "practice") return <PracticeScreen workbook={screen.workbook} count={screen.count} onCancel={() => setScreen({ kind: "vocab", workbook: refreshWorkbook(screen.workbook.id) })} onQuit={quit} onDone={() => setScreen({ kind: "vocab", workbook: refreshWorkbook(screen.workbook.id) })} />;

  if (screen.kind === "settings") return <SettingsScreen workbook={screen.workbook} backend={backend} onCancel={() => setScreen({ kind: "vocab", workbook: refreshWorkbook(screen.workbook.id) })} onQuit={quit}
    renderAttributes={(panel, draft, leave) => <MetadataSettingsScreen panel={panel} draft={draft} onCancel={leave} />}
    renderTags={(panel, draft, leave) => <TagSettingsScreen workbook={screen.workbook} panel={panel} draft={draft} onSave={() => {}} onCancel={leave} onQuit={quit} />} />;
  if (screen.kind === "tags") return <TagSettingsScreen workbook={screen.workbook} onSave={(draft, confirmDataLoss) => { backend.updateWorkbookTags(screen.workbook.id, draft, confirmDataLoss); setScreen({ kind: "vocab", workbook: refreshWorkbook(screen.workbook.id) }); }} onCancel={() => setScreen({ kind: "vocab", workbook: refreshWorkbook(screen.workbook.id) })} onQuit={quit} />;
  if (screen.kind === "view" || screen.kind === "add") return <EntryViewScreen key={screen.kind === "add" ? "new" : screen.entryId} workbook={screen.workbook} entryId={screen.kind === "add" ? null : screen.entryId} editing={screen.kind === "view" && screen.editing} onCreated={(entryId) => setScreen({ kind: "view", workbook: screen.workbook, entryId })} onNavigate={(entryId) => setScreen({ kind: "view", workbook: screen.workbook, entryId, editing: false, listContext: screen.kind === "view" ? screen.listContext : undefined })} onCancel={() => setScreen({ kind: "vocab", workbook: refreshWorkbook(screen.workbook.id), listContext: screen.kind === "view" ? screen.listContext : undefined })} onQuit={quit} />;

  if (screen.kind === "create-workbook") {
    return <WorkbookWizard onSave={(input) => handleCreateWorkbook(input)} onCancel={backToMenu} onQuit={quit} />;
  }

  if (screen.kind === "edit-workbook") {
    return (
      <WorkbookWizard
        existingWorkbook={screen.workbook}
        onSave={(input, confirmDataLoss) => handleUpdateWorkbook(screen.workbook.id, input, confirmDataLoss)}
        onCancel={backToMenu}
        onQuit={quit}
      />
    );
  }

  if (screen.kind === "delete-workbook") {
    return (
      <WorkbookDeleteConfirmScreen
        workbook={screen.workbook}
        onConfirm={() => handleDeleteWorkbook(screen.workbook)}
        onCancel={backToMenu}
        onQuit={quit}
      />
    );
  }

  return (
    <WorkbookMenuScreen
      workbooks={workbooks}
      selectedIndex={selectedIndex}
      onSelectedIndexChange={setSelectedIndex}
      onOpenWorkbook={openWorkbook}
      onCreateWorkbook={createWorkbook}
      onEditWorkbook={editWorkbook}
      onDeleteWorkbook={(workbook) => setScreen({ kind: "delete-workbook", workbook, confirm: "" })}
      onQuit={quit}
    />
  );
}

function run(): void {
  const interactive = Boolean(process.stdin?.isTTY && process.stdout?.isTTY);
  if (!interactive) {
    console.log(TITLE);
    console.log("Open this app in an interactive terminal to use the TUI.");
    backend.close();
    return;
  }

  enterAlternateScreen();
  render(<App />);
}

run();

