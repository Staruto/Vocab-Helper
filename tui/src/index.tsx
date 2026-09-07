import React, { useEffect, useMemo, useState } from "react";
import { Box, Text, render, useApp, useInput, useStdout } from "ink";
import { CreateWorkbookInput, EntryRow, LANGUAGE_PRESET_DEFINITIONS, MeaningAttribute, MetadataAttribute, PosTag, VocabularyKind, WorkbookAttributesDraft, WorkbookConfigurationInput, WorkbookDataLossError, WorkbookRow } from "./db.js";
import { VocabularyBackend } from "./backend.js";

type UiMode =
  | { kind: "command" }
  | { kind: "commandArg"; command: ParameterizedCommand }
  | { kind: "add"; stage: "vocabulary" | "meaning" | "metadata" | "pos"; vocabulary: string; meanings: string[]; meaningIndex: number; metadata: Record<string, string>; metadataIndex: number; selectedTagIds: number[]; posIndex: number }
  | { kind: "edit"; stage: "vocabulary" | "meaning" | "metadata" | "pos"; entryId: number; vocabulary: string; meanings: string[]; meaningIndex: number; metadata: Record<string, string>; metadataIndex: number; selectedTagIds: number[]; posIndex: number }
  | { kind: "delete"; entryId: number; label: string };

type AppScreen =
  | { kind: "menu" }
  | { kind: "create-workbook" }
  | { kind: "edit-workbook"; workbook: WorkbookRow }
  | { kind: "delete-workbook"; workbook: WorkbookRow; confirm: string }
  | { kind: "settings"; workbook: WorkbookRow }
  | { kind: "settings-attributes"; workbook: WorkbookRow }
  | { kind: "settings-pos"; workbook: WorkbookRow }
  | { kind: "settings-appearance"; workbook: WorkbookRow }
  | { kind: "tags"; workbook: WorkbookRow; returnTo?: "vocab" | "settings-pos" }
  | { kind: "view"; workbook: WorkbookRow; entry: EntryRow }
  | { kind: "practice"; workbook: WorkbookRow; count: number }
  | { kind: "vocab"; workbook: WorkbookRow };

type VocabularyScreenProps = {
  workbook: WorkbookRow;
  onBackToMenu: () => void;
  onQuit: () => void;
  onOpenSettings: () => void;
  onOpenTags: () => void;
  onViewEntry: (entry: EntryRow) => void;
  onStartPractice: (count?: number) => void;
};

type CommandSpec = {
  name: string;
  hint: string;
};

type ParameterizedCommand = "edit" | "delete";

type LanguagePreset = { code: string; label: string };

const PAGE_SIZE = 20;
const TITLE = "VocabHelper 3.0.0";
const FOOTER_HINT = "Navigate pages with <- -> | Esc returns to menu";
const AUXILIARY_TEXT_COLOR = "#979797";
const GRAY_TIER_COLOR = "#777777";
const SELECTED_TEXT_COLOR = "#cea8ff";
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
const WORKBOOK_MENU_HINT = "↑↓ select | Enter open | Ctrl+E edit | Del delete | + create | Esc exit";
const WORKBOOK_CREATE_HINT = "Type a name and press Enter. Esc returns to the menu.";
const WORKBOOK_DELETE_HINT = "Type yes to confirm. Enter deletes. Esc cancels.";
const COMMANDS: CommandSpec[] = [
  { name: "list", hint: "Refresh and show entries" },
  { name: "add", hint: "Add a new entry" },
  { name: "edit", hint: "Edit an entry by id" },
  { name: "delete", hint: "Delete an entry by id" },
  { name: "menu", hint: "Return to the workbook menu" },
  { name: "help", hint: "Show command help" },
  { name: "setting", hint: "Configure workbook fields" },
  { name: "tag", hint: "Manage part-of-speech tags" },
  { name: "view", hint: "View entry details and status" },
  { name: "test", hint: "Practice vocabulary (15 questions)" },
  { name: "exit", hint: "Exit the app" },
];
const backend = new VocabularyBackend();

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

function buildFooterLine(width: number, pageText: string, hintText: string): string {
  if (width <= 0) {
    return "";
  }

  const left = truncate(pageText, width);
  const right = truncate(hintText, width);
  const leftWidth = displayWidth(left);
  const rightWidth = displayWidth(right);
  if (leftWidth + rightWidth + 2 >= width) {
    const gap = Math.max(1, width - leftWidth);
    return `${left}${" ".repeat(gap)}${truncate(right, width - leftWidth - gap)}`;
  }

  return `${left}${" ".repeat(width - leftWidth - rightWidth)}${right}`;
}

function buildTableLines(entries: EntryRow[], pageIndex: number, width: number, availableRows: number, vocabularyLabel: string, meaningLabel: string, attributes?: MetadataAttribute[]): string[] {
  const totalWidth = Math.max(width, 40);
  const innerWidth = Math.max(20, totalWidth - 2);
  const columns = (attributes?.filter((a) => a.visible) ?? [
    { key: "vocab", label: vocabularyLabel }, { key: "meaning_1", label: meaningLabel },
  ]).map((a) => ({ key: a.key, label: a.label }));
  const columnWidth = Math.max(12, Math.floor((innerWidth - columns.length + 1) / columns.length));
  const widths = columns.map((_c, i) => i === columns.length - 1 ? innerWidth - (columnWidth + 1) * (columns.length - 1) : columnWidth);
  const border = `+${widths.map((w) => "-".repeat(w)).join("+")}+`;
  const valueFor = (entry: EntryRow, key: string): string => {
    if (key === "vocab") return buildEntryLabel(entry);
    if (key.startsWith("meaning_")) return entry.meanings[Number(key.slice("meaning_".length)) - 1] ?? "";
    return entry.attributes[key] ?? "";
  };
  const visibleRows = Math.max(1, Math.min(PAGE_SIZE, availableRows));
  const pageEntries = entries.slice(pageIndex * PAGE_SIZE, pageIndex * PAGE_SIZE + visibleRows);
  const header = `|${columns.map((c, i) => padLine(c.label, widths[i])).join("|")}|`;

  const rows = pageEntries.map((entry) => {
    return `|${columns.map((c, i) => padLine(valueFor(entry, c.key), widths[i])).join("|")}|`;
  });

  return [border, header, border, ...rows, border];
}

function buildHelpText(): string {
  return ["Commands:", ...COMMANDS.map((command) => `/${command.name}  ${command.hint}`), "Esc cancels forms.", "Use <- -> to change pages."].join("\n");
}

function buildPendingCommandText(command: ParameterizedCommand): string {
  return `Enter id for /${command}.`;
}

function tierColor(tier: EntryRow["tier"]): string {
  return tier === "gray" ? GRAY_TIER_COLOR : tier === "green" ? "green" : tier === "yellow" ? "yellow" : "red";
}

function formatLastTested(value: string | null): string {
  if (!value) return "Never";
  const date = new Date(value.replace(" ", "T") + (value.includes("Z") ? "" : "Z"));
  return Number.isNaN(date.getTime()) ? value.slice(0, 10) : date.toLocaleDateString();
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

function buildWorkbookMenuLines(
  workbooks: WorkbookRow[],
  width: number,
  selectedIndex: number,
): { border: string; header: string; rows: Array<{ line: string; selected: boolean }> } {
  const totalWidth = Math.max(width, 56);
  const selectWidth = 3;
  const countWidth = 8;
  const nameWidth = Math.max(24, totalWidth - selectWidth - countWidth - 4);
  const border = `+${"-".repeat(selectWidth)}+${"-".repeat(nameWidth)}+${"-".repeat(countWidth)}+`;
  const header = `|${padLine("", selectWidth)}|${padLine("Name", nameWidth)}|${padLine("Words", countWidth)}|`;

  const rows = workbooks.map((workbook, index) => ({
    selected: index === selectedIndex,
    line: `|${padLine(index === selectedIndex ? ">" : " ", selectWidth)}|${padLine(workbook.name, nameWidth)}|${padLine(String(workbook.wordCount), countWidth)}|`,
  }));

  rows.push({
    selected: selectedIndex === workbooks.length,
    line: `|${padLine(selectedIndex === workbooks.length ? ">" : " ", selectWidth)}|${padLine("+ Create new workbook", nameWidth)}|${padLine("", countWidth)}|`,
  });
  return { border, header, rows };
}

function buildWorkbookCreateLines(name: string, width: number): string[] {
  return [
    padLine(`Name: > ${name}_`, width),
    padLine("", width),
  ];
}

function buildWorkbookDeleteLines(workbook: WorkbookRow, confirm: string, width: number): string[] {
  return [
    padLine(`Delete ${workbook.name}?`, width),
    padLine("", width),
    padLine(`> ${confirm}_`, width),
    padLine("", width),
  ];
}

function VocabularyScreen({ workbook, onBackToMenu, onQuit, onOpenSettings, onOpenTags, onViewEntry, onStartPractice }: VocabularyScreenProps): JSX.Element {
  const { stdout } = useStdout();
  const [width, setWidth] = useState(() => stdout?.columns ?? 80);
  const [rows, setRows] = useState(() => stdout?.rows ?? 24);
  const [entries, setEntries] = useState<EntryRow[]>(() => backend.listEntries(workbook.id));
  const [pageIndex, setPageIndex] = useState(0);
  const [mode, setMode] = useState<UiMode>({ kind: "command" });
  const [buffer, setBuffer] = useState("");
  const [suggestionIndex, setSuggestionIndex] = useState(0);
  const [statusLines, setStatusLines] = useState<string[]>(() => buildStatusLines("Ready."));
  const [tierColorsEnabled, setTierColorsEnabled] = useState(() => backend.getTierColorsEnabled());
  const activeMetadata = workbook.metadataAttributes.filter((attribute) => attribute.role === "optional");
  const posTags = workbook.posEnabled ? backend.listPosTags(workbook.id) : [];

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

  useEffect(() => {
    setPageIndex((current) => clampPageIndex(current, entries.length));
  }, [entries.length]);

  const commandSuggestions = useMemo(() => buildCommandSuggestions(buffer), [buffer]);
  const commandSuggestionIndex = commandSuggestions.length === 0 ? 0 : Math.min(suggestionIndex, commandSuggestions.length - 1);
  const commandSuggestionScrollOffset = commandSuggestions.length <= COMMAND_SUGGESTION_ROWS
    ? 0
    : Math.min(
        Math.max(0, commandSuggestionIndex - COMMAND_SUGGESTION_ROWS + 1),
        commandSuggestions.length - COMMAND_SUGGESTION_ROWS,
      );
  const commandPaletteActive = mode.kind === "command" && getCommandPrefix(buffer) !== null;
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

  function refreshEntries(message?: string): void {
    const next = backend.listEntries(workbook.id);
    setEntries(next);
    setPageIndex((current) => clampPageIndex(current, next.length));
    setStatusLines(buildStatusLines(message ?? `Loaded ${next.length} entr${next.length === 1 ? "y" : "ies"}.`));
  }

  function beginAdd(): void {
    setMode({ kind: "add", stage: "vocabulary", vocabulary: "", meanings: [], meaningIndex: 0, metadata: {}, metadataIndex: 0, selectedTagIds: [], posIndex: 0 });
    setBuffer("");
    setStatusLines(buildStatusLines(`Adding a new entry.\nEnter ${workbook.vocabularyLabel}.`));
  }

  function beginPendingCommand(command: ParameterizedCommand): void {
    setMode({ kind: "commandArg", command });
    setBuffer(`/${command} `);
    setStatusLines(buildStatusLines(buildPendingCommandText(command)));
  }

  function beginEdit(entryId: number): void {
    const entry = backend.getEntry(entryId);
    if (!entry) {
      setStatusLines(buildStatusLines(`Entry #${entryId} was not found.`));
      return;
    }

    setMode({
      kind: "edit",
      stage: "vocabulary",
      entryId,
      vocabulary: entry.vocabulary,
      meanings: entry.meanings,
      meaningIndex: 0,
      metadata: entry.attributes,
      metadataIndex: 0,
      selectedTagIds: entry.posTags.map((tag) => tag.id),
      posIndex: 0,
    });
    setBuffer(entry.vocabulary);
    setStatusLines(buildStatusLines(`Editing #${entryId}.\nEdit vocabulary.`));
  }

  function beginDelete(entryId: number): void {
    const entry = backend.getEntry(entryId);
    if (!entry) {
      setStatusLines(buildStatusLines(`Entry #${entryId} was not found.`));
      return;
    }

    setMode({ kind: "delete", entryId, label: buildEntryLabel(entry) });
    setBuffer("");
    setStatusLines(buildStatusLines(`Type yes to delete ${buildEntryLabel(entry)}.`));
  }

  function cancelActiveMode(message = "Cancelled."): void {
    setMode({ kind: "command" });
    setBuffer("");
    setSuggestionIndex(0);
    setStatusLines(buildStatusLines(message));
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
    if (lower === "tag" || lower === "tags") { if (!workbook.posEnabled) setStatusLines(buildStatusLines("Part of speech is disabled. Enable it under /setting.")); else onOpenTags(); return; }
    if (lower === "help") {
      setStatusLines(buildStatusLines(buildHelpText()));
      return;
    }

    if (lower === "view") {
      const entryId = Number(args[0]);
      if (!args[0] || !Number.isInteger(entryId)) { setStatusLines(buildStatusLines("Usage: /view <id>")); return; }
      const entry = backend.getEntry(entryId);
      if (!entry || entry.workbookId !== workbook.id) { setStatusLines(buildStatusLines(`Entry #${entryId} was not found.`)); return; }
      onViewEntry(entry);
      return;
    }

    if (lower === "test") {
      if (args.length > 1 || (args[0] && (!/^\d+$/.test(args[0]) || Number(args[0]) <= 0))) { setStatusLines(buildStatusLines("Usage: /test [count]")); return; }
      onStartPractice(args[0] ? Number(args[0]) : undefined);
      return;
    }

    if (lower === "list") {
      refreshEntries();
      return;
    }

    if (lower === "menu") {
      onBackToMenu();
      return;
    }

    if (lower === "add") {
      beginAdd();
      return;
    }

    if (lower === "edit") {
      const entryId = Number(args[0]);
      if (!args[0] || Number.isNaN(entryId)) {
        beginPendingCommand("edit");
        return;
      }
      beginEdit(entryId);
      return;
    }

    if (lower === "delete") {
      const entryId = Number(args[0]);
      if (!args[0] || Number.isNaN(entryId)) {
        beginPendingCommand("delete");
        return;
      }
      beginDelete(entryId);
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

    if (mode.kind === "add") {
      if (mode.stage === "vocabulary") {
        if (!text) {
          setStatusLines(buildStatusLines("Vocabulary is required."));
          return;
        }

        setMode({
          kind: "add",
          stage: "meaning",
          vocabulary: text,
          meanings: [],
          meaningIndex: 0,
          metadata: {}, metadataIndex: 0, selectedTagIds: [], posIndex: 0,
        });
        setBuffer("");
        setStatusLines(buildStatusLines(`Enter ${workbook.meaningAttributes[0]?.label ?? "Meaning 1"}.`));
        return;
      }

      if (mode.stage === "metadata") {
        const field = activeMetadata[mode.metadataIndex];
        const metadata = { ...mode.metadata, [field.key]: text };
        if (mode.metadataIndex + 1 < activeMetadata.length) { setMode({ ...mode, metadata, metadataIndex: mode.metadataIndex + 1 }); setBuffer(""); setStatusLines(buildStatusLines(`Enter ${activeMetadata[mode.metadataIndex + 1].label}. Optional.`)); return; }
        if (posTags.length > 0) { setMode({ ...mode, metadata, stage: "pos", posIndex: 0 }); setBuffer(""); setStatusLines(buildStatusLines("Select part-of-speech tags with Space, then press Enter.")); return; }
        const entry = backend.addEntry(workbook.id, mode.vocabulary, mode.meanings[0], mode.meanings, metadata, mode.selectedTagIds); setMode({ kind: "command" }); setBuffer(""); refreshEntries(`Added #${entry.id}.`); return;
      }
      if (mode.stage === "pos") {
        const entry = backend.addEntry(workbook.id, mode.vocabulary, mode.meanings[0], mode.meanings, mode.metadata, mode.selectedTagIds); setMode({ kind: "command" }); setBuffer(""); refreshEntries(`Added #${entry.id}.`); return;
      }

      if (mode.meaningIndex === 0 && !text) {
        setStatusLines(buildStatusLines("Meaning is required."));
        return;
      }

      const meanings = [...mode.meanings, text];
      const attributeCount = workbook.meaningAttributes.length;
      if (mode.meaningIndex + 1 < attributeCount) {
        setMode({ ...mode, meanings, meaningIndex: mode.meaningIndex + 1 });
        setBuffer("");
        setStatusLines(buildStatusLines(`Enter ${workbook.meaningAttributes[mode.meaningIndex + 1].label}. Optional.`));
        return;
      }

      if (activeMetadata.length > 0) {
        setMode({ ...mode, meanings, stage: "metadata", metadataIndex: 0 });
        setBuffer(""); setStatusLines(buildStatusLines(`Enter ${activeMetadata[0].label}. Optional.`)); return;
      }
      if (posTags.length > 0) { setMode({ ...mode, meanings, stage: "pos", posIndex: 0 }); setBuffer(""); setStatusLines(buildStatusLines("Select part-of-speech tags with Space, then press Enter.")); return; }
      const entry = backend.addEntry(workbook.id, mode.vocabulary, meanings[0], meanings, mode.metadata, mode.selectedTagIds);
      setMode({ kind: "command" });
      setBuffer("");
      refreshEntries(`Added #${entry.id}.`);
      return;
    }

    if (mode.kind === "edit") {
      if (mode.stage === "vocabulary") {
        if (!text) {
          setStatusLines(buildStatusLines("Vocabulary is required."));
          return;
        }

        setMode({
          kind: "edit",
          stage: "meaning",
          entryId: mode.entryId,
          vocabulary: text,
          meanings: mode.meanings,
          meaningIndex: 0,
          metadata: mode.metadata, metadataIndex: 0, selectedTagIds: mode.selectedTagIds, posIndex: 0,
        });
        setBuffer(mode.meanings[0] ?? "");
        setStatusLines(buildStatusLines(`Update ${workbook.meaningAttributes[0]?.label ?? "Meaning 1"}.`));
        return;
      }

      if (mode.stage === "metadata") {
        const field = activeMetadata[mode.metadataIndex];
        const metadata = { ...mode.metadata, [field.key]: text };
        if (mode.metadataIndex + 1 < activeMetadata.length) { setMode({ ...mode, metadata, metadataIndex: mode.metadataIndex + 1 }); setBuffer(mode.metadata[activeMetadata[mode.metadataIndex + 1].key] ?? ""); setStatusLines(buildStatusLines(`Update ${activeMetadata[mode.metadataIndex + 1].label}. Optional.`)); return; }
        if (posTags.length > 0) { setMode({ ...mode, metadata, stage: "pos", posIndex: 0 }); setBuffer(""); setStatusLines(buildStatusLines("Select part-of-speech tags with Space, then press Enter.")); return; }
        const entry = backend.updateEntry(mode.entryId, mode.vocabulary, mode.meanings[0], mode.meanings, metadata, mode.selectedTagIds); setMode({ kind: "command" }); setBuffer(""); refreshEntries(`Updated #${entry.id}.`); return;
      }
      if (mode.stage === "pos") {
        const entry = backend.updateEntry(mode.entryId, mode.vocabulary, mode.meanings[0], mode.meanings, mode.metadata, mode.selectedTagIds); setMode({ kind: "command" }); setBuffer(""); refreshEntries(`Updated #${entry.id}.`); return;
      }

      if (mode.meaningIndex === 0 && !text) {
        setStatusLines(buildStatusLines("Meaning is required."));
        return;
      }

      const meanings = [...mode.meanings];
      meanings[mode.meaningIndex] = text;
      if (mode.meaningIndex + 1 < workbook.meaningAttributes.length) {
        setMode({ ...mode, meanings, meaningIndex: mode.meaningIndex + 1 });
        setBuffer(mode.meanings[mode.meaningIndex + 1] ?? "");
        setStatusLines(buildStatusLines(`Update ${workbook.meaningAttributes[mode.meaningIndex + 1].label}. Optional.`));
        return;
      }
      if (activeMetadata.length > 0) { setMode({ ...mode, meanings, stage: "metadata", metadataIndex: 0 }); setBuffer(mode.metadata[activeMetadata[0].key] ?? ""); setStatusLines(buildStatusLines(`Update ${activeMetadata[0].label}. Optional.`)); return; }
      if (posTags.length > 0) { setMode({ ...mode, meanings, stage: "pos", posIndex: 0 }); setBuffer(""); setStatusLines(buildStatusLines("Select part-of-speech tags with Space, then press Enter.")); return; }
      const entry = backend.updateEntry(mode.entryId, mode.vocabulary, meanings[0], meanings, mode.metadata, mode.selectedTagIds);
      setMode({ kind: "command" });
      setBuffer("");
      refreshEntries(`Updated #${entry.id}.`);
      return;
    }

    if (mode.kind === "delete") {
      if (text.toLowerCase() !== "yes") {
        cancelActiveMode("Delete cancelled.");
        return;
      }

      backend.deleteEntry(mode.entryId);
      setMode({ kind: "command" });
      setBuffer("");
      refreshEntries(`Deleted ${mode.label}.`);
    }
  }

  useInput((input, key) => {
    if (key.ctrl && input === "c") {
      onQuit();
      return;
    }

    if (key.escape) {
      if (mode.kind !== "command") {
        cancelActiveMode("Cancelled.");
      } else {
        onBackToMenu();
      }
      return;
    }

    if ((mode.kind === "add" || mode.kind === "edit") && mode.stage === "pos" && posTags.length > 0) {
      if (key.upArrow) { setMode({ ...mode, posIndex: Math.max(0, mode.posIndex - 1) }); return; }
      if (key.downArrow) { setMode({ ...mode, posIndex: Math.min(posTags.length - 1, mode.posIndex + 1) }); return; }
      if (input === " ") { const id = posTags[mode.posIndex].id; setMode({ ...mode, selectedTagIds: mode.selectedTagIds.includes(id) ? mode.selectedTagIds.filter((v) => v !== id) : [...mode.selectedTagIds, id] }); return; }
    }

    if (key.upArrow && commandPaletteActive) {
      setSuggestionIndex((current) => (current <= 0 ? commandSuggestions.length - 1 : current - 1));
      return;
    }

    if (key.downArrow && commandPaletteActive) {
      setSuggestionIndex((current) => (current >= commandSuggestions.length - 1 ? 0 : current + 1));
      return;
    }

    if (key.leftArrow && mode.kind === "command" && !commandPaletteActive) {
      setPageIndex((current) => Math.max(0, current - 1));
      return;
    }

    if (key.rightArrow && mode.kind === "command" && !commandPaletteActive) {
      setPageIndex((current) => Math.min(getPageCount(entries.length) - 1, current + 1));
      return;
    }

    if (key.backspace || key.delete) {
      setBuffer((current) => current.slice(0, -1));
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
      if (commandPaletteActive && commandSuggestions.length > 0) {
        submitHighlightedCommand();
      } else if (mode.kind === "command" || mode.kind === "commandArg") {
        submitCommand(current);
      } else {
        submitForm(current);
      }
      return;
    }

    if (!key.ctrl && !key.meta && input) {
      setBuffer((current) => current + input);
    }
  });

  const pageCount = getPageCount(entries.length);
  const safePageIndex = clampPageIndex(pageIndex, entries.length);
  const pageText = `Page ${safePageIndex + 1}/${pageCount}`;
  const tableLines = useMemo(
    () => buildTableLines(entries, safePageIndex, width, PAGE_SIZE, workbook.vocabularyLabel, workbook.meaningAttributes[0]?.label ?? "Meaning 1", workbook.metadataAttributes),
    [entries, safePageIndex, width, workbook.vocabularyLabel, workbook.meaningAttributes, workbook.metadataAttributes],
  );
  const promptLine = `> ${buffer}_`;
  const screenTitle = `${TITLE} — ${workbook.name}`;

  return (
    <Box flexDirection="column">
      <Text color="cyan" bold>
        {centerLine(screenTitle, width)}
      </Text>
      {tableLines.map((line, index) => (
        <Text key={`${index}-${line}`} color={tierColorsEnabled && index >= 3 && index < 3 + entries.slice(safePageIndex * PAGE_SIZE, safePageIndex * PAGE_SIZE + PAGE_SIZE).length ? tierColor(entries[safePageIndex * PAGE_SIZE + index - 3].tier) : undefined}>{line}</Text>
      ))}
      <Text>{padLine("", width)}</Text>
      <Text>{padLine("", width)}</Text>
      <Text color={AUXILIARY_TEXT_COLOR}>{padLine(buildFooterLine(width, pageText, FOOTER_HINT), width)}</Text>
      <Text>{padLine("", width)}</Text>
      <Text>{padLine("", width)}</Text>
      <Text color="cyan">{padLine(promptLine, width)}</Text>
      {(mode.kind === "add" || mode.kind === "edit") && mode.stage === "pos" ? (
        <>
          <Text>{padLine("", width)}</Text>
          {posTags.map((tag, index) => (
            <Text key={tag.id} color={index === mode.posIndex ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>
              {padLine(`${index === mode.posIndex ? ">" : " "} [${mode.selectedTagIds.includes(tag.id) ? "x" : " "}] ${tag.name}`, width)}
            </Text>
          ))}
          <Text color={AUXILIARY_TEXT_COLOR}>{padLine("Space selects | ↑↓ moves | Enter confirms", width)}</Text>
        </>
      ) : null}
      <Text>{padLine("", width)}</Text>
      {commandPaletteActive
        ? suggestionLines.map((line, index) => (
            <Text
              key={`suggestion-${index}-${line.commandText}-${line.hintText}`}
              color={index + commandSuggestionScrollOffset === commandSuggestionIndex ? SELECTED_TEXT_COLOR : undefined}
            >
              <Text color={index + commandSuggestionScrollOffset === commandSuggestionIndex ? SELECTED_TEXT_COLOR : "white"}>{padLine("", 2)}{line.commandText}</Text>
              <Text color={index + commandSuggestionScrollOffset === commandSuggestionIndex ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}> {line.hintText}</Text>
            </Text>
          ))
        : null}
      {commandPaletteActive ? <Text>{padLine("", width)}</Text> : null}
      {!commandPaletteActive
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

  const displayRows = useMemo(() => buildWorkbookMenuLines(workbooks, width, selectedIndex), [workbooks, selectedIndex, width]);

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
      {displayRows.rows.map((row, index) => (
        <Text key={`${index}-${row.line}`} color={row.selected ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>
          {padLine(row.line, width)}
        </Text>
      ))}
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
  const [presetEnabled, setPresetEnabled] = useState(existingWorkbook?.presetEnabled ?? true);
  const [posEnabled, setPosEnabled] = useState(existingWorkbook?.posEnabled ?? true);
  const [tags, setTags] = useState<Array<{ id?: number; name: string; predefined: boolean }>>(() => existingWorkbook ? backend.listStoredPosTags(existingWorkbook.id) : []);
  const [meaningCount, setMeaningCount] = useState(existingWorkbook?.meaningAttributes.length ?? 1);
  const [meanings, setMeanings] = useState<MeaningAttribute[]>(existingWorkbook?.meaningAttributes ?? [{ position: 1, label: "Meaning 1", languageCode: null }]);
  const [meaningIndex, setMeaningIndex] = useState(0);
  const [meaningPalette, setMeaningPalette] = useState<number | null>(null);
  const [attributes, setAttributes] = useState<MetadataAttribute[]>(() => existingWorkbook?.metadataAttributes.filter((item) => item.role === "optional") ?? []);
  const [selected, setSelected] = useState(0);
  const [editAction, setEditAction] = useState<"none" | "add" | "rename">("none");
  const [editBuffer, setEditBuffer] = useState("");
  const [error, setError] = useState("");
  const [confirmBuffer, setConfirmBuffer] = useState("");
  const [pendingConfiguration, setPendingConfiguration] = useState<WorkbookConfigurationInput | null>(null);
  const [destructiveFields, setDestructiveFields] = useState<Array<{ label: string; valueCount: number }>>([]);
  const selectedType = VOCABULARY_TYPES[typeIndex];
  const presetDefinition = selectedType.code ? LANGUAGE_PRESET_DEFINITIONS[selectedType.code] : undefined;
  const presetOptionalAttributes = (presetDefinition?.optionalAttributes ?? []).map((item) => ({ ...item, required: false, visible: false, displayOrder: 0 }));
  const presetKeys = new Set(presetEnabled && selectedType.kind === "preset_language" ? presetOptionalAttributes.map((item) => item.key) : []);
  useEffect(() => { if (!stdout) return; const f = () => setWidth(stdout.columns ?? 80); stdout.on("resize", f); return () => stdout.off("resize", f); }, [stdout]);

  function go(next: CreateStage): void { setHistory((items) => [...items, stage]); setStage(next); setError(""); setSelected(0); setEditAction("none"); setEditBuffer(""); }
  function back(): void { const previous = history.at(-1); if (!previous) return; setHistory((items) => items.slice(0, -1)); setStage(previous); setError(""); setEditAction("none"); }
  function optionalKey(label: string): string {
    const base = label.trim().toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "") || "field";
    const used = new Set(attributes.map((item) => item.key)); let key = base; let suffix = 2;
    while (used.has(key)) key = `${base}_${suffix++}`;
    return key;
  }
  function commitListEdit(): void {
    const label = editBuffer.trim(); if (!label) { setError("A name is required."); return; }
    if (editAction === "add") {
      if (stage === "tags") setTags((items) => items.some((item) => item.name.toLocaleLowerCase() === label.toLocaleLowerCase()) ? items : [...items, { name: label, predefined: false }]);
      else setAttributes((items) => [...items, { key: optionalKey(label), label, languageCode: null, required: false, visible: false, displayOrder: items.length }]);
    } else if (editAction === "rename") {
      if (stage === "tags") setTags((items) => items.map((item, index) => index === selected ? { ...item, name: label } : item));
      else setAttributes((items) => items.map((item, index) => index === selected ? { ...item, label } : item));
    }
    setEditAction("none"); setEditBuffer(""); setError("");
  }
  function configuration(): WorkbookConfigurationInput {
    return { name: name.trim(), vocabularyKind: selectedType.kind, vocabularyLabel: vocabularyLabel.trim(), vocabularyLanguageCode: selectedType.code,
      presetEnabled: selectedType.kind === "preset_language" && presetEnabled, posEnabled, meaningAttributes: meanings.slice(0, meaningCount),
      optionalAttributes: attributes, posTags: tags };
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
        if (!existingWorkbook) { setAttributes([]); setTags([]); setPresetEnabled(selectedType.kind === "preset_language"); }
      }
      if (selectedType.kind === "other_language") setPosEnabled(true);
      if (selectedType.kind === "non_language") setPosEnabled(false);
      go("label"); return;
    }
    if (stage === "label") {
      if (!vocabularyLabel.trim()) { setError("Vocabulary label is required."); return; }
      go(existingWorkbook ? "pos" : selectedType.kind === "preset_language" ? "preset" : "meaning-count"); return;
    }
    if (stage === "preset") {
      if (presetEnabled && selectedType.kind === "preset_language") setAttributes((current) => [...current, ...presetOptionalAttributes.filter((preset) => !current.some((field) => field.key === preset.key)).map((field) => ({ ...field, provenance: "preset" as const }))]);
      else if (!existingWorkbook) setAttributes([]);
      go("pos"); return;
    }
    if (stage === "pos") {
      if (!existingWorkbook && posEnabled && selectedType.kind === "preset_language") setTags((current) => [...current, ...(presetDefinition?.posTags ?? []).filter((name) => !current.some((tag) => tag.name.toLocaleLowerCase() === name.toLocaleLowerCase())).map((tagName) => ({ name: tagName, predefined: true }))]);
      else if (!existingWorkbook && !posEnabled) setTags([]);
      go(posEnabled ? "tags" : "meaning-count"); return;
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
      if (key.backspace || key.delete) setConfirmBuffer((value) => value.slice(0, -1));
      else if (key.return) {
        if (confirmBuffer.trim().toLowerCase() !== "yes") setError("Type yes to confirm removal of populated fields.");
        else if (pendingConfiguration) submit(pendingConfiguration, true);
      } else if (!key.ctrl && !key.meta && input) { setConfirmBuffer((value) => value + input); setError(""); }
      return;
    }
    if (key.leftArrow) { if (editAction !== "none") { setEditAction("none"); setEditBuffer(""); } else if (stage === "meaning" && meaningIndex > 0) { setMeaningIndex((index) => index - 1); setMeaningPalette(null); } else back(); return; }
    if (key.rightArrow) { if (editAction === "none") advance(); return; }
    if (editAction !== "none") {
      if (key.backspace || key.delete) setEditBuffer((value) => value.slice(0, -1));
      else if (key.return) commitListEdit();
      else if (!key.ctrl && !key.meta && input) setEditBuffer((value) => value + input);
      return;
    }
    if (!existingWorkbook && stage === "type" && (key.upArrow || key.downArrow)) { setTypeIndex((value) => key.upArrow ? (value <= 0 ? VOCABULARY_TYPES.length - 1 : value - 1) : (value + 1) % VOCABULARY_TYPES.length); return; }
    if ((stage === "preset" || stage === "pos") && (key.upArrow || key.downArrow || input === " ")) { stage === "preset" ? setPresetEnabled((value) => !value) : setPosEnabled((value) => !value); return; }
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
    if (key.backspace || key.delete) {
      if (stage === "name") setName((value) => value.slice(0, -1));
      else if (stage === "label") setVocabularyLabel((value) => value.slice(0, -1));
      else if (stage === "meaning") setMeanings((items) => items.map((item, index) => index === meaningIndex ? { ...item, label: item.label.slice(0, -1), languageCode: null } : item));
      return;
    }
    if (!key.ctrl && !key.meta && input) {
      if (stage === "name") setName((value) => value + input);
      else if (stage === "label") setVocabularyLabel((value) => value + input);
      else if (stage === "meaning") { setMeaningPalette(null); setMeanings((items) => items.map((item, index) => index === meaningIndex ? { ...item, label: item.label + input, languageCode: null } : item)); }
      setError("");
    }
  });

  const question = ["name"].includes(stage) ? 1 : ["type", "label", "preset", "pos", "tags"].includes(stage) ? 2 : ["meaning-count", "meaning"].includes(stage) ? 3 : stage === "attributes" ? 4 : 4;
  const title = stage === "destructive-confirm" ? "Confirm data removal" : stage === "confirm" ? `Confirm ${existingWorkbook ? "changes" : "creation"}` : `Question ${question}/4`;
  const listItems = stage === "tags" ? tags.map((item) => `${item.name}${item.predefined ? " (preset)" : ""}`) : attributes.map((item) => `${item.label}${presetKeys.has(item.key) ? " (preset)" : ""}`);
  const summary = [`Name: ${name}`, `Vocabulary: ${vocabularyLabel} — ${selectedType.label}`, `Preset attributes: ${presetEnabled && selectedType.kind === "preset_language" ? "enabled" : "disabled"}`, `Part of speech: ${posEnabled ? `enabled (${tags.length} tags)` : "disabled"}`, `Meanings: ${meanings.map((item) => item.label).join(", ")}`, `Optional attributes: ${attributes.map((item) => item.label).join(", ") || "None"}`];
  let description = `Enter a name for the ${existingWorkbook ? "workbook" : "new workbook"}.`;
  let footer = "Enter next | Esc cancel";
  if (stage === "type") { description = existingWorkbook ? "The vocabulary type is fixed after workbook creation." : "Choose a vocabulary type with Up/Down. This selection is required."; footer = existingWorkbook ? "←/→ navigate questions | Esc cancel" : "↑↓ choose type | ←/→ navigate questions | Esc cancel"; }
  if (stage === "label") { description = "Choose the label users will see for vocabulary entries. You can edit the default."; footer = "←/→ to navigate questions | Enter next | Esc cancel"; }
  if (stage === "preset") { description = "Choose whether to add the language preset fields now. After creation they become ordinary attributes managed in Settings."; footer = "↑↓/Space toggle | ←/→ navigate questions | Esc cancel"; }
  if (stage === "pos") { description = "Choose whether this workbook uses Part of Speech tags. You can change this later in Settings."; footer = "↑↓/Space toggle | ←/→ to navigate questions | Esc cancel"; }
  if (stage === "tags") { description = "Review and customize the Part of Speech tags for this workbook."; footer = "↑↓ select | Ctrl+A add | Ctrl+R rename | Del delete | ←/→ navigate"; }
  if (stage === "meaning-count") { description = "Choose how many meaning fields each vocabulary entry will have."; footer = "↑↓ choose number | ←/→ to navigate questions | Esc cancel"; }
  if (stage === "meaning") { description = "Name each meaning field. Use a language preset or type your own label."; footer = "↑↓ choose preset | ←/→ to navigate questions | Esc cancel"; }
  if (stage === "attributes") { description = "Add any other fields you want to store, such as examples or notes. Preset fields are marked."; footer = "↑↓ select | Ctrl+A add | Ctrl+R rename | Del delete | ←/→ navigate"; }
  if (stage === "confirm") { description = `Review the workbook configuration before ${existingWorkbook ? "saving" : "creating"} it.`; footer = `← back | Enter ${existingWorkbook ? "save" : "create"} | Esc cancel`; }
  if (stage === "destructive-confirm") { description = "This change will permanently remove stored field values."; footer = "Type yes and press Enter | Esc cancel"; }
  return <Box flexDirection="column"><Text color="cyan" bold>{centerLine(title, width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine(description, width)}</Text><Text>{padLine("", width)}</Text>
    {stage === "name" ? <Text color="cyan">{padLine(`Workbook name: ${name}_`, width)}</Text> : null}
    {stage === "type" ? existingWorkbook ? <Text color={SELECTED_TEXT_COLOR}>{padLine(selectedType.label, width)}</Text> : VOCABULARY_TYPES.map((item, index) => <Text key={item.label} color={index === typeIndex ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>{padLine(`${index === typeIndex ? ">" : " "} ${item.label}`, width)}</Text>) : null}
    {stage === "label" ? <Text color="cyan">{padLine(`Vocabulary label: ${vocabularyLabel}_`, width)}</Text> : null}
    {stage === "preset" ? <><Text color={presetEnabled ? "green" : AUXILIARY_TEXT_COLOR}>{padLine(`Exclusive attributes: ${presetEnabled ? "Enabled" : "Disabled"}`, width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine(presetOptionalAttributes.length ? `Available: ${presetOptionalAttributes.map((item) => item.label).join(", ")}` : "No exclusive attributes are currently defined for this language.", width)}</Text></> : null}
    {stage === "pos" ? <><Text color={posEnabled ? "green" : AUXILIARY_TEXT_COLOR}>{padLine(`Part of speech: ${posEnabled ? "Enabled" : "Disabled"}`, width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine((presetDefinition?.posTags ?? []).length ? `Preset tags: ${(presetDefinition?.posTags ?? []).join(", ")}` : "No preset tags are currently defined for this language.", width)}</Text></> : null}
    {stage === "meaning-count" ? <Text color="cyan">{padLine(`Meaning attributes: ${meaningCount}`, width)}</Text> : null}
    {stage === "meaning" ? <><Text color="cyan">{padLine(`Meaning ${meaningIndex + 1}/${meaningCount}: ${meanings[meaningIndex]?.label ?? ""}_`, width)}</Text>{meaningPalette !== null ? LANGUAGE_PRESETS.map((item, index) => <Text key={item.code} color={index === meaningPalette ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>{padLine(`${index === meaningPalette ? ">" : " "} ${item.label} (${item.code})`, width)}</Text>) : null}</> : null}
    {(stage === "tags" || stage === "attributes") ? <>{listItems.length ? listItems.map((item, index) => <Text key={`${index}-${typeof item === "string" ? item : ""}`} color={index === selected ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>{padLine(`${index === selected ? ">" : " "} ${item}`, width)}</Text>) : <Text color={AUXILIARY_TEXT_COLOR}>{padLine(stage === "tags" ? "No tags." : "No optional attributes.", width)}</Text>}<Text color="cyan">{padLine(editAction === "none" ? "" : `${editAction === "add" ? "Add" : "Rename"}: ${editBuffer}_`, width)}</Text></> : null}
    {stage === "confirm" ? <>{summary.map((line) => <Text key={line} color={AUXILIARY_TEXT_COLOR}>{padLine(line, width)}</Text>)}<Text>{padLine("", width)}</Text><Text color="green">{padLine(`Press Enter to ${existingWorkbook ? "save changes" : "create the workbook"}.`, width)}</Text></> : null}
    {stage === "destructive-confirm" ? <>{destructiveFields.map((field) => <Text key={field.label} color="red">{padLine(`${field.label}: ${field.valueCount} populated value(s)`, width)}</Text>)}<Text>{padLine("", width)}</Text><Text color="cyan">{padLine(`Confirmation: ${confirmBuffer}_`, width)}</Text></> : null}
    <Text>{padLine("", width)}</Text><Text color="red">{padLine(error, width)}</Text><Text>{padLine("", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{rightLine(footer, width)}</Text></Box>;
}

function WorkbookEditScreen({
  onCreate,
  onCancel,
  onQuit,
  existingWorkbook,
}: {
  onCreate: (name: string, vocabularyLabel: string, vocabularyLanguageCode: string | null, meaningAttributes: MeaningAttribute[]) => void;
  onCancel: () => void;
  onQuit: () => void;
  existingWorkbook?: WorkbookRow;
}): JSX.Element {
  const { stdout } = useStdout();
  const [width, setWidth] = useState(() => stdout?.columns ?? 80);
  const [stage, setStage] = useState<"name" | "vocabulary" | "count" | "meaning">("name");
  const [name, setName] = useState(existingWorkbook?.name ?? "");
  const [vocabularyLabel, setVocabularyLabel] = useState(existingWorkbook?.vocabularyLabel ?? "");
  const [vocabularyLanguageCode, setVocabularyLanguageCode] = useState<string | null>(existingWorkbook?.vocabularyLanguageCode ?? null);
  const [meaningCount, setMeaningCount] = useState(existingWorkbook?.meaningAttributes.length ?? 1);
  const [meaningAttributes, setMeaningAttributes] = useState<MeaningAttribute[]>(existingWorkbook?.meaningAttributes ?? [
    { position: 1, label: "Meaning 1", languageCode: null },
  ]);
  const [meaningIndex, setMeaningIndex] = useState(0);
  const [buffer, setBuffer] = useState(existingWorkbook?.name ?? "");
  const [paletteIndex, setPaletteIndex] = useState(0);
  const [paletteActive, setPaletteActive] = useState(false);
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

    if (stage === "count") {
      if (key.upArrow) {
        setMeaningCount((current) => Math.min(5, current + 1));
        return;
      }
      if (key.downArrow) {
        setMeaningCount((current) => Math.max(1, current - 1));
        return;
      }
      if (key.return) {
        const nextAttributes = Array.from({ length: meaningCount }, (_, index) => meaningAttributes[index] ?? {
          position: index + 1,
          label: `Meaning ${index + 1}`,
          languageCode: null,
        });
        setMeaningAttributes(nextAttributes);
        setMeaningIndex(0);
        setBuffer(nextAttributes[0]?.label ?? "Meaning 1");
        setStage("meaning");
        setError("");
      }
      return;
    }

    if (key.upArrow || key.downArrow) {
      if (stage === "vocabulary" || stage === "meaning") {
        setPaletteActive(true);
        setPaletteIndex((current) => key.upArrow
          ? (current <= 0 ? LANGUAGE_PRESETS.length - 1 : current - 1)
          : (current >= LANGUAGE_PRESETS.length - 1 ? 0 : current + 1));
      }
      return;
    }

    if (key.backspace || key.delete) {
      setBuffer((current) => current.slice(0, -1));
      setPaletteActive(false);
      setError("");
      return;
    }

    if (key.return) {
      if (stage === "name") {
        const trimmed = buffer.trim();
        if (!trimmed) {
          setError("Workbook name is required.");
          return;
        }
        setName(trimmed);
        setBuffer("");
        setStage("vocabulary");
        setError("");
        return;
      }

      if (stage === "vocabulary") {
        if (paletteActive) {
          const preset = LANGUAGE_PRESETS[paletteIndex];
          setVocabularyLabel(preset.label);
          setVocabularyLanguageCode(preset.code);
        } else {
          const label = buffer.trim() || "Vocabulary";
          setVocabularyLabel(label);
          setVocabularyLanguageCode(existingWorkbook && label === existingWorkbook.vocabularyLabel ? existingWorkbook.vocabularyLanguageCode : null);
        }
        setBuffer("");
        setStage("count");
        setPaletteActive(false);
        setError("");
        return;
      }

      if (stage === "meaning") {
        const label = paletteActive ? LANGUAGE_PRESETS[paletteIndex].label : (buffer.trim() || `Meaning ${meaningIndex + 1}`);
        if (!label) {
          setError("Meaning label is required.");
          return;
        }
        const nextAttributes = meaningAttributes.map((attribute, index) => index === meaningIndex
          ? { ...attribute, label, languageCode: paletteActive ? LANGUAGE_PRESETS[paletteIndex].code : (existingWorkbook && label === attribute.label ? attribute.languageCode : null) }
          : attribute);
        if (new Set(nextAttributes.map((attribute) => attribute.label.toLocaleLowerCase())).size !== nextAttributes.length) {
          setError("Meaning attribute labels must be unique.");
          return;
        }
        setMeaningAttributes(nextAttributes);
        if (meaningIndex + 1 < meaningCount) {
          setMeaningIndex((current) => current + 1);
          setBuffer(nextAttributes[meaningIndex + 1]?.label ?? `Meaning ${meaningIndex + 2}`);
          setPaletteActive(false);
          setError("");
          return;
        }
        try {
          onCreate(name, vocabularyLabel || "Vocabulary", vocabularyLanguageCode, nextAttributes);
        } catch (caught) {
          setError(caught instanceof Error ? caught.message : "Could not create workbook.");
        }
        return;
      }
    }

    if (!key.ctrl && !key.meta && input) {
      setBuffer((current) => current + input);
      setPaletteActive(false);
      setError("");
    }
  });

  const currentPrompt = stage === "name"
    ? `Name: > ${buffer}_`
    : stage === "vocabulary"
      ? `Vocabulary: > ${buffer}_`
      : stage === "count"
        ? `Meaning attributes: ${meaningCount}`
        : `${meaningAttributes[meaningIndex]?.label ?? `Meaning ${meaningIndex + 1}`}: > ${buffer}_`;
  const stageHint = stage === "name"
    ? "Enter a workbook name."
    : stage === "vocabulary"
      ? "Type a custom label, or use ↑↓ to choose a language. Blank uses Vocabulary."
      : stage === "count"
        ? "Use ↑↓ to choose 1–5 meaning attributes, then press Enter."
        : `Meaning ${meaningIndex + 1}/${meaningCount}: type a label or use ↑↓ for a language preset.`;

  return (
    <Box flexDirection="column">
      <Text color="cyan" bold>
        {centerLine(existingWorkbook ? "Edit workbook settings" : "Create workbook", width)}
      </Text>
      <Text color={AUXILIARY_TEXT_COLOR}>{padLine(stageHint, width)}</Text>
      <Text>{padLine("", width)}</Text>
      <Text color="cyan">{padLine(currentPrompt, width)}</Text>
      {(stage === "vocabulary" || stage === "meaning") && paletteActive ? (
        <>
          <Text>{padLine("", width)}</Text>
          {LANGUAGE_PRESETS.map((preset, index) => (
            <Text key={preset.code} color={index === paletteIndex ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>
              {padLine(`${index === paletteIndex ? ">" : " "} ${preset.label} (${preset.code})`, width)}
            </Text>
          ))}
        </>
      ) : null}
      <Text>{padLine("", width)}</Text>
      <Text color={AUXILIARY_TEXT_COLOR}>{padLine(error || "Enter advances. Esc cancels. Use /menu to return.", width)}</Text>
    </Box>
  );
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

    if (key.backspace || key.delete) {
      setConfirm((current) => current.slice(0, -1));
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

    if (!key.ctrl && !key.meta && input) {
      setConfirm((current) => current + input);
      setError("");
    }
  });

  const lines = buildWorkbookDeleteLines(workbook, confirm, width);

  return (
    <Box flexDirection="column">
      <Text color="cyan" bold>
        {centerLine("Delete workbook", width)}
      </Text>
      <Text color={AUXILIARY_TEXT_COLOR}>{padLine(`Workbook: ${workbook.name}`, width)}</Text>
      <Text>{padLine("", width)}</Text>
      {lines.map((line, index) => (
        <Text key={`${index}-${line}`} color={AUXILIARY_TEXT_COLOR}>
          {padLine(line, width)}
        </Text>
      ))}
      <Text>{padLine("", width)}</Text>
      <Text color={AUXILIARY_TEXT_COLOR}>{padLine(error || WORKBOOK_DELETE_HINT, width)}</Text>
    </Box>
  );
}

function SettingsHomeScreen({ workbook, onAttributes, onPos, onAppearance, onCancel, onQuit }: { workbook: WorkbookRow; onAttributes: () => void; onPos: () => void; onAppearance: () => void; onCancel: () => void; onQuit: () => void }): JSX.Element {
  const { stdout } = useStdout(); const [width, setWidth] = useState(() => stdout?.columns ?? 80); const [selected, setSelected] = useState(0);
  const sections = ["Attributes", "Part of Speech", "Appearance"];
  useEffect(() => { if (!stdout) return; const f = () => setWidth(stdout.columns ?? 80); stdout.on("resize", f); return () => stdout.off("resize", f); }, [stdout]);
  useInput((input, key) => { if (key.ctrl && input === "c") onQuit(); else if (key.escape) onCancel(); else if (key.upArrow) setSelected((v) => v <= 0 ? sections.length - 1 : v - 1); else if (key.downArrow) setSelected((v) => (v + 1) % sections.length); else if (key.return) [onAttributes, onPos, onAppearance][selected](); });
  return <Box flexDirection="column"><Text color="cyan" bold>{centerLine(`Settings — ${workbook.name}`, width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine("Choose a settings section.", width)}</Text><Text>{padLine("", width)}</Text>{sections.map((section, index) => <Text key={section} color={index === selected ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>{padLine(`${index === selected ? ">" : " "} ${section}`, width)}</Text>)}<Text>{padLine("", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine("Up/Down select | Enter open | Esc back", width)}</Text></Box>;
}

function PosSettingsScreen({ workbook, onToggle, onManage, onCancel, onQuit }: { workbook: WorkbookRow; onToggle: (enabled: boolean) => void; onManage: () => void; onCancel: () => void; onQuit: () => void }): JSX.Element {
  const { stdout } = useStdout(); const [width, setWidth] = useState(() => stdout?.columns ?? 80); const [enabled, setEnabled] = useState(workbook.posEnabled);
  useEffect(() => { if (!stdout) return; const f = () => setWidth(stdout.columns ?? 80); stdout.on("resize", f); return () => stdout.off("resize", f); }, [stdout]);
  useInput((input, key) => { if (key.ctrl && input === "c") onQuit(); else if (key.escape) onCancel(); else if (input === " ") { const next = !enabled; setEnabled(next); onToggle(next); } else if (key.return && enabled) onManage(); });
  return <Box flexDirection="column"><Text color="cyan" bold>{centerLine(`Part of Speech — ${workbook.name}`, width)}</Text><Text>{padLine("", width)}</Text><Text color={enabled ? "green" : AUXILIARY_TEXT_COLOR}>{padLine(`[${enabled ? "x" : " "}] Part of Speech enabled`, width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine(enabled ? "Press Enter to manage tags." : "Stored tags and assignments are preserved while disabled.", width)}</Text><Text>{padLine("", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine("Space toggles | Enter manages tags | Esc back", width)}</Text></Box>;
}

function AppearanceSettingsScreen({ onCancel, onQuit }: { onCancel: () => void; onQuit: () => void }): JSX.Element {
  const { stdout } = useStdout(); const [width, setWidth] = useState(() => stdout?.columns ?? 80); const [enabled, setEnabled] = useState(() => backend.getTierColorsEnabled());
  useEffect(() => { if (!stdout) return; const f = () => setWidth(stdout.columns ?? 80); stdout.on("resize", f); return () => stdout.off("resize", f); }, [stdout]);
  useInput((input, key) => { if (key.ctrl && input === "c") onQuit(); else if (key.escape) onCancel(); else if (input === " ") { const next = !enabled; backend.setTierColorsEnabled(next); setEnabled(next); } });
  return <Box flexDirection="column"><Text color="cyan" bold>{centerLine("Appearance", width)}</Text><Text>{padLine("", width)}</Text><Text color={enabled ? "green" : AUXILIARY_TEXT_COLOR}>{padLine(`[${enabled ? "x" : " "}] Tier colors`, width)}</Text><Text>{padLine("", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine("Space toggles | Esc back", width)}</Text></Box>;
}

type AttributeSection = "vocabulary" | "meaning" | "optional";
type AttributeSelection = { section: AttributeSection; fieldIndex: number | null };

function attributeDraftSignature(vocabularyLabel: string, fields: MetadataAttribute[]): string {
  return JSON.stringify({ vocabularyLabel, fields: fields.map((field) => ({ id: field.id, key: field.key, role: field.role, label: field.label, languageCode: field.languageCode, visible: field.visible })) });
}

function MetadataSettingsScreen({ workbook, onSave, onCancel, onQuit }: { workbook: WorkbookRow; onSave: (draft: WorkbookAttributesDraft, confirmDataLoss: boolean) => void; onCancel: () => void; onQuit: () => void }): JSX.Element {
  const { stdout } = useStdout();
  const [width, setWidth] = useState(() => stdout?.columns ?? 80);
  const initialFields = useMemo(() => workbook.metadataAttributes.filter((field) => field.role !== "vocabulary").map((field) => ({ ...field })), [workbook]);
  const [vocabularyLabel, setVocabularyLabel] = useState(workbook.vocabularyLabel);
  const [fields, setFields] = useState<MetadataAttribute[]>(initialFields);
  const [selectedIndex, setSelectedIndex] = useState(0);
  const [editAction, setEditAction] = useState<"none" | "add" | "rename">("none");
  const [buffer, setBuffer] = useState("");
  const [screenMode, setScreenMode] = useState<"edit" | "exit-confirm" | "destructive-confirm">("edit");
  const [exitChoice, setExitChoice] = useState(0);
  const [confirmBuffer, setConfirmBuffer] = useState("");
  const [destructiveFields, setDestructiveFields] = useState<Array<{ label: string; valueCount: number }>>([]);
  const [message, setMessage] = useState("");
  const meanings = fields.filter((field) => field.role === "meaning");
  const optional = fields.filter((field) => field.role === "optional");
  const selections: AttributeSelection[] = [
    { section: "vocabulary", fieldIndex: null },
    ...fields.map((field, fieldIndex) => ({ section: field.role as "meaning" | "optional", fieldIndex })),
    ...(optional.length === 0 ? [{ section: "optional" as const, fieldIndex: null }] : []),
  ];
  const safeSelectedIndex = Math.min(selectedIndex, Math.max(0, selections.length - 1));
  const selected = selections[safeSelectedIndex];
  const selectedField = selected?.fieldIndex == null ? null : fields[selected.fieldIndex];
  const originalSignature = useMemo(() => attributeDraftSignature(workbook.vocabularyLabel, initialFields), [workbook.vocabularyLabel, initialFields]);
  const dirty = attributeDraftSignature(vocabularyLabel, fields) !== originalSignature;

  useEffect(() => { if (!stdout) return; const f = () => setWidth(stdout.columns ?? 80); stdout.on("resize", f); return () => stdout.off("resize", f); }, [stdout]);
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
  function labelsAreUnique(section: "meaning" | "optional", label: string, ignoredField?: MetadataAttribute): boolean {
    return !fields.some((field) => field !== ignoredField && field.role === section && field.label.trim().toLocaleLowerCase() === label.trim().toLocaleLowerCase());
  }
  function beginAdd(): void {
    if (selected.section === "vocabulary") { setMessage("Vocabulary cannot be added or removed."); return; }
    if (selected.section === "meaning" && meanings.length >= 5) { setMessage("A workbook can have at most five meaning attributes."); return; }
    setEditAction("add"); setBuffer(""); setMessage(`Enter a new ${selected.section === "meaning" ? "meaning" : "optional"} attribute name.`);
  }
  function beginRename(): void {
    setEditAction("rename"); setBuffer(selected.section === "vocabulary" ? vocabularyLabel : selectedField?.label ?? ""); setMessage("Enter a new display name.");
  }
  function commitEdit(): void {
    const label = buffer.trim();
    if (!label) { setMessage("Attribute name is required."); return; }
    if (editAction === "rename") {
      if (selected.section === "vocabulary") setVocabularyLabel(label);
      else if (selectedField) {
        if (!labelsAreUnique(selected.section as "meaning" | "optional", label, selectedField)) { setMessage("Attribute names must be unique within this section."); return; }
        setFields((current) => current.map((field) => field === selectedField ? { ...field, label } : field));
      }
    } else {
      const section = selected.section as "meaning" | "optional";
      if (!labelsAreUnique(section, label)) { setMessage("Attribute names must be unique within this section."); return; }
      const newField: MetadataAttribute = { key: nextKey(section, label), role: section, label, languageCode: null, required: false, visible: false, displayOrder: section === "meaning" ? meanings.length + 1 : optional.length + 1 };
      setFields((current) => section === "meaning" ? [...current.filter((field) => field.role === "meaning"), newField, ...current.filter((field) => field.role === "optional")] : [...current, newField]);
      setSelectedIndex(section === "meaning" ? 1 + meanings.length : 1 + meanings.length + optional.length);
    }
    setEditAction("none"); setBuffer(""); setMessage("");
  }
  function currentDraft(): WorkbookAttributesDraft { return { vocabularyLabel, fields }; }
  function save(confirmDataLoss = false): void {
    try { onSave(currentDraft(), confirmDataLoss); }
    catch (caught) {
      if (caught instanceof WorkbookDataLossError) {
        setDestructiveFields(caught.impact.populatedFields); setConfirmBuffer(""); setScreenMode("destructive-confirm"); setMessage("");
      } else { setScreenMode("edit"); setMessage(caught instanceof Error ? caught.message : "Could not save attributes."); }
    }
  }

  useInput((input, key) => {
    if (key.ctrl && input === "c") return onQuit();
    if (screenMode === "destructive-confirm") {
      if (key.escape) { setScreenMode("edit"); setConfirmBuffer(""); return; }
      if (key.backspace || key.delete) { setConfirmBuffer((value) => value.slice(0, -1)); return; }
      if (key.return) { if (confirmBuffer.trim().toLowerCase() === "yes") save(true); else setMessage("Type yes to confirm removal of populated fields."); return; }
      if (!key.ctrl && !key.meta && input) { setConfirmBuffer((value) => value + input); setMessage(""); }
      return;
    }
    if (screenMode === "exit-confirm") {
      if (key.escape) { setScreenMode("edit"); return; }
      if (key.upArrow) { setExitChoice((value) => value <= 0 ? 2 : value - 1); return; }
      if (key.downArrow) { setExitChoice((value) => (value + 1) % 3); return; }
      if (key.return) { if (exitChoice === 0) save(); else if (exitChoice === 1) onCancel(); else setScreenMode("edit"); }
      return;
    }
    if (editAction !== "none") {
      if (key.escape) { setEditAction("none"); setBuffer(""); setMessage(""); return; }
      if (key.backspace || key.delete) { setBuffer((value) => value.slice(0, -1)); return; }
      if (key.return) { commitEdit(); return; }
      if (!key.ctrl && !key.meta && input) { setBuffer((value) => value + input); setMessage(""); }
      return;
    }
    if (key.escape) { if (dirty) { setExitChoice(0); setScreenMode("exit-confirm"); } else onCancel(); return; }
    if (key.upArrow) { setSelectedIndex((value) => value <= 0 ? selections.length - 1 : value - 1); setMessage(""); return; }
    if (key.downArrow) { setSelectedIndex((value) => (value + 1) % selections.length); setMessage(""); return; }
    if (key.ctrl && input.toLowerCase() === "a") { beginAdd(); return; }
    if (key.ctrl && input.toLowerCase() === "r") { beginRename(); return; }
    if (input === " ") {
      if (selected.section === "vocabulary" || (selected.section === "meaning" && meanings[0] === selectedField)) { setMessage(selected.section === "vocabulary" ? "Vocabulary is always shown." : "Meaning 1 is required and always shown."); return; }
      if (selectedField) setFields((current) => current.map((field) => field === selectedField ? { ...field, visible: !field.visible } : field));
      return;
    }
    if (key.delete) {
      if (selected.section === "vocabulary") { setMessage("Vocabulary cannot be deleted."); return; }
      if (!selectedField) return;
      if (selected.section === "meaning" && meanings[0] === selectedField) { setMessage("Meaning 1 cannot be deleted."); return; }
      setFields((current) => current.filter((field) => field !== selectedField)); setSelectedIndex((value) => Math.max(0, value - 1)); setMessage(`${selectedField.label} will be removed if changes are saved.`);
    }
  });

  if (screenMode === "exit-confirm") {
    const choices = ["Save changes", "Discard changes", "Continue editing"];
    return <Box flexDirection="column"><Text color="cyan" bold>{centerLine("Unsaved attribute changes", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine("Choose what to do with your changes.", width)}</Text><Text>{padLine("", width)}</Text>{choices.map((choice, index) => <Text key={choice} color={index === exitChoice ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>{padLine(`${index === exitChoice ? ">" : " "} ${choice}`, width)}</Text>)}<Text>{padLine("", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{rightLine("↑↓ select | Enter confirm | Esc continue editing", width)}</Text></Box>;
  }
  if (screenMode === "destructive-confirm") {
    return <Box flexDirection="column"><Text color="cyan" bold>{centerLine("Confirm data removal", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine("Saving will permanently remove stored values from these fields:", width)}</Text><Text>{padLine("", width)}</Text>{destructiveFields.map((field) => <Text key={field.label} color="red">{padLine(`${field.label}: ${field.valueCount} populated value(s)`, width)}</Text>)}<Text>{padLine("", width)}</Text><Text color="cyan">{padLine(`Type yes: ${confirmBuffer}_`, width)}</Text><Text color="red">{padLine(message, width)}</Text><Text>{padLine("", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{rightLine("Enter confirm | Esc continue editing", width)}</Text></Box>;
  }

  const meaningFooter = meanings[0] === selectedField ? "↑↓ navigate | Ctrl+A add | Ctrl+R rename | Esc leave" : "↑↓ navigate | Ctrl+A add | Ctrl+R rename | Space show/hide | Del remove | Esc leave";
  const footer = editAction !== "none" ? "Enter confirm | Esc cancel edit" : selected.section === "vocabulary" ? "↑↓ navigate | Ctrl+R rename | Esc leave" : selected.section === "meaning" ? meaningFooter : "↑↓ navigate | Ctrl+A add | Ctrl+R rename | Space show/hide | Del remove | Esc leave";
  const row = (field: MetadataAttribute, fieldIndex: number) => <Text key={field.id ?? field.key} color={selected.fieldIndex === fieldIndex ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>{`${selected.fieldIndex === fieldIndex ? ">" : " "} ${field.label} [${field.visible ? "shown" : "hidden"}]${field.required ? " (required)" : ""}`}</Text>;
  return <Box flexDirection="column"><Text color="cyan" bold>{centerLine(`Attributes — ${workbook.name}`, width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine("Configure field names and visibility. Changes are staged until you leave this page.", width)}</Text><Text color={message ? "red" : AUXILIARY_TEXT_COLOR}>{padLine(message, width)}</Text>
    <Box flexDirection="column" borderStyle="single" borderColor={selected.section === "vocabulary" ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR} paddingX={1}><Text bold color={selected.section === "vocabulary" ? SELECTED_TEXT_COLOR : undefined}>Vocabulary</Text><Text color={AUXILIARY_TEXT_COLOR}>The workbook type is fixed. Only this display name can be changed.</Text><Text color={selected.section === "vocabulary" ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>{`${selected.section === "vocabulary" ? ">" : " "} ${vocabularyLabel} [shown] (required)`}</Text></Box>
    <Box flexDirection="column" borderStyle="single" borderColor={selected.section === "meaning" ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR} paddingX={1}><Text bold color={selected.section === "meaning" ? SELECTED_TEXT_COLOR : undefined}>Meanings</Text><Text color={AUXILIARY_TEXT_COLOR}>Meaning 1 is required and always shown. Up to five meanings are supported.</Text>{fields.map((field, index) => field.role === "meaning" ? row(field, index) : null)}</Box>
    <Box flexDirection="column" borderStyle="single" borderColor={selected.section === "optional" ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR} paddingX={1}><Text bold color={selected.section === "optional" ? SELECTED_TEXT_COLOR : undefined}>Optional Attributes</Text><Text color={AUXILIARY_TEXT_COLOR}>Supplemental fields can be freely added, renamed, hidden, shown, or removed.</Text>{optional.length ? fields.map((field, index) => field.role === "optional" ? row(field, index) : null) : <Text color={selected.section === "optional" ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>{`${selected.section === "optional" ? ">" : " "} No optional attributes`}</Text>}</Box>
    {editAction !== "none" ? <Text color="cyan">{padLine(`${editAction === "add" ? "New attribute" : "Display name"}: ${buffer}_`, width)}</Text> : null}<Text>{padLine("", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{rightLine(footer, width)}</Text></Box>;
}

function PosTagScreen({ workbook, onCancel, onQuit }: { workbook: WorkbookRow; onCancel: () => void; onQuit: () => void }): JSX.Element {
  const { stdout } = useStdout(); const [width, setWidth] = useState(() => stdout?.columns ?? 80); const [tags, setTags] = useState<PosTag[]>(() => backend.listPosTags(workbook.id)); const [selected, setSelected] = useState(0); const [buffer, setBuffer] = useState(""); const [action, setAction] = useState<"none" | "add" | "rename">("none"); const [message, setMessage] = useState(tags.length ? "↑↓ select | Ctrl+A add | Ctrl+R rename | Del delete | Esc back" : "No POS tags. Press Ctrl+A to add one, or Esc to return.");
  useEffect(() => { if (!stdout) return; const f = () => setWidth(stdout.columns ?? 80); stdout.on("resize", f); return () => stdout.off("resize", f); }, [stdout]);
  useInput((input, key) => {
    if (key.ctrl && input === "c") return onQuit();
    if (key.escape) { if (action !== "none") { setAction("none"); setBuffer(""); setMessage("↑↓ select | Ctrl+A add | Ctrl+R rename | Del delete | Esc back"); } else onCancel(); return; }
    if (action !== "none") {
      if (key.backspace || key.delete) setBuffer((value) => value.slice(0, -1));
      else if (key.return && buffer.trim()) {
        try {
          if (action === "add") backend.addPosTag(workbook.id, buffer); else if (tags[selected]) backend.renamePosTag(tags[selected].id, buffer);
          setTags(backend.listPosTags(workbook.id)); setBuffer(""); setAction("none"); setMessage("↑↓ select | Ctrl+A add | Ctrl+R rename | Del delete | Esc back");
        } catch (caught) { setMessage(caught instanceof Error ? caught.message : "Could not update tag."); }
      } else if (!key.ctrl && !key.meta && input) setBuffer((value) => value + input);
      return;
    }
    if (key.upArrow) return setSelected((value) => Math.max(0, value - 1));
    if (key.downArrow) return setSelected((value) => Math.min(Math.max(0, tags.length - 1), value + 1));
    if (key.ctrl && input.toLowerCase() === "a") { setAction("add"); setBuffer(""); setMessage("Type tag name and press Enter."); return; }
    if (key.ctrl && input.toLowerCase() === "r" && tags[selected]) { setAction("rename"); setBuffer(tags[selected].name); setMessage("Type replacement name and press Enter."); return; }
    if (key.delete && tags[selected]) { backend.deletePosTag(tags[selected].id); setTags(backend.listPosTags(workbook.id)); setSelected(0); }
  });
  return <Box flexDirection="column"><Text color="cyan" bold>{centerLine(`Part of speech — ${workbook.name}`, width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine(message, width)}</Text><Text>{padLine("", width)}</Text>{tags.map((tag, i) => <Text key={tag.id} color={i === selected ? SELECTED_TEXT_COLOR : AUXILIARY_TEXT_COLOR}>{padLine(`${i === selected ? ">" : " "} ${tag.name}${tag.predefined ? " (preset)" : ""}`, width)}</Text>)}<Text>{padLine("", width)}</Text><Text color="cyan">{padLine(buffer ? `> ${buffer}_` : "", width)}</Text></Box>;
}

function EntryViewScreen({ workbook, entry, onCancel, onQuit }: { workbook: WorkbookRow; entry: EntryRow; onCancel: () => void; onQuit: () => void }): JSX.Element {
  const { stdout } = useStdout();
  const [width, setWidth] = useState(() => stdout?.columns ?? 80);
  useEffect(() => { if (!stdout) return; const f = () => setWidth(stdout.columns ?? 80); stdout.on("resize", f); return () => stdout.off("resize", f); }, [stdout]);
  useInput((_input, key) => { if (key.ctrl && _input === "c") onQuit(); else if (key.escape) onCancel(); });
  const lines = [
    ...buildExplicitEntryLines(workbook, entry),
    "",
    `Tests: ${entry.testCount}`,
    `Errors: ${entry.errorCount}`,
    `Tier: ${entry.tier[0].toUpperCase()}${entry.tier.slice(1)}`,
    `Last tested: ${formatLastTested(entry.lastTested)}`,
  ];
  return <Box flexDirection="column"><Text color="cyan" bold>{centerLine(`Entry #${entry.id}`, width)}</Text><Text>{padLine("", width)}</Text>{lines.map((line, index) => <Text key={`${index}-${line}`} color={index >= lines.length - 4 ? tierColor(entry.tier) : AUXILIARY_TEXT_COLOR}>{padLine(line, width)}</Text>)}<Text>{padLine("", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine("Read-only view. Esc returns to the vocabulary list.", width)}</Text></Box>;
}

function buildExplicitEntryLines(workbook: WorkbookRow, entry: EntryRow): string[] {
  return [
    `ID: #${entry.id}`,
    `${workbook.vocabularyLabel}: ${entry.vocabulary}`,
    ...entry.meanings.map((meaning, index) => `${workbook.meaningAttributes[index]?.label ?? `Meaning ${index + 1}`}: ${meaning}`),
    ...workbook.metadataAttributes.filter((attribute) => attribute.role === "optional").map((attribute) => `${attribute.label}: ${entry.attributes[attribute.key] ?? ""}`),
    `Part of speech: ${entry.posTags.map((tag) => tag.name).join(", ") || "None"}`,
  ];
}

function PracticeScreen({ workbook, count, onCancel, onQuit, onDone }: { workbook: WorkbookRow; count: number; onCancel: () => void; onQuit: () => void; onDone: (score: number, total: number) => void }): JSX.Element {
  const { stdout } = useStdout();
  const [width, setWidth] = useState(() => stdout?.columns ?? 80);
  const [questions] = useState(() => backend.selectPracticeCandidates(workbook.id, count));
  const [phase, setPhase] = useState<"initial" | "retry" | "detail" | "done">("initial");
  const [index, setIndex] = useState(0);
  const [retryRound, setRetryRound] = useState<EntryRow[]>([]);
  const [nextRetryRound, setNextRetryRound] = useState<EntryRow[]>([]);
  const [retryNumber, setRetryNumber] = useState(1);
  const [detailEntry, setDetailEntry] = useState<EntryRow | null>(null);
  const [detailSourcePhase, setDetailSourcePhase] = useState<"initial" | "retry">("initial");
  const [answer, setAnswer] = useState("");
  const [feedback, setFeedback] = useState<string | null>(null);
  const [score, setScore] = useState(0);
  const current = phase === "initial" ? questions[index] : phase === "retry" ? retryRound[index] : null;
  useEffect(() => { if (!stdout) return; const f = () => setWidth(stdout.columns ?? 80); stdout.on("resize", f); return () => stdout.off("resize", f); }, [stdout]);

  function advanceAfterAnswer(sourcePhase: "initial" | "retry" = phase as "initial" | "retry", queuedRetry = nextRetryRound): void {
    if (sourcePhase === "initial") {
      const nextIndex = index + 1;
      if (nextIndex < questions.length) { setIndex(nextIndex); return; }
      if (queuedRetry.length > 0) { setRetryRound(queuedRetry); setNextRetryRound([]); setIndex(0); setRetryNumber(1); setPhase("retry"); return; }
      setPhase("done");
      return;
    }
    if (sourcePhase === "retry") {
      const nextIndex = index + 1;
      if (nextIndex < retryRound.length) { setIndex(nextIndex); return; }
      if (queuedRetry.length > 0) { setRetryRound(queuedRetry); setNextRetryRound([]); setIndex(0); setRetryNumber((n) => n + 1); return; }
      setPhase("done");
    }
  }

  useInput((input, key) => {
    if (key.ctrl && input === "c") return onQuit();
    if (key.escape) return onCancel();
    if (phase === "done") { if (key.return) onDone(score, questions.length); return; }
    if (phase === "detail") {
      if (key.return) {
        const source = detailSourcePhase;
        const queuedRetry = detailEntry ? [...nextRetryRound, detailEntry] : nextRetryRound;
        setNextRetryRound(queuedRetry);
        setDetailEntry(null);
        setFeedback(null);
        setAnswer("");
        setPhase(source);
        advanceAfterAnswer(source, queuedRetry);
      }
      return;
    }
    if (feedback !== null) {
      if (key.return) { setFeedback(null); setAnswer(""); advanceAfterAnswer(); }
      return;
    }
    if (key.backspace || key.delete) { setAnswer((v) => v.slice(0, -1)); return; }
    if (key.return) {
      if (!current) { setPhase("done"); return; }
      const given = answer.trim(); const correct = given === current.vocabulary;
      const updated = backend.recordTestResult(current.id, correct, phase === "initial");
      if (phase === "initial" && correct) setScore((v) => v + 1);
      if (!correct) {
        setDetailEntry(updated);
        setDetailSourcePhase(phase);
        setPhase("detail");
      } else {
        setFeedback("Correct!");
      }
      return;
    }
    if (!key.ctrl && !key.meta && input) setAnswer((v) => v + input);
  });
  if (questions.length === 0) return <PracticeEmptyScreen workbook={workbook} onCancel={onCancel} onQuit={onQuit} />;
  if (phase === "done") return <Box flexDirection="column"><Text color="cyan" bold>{centerLine(`Practice — ${workbook.name}`, width)}</Text><Text>{padLine("", width)}</Text><Text color="green">{padLine(`Final initial-round score: ${score}/${questions.length}`, width)}</Text><Text>{padLine("Press Enter to return.", width)}</Text></Box>;
  if (phase === "detail" && detailEntry) return <Box flexDirection="column"><Text color="cyan" bold>{centerLine(`Entry #${detailEntry.id}`, width)}</Text><Text color="red">{padLine(`Incorrect — expected: ${detailEntry.vocabulary}`, width)}</Text><Text>{padLine("", width)}</Text>{buildExplicitEntryLines(workbook, detailEntry).map((line, i) => <Text key={`${i}-${line}`} color={AUXILIARY_TEXT_COLOR}>{padLine(line, width)}</Text>)}<Text>{padLine("", width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine("Enter advances. Esc cancels.", width)}</Text></Box>;
  const roundLabel = phase === "retry" ? `Retry round ${retryNumber} — Question ${index + 1}/${retryRound.length}` : `Question ${index + 1}/${questions.length}`;
  return <Box flexDirection="column"><Text color="cyan" bold>{centerLine(`Practice — ${workbook.name}`, width)}</Text><Text color={AUXILIARY_TEXT_COLOR}>{padLine(roundLabel, width)}</Text><Text>{padLine("", width)}</Text><Text color="yellow">{padLine(`${workbook.meaningAttributes[0]?.label ?? "Meaning 1"}: ${current?.meaning ?? ""}`, width)}</Text><Text>{padLine("", width)}</Text><Text color="cyan">{padLine(`Answer: ${answer}_`, width)}</Text><Text>{padLine("", width)}</Text><Text color="green">{padLine(feedback ?? "Enter submits. Esc cancels.", width)}</Text></Box>;
}

function PracticeEmptyScreen({ workbook, onCancel, onQuit }: { workbook: WorkbookRow; onCancel: () => void; onQuit: () => void }): JSX.Element {
  useInput((input, key) => { if (key.ctrl && input === "c") onQuit(); else if (key.escape || key.return) onCancel(); });
  return <Box flexDirection="column"><Text color="cyan">{`No entries available in ${workbook.name}.`}</Text><Text color={AUXILIARY_TEXT_COLOR}>Press Enter or Esc to return.</Text></Box>;
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
    return <VocabularyScreen workbook={screen.workbook} onBackToMenu={backToMenu} onQuit={quit} onOpenSettings={() => setScreen({ kind: "settings", workbook: refreshWorkbook(screen.workbook.id) })} onOpenTags={() => setScreen({ kind: "tags", workbook: refreshWorkbook(screen.workbook.id) })} onViewEntry={(entry) => setScreen({ kind: "view", workbook: screen.workbook, entry })} onStartPractice={(count) => { const n = count ?? 15; setScreen({ kind: "practice", workbook: refreshWorkbook(screen.workbook.id), count: Math.min(n, backend.countEntries(screen.workbook.id)) }); }} />;
  }

  if (screen.kind === "practice") return <PracticeScreen workbook={screen.workbook} count={screen.count} onCancel={() => setScreen({ kind: "vocab", workbook: refreshWorkbook(screen.workbook.id) })} onQuit={quit} onDone={() => setScreen({ kind: "vocab", workbook: refreshWorkbook(screen.workbook.id) })} />;

  if (screen.kind === "settings") return <SettingsHomeScreen workbook={screen.workbook} onAttributes={() => setScreen({ kind: "settings-attributes", workbook: refreshWorkbook(screen.workbook.id) })} onPos={() => setScreen({ kind: "settings-pos", workbook: refreshWorkbook(screen.workbook.id) })} onAppearance={() => setScreen({ kind: "settings-appearance", workbook: refreshWorkbook(screen.workbook.id) })} onCancel={() => setScreen({ kind: "vocab", workbook: refreshWorkbook(screen.workbook.id) })} onQuit={quit} />;
  if (screen.kind === "settings-attributes") return <MetadataSettingsScreen workbook={screen.workbook} onSave={(draft, confirmDataLoss) => { backend.updateWorkbookAttributes(screen.workbook.id, draft, confirmDataLoss); setScreen({ kind: "settings", workbook: refreshWorkbook(screen.workbook.id) }); }} onCancel={() => setScreen({ kind: "settings", workbook: refreshWorkbook(screen.workbook.id) })} onQuit={quit} />;
  if (screen.kind === "settings-pos") return <PosSettingsScreen workbook={screen.workbook} onToggle={(enabled) => { backend.setPosEnabled(screen.workbook.id, enabled); }} onManage={() => setScreen({ kind: "tags", workbook: refreshWorkbook(screen.workbook.id), returnTo: "settings-pos" })} onCancel={() => setScreen({ kind: "settings", workbook: refreshWorkbook(screen.workbook.id) })} onQuit={quit} />;
  if (screen.kind === "settings-appearance") return <AppearanceSettingsScreen onCancel={() => setScreen({ kind: "settings", workbook: refreshWorkbook(screen.workbook.id) })} onQuit={quit} />;
  if (screen.kind === "tags") return <PosTagScreen workbook={screen.workbook} onCancel={() => setScreen(screen.returnTo === "settings-pos" ? { kind: "settings-pos", workbook: refreshWorkbook(screen.workbook.id) } : { kind: "vocab", workbook: refreshWorkbook(screen.workbook.id) })} onQuit={quit} />;
  if (screen.kind === "view") return <EntryViewScreen workbook={screen.workbook} entry={backend.getEntry(screen.entry.id) ?? screen.entry} onCancel={() => setScreen({ kind: "vocab", workbook: refreshWorkbook(screen.workbook.id) })} onQuit={quit} />;

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
