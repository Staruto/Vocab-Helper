import React, { useEffect, useMemo, useState } from "react";
import { Box, Text, render, useApp, useInput, useStdout } from "ink";
import { EntryRow, MeaningAttribute, WorkbookRow } from "./db.js";
import { VocabularyBackend } from "./backend.js";

type UiMode =
  | { kind: "command" }
  | { kind: "commandArg"; command: ParameterizedCommand }
  | { kind: "add"; stage: "vocabulary" | "meaning"; vocabulary: string; meanings: string[]; meaningIndex: number }
  | { kind: "edit"; stage: "vocabulary" | "meaning"; entryId: number; vocabulary: string; meanings: string[]; meaningIndex: number }
  | { kind: "delete"; entryId: number; label: string };

type AppScreen =
  | { kind: "menu" }
  | { kind: "create-workbook" }
  | { kind: "edit-workbook"; workbook: WorkbookRow }
  | { kind: "delete-workbook"; workbook: WorkbookRow; confirm: string }
  | { kind: "vocab"; workbook: WorkbookRow };

type VocabularyScreenProps = {
  workbook: WorkbookRow;
  onBackToMenu: () => void;
  onQuit: () => void;
};

type CommandSpec = {
  name: string;
  hint: string;
};

type ParameterizedCommand = "edit" | "delete";

type LanguagePreset = { code: string; label: string };

const PAGE_SIZE = 20;
const TITLE = "VocabHelper MVP";
const FOOTER_HINT = "Navigate pages with <- -> | Esc returns to menu";
const AUXILIARY_TEXT_COLOR = "gray";
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
const WORKBOOK_MENU_HINT = "↑↓ select | Enter open | E edit | Del delete | + create | Esc quit";
const WORKBOOK_CREATE_HINT = "Type a name and press Enter. Esc returns to the menu.";
const WORKBOOK_DELETE_HINT = "Type yes to confirm. Enter deletes. Esc cancels.";
const COMMANDS: CommandSpec[] = [
  { name: "list", hint: "Refresh and show entries" },
  { name: "add", hint: "Add a new entry" },
  { name: "edit", hint: "Edit an entry by id" },
  { name: "delete", hint: "Delete an entry by id" },
  { name: "help", hint: "Show command help" },
  { name: "quit", hint: "Exit the app" },
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

function buildTableLines(entries: EntryRow[], pageIndex: number, width: number, availableRows: number, vocabularyLabel: string, meaningLabel: string): string[] {
  const totalWidth = Math.max(width, 40);
  const innerWidth = Math.max(20, totalWidth - 2);
  const gap = 1;
  const minimumLeft = Math.max(14, Math.floor(innerWidth * 0.4));
  const minimumRight = Math.max(12, Math.floor(innerWidth * 0.25));
  const desiredLeft = Math.floor(innerWidth * 0.58);
  const maxLeft = Math.max(minimumLeft, innerWidth - gap - minimumRight);
  const leftWidth = Math.max(minimumLeft, Math.min(desiredLeft, maxLeft));
  const rightWidth = Math.max(minimumRight, innerWidth - gap - leftWidth);
  const visibleRows = Math.max(1, Math.min(PAGE_SIZE, availableRows));
  const pageEntries = entries.slice(pageIndex * PAGE_SIZE, pageIndex * PAGE_SIZE + visibleRows);
  const border = `+${"-".repeat(leftWidth)}+${"-".repeat(rightWidth)}+`;
  const header = `|${padLine(vocabularyLabel, leftWidth)}|${padLine(meaningLabel, rightWidth)}|`;

  const rows = pageEntries.map((entry) => {
    const left = padLine(buildEntryLabel(entry), leftWidth);
    const right = padLine(entry.meaning, rightWidth);
    return `|${left}|${right}|`;
  });

  return [border, header, border, ...rows, border];
}

function buildHelpText(): string {
  return ["Commands:", ...COMMANDS.map((command) => `/${command.name}  ${command.hint}`), "Esc cancels forms.", "Use <- -> to change pages."].join("\n");
}

function buildPendingCommandText(command: ParameterizedCommand): string {
  return `Enter id for /${command}.`;
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

function buildSuggestionLine(command: CommandSpec, selected: boolean, width: number): { commandText: string; hintText: string } {
  const totalWidth = Math.max(width, 40);
  const leftWidth = Math.max(12, Math.min(16, Math.floor(totalWidth * 0.22)));
  const rightWidth = Math.max(1, totalWidth - 3 - leftWidth);
  return {
    commandText: padLine(`/${command.name}`, leftWidth),
    hintText: padLine(command.hint, rightWidth),
  };
}

function buildSuggestionLines(suggestions: CommandSpec[], selectedIndex: number, width: number): Array<{ commandText: string; hintText: string }> {
  if (suggestions.length === 0) {
    return [
      { commandText: padLine("No matching commands.", width), hintText: "" },
      ...Array.from({ length: COMMAND_SUGGESTION_ROWS - 1 }, () => ({ commandText: "", hintText: "" })),
    ];
  }

  const rows = suggestions.slice(0, COMMAND_SUGGESTION_ROWS).map((command, index) =>
    buildSuggestionLine(command, index === selectedIndex, width),
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

function VocabularyScreen({ workbook, onBackToMenu, onQuit }: VocabularyScreenProps): JSX.Element {
  const { stdout } = useStdout();
  const [width, setWidth] = useState(() => stdout?.columns ?? 80);
  const [rows, setRows] = useState(() => stdout?.rows ?? 24);
  const [entries, setEntries] = useState<EntryRow[]>(() => backend.listEntries(workbook.id));
  const [pageIndex, setPageIndex] = useState(0);
  const [mode, setMode] = useState<UiMode>({ kind: "command" });
  const [buffer, setBuffer] = useState("");
  const [suggestionIndex, setSuggestionIndex] = useState(0);
  const [statusLines, setStatusLines] = useState<string[]>(() => buildStatusLines("Ready."));

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
  const commandPaletteActive = mode.kind === "command" && getCommandPrefix(buffer) !== null;
  const suggestionLines = useMemo(
    () => buildSuggestionLines(commandSuggestions, commandSuggestionIndex, width),
    [commandSuggestions, commandSuggestionIndex, width],
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
    setMode({ kind: "add", stage: "vocabulary", vocabulary: "", meanings: [], meaningIndex: 0 });
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

    if (lower === "help") {
      setStatusLines(buildStatusLines(buildHelpText()));
      return;
    }

    if (lower === "list") {
      refreshEntries();
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

    if (lower === "quit" || lower === "exit") {
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
        });
        setBuffer("");
        setStatusLines(buildStatusLines(`Enter ${workbook.meaningAttributes[0]?.label ?? "Meaning 1"}.`));
        return;
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

      const entry = backend.addEntry(workbook.id, mode.vocabulary, meanings[0], meanings);
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
        });
        setBuffer(mode.meanings[0] ?? "");
        setStatusLines(buildStatusLines(`Update ${workbook.meaningAttributes[0]?.label ?? "Meaning 1"}.`));
        return;
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

      const entry = backend.updateEntry(mode.entryId, mode.vocabulary, meanings[0], meanings);
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
      onBackToMenu();
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
    () => buildTableLines(entries, safePageIndex, width, PAGE_SIZE, workbook.vocabularyLabel, workbook.meaningAttributes[0]?.label ?? "Meaning 1"),
    [entries, safePageIndex, width, workbook.vocabularyLabel, workbook.meaningAttributes],
  );
  const promptLine = `> ${buffer}_`;
  const screenTitle = `${TITLE} — ${workbook.name}`;

  return (
    <Box flexDirection="column">
      <Text color="cyan" bold>
        {centerLine(screenTitle, width)}
      </Text>
      {tableLines.map((line, index) => (
        <Text key={`${index}-${line}`}>{line}</Text>
      ))}
      <Text>{padLine("", width)}</Text>
      <Text>{padLine("", width)}</Text>
      <Text color={AUXILIARY_TEXT_COLOR}>{padLine(buildFooterLine(width, pageText, FOOTER_HINT), width)}</Text>
      <Text>{padLine("", width)}</Text>
      <Text>{padLine("", width)}</Text>
      <Text color="cyan">{padLine(promptLine, width)}</Text>
      <Text>{padLine("", width)}</Text>
      {commandPaletteActive
        ? suggestionLines.map((line, index) => (
            <Text key={`suggestion-${index}-${line.commandText}-${line.hintText}`}>
              {index === commandSuggestionIndex ? (
                <>
                  <Text color="yellow">{padLine(">", 2)}</Text>
                  <Text color="yellow">{line.commandText}</Text>
                </>
              ) : (
                <>
                  <Text color={AUXILIARY_TEXT_COLOR}>{padLine(" ", 2)}</Text>
                  <Text color="white">{line.commandText}</Text>
                </>
              )}
              <Text color={AUXILIARY_TEXT_COLOR}> {line.hintText}</Text>
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

    if (input.toLowerCase() === "e") {
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
        <Text key={`${index}-${row.line}`} color={row.selected ? "yellow" : AUXILIARY_TEXT_COLOR}>
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

function WorkbookCreateScreen({
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
  const [stage, setStage] = useState<"name" | "vocabulary" | "count" | "meaning">(existingWorkbook ? "vocabulary" : "name");
  const [name, setName] = useState(existingWorkbook?.name ?? "");
  const [vocabularyLabel, setVocabularyLabel] = useState(existingWorkbook?.vocabularyLabel ?? "");
  const [vocabularyLanguageCode, setVocabularyLanguageCode] = useState<string | null>(existingWorkbook?.vocabularyLanguageCode ?? null);
  const [meaningCount, setMeaningCount] = useState(existingWorkbook?.meaningAttributes.length ?? 1);
  const [meaningAttributes, setMeaningAttributes] = useState<MeaningAttribute[]>(existingWorkbook?.meaningAttributes ?? [
    { position: 1, label: "Meaning 1", languageCode: null },
  ]);
  const [meaningIndex, setMeaningIndex] = useState(0);
  const [buffer, setBuffer] = useState("");
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
            <Text key={preset.code} color={index === paletteIndex ? "yellow" : AUXILIARY_TEXT_COLOR}>
              {padLine(`${index === paletteIndex ? ">" : " "} ${preset.label} (${preset.code})`, width)}
            </Text>
          ))}
        </>
      ) : null}
      <Text>{padLine("", width)}</Text>
      <Text color={AUXILIARY_TEXT_COLOR}>{padLine(error || "Enter advances. Esc returns to the menu.", width)}</Text>
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

  function handleCreateWorkbook(
    name: string,
    vocabularyLabel: string,
    vocabularyLanguageCode: string | null,
    meaningAttributes: MeaningAttribute[],
  ): void {
    const workbook = backend.createWorkbook(name, vocabularyLabel, vocabularyLanguageCode, meaningAttributes);
    backend.setCurrentWorkbookId(workbook.id);
    const nextWorkbooks = backend.listWorkbooks();
    setWorkbooks(nextWorkbooks);
    const nextIndex = nextWorkbooks.findIndex((item) => item.id === workbook.id);
    setSelectedIndex(nextIndex >= 0 ? nextIndex : (nextWorkbooks.length > 0 ? 0 : 0));
    setScreen({ kind: "menu" });
  }

  function handleUpdateWorkbook(
    workbookId: number,
    _name: string,
    vocabularyLabel: string,
    vocabularyLanguageCode: string | null,
    meaningAttributes: MeaningAttribute[],
  ): void {
    const updated = backend.updateWorkbookSettings(workbookId, vocabularyLabel, vocabularyLanguageCode, meaningAttributes);
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

  if (screen.kind === "vocab") {
    return <VocabularyScreen workbook={screen.workbook} onBackToMenu={backToMenu} onQuit={quit} />;
  }

  if (screen.kind === "create-workbook") {
    return <WorkbookCreateScreen onCreate={handleCreateWorkbook} onCancel={backToMenu} onQuit={quit} />;
  }

  if (screen.kind === "edit-workbook") {
    return (
      <WorkbookCreateScreen
        existingWorkbook={screen.workbook}
        onCreate={(name, label, languageCode, attributes) => handleUpdateWorkbook(screen.workbook.id, name, label, languageCode, attributes)}
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
