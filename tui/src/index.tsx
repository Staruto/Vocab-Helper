import React, { useEffect, useMemo, useState } from "react";
import { Box, Text, render, useApp, useInput, useStdout } from "ink";
import { defaultDbPath, EntryRow, VocabularyRepository } from "./db.js";

type UiMode =
  | { kind: "command" }
  | { kind: "add"; stage: "vocabulary" | "meaning"; vocabulary: string; meaning: string }
  | { kind: "edit"; stage: "vocabulary" | "meaning"; entryId: number; vocabulary: string; meaning: string }
  | { kind: "delete"; entryId: number; label: string };

const PAGE_SIZE = 20;
const TITLE = "VocabHelper MVP";
const FOOTER_HINT = "Navigate pages with <- ->";
const repository = new VocabularyRepository(defaultDbPath());

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

function buildTableLines(entries: EntryRow[], pageIndex: number, width: number, availableRows: number): string[] {
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
  const header = `|${padLine("Vocabulary", leftWidth)}|${padLine("Meaning", rightWidth)}|`;

  const rows = pageEntries.map((entry) => {
    const left = padLine(buildEntryLabel(entry), leftWidth);
    const right = padLine(entry.meaning, rightWidth);
    return `|${left}|${right}|`;
  });

  return [border, header, border, ...rows, border];
}

function buildHelpText(): string {
  return "Commands: /list /add /edit <id> /delete <id> /help /quit | Esc cancels forms.";
}

function App(): JSX.Element {
  const { exit } = useApp();
  const { stdout } = useStdout();
  const [width, setWidth] = useState(() => stdout?.columns ?? 80);
  const [rows, setRows] = useState(() => stdout?.rows ?? 24);
  const [entries, setEntries] = useState<EntryRow[]>(() => repository.listEntries());
  const [pageIndex, setPageIndex] = useState(0);
  const [mode, setMode] = useState<UiMode>({ kind: "command" });
  const [buffer, setBuffer] = useState("");
  const [status, setStatus] = useState<string>("Ready.");

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
    return () => {
      leaveAlternateScreen();
      repository.close();
    };
  }, []);

  useEffect(() => {
    setPageIndex((current) => clampPageIndex(current, entries.length));
  }, [entries.length]);

  function refreshEntries(message?: string): void {
    const next = repository.listEntries();
    setEntries(next);
    setPageIndex((current) => clampPageIndex(current, next.length));
    setStatus(message ?? `Loaded ${next.length} entr${next.length === 1 ? "y" : "ies"}.`);
  }

  function beginAdd(): void {
    setMode({ kind: "add", stage: "vocabulary", vocabulary: "", meaning: "" });
    setBuffer("");
    setStatus("Adding a new entry.");
  }

  function beginEdit(entryId: number): void {
    const entry = repository.getEntry(entryId);
    if (!entry) {
      setStatus(`Entry #${entryId} was not found.`);
      return;
    }

    setMode({
      kind: "edit",
      stage: "vocabulary",
      entryId,
      vocabulary: entry.vocabulary,
      meaning: entry.meaning,
    });
    setBuffer(entry.vocabulary);
    setStatus(`Editing #${entryId}.`);
  }

  function beginDelete(entryId: number): void {
    const entry = repository.getEntry(entryId);
    if (!entry) {
      setStatus(`Entry #${entryId} was not found.`);
      return;
    }

    setMode({ kind: "delete", entryId, label: buildEntryLabel(entry) });
    setBuffer("");
    setStatus(`Type yes to delete ${buildEntryLabel(entry)}.`);
  }

  function cancelActiveMode(message = "Cancelled."): void {
    setMode({ kind: "command" });
    setBuffer("");
    setStatus(message);
  }

  function submitCommand(raw: string): void {
    const parts = normalizeCommand(raw);
    if (parts.length === 0) {
      return;
    }

    const [command, ...args] = parts;
    const lower = command.toLowerCase();

    if (lower === "help") {
      setStatus(buildHelpText());
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
        setStatus("Usage: /edit <id>");
        return;
      }
      beginEdit(entryId);
      return;
    }

    if (lower === "delete") {
      const entryId = Number(args[0]);
      if (!args[0] || Number.isNaN(entryId)) {
        setStatus("Usage: /delete <id>");
        return;
      }
      beginDelete(entryId);
      return;
    }

    if (lower === "quit" || lower === "exit") {
      leaveAlternateScreen();
      repository.close();
      exit();
      return;
    }

    setStatus(`Unknown command: ${command}`);
  }

  function submitForm(value: string): void {
    const text = value.trim();

    if (mode.kind === "add") {
      if (mode.stage === "vocabulary") {
        if (!text) {
          setStatus("Vocabulary is required.");
          return;
        }

        setMode({
          kind: "add",
          stage: "meaning",
          vocabulary: text,
          meaning: "",
        });
        setBuffer("");
        setStatus("Enter meaning.");
        return;
      }

      if (!text) {
        setStatus("Meaning is required.");
        return;
      }

      const entry = repository.addEntry(mode.vocabulary, text);
      setMode({ kind: "command" });
      setBuffer("");
      refreshEntries(`Added #${entry.id}.`);
      return;
    }

    if (mode.kind === "edit") {
      if (mode.stage === "vocabulary") {
        if (!text) {
          setStatus("Vocabulary is required.");
          return;
        }

        setMode({
          kind: "edit",
          stage: "meaning",
          entryId: mode.entryId,
          vocabulary: text,
          meaning: mode.meaning,
        });
        setBuffer(mode.meaning);
        setStatus("Update meaning.");
        return;
      }

      if (!text) {
        setStatus("Meaning is required.");
        return;
      }

      const entry = repository.updateEntry(mode.entryId, mode.vocabulary, text);
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

      repository.deleteEntry(mode.entryId);
      setMode({ kind: "command" });
      setBuffer("");
      refreshEntries(`Deleted ${mode.label}.`);
    }
  }

  useInput((input, key) => {
    if (key.ctrl && input === "c") {
      leaveAlternateScreen();
      repository.close();
      exit();
      return;
    }

    if (key.escape) {
      if (mode.kind === "command") {
        setBuffer("");
        setStatus("Ready.");
      } else {
        cancelActiveMode();
      }
      return;
    }

    if (key.leftArrow && mode.kind === "command") {
      setPageIndex((current) => Math.max(0, current - 1));
      return;
    }

    if (key.rightArrow && mode.kind === "command") {
      setPageIndex((current) => Math.min(getPageCount(entries.length) - 1, current + 1));
      return;
    }

    if (key.backspace || key.delete) {
      setBuffer((current) => current.slice(0, -1));
      return;
    }

    if (key.return) {
      const current = buffer;
      if (mode.kind === "command") {
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
    () => buildTableLines(entries, safePageIndex, width, Math.max(1, rows - 6)),
    [entries, safePageIndex, rows, width],
  );
  const promptLabel =
    mode.kind === "command"
      ? "Command"
      : mode.kind === "add"
        ? mode.stage === "vocabulary"
          ? "Vocabulary"
          : "Meaning"
        : mode.kind === "edit"
          ? mode.stage === "vocabulary"
            ? "Vocabulary"
            : "Meaning"
          : "Confirm delete";
  const promptLine = `${promptLabel}: ${buffer}_`;
  const statusLine = status.split("\n")[0] ?? "";

  const bodyLines = useMemo(() => {
    const lines: string[] = [];
    lines.push("");
    for (const line of tableLines) {
      lines.push(padLine(line, width));
    }

    const reservedLines = 2;
    while (lines.length < Math.max(0, rows - reservedLines - 3)) {
      lines.push("");
    }

    lines.push(padLine("", width));
    lines.push(padLine(buildFooterLine(width, pageText, FOOTER_HINT), width));
    return lines;
  }, [pageText, rows, tableLines, width]);

  return (
    <Box flexDirection="column">
      <Text color="cyan" bold>
        {centerLine(TITLE, width)}
      </Text>
      <Text color="cyan">{padLine(promptLine, width)}</Text>
      <Text>{padLine(statusLine, width)}</Text>
      {bodyLines.map((line, index) => (
        <Text key={`${index}-${line}`}>{line}</Text>
      ))}
    </Box>
  );
}

function run(): void {
  const interactive = Boolean(process.stdin?.isTTY && process.stdout?.isTTY);
  if (!interactive) {
    const entries = repository.listEntries();
    console.log(TITLE);
    console.log(`Loaded ${entries.length} entries.`);
    console.log("Open this app in an interactive terminal to use the TUI.");
    repository.close();
    return;
  }

  enterAlternateScreen();
  render(<App />);
}

run();
