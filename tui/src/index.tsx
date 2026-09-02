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
const STATUS_HINT = "Esc cancels forms.";
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

function truncate(text: string, width: number): string {
  if (width <= 0) {
    return "";
  }
  if (text.length <= width) {
    return text;
  }
  if (width === 1) {
    return ".";
  }
  return `${text.slice(0, width - 1)}.`;
}

function padLine(text: string, width: number): string {
  const clipped = truncate(text, width);
  return clipped.padEnd(width, " ");
}

function centerLine(text: string, width: number): string {
  if (width <= 0) {
    return "";
  }

  const clipped = truncate(text, width);
  const leftPadding = Math.max(0, Math.floor((width - clipped.length) / 2));
  const rightPadding = Math.max(0, width - clipped.length - leftPadding);
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
  if (left.length + right.length + 2 >= width) {
    const gap = Math.max(1, width - left.length);
    return `${left}${" ".repeat(gap)}${truncate(right, width - left.length - gap)}`;
  }

  return `${left}${" ".repeat(width - left.length - right.length)}${right}`;
}

function buildEntryLines(entries: EntryRow[], pageIndex: number, width: number): string[] {
  const totalWidth = Math.max(width, 80);
  const gap = 4;
  const minimumLeft = 28;
  const minimumRight = 28;
  const desiredLeft = Math.floor(totalWidth * 0.56);
  const maxLeft = totalWidth - gap - minimumRight;
  const leftWidth = Math.max(minimumLeft, Math.min(desiredLeft, maxLeft));
  const rightWidth = Math.max(minimumRight, totalWidth - gap - leftWidth);

  const header = `${truncate("Vocabulary", leftWidth)}${" ".repeat(gap)}${truncate("Meaning", rightWidth)}`;
  const pageEntries = entries.slice(pageIndex * PAGE_SIZE, pageIndex * PAGE_SIZE + PAGE_SIZE);
  const rows = pageEntries.map((entry) => {
    const left = truncate(buildEntryLabel(entry), leftWidth);
    const right = truncate(entry.meaning, rightWidth);
    return `${left}${" ".repeat(gap)}${right}`;
  });

  return [header, ...rows];
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
  const entryLines = useMemo(
    () => buildEntryLines(entries, safePageIndex, width),
    [entries, safePageIndex, width],
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
    for (const line of entryLines) {
      lines.push(padLine(line, width));
    }

    const reservedLines = 2;
    while (lines.length < Math.max(0, rows - reservedLines - 3)) {
      lines.push("");
    }

    lines.push(padLine("", width));
    lines.push(padLine(buildFooterLine(width, pageText, FOOTER_HINT), width));
    return lines;
  }, [entryLines, pageText, rows, width]);

  return (
    <Box flexDirection="column">
      <Text color="magenta" bold>
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
