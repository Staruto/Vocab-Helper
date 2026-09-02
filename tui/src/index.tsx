import React, { useEffect, useMemo, useState } from "react";
import { Box, Text, render, useApp, useInput, useStdout } from "ink";
import { defaultDbPath, EntryRow, VocabularyRepository } from "./db.js";

type UiMode =
  | { kind: "command" }
  | { kind: "add"; stage: "vocabulary" | "meaning"; vocabulary: string; meaning: string }
  | { kind: "edit"; stage: "vocabulary" | "meaning"; entryId: number; vocabulary: string; meaning: string }
  | { kind: "delete"; entryId: number; label: string };

const repository = new VocabularyRepository(defaultDbPath());

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
    return text.padEnd(width, " ");
  }
  if (width === 1) {
    return "…";
  }
  return `${text.slice(0, width - 1)}…`;
}

function buildEntryLabel(entry: EntryRow): string {
  return `#${entry.id} ${entry.vocabulary}`;
}

function renderEntries(entries: EntryRow[], width: number): string[] {
  const totalWidth = Math.max(width, 60);
  const gap = 3;
  const leftWidth = Math.max(22, Math.floor((totalWidth - gap) * 0.48));
  const rightWidth = Math.max(22, totalWidth - gap - leftWidth);
  const header = `${truncate("Vocabulary", leftWidth)}   ${truncate("Meaning", rightWidth)}`;
  const rows = entries.slice(0, 12).map((entry) => {
    const left = truncate(buildEntryLabel(entry), leftWidth);
    const right = truncate(entry.meaning, rightWidth);
    return `${left}   ${right}`;
  });
  const footer =
    entries.length === 0
      ? "No vocabulary yet. Use /add."
      : entries.length > 12
        ? `Showing 12 of ${entries.length}.`
        : `Total entries: ${entries.length}.`;
  return [header, ...rows, "", footer];
}

function helpText(): string {
  return [
    "Commands:",
    "/list",
    "/add",
    "/edit <id>",
    "/delete <id>",
    "/help",
    "/quit",
    "",
    "Esc cancels add/edit/delete.",
  ].join("\n");
}

function App(): JSX.Element {
  const { exit } = useApp();
  const { stdout } = useStdout();
  const [width, setWidth] = useState(() => stdout?.columns ?? 80);
  const [entries, setEntries] = useState<EntryRow[]>(() => repository.listEntries());
  const [mode, setMode] = useState<UiMode>({ kind: "command" });
  const [buffer, setBuffer] = useState("");
  const [status, setStatus] = useState<string>("Ready.");

  useEffect(() => {
    if (!stdout) {
      return;
    }
    const onResize = () => setWidth(stdout.columns ?? 80);
    stdout.on("resize", onResize);
    return () => {
      stdout.off("resize", onResize);
    };
  }, [stdout]);

  useEffect(() => {
    return () => {
      repository.close();
    };
  }, []);

  function refreshEntries(message?: string): void {
    const next = repository.listEntries();
    setEntries(next);
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
      setStatus(helpText());
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

  const entryLines = useMemo(() => renderEntries(entries, width), [entries, width]);
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

  return (
    <Box flexDirection="column" padding={1}>
      <Text color="cyan">VocabHelper MVP</Text>
      <Text>{status.split("\n")[0] || ""}</Text>
      {status.split("\n").slice(1).map((line, index) => (
        <Text key={`${line}-${index}`}>{line}</Text>
      ))}
      <Box marginTop={1} flexDirection="column">
        {entryLines.map((line, index) => (
          <Text key={`${index}-${line}`}>{line}</Text>
        ))}
      </Box>
      <Box marginTop={1}>
        <Text color="cyan">{promptLabel}: </Text>
        <Text>{buffer}</Text>
        <Text color="gray">▌</Text>
      </Box>
    </Box>
  );
}

function run(): void {
  const interactive = Boolean(process.stdin?.isTTY && process.stdout?.isTTY);
  if (!interactive) {
    const entries = repository.listEntries();
    console.log("VocabHelper MVP");
    console.log(`Loaded ${entries.length} entries.`);
    console.log("Open this app in an interactive terminal to use the TUI.");
    repository.close();
    return;
  }

  render(<App />);
}

run();
