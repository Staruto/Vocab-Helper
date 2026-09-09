import { readFile } from "node:fs/promises";
import { extname, resolve } from "node:path";
import { ImportEntryInput, TagType, WorkbookRow } from "./db.js";

export type ImportDiagnosticKind = "ignored-field" | "ignored-tag" | "invalid" | "duplicate";
export type ImportDiagnostic = { recordNumber: number; kind: ImportDiagnosticKind; message: string };
export type PreparedImportEntry = ImportEntryInput & { recordNumber: number };
export type ImportPreview = {
  totalRecords: number;
  entries: PreparedImportEntry[];
  skippedInvalid: number;
  skippedDuplicates: number;
  diagnostics: ImportDiagnostic[];
};

type Destination =
  | { kind: "vocabulary" }
  | { kind: "meaning"; index: number }
  | { kind: "optional"; key: string }
  | { kind: "tags"; type: TagType };

export async function loadLabeledTextFile(inputPath: string): Promise<{ path: string; text: string }> {
  const path = resolve(inputPath.trim());
  if (extname(path).toLocaleLowerCase() !== ".txt") throw new Error("Import supports .txt files only.");
  const bytes = await readFile(path);
  let text: string;
  try {
    text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    throw new Error("Import files must contain valid UTF-8 text.");
  }
  return { path, text: text.replace(/^\uFEFF/, "") };
}

export function parseLabeledTextImport(text: string, workbook: WorkbookRow, tagTypes: TagType[], existingVocabulary: string[]): ImportPreview {
  const normalized = text.replace(/\r\n?/g, "\n").trim();
  const records = normalized ? normalized.split(/\n[ \t]*\n+/) : [];
  const destinations = new Map<string, Destination[]>();
  const addDestination = (label: string, destination: Destination) => destinations.set(label, [...(destinations.get(label) ?? []), destination]);
  addDestination(workbook.vocabularyLabel, { kind: "vocabulary" });
  workbook.meaningAttributes.forEach((field, index) => addDestination(field.label, { kind: "meaning", index }));
  workbook.metadataAttributes.filter((field) => field.role === "optional").forEach((field) => addDestination(field.label, { kind: "optional", key: field.key }));
  tagTypes.forEach((type) => addDestination(type.name, { kind: "tags", type }));

  const seenVocabulary = new Set(existingVocabulary.map((value) => value.trim()));
  const entries: PreparedImportEntry[] = [];
  const diagnostics: ImportDiagnostic[] = [];
  let skippedInvalid = 0;
  let skippedDuplicates = 0;

  records.forEach((record, recordIndex) => {
    const recordNumber = recordIndex + 1;
    let malformed = false;
    let vocabulary = "";
    const meanings = workbook.meaningAttributes.map(() => "");
    const attributes: Record<string, string> = {};
    const tagIds: number[] = [];
    const recognizedLabels = new Set<string>();

    for (const rawLine of record.split("\n")) {
      const line = rawLine.trim();
      if (!line) continue;
      const separator = line.indexOf(":");
      if (separator < 0) {
        diagnostics.push({ recordNumber, kind: "invalid", message: `line has no ':' separator: ${line}` });
        malformed = true;
        continue;
      }
      const label = line.slice(0, separator).trim();
      const value = line.slice(separator + 1).trim();
      const matches = destinations.get(label) ?? [];
      if (matches.length === 0) {
        diagnostics.push({ recordNumber, kind: "ignored-field", message: `ignored field '${label}'` });
        continue;
      }
      if (matches.length > 1) {
        diagnostics.push({ recordNumber, kind: "ignored-field", message: `ignored ambiguous field '${label}'` });
        continue;
      }
      if (recognizedLabels.has(label)) {
        diagnostics.push({ recordNumber, kind: "invalid", message: `recognized field '${label}' appears more than once` });
        malformed = true;
        continue;
      }
      recognizedLabels.add(label);
      const destination = matches[0];
      if (destination.kind === "vocabulary") vocabulary = value;
      else if (destination.kind === "meaning") meanings[destination.index] = value;
      else if (destination.kind === "optional") attributes[destination.key] = value;
      else {
        const tagsByName = new Map(destination.type.tags.map((tag) => [tag.name, tag.id]));
        for (const tagName of value.split(",").map((item) => item.trim()).filter(Boolean)) {
          const tagId = tagsByName.get(tagName);
          if (tagId === undefined) diagnostics.push({ recordNumber, kind: "ignored-tag", message: `ignored ${destination.type.name} tag '${tagName}'` });
          else tagIds.push(tagId);
        }
      }
    }

    if (!vocabulary) {
      diagnostics.push({ recordNumber, kind: "invalid", message: `missing required field '${workbook.vocabularyLabel}'` });
      malformed = true;
    }
    const primaryLabel = workbook.meaningAttributes[0]?.label ?? "Primary Meaning";
    if (!meanings[0]) {
      diagnostics.push({ recordNumber, kind: "invalid", message: `missing required field '${primaryLabel}'` });
      malformed = true;
    }
    if (malformed) {
      skippedInvalid += 1;
      return;
    }
    if (seenVocabulary.has(vocabulary)) {
      diagnostics.push({ recordNumber, kind: "duplicate", message: `skipped duplicate vocabulary '${vocabulary}'` });
      skippedDuplicates += 1;
      return;
    }
    seenVocabulary.add(vocabulary);
    entries.push({ recordNumber, vocabulary, meanings, attributes, tagIds: [...new Set(tagIds)] });
  });

  return { totalRecords: records.length, entries, skippedInvalid, skippedDuplicates, diagnostics };
}

export function buildImportPreviewLines(preview: ImportPreview): string[] {
  const ignoredFields = preview.diagnostics.filter((item) => item.kind === "ignored-field").length;
  const ignoredTags = preview.diagnostics.filter((item) => item.kind === "ignored-tag").length;
  return [
    `Records: ${preview.totalRecords} | Ready: ${preview.entries.length} | Invalid: ${preview.skippedInvalid} | Duplicates: ${preview.skippedDuplicates}`,
    `Ignored fields: ${ignoredFields} | Ignored tags: ${ignoredTags}`,
    ...preview.diagnostics.map((item) => `Record ${item.recordNumber}: ${item.message}`),
  ];
}
