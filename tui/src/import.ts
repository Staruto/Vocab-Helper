import { readFile } from "node:fs/promises";
import { extname, resolve } from "node:path";
import { ImportEntryInput, TagType, WorkbookRow } from "./db.js";

export type ImportDiagnosticKind = "ignored-field" | "ignored-tag" | "invalid" | "duplicate";
export type ImportDiagnostic = { recordNumber: number; kind: ImportDiagnosticKind; message: string };
export type PreparedImportEntry = ImportEntryInput & { recordNumber: number };
export type ImportRecordStatus = "ready" | "invalid" | "duplicate";
export type ImportPreviewRecord = {
  recordNumber: number;
  status: ImportRecordStatus;
  vocabulary: string;
  entry?: PreparedImportEntry;
  ignoredFieldCount: number;
  ignoredTagCount: number;
};
export type ImportPreviewFilter = "records" | "ready" | "invalid" | "duplicates";
export type ImportPreview = {
  totalRecords: number;
  entries: PreparedImportEntry[];
  skippedInvalid: number;
  skippedDuplicates: number;
  diagnostics: ImportDiagnostic[];
  records: ImportPreviewRecord[];
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

export function importFilePathsEqual(left: string, right: string | null): boolean {
  if (!right) return false;
  const resolvedLeft = resolve(left);
  const resolvedRight = resolve(right);
  return process.platform === "win32"
    ? resolvedLeft.toLocaleLowerCase() === resolvedRight.toLocaleLowerCase()
    : resolvedLeft === resolvedRight;
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
  const previewRecords: ImportPreviewRecord[] = [];
  let skippedInvalid = 0;
  let skippedDuplicates = 0;

  records.forEach((record, recordIndex) => {
    const recordNumber = recordIndex + 1;
    let malformed = false;
    let vocabulary = "";
    let ignoredFieldCount = 0;
    let ignoredTagCount = 0;
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
        ignoredFieldCount += 1;
        continue;
      }
      if (matches.length > 1) {
        diagnostics.push({ recordNumber, kind: "ignored-field", message: `ignored ambiguous field '${label}'` });
        ignoredFieldCount += 1;
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
          if (tagId === undefined) ignoredTagCount += 1;
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
      previewRecords.push({ recordNumber, status: "invalid", vocabulary: vocabulary || "<missing vocabulary>", ignoredFieldCount, ignoredTagCount });
      return;
    }
    if (seenVocabulary.has(vocabulary)) {
      diagnostics.push({ recordNumber, kind: "duplicate", message: `skipped duplicate vocabulary '${vocabulary}'` });
      skippedDuplicates += 1;
      previewRecords.push({ recordNumber, status: "duplicate", vocabulary, ignoredFieldCount, ignoredTagCount });
      return;
    }
    seenVocabulary.add(vocabulary);
    const entry = { recordNumber, vocabulary, meanings, attributes, tagIds: [...new Set(tagIds)] };
    entries.push(entry);
    previewRecords.push({ recordNumber, status: "ready", vocabulary, entry, ignoredFieldCount, ignoredTagCount });
  });

  return { totalRecords: records.length, entries, skippedInvalid, skippedDuplicates, diagnostics, records: previewRecords };
}

export function importPreviewRecords(preview: ImportPreview, filter: ImportPreviewFilter): ImportPreviewRecord[] {
  if (filter === "records") return preview.records;
  const status = filter === "duplicates" ? "duplicate" : filter;
  return preview.records.filter((record) => record.status === status);
}

export function formatImportPreviewFilter(label: string, count: number, selected: boolean): string {
  const text = `${label}: ${count}`;
  return selected ? `[${text}]` : text;
}

export function paginateImportPreview(preview: ImportPreview, filter: ImportPreviewFilter, requestedPage: number, pageSize: number): { records: ImportPreviewRecord[]; pageIndex: number; pageCount: number } {
  const filtered = importPreviewRecords(preview, filter);
  const safePageSize = Math.max(1, pageSize);
  const pageCount = Math.max(1, Math.ceil(filtered.length / safePageSize));
  const pageIndex = Math.max(0, Math.min(requestedPage, pageCount - 1));
  return { records: filtered.slice(pageIndex * safePageSize, pageIndex * safePageSize + safePageSize), pageIndex, pageCount };
}

export function shouldPromptToSaveImportPath(importedPath: string, savedPath: string | null): boolean {
  return !importFilePathsEqual(importedPath, savedPath);
}

export function formatImportPreviewRecord(record: ImportPreviewRecord): string {
  const marker = record.status === "ready" ? "+" : record.status === "invalid" ? "!" : "=";
  const suffix = [
    record.ignoredFieldCount ? `+${record.ignoredFieldCount} ignored field${record.ignoredFieldCount === 1 ? "" : "s"}` : "",
    record.ignoredTagCount ? `+${record.ignoredTagCount} ignored tag${record.ignoredTagCount === 1 ? "" : "s"}` : "",
  ].filter(Boolean);
  return `${marker} ${record.vocabulary}${suffix.length ? ` (${suffix.join(", ")})` : ""}`;
}

export function buildImportPreviewLines(preview: ImportPreview, filter: ImportPreviewFilter = "records"): string[] {
  const ignoredFields = preview.diagnostics.filter((item) => item.kind === "ignored-field").length;
  const ignoredTags = preview.diagnostics.filter((item) => item.kind === "ignored-tag").length;
  const rows = importPreviewRecords(preview, filter).map(formatImportPreviewRecord);
  return [
    `Records: ${preview.totalRecords} | Ready: ${preview.entries.length} | Invalid: ${preview.skippedInvalid} | Duplicates: ${preview.skippedDuplicates}`,
    `Ignored fields: ${ignoredFields} | Ignored tags: ${ignoredTags}`,
    ...(rows.length ? rows : ["No records in this category."]),
  ];
}
