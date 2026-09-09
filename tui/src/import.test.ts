import assert from "node:assert/strict";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import test from "node:test";
import { TagType, WorkbookRow } from "./db.js";
import { buildImportPreviewLines, importFilePathsEqual, loadLabeledTextFile, parseLabeledTextImport } from "./import.js";

const workbook: WorkbookRow = {
  id: 1, name: "Japanese", wordCount: 0, createdAt: "2026-01-01",
  vocabularyLabel: "Japanese", vocabularyLanguageCode: "JP", presetEnabled: true, vocabularyKind: "preset_language",
  importFilePath: null,
  meaningAttributes: [
    { id: 1, key: "meaning_1", position: 1, label: "English", languageCode: "EN" },
    { id: 2, key: "meaning_2", position: 2, label: "Chinese", languageCode: "ZH" },
    { id: 3, key: "meaning_3", position: 3, label: "日本語訳 (Simple JPN)", languageCode: "JP" },
  ],
  metadataAttributes: [
    { key: "vocab", role: "vocabulary", label: "Japanese", languageCode: "JP", required: true, visible: true, displayOrder: 0 },
    { id: 1, key: "meaning_1", role: "meaning", label: "English", languageCode: "EN", required: true, visible: true, displayOrder: 1 },
    { id: 2, key: "meaning_2", role: "meaning", label: "Chinese", languageCode: "ZH", required: false, visible: false, displayOrder: 2 },
    { id: 3, key: "meaning_3", role: "meaning", label: "日本語訳 (Simple JPN)", languageCode: "JP", required: false, visible: false, displayOrder: 3 },
    { id: 4, key: "kana", role: "optional", label: "読み方 (kana)", languageCode: "JP", required: false, visible: false, displayOrder: 4 },
    { id: 5, key: "example_sentence_1", role: "optional", label: "Example Sentence 1", languageCode: "JP", required: false, visible: false, displayOrder: 5 },
    { id: 6, key: "example_sentence_2", role: "optional", label: "Example Sentence 2", languageCode: "JP", required: false, visible: false, displayOrder: 6 },
  ],
};

const tagTypes: TagType[] = [{
  id: 1, workbookId: 1, name: "Part of Speech", position: 1, visible: true,
  tags: [{ id: 10, tagTypeId: 1, name: "名詞" }, { id: 11, tagTypeId: 1, name: "動詞" }],
}];

test("labeled text maps the supplied Japanese example by exact label regardless of order", () => {
  const preview = parseLabeledTextImport(`Japanese: 連休
読み方 (kana): れんきゅう
English: consecutive holidays / long weekend
Chinese: 连休，连续假期
日本語訳 (Simple JPN): 何日も続いて休みの日があること。
Part of Speech: 名詞
Example Sentence 1: 来週は連休です。 (Next week is a long weekend.)
Example Sentence 2: 連休に海へ行きます。 (I will go to the sea during the consecutive holidays.)`, workbook, tagTypes, []);

  assert.equal(preview.totalRecords, 1);
  assert.equal(preview.entries.length, 1);
  assert.deepEqual(preview.entries[0], {
    recordNumber: 1,
    vocabulary: "連休",
    meanings: ["consecutive holidays / long weekend", "连休，连续假期", "何日も続いて休みの日があること。"],
    attributes: {
      kana: "れんきゅう",
      example_sentence_1: "来週は連休です。 (Next week is a long weekend.)",
      example_sentence_2: "連休に海へ行きます。 (I will go to the sea during the consecutive holidays.)",
    },
    tagIds: [10],
  });
  assert.deepEqual(preview.diagnostics, []);
});

test("records report ignored data and skip malformed, missing, and duplicate entries", () => {
  const preview = parseLabeledTextImport(`\uFEFFJapanese: 猫\r
English: cat: a small feline\r
Unused: value\r
Part of Speech: 名詞, unknown\r
\r
Japanese: 犬\r
English: dog\r
\r
Japanese: 犬\r
English: duplicate\r
\r
Japanese: 鳥\r
bad line\r
English: bird\r
\r
Japanese: 魚`, workbook, tagTypes, ["猫"]);

  assert.deepEqual(preview.entries.map((entry) => entry.vocabulary), ["犬"]);
  assert.equal(preview.skippedDuplicates, 2);
  assert.equal(preview.skippedInvalid, 2);
  assert.match(buildImportPreviewLines(preview).join("\n"), /ignored field 'Unused'/);
  assert.match(buildImportPreviewLines(preview).join("\n"), /ignored Part of Speech tag 'unknown'/);
  assert.match(buildImportPreviewLines(preview).join("\n"), /line has no ':' separator/);
  assert.match(buildImportPreviewLines(preview).join("\n"), /missing required field 'English'/);
});

test("recognized labels are exact, unique, and unambiguous", () => {
  const ambiguousWorkbook = { ...workbook, meaningAttributes: [{ ...workbook.meaningAttributes[0], label: "Japanese" }] };
  const ambiguous = parseLabeledTextImport("Japanese: value", ambiguousWorkbook, tagTypes, []);
  assert.equal(ambiguous.entries.length, 0);
  assert.match(ambiguous.diagnostics.map((item) => item.message).join("\n"), /ambiguous field 'Japanese'/);

  const repeated = parseLabeledTextImport("Japanese: 猫\nJapanese: 犬\nEnglish: cat", workbook, tagTypes, []);
  assert.equal(repeated.entries.length, 0);
  assert.match(repeated.diagnostics.map((item) => item.message).join("\n"), /appears more than once/);

  const wrongCase = parseLabeledTextImport("japanese: 猫\nEnglish: cat", workbook, tagTypes, []);
  assert.equal(wrongCase.entries.length, 0);
  assert.match(wrongCase.diagnostics.map((item) => item.message).join("\n"), /ignored field 'japanese'/);

  const multipleTags = parseLabeledTextImport("Japanese: 走る\nEnglish: run\nPart of Speech: 名詞, 動詞", workbook, tagTypes, []);
  assert.deepEqual(multipleTags.entries[0].tagIds, [10, 11]);
});

test("the file loader accepts UTF-8 txt files and rejects other extensions and invalid UTF-8", async () => {
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-import-"));
  try {
    const valid = join(directory, "words.txt");
    writeFileSync(valid, Buffer.from("\uFEFFJapanese: 猫", "utf8"));
    assert.equal((await loadLabeledTextFile(valid)).text, "Japanese: 猫");
    await assert.rejects(() => loadLabeledTextFile(join(directory, "words.csv")), /\.txt files only/);
    await assert.rejects(() => loadLabeledTextFile(join(directory, "missing.txt")), /ENOENT/);
    const invalid = join(directory, "invalid.txt");
    writeFileSync(invalid, Buffer.from([0xc3, 0x28]));
    await assert.rejects(() => loadLabeledTextFile(invalid), /valid UTF-8/);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});

test("empty files produce an empty, non-importable preview", () => {
  const preview = parseLabeledTextImport(" \r\n\r\n ", workbook, tagTypes, []);
  assert.equal(preview.totalRecords, 0);
  assert.deepEqual(preview.entries, []);
});

test("import file path comparison resolves equivalent paths", () => {
  assert.equal(importFilePathsEqual("words.txt", resolve("words.txt")), true);
  assert.equal(importFilePathsEqual("words.txt", null), false);
  if (process.platform === "win32") assert.equal(importFilePathsEqual("C:\\IMPORTS\\WORDS.TXT", "c:\\imports\\words.txt"), true);
});
