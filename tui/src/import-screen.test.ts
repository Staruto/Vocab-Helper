import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { once } from "node:events";
import { mkdtempSync, rmdirSync, unlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { setTimeout as delay } from "node:timers/promises";
import { stripVTControlCharacters } from "node:util";
import { VocabularyRepository } from "./db.js";

test("import keyboard confirmation stages writes and preview header follows filters and focus", { timeout: 30_000 }, async () => {
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-import-screen-"));
  const path = join(directory, "vocab.db");
  const file = join(directory, "words.txt");
  const repository = new VocabularyRepository(path);
  const created = repository.createWorkbook("Import keyboard test");
  const workbook = repository.setWorkbookImportFilePath(created.id, file);
  const entry = repository.addEntry(workbook.id, "existing", "old");
  repository.completePracticeSession(workbook.id, [{ entryId: entry.id, answers: [false, true] }], "2026-10-25");
  const stats = repository.getEntryStats(entry.id);
  const record = (word: string, meaning: string) => `${workbook.vocabularyLabel}: ${word}\n${workbook.meaningAttributes[0].label}: ${meaning}`;
  writeFileSync(file, [record("existing", "new"), ...Array.from({ length: 10 }, (_, i) => record(`word${i}`, "meaning")), `${workbook.vocabularyLabel}: invalid`].join("\n\n"));
  repository.close();
  const source = `
    process.stdin.isTTY = true;
    process.stdin.setRawMode = () => {};
    process.stdout.isTTY = true;
    process.stdout.columns = 80;
    process.stdout.rows = 24;
    await import(${JSON.stringify(new URL("./index.tsx", import.meta.url).href)});
  `;
  const child = spawn(process.execPath, ["--import", "tsx", "--input-type=module", "-e", source], {
    env: { ...process.env, VOCAB_HELPER_DB_PATH: path, FORCE_COLOR: "3", TERM: "xterm-256color", CI: "false", NODE_NO_WARNINGS: "1" },
    stdio: ["pipe", "pipe", "pipe"],
  });
  let output = ""; let stderr = "";
  child.stdout.on("data", (chunk: Buffer) => { output += chunk.toString(); });
  child.stderr.on("data", (chunk: Buffer) => { stderr += chunk.toString(); });
  async function frame(expected: RegExp): Promise<string> {
    const deadline = Date.now() + 5_000;
    while (!expected.test(stripVTControlCharacters(output))) {
      assert.equal(child.exitCode, null, `App exited: ${stderr}`);
      if (Date.now() > deadline) assert.fail(`Missing ${expected}: ${stripVTControlCharacters(output)}\n${stderr}`);
      await delay(20);
    }
    await delay(80);
    const start = Math.max(output.lastIndexOf("\x1b[G"), output.lastIndexOf("\x1b[H"));
    return output.slice(start < 0 ? 0 : start + 3);
  }
  async function key(input: string, expected: RegExp): Promise<string> {
    output = ""; child.stdin.write(input); return frame(expected);
  }
  const down = "\x1b[B"; const right = "\x1b[C"; const up = "\x1b[A";
  try {
    await frame(/Selected: Import keyboard test/);
    await key("\r", /Page 1\/1/);
    await key("/import", /\/import/);
    let rendered = await key("\r", /Preview\s+12 words/);
    assert.match(stripVTControlCharacters(rendered), /Filters\s+\[All\]\s+Ready\s+Invalid\s+Duplicates/);
    let header = rendered.split("\n").find((line) => line.includes("Preview"))!;
    assert.match(header, /\x1b\[37mPreview/);
    assert.match(stripVTControlCharacters(header), /^│ Preview +12 words │$/);
    rendered = await key(down, /Preview\s+12 words/);
    header = rendered.split("\n").find((line) => line.includes("Preview"))!;
    assert.match(header, /38;2;206;168;255mPreview/);
    rendered = await key(right, /Page 2\/2/);
    assert.match(stripVTControlCharacters(rendered), /Preview\s+12 words/);
    assert.ok(stripVTControlCharacters(rendered).trimEnd().split("\n").length <= 24);
    await key(up, /\[All\]/);
    await key(right, /Preview\s+10 words/);
    await key(right, /Preview\s+1 word/);
    await key(right, /\[Duplicates\]/);
    await key("\r", /Unsaved import changes/);
    assert.match(stripVTControlCharacters(output), /> Save Changes \(Preserve Test Data\)/);
    const verify = new VocabularyRepository(path);
    try {
      assert.equal(verify.getEntry(entry.id)!.meaning, "old");
      assert.equal(verify.countEntries(workbook.id), 1);
    } finally { verify.close(); }
    await key("\x1b", /\[Duplicates\]/);
    await key("\r", /Unsaved import changes/);
    await key(down, /> Save Changes \(Reset Test Data\)/);
    await key(down, /> Discard Changes/);
    await key(down, /> Return to Preview/);
    await key("\r", /\[Duplicates\]/);
    await key("\r", /Unsaved import changes/);
    await key(down, /> Save Changes \(Reset Test Data\)/);
    await key(down, /> Discard Changes/);
    await key("\r", /Import discarded/);
    await key("/import", /\/import/);
    await key("\r", /Preview\s+12 words/);
    await key("\r", /Unsaved import changes/);
    await key("\r", /Page 1\/1/);
    const saved = new VocabularyRepository(path);
    try {
      assert.equal(saved.getEntry(entry.id)!.meaning, "new");
      assert.deepEqual(saved.getEntryStats(entry.id), stats);
      assert.equal(saved.countEntries(workbook.id), 11);
    } finally { saved.close(); }
  } finally {
    if (child.exitCode === null && child.signalCode === null) {
      const closed = once(child, "close"); child.kill(); await closed;
    }
    unlinkSync(path); unlinkSync(file); rmdirSync(directory);
  }
});
