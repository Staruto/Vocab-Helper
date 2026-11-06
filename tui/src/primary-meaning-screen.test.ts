import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { once } from "node:events";
import { mkdtempSync, unlinkSync, rmdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { setTimeout as delay } from "node:timers/promises";
import { stripVTControlCharacters } from "node:util";
import { VocabularyRepository } from "./db.js";

test("primary meaning visibility toggles in settings and blank meanings remain usable in practice", { timeout: 30_000 }, async () => {
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-meaning-screen-"));
  const path = join(directory, "vocab.db");
  const repository = new VocabularyRepository(path);
  const workbook = repository.createWorkbook("Optional meaning test");
  repository.addEntry(workbook.id, "word", "");
  repository.close();
  const source = `
    process.stdin.isTTY = true;
    process.stdin.setRawMode = () => {};
    process.stdout.isTTY = true;
    process.stdout.columns = 100;
    process.stdout.rows = 35;
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
    return stripVTControlCharacters(output.slice(start < 0 ? 0 : start + 3));
  }
  async function key(input: string, expected: RegExp): Promise<string> {
    output = ""; child.stdin.write(input); return frame(expected);
  }
  async function command(raw: string, expected: RegExp): Promise<string> {
    await key(raw, new RegExp(raw.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
    return key("\r", expected);
  }
  try {
    await frame(/Selected: Optional meaning test/);
    let rendered = await key("\r", /Page 1\/1/);
    assert.match(rendered, /Primary Meaning/);
    await command("/setting", /Settings.*Attributes.*Tags/);
    await key("\x1b[B", /> Vocabulary|> Word/);
    await key("\x1b[B", /> Primary Meaning.*\[shown\]/);
    rendered = await key(" ", /> Primary Meaning.*\[hidden\]/);
    assert.doesNotMatch(rendered, /primary, required/);
    await key("\x1b", /Unsaved settings changes/);
    rendered = await key("\r", /Page 1\/1/);
    assert.match(rendered, /word/);
    assert.doesNotMatch(rendered, /Primary Meaning/);
    const verify = new VocabularyRepository(path);
    try { assert.equal(verify.getWorkbook(workbook.id)?.meaningAttributes[0].visible, false); }
    finally { verify.close(); }
    await command("/test 1", /Primary Meaning: -/);
    await key("\x1b", /Page 1\/1/);
    await command("/setting", /Settings.*Attributes.*Tags/);
    await key("\x1b[B", /> Vocabulary|> Word/);
    await key("\x1b[B", /> Primary Meaning.*\[hidden\]/);
    await key(" ", /> Primary Meaning.*\[shown\]/);
    await key("\x1b", /Unsaved settings changes/);
    rendered = await key("\r", /Page 1\/1/);
    assert.match(rendered, /Primary Meaning/);
    assert.equal(stderr, "");
  } finally {
    if (child.exitCode === null && child.signalCode === null) {
      const closed = once(child, "close"); child.kill(); await closed;
    }
    unlinkSync(path); rmdirSync(directory);
  }
});

test("promoting a meaning keeps list columns and search paired with their values", { timeout: 30_000 }, async () => {
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-meaning-screen-"));
  const path = join(directory, "vocab.db");
  const repository = new VocabularyRepository(path);
  const workbook = repository.createWorkbook("Optional meaning test", "Word", null, [{ position: 1, label: "English", languageCode: "en", visible: true }, { position: 2, label: "Japanese", languageCode: "ja", visible: false }]);
  repository.addEntry(workbook.id, "neko", "cat", ["cat", "neko-jp"]);
  repository.close();
  const source = `
    process.stdin.isTTY = true;
    process.stdin.setRawMode = () => {};
    process.stdout.isTTY = true;
    process.stdout.columns = 100;
    process.stdout.rows = 35;
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
    return stripVTControlCharacters(output.slice(start < 0 ? 0 : start + 3));
  }
  async function key(input: string, expected: RegExp): Promise<string> {
    output = ""; child.stdin.write(input); return frame(expected);
  }
  async function command(raw: string, expected: RegExp): Promise<string> {
    await key(raw, new RegExp(raw.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
    return key("\r", expected);
  }
  try {
    await frame(/Selected: Optional meaning test/);
    let rendered = await key("\r", /Page 1\/1/);
    assert.match(rendered, /\|\s*cat\s*\|/);
    await command("/setting", /Settings.*Attributes.*Tags/);
    await key("\x1b[B", /> Vocabulary|> Word/);
    await key("\x1b[B", /> English/);
    await key(" ", /> English.*\[hidden\]/);
    await key("\x1b[B", /> Japanese/);
    await key("\x10", /Japanese will become the Primary Meaning/);
    await key(" ", /> Japanese.*\[shown\]/);
    await key("\x1b", /Unsaved settings changes/);
    rendered = await key("\r", /Page 1\/1/);
    assert.match(rendered, /Japanese/);
    assert.match(rendered, /\|\s*neko-jp\s*\|/);
    assert.doesNotMatch(rendered, /\|\s*cat\s*\|/);
    await command("/search neko-jp", /1 matches/);
    await key("\x1b", /Page 1\/1/);
    await command("/search cat", /0 matches/);
    await key("\x1b", /Page 1\/1/);
    await command("/setting", /Settings.*Attributes.*Tags/);
    await key("\x1b[B", /> Vocabulary|> Word/);
    await key("\x1b[B", /> Japanese/);
    await key(" ", /> Japanese.*\[hidden\]/);
    await key("\x1b[B", /> English/);
    await key(" ", /> English.*\[shown\]/);
    await key("\x1b", /Unsaved settings changes/);
    rendered = await key("\r", /Page 1\/1/);
    assert.match(rendered, /English/);
    assert.match(rendered, /\|\s*cat\s*\|/);
    assert.doesNotMatch(rendered, /\|\s*neko-jp\s*\|/);
    await command("/search cat", /1 matches/);
    await key("\x1b", /Page 1\/1/);
    await command("/test 1", /Japanese: neko-jp/);
    assert.equal(stderr, "");
  } finally {
    if (child.exitCode === null && child.signalCode === null) {
      const closed = once(child, "close"); child.kill(); await closed;
    }
    unlinkSync(path); rmdirSync(directory);
  }
});
