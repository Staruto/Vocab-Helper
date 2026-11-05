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

test("list picker uses the palette, remembers choices, and direct commands preserve search", { timeout: 30_000 }, async () => {
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-list-screen-"));
  const path = join(directory, "vocab.db");
  const repository = new VocabularyRepository(path);
  const workbook = repository.createWorkbook("List keyboard test");
  repository.addEntry(workbook.id, "matchAlpha", "meaning");
  repository.addEntry(workbook.id, "unrelated", "meaning");
  repository.addEntry(workbook.id, "matchBeta", "meaning");
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
  function savedOrder(): string {
    const verify = new VocabularyRepository(path);
    try { return verify.getWorkbook(workbook.id)!.listOrder; } finally { verify.close(); }
  }
  try {
    await frame(/Selected: List keyboard test/);
    await key("\r", /Page 1\/1/);
    let rendered = await command("/list", /> Added: oldest first/);
    assert.match(rendered, /matchAlpha/);
    assert.match(rendered, /Search:/);
    assert.ok(rendered.indexOf("matchAlpha") < rendered.indexOf("> Added: oldest first"));
    await key("\x1b[A", /> Error count: descending/);
    await key("\x1b[B", /> Added: oldest first/);
    await key("\x1b[B", /> Added: newest first/);
    await key("\x1b", /Cancelled/);
    assert.equal(savedOrder(), "add_asc");
    await command("/list", /> Added: oldest first/);
    await key("\x1b[B", /> Added: newest first/);
    rendered = await key("\r", /List order saved: Added: newest first/);
    assert.ok(rendered.indexOf("matchBeta") < rendered.indexOf("matchAlpha"));
    assert.equal(savedOrder(), "add_desc");
    await command("/list", /> Added: newest first/);
    await key("\x1b", /Cancelled/);
    await command("/search match", /2 matches/);
    rendered = await command("/list add asc", /List order saved: Added: oldest first/);
    assert.match(rendered, /2 matches/);
    assert.doesNotMatch(rendered, /unrelated/);
    assert.doesNotMatch(rendered, /Up\/Down select/);
    assert.ok(rendered.indexOf("matchAlpha") < rendered.indexOf("matchBeta"));
    assert.equal(savedOrder(), "add_asc");
    await command("/list error desc", /List order saved: Error count: descending/);
    assert.equal(savedOrder(), "error_desc");
    await command("/list wrong", /Usage: \/list/);
    assert.equal(savedOrder(), "error_desc");
    await key("\x1b", /Page 1\/1/);
    await key("\x1b", /Selected: List keyboard test/);
    await key("\r", /Page 1\/1/);
    await command("/list", /> Error count: descending/);
  } finally {
    if (child.exitCode === null && child.signalCode === null) {
      const closed = once(child, "close"); child.kill(); await closed;
    }
    unlinkSync(path); rmdirSync(directory);
  }
});
