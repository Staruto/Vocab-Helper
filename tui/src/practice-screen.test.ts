import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { once } from "node:events";
import { mkdtempSync, rmdirSync, unlinkSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { setTimeout as delay } from "node:timers/promises";
import { stripVTControlCharacters } from "node:util";
import { VocabularyRepository } from "./db.js";

test("Tab opens correct-answer details and Esc or Enter continues initial and retry practice", { timeout: 30_000 }, async () => {
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-practice-screen-"));
  const path = join(directory, "vocab.db");
  const repository = new VocabularyRepository(path);
  const workbook = repository.createWorkbook("Practice keyboard test");
  const entry = repository.addEntry(workbook.id, "word", "definition");
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
  async function frame(expected: RegExp): Promise<void> {
    const deadline = Date.now() + 5_000;
    while (!expected.test(stripVTControlCharacters(output))) {
      assert.equal(child.exitCode, null, `App exited: ${stderr}`);
      if (Date.now() > deadline) assert.fail(`Missing ${expected}: ${stripVTControlCharacters(output)}\n${stderr}`);
      await delay(20);
    }
    await delay(50);
  }
  async function key(input: string, expected: RegExp): Promise<void> {
    output = "";
    child.stdin.write(input);
    await frame(expected);
  }
  try {
    await frame(/Selected: Practice keyboard test/);
    await key("\r", /Page 1\/1/);
    for (const retry of [false, true]) {
      await key("/test 1", /\/test 1/);
      await key("\r", /Question 1\/1/);
      if (retry) {
        await key("wrong", /Answer:.*wrong/);
        await key("\r", /Incorrect!/);
        await key("\r", /Enter \/ Esc returns to practice/);
        await key("\r", /Retry round 1/);
      }
      for (const retryKey of ["\t", "\b", "\x7f"]) {
        await key("mistake", /Answer:.*mistake/);
        await key("\r", /Tab \/ Backspace retry/);
        assert.match(stripVTControlCharacters(output), /Answer:.*mistake/);
        assert.doesNotMatch(stripVTControlCharacters(output), /Review/);
        await key(retryKey, /Tab cycle meanings/);
        assert.doesNotMatch(stripVTControlCharacters(output), /mistake|Incorrect!/);
      }
      await key("word", /Answer:.*word/);
      await key("\r", /Tab view details/);
      await key("\t", /Enter \/ Esc returns to practice/);
      assert.match(stripVTControlCharacters(output), /^\s*word\s*$/m);
      await key(retry ? "\r" : "\x1b", retry ? /Final initial-round score: 0\/1/ : /Final initial-round score: 1\/1/);
      await key("\r", /Page 1\/1/);
    }
    assert.equal(stderr, "");
    const verify = new VocabularyRepository(path);
    try {
      assert.equal(verify.getEntry(entry.id)?.testCount, 3);
      assert.equal(verify.getEntry(entry.id)?.errorCount, 1);
    } finally { verify.close(); }
  } finally {
    if (child.exitCode === null && child.signalCode === null) {
      const closed = once(child, "close");
      child.kill();
      await closed;
    }
    unlinkSync(path);
    rmdirSync(directory);
  }
});
