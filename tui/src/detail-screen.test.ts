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

test("detail page keyboard flow preserves carets and drafts, owns navigation, and saves only after confirmation", { timeout: 30_000 }, async () => {
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-detail-screen-"));
  const path = join(directory, "vocab.db");
  const repository = new VocabularyRepository(path);
  const workbook = repository.createConfiguredWorkbook({
    name: "Detail keyboard test", vocabularyKind: "non_language", vocabularyLabel: "Word", vocabularyLanguageCode: null,
    presetAttributes: [], meaningAttributes: [{ position: 1, label: "Meaning", languageCode: null }, { position: 2, label: "Other meaning", languageCode: null }],
    optionalAttributes: Array.from({ length: 3 }, (_, index) => ({ key: `note_${index}`, label: `Note ${index}`, role: "optional", languageCode: null, required: false, visible: false, displayOrder: index + 1 })),
    tagTypes: [
      { name: "Labels", visible: true, tags: [{ name: "one" }, { name: "two" }] },
      { name: "Hidden", visible: false, tags: [{ name: "secret" }] },
      { name: "Empty", visible: false, tags: [] },
    ],
  });
  const types = repository.listTagTypes(workbook.id);
  const entry = repository.addEntry(workbook.id, "猫", "primary", ["primary", "secondary"], { note_0: "long text ".repeat(25) }, [types[0].tags[0].id, types[1].tags[0].id]);
  repository.addEntry(workbook.id, "second", "next meaning");
  repository.completePracticeSession(workbook.id, [{ entryId: entry.id, answers: [false, true] }], "2026-10-30");
  const historyBefore = repository.listDailyPractice(workbook.id);
  repository.close();

  // Exercise the real app and Ink input/rendering using terminal-shaped pipe streams.
  const source = `
    process.stdin.isTTY = true;
    process.stdin.setRawMode = () => {};
    process.stdout.isTTY = true;
    process.stdout.columns = 80;
    process.stdout.rows = 24;
    process.on('message', ({width, height}) => {
      process.stdout.columns = width; process.stdout.rows = height;
      process.stdout.emit('resize');
    });
    await import(${JSON.stringify(new URL("./index.tsx", import.meta.url).href)});
  `;
  const child = spawn(process.execPath, ["--import", "tsx", "--input-type=module", "-e", source], {
    env: { ...process.env, VOCAB_HELPER_DB_PATH: path, FORCE_COLOR: "3", TERM: "xterm-256color", CI: "false", NODE_NO_WARNINGS: "1" },
    stdio: ["pipe", "pipe", "pipe", "ipc"],
  });
  const { stdin, stdout, stderr: errorStream } = child;
  assert.ok(stdin && stdout && errorStream);
  let output = ""; let stderr = "";
  stdout.on("data", (chunk: Buffer) => { output += chunk.toString(); });
  errorStream.on("data", (chunk: Buffer) => { stderr += chunk.toString(); });
  function latestFrame(): string {
    const start = Math.max(output.lastIndexOf("\x1b[G"), output.lastIndexOf("\x1b[H"));
    return output.slice(start < 0 ? 0 : start + 3);
  }
  async function frame(expected: RegExp): Promise<string> {
    const deadline = Date.now() + 5_000;
    while (!expected.test(stripVTControlCharacters(output))) {
      assert.equal(child.exitCode, null, `App exited: ${stderr}`);
      if (Date.now() > deadline) assert.fail(`Missing ${expected}: ${stripVTControlCharacters(output)}\n${stderr}`);
      await delay(20);
    }
    await delay(50); // Let input subscriptions and throttled terminal rendering settle.
    return latestFrame();
  }
  const key = async (input: string, expected: RegExp): Promise<string> => {
    output = ""; stdin.write(input); return frame(expected);
  };
  function attributeLine(rendered: string, label: string): string {
    const line = rendered.split("\n").find((line) => stripVTControlCharacters(line).includes(`${label}:`));
    assert.ok(line, `Missing attribute row ${label}`);
    return line;
  }
  function assertSelectedRow(rendered: string, label: string, editable: boolean): void {
    const line = attributeLine(rendered, label);
    assert.match(stripVTControlCharacters(line), /^> /);
    assert.match(line, /38;2;206;168;255m/);
    assert.doesNotMatch(line, /38;2;229;198;95m/);
    assert.doesNotMatch(line, /\x1b\[40m/);
    // A text row has only the standard white block caret; tag rows have no background.
    assert.equal((line.match(/\x1b\[47m/g) ?? []).length, editable ? 1 : 0);
  }
  const down = "\x1b[B"; const up = "\x1b[A";
  try {
    await frame(/Selected: Detail keyboard test/);
    await key("\r", /Page 1\/1/);
    await key("/edit", /\/edit/);
    await key("\r", /Enter id for \/edit/);
    await key(String(entry.id), /\/edit 1/);
    let rendered = await key("\r", /↑↓ rows/);
    assert.match(stripVTControlCharacters(rendered), /\nVocabulary\s*\n>\s+Word:/);
    const vocabularyHeading = rendered.split("\n").find((line) => stripVTControlCharacters(line).trim() === "Vocabulary")!;
    assert.match(vocabularyHeading, /\x1b\[1m/);
    assert.match(vocabularyHeading, /\x1b\[36m/);
    assert.match(stripVTControlCharacters(rendered), /Word:\s+猫/);
    assertSelectedRow(rendered, "Word", true);
    rendered = await key("X", /Word:\s+猫X/);
    assertSelectedRow(rendered, "Word", true);
    assert.match(stripVTControlCharacters(attributeLine(rendered, "Word")), /^> \*Word:/);
    await key("\x1b[D", /↑↓ rows/);
    rendered = await key(down, /Meaning:\s+primary/);
    assertSelectedRow(rendered, "Meaning", true);
    assert.match(attributeLine(rendered, "Word"), /38;2;229;198;95m/);
    assert.match(stripVTControlCharacters(attributeLine(rendered, "Word")), /^  \*Word:/);
    assert.equal(stripVTControlCharacters(attributeLine(rendered, "Word")).indexOf("猫"), stripVTControlCharacters(attributeLine(rendered, "Meaning")).indexOf("primary"));
    await key("Z", /Meaning:\s+primaryZ/);
    await key(up, /Word:\s+猫X/);
    await key("!", /Word:\s+猫X!/); // Revisiting resets the caret to the end.
    rendered = await key("\x1b", /Unsaved vocabulary changes/);
    assert.match(stripVTControlCharacters(rendered), /> Save Changes \(Preserve Test Data\)/);
    await key(down, /> Save Changes \(Reset Test Data\)/);
    await key(down, /> Discard Changes/);
    await key(down, /> Continue Editing/);
    await key("\r", /Word:\s+猫X!/);
    await key("\x13", /Unsaved vocabulary changes/);
    await key(down, /> Save Changes \(Reset Test Data\)/);
    await key(down, /> Discard Changes/);
    rendered = await key("\r", /Ctrl\+E edit/);
    assert.match(stripVTControlCharacters(rendered), /\nVocabulary\s*\nWord:/);
    await key("\x05", /↑↓ rows/);
    await key("\x1b", /Ctrl\+E edit/); // An unchanged exit skips confirmation.
    await key("\x1b[C", /Word:\s+second/);
    await key("\x1b[D", /Word:\s+猫/);
    await key("\x1b", /Page 1\/1/);
    await key("/view 1", /\/view 1/);
    await key("\r", /Ctrl\+E edit/);
    await key("\x05", /↑↓ rows/);
    await key("!", /Word:\s+猫!/);
    for (let index = 0; index < 7; index++) await key(down, /↑↓ rows/);
    rendered = output;
    assert.match(stripVTControlCharacters(rendered), /Hidden:\s+\[secret\]/);
    assert.match(stripVTControlCharacters(rendered), /Empty:/);
    output = ""; child.send({ width: 34, height: 12 });
    rendered = await frame(/Hidden:/);
    assertSelectedRow(rendered, "Hidden", false);
    assert.ok(stripVTControlCharacters(rendered).split("\n").length <= 13);
    await key("\r", /Select tags/);
    await key(" ", /\[ \] secret/);
    await key("\x1b", /Hidden:\s+\[secret\]/); // Cancelling the picker keeps staged assignments.
    await key("\r", /Select tags/);
    await key(" ", /\[ \] secret/);
    await key("\x13", /Unsaved vocabulary changes/); // Pending picker selection participates in exit.
    rendered = await key("\x1b", /\*Hidden:/);
    assertSelectedRow(rendered, "Hidden", false);
    rendered = await key(down, /Empty:/);
    assertSelectedRow(rendered, "Empty", false);
    assert.match(attributeLine(rendered, "Hidden"), /38;2;229;198;95m/);
    await key("\r", /No tags configured/);
    await key("\r", /Empty:/);
    await key("\x13", /Unsaved vocabulary changes/);
    rendered = await key("\r", /Ctrl\+E edit/);
    assert.match(stripVTControlCharacters(rendered), /Tests:\s+2/);
    assert.match(stripVTControlCharacters(rendered), /Errors:\s+1/);
    assert.match(stripVTControlCharacters(rendered), /Tier:\s+Yellow/);
    assert.match(stripVTControlCharacters(rendered), /Last tested:\s+2026-10-30/);
    assert.match(stripVTControlCharacters(rendered), /Next deadline:\s+2026-11-06/);
    assert.doesNotMatch(stripVTControlCharacters(output), /Hidden:/);
    await key("\x05", /↑↓ rows/);
    await key("?", /Word:\s+猫!\?/);
    await key("\x13", /Unsaved vocabulary changes/);
    await key(down, /> Save Changes \(Reset Test Data\)/);
    rendered = await key("\r", /Ctrl\+E edit/);
    assert.match(stripVTControlCharacters(rendered), /Tests:\s+0/);
    assert.match(stripVTControlCharacters(rendered), /Tier:\s+Gray/);
    assert.equal(stderr, "");
    const verify = new VocabularyRepository(path);
    try {
      const saved = verify.getEntry(entry.id)!;
      assert.equal(saved.vocabulary, "猫!?");
      assert.deepEqual(saved.meanings, ["primary", "secondary"]);
      assert.deepEqual(saved.tags.map((tag) => tag.id), [types[0].tags[0].id]);
      assert.equal(saved.testCount, 0);
      assert.equal(saved.tier, "gray");
      assert.equal(verify.listEntries(workbook.id)[1].vocabulary, "second");
      assert.deepEqual(verify.listDailyPractice(workbook.id), historyBefore);
    } finally { verify.close(); }
  } finally {
    if (child.exitCode === null && child.signalCode === null) {
      const closed = once(child, "close");
      child.kill(); await closed;
    }
    unlinkSync(path); rmdirSync(directory);
  }
});
