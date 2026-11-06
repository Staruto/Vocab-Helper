import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { once } from "node:events";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import { setTimeout as delay } from "node:timers/promises";
import { stripVTControlCharacters as plain } from "node:util";
import { VocabularyRepository } from "./db.js";
import { settingsTabWindow } from "./settings-navigation.js";

const up = "\x1b[A", down = "\x1b[B", left = "\x1b[D", right = "\x1b[C", esc = "\x1b";

test("settings row shows all sections whenever their labels fit", () => {
  for (let selected = 0; selected < 5; selected++) {
    assert.deepEqual(settingsTabWindow(65, selected), [0, 1, 2, 3, 4]);
    assert.ok(settingsTabWindow(40, selected).includes(selected));
  }
});
async function app(width: number) {
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-settings-screen-"));
  const path = join(directory, "vocab.db");
  const file = join(directory, "words.txt");
  const r = new VocabularyRepository(path);
  const workbook = r.createConfiguredWorkbook({
    name: "Settings keyboard test", vocabularyKind: "non_language", vocabularyLabel: "Word", vocabularyLanguageCode: null,
    meaningAttributes: [{ position: 1, label: "Meaning", languageCode: null }], presetAttributes: [],
    optionalAttributes: [{ key: "note", role: "optional", label: "Note", languageCode: null, required: false, visible: false, displayOrder: 1 }],
    tagTypes: [{ name: "Topic", visible: true, tags: Array.from({ length: 20 }, (_, index) => ({ name: `Tag ${String(index).padStart(2, "0")}` })) }],
  });
  const entry = r.addEntry(workbook.id, "atom", "small", ["small"], { note: "stored note" }, [r.listTagTypes(workbook.id)[0].tags[0].id]);
  r.close();
  writeFileSync(file, "Word: atom\nMeaning: small\n");
  const child = spawn(process.execPath, ["--import", "tsx", "--input-type=module", "-e", `
    process.stdin.isTTY = true; process.stdin.setRawMode = () => {};
    process.stdout.isTTY = true; process.stdout.columns = ${width}; process.stdout.rows = 24;
    await import(${JSON.stringify(new URL("./index.tsx", import.meta.url).href)});
  `], { env: { ...process.env, VOCAB_HELPER_DB_PATH: path, FORCE_COLOR: "3", TERM: "xterm-256color", CI: "false", NODE_NO_WARNINGS: "1" }, stdio: ["pipe", "pipe", "pipe"] });
  let output = "", stderr = "", rendered = "";
  child.stdout.on("data", (chunk: Buffer) => { output += chunk.toString(); });
  child.stderr.on("data", (chunk: Buffer) => { stderr += chunk.toString(); });
  async function frame(expected?: RegExp): Promise<string> {
    const deadline = Date.now() + 5000;
    while (expected && !expected.test(plain(output))) {
      assert.equal(child.exitCode, null, stderr);
      if (Date.now() > deadline) assert.fail(`Missing ${expected}: ${plain(output)}\nPrevious frame: ${plain(rendered)}\n${stderr.slice(-2500)}`);
      await delay(20);
    }
    await delay(100);
    if (output) {
      const start = Math.max(output.lastIndexOf("\x1b[G"), output.lastIndexOf("\x1b[H"));
      rendered = output.slice(start < 0 ? 0 : start + 3);
    }
    return rendered;
  }
  async function press(input: string, expected?: RegExp): Promise<string> {
    output = ""; child.stdin.write(input); return frame(expected);
  }
  async function open(): Promise<string> {
    await frame(/Selected: Settings keyboard test/);
    await press("\r", /Page 1\/1/);
    await press("/setting", /\/setting/);
    return press("\r", /Settings.*Attributes/);
  }
  function read<T>(fn: (repository: VocabularyRepository) => T): T {
    const repository = new VocabularyRepository(path);
    try { return fn(repository); } finally { repository.close(); }
  }
  return { press, open, read, workbook, entry, file, async close() {
    if (child.exitCode === null && child.signalCode === null) { const closed = once(child, "close"); child.kill(); await closed; }
    rmSync(directory, { recursive: true, force: true });
    assert.equal(stderr, "");
  } };
}

function selectedTab(rendered: string, label: string, focused: boolean): void {
  const line = rendered.split("\n").find((line) => /^Settings\s{3}/.test(plain(line)));
  assert.ok(line, "Missing settings navigation row");
  assert.match(line, new RegExp(`\\x1b\\[7m(?:\\x1b\\[[\\d;]*m)* ${label} (?:\\x1b\\[[\\d;]*m)*\\x1b\\[27m`));
  assert.match(line, new RegExp(`\\x1b\\[${focused ? "38;2;197;168;255" : "37"}m(?:\\x1b\\[[\\d;]*m)* ${label} `));
  assert.equal((line.match(/\x1b\[7m/g) ?? []).length, 1);
}

test("settings tabs retain drafts, validate together and save all sections", { timeout: 60_000 }, async () => {
  const a = await app(80);
  try {
    let rendered = await a.open();
    assert.match(plain(rendered), /Settings - Settings keyboard test[^\n]*\n[ \t]*─+[ \t]*\nSettings\s+Attributes/);
    assert.match(plain(rendered), /Settings\s+Attributes\s+Tags\s+Import\s+Appearance\s+Daily Goal/);
    assert.match(plain(rendered), /Configure field names/);
    assert.doesNotMatch(plain(rendered), /> Word/);
    selectedTab(rendered, "Attributes", true);
    rendered = await a.press(left, /Words per day/);
    selectedTab(rendered, "Daily Goal", true);
    await a.press(right, /Configure field names/);
    rendered = await a.press(down, /> Word/);
    selectedTab(rendered, "Attributes", false);
    // Cursor keys in content must not switch sections.
    selectedTab(await a.press(right), "Attributes", false);
    await a.press(down, /> Meaning/);
    await a.press(" ", /> Meaning.*hidden/);
    await a.press(down, /> Note/);
    rendered = await a.press(down);
    assert.match(plain(rendered), /> Note/); // no vertical wrapping
    await a.press(up, /> Meaning/); await a.press(up, /> Word/);
    rendered = await a.press(up, /sections/);
    selectedTab(rendered, "Attributes", true);
    rendered = await a.press(right, /Organize tags/);
    assert.match(plain(rendered), /┌─+┐/);
    assert.match(plain(rendered), /│\s+Topic/);
    await a.press(down, /> Topic/);
    await a.press(" ", /Topic.*hidden/);
    // Long tag lists scroll within the available screen.
    for (let i = 0; i < 21; i++) {
      rendered = await a.press(down, i < 20 ? new RegExp(`> Tag ${String(i).padStart(2, "0")}`) : />\s+\+ Add tag type/);
    }
    assert.match(plain(rendered), />\s+\+ Add tag type/);
    assert.ok(plain(rendered).trimEnd().split("\n").length < 24);
    rendered = await a.press(down);
    assert.match(plain(rendered), />\s+\+ Add tag type/);
    for (let i = 19; i >= 0; i--) await a.press(up, new RegExp(`> Tag ${String(i).padStart(2, "0")}`));
    await a.press(up, /> Topic/);
    rendered = await a.press(up, /←\/→ sections/);
    selectedTab(rendered, "Tags", true);
    await a.press(left, /Configure field names/);
    assert.match(plain(await a.press(down, /> Word/)), /Meaning.*hidden/);
    await a.press(up, /←\/→ sections/); await a.press(left, /Words per day/);
    await a.press(down); await a.press("0", /Daily goal \(Words per day\): 0/);
    await a.press(esc, /Unsaved settings changes/);
    await a.press("\r", /positive whole number/);
    assert.equal(a.read((r) => r.getWorkbook(a.workbook.id)?.meaningAttributes[0].visible), true);
    await a.press("\x15"); await a.press("23", /Daily goal \(Words per day\): 23/);
    // Text cursor navigation keeps focus and section.
    selectedTab(await a.press(left), "Daily Goal", false);
    await a.press(up, /←\/→ sections/); await a.press(left, /Tier colors/);
    await a.press(down, />.*Tier colors/); await a.press(" ", /> \[ \] Tier colors/); // immediate Appearance preference
    await a.press(up, /←\/→ sections/); await a.press(left, /Default import file/);
    await a.press(down); await a.press("missing-settings-file.txt");
    await a.press(esc, /Unsaved settings changes/);
    await a.press("\r", /Could not|ENOENT|Cannot|Unable/);
    selectedTab(await a.press(left), "Import", false);
    await a.press("\x15"); await a.press(a.file);
    await a.press("\r", /Saved\./); // Enter saves just Import, staying in place.
    await a.press(up, /←\/→ sections/); await a.press(right, /Tier colors/); await a.press(right, /Words per day/);
    assert.match(plain(await a.press(down)), /Daily goal \(Words per day\): 23/);
    await a.press(esc, /Unsaved settings changes/);
    await a.press(down); await a.press(down); await a.press("\r", /Daily goal \(Words per day\): 23/); // continue
    await a.press(esc, /Unsaved settings changes/);
    await a.press("\r", /Page 1\/1/);
    a.read((r) => {
      const saved = r.getWorkbook(a.workbook.id)!;
      assert.equal(saved.dailyGoal, 23);
      assert.equal(saved.importFilePath, a.file);
      assert.equal(saved.meaningAttributes[0].visible, false);
      assert.equal(r.listTagTypes(saved.id)[0].visible, false);
      assert.equal(r.getTierColorsEnabled(), false);
    });
  } finally { await a.close(); }
});

test("narrow settings retain the selected tab and discard drafts without reverting Appearance", { timeout: 30_000 }, async () => {
  const a = await app(40);
  try {
    let rendered = await a.open();
    selectedTab(rendered, "Attributes", true);
    await a.press(left, /Words per day/);
    selectedTab(await a.press(down), "Daily Goal", false);
    await a.press("8", /Daily goal \(Words per day\): 8/);
    await a.press(up, /←\/→ sections/); await a.press(left, /Tier colors/);
    await a.press(down, />.*Tier colors/); await a.press(" ", /> \[ \] Tier colors/);
    rendered = await a.press(down);
    assert.match(plain(rendered), />.*Workbook homepage progress/);
    assert.match(plain(await a.press(down)), />.*Workbook homepage progress/);
    await a.press(esc, /Unsaved settings changes/);
    await a.press(down); await a.press("\r", /Page 1\/1/);
    a.read((r) => { assert.equal(r.getWorkbook(a.workbook.id)?.dailyGoal, null); assert.equal(r.getTierColorsEnabled(), false); });
    await a.press("/setting"); await a.press("\r", /Settings.*Attributes/);
    await a.press(left, /Words per day/); await a.press(down);
    await a.press("8");
    rendered = await a.press("\r", /Saved\./);
    selectedTab(rendered, "Daily Goal", false);
    assert.equal(a.read((r) => r.getWorkbook(a.workbook.id)?.dailyGoal), 8);
    await a.press("\x15"); await a.press("\r", /Saved\./);
    assert.equal(a.read((r) => r.getWorkbook(a.workbook.id)?.dailyGoal), null);
    await a.press(esc, /Page 1\/1/);
  } finally { await a.close(); }
});

test("combined removal confirmation covers fields and tag assignments and can be cancelled", { timeout: 30_000 }, async () => {
  const a = await app(80);
  try {
    await a.open(); await a.press(down); await a.press(down); await a.press(down, /> Note/);
    await a.press("\x1b[3~", /Note will be removed/);
    await a.press(up); await a.press(up); await a.press(right, /Organize tags/); await a.press(down, /> Topic/);
    await a.press("\x1b[3~", /will be removed/);
    await a.press(esc, /Unsaved settings changes/);
    let rendered = await a.press("\r", /Confirm data removal/);
    assert.match(plain(rendered), /Note: 1 populated value/);
    assert.match(plain(rendered), /Topic: 1 assignment/);
    a.read((r) => { assert.equal(r.getEntry(a.entry.id)?.attributes.note, "stored note"); assert.equal(r.getEntry(a.entry.id)?.tags.length, 1); });
    await a.press(esc, /Organize tags/);
    await a.press(esc, /Unsaved settings changes/); await a.press("\r", /Confirm data removal/);
    await a.press("\r", /Type yes to confirm/);
    await a.press("yes"); await a.press("\r", /Page 1\/1/);
    a.read((r) => { assert.equal(r.getEntry(a.entry.id)?.attributes.note, undefined); assert.deepEqual(r.getEntry(a.entry.id)?.tags, []); });
  } finally { await a.close(); }
});
