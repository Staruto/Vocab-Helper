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

test("inverse row selection opens details and restores filtered browsing and input drafts", { timeout: 30_000 }, async () => {
  const directory = mkdtempSync(join(tmpdir(), "vocabhelper-selection-screen-"));
  const path = join(directory, "vocab.db");
  const repository = new VocabularyRepository(path);
  const workbook = repository.createConfiguredWorkbook({
    name: "Selection keyboard test", vocabularyKind: "non_language", vocabularyLabel: "Word", vocabularyLanguageCode: null,
    presetAttributes: [], meaningAttributes: [{ position: 1, label: "Meaning", languageCode: null }], optionalAttributes: [],
    tagTypes: [{ name: "Labels", visible: true, tags: [{ name: "noun" }] }],
  });
  const tagId = repository.listTagTypes(workbook.id)[0].tags[0].id;
  for (let index = 1; index <= 25; index++) repository.addEntry(workbook.id, `match${index}`, `meaning${index}`, undefined, undefined, [tagId]);
  repository.addEntry(workbook.id, "unrelated", "other meaning");
  repository.setTierColorsEnabled(true);
  repository.close();
  const source = `
    process.stdin.isTTY = true;
    process.stdin.setRawMode = () => {};
    process.stdout.isTTY = true;
    process.stdout.columns = 120;
    process.stdout.rows = 50;
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
  async function key(input: string, expected = /Page \d\/\d/): Promise<string> {
    output = ""; child.stdin.write(input); return frame(expected);
  }
  async function command(raw: string, expected: RegExp): Promise<string> {
    await key(raw); return key("\r", expected);
  }
  function assertSelection(rendered: string, id: number | null): void {
    const rows = rendered.split("\n").filter((line) => /^\|#\d+ /.test(stripVTControlCharacters(line)));
    const selected = rows.filter((line) => /\x1b\[7m/.test(line));
    assert.equal(selected.length, id === null ? 0 : 1, `Unexpected selection: ${stripVTControlCharacters(rendered)}`);
    if (id === null) return;
    assert.match(stripVTControlCharacters(selected[0]), new RegExp(`^\\|#${id} match${id}\\s`));
    assert.match(stripVTControlCharacters(selected[0]), /\[noun\].*\|.*\|$/);
    assert.doesNotMatch(selected[0], /38;2;(119;119;119|95;215;215)m/);
    const unselected = rows.find((line) => !/\x1b\[7m/.test(line));
    assert.ok(unselected);
    assert.match(unselected, /38;2;119;119;119m/);
    assert.match(unselected, /38;2;95;215;215m/);
  }
  const up = "\x1b[A"; const down = "\x1b[B"; const left = "\x1b[D"; const right = "\x1b[C";
  try {
    await frame(/Selected: Selection keyboard test/);
    assertSelection(await key("\r"), null);
    await command("/search match", /25 matches/);
    await key("draft");
    assertSelection(await key(up), null);
    await key("!"); // Leave an unapplied search draft while retaining the active filter.
    await key(left); // A nonempty input owns caret movement, rather than paging.
    let rendered = await key(up);
    assertSelection(rendered, 20);
    assert.match(stripVTControlCharacters(rendered), /Search: match!/);
    assert.match(stripVTControlCharacters(rendered), /> draft/);
    child.stdin.write("ignored"); await delay(80); // Row focus must not edit either draft.
    rendered = await key(right, /Page 2\/2/);
    assertSelection(rendered, 25);
    assert.doesNotMatch(stripVTControlCharacters(rendered), /unrelated/);
    assertSelection(await key(up), 24);
    rendered = await key("\r", /Ctrl\+E edit/);
    assert.match(stripVTControlCharacters(rendered), /Word:\s+match24\s/);
    await key(right, /Word:\s+match25\s/);
    await key("\x05", /↑↓ rows/);
    await key("\x1b", /Ctrl\+E edit/);
    rendered = await key("\x1b", /Page 2\/2/);
    assertSelection(rendered, 24); // Detail navigation preserves the original list selection.
    assert.match(stripVTControlCharacters(rendered), /25 matches/);
    assert.match(stripVTControlCharacters(rendered), /Search: match!/);
    assert.match(stripVTControlCharacters(rendered), /> draft/);
    assert.doesNotMatch(stripVTControlCharacters(rendered), /ignored|unrelated/);
    assertSelection(await key(down), 25);
    rendered = await key(down);
    assertSelection(rendered, null);
    const searchLine = rendered.split("\n").find((line) => stripVTControlCharacters(line).includes("Search:"))!;
    assert.match(searchLine, /\x1b\[47m/);
    await key(down);
    rendered = await key("!");
    assert.match(stripVTControlCharacters(rendered), /> draft!/);
    for (let index = 0; index < 6; index++) await key("\x7f");
    await command("/view 1", /Ctrl\+E edit/);
    rendered = await key("\x1b", /Page 2\/2/);
    assertSelection(rendered, null);
    assert.match(stripVTControlCharacters(rendered), /Search: match!/);
    await command("/edit 1", /↑↓ rows/);
    await key("\x1b", /Ctrl\+E edit/);
    await key("\x1b", /Page 2\/2/);
    await key("/li");
    await key("\t", /\/list /);
    await key("add asc");
    await key("\r", /List order saved/);
    await key(up);
    assertSelection(await key(up), 20);
    for (let index = 0; index < 19; index++) rendered = await key(up);
    assertSelection(rendered, 1);
    await key("\x1b"); // Clear search, then leave and reopen to verify a fresh command focus.
    await key("\x1b", /Selected: Selection keyboard test/);
    assertSelection(await key("\r"), null);
  } finally {
    if (child.exitCode === null && child.signalCode === null) {
      const closed = once(child, "close"); child.kill(); await closed;
    }
    unlinkSync(path); rmdirSync(directory);
  }
});
