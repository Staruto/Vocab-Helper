import assert from "node:assert/strict";
import test from "node:test";
import { buildInputLine, editText, splitGraphemes, textActionFromRawInput, TextEditState } from "./text-input.js";

test("caret movement clamps to the text boundaries", () => {
  assert.deepEqual(editText({ value: "word", cursor: 0 }, { kind: "left" }), { value: "word", cursor: 0 });
  assert.deepEqual(editText({ value: "word", cursor: 2 }, { kind: "left" }), { value: "word", cursor: 1 });
  assert.deepEqual(editText({ value: "word", cursor: 2 }, { kind: "right" }), { value: "word", cursor: 3 });
  assert.deepEqual(editText({ value: "word", cursor: 4 }, { kind: "right" }), { value: "word", cursor: 4 });
});

test("text inserts at the caret and advances past pasted graphemes", () => {
  assert.deepEqual(
    editText({ value: "abcd", cursor: 2 }, { kind: "insert", text: "XY" }),
    { value: "abXYcd", cursor: 4 },
  );
});

test("backspace removes before the caret and delete removes after it", () => {
  const state: TextEditState = { value: "abcd", cursor: 2 };
  assert.deepEqual(editText(state, { kind: "backspace" }), { value: "acd", cursor: 1 });
  assert.deepEqual(editText(state, { kind: "delete" }), { value: "abd", cursor: 2 });
  assert.deepEqual(editText({ value: "abcd", cursor: 0 }, { kind: "backspace" }), { value: "abcd", cursor: 0 });
  assert.deepEqual(editText({ value: "abcd", cursor: 4 }, { kind: "delete" }), { value: "abcd", cursor: 4 });
});

test("editing preserves complete Unicode grapheme clusters", () => {
  const value = "e\u0301語👨‍👩‍👧‍👦";
  assert.deepEqual(splitGraphemes(value), ["e\u0301", "語", "👨‍👩‍👧‍👦"]);
  assert.deepEqual(editText({ value, cursor: 1 }, { kind: "delete" }), { value: "e\u0301👨‍👩‍👧‍👦", cursor: 1 });
  assert.deepEqual(editText({ value, cursor: 3 }, { kind: "backspace" }), { value: "e\u0301語", cursor: 2 });
});

test("input rendering uses a width-stable blinking vertical caret", () => {
  assert.equal(buildInputLine("> ", "abcd", 2, 10, true), "> ab|d    ");
  assert.equal(buildInputLine("> ", "abcd", 2, 10, false), "> abcd    ");
  assert.equal(buildInputLine("> ", "abcdef", 5, 7, true), "> bcde|");
  assert.equal(buildInputLine("> ", "日本", 2, 8, true), "> 日本| ");
});

test("raw Windows Backspace and forward Delete map to distinct edits", () => {
  assert.deepEqual(textActionFromRawInput("\x7f"), { kind: "backspace" });
  assert.deepEqual(textActionFromRawInput("\x1b[3~"), { kind: "delete" });
  assert.deepEqual(textActionFromRawInput("\x1b[D"), { kind: "left" });
  assert.deepEqual(textActionFromRawInput("\x1b[C"), { kind: "right" });
});
