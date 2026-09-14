import React, { useEffect, useRef, useState } from "react";
import { Text, useStdin } from "ink";

export type TextEditState = {
  value: string;
  cursor: number;
};

export type TextEditAction =
  | { kind: "left" }
  | { kind: "right" }
  | { kind: "backspace" }
  | { kind: "delete" }
  | { kind: "insert"; text: string };

const graphemeSegmenter = new Intl.Segmenter(undefined, { granularity: "grapheme" });

export function splitGraphemes(value: string): string[] {
  return Array.from(graphemeSegmenter.segment(value), ({ segment }) => segment);
}

export function editText(state: TextEditState, action: TextEditAction): TextEditState {
  const graphemes = splitGraphemes(state.value);
  const cursor = Math.max(0, Math.min(state.cursor, graphemes.length));

  if (action.kind === "left") return { value: state.value, cursor: Math.max(0, cursor - 1) };
  if (action.kind === "right") return { value: state.value, cursor: Math.min(graphemes.length, cursor + 1) };
  if (action.kind === "backspace") {
    if (cursor === 0) return { value: state.value, cursor };
    graphemes.splice(cursor - 1, 1);
    return { value: graphemes.join(""), cursor: cursor - 1 };
  }
  if (action.kind === "delete") {
    if (cursor === graphemes.length) return { value: state.value, cursor };
    graphemes.splice(cursor, 1);
    return { value: graphemes.join(""), cursor };
  }

  const inserted = splitGraphemes(action.text);
  graphemes.splice(cursor, 0, ...inserted);
  return { value: graphemes.join(""), cursor: cursor + inserted.length };
}

function isCombiningMark(codePoint: number): boolean {
  return (
    (codePoint >= 0x0300 && codePoint <= 0x036f) ||
    (codePoint >= 0x1ab0 && codePoint <= 0x1aff) ||
    (codePoint >= 0x1dc0 && codePoint <= 0x1dff) ||
    (codePoint >= 0x20d0 && codePoint <= 0x20ff) ||
    (codePoint >= 0xfe20 && codePoint <= 0xfe2f)
  );
}

function isWide(codePoint: number): boolean {
  return (
    codePoint >= 0x1100 &&
    (
      codePoint <= 0x115f ||
      codePoint === 0x2329 ||
      codePoint === 0x232a ||
      (codePoint >= 0x2e80 && codePoint <= 0xa4cf && codePoint !== 0x303f) ||
      (codePoint >= 0xac00 && codePoint <= 0xd7a3) ||
      (codePoint >= 0xf900 && codePoint <= 0xfaff) ||
      (codePoint >= 0xfe10 && codePoint <= 0xfe19) ||
      (codePoint >= 0xfe30 && codePoint <= 0xfe6f) ||
      (codePoint >= 0xff00 && codePoint <= 0xff60) ||
      (codePoint >= 0xffe0 && codePoint <= 0xffe6) ||
      (codePoint >= 0x1f300 && codePoint <= 0x1faff) ||
      (codePoint >= 0x20000 && codePoint <= 0x3fffd)
    )
  );
}

function graphemeWidth(grapheme: string): number {
  const codePoints = Array.from(grapheme, (char) => char.codePointAt(0) ?? 0);
  if (grapheme.includes("\u200d") || codePoints.some((codePoint) => codePoint >= 0x1f300 && codePoint <= 0x1faff)) return 2;
  return codePoints.reduce((width, codePoint) => width + (isCombiningMark(codePoint) ? 0 : isWide(codePoint) ? 2 : 1), 0);
}

function takePrefixToWidth(value: string, width: number): string {
  let result = "";
  let used = 0;
  for (const grapheme of splitGraphemes(value)) {
    const nextWidth = graphemeWidth(grapheme);
    if (used + nextWidth > width) break;
    result += grapheme;
    used += nextWidth;
  }
  return result;
}

type InputLineSegments = {
  prefix: string;
  before: string;
  cursorCell: string;
  after: string;
  padding: string;
};

export function buildInputSegments(prefix: string, value: string, cursorPosition: number, width: number): InputLineSegments {
  if (width <= 0) return { prefix: "", before: "", cursorCell: "", after: "", padding: "" };

  const fittedPrefix = takePrefixToWidth(prefix, Math.max(0, width - 1));
  const prefixWidth = splitGraphemes(fittedPrefix).reduce((total, grapheme) => total + graphemeWidth(grapheme), 0);
  const viewportWidth = Math.max(1, width - prefixWidth);
  const graphemes = splitGraphemes(value);
  const cursor = Math.max(0, Math.min(cursorPosition, graphemes.length));
  const cursorGrapheme = graphemes[cursor] ?? " ";
  const cursorWidth = Math.max(1, graphemeWidth(cursorGrapheme));
  const textWidth = Math.max(0, viewportWidth - cursorWidth);

  const before: string[] = [];
  let used = 0;
  for (let index = cursor - 1; index >= 0; index--) {
    const nextWidth = graphemeWidth(graphemes[index]);
    if (used + nextWidth > textWidth) break;
    before.unshift(graphemes[index]);
    used += nextWidth;
  }

  const after: string[] = [];
  for (let index = cursor + 1; index < graphemes.length; index++) {
    const nextWidth = graphemeWidth(graphemes[index]);
    if (used + nextWidth > textWidth) break;
    after.push(graphemes[index]);
    used += nextWidth;
  }

  return {
    prefix: fittedPrefix,
    before: before.join(""),
    cursorCell: cursorGrapheme,
    after: after.join(""),
    padding: " ".repeat(Math.max(0, viewportWidth - used - cursorWidth)),
  };
}

export function buildInputLine(prefix: string, value: string, cursorPosition: number, width: number): string {
  const segments = buildInputSegments(prefix, value, cursorPosition, width);
  return `${segments.prefix}${segments.before}${segments.cursorCell}${segments.after}${segments.padding}`;
}

export function textActionFromRawInput(data: Buffer | string): TextEditAction | null {
  const input = data.toString();
  if (input === "\b" || input === "\x7f" || input === "\x1b\b" || input === "\x1b\x7f") return { kind: "backspace" };
  if (/^\x1b(?:\[3[~$^]|\[\[?3~)$/.test(input)) return { kind: "delete" };
  if (/^\x1b(?:\[D|OD|\[1;\d+D)$/.test(input)) return { kind: "left" };
  if (/^\x1b(?:\[C|OC|\[1;\d+C)$/.test(input)) return { kind: "right" };
  if (!input || /[\x00-\x1f\x7f]/.test(input)) return null;
  return { kind: "insert", text: input };
}

type CaretInputLineProps = {
  value: string;
  onChange: (value: string) => void;
  prefix?: string;
  width: number;
  color?: string;
  focus?: boolean;
  inputKey?: string;
};

export function CaretInputLine({ value, onChange, prefix = "", width, color, focus = true, inputKey = "default" }: CaretInputLineProps): JSX.Element {
  const { internal_eventEmitter: inputEvents } = useStdin();
  const [cursor, setCursor] = useState(() => splitGraphemes(value).length);
  const lastEmittedValue = useRef(value);
  const valueRef = useRef(value);
  const cursorRef = useRef(cursor);

  useEffect(() => {
    const nextCursor = splitGraphemes(value).length;
    valueRef.current = value;
    cursorRef.current = nextCursor;
    setCursor(nextCursor);
    lastEmittedValue.current = value;
  }, [inputKey]);

  useEffect(() => {
    valueRef.current = value;
    if (value === lastEmittedValue.current) return;
    const nextCursor = splitGraphemes(value).length;
    cursorRef.current = nextCursor;
    setCursor(nextCursor);
    lastEmittedValue.current = value;
  }, [value]);

  useEffect(() => {
    if (!focus) return;
    const handleInput = (data: Buffer | string): void => {
      const action = textActionFromRawInput(data);
      if (!action) return;
      const next = editText({ value: valueRef.current, cursor: cursorRef.current }, action);
      valueRef.current = next.value;
      cursorRef.current = next.cursor;
      setCursor(next.cursor);
      if (next.value !== lastEmittedValue.current) {
        lastEmittedValue.current = next.value;
        onChange(next.value);
      }
    };
    inputEvents.on("input", handleInput);
    return () => { inputEvents.removeListener("input", handleInput); };
  }, [focus, inputEvents, onChange]);

  const segments = buildInputSegments(prefix, value, cursor, width);
  return <Text color={color}>{segments.prefix}{segments.before}{focus ? <Text color="black" backgroundColor="white">{segments.cursorCell}</Text> : segments.cursorCell}{segments.after}{segments.padding}</Text>;
}
