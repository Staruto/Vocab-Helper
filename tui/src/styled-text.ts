export type StyledTextSpan = { text: string; highlight: boolean };

export function parseExampleSentence(value: string): StyledTextSpan[] | null {
  const pairs = new Map<number, { end: number; nested: boolean }>();
  const stack: Array<{ start: number; nested: boolean }> = [];
  for (let index = 0; index < value.length; index++) {
    if (value[index] === "{") {
      const parent = stack.at(-1);
      if (parent) parent.nested = true;
      stack.push({ start: index, nested: false });
    } else if (value[index] === "}") {
      const opening = stack.pop();
      if (opening) pairs.set(opening.start, { end: index, nested: opening.nested });
    }
  }
  const spans: StyledTextSpan[] = [];
  let found = false;
  let plainStart = 0;
  for (let index = 0; index < value.length; index++) {
    const pair = pairs.get(index);
    if (!pair) continue;
    if (!pair.nested) {
      if (index > plainStart) spans.push({ text: value.slice(plainStart, index), highlight: false });
      if (pair.end > index + 1) spans.push({ text: value.slice(index + 1, pair.end), highlight: true });
      plainStart = pair.end + 1;
      found = true;
    }
    // Completed nested groups stay literal, including their inner pairs.
    index = pair.end;
  }
  if (!found) return null;
  if (plainStart < value.length) spans.push({ text: value.slice(plainStart), highlight: false });
  return spans;
}

export function wrapStyledText(spans: StyledTextSpan[], width: number, measure: (text: string) => number): StyledTextSpan[][] {
  const lines: StyledTextSpan[][] = [];
  let line: StyledTextSpan[] = [];
  let text = "";
  for (const span of spans) {
    for (const char of Array.from(span.text)) {
      if (text && measure(text + char) > Math.max(1, width)) {
        lines.push(line);
        line = [];
        text = "";
      }
      const last = line.at(-1);
      if (last?.highlight === span.highlight) last.text += char;
      else line.push({ text: char, highlight: span.highlight });
      text += char;
    }
  }
  lines.push(line);
  return lines;
}
