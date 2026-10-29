export type SimilarWord = { word: string; explanation?: string };
export type SimilarWordSpan = { text: string; word: boolean };

// Parse the whole value so malformed or legacy text can be displayed unchanged.
export function parseSimilarWords(value: string): SimilarWord[] | null {
  const text = value.trim();
  const items: SimilarWord[] = [];
  let cursor = 0;
  while (cursor < text.length) {
    if (text[cursor] !== "{") return null;
    const end = text.indexOf("}", cursor + 1);
    if (end < 0) return null;
    const word = text.slice(cursor + 1, end).trim();
    if (!word || word.includes("{")) return null;
    cursor = end + 1;
    const afterWord = cursor;
    while (/\s/.test(text[cursor] ?? "") && cursor < text.length) cursor++;
    let explanation: string | undefined;
    if (text[cursor] === "(") {
      const start = ++cursor;
      let depth = 1;
      while (cursor < text.length && depth > 0) {
        if (text[cursor] === "(") depth++;
        if (text[cursor] === ")") depth--;
        cursor++;
      }
      if (depth !== 0) return null;
      explanation = text.slice(start, cursor - 1).trim() || undefined;
    } else {
      cursor = afterWord;
    }
    items.push(explanation ? { word, explanation } : { word });
    if (cursor === text.length) break;
    const separatorStart = cursor;
    while (/\s/.test(text[cursor] ?? "") && cursor < text.length) cursor++;
    if (text[cursor] === ",") {
      cursor++;
      while (/\s/.test(text[cursor] ?? "") && cursor < text.length) cursor++;
      if (cursor === text.length) return null;
    } else if (cursor === separatorStart) return null;
  }
  return items.length ? items : null;
}

export function similarWordLines(items: SimilarWord[], width: number, measure: (text: string) => number): SimilarWordSpan[][] {
  const lines: SimilarWordSpan[][] = [];
  for (const item of items) {
    let line: SimilarWordSpan[] = [];
    let text = "";
    const spans = [{ text: item.word, word: true }, ...(item.explanation ? [{ text: `  ${item.explanation}`, word: false }] : [])];
    for (const span of spans) {
      for (const char of Array.from(span.text)) {
        if (text && measure(text + char) > Math.max(1, width)) {
          lines.push(line);
          line = [];
          text = "";
        }
        const last = line.at(-1);
        if (last?.word === span.word) last.text += char;
        else line.push({ text: char, word: span.word });
        text += char;
      }
    }
    lines.push(line);
  }
  return lines;
}
