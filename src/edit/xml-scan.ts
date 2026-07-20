export type XmlPiece = {
  start: number;
  end: number;
  raw: string;
  kind: 'tag' | 'text';
  name?: string;
  closing?: boolean;
  selfClosing?: boolean;
};

function tagEnd(source: string, start: number): number {
  let quote: '"' | "'" | null = null;
  for (let index = start + 1; index < source.length; index++) {
    const char = source[index];
    if (quote) {
      if (char === quote) quote = null;
      continue;
    }
    if (char === '"' || char === "'") {
      quote = char;
      continue;
    }
    if (char === '>') return index + 1;
  }
  return source.length;
}

/** Minimal quote-aware XML scanner used only for source-preserving edits. */
export function scanXml(source: string): XmlPiece[] {
  const pieces: XmlPiece[] = [];
  let index = 0;
  while (index < source.length) {
    if (source[index] !== '<') {
      const next = source.indexOf('<', index);
      const end = next < 0 ? source.length : next;
      pieces.push({
        start: index,
        end,
        raw: source.slice(index, end),
        kind: 'text',
      });
      index = end;
      continue;
    }
    if (source.startsWith('<!--', index)) {
      const close = source.indexOf('-->', index + 4);
      const end = close < 0 ? source.length : close + 3;
      pieces.push({
        start: index,
        end,
        raw: source.slice(index, end),
        kind: 'tag',
      });
      index = end;
      continue;
    }
    if (source.startsWith('<![CDATA[', index)) {
      const close = source.indexOf(']]>', index + 9);
      const contentEnd = close < 0 ? source.length : close;
      pieces.push({
        start: index,
        end: index + 9,
        raw: '<![CDATA[',
        kind: 'tag',
      });
      pieces.push({
        start: index + 9,
        end: contentEnd,
        raw: source.slice(index + 9, contentEnd),
        kind: 'text',
      });
      if (close >= 0) {
        pieces.push({
          start: close,
          end: close + 3,
          raw: ']]>',
          kind: 'tag',
        });
      }
      index = close < 0 ? source.length : close + 3;
      continue;
    }

    const end = tagEnd(source, index);
    const raw = source.slice(index, end);
    const parsed = /^<\s*(\/?)\s*([A-Za-z_][\w:.-]*)/.exec(raw);
    pieces.push({
      start: index,
      end,
      raw,
      kind: 'tag',
      ...(parsed
        ? {
            name: parsed[2].toLowerCase(),
            closing: parsed[1] === '/',
            selfClosing: /\/\s*>$/.test(raw),
          }
        : {}),
    });
    index = end;
  }
  return pieces;
}

export type XmlElementRange = {
  name: string;
  start: number;
  end: number;
};

/** Return balanced source ranges for selected element names. */
export function xmlElementRanges(
  source: string,
  acceptedNames: Set<string>
): XmlElementRange[] {
  const stack: Array<{ name: string; start: number }> = [];
  const ranges: XmlElementRange[] = [];
  for (const piece of scanXml(source)) {
    if (piece.kind !== 'tag' || !piece.name) continue;
    if (piece.closing) {
      const open = stack.pop();
      if (!open || open.name !== piece.name) continue;
      const localName = open.name.includes(':')
        ? open.name.slice(open.name.lastIndexOf(':') + 1)
        : open.name;
      if (acceptedNames.has(open.name) || acceptedNames.has(localName)) {
        ranges.push({ name: open.name, start: open.start, end: piece.end });
      }
    } else if (!piece.selfClosing) {
      stack.push({ name: piece.name, start: piece.start });
    }
  }
  return ranges;
}
