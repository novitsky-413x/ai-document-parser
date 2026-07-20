/**
 * One visible-text leaf in a source container.
 *
 * `start` / `end` are adapter-owned source offsets. The shared replacement
 * algorithm only operates on `decoded`.
 */
export type VisibleTextToken = {
  start: number;
  end: number;
  decoded: string;
};

function applyOne(
  texts: string[],
  start: number,
  end: number,
  insert: string
): string[] {
  let pos = 0;
  let inserted = false;
  return texts.map((decoded) => {
    const tokenStart = pos;
    const tokenEnd = pos + decoded.length;
    pos = tokenEnd;
    if (tokenEnd <= start || tokenStart >= end) return decoded;

    const localStart = Math.max(0, start - tokenStart);
    const localEnd = Math.min(decoded.length, end - tokenStart);
    const before = decoded.slice(0, localStart);
    const after = decoded.slice(localEnd);
    const middle = inserted ? '' : insert;
    inserted = true;
    return before + middle + after;
  });
}

/**
 * Replace visible text while retaining the original token partition.
 *
 * The replacement lands in the first overlapping token. This lets format
 * adapters splice only changed text leaves while preserving surrounding markup.
 */
export function replaceVisibleTokenText(
  tokens: VisibleTextToken[],
  oldString: string,
  newString: string,
  replaceAll: boolean
): { texts: string[]; count: number } {
  let texts = tokens.map((token) => token.decoded);
  if (!oldString) return { texts, count: 0 };
  const original = texts.join('');
  const matches: Array<{ start: number; end: number }> = [];
  let searchFrom = 0;
  while (searchFrom <= original.length) {
    const index = original.indexOf(oldString, searchFrom);
    if (index === -1) break;
    matches.push({ start: index, end: index + oldString.length });
    if (!replaceAll) break;
    searchFrom = index + oldString.length;
  }
  for (let index = matches.length - 1; index >= 0; index--) {
    texts = applyOne(
      texts,
      matches[index].start,
      matches[index].end,
      newString
    );
  }

  return { texts, count: matches.length };
}
