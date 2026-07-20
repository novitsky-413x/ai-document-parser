import { readFile } from 'node:fs/promises';
import { DocumentError } from '../utils/errors.js';
import { writeFileAtomic } from '../utils/paths.js';
import { decodeAnsiByte, escapeRtf } from '../parse/rtf.js';
import {
  replaceVisibleTokenText,
  type VisibleTextToken,
} from './text-selection.js';

type RtfState = {
  skipDestination: boolean;
  codePage: number;
  ucSkip: number;
  inTable: boolean;
};

const SKIPPED_DESTINATIONS = new Set([
  'fonttbl',
  'colortbl',
  'stylesheet',
  'info',
  'pict',
  'object',
  'header',
  'headerl',
  'headerr',
  'headerf',
  'footer',
  'footerl',
  'footerr',
  'footerf',
  'footnote',
  'annotation',
  'field',
  'fldinst',
  'fldrslt',
  'listtable',
  'listoverridetable',
  'generator',
  'xmlnstbl',
  'datastore',
  'themedata',
]);

function signedCodeUnit(value: number): string {
  const code = value < 0 ? value + 0x10000 : value;
  return String.fromCharCode(code & 0xffff);
}

/** Skip `count` ANSI fallback characters after an RTF `\uN` token. */
function skipUnicodeFallback(source: string, from: number, count: number): number {
  let index = from;
  let remaining = count;
  while (index < source.length && remaining > 0) {
    const char = source[index];
    if (char === '{' || char === '}') {
      index++;
      continue;
    }
    if (char !== '\\') {
      index++;
      if (char !== '\r' && char !== '\n') remaining--;
      continue;
    }
    if (source[index + 1] === "'" && /^[0-9a-f]{2}$/i.test(source.slice(index + 2, index + 4))) {
      index += 4;
      remaining--;
      continue;
    }
    const symbol = source[index + 1];
    if (symbol && !/[A-Za-z]/.test(symbol)) {
      index += 2;
      remaining--;
      continue;
    }
    const control = /^\\[A-Za-z]+-?\d* ?/.exec(source.slice(index));
    if (control) {
      index += control[0].length;
      remaining--;
      continue;
    }
    index++;
  }
  return index;
}

function token(
  start: number,
  end: number,
  decoded: string
): VisibleTextToken {
  return { start, end, decoded };
}

function decodeRawAnsi(text: string, codePage: number): string {
  let out = '';
  for (let index = 0; index < text.length; index++) {
    out += decodeAnsiByte(text.charCodeAt(index) & 0xff, codePage);
  }
  return out;
}

/**
 * Tokenize visible RTF body text while retaining raw source offsets.
 *
 * The tokenizer intentionally excludes fields, headers/footers, pictures,
 * objects, and table rows from the first surgical-edit contract.
 */
export function tokenizeRtfParagraphs(source: string): VisibleTextToken[][] {
  const initialCodePage = Number(
    /\\ansicpg(\d+)\b/i.exec(source)?.[1] ?? 1252
  );
  let state: RtfState = {
    skipDestination: false,
    codePage: initialCodePage,
    ucSkip: 1,
    inTable: false,
  };
  const stack: RtfState[] = [];
  const paragraphs: VisibleTextToken[][] = [];
  let current: VisibleTextToken[] = [];
  const finishParagraph = () => {
    if (current.length) paragraphs.push(current);
    current = [];
  };
  const push = (entry: VisibleTextToken) => {
    if (!state.skipDestination && !state.inTable) current.push(entry);
  };

  let index = 0;
  while (index < source.length) {
    const char = source[index];
    if (char === '{') {
      stack.push({ ...state });
      index++;
      continue;
    }
    if (char === '}') {
      state = stack.pop() ?? state;
      index++;
      continue;
    }
    if (char === '\r' || char === '\n') {
      index++;
      continue;
    }
    if (char !== '\\') {
      const start = index;
      while (
        index < source.length &&
        !['\\', '{', '}', '\r', '\n'].includes(source[index])
      ) {
        index++;
      }
      push(
        token(
          start,
          index,
          decodeRawAnsi(source.slice(start, index), state.codePage)
        )
      );
      continue;
    }

    const start = index;
    const next = source[index + 1];
    if (next === '\\' || next === '{' || next === '}') {
      index += 2;
      push(token(start, index, next));
      continue;
    }
    if (next === '*' ) {
      state.skipDestination = true;
      index += 2;
      continue;
    }
    if (next === "'" && /^[0-9a-f]{2}$/i.test(source.slice(index + 2, index + 4))) {
      index += 4;
      push(
        token(
          start,
          index,
          decodeAnsiByte(parseInt(source.slice(start + 2, start + 4), 16), state.codePage)
        )
      );
      continue;
    }
    if (next && !/[A-Za-z]/.test(next)) {
      index += 2;
      const visible =
        next === '~' ? '\u00a0' : next === '_' ? '\u2011' : next === '-' ? '' : null;
      if (visible !== null) push(token(start, index, visible));
      continue;
    }

    const match = /^\\([A-Za-z]+)(-?\d+)? ?/.exec(source.slice(index));
    if (!match) {
      index++;
      continue;
    }
    const word = match[1].toLowerCase();
    const argument = match[2] === undefined ? undefined : Number(match[2]);
    index += match[0].length;

    if (SKIPPED_DESTINATIONS.has(word)) {
      if (!state.skipDestination) finishParagraph();
      state.skipDestination = true;
      continue;
    }
    if (word === 'ansicpg' && argument !== undefined) {
      state.codePage = argument;
      continue;
    }
    if (word === 'uc' && argument !== undefined) {
      state.ucSkip = Math.max(0, argument);
      continue;
    }
    if (word === 'bin' && argument !== undefined) {
      index = Math.min(source.length, index + Math.max(0, argument));
      continue;
    }
    if (word === 'trowd' || word === 'intbl') {
      if (!state.inTable) finishParagraph();
      state.inTable = true;
      continue;
    }
    if (word === 'row') {
      state.inTable = false;
      finishParagraph();
      continue;
    }
    if (word === 'par') {
      if (!state.skipDestination && !state.inTable) finishParagraph();
      continue;
    }
    if (word === 'u' && argument !== undefined) {
      const end = skipUnicodeFallback(source, index, state.ucSkip);
      push(token(start, end, signedCodeUnit(argument)));
      index = end;
      continue;
    }

    const visible =
      word === 'tab'
        ? '\t'
        : word === 'line'
          ? '\n'
          : word === 'emdash'
            ? '—'
            : word === 'endash'
              ? '–'
              : word === 'bullet'
                ? '•'
                : null;
    if (visible !== null) push(token(start, index, visible));
  }
  finishParagraph();
  return paragraphs;
}

/**
 * Render a changed RTF text leaf while preserving unchanged raw ANSI bytes.
 *
 * Single-byte code-page runs are commonly stored as raw bytes (e.g. 0xE9 for
 * "é" under windows-1252). Re-escaping the whole leaf would rewrite those
 * bytes; instead keep the shared prefix/suffix raw and escape only the delta.
 */
function renderRtfTokenChange(
  source: string,
  token: VisibleTextToken,
  nextText: string
): string {
  const previous = token.decoded;
  if (previous === nextText) return source.slice(token.start, token.end);

  const rawLength = token.end - token.start;
  const oneToOne = previous.length === rawLength;
  if (!oneToOne) return escapeRtf(nextText);

  let prefix = 0;
  const maxPrefix = Math.min(previous.length, nextText.length);
  while (prefix < maxPrefix && previous[prefix] === nextText[prefix]) prefix++;

  let suffix = 0;
  const maxSuffix = Math.min(previous.length - prefix, nextText.length - prefix);
  while (
    suffix < maxSuffix &&
    previous[previous.length - 1 - suffix] === nextText[nextText.length - 1 - suffix]
  ) {
    suffix++;
  }

  return (
    source.slice(token.start, token.start + prefix) +
    escapeRtf(nextText.slice(prefix, nextText.length - suffix)) +
    source.slice(token.end - suffix, token.end)
  );
}

export async function editRtfText(
  filePath: string,
  oldString: string,
  newString: string,
  replaceAll = false
): Promise<{ message: string; changes: number }> {
  if (!oldString) {
    throw new DocumentError('INVALID_ARGUMENT', 'oldString cannot be empty');
  }
  const source = (await readFile(filePath)).toString('latin1');
  if (!/^\s*\{\\rtf/i.test(source)) {
    throw new DocumentError('PARSE_FAILED', 'Invalid RTF: missing {\\rtf header');
  }

  let total = 0;
  const changes: Array<{ token: VisibleTextToken; text: string }> = [];
  for (const paragraph of tokenizeRtfParagraphs(source)) {
    const result = replaceVisibleTokenText(
      paragraph,
      oldString,
      newString,
      true
    );
    total += result.count;
    paragraph.forEach((entry, index) => {
      if (entry.decoded !== result.texts[index]) {
        changes.push({ token: entry, text: result.texts[index] });
      }
    });
  }

  if (total === 0) {
    throw new DocumentError(
      'EDIT_TARGET_NOT_FOUND',
      'Text not found in editable RTF body paragraphs.'
    );
  }
  if (total > 1 && !replaceAll) {
    throw new DocumentError(
      'MULTIPLE_MATCHES',
      `Found ${total} matches. Narrow oldString or set replaceAll: true.`
    );
  }

  let output = source;
  changes
    .sort((a, b) => b.token.start - a.token.start)
    .forEach(({ token: entry, text }) => {
      output =
        output.slice(0, entry.start) +
        renderRtfTokenChange(source, entry, text) +
        output.slice(entry.end);
    });

  await writeFileAtomic(filePath, Buffer.from(output, 'latin1'));
  return {
    message: `Replaced ${total} occurrence(s) in RTF: ${filePath}`,
    changes: total,
  };
}
