import { DocumentError, wrapError } from '../utils/errors.js';
import type {
  CanonicalFlowBlock,
  CanonicalInline,
  CanonicalTableCell,
  ParsedCanonical,
} from '../ir/types.js';
import { buildParsedCanonical } from '../ir/projections.js';
import { codeStyleId } from '../utils/code-style.js';
import { coalesceAdjacentBlockquotes } from '../utils/quote-style.js';

const TABLE_WIDTH_TWIPS = 9000;
const LIST_INDENT = 720;

/** Decode a single ANSI byte using a Windows code page when available. */
export function decodeAnsiByte(byte: number, codePage: number): string {
  try {
    const label =
      codePage === 1251
        ? 'windows-1251'
        : codePage === 1252
          ? 'windows-1252'
          : codePage === 1250
            ? 'windows-1250'
            : codePage === 866
              ? 'ibm866'
              : null;
    if (label) return new TextDecoder(label).decode(Uint8Array.of(byte));
  } catch {
    // fall through
  }
  return String.fromCharCode(byte);
}

function decodeRtfEscapes(text: string, codePage = 1252): string {
  const BS = '\u0000'; // placeholder so literal backslashes survive control-word stripping
  return (
    text
      .replace(/\\'([0-9a-fA-F]{2})/g, (_, hex) =>
        decodeAnsiByte(parseInt(hex, 16), codePage)
      )
      // \uN is a signed 16-bit code unit; with \uc1 the next char is an ANSI fallback (often '?').
      .replace(/\\u(-?\d+)(?:\\'[0-9a-fA-F]{2}|.)?/g, (_, n) => {
        let code = Number(n);
        if (code < 0) code += 0x10000;
        return String.fromCharCode(code & 0xffff);
      })
      .replace(/\\par\b/g, '\n')
      .replace(/\\tab\b/g, '\t')
      .replace(/\\line\b/g, '\n')
      .replace(/\\emdash\b/g, '—')
      .replace(/\\endash\b/g, '–')
      .replace(/\\bullet\b/g, '•')
      .replace(/\\~ /g, ' ')
      .replace(/\\-/g, '')
      // Literal escapes must be protected before stripping \\control words
      // (otherwise \\Users is eaten as a fake control word "\\Users").
      .replace(/\\\\/g, BS)
      .replace(/\\\{/g, '{')
      .replace(/\\\}/g, '}')
      // Control words may carry a signed numeric arg: \fi-360, \li720, \fs18, ...
      .replace(/\\[a-z]+-?\d* ?/gi, '')
      // Empty groups used as control-word delimiters (e.g. \line{}).
      .replace(/\{\}/g, '')
      .replace(new RegExp(BS, 'g'), '\\')
  );
}

function extractAdpCodeLang(segment: string): string | undefined {
  return /\{\\\*\\adpcodelang\s+([A-Za-z0-9_+-]+)\}/.exec(segment)?.[1];
}

function extractAdpListMeta(
  segment: string
): { ordered: boolean; level: number } | null {
  const m = /\{\\\*\\adplist\s+([ou])\s+(\d+)\}/.exec(segment);
  if (m) return { ordered: m[1] === 'o', level: Number(m[2]) };

  // Third-party Word/LibreOffice list hints.
  if (/\\(?:pntext|listtext)\b/i.test(segment) || /\\ls\d+\b/.test(segment)) {
    const li = /\\li(\d+)\b/.exec(segment);
    const level = li ? Math.max(0, Math.round(Number(li[1]) / LIST_INDENT) - 1) : 0;
    const ordered = /\\pnlvlbody\b|\\pndec\b|\\listtemplateid\b/i.test(segment)
      ? true
      : /\d+\.\\tab|\d+\.\s/.test(segment);
    return { ordered: !!ordered, level: Math.max(0, level) };
  }

  // Indented bullet / numbered paragraph without explicit list controls.
  if (/\\li(\d+)\b/.test(segment)) {
    const li = Number(/\\li(\d+)\b/.exec(segment)?.[1] ?? 0);
    const level = Math.max(0, Math.round(li / LIST_INDENT) - 1);
    if (/\\bullet\b|\\u8226\?|\\u\'95/.test(segment)) {
      return { ordered: false, level };
    }
    if (/(?:^|\{)\\pard[^\\]*?(?:\\[a-z]+-?\d* ?)*\d+\./i.test(segment)) {
      return { ordered: true, level };
    }
  }
  return null;
}

function isRtfCodeParagraph(segment: string): boolean {
  if (extractAdpCodeLang(segment)) return true;
  // Our emitter starts fenced code with \pard\f1; inline code uses mid-run \f1 after \pard\f0.
  return /\\pard\\f1\b/.test(segment);
}

function isRtfQuoteParagraph(segment: string): boolean {
  return /\{\\\*\\adpquote\}/.test(segment);
}

/** End index (exclusive) of a `{ ... }` group starting at `start`. */
function findBalancedBraceGroup(text: string, start: number): number {
  if (text[start] !== '{') return -1;
  let depth = 0;
  for (let i = start; i < text.length; i++) {
    const ch = text[i];
    if (ch === '\\' && i + 1 < text.length) {
      // Skip escaped brace / control backslash so \{ doesn't affect depth.
      i++;
      continue;
    }
    if (ch === '{') depth++;
    else if (ch === '}') {
      depth--;
      if (depth === 0) return i + 1;
    }
  }
  return -1;
}

function parseHyperlinkField(
  fieldXml: string,
  codePage: number
): CanonicalInline | null {
  const href =
    /HYPERLINK\s+"([^"]+)"/i.exec(fieldXml)?.[1] ??
    /HYPERLINK\s+'([^']+)'/i.exec(fieldXml)?.[1] ??
    /HYPERLINK\s+([^\s\\{}]+)/i.exec(fieldXml)?.[1];
  if (!href) return null;

  const resultAt = fieldXml.search(/\{\\fldrslt\b/);
  let label = href;
  if (resultAt >= 0) {
    const end = findBalancedBraceGroup(fieldXml, resultAt);
    if (end > resultAt) {
      const inner = fieldXml.slice(resultAt, end).replace(/^\{\\fldrslt\b\s?/, '').replace(/\}$/, '');
      const decoded = decodeRtfEscapes(inner, codePage).trim();
      if (decoded) label = decoded;
    }
  }
  return { text: label, href };
}

function parseRtfInline(segment: string, codePage = 1252): CanonicalInline[] {
  const runs: CanonicalInline[] = [];
  let i = 0;
  while (i < segment.length) {
    const fieldAt = segment.indexOf('{\\field', i);
    if (fieldAt < 0) {
      runs.push(...parseRtfInlinePlain(segment.slice(i), codePage));
      break;
    }
    if (fieldAt > i) {
      runs.push(...parseRtfInlinePlain(segment.slice(i, fieldAt), codePage));
    }
    const end = findBalancedBraceGroup(segment, fieldAt);
    if (end < 0) {
      runs.push(...parseRtfInlinePlain(segment.slice(fieldAt), codePage));
      break;
    }
    const field = segment.slice(fieldAt, end);
    const link = /HYPERLINK/i.test(field) ? parseHyperlinkField(field, codePage) : null;
    if (link) runs.push(link);
    else runs.push(...parseRtfInlinePlain(field, codePage));
    i = end;
  }
  return runs.filter((r) => r.text);
}

function parseRtfInlinePlain(segment: string, codePage = 1252): CanonicalInline[] {
  const runs: CanonicalInline[] = [];
  let bold = false;
  let italic = false;
  let strike = false;
  let code = false;
  // Consume the optional trailing space that terminates an RTF control word —
  // that space is a delimiter, not document text (must match surgical tokenizer).
  const parts = segment.split(
    /(\\(?:b0|i0|strike0|b|i|strike|f1|f0)(?![A-Za-z]) ?)/
  );
  let buf = '';
  const flush = () => {
    if (!buf) return;
    const text = decodeRtfEscapes(buf, codePage);
    buf = '';
    if (text) runs.push({ text, bold, italic, strike, ...(code ? { code: true } : {}) });
  };
  for (const part of parts) {
    const ctrl = /^\\(b0|i0|strike0|b|i|strike|f1|f0) ?$/.exec(part);
    if (ctrl) {
      flush();
      const word = ctrl[1];
      if (word === 'b') bold = true;
      else if (word === 'b0') bold = false;
      else if (word === 'i') italic = true;
      else if (word === 'i0') italic = false;
      else if (word === 'strike') strike = true;
      else if (word === 'strike0') strike = false;
      else if (word === 'f1') code = true;
      else if (word === 'f0') code = false;
      continue;
    }
    buf += part;
  }
  flush();
  return runs.filter((r) => r.text);
}

function stripAdpMarkers(segment: string): string {
  return segment
    .replace(/\{\\\*\\adpcodelang\s+[A-Za-z0-9_+-]+\}/g, '')
    .replace(/\{\\\*\\adplist\s+[ou]\s+\d+\}/g, '')
    .replace(/\{\\\*\\adpquote\}/g, '');
}

type FlatListItem = { level: number; ordered?: boolean; runs: CanonicalInline[] };

function buildNestedList(
  items: FlatListItem[],
  depth: number
): Extract<CanonicalFlowBlock, { type: 'list' }> {
  const list: Extract<CanonicalFlowBlock, { type: 'list' }> = {
    type: 'list',
    ordered: items[0]?.ordered,
    items: [],
  };
  let i = 0;
  while (i < items.length) {
    if (items[i].level <= depth) {
      list.items.push([{ type: 'paragraph', runs: items[i].runs }]);
      i++;
    } else {
      const sub: FlatListItem[] = [];
      while (i < items.length && items[i].level > depth) {
        sub.push(items[i]);
        i++;
      }
      const nested = buildNestedList(sub, depth + 1);
      if (list.items.length) list.items[list.items.length - 1].push(nested);
      else list.items.push([nested]);
    }
  }
  return list;
}

function parseTableRegion(region: string, codePage = 1252): CanonicalFlowBlock | null {
  const rows: CanonicalTableCell[][] = [];
  const rowRe = /\\trowd[\s\S]*?\\row/g;
  let rm: RegExpExecArray | null;
  while ((rm = rowRe.exec(region)) !== null) {
    const raw = rm[0].replace(/\\row\s*$/, '');
    const parts = raw.split(/\\cell\b/);
    const cells: CanonicalTableCell[] = [];
    for (let i = 0; i < parts.length; i++) {
      const part = parts[i];
      if (
        i === parts.length - 1 &&
        !/\\intbl\b/.test(part) &&
        !decodeRtfEscapes(part, codePage).trim()
      ) {
        continue;
      }
      let cell = part.replace(/\\trowd\b[\s\S]*?(?=\\intbl|$)/, '');
      cell = cell.replace(/\\intbl\b/g, '');
      cells.push({ text: decodeRtfEscapes(cell, codePage).replace(/^\s+|\s+$/g, '') });
    }
    if (cells.length) rows.push(cells);
  }
  return rows.length ? { type: 'table', headerRows: 1, rows } : null;
}

function parseParagraphRegion(region: string, codePage = 1252): CanonicalFlowBlock[] {
  const paragraphs = region.split(/\\par\b/).map((p) => p.trim()).filter(Boolean);
  const out: CanonicalFlowBlock[] = [];
  let listBuf: FlatListItem[] = [];
  let listOrdered: boolean | undefined;

  const flushList = () => {
    if (listBuf.length) {
      out.push(buildNestedList(listBuf, 0));
      listBuf = [];
      listOrdered = undefined;
    }
  };

  for (const segment of paragraphs) {
    if (isRtfCodeParagraph(segment)) {
      flushList();
      const language = extractAdpCodeLang(segment);
      out.push({
        type: 'code',
        text: decodeRtfEscapes(stripAdpMarkers(segment), codePage),
        ...(language ? { language } : {}),
      });
      continue;
    }

    if (isRtfQuoteParagraph(segment)) {
      flushList();
      const runs = parseRtfInline(stripAdpMarkers(segment), codePage);
      if (runs.length) {
        out.push({ type: 'blockquote', blocks: [{ type: 'paragraph', runs }] });
      }
      continue;
    }

    const listMeta = extractAdpListMeta(segment);
    if (listMeta) {
      if (listBuf.length && listMeta.level === 0 && listOrdered !== listMeta.ordered) {
        flushList();
      }
      if (!listBuf.length) listOrdered = listMeta.ordered;
      const runs = parseRtfInline(stripAdpMarkers(segment), codePage);
      const cleaned = stripListPrefix(runs);
      if (cleaned.length) {
        listBuf.push({
          level: listMeta.level,
          ordered: listMeta.ordered,
          runs: cleaned,
        });
      }
      continue;
    }

    flushList();
    const runs = parseRtfInline(segment, codePage);
    if (!runs.length) continue;
    const outline = /\\outlinelevel(\d+)\b/i.exec(segment);
    if (outline) {
      const level = Math.min(6, Math.max(1, Number(outline[1]) + 1)) as 1 | 2 | 3 | 4 | 5 | 6;
      out.push({ type: 'heading', level, runs: runs.map((r) => ({ text: r.text })) });
    } else if (/\\pard\\f0\\b\b/.test(segment) || /\\pard\\f0\\b /.test(segment)) {
      out.push({
        type: 'heading',
        level: 1,
        runs: runs.map((r) => ({ text: r.text })),
      });
    } else {
      out.push({ type: 'paragraph', runs });
    }
  }
  flushList();
  return coalesceAdjacentBlockquotes(out);
}

function stripListPrefix(runs: CanonicalInline[]): CanonicalInline[] {
  if (!runs.length) return runs;
  const first = { ...runs[0] };
  first.text = first.text
    .replace(/^•\s*/, '')
    .replace(/^\d+\.\s+/, '')
    .replace(/^[-*+]\s+/, '');
  if (!first.text && runs.length === 1) return [];
  if (!first.text) return runs.slice(1);
  return [first, ...runs.slice(1)];
}

function parseRtfBody(body: string, codePage = 1252): CanonicalFlowBlock[] {
  const blocks: CanonicalFlowBlock[] = [];
  const tableRe = /(?:\\trowd[\s\S]*?\\row\s*)+/g;
  let last = 0;
  let m: RegExpExecArray | null;
  while ((m = tableRe.exec(body)) !== null) {
    if (m.index > last) {
      blocks.push(...parseParagraphRegion(body.slice(last, m.index), codePage));
    }
    const table = parseTableRegion(m[0], codePage);
    if (table) blocks.push(table);
    last = m.index + m[0].length;
  }
  if (last < body.length) {
    blocks.push(...parseParagraphRegion(body.slice(last), codePage));
  }
  return blocks;
}

/** Word metadata groups that leak control text. Field instructions and adp markers stay. */
const IGNORABLE_STAR_PREFIXES = [
  '{\\*\\generator',
  '{\\*\\datastore',
  '{\\*\\latentstyles',
  '{\\*\\colorschememapping',
  '{\\*\\themedata',
  '{\\*\\xmlnstbl',
  '{\\*\\rsidtbl',
  '{\\*\\wgrffmtfilter',
  '{\\*\\pgptbl',
  '{\\*\\listoverridetable',
  '{\\*\\revtbl',
  '{\\*\\stylesheet',
];

/**
 * Drop noisy `{\\*...}` destinations (generator, datastore, stylesheet extras).
 * Keep `{\*\fldinst}` hyperlink instructions and this library's `{\*\adp...}` markers.
 */
function removeIgnorableDestinations(text: string): string {
  let out = '';
  for (let i = 0; i < text.length; i++) {
    const noisy = IGNORABLE_STAR_PREFIXES.some((prefix) => text.startsWith(prefix, i));
    if (noisy) {
      const end = findBalancedBraceGroup(text, i);
      if (end > i) {
        i = end - 1;
        continue;
      }
    }
    out += text[i];
  }
  return out;
}

export function parseRtf(text: string): ParsedCanonical {
  try {
    // RTF source line breaks are insignificant (spec); only \par / \\n break lines.
    // {\*...} destinations are ignorable and otherwise leak control digits into text.
    const normalized = removeIgnorableDestinations(
      text
        .replace(/\r\n/g, '\n')
        // A backslash before a source newline is a line continuation, not a paragraph.
        .replace(/\\\n/g, '')
        .replace(/\n/g, '')
    );
    if (!/^\{\\rtf/i.test(normalized.trimStart())) {
      throw new DocumentError('PARSE_FAILED', 'Invalid RTF: missing {\\rtf header');
    }
    const codePage = Number(/\\ansicpg(\d+)\b/i.exec(normalized)?.[1] ?? 1252);
    let body = normalized.replace(/^\{\\rtf[\s\S]*?\\viewkind\d+/i, '');
    body = removeRtfGroup(body, 'fonttbl');
    body = removeRtfGroup(body, 'colortbl');
    body = body.replace(/^\{\\rtf1?\\ansi\\uc1/, '').replace(/^\{\\rtf1?\\ansi/, '');
    // Drop only the document's trailing closer — never a leading `{` that opens {\*\...}.
    body = body.replace(/\s*\}\s*$/, '');
    const blocks = parseRtfBody(body, codePage);
    return buildParsedCanonical(
      { kind: 'flow', format: 'rtf', blocks, raw: text },
      [{ code: 'RTF_BASIC_PARSE', message: 'RTF parsed with basic control-word extraction.' }]
    );
  } catch (err) {
    if (err instanceof DocumentError) throw err;
    throw wrapError('PARSE_FAILED', 'Failed to parse RTF', err);
  }
}

/** Remove a top-level `{\groupName ...}` destination (balanced braces). */
function removeRtfGroup(text: string, groupName: string): string {
  const re = new RegExp(`\\{\\\\${groupName}\\b`);
  const start = text.search(re);
  if (start < 0) return text;
  let depth = 0;
  for (let i = start; i < text.length; i++) {
    const ch = text[i];
    if (ch === '{') depth++;
    else if (ch === '}') {
      depth--;
      if (depth === 0) return text.slice(0, start) + text.slice(i + 1);
    }
  }
  return text;
}

export function emitRtfFromFlow(blocks: CanonicalFlowBlock[]): string {
  const parts = [
    '{\\rtf1\\ansi\\uc1',
    '{\\fonttbl{\\f0\\froman Times New Roman;}{\\f1\\fmodern Courier New;}}',
  ];
  for (const b of blocks) emitBlock(b, parts);
  parts.push('}');
  return parts.join('');
}

function emitBlock(b: CanonicalFlowBlock, parts: string[]): void {
  if (b.type === 'heading') {
    const level = Math.min(6, Math.max(1, b.level));
    parts.push(`\\pard\\f0\\outlinelevel${level - 1}\\b `);
    parts.push(escapeRtf(runsToPlain(b.runs)));
    parts.push('\\b0\\par ');
  } else if (b.type === 'paragraph') {
    parts.push('\\pard\\f0 ');
    emitRuns(b.runs, parts);
    parts.push('\\par ');
  } else if (b.type === 'code') {
    const style = codeStyleId(b.language);
    const lang = style.startsWith('ADPCode_') ? style.slice('ADPCode_'.length) : '';
    if (lang) parts.push(`{\\*\\adpcodelang ${lang}}`);
    parts.push('\\pard\\f1\\fs18 ');
    parts.push(escapeRtf(b.text));
    parts.push('\\pard\\f0\\par ');
  } else if (b.type === 'thematic_break') {
    parts.push('\\pard\\f0 ');
    parts.push(escapeRtf('---'));
    parts.push('\\par ');
  } else if (b.type === 'blockquote') {
    for (const inner of b.blocks) {
      if (inner.type === 'paragraph' || inner.type === 'heading') {
        parts.push('{\\*\\adpquote}');
        parts.push('\\pard\\f0\\li720\\i ');
        emitRuns(inner.runs, parts);
        parts.push('\\i0\\par ');
      } else {
        emitBlock(inner, parts);
      }
    }
  } else if (b.type === 'list') {
    emitList(b, 0, parts);
  } else if (b.type === 'table') {
    emitTable(b, parts);
  }
}

function emitRuns(runs: CanonicalInline[], parts: string[]): void {
  for (const r of runs) {
    if (r.href) {
      const href = r.href.replace(/\\/g, '\\\\').replace(/"/g, '');
      parts.push(`{\\field{\\*\\fldinst HYPERLINK "${href}"}{\\fldrslt `);
    }
    if (r.code) parts.push('\\f1\\fs18 ');
    if (r.bold) parts.push('\\b ');
    if (r.italic) parts.push('\\i ');
    if (r.strike) parts.push('\\strike ');
    parts.push(escapeRtf(r.text));
    if (r.strike) parts.push('\\strike0 ');
    if (r.italic) parts.push('\\i0 ');
    if (r.bold) parts.push('\\b0 ');
    if (r.code) parts.push('\\f0 ');
    if (r.href) parts.push('}}');
  }
}

function emitList(
  list: Extract<CanonicalFlowBlock, { type: 'list' }>,
  depth: number,
  parts: string[]
): void {
  let n = 1;
  const indent = LIST_INDENT * (depth + 1);
  const flag = list.ordered ? 'o' : 'u';
  for (const item of list.items) {
    for (const sub of item) {
      if (sub.type === 'paragraph' || sub.type === 'heading') {
        parts.push(`{\\*\\adplist ${flag} ${depth}}`);
        parts.push(`\\pard\\f0\\fi-360\\li${indent} `);
        if (list.ordered) {
          parts.push(escapeRtf(`${n}. `));
          n++;
        } else {
          parts.push('\\bullet ');
        }
        emitRuns(sub.runs, parts);
        parts.push('\\par ');
      } else if (sub.type === 'list') {
        emitList(sub, depth + 1, parts);
      }
    }
  }
}

function emitTable(
  table: Extract<CanonicalFlowBlock, { type: 'table' }>,
  parts: string[]
): void {
  const colCount = Math.max(1, ...table.rows.map((r) => r.length));
  const colWidth = Math.max(1, Math.floor(TABLE_WIDTH_TWIPS / colCount));
  for (const row of table.rows) {
    parts.push('\\trowd\\trgaph108');
    for (let c = 0; c < colCount; c++) {
      parts.push(`\\cellx${(c + 1) * colWidth}`);
    }
    parts.push(' ');
    for (let c = 0; c < colCount; c++) {
      const cell = row[c] ?? { text: '' };
      parts.push('\\intbl ');
      if (cell.runs?.length) emitRuns(cell.runs, parts);
      else parts.push(escapeRtf(cell.text ?? ''));
      parts.push('\\cell ');
    }
    parts.push('\\row ');
  }
  parts.push('\\pard\\f0 ');
}

function runsToPlain(runs: CanonicalInline[]): string {
  return runs.map((r) => r.text).join('');
}

export function escapeRtf(s: string): string {
  const out: string[] = [];
  const normalized = s.replace(/\r\n/g, '\n');
  for (const ch of normalized) {
    if (ch === '\\') {
      out.push('\\\\');
      continue;
    }
    if (ch === '{') {
      out.push('\\{');
      continue;
    }
    if (ch === '}') {
      out.push('\\}');
      continue;
    }
    if (ch === '\n') {
      out.push('\\line{}');
      continue;
    }
    if (ch === '\t') {
      out.push('\\tab ');
      continue;
    }
    const cp = ch.codePointAt(0)!;
    if (cp >= 0x20 && cp <= 0x7e) {
      out.push(ch);
      continue;
    }
    if (cp <= 0xffff) {
      out.push(`\\u${toSigned16(cp)}?`);
      continue;
    }
    const pair = toSurrogatePair(cp);
    out.push(`\\u${toSigned16(pair.high)}?\\u${toSigned16(pair.low)}?`);
  }
  return out.join('');
}

function toSigned16(n: number): number {
  return n > 0x7fff ? n - 0x10000 : n;
}

function toSurrogatePair(cp: number): { high: number; low: number } {
  const v = cp - 0x10000;
  return {
    high: 0xd800 + (v >> 10),
    low: 0xdc00 + (v & 0x3ff),
  };
}
