import type { ContentBlock, DocumentWarning, InlineRun } from '../types.js';

/**
 * Pages with fewer letters than this are scans: OCR may replace the layer.
 * Short real text (the hygiene fixture is ~18 letters) stays on pdf.js.
 */
export const PDF_EMPTY_PAGE_LETTERS = 15;

export interface PdfWord {
  text: string;
  x: number;
  y: number;
  width: number;
  height: number;
  page: number;
}

export interface PdfLine {
  page: number;
  x: number;
  y: number;
  height: number;
  text: string;
  words: PdfWord[];
  gapXs: number[];
}

const Y_BUCKET = 4;
const LETTER_RE = /[A-Za-z\u0410-\u042F\u0430-\u044F\u0401\u0451]/g;
const HEADING_RE = /^(Глава|Раздел|Приложение|Статья|Часть)(?=\s|$)/i;
const PAGE_NUM_RE = /^(?:[|_\u2014\u2013-]\s*)?\d{1,4}(?:\s*[|_\u2014\u2013-])?\s*$/;

export function countLetters(text: string): number {
  const matched = text.match(LETTER_RE);
  return matched ? matched.length : 0;
}

function median(nums: number[]): number {
  if (!nums.length) return 10;
  const sorted = nums.slice().sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  if (sorted.length % 2) return sorted[mid];
  return (sorted[mid - 1] + sorted[mid]) / 2;
}

export function wordsToLines(words: PdfWord[]): PdfLine[] {
  const byPage = new Map<number, PdfWord[]>();
  for (const word of words) {
    if (!word.text || !word.text.trim()) continue;
    const list = byPage.get(word.page);
    if (list) list.push(word);
    else byPage.set(word.page, [word]);
  }
  const pages = Array.from(byPage.keys()).sort((a, b) => a - b);
  const lines: PdfLine[] = [];
  for (const page of pages) {
    const pageWords = (byPage.get(page) || []).slice().sort((a, b) => b.y - a.y || a.x - b.x);
    const heights = pageWords.map((word) => word.height).filter((height) => height > 0);
    const tol = Math.max(Y_BUCKET, median(heights) * 0.6);
    const rows: Array<{ y: number; words: PdfWord[] }> = [];
    for (const word of pageWords) {
      const bucket = rows.find((row) => Math.abs(row.y - word.y) < tol);
      if (bucket) bucket.words.push(word);
      else rows.push({ y: word.y, words: [word] });
    }
    rows.sort((a, b) => b.y - a.y);
    for (const row of rows) {
      row.words.sort((a, b) => a.x - b.x);
      const gapXs: number[] = [];
      const parts: string[] = [];
      let height = 0;
      for (let i = 0; i < row.words.length; i++) {
        const word = row.words[i];
        if (word.height > height) height = word.height;
        if (i > 0) {
          const prev = row.words[i - 1];
          const gap = word.x - (prev.x + Math.max(prev.width, 0));
          const font = Math.max(prev.height || 0, word.height || 0, 8);
          if (gap > Math.max(font * 3, 14)) {
            gapXs.push(prev.x + Math.max(prev.width, 0) + gap / 2);
          }
        }
        parts.push(word.text.trim());
      }
      const text = parts.join(' ').replace(/\s+/g, ' ').trim();
      if (!text) continue;
      lines.push({
        page,
        x: row.words[0].x,
        y: row.y,
        height: height || 10,
        text,
        words: row.words,
        gapXs,
      });
    }
  }
  return lines;
}

function fixNumero(text: string): string {
  if (!/[\u0410-\u042F\u0430-\u044F\u0401\u0451]/.test(text)) return text;
  return text.replace(/(^|[\s(])No(?=\s*\d)/g, '$1\u2116');
}

function isClauseToken(num: string): boolean {
  if (/^\d{2}\.\d{2}\.\d{2,4}$/.test(num)) return false;
  const parts = num.split('.');
  for (const part of parts) {
    if (part.length > 3) return false;
  }
  return true;
}

function headingLevel(text: string): 1 | 2 | 3 {
  if (/^Глава(?:\s|$)/i.test(text) || /^Часть(?:\s|$)/i.test(text)) return 1;
  if (/^Раздел(?:\s|$)/i.test(text) || /^Приложение(?:\s|$)/i.test(text)) return 2;
  return 3;
}

function headingBlock(text: string): ContentBlock {
  return { type: 'heading', level: headingLevel(text), runs: [{ text }] };
}

function paragraph(text: string): ContentBlock {
  return { type: 'paragraph', runs: [{ text }] };
}

function clauseParagraph(num: string, rest: string): ContentBlock {
  const runs: InlineRun[] = [{ text: num + '.', bold: true }];
  if (rest) runs.push({ text: ' ' + rest });
  return { type: 'paragraph', runs };
}

function isPageNumber(text: string): boolean {
  return PAGE_NUM_RE.test(text.trim());
}

function dropRepeatedHeaders(lines: PdfLine[]): PdfLine[] {
  const counts = new Map<string, number>();
  for (const line of lines) {
    const text = line.text.trim();
    if (text.length < 2 || text.length > 60) continue;
    counts.set(text, (counts.get(text) || 0) + 1);
  }
  const repeated = new Set<string>();
  counts.forEach((count, text) => {
    if (count >= 3) repeated.add(text);
  });
  if (!repeated.size) return lines;
  const byPage = new Map<number, PdfLine[]>();
  for (const line of lines) {
    const list = byPage.get(line.page);
    if (list) list.push(line);
    else byPage.set(line.page, [line]);
  }
  const out: PdfLine[] = [];
  const pages = Array.from(byPage.keys()).sort((a, b) => a - b);
  for (const page of pages) {
    const pageLines = byPage.get(page) || [];
    for (let i = 0; i < pageLines.length; i++) {
      const edge = i === 0 || i === pageLines.length - 1;
      if (edge && repeated.has(pageLines[i].text.trim())) continue;
      out.push(pageLines[i]);
    }
  }
  return out;
}

function joinHyphens(lines: PdfLine[]): PdfLine[] {
  const out: PdfLine[] = [];
  for (const line of lines) {
    const prev = out[out.length - 1];
    const cont = line.text.trim();
    if (
      prev &&
      /[A-Za-z\u0410-\u042F\u0430-\u044F\u0401\u0451]-$/.test(prev.text) &&
      /^[a-z\u0430-\u044F\u0451]/.test(cont) &&
      !HEADING_RE.test(cont)
    ) {
      prev.text = prev.text.replace(/-$/, '') + cont;
      prev.words = prev.words.concat(line.words);
      prev.gapXs = prev.gapXs.concat(line.gapXs);
      if (line.height > prev.height) prev.height = line.height;
      continue;
    }
    out.push({
      page: line.page,
      x: line.x,
      y: line.y,
      height: line.height,
      text: line.text,
      words: line.words.slice(),
      gapXs: line.gapXs.slice(),
    });
  }
  return out;
}

function looksLikeTable(group: PdfLine[]): boolean {
  if (group.length < 3) return false;
  const withGaps = group.filter((line) => line.gapXs.length >= 2);
  if (withGaps.length < 3) return false;
  const medH = median(withGaps.map((line) => line.height));
  const tol = Math.max(18, medH * 1.5);
  const first = withGaps[0].gapXs;
  let aligned = 0;
  for (const line of withGaps) {
    const ok = first.every((x) => line.gapXs.some((gap) => Math.abs(gap - x) < tol));
    if (ok) aligned++;
  }
  return aligned >= 3;
}

const LONE_MARKER_RE = /^[\u2014\u2013\-\u2022\u2219\u00b7]$/;

function isPipeNoise(text: string): boolean {
  const trimmed = text.trim();
  return trimmed.length > 0 && /^[|\s]+$/.test(trimmed);
}

function startsNewBlock(text: string): boolean {
  const body = text.trim();
  if (HEADING_RE.test(body)) return true;
  const match = body.match(/^(\d+\.\d+(?:\.\d+)*)\.?(?:\s|$)/);
  return !!(match && isClauseToken(match[1]));
}

function medianSamePageGap(lines: PdfLine[]): number {
  const gaps: number[] = [];
  for (let n = 1; n < lines.length; n++) {
    if (lines[n].page !== lines[n - 1].page) continue;
    const dy = Math.abs(lines[n - 1].y - lines[n].y);
    if (dy > 0.5) gaps.push(dy);
  }
  if (!gaps.length) return 14;
  return median(gaps);
}

function lineToCells(line: PdfLine): string[] {
  const gaps = line.gapXs.slice().sort((a, b) => a - b);
  const cells: string[] = [];
  let buf: string[] = [];
  let gapI = 0;
  const words = line.words.slice().sort((a, b) => a.x - b.x);
  for (const word of words) {
    while (gapI < gaps.length && word.x > gaps[gapI]) {
      cells.push(buf.join(' ').trim());
      buf = [];
      gapI++;
    }
    buf.push(word.text.trim());
  }
  cells.push(buf.join(' ').trim());
  return cells;
}

function emitLine(blocks: ContentBlock[], raw: string, tall: boolean): void {
  let body = fixNumero(raw.trim());
  if (!body) return;
  if (HEADING_RE.test(body)) {
    const idx = body.search(/\d+\.\d+(?:\.\d+)*\.?\s+/);
    if (idx > 0) {
      const head = body.slice(0, idx).trim();
      blocks.push(headingBlock(head));
      body = body.slice(idx).trim();
    } else {
      blocks.push(headingBlock(body));
      return;
    }
  } else if (tall && body.length < 120 && !/^\d+\.\d+/.test(body)) {
    blocks.push(headingBlock(body));
    return;
  }

  const re = /(\d+\.\d+(?:\.\d+)*)\.?\s+/g;
  const indexes: Array<{ num: string; end: number; start: number }> = [];
  let match: RegExpExecArray | null;
  while ((match = re.exec(body))) {
    if (!isClauseToken(match[1])) continue;
    indexes.push({ num: match[1], start: match.index, end: re.lastIndex });
  }
  if (!indexes.length) {
    blocks.push(paragraph(body));
    return;
  }
  if (indexes[0].start > 0) {
    const prefix = body.slice(0, indexes[0].start).trim();
    if (prefix) blocks.push(paragraph(prefix));
  }
  for (let i = 0; i < indexes.length; i++) {
    const from = indexes[i].end;
    const to = i + 1 < indexes.length ? indexes[i + 1].start : body.length;
    blocks.push(clauseParagraph(indexes[i].num, body.slice(from, to).trim()));
  }
}

export function linesToBlocks(lines: PdfLine[]): {
  blocks: ContentBlock[];
  warnings: DocumentWarning[];
} {
  const warnings: DocumentWarning[] = [];
  const cleaned = joinHyphens(
    dropRepeatedHeaders(lines.filter((line) => !isPageNumber(line.text)))
  );
  const med = median(cleaned.map((line) => line.height).filter((height) => height > 0));
  const gapLimit = medianSamePageGap(cleaned) * 1.6;
  const blocks: ContentBlock[] = [];
  let tableUsed = false;
  let i = 0;
  let pendingList = false;
  let buf: {
    text: string;
    height: number;
    y: number;
    page: number;
    list: boolean;
  } | null = null;

  const flush = (): void => {
    if (!buf || !buf.text.trim()) {
      buf = null;
      return;
    }
    const text = buf.text.trim();
    const words = text.split(/\s+/).length;
    const tall =
      buf.height > med * 1.85 &&
      text.length < 80 &&
      words <= 8 &&
      !/[.;,]/.test(text);
    const asList = buf.list;
    buf = null;
    if (asList) {
      blocks.push({
        type: 'list_item',
        level: 0,
        ordered: false,
        runs: [{ text: fixNumero(text) }],
      });
      return;
    }
    emitLine(blocks, text, tall);
  };

  const take = (line: PdfLine, asList: boolean): void => {
    const piece = line.text.trim();
    if (!buf) {
      buf = {
        text: piece,
        height: line.height,
        y: line.y,
        page: line.page,
        list: asList,
      };
      return;
    }
    buf.text = buf.text + ' ' + piece;
    if (line.height > buf.height) buf.height = line.height;
    buf.y = line.y;
  };

  while (i < cleaned.length) {
    const line = cleaned[i];
    if (isPipeNoise(line.text)) {
      i++;
      continue;
    }
    if (line.gapXs.length >= 2) {
      const group: PdfLine[] = [];
      let k = i;
      while (
        k < cleaned.length &&
        cleaned[k].page === line.page &&
        cleaned[k].gapXs.length >= 2
      ) {
        group.push(cleaned[k]);
        k++;
      }
      if (looksLikeTable(group)) {
        flush();
        pendingList = false;
        const rows = group.map(lineToCells);
        let width = 0;
        for (const row of rows) if (row.length > width) width = row.length;
        const norm = rows.map((row) => {
          const copy = row.slice();
          while (copy.length < width) copy.push('');
          return copy;
        });
        blocks.push({ type: 'table', rows: norm });
        tableUsed = true;
        i = k;
        continue;
      }
    }
    if (LONE_MARKER_RE.test(line.text.trim())) {
      flush();
      pendingList = true;
      i++;
      continue;
    }
    const join = (function (
      prev: { page: number; y: number } | null
    ): boolean {
      if (prev === null) return false;
      if (prev.page !== line.page) return false;
      if (startsNewBlock(line.text)) return false;
      return Math.abs(prev.y - line.y) <= gapLimit;
    })(buf);
    if (!join) {
      flush();
      take(line, pendingList);
      pendingList = false;
    } else {
      take(line, false);
    }
    i++;
  }
  flush();
  if (tableUsed) {
    warnings.push({
      code: 'PDF_TABLE_HEURISTIC',
      message:
        'Таблицы PDF собраны по большим промежуткам между словами и могут расходиться с формой (журналы, акты). Сверьте с оригиналом.',
    });
  }
  return { blocks, warnings };
}
