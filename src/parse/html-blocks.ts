import type { ContentBlock, InlineRun } from '../types.js';
import { flattenBlocksToParagraphs } from '../convert/blocks.js';

function decodeEntities(s: string): string {
  return s
    .replace(/&nbsp;/gi, ' ')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&apos;/gi, "'")
    .replace(/&amp;/gi, '&');
}

function parseInline(html: string): InlineRun[] {
  const runs: InlineRun[] = [];
  const re =
    /<(strong|b)(\s[^>]*)?>((?:[\s\S]*?))<\/\1>|<(em|i)(\s[^>]*)?>((?:[\s\S]*?))<\/\4>|([^<]+)/gi;
  let m: RegExpExecArray | null;
  while ((m = re.exec(html)) !== null) {
    if (m[1]) {
      const inner = m[3];
      // Nested em/i inside strong/b
      const nested = /<(em|i)(\s[^>]*)?>([\s\S]*?)<\/\1>/i.exec(inner);
      if (nested) {
        const before = decodeEntities(inner.slice(0, nested.index).replace(/<[^>]+>/g, ''));
        if (before) runs.push({ text: before, bold: true });
        const mid = decodeEntities(nested[3].replace(/<[^>]+>/g, ''));
        if (mid) runs.push({ text: mid, bold: true, italic: true });
        const after = decodeEntities(inner.slice(nested.index + nested[0].length).replace(/<[^>]+>/g, ''));
        if (after) runs.push({ text: after, bold: true });
      } else {
        const text = decodeEntities(inner.replace(/<[^>]+>/g, ''));
        if (text) runs.push({ text, bold: true });
      }
    } else if (m[4]) {
      const text = decodeEntities(m[6].replace(/<[^>]+>/g, ''));
      if (text) runs.push({ text, italic: true });
    } else if (m[7]) {
      const t = decodeEntities(m[7]);
      if (t) runs.push({ text: t });
    }
  }
  if (!runs.length) {
    const plain = decodeEntities(html.replace(/<[^>]+>/g, ''));
    if (plain) runs.push({ text: plain });
  }
  return runs;
}

function parseTableHtml(tableHtml: string): ContentBlock {
  const rows: string[][] = [];
  const trRe = /<tr\b[^>]*>([\s\S]*?)<\/tr>/gi;
  let tr: RegExpExecArray | null;
  while ((tr = trRe.exec(tableHtml)) !== null) {
    const cells: string[] = [];
    const cellRe = /<(td|th)\b[^>]*>([\s\S]*?)<\/\1>/gi;
    let c: RegExpExecArray | null;
    while ((c = cellRe.exec(tr[1])) !== null) {
      cells.push(decodeEntities(c[2].replace(/<[^>]+>/g, '')).trim());
    }
    if (cells.length) rows.push(cells);
  }
  return { type: 'table', rows };
}

/**
 * Convert HTML snippet into ContentBlocks (structural convert path).
 */
export function htmlToBlocks(html: string): ContentBlock[] {
  const normalized = html
    .replace(/\r\n/g, '\n')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '');

  const blocks: ContentBlock[] = [];
  const re =
    /<(h([1-6]))\b[^>]*>([\s\S]*?)<\/\1>|<p\b[^>]*>([\s\S]*?)<\/p>|<li\b[^>]*>([\s\S]*?)<\/li>|<table\b[^>]*>([\s\S]*?)<\/table>|<hr\s*\/?>|<pre\b[^>]*>([\s\S]*?)<\/pre>|<ul\b[^>]*>([\s\S]*?)<\/ul>|<ol\b[^>]*>([\s\S]*?)<\/ol>/gi;

  let m: RegExpExecArray | null;
  let last = 0;
  const pushLoose = (chunk: string) => {
    const plain = decodeEntities(chunk.replace(/<[^>]+>/g, '')).trim();
    if (plain) blocks.push({ type: 'paragraph', runs: [{ text: plain }] });
  };

  while ((m = re.exec(normalized)) !== null) {
    if (m.index > last) pushLoose(normalized.slice(last, m.index));
    last = m.index + m[0].length;

    if (m[2]) {
      const level = Number(m[2]) as 1 | 2 | 3 | 4 | 5 | 6;
      blocks.push({ type: 'heading', level, runs: parseInline(m[3]) });
    } else if (m[4] != null && m[0].toLowerCase().startsWith('<p')) {
      const runs = parseInline(m[4]);
      if (runsToText(runs).trim()) blocks.push({ type: 'paragraph', runs });
    } else if (m[5] != null && /<li/i.test(m[0])) {
      blocks.push({ type: 'list_item', level: 0, runs: parseInline(m[5]) });
    } else if (m[6] != null && /<table/i.test(m[0])) {
      const table = parseTableHtml(m[0]);
      if (table.type === 'table' && table.rows.length) blocks.push(table);
    } else if (/^<hr/i.test(m[0])) {
      blocks.push({ type: 'thematic_break' });
    } else if (m[7] != null && /<pre/i.test(m[0])) {
      blocks.push({
        type: 'code',
        text: decodeEntities(m[7].replace(/<[^>]+>/g, '')),
      });
    } else if (m[8] != null && /<ul/i.test(m[0])) {
      const liRe = /<li\b[^>]*>([\s\S]*?)<\/li>/gi;
      let li: RegExpExecArray | null;
      while ((li = liRe.exec(m[8])) !== null) {
        blocks.push({ type: 'list_item', level: 0, runs: parseInline(li[1]) });
      }
    } else if (m[9] != null && /<ol/i.test(m[0])) {
      const liRe = /<li\b[^>]*>([\s\S]*?)<\/li>/gi;
      let li: RegExpExecArray | null;
      while ((li = liRe.exec(m[9])) !== null) {
        blocks.push({
          type: 'list_item',
          level: 0,
          ordered: true,
          runs: parseInline(li[1]),
        });
      }
    }
  }
  if (last < normalized.length) pushLoose(normalized.slice(last));

  return blocks;
}

function runsToText(runs: InlineRun[]): string {
  return runs.map((r) => r.text).join('');
}

export function htmlToStructured(html: string): {
  format: 'html';
  paragraphs: string[];
  blocks: ContentBlock[];
  text: string;
} {
  const blocks = htmlToBlocks(html);
  const paragraphs = flattenBlocksToParagraphs(blocks);
  return {
    format: 'html',
    blocks,
    paragraphs,
    text: paragraphs.join('\n\n'),
  };
}

/** Lightweight MD → blocks (headings, lists, GFM tables, paragraphs). */
export function markdownToBlocks(md: string): ContentBlock[] {
  const lines = md.replace(/\r\n/g, '\n').split('\n');
  const blocks: ContentBlock[] = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (/^\s*$/.test(line)) {
      i++;
      continue;
    }
    const hm = /^(#{1,6})\s+(.*)$/.exec(line);
    if (hm) {
      blocks.push({
        type: 'heading',
        level: hm[1].length as 1 | 2 | 3 | 4 | 5 | 6,
        runs: [{ text: hm[2] }],
      });
      i++;
      continue;
    }
    if (/^---+$/.test(line.trim())) {
      blocks.push({ type: 'thematic_break' });
      i++;
      continue;
    }
    if (/^\s*(?:[-*+]|\d+\.)\s+/.test(line)) {
      while (i < lines.length && /^\s*(?:[-*+]|\d+\.)\s+/.test(lines[i])) {
        const m = /^(\s*)([-*+]|\d+\.)\s+(.*)$/.exec(lines[i])!;
        const level = Math.floor(m[1].replace(/\t/g, '  ').length / 2);
        const ordered = /^\d+\.$/.test(m[2]);
        blocks.push({
          type: 'list_item',
          level,
          ...(ordered ? { ordered: true } : {}),
          runs: parseMarkdownInline(m[3]),
        });
        i++;
      }
      continue;
    }
    if (/^\s*\|/.test(line) && i + 1 < lines.length && /^\s*\|?\s*:?-+:?/.test(lines[i + 1])) {
      const rows: string[][] = [];
      while (i < lines.length && /^\s*\|/.test(lines[i])) {
        if (/^\s*\|?\s*:?-+:?/.test(lines[i])) {
          i++;
          continue;
        }
        const cells = lines[i]
          .replace(/^\s*\|/, '')
          .replace(/\|\s*$/, '')
          .split('|')
          .map((c) => c.trim());
        rows.push(cells);
        i++;
      }
      if (rows.length) blocks.push({ type: 'table', rows });
      continue;
    }
    if (/^```/.test(line)) {
      const lang = line.replace(/^```/, '').trim();
      i++;
      const buf: string[] = [];
      while (i < lines.length && !/^```/.test(lines[i])) {
        buf.push(lines[i]);
        i++;
      }
      if (i < lines.length) i++;
      blocks.push({ type: 'code', text: buf.join('\n'), language: lang || undefined });
      continue;
    }
    const para: string[] = [line];
    i++;
    while (i < lines.length && !/^\s*$/.test(lines[i]) && !/^(#{1,6})\s+|^\s*[-*+]\s+|^\s*\||^```/.test(lines[i])) {
      para.push(lines[i]);
      i++;
    }
    blocks.push({ type: 'paragraph', runs: parseMarkdownInline(para.join(' ')) });
  }
  return blocks;
}

/** Parse inline markdown emphasis/links/code in a paragraph string. */
export function parseMarkdownInline(text: string): InlineRun[] {
  const runs: InlineRun[] = [];
  const re =
    /(\*\*([^*]+)\*\*)|(\*([^*]+)\*)|(\[([^\]]+)\]\(([^)]+)\))|(`([^`]+)`)|([^*`[\]]+)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    if (m[2]) runs.push({ text: m[2], bold: true });
    else if (m[4]) runs.push({ text: m[4], italic: true });
    else if (m[6] && m[7]) runs.push({ text: m[6], href: m[7] });
    else if (m[9]) runs.push({ text: m[9], code: true });
    else if (m[10]) runs.push({ text: m[10] });
  }
  if (!runs.length && text) runs.push({ text });
  return runs;
}
