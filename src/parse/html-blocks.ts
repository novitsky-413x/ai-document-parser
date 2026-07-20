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

/** Find the end index (exclusive) of a balanced HTML element starting at `start`. */
function findBalancedHtmlEnd(html: string, start: number): number {
  const open = /^<([A-Za-z][\w:-]*)\b[^>]*>/.exec(html.slice(start));
  if (!open) return -1;
  const tag = open[1].toLowerCase();
  if (/\/>$/.test(open[0]) || /^(br|hr|img|meta|link|input)$/i.test(tag)) {
    return start + open[0].length;
  }
  const openRe = new RegExp(`<${tag}\\b[^>]*>`, 'gi');
  const closeRe = new RegExp(`</${tag}\\s*>`, 'gi');
  let depth = 0;
  let i = start;
  while (i < html.length) {
    openRe.lastIndex = i;
    closeRe.lastIndex = i;
    const om = openRe.exec(html);
    const cm = closeRe.exec(html);
    const oAt = om ? om.index : -1;
    const cAt = cm ? cm.index : -1;
    if (oAt < 0 && cAt < 0) return -1;
    if (oAt >= 0 && (cAt < 0 || oAt < cAt)) {
      const selfClosing = /\/>$/.test(om![0]);
      if (!selfClosing) depth++;
      i = oAt + om![0].length;
    } else {
      depth--;
      i = cAt + cm![0].length;
      if (depth === 0) return i;
    }
  }
  return -1;
}

function indexOfNextBlockTag(html: string, from: number): number {
  const re = /<(h[1-6]|p|ul|ol|table|pre|hr|blockquote)\b/gi;
  re.lastIndex = from;
  const m = re.exec(html);
  return m ? m.index : -1;
}

function parseInline(html: string): InlineRun[] {
  const runs: InlineRun[] = [];
  // Prefer specific inline tags; fall back to text nodes.
  const re =
    /<a\b[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>|<code\b[^>]*>([\s\S]*?)<\/code>|<(strong|b)(\s[^>]*)?>((?:[\s\S]*?))<\/\4>|<(em|i)(\s[^>]*)?>((?:[\s\S]*?))<\/\7>|<(s|strike|del)(\s[^>]*)?>((?:[\s\S]*?))<\/\10>|([^<]+)/gi;
  let m: RegExpExecArray | null;
  while ((m = re.exec(html)) !== null) {
    if (m[1] != null && m[2] != null) {
      const text = decodeEntities(m[2].replace(/<[^>]+>/g, ''));
      if (text) runs.push({ text, href: decodeEntities(m[1]) });
    } else if (m[3] != null) {
      const text = decodeEntities(m[3].replace(/<[^>]+>/g, ''));
      if (text) runs.push({ text, code: true });
    } else if (m[4]) {
      const inner = m[6];
      const nested = /<(em|i)(\s[^>]*)?>([\s\S]*?)<\/\1>/i.exec(inner);
      if (nested) {
        const before = decodeEntities(inner.slice(0, nested.index).replace(/<[^>]+>/g, ''));
        if (before) runs.push({ text: before, bold: true });
        const mid = decodeEntities(nested[3].replace(/<[^>]+>/g, ''));
        if (mid) runs.push({ text: mid, bold: true, italic: true });
        const after = decodeEntities(
          inner.slice(nested.index + nested[0].length).replace(/<[^>]+>/g, '')
        );
        if (after) runs.push({ text: after, bold: true });
      } else {
        const text = decodeEntities(inner.replace(/<[^>]+>/g, ''));
        if (text) runs.push({ text, bold: true });
      }
    } else if (m[7]) {
      const text = decodeEntities(m[9].replace(/<[^>]+>/g, ''));
      if (text) runs.push({ text, italic: true });
    } else if (m[10]) {
      const text = decodeEntities(m[12].replace(/<[^>]+>/g, ''));
      if (text) runs.push({ text, strike: true });
    } else if (m[13]) {
      const t = decodeEntities(m[13]);
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

function extractCodeLanguage(preOpen: string, codeOpen: string): string | undefined {
  const fromClass =
    /(?:^|\s)class=["'][^"']*\blanguage-([A-Za-z0-9_+-]+)/i.exec(codeOpen)?.[1] ??
    /(?:^|\s)class=["'][^"']*\blanguage-([A-Za-z0-9_+-]+)/i.exec(preOpen)?.[1];
  return fromClass || undefined;
}

function parseListHtml(
  listHtml: string,
  depth: number,
  ordered: boolean
): ContentBlock[] {
  const blocks: ContentBlock[] = [];
  const inner = listHtml.replace(/^<(ul|ol)\b[^>]*>/i, '').replace(/<\/(ul|ol)\s*>$/i, '');
  let i = 0;
  while (i < inner.length) {
    const liAt = inner.toLowerCase().indexOf('<li', i);
    if (liAt < 0) break;
    const after = inner.charAt(liAt + 3);
    if (after && !/[\s/>]/.test(after)) {
      i = liAt + 3;
      continue;
    }
    const end = findBalancedHtmlEnd(inner, liAt);
    if (end < 0) {
      i = liAt + 3;
      continue;
    }
    const itemXml = inner.slice(liAt, end);
    const openEnd = itemXml.indexOf('>') + 1;
    let body = itemXml.slice(openEnd).replace(/<\/li\s*>$/i, '');

    // Nested lists inside this item.
    const nestedBlocks: ContentBlock[] = [];
    const nestedRe = /<(ul|ol)\b/gi;
    let nm: RegExpExecArray | null;
    let cursor = 0;
    const textParts: string[] = [];
    while ((nm = nestedRe.exec(body)) !== null) {
      textParts.push(body.slice(cursor, nm.index));
      const nEnd = findBalancedHtmlEnd(body, nm.index);
      if (nEnd < 0) break;
      const nestedXml = body.slice(nm.index, nEnd);
      const isOl = /^<ol\b/i.test(nestedXml);
      nestedBlocks.push(...parseListHtml(nestedXml, depth + 1, isOl));
      cursor = nEnd;
      nestedRe.lastIndex = nEnd;
    }
    textParts.push(body.slice(cursor));
    const runs = parseInline(textParts.join(''));
    if (runsToText(runs).trim() || !nestedBlocks.length) {
      blocks.push({
        type: 'list_item',
        level: depth,
        ...(ordered ? { ordered: true } : {}),
        runs: runs.length ? runs : [{ text: ' ' }],
      });
    }
    blocks.push(...nestedBlocks);
    i = end;
  }
  return blocks;
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
  let i = 0;
  const pushLoose = (chunk: string) => {
    const plain = decodeEntities(chunk.replace(/<[^>]+>/g, '')).trim();
    if (plain) blocks.push({ type: 'paragraph', runs: [{ text: plain }] });
  };

  while (i < normalized.length) {
    const at = indexOfNextBlockTag(normalized, i);
    if (at < 0) {
      pushLoose(normalized.slice(i));
      break;
    }
    if (at > i) pushLoose(normalized.slice(i, at));

    const open = /^<([A-Za-z][\w:-]*)\b[^>]*>/i.exec(normalized.slice(at));
    if (!open) {
      i = at + 1;
      continue;
    }
    const tag = open[1].toLowerCase();

    if (tag === 'hr') {
      blocks.push({ type: 'thematic_break' });
      i = at + open[0].length;
      continue;
    }

    const end = findBalancedHtmlEnd(normalized, at);
    if (end < 0) {
      i = at + open[0].length;
      continue;
    }
    const el = normalized.slice(at, end);

    if (/^h[1-6]$/.test(tag)) {
      const level = Number(tag[1]) as 1 | 2 | 3 | 4 | 5 | 6;
      const inner = el.replace(/^<h[1-6]\b[^>]*>/i, '').replace(/<\/h[1-6]\s*>$/i, '');
      blocks.push({ type: 'heading', level, runs: parseInline(inner) });
    } else if (tag === 'p') {
      const inner = el.replace(/^<p\b[^>]*>/i, '').replace(/<\/p\s*>$/i, '');
      const runs = parseInline(inner);
      if (runsToText(runs).trim()) blocks.push({ type: 'paragraph', runs });
    } else if (tag === 'table') {
      const table = parseTableHtml(el);
      if (table.type === 'table' && table.rows.length) blocks.push(table);
    } else if (tag === 'pre') {
      const preOpen = /^<pre\b[^>]*>/i.exec(el)?.[0] ?? '';
      const codeOpen = /<code\b[^>]*>/i.exec(el)?.[0] ?? '';
      const language = extractCodeLanguage(preOpen, codeOpen);
      const inner = el
        .replace(/^<pre\b[^>]*>/i, '')
        .replace(/<\/pre\s*>$/i, '')
        .replace(/^<code\b[^>]*>/i, '')
        .replace(/<\/code\s*>$/i, '');
      blocks.push({
        type: 'code',
        text: decodeEntities(inner.replace(/<[^>]+>/g, '')),
        ...(language ? { language } : {}),
      });
    } else if (tag === 'ul' || tag === 'ol') {
      blocks.push(...parseListHtml(el, 0, tag === 'ol'));
    } else if (tag === 'blockquote') {
      const inner = el.replace(/^<blockquote\b[^>]*>/i, '').replace(/<\/blockquote\s*>$/i, '');
      const nested = htmlToBlocks(inner);
      if (nested.length) blocks.push({ type: 'blockquote', blocks: nested });
      else {
        const runs = parseInline(inner);
        if (runsToText(runs).trim()) {
          blocks.push({ type: 'blockquote', blocks: [{ type: 'paragraph', runs }] });
        }
      }
    }

    i = end;
  }

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
        runs: parseMarkdownInline(hm[2]),
      });
      i++;
      continue;
    }
    if (/^---+$/.test(line.trim())) {
      blocks.push({ type: 'thematic_break' });
      i++;
      continue;
    }
    if (/^\s*>\s?/.test(line)) {
      const quoteLines: string[] = [];
      while (i < lines.length && /^\s*>\s?/.test(lines[i])) {
        quoteLines.push(lines[i].replace(/^\s*>\s?/, ''));
        i++;
      }
      const nested = markdownToBlocks(quoteLines.join('\n'));
      blocks.push({
        type: 'blockquote',
        blocks: nested.length
          ? nested
          : [{ type: 'paragraph', runs: parseMarkdownInline(quoteLines.join(' ')) }],
      });
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
    while (
      i < lines.length &&
      !/^\s*$/.test(lines[i]) &&
      !/^(#{1,6})\s+|^\s*>|^\s*[-*+]\s+|^\s*\d+\.\s+|^\s*\||^```/.test(lines[i])
    ) {
      para.push(lines[i]);
      i++;
    }
    blocks.push({ type: 'paragraph', runs: parseMarkdownInline(para.join(' ')) });
  }
  return blocks;
}

/**
 * Parse inline markdown emphasis/links/code.
 * Supports both *italic* / _italic_ and ~~strike~~ so emit↔parse stay aligned.
 */
export function parseMarkdownInline(text: string): InlineRun[] {
  const runs: InlineRun[] = [];
  // Order: code, links, bold+italic combos, bold, strike, italic, plain.
  const re =
    /(`([^`]+)`)|(\[([^\]]+)\]\(([^)]+)\))|(\*\*_([^*]+)_\*\*)|(\*\*([^*]+)\*\*)|(~~([^~]+)~~)|(\*([^*]+)\*)|((?<!\w)_([^_]+)_(?!\w))|([^*`[~_\]]+|[_*~\[\]])/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    if (m[2]) runs.push({ text: m[2], code: true });
    else if (m[4] && m[5]) runs.push({ text: m[4], href: m[5] });
    else if (m[7]) runs.push({ text: m[7], bold: true, italic: true });
    else if (m[9]) runs.push({ text: m[9], bold: true });
    else if (m[11]) runs.push({ text: m[11], strike: true });
    else if (m[13]) runs.push({ text: m[13], italic: true });
    else if (m[15]) runs.push({ text: m[15], italic: true });
    else if (m[16]) runs.push({ text: m[16] });
  }
  if (!runs.length && text) runs.push({ text });
  return mergeAdjacentPlain(runs);
}

function mergeAdjacentPlain(runs: InlineRun[]): InlineRun[] {
  if (runs.length < 2) return runs;
  const out: InlineRun[] = [{ ...runs[0] }];
  for (let i = 1; i < runs.length; i++) {
    const prev = out[out.length - 1];
    const cur = runs[i];
    const plain =
      !prev.bold &&
      !prev.italic &&
      !prev.strike &&
      !prev.code &&
      !prev.href &&
      !cur.bold &&
      !cur.italic &&
      !cur.strike &&
      !cur.code &&
      !cur.href;
    if (plain) prev.text += cur.text;
    else out.push({ ...cur });
  }
  return out;
}
