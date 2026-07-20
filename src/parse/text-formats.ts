import { readFile } from 'node:fs/promises';

export type TextFormat = 'md' | 'html' | 'csv';

export async function readTextFile(path: string): Promise<string> {
  const text = await readFile(path, 'utf8');
  return text.startsWith('\uFEFF') ? text.slice(1) : text;
}

export function parseMarkdown(text: string): {
  format: 'md';
  text: string;
  paragraphs: string[];
} {
  const paragraphs = text
    .split(/\n\s*\n/)
    .map((p) => p.trim())
    .filter(Boolean);
  return { format: 'md', text: text.replace(/\r\n/g, '\n'), paragraphs };
}

/** Strip tags for plain text; keep a simple paragraph list from block tags. */
export function parseHtml(text: string): {
  format: 'html';
  text: string;
  paragraphs: string[];
} {
  const normalized = text.replace(/\r\n/g, '\n');
  const withoutScripts = normalized
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '');
  const withBreaks = withoutScripts
    .replace(/<\/(p|div|h[1-6]|li|tr|br)\s*>/gi, '\n')
    .replace(/<br\s*\/?>/gi, '\n');
  const plain = withBreaks
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  const paragraphs = plain
    .split(/\n\s*\n/)
    .map((p) => p.trim())
    .filter(Boolean);
  return { format: 'html', text: plain, paragraphs };
}

export function parseCsv(text: string): {
  format: 'csv';
  text: string;
  rows: string[][];
} {
  const normalized = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const rows: string[][] = [];
  let row: string[] = [];
  let cur = '';
  let inQuotes = false;

  for (let i = 0; i < normalized.length; i++) {
    const ch = normalized[i];
    if (inQuotes) {
      if (ch === '"' && normalized[i + 1] === '"') {
        cur += '"';
        i++;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        cur += ch;
      }
      continue;
    }
    if (ch === '"') {
      inQuotes = true;
    } else if (ch === ',') {
      row.push(cur);
      cur = '';
    } else if (ch === '\n') {
      row.push(cur);
      if (row.some((c) => c !== '') || rows.length > 0) rows.push(row);
      row = [];
      cur = '';
    } else {
      cur += ch;
    }
  }
  row.push(cur);
  if (row.some((c) => c !== '') || rows.length > 0) rows.push(row);

  while (rows.length && rows[rows.length - 1].every((c) => c === '')) {
    rows.pop();
  }
  return { format: 'csv', text: normalized, rows };
}
