import type { ContentBlock, InlineRun } from '../types.js';
import { escapeHtml } from '../utils/text.js';

export type { ContentBlock, InlineRun };

export function runsToPlain(runs: InlineRun[]): string {
  return runs.map((r) => r.text).join('');
}

export function flattenBlocksToParagraphs(blocks: ContentBlock[]): string[] {
  const out: string[] = [];
  for (const b of blocks) {
    if (b.type === 'heading' || b.type === 'paragraph' || b.type === 'list_item') {
      const t = runsToPlain(b.runs).trim();
      // Keep edit targets free of GFM footnote definition lines.
      if (t && !/^\[\^[\w-]+\]:/.test(t)) out.push(t);
    } else if (b.type === 'table') {
      for (const row of b.rows) {
        const line = row.map((c) => String(c ?? '').trim()).join(' | ');
        if (line.replace(/\s*\|\s*/g, '').trim()) out.push(line);
      }
    } else if (b.type === 'code') {
      const t = b.text.trim();
      if (t) out.push(t);
    }
  }
  return out;
}

function escapeMdCell(s: string): string {
  return String(s ?? '')
    .replace(/\|/g, '\\|')
    .replace(/\n/g, ' ')
    .trim();
}

/** Escape only characters that break inline markdown emphasis/links. */
function escapeMdInline(s: string): string {
  return s.replace(/([\\`*_\[\]])/g, '\\$1');
}

function runKey(r: InlineRun): string {
  return `${r.bold ? 1 : 0}|${r.italic ? 1 : 0}|${r.strike ? 1 : 0}|${r.href ?? ''}`;
}

function isFootnoteAtom(text: string): boolean {
  return /^\[\^[\w-]+\](?::[\s\S]*)?$/.test(text);
}

/** Merge adjacent runs with identical formatting. */
export function mergeRuns(runs: InlineRun[]): InlineRun[] {
  if (!runs.length) return runs;
  const out: InlineRun[] = [{ ...runs[0], text: runs[0].text }];
  for (let i = 1; i < runs.length; i++) {
    const prev = out[out.length - 1];
    const cur = runs[i];
    // Keep GFM footnote markers as standalone atoms so emit can leave brackets raw.
    if (isFootnoteAtom(prev.text) || isFootnoteAtom(cur.text)) {
      out.push({ ...cur, text: cur.text });
      continue;
    }
    if (runKey(prev) === runKey(cur)) {
      prev.text += cur.text;
    } else {
      out.push({ ...cur, text: cur.text });
    }
  }
  return out;
}

export function runsToMarkdown(runs: InlineRun[]): string {
  // Use ** for bold and _ for italic so adjacent runs cannot form **** / *** traps
  // that hang markdown-it (VS Code / Cursor preview). Keep spaces outside markers.
  return mergeRuns(runs)
    .map((r) => {
      if (/^\[\^[\w-]+\](?::[\s\S]*)?$/.test(r.text) && !r.href && !r.bold && !r.italic) {
        return r.text;
      }
      const m = /^(?<lead>\s*)(?<core>[\s\S]*?)(?<trail>\s*)$/.exec(r.text);
      const lead = m?.groups?.lead ?? '';
      const core = m?.groups?.core ?? r.text;
      const trail = m?.groups?.trail ?? '';
      if (!core) return escapeMdInline(r.text);

      let t = escapeMdInline(core);
      if (r.strike) t = `~~${t}~~`;
      if (r.bold && r.italic) t = `**_${t}_**`;
      else if (r.bold) t = `**${t}**`;
      else if (r.italic) t = `_${t}_`;
      if (r.href) t = `[${t}](${r.href})`;
      return `${lead}${t}${trail}`;
    })
    .join('');
}

export function emitMarkdownTable(rows: string[][]): string {
  if (!rows.length) return '';
  const width = Math.max(...rows.map((r) => r.length), 1);
  let norm = rows.map((r) => {
    const cells = r.map((c) => String(c ?? '').trim());
    while (cells.length < width) cells.push('');
    return cells.slice(0, width);
  });

  // Drop columns that are empty in every row (common after Word merged-cell calendars).
  const keepCols: number[] = [];
  for (let c = 0; c < width; c++) {
    if (norm.some((row) => row[c])) keepCols.push(c);
  }
  if (!keepCols.length) return '';
  norm = norm.map((row) => keepCols.map((c) => row[c]));

  // Drop fully empty body rows (keep header even if sparse).
  norm = norm.filter((row, i) => i === 0 || row.some((c) => c));

  const escaped = norm.map((r) => r.map(escapeMdCell));
  const header = escaped[0];
  const sep = header.map(() => '---');
  const lines = [
    `| ${header.join(' | ')} |`,
    `| ${sep.join(' | ')} |`,
    ...escaped.slice(1).map((r) => `| ${r.join(' | ')} |`),
  ];
  return lines.join('\n');
}

export function emitMarkdownFromBlocks(blocks: ContentBlock[]): string {
  const parts: string[] = [];
  let listBuf: string[] = [];

  const flushList = () => {
    if (listBuf.length) {
      parts.push(listBuf.join('\n'));
      listBuf = [];
    }
  };

  for (const b of blocks) {
    if (b.type === 'list_item') {
      const indent = '  '.repeat(Math.max(0, b.level));
      const marker = b.ordered ? '1.' : '-';
      listBuf.push(`${indent}${marker} ${runsToMarkdown(b.runs)}`.trimEnd());
      continue;
    }
    flushList();
    if (b.type === 'heading') {
      const hashes = '#'.repeat(b.level);
      parts.push(`${hashes} ${runsToMarkdown(b.runs)}`.trimEnd());
    } else if (b.type === 'paragraph') {
      const line = runsToMarkdown(b.runs).trim();
      if (line) parts.push(line);
    } else if (b.type === 'table') {
      const md = emitMarkdownTable(b.rows);
      if (md) parts.push(md);
    } else if (b.type === 'thematic_break') {
      parts.push('---');
    } else if (b.type === 'code') {
      const lang = b.language ?? '';
      parts.push(`\`\`\`${lang}\n${b.text.replace(/\n$/, '')}\n\`\`\``);
    }
  }
  flushList();
  return parts.join('\n\n').trim();
}

export function emitHtmlFromBlocks(blocks: ContentBlock[]): string {
  const parts: string[] = [];
  let listOpen: 'ul' | 'ol' | null = null;

  const closeList = () => {
    if (listOpen) {
      parts.push(`</${listOpen}>`);
      listOpen = null;
    }
  };

  const runsToHtml = (runs: InlineRun[]): string =>
    mergeRuns(runs)
      .map((r) => {
        let t = escapeHtml(r.text);
        if (r.strike) t = `<s>${t}</s>`;
        if (r.bold) t = `<strong>${t}</strong>`;
        if (r.italic) t = `<em>${t}</em>`;
        if (r.href) t = `<a href="${escapeHtml(r.href)}">${t}</a>`;
        return t;
      })
      .join('');

  for (const b of blocks) {
    if (b.type === 'list_item') {
      const tag = b.ordered ? 'ol' : 'ul';
      if (listOpen !== tag) {
        closeList();
        parts.push(`<${tag}>`);
        listOpen = tag;
      }
      parts.push(`<li>${runsToHtml(b.runs)}</li>`);
      continue;
    }
    closeList();
    if (b.type === 'heading') {
      parts.push(`<h${b.level}>${runsToHtml(b.runs)}</h${b.level}>`);
    } else if (b.type === 'paragraph') {
      parts.push(`<p>${runsToHtml(b.runs)}</p>`);
    } else if (b.type === 'table') {
      const body = b.rows
        .map((row, i) => {
          const cell = i === 0 ? 'th' : 'td';
          return `<tr>${row
            .map((c) => `<${cell}>${escapeHtml(String(c ?? ''))}</${cell}>`)
            .join('')}</tr>`;
        })
        .join('');
      parts.push(`<table>${body}</table>`);
    } else if (b.type === 'thematic_break') {
      parts.push('<hr />');
    } else if (b.type === 'code') {
      parts.push(`<pre><code>${escapeHtml(b.text)}</code></pre>`);
    }
  }
  closeList();
  return `<!DOCTYPE html>\n<html><body>\n${parts.join('\n')}\n</body></html>\n`;
}

/** Build blocks from spreadsheet sheets (heading per sheet + table). */
export function blocksFromSpreadsheet(
  sheets: Record<string, (string | number | boolean | null)[][]>
): ContentBlock[] {
  const blocks: ContentBlock[] = [];
  for (const [name, grid] of Object.entries(sheets)) {
    blocks.push({
      type: 'heading',
      level: 2,
      runs: [{ text: name }],
    });
    if (grid.length) {
      blocks.push({
        type: 'table',
        rows: grid.map((row) => row.map((c) => (c == null ? '' : String(c)))),
      });
    }
  }
  return blocks;
}

export function blocksFromCsvRows(rows: string[][]): ContentBlock[] {
  if (!rows.length) return [];
  return [{ type: 'table', rows }];
}
