import { csvEscapeCell, escapeHtml } from '../utils/text.js';
import type { ContentBlock, SpreadsheetStructured } from '../types.js';
import { spreadsheetToText } from '../edit/spreadsheet.js';
import {
  blocksFromSpreadsheet,
  emitHtmlFromBlocks,
  emitMarkdownFromBlocks,
  emitMarkdownTable,
} from './blocks.js';

export { csvEscapeCell, escapeHtml } from '../utils/text.js';
export {
  emitMarkdownFromBlocks,
  emitHtmlFromBlocks,
  emitMarkdownTable,
  blocksFromSpreadsheet,
  blocksFromCsvRows,
  flattenBlocksToParagraphs,
} from './blocks.js';

/** Emit markdown / html / text / csv from already-extracted content. */
export function emitText(text: string): string {
  return text.replace(/\r\n/g, '\n').trim();
}

export function emitMarkdownFromParagraphs(paragraphs: string[]): string {
  return paragraphs.map((p) => p.trim()).filter(Boolean).join('\n\n');
}

export function emitHtmlFromParagraphs(paragraphs: string[]): string {
  const body = paragraphs
    .map((p) => p.trim())
    .filter(Boolean)
    .map((p) => `<p>${escapeHtml(p)}</p>`)
    .join('\n');
  return `<!DOCTYPE html>\n<html><body>\n${body}\n</body></html>\n`;
}

export function emitMarkdownFromText(text: string): string {
  return emitText(text);
}

export function emitHtmlFromText(text: string): string {
  const paragraphs = text
    .split(/\n\s*\n/)
    .map((p) => p.trim())
    .filter(Boolean);
  if (!paragraphs.length) {
    return `<!DOCTYPE html>\n<html><body><pre>${escapeHtml(text)}</pre></body></html>\n`;
  }
  return emitHtmlFromParagraphs(paragraphs);
}

export function emitCsvFromSheet(
  structured: SpreadsheetStructured,
  sheetName?: string
): string {
  const name = sheetName ?? Object.keys(structured.sheets)[0];
  const rows = structured.sheets[name] ?? [];
  return rows
    .map((row) => row.map((c) => csvEscapeCell(c)).join(','))
    .join('\n');
}

export function emitMarkdownFromSpreadsheet(structured: SpreadsheetStructured): string {
  return emitMarkdownFromBlocks(blocksFromSpreadsheet(structured.sheets));
}

export function emitMarkdownFromPptx(slides: string[][]): string {
  const blocks: ContentBlock[] = [];
  slides.forEach((texts, i) => {
    blocks.push({
      type: 'heading',
      level: 2,
      runs: [{ text: `Slide ${i + 1}` }],
    });
    texts.forEach((t, j) => {
      if (j === 0) {
        blocks.push({ type: 'heading', level: 3, runs: [{ text: t }] });
      } else {
        blocks.push({ type: 'paragraph', runs: [{ text: t }] });
      }
    });
  });
  return emitMarkdownFromBlocks(blocks);
}

export function emitHtmlFromPptx(slides: string[][]): string {
  const blocks: ContentBlock[] = [];
  slides.forEach((texts, i) => {
    blocks.push({
      type: 'heading',
      level: 2,
      runs: [{ text: `Slide ${i + 1}` }],
    });
    for (const t of texts) {
      blocks.push({ type: 'paragraph', runs: [{ text: t }] });
    }
  });
  return emitHtmlFromBlocks(blocks);
}

export function emitTextFromSpreadsheet(structured: SpreadsheetStructured): string {
  return spreadsheetToText(structured);
}

export function chunkText(
  text: string,
  blocks?: ContentBlock[],
  maxChars = 1200,
  overlap = 120
): string[] {
  maxChars = Math.max(1, Math.floor(maxChars));
  overlap = Math.min(Math.max(0, Math.floor(overlap)), maxChars - 1);
  const trimmed = text.trim();
  if (!trimmed) return [];

  if (blocks?.length) {
    const segments: string[] = [];
    for (const b of blocks) {
      if (b.type === 'heading' || b.type === 'paragraph' || b.type === 'list_item') {
        const t = b.runs.map((r) => r.text).join('').trim();
        if (t) segments.push(t);
      } else if (b.type === 'table') {
        segments.push(
          b.rows.map((row) => row.map((c) => String(c ?? '').trim()).join(' | ')).join('\n')
        );
      } else if (b.type === 'code') {
        segments.push(b.text);
      }
    }
    const joined = segments.join('\n\n');
    return chunkText(joined, undefined, maxChars, overlap);
  }

  if (trimmed.length <= maxChars) return [trimmed];

  const chunks: string[] = [];
  const paragraphs = trimmed.split(/\n\s*\n/);
  let buf = '';
  for (const p of paragraphs) {
    const piece = p.trim();
    if (!piece) continue;
    if ((buf + '\n\n' + piece).trim().length > maxChars) {
      if (buf) chunks.push(buf.trim());
      if (piece.length <= maxChars) {
        buf = piece;
      } else {
        let i = 0;
        while (i < piece.length) {
          chunks.push(piece.slice(i, i + maxChars));
          i += maxChars - overlap;
        }
        buf = '';
      }
    } else {
      buf = buf ? `${buf}\n\n${piece}` : piece;
    }
  }
  if (buf.trim()) chunks.push(buf.trim());
  return chunks;
}
