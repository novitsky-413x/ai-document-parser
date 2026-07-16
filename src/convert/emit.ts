import { csvEscapeCell, escapeHtml } from '../utils/text.js';
import type { SpreadsheetStructured } from '../types.js';
import { spreadsheetToText } from '../edit/spreadsheet.js';

export { csvEscapeCell, escapeHtml } from '../utils/text.js';

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
  const parts: string[] = [];
  for (const [name, rows] of Object.entries(structured.sheets)) {
    parts.push(`## ${name}`);
    if (!rows.length) continue;
    const header = rows[0].map((c) => String(c ?? ''));
    const sep = header.map(() => '---');
    parts.push(`| ${header.join(' | ')} |`);
    parts.push(`| ${sep.join(' | ')} |`);
    for (const row of rows.slice(1)) {
      parts.push(`| ${row.map((c) => String(c ?? '')).join(' | ')} |`);
    }
    parts.push('');
  }
  return parts.join('\n').trim();
}

export function emitMarkdownFromPptx(slides: string[][]): string {
  return slides
    .map((texts, i) => {
      const body = texts.map((t) => t.trim()).filter(Boolean).join('\n\n');
      return body ? `## Slide ${i + 1}\n\n${body}` : `## Slide ${i + 1}`;
    })
    .join('\n\n')
    .trim();
}

export function emitHtmlFromPptx(slides: string[][]): string {
  const sections = slides
    .map((texts, i) => {
      const paras = texts
        .map((t) => t.trim())
        .filter(Boolean)
        .map((t) => `<p>${escapeHtml(t)}</p>`)
        .join('\n');
      return `<section>\n<h2>Slide ${i + 1}</h2>\n${paras || '<p></p>'}\n</section>`;
    })
    .join('\n');
  return `<!DOCTYPE html>\n<html><body>\n${sections}\n</body></html>\n`;
}

export function emitTextFromSpreadsheet(structured: SpreadsheetStructured): string {
  return spreadsheetToText(structured);
}

export function chunkText(text: string, maxChars = 1200): string[] {
  const trimmed = text.trim();
  if (!trimmed) return [];
  if (trimmed.length <= maxChars) return [trimmed];
  const chunks: string[] = [];
  let i = 0;
  while (i < trimmed.length) {
    chunks.push(trimmed.slice(i, i + maxChars));
    i += maxChars;
  }
  return chunks;
}
