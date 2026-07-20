import { readFile } from 'node:fs/promises';
import { DocumentError, wrapError } from '../utils/errors.js';
import { assertMaxPdfPages } from '../utils/limits.js';
import { readFileLimited } from '../utils/zip.js';

/**
 * Extract plain text from a PDF using pdfjs-dist (public PDF.js API).
 */
export async function extractPdfText(
  input: string | Buffer
): Promise<{ text: string; format: 'pdf' }> {
  let pdf: { numPages: number; getPage: (n: number) => Promise<{ getTextContent: () => Promise<{ items: unknown[] }> }>; destroy?: () => void } | null = null;
  try {
    const data =
      typeof input === 'string'
        ? new Uint8Array(await readFileLimited(input))
        : new Uint8Array(input);

    const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
    const loadingTask = pdfjs.getDocument({
      data,
      useSystemFonts: true,
    });
    pdf = await loadingTask.promise;
    assertMaxPdfPages(pdf.numPages);
    const parts: string[] = [];

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      const page = await pdf.getPage(pageNum);
      const content = await page.getTextContent();
      const items = content.items as Array<{ str?: string; transform?: number[] }>;
      const lines = groupPdfItems(items);
      if (lines.length) parts.push(lines.join('\n'));
    }

    return { text: parts.join('\n\n').trim(), format: 'pdf' };
  } catch (err) {
    if (err instanceof DocumentError) throw err;
    throw wrapError('PARSE_FAILED', 'Failed to parse PDF', err);
  } finally {
    pdf?.destroy?.();
  }
}

function groupPdfItems(items: Array<{ str?: string; transform?: number[] }>): string[] {
  const rows: Array<{ y: number; texts: Array<{ x: number; t: string }> }> = [];
  for (const item of items) {
    const t = typeof item.str === 'string' ? item.str : '';
    if (!t.trim()) continue;
    const x = item.transform?.[4] ?? 0;
    const y = item.transform?.[5] ?? 0;
    const bucket = rows.find((r) => Math.abs(r.y - y) < 4);
    if (bucket) bucket.texts.push({ x, t });
    else rows.push({ y, texts: [{ x, t }] });
  }
  rows.sort((a, b) => b.y - a.y);
  return rows.map((r) =>
    r.texts
      .sort((a, b) => a.x - b.x)
      .map((p) => p.t)
      .join(' ')
      .trim()
  ).filter(Boolean);
}
