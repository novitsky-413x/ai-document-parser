import { readFile } from 'node:fs/promises';
import { DocumentError, wrapError } from '../utils/errors.js';

/**
 * Extract plain text from a PDF using pdfjs-dist (public PDF.js API).
 */
export async function extractPdfText(
  input: string | Buffer
): Promise<{ text: string; format: 'pdf' }> {
  try {
    const data =
      typeof input === 'string' ? new Uint8Array(await readFile(input)) : new Uint8Array(input);

    const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
    const loadingTask = pdfjs.getDocument({
      data,
      useSystemFonts: true,
    });
    const pdf = await loadingTask.promise;
    const parts: string[] = [];

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      const page = await pdf.getPage(pageNum);
      const content = await page.getTextContent();
      const line = content.items
        .map((item) => ('str' in item && typeof item.str === 'string' ? item.str : ''))
        .filter(Boolean)
        .join(' ');
      if (line.trim()) parts.push(line);
    }

    return { text: parts.join('\n').trim(), format: 'pdf' };
  } catch (err) {
    if (err instanceof DocumentError) throw err;
    throw wrapError('PARSE_FAILED', 'Failed to parse PDF', err);
  }
}
