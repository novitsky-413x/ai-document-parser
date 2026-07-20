import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { jsPDF } from 'jspdf';
import { ensureParentDir, writeFileAtomic } from '../utils/paths.js';
import { wrapError } from '../utils/errors.js';

const PDF_FONT = 'DejaVuSans';
const PDF_FONT_FILE = 'DejaVuSans.ttf';

function resolveFontPath(): string {
  // Prefer packaged fonts next to dist/, then repo fonts/
  const here = path.dirname(fileURLToPath(import.meta.url));
  const candidates = [
    path.join(here, '..', '..', 'fonts', PDF_FONT_FILE),
    path.join(here, '..', 'fonts', PDF_FONT_FILE),
    path.join(process.cwd(), 'fonts', PDF_FONT_FILE),
  ];
  for (const c of candidates) {
    try {
      readFileSync(c);
      return c;
    } catch {
      /* try next */
    }
  }
  throw new Error(
    `Font ${PDF_FONT_FILE} not found. Expected under package fonts/`
  );
}

function registerCyrillicFont(doc: jsPDF): void {
  const fontPath = resolveFontPath();
  const fontData = readFileSync(fontPath).toString('binary');
  doc.addFileToVFS(PDF_FONT_FILE, fontData);
  doc.addFont(PDF_FONT_FILE, PDF_FONT, 'normal');
  doc.setFont(PDF_FONT);
}

export async function writePdf(
  filePath: string,
  content: string,
  title?: string
): Promise<{ message: string }> {
  try {
    const doc = new jsPDF({ unit: 'mm', format: 'a4' });
    registerCyrillicFont(doc);

    const margin = 15;
    const pageWidth = doc.internal.pageSize.getWidth() - margin * 2;
    let y = margin;

    if (title) {
      doc.setFontSize(14);
      const titleLines = doc.splitTextToSize(title, pageWidth) as string[];
      doc.text(titleLines, margin, y);
      y += titleLines.length * 7 + 4;
    }

    doc.setFontSize(11);
    const lines = doc.splitTextToSize(content, pageWidth) as string[];
    const lineHeight = 5;

    for (const line of lines) {
      if (y > doc.internal.pageSize.getHeight() - margin) {
        doc.addPage();
        doc.setFont(PDF_FONT);
        y = margin;
      }
      doc.text(line, margin, y);
      y += lineHeight;
    }

    const buffer = Buffer.from(doc.output('arraybuffer'));
    await ensureParentDir(filePath);
    await writeFileAtomic(filePath, buffer);
    return { message: `Created PDF: ${filePath} (${lines.length} line(s))` };
  } catch (err) {
    throw wrapError('WRITE_FAILED', `Failed to write pdf ${filePath}`, err);
  }
}
