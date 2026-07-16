import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { DocumentError, wrapError } from '../utils/errors.js';
import type { PptxStructured } from '../types.js';

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

/** Extract visible text runs from a PPTX slide DrawingML fragment. */
function extractSlideTexts(xml: string): string[] {
  const texts: string[] = [];
  const paraRegex = /<a:p[\s\S]*?<\/a:p>/g;
  let pm: RegExpExecArray | null;
  while ((pm = paraRegex.exec(xml)) !== null) {
    const parts: string[] = [];
    const tRegex = /<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g;
    let tm: RegExpExecArray | null;
    while ((tm = tRegex.exec(pm[0])) !== null) {
      parts.push(decodeXmlEntities(tm[1]));
    }
    const line = parts.join('').trim();
    if (line) texts.push(line);
  }
  // Fallback: any a:t outside well-formed a:p (rare)
  if (texts.length === 0) {
    const tRegex = /<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g;
    let tm: RegExpExecArray | null;
    while ((tm = tRegex.exec(xml)) !== null) {
      const line = decodeXmlEntities(tm[1]).trim();
      if (line) texts.push(line);
    }
  }
  return texts;
}

function listSlidePaths(zip: JSZip): string[] {
  return Object.keys(zip.files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/i.test(n) && !zip.files[n].dir)
    .sort((a, b) => {
      const na = Number(/slide(\d+)/i.exec(a)?.[1] ?? 0);
      const nb = Number(/slide(\d+)/i.exec(b)?.[1] ?? 0);
      return na - nb;
    });
}

async function readPptxFromZip(zip: JSZip): Promise<PptxStructured> {
  if (!zip.file('ppt/presentation.xml') && listSlidePaths(zip).length === 0) {
    throw new DocumentError(
      'PARSE_FAILED',
      'Invalid pptx: missing ppt/presentation.xml / slides'
    );
  }

  const slides: string[][] = [];
  for (const path of listSlidePaths(zip)) {
    const file = zip.file(path);
    if (!file) continue;
    const xml = await file.async('string');
    slides.push(extractSlideTexts(xml));
  }

  return { format: 'pptx', slides };
}

export async function readPptxStructured(filePath: string): Promise<PptxStructured> {
  try {
    const buf = await readFile(filePath);
    return readPptxStructuredFromBuffer(buf);
  } catch (err) {
    if (err instanceof DocumentError) throw err;
    throw wrapError('PARSE_FAILED', `Failed to parse pptx ${filePath}`, err);
  }
}

export async function readPptxStructuredFromBuffer(
  buf: Buffer
): Promise<PptxStructured> {
  try {
    const zip = await JSZip.loadAsync(buf);
    return readPptxFromZip(zip);
  } catch (err) {
    if (err instanceof DocumentError) throw err;
    throw wrapError('PARSE_FAILED', 'Failed to parse pptx buffer', err);
  }
}

/** Flatten slides to paragraphs (for text / convert emit). */
export function pptxToParagraphs(structured: PptxStructured): string[] {
  const out: string[] = [];
  structured.slides.forEach((texts, i) => {
    if (texts.length === 0) {
      out.push(`Slide ${i + 1}`);
      return;
    }
    for (const t of texts) out.push(t);
  });
  return out;
}

export function pptxToText(structured: PptxStructured): string {
  return structured.slides
    .map((texts, i) => {
      const body = texts.join('\n');
      return body ? `## Slide ${i + 1}\n${body}` : `## Slide ${i + 1}`;
    })
    .join('\n\n')
    .trim();
}
