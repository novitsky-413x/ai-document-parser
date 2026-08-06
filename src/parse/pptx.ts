import type JSZip from 'jszip';
import { DocumentError, wrapError } from '../utils/errors.js';
import { loadZipBuffer, readFileLimited } from '../utils/zip.js';
import type { ContentBlock, InlineRun, PptxStructured } from '../types.js';
import { flattenBlocksToParagraphs } from '../convert/blocks.js';

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function parsePptxRuns(paraXml: string): InlineRun[] {
  const runs: InlineRun[] = [];
  const runRe = /<a:r\b[^>]*>([\s\S]*?)<\/a:r>/g;
  let rm: RegExpExecArray | null;
  while ((rm = runRe.exec(paraXml)) !== null) {
    const inner = rm[1];
    const bold = /b="1"/.test(inner) || /<a:b\b/.test(inner);
    const italic = /i="1"/.test(inner) || /<a:i\b/.test(inner);
    const tm = /<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/.exec(inner);
    if (!tm) continue;
    const text = decodeXmlEntities(tm[1]);
    if (!text) continue;
    runs.push({
      text,
      ...(bold ? { bold: true } : {}),
      ...(italic ? { italic: true } : {}),
    });
  }
  if (!runs.length) {
    const tRe = /<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g;
    let tm: RegExpExecArray | null;
    while ((tm = tRe.exec(paraXml)) !== null) {
      const text = decodeXmlEntities(tm[1]);
      if (text) runs.push({ text });
    }
  }
  return runs;
}

function extractSlideTables(xml: string): ContentBlock[] {
  const blocks: ContentBlock[] = [];
  const tblRe = /<a:tbl\b[\s\S]*?<\/a:tbl>/g;
  let tm: RegExpExecArray | null;
  while ((tm = tblRe.exec(xml)) !== null) {
    const rows: string[][] = [];
    const trRe = /<a:tr\b[\s\S]*?<\/a:tr>/g;
    let tr: RegExpExecArray | null;
    while ((tr = trRe.exec(tm[0])) !== null) {
      const cells: string[] = [];
      const tcRe = /<a:tc\b[\s\S]*?<\/a:tc>/g;
      let tc: RegExpExecArray | null;
      while ((tc = tcRe.exec(tr[0])) !== null) {
        const texts: string[] = [];
        const tRe = /<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g;
        let t: RegExpExecArray | null;
        while ((t = tRe.exec(tc[0])) !== null) {
          texts.push(decodeXmlEntities(t[1]));
        }
        cells.push(texts.join('').trim());
      }
      if (cells.length) rows.push(cells);
    }
    if (rows.length) blocks.push({ type: 'table', rows });
  }
  return blocks;
}

/** Detect DrawingML bullet / numbering on a paragraph (`buChar` / `buAutoNum`). */
function detectPptxListMeta(paraXml: string): {
  isList: boolean;
  ordered: boolean;
  level: number;
} | null {
  const pPr = /<a:pPr\b([^>]*)(?:\/>|>([\s\S]*?)<\/a:pPr>)/.exec(paraXml);
  if (!pPr) return null;
  const pPrInner = `${pPr[1] ?? ''}${pPr[2] ?? ''}`;
  if (/<a:buNone\b/.test(pPrInner)) return null;

  const ordered = /<a:buAutoNum\b/.test(pPrInner);
  const bullet = /<a:buChar\b/.test(pPrInner);
  if (!ordered && !bullet) return null;

  let level = 0;
  const lvlAttr = /\blvl="(\d+)"/.exec(pPrInner);
  if (lvlAttr) {
    level = Math.max(0, Number(lvlAttr[1]) || 0);
  } else {
    const marL = /\bmarL="(\d+)"/.exec(pPrInner);
    if (marL) {
      const emu = Number(marL[1]) || 0;
      // Writer uses marL="342900" per indent step.
      level = Math.max(0, Math.round(emu / 342900) - 1);
    }
  }
  return { isList: true, ordered, level };
}

function extractSlideBlocks(xml: string, slideIndex: number): {
  texts: string[];
  blocks: ContentBlock[];
} {
  const texts: string[] = [];
  const blocks: ContentBlock[] = [
    {
      type: 'heading',
      level: 2,
      runs: [{ text: `Slide ${slideIndex + 1}` }],
    },
  ];

  // Remove tables from paragraph scan to avoid duplicating cell text as paras
  const withoutTables = xml.replace(/<a:tbl\b[\s\S]*?<\/a:tbl>/g, '');

  const paraRegex = /<a:p[\s\S]*?<\/a:p>/g;
  let pm: RegExpExecArray | null;
  let paraIdx = 0;
  while ((pm = paraRegex.exec(withoutTables)) !== null) {
    const paraXml = pm[0];
    const runs = parsePptxRuns(paraXml);
    const line = runs.map((r) => r.text).join('').trim();
    if (!line) continue;
    texts.push(line);
    const listMeta = detectPptxListMeta(paraXml);
    const runPayload = runs.length ? runs : [{ text: line }];
    if (listMeta?.isList) {
      blocks.push({
        type: 'list_item',
        level: listMeta.level,
        ordered: listMeta.ordered || undefined,
        runs: runPayload,
      });
    } else if (paraIdx === 0) {
      blocks.push({ type: 'heading', level: 3, runs: runPayload });
    } else {
      blocks.push({ type: 'paragraph', runs: runPayload });
    }
    paraIdx++;
  }

  blocks.push(...extractSlideTables(xml));
  return { texts, blocks };
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
  const blocks: ContentBlock[] = [];
  let i = 0;
  for (const path of listSlidePaths(zip)) {
    const file = zip.file(path);
    if (!file) continue;
    const xml = await file.async('string');
    const { texts, blocks: slideBlocks } = extractSlideBlocks(xml, i);
    slides.push(texts);
    blocks.push(...slideBlocks);
    i++;
  }

  return { format: 'pptx', slides, blocks };
}

export async function readPptxStructured(filePath: string): Promise<PptxStructured> {
  try {
    const buf = await readFileLimited(filePath);
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
    const zip = await loadZipBuffer(buf);
    return readPptxFromZip(zip);
  } catch (err) {
    if (err instanceof DocumentError) throw err;
    throw wrapError('PARSE_FAILED', 'Failed to parse pptx buffer', err);
  }
}

export function pptxToParagraphs(structured: PptxStructured): string[] {
  if (structured.blocks?.length) {
    return flattenBlocksToParagraphs(structured.blocks);
  }
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
