import JSZip from 'jszip';
import { DocumentError, wrapError } from '../utils/errors.js';
import { decodeXmlEntities } from '../utils/xml.js';
import { loadZipBuffer } from '../utils/zip.js';
import type { CanonicalFlowBlock, CanonicalInline, ParsedCanonical } from '../ir/types.js';
import { buildParsedCanonical } from '../ir/projections.js';

export type OdfKind = 'odt' | 'ods' | 'odp';

function stripTags(xml: string): string {
  return decodeXmlEntities(xml.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim());
}

function parseInlineOdf(text: string): CanonicalInline[] {
  const runs: CanonicalInline[] = [];
  const re = /<text:span\b[^>]*>([\s\S]*?)<\/text:span>|([^<]+)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    const span = m[1];
    const plain = m[2];
    if (span != null) {
      const bold = /font-weight="bold"/i.test(m[0]) || /fo:font-weight="bold"/i.test(m[0]);
      const italic = /font-style="italic"/i.test(m[0]) || /fo:font-style="italic"/i.test(m[0]);
      const t = stripTags(span);
      if (t) runs.push({ text: t, ...(bold ? { bold: true } : {}), ...(italic ? { italic: true } : {}) });
    } else if (plain) {
      const t = decodeXmlEntities(plain);
      if (t.trim()) runs.push({ text: t });
    }
  }
  if (!runs.length) {
    const t = stripTags(text);
    if (t) runs.push({ text: t });
  }
  return runs;
}

function parseOdtTable(tableXml: string): CanonicalFlowBlock | null {
  const rows: import('../ir/types.js').CanonicalTableCell[][] = [];
  const trRe = /<table:table-row\b[\s\S]*?<\/table:table-row>/g;
  let tr: RegExpExecArray | null;
  while ((tr = trRe.exec(tableXml)) !== null) {
    const cells: import('../ir/types.js').CanonicalTableCell[] = [];
    const cellRe = /<table:(?:table-)?cell\b[\s\S]*?<\/table:(?:table-)?cell>/g;
    let c: RegExpExecArray | null;
    while ((c = cellRe.exec(tr[0])) !== null) {
      cells.push({ text: stripTags(c[0]) });
    }
    if (cells.length) rows.push(cells);
  }
  return rows.length ? { type: 'table', headerRows: 1, rows } : null;
}

function parseOdtBody(xml: string): CanonicalFlowBlock[] {
  const blocks: CanonicalFlowBlock[] = [];
  // Single document-order scan: a matched <table:table> consumes its inner
  // <text:p> cells so paragraph content is not duplicated outside the table.
  const blockRe =
    /<table:table\b[\s\S]*?<\/table:table>|<text:h\b[^>]*>([\s\S]*?)<\/text:h>|<text:p\b[^>]*>([\s\S]*?)<\/text:p>/g;
  let m: RegExpExecArray | null;
  while ((m = blockRe.exec(xml)) !== null) {
    if (m[0].startsWith('<table:table')) {
      const table = parseOdtTable(m[0]);
      if (table) blocks.push(table);
    } else if (m[0].startsWith('<text:h')) {
      const levelAttr = /text:outline-level="(\d+)"/.exec(m[0])?.[1];
      const level = Math.min(6, Math.max(1, Number(levelAttr ?? 1))) as 1 | 2 | 3 | 4 | 5 | 6;
      blocks.push({ type: 'heading', level, runs: parseInlineOdf(m[1]) });
    } else {
      const runs = parseInlineOdf(m[2]);
      if (runsToText(runs)) blocks.push({ type: 'paragraph', runs });
    }
  }
  return blocks;
}

function runsToText(runs: CanonicalInline[]): string {
  return runs.map((r) => r.text).join('').trim();
}

const MAX_ODS_REPEAT = 1024;

function repeatCount(tag: string, attr: string): number {
  const raw = new RegExp(`${attr}="(\\d+)"`).exec(tag)?.[1];
  const n = raw ? Number(raw) : 1;
  return Math.min(Math.max(1, n), MAX_ODS_REPEAT);
}

function parseOdsCellValue(cellXml: string): string | number | boolean | null {
  const openTag = /^<[^>]*>/.exec(cellXml)?.[0] ?? '';
  const typed = /office:value-type="([^"]*)"/.exec(openTag)?.[1];
  if (typed === 'float' || typed === 'currency' || typed === 'percentage') {
    const v = /office:value="([^"]*)"/.exec(openTag)?.[1];
    if (v != null && !Number.isNaN(Number(v))) return Number(v);
  }
  if (typed === 'boolean') {
    const v = /office:boolean-value="([^"]*)"/.exec(openTag)?.[1];
    if (v != null) return v === 'true';
  }
  const val = stripTags(cellXml);
  if (val === '') return null;
  const n = Number(val);
  return !Number.isNaN(n) && String(n) === val ? n : val;
}

function parseOdsBody(xml: string): import('../ir/types.js').CanonicalWorkbook {
  const sheets: import('../ir/types.js').CanonicalSheet[] = [];
  const sheetRe = /<table:table\b[^>]*table:name="([^"]*)"[\s\S]*?<\/table:table>/g;
  let sm: RegExpExecArray | null;
  while ((sm = sheetRe.exec(xml)) !== null) {
    const rows: (string | number | boolean | null)[][] = [];
    const trRe = /<table:table-row\b[^>]*\/>|<table:table-row\b[\s\S]*?<\/table:table-row>/g;
    let tr: RegExpExecArray | null;
    while ((tr = trRe.exec(sm[0])) !== null) {
      const rowRepeat = repeatCount(tr[0], 'table:number-rows-repeated');
      const cells: (string | number | boolean | null)[] = [];
      const cellRe =
        /<table:(?:table-|covered-table-)?cell\b[^>]*\/>|<table:(?:table-|covered-table-)?cell\b[\s\S]*?<\/table:(?:table-)?cell>/g;
      let c: RegExpExecArray | null;
      while ((c = cellRe.exec(tr[0])) !== null) {
        const colRepeat = repeatCount(c[0], 'table:number-columns-repeated');
        const value = parseOdsCellValue(c[0]);
        for (let k = 0; k < colRepeat; k++) cells.push(value);
      }
      // Trailing repeated empties pad rows to the sheet max width; drop them.
      while (cells.length && cells[cells.length - 1] === null) cells.pop();
      if (cells.length) {
        for (let k = 0; k < rowRepeat; k++) rows.push([...cells]);
      }
    }
    sheets.push({ name: sm[1], rows });
  }
  return { kind: 'workbook', format: 'ods', sheets };
}

function parseOdpBody(xml: string): import('../ir/types.js').CanonicalPresentation {
  const slides: import('../ir/types.js').CanonicalSlide[] = [];
  const pageRe = /<draw:page\b[^>]*>([\s\S]*?)<\/draw:page>/g;
  let i = 0;
  let pm: RegExpExecArray | null;
  while ((pm = pageRe.exec(xml)) !== null) {
    const body = parseOdtBody(pm[1]);
    const titleBlock = body.find((b) => b.type === 'heading' || b.type === 'paragraph');
    slides.push({
      index: i,
      title:
        titleBlock && (titleBlock.type === 'heading' || titleBlock.type === 'paragraph')
          ? titleBlock.runs
          : undefined,
      body: body.filter((b) => b !== titleBlock),
    });
    i++;
  }
  return { kind: 'presentation', format: 'odp', slides };
}

export async function parseOdfBuffer(buf: Buffer, kind: OdfKind): Promise<ParsedCanonical> {
  try {
    const zip = await loadZipBuffer(buf);
    const content = zip.file('content.xml');
    if (!content) throw new DocumentError('PARSE_FAILED', 'Invalid ODF: missing content.xml');
    const xml = await content.async('string');
    if (kind === 'ods') {
      const workbook = parseOdsBody(xml);
      return buildParsedCanonical(workbook, [
        { code: 'ODF_BASIC_PARSE', message: 'ODS parsed with basic ODF table extraction.' },
      ]);
    }
    if (kind === 'odp') {
      const presentation = parseOdpBody(xml);
      return buildParsedCanonical(presentation, [
        { code: 'ODF_BASIC_PARSE', message: 'ODP parsed with basic slide text extraction.' },
      ]);
    }
    const blocks = parseOdtBody(xml);
    return buildParsedCanonical(
      { kind: 'flow', format: 'odt', blocks },
      [{ code: 'ODF_BASIC_PARSE', message: 'ODT parsed with basic ODF text/table extraction.' }]
    );
  } catch (err) {
    if (err instanceof DocumentError) throw err;
    throw wrapError('PARSE_FAILED', `Failed to parse ${kind}`, err);
  }
}

export async function detectOdfKind(buf: Buffer): Promise<OdfKind | null> {
  try {
    const zip = await loadZipBuffer(buf);
    const mimeFile = zip.file('mimetype');
    if (mimeFile) {
      const mime = (await mimeFile.async('string')).trim();
      if (mime.includes('spreadsheet')) return 'ods';
      if (mime.includes('presentation')) return 'odp';
      if (mime.includes('text')) return 'odt';
    }
    const content = zip.file('content.xml');
    if (!content) return null;
    const xml = await content.async('string');
    if (/<office:spreadsheet/i.test(xml)) return 'ods';
    if (/<office:presentation/i.test(xml)) return 'odp';
    if (/<office:text/i.test(xml)) return 'odt';
  } catch {
    return null;
  }
  return null;
}
