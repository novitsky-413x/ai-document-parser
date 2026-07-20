import JSZip from 'jszip';
import { DocumentError, wrapError } from '../utils/errors.js';
import { decodeXmlEntities } from '../utils/xml.js';
import { loadZipBuffer } from '../utils/zip.js';
import { parseCodeStyleId } from '../utils/code-style.js';
import type { CanonicalFlowBlock, CanonicalInline, ParsedCanonical } from '../ir/types.js';
import { buildParsedCanonical } from '../ir/projections.js';

export type OdfKind = 'odt' | 'ods' | 'odp';

function stripTags(xml: string): string {
  return decodeXmlEntities(xml.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim());
}

function extractOdfCodeText(inner: string): string {
  return decodeXmlEntities(
    inner
      .replace(/<text:line-break\b[^/]*\/>/g, '\n')
      .replace(/<text:line-break\b[^>]*>\s*<\/text:line-break>/g, '\n')
      .replace(/<text:s\b[^>]*\/>/g, ' ')
      .replace(/<text:tab\b[^>]*\/>/g, '\t')
      .replace(/<[^>]+>/g, '')
  );
}

function parseInlineOdf(text: string): CanonicalInline[] {
  const runs: CanonicalInline[] = [];
  const re =
    /<text:a\b[^>]*xlink:href="([^"]*)"[^>]*>([\s\S]*?)<\/text:a>|<text:span\b[^>]*>([\s\S]*?)<\/text:span>|([^<]+)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    if (m[1] != null && m[2] != null) {
      const innerRuns = parseInlineOdf(m[2]);
      const href = decodeXmlEntities(m[1]);
      for (const r of innerRuns.length ? innerRuns : [{ text: stripTags(m[2]) }]) {
        if (r.text) runs.push({ ...r, href });
      }
      continue;
    }
    const span = m[3];
    const plain = m[4];
    if (span != null) {
      const style = /text:style-name="([^"]+)"/.exec(m[0])?.[1] ?? '';
      const bold =
        /font-weight="bold"/i.test(m[0]) ||
        /fo:font-weight="bold"/i.test(m[0]) ||
        style === 'ADPB' ||
        style === 'ADPBI';
      const italic =
        /font-style="italic"/i.test(m[0]) ||
        /fo:font-style="italic"/i.test(m[0]) ||
        style === 'ADPI' ||
        style === 'ADPBI';
      const strike =
        /line-through/i.test(m[0]) || style === 'ADPS';
      const code = style === 'ADPCodeInline' || /Courier|Consolas|Mono/i.test(m[0]);
      const t = stripTags(span);
      if (t) {
        runs.push({
          text: t,
          ...(bold ? { bold: true } : {}),
          ...(italic ? { italic: true } : {}),
          ...(strike ? { strike: true } : {}),
          ...(code ? { code: true } : {}),
        });
      }
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

/** Balanced end for an ODF element whose local name is `local` (e.g. list, table). */
function findBalancedEnd(xml: string, start: number, local: string): number {
  // (?=[\s/>]) so "list" does not match "list-item".
  const openRe = new RegExp(`<(?:[A-Za-z0-9_]+:)?${local}(?=[\\s/>])`);
  const closeRe = new RegExp(`</(?:[A-Za-z0-9_]+:)?${local}\\s*>`, 'g');
  let depth = 0;
  let i = start;
  while (i < xml.length) {
    const slice = xml.slice(i);
    const open = openRe.exec(slice);
    closeRe.lastIndex = 0;
    const close = closeRe.exec(slice);
    const openAt = open ? open.index : -1;
    const closeAt = close ? close.index : -1;
    if (openAt < 0 && closeAt < 0) return -1;
    if (openAt >= 0 && (closeAt < 0 || openAt < closeAt)) {
      const abs = i + openAt;
      const gt = xml.indexOf('>', abs);
      if (gt < 0) return -1;
      if (xml.charAt(gt - 1) !== '/') depth++;
      i = gt + 1;
    } else {
      depth--;
      i = i + closeAt + (close?.[0].length ?? 0);
      if (depth === 0) return i;
    }
  }
  return -1;
}

function parseListStyleOrdered(xml: string): Map<string, boolean> {
  const map = new Map<string, boolean>();
  const re = /<text:list-style\b[^>]*style:name="([^"]+)"[\s\S]*?<\/text:list-style>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) {
    const name = m[1];
    if (/list-level-style-number\b/i.test(m[0])) map.set(name, true);
    else if (/list-level-style-bullet\b/i.test(m[0])) map.set(name, false);
  }
  return map;
}

function isOrderedListStyle(
  styleName: string | undefined,
  styleMap: Map<string, boolean>
): boolean | undefined {
  if (!styleName) return undefined;
  if (styleMap.has(styleName)) return styleMap.get(styleName);
  if (/ADPNum/i.test(styleName) || /num/i.test(styleName)) return true;
  if (/ADPBullet/i.test(styleName) || /bullet/i.test(styleName)) return false;
  return undefined;
}

function parseOdtList(
  listXml: string,
  styleMap: Map<string, boolean>
): CanonicalFlowBlock | null {
  const open = /^<text:list\b[^>]*>/.exec(listXml)?.[0] ?? '';
  const style = /text:style-name="([^"]+)"/.exec(open)?.[1];
  const ordered = isOrderedListStyle(style, styleMap);
  const items: CanonicalFlowBlock[][] = [];

  // Walk top-level list-items only (balanced), so nested lists stay inside their item.
  let i = 0;
  const inner = listXml.replace(/^<text:list\b[^>]*>/, '').replace(/<\/text:list>\s*$/, '');
  while (i < inner.length) {
    const at = inner.indexOf('<text:list-item', i);
    if (at < 0) break;
    const after = inner.charAt(at + '<text:list-item'.length);
    if (after && !/[\s/>]/.test(after)) {
      i = at + 1;
      continue;
    }
    const end = findBalancedEnd(inner, at, 'list-item');
    if (end < 0) {
      i = at + 1;
      continue;
    }
    const itemXml = inner.slice(at, end);
    const itemBlocks = parseOdtListItemContent(itemXml, styleMap);
    if (itemBlocks.length) items.push(itemBlocks);
    i = end;
  }

  if (!items.length) return null;
  return {
    type: 'list',
    ...(ordered ? { ordered: true } : {}),
    items,
  };
}

function parseOdtListItemContent(
  itemXml: string,
  styleMap: Map<string, boolean>
): CanonicalFlowBlock[] {
  const body = itemXml
    .replace(/^<text:list-item\b[^>]*>/, '')
    .replace(/<\/text:list-item>\s*$/, '');
  return parseOdtFragment(body, styleMap);
}

function parseOdtParagraph(pXml: string): CanonicalFlowBlock | null {
  const open = /^<text:p\b[^>]*>/.exec(pXml)?.[0] ?? '';
  const style = /text:style-name="([^"]+)"/.exec(open)?.[1];
  const codeMeta = parseCodeStyleId(style);
  const inner = pXml.replace(/^<text:p\b[^>]*>/, '').replace(/<\/text:p>\s*$/, '');
  if (codeMeta.isCode) {
    return {
      type: 'code',
      text: extractOdfCodeText(inner),
      ...(codeMeta.language ? { language: codeMeta.language } : {}),
    };
  }
  const runs = parseInlineOdf(inner);
  if (!runsToText(runs)) return null;
  return { type: 'paragraph', runs };
}

function parseOdtHeading(hXml: string): CanonicalFlowBlock {
  const levelAttr = /text:outline-level="(\d+)"/.exec(hXml)?.[1];
  const level = Math.min(6, Math.max(1, Number(levelAttr ?? 1))) as 1 | 2 | 3 | 4 | 5 | 6;
  const inner = hXml.replace(/^<text:h\b[^>]*>/, '').replace(/<\/text:h>\s*$/, '');
  return { type: 'heading', level, runs: parseInlineOdf(inner) };
}

/** Parse a fragment of office:text children in document order. */
function parseOdtFragment(
  xml: string,
  styleMap: Map<string, boolean>
): CanonicalFlowBlock[] {
  const blocks: CanonicalFlowBlock[] = [];
  let i = 0;
  while (i < xml.length) {
    const tableAt = indexOfTag(xml, 'table:table', i);
    const listAt = indexOfTag(xml, 'text:list', i);
    const hAt = indexOfTag(xml, 'text:h', i);
    const pAt = indexOfTag(xml, 'text:p', i);

    const candidates = [
      { kind: 'table' as const, at: tableAt },
      { kind: 'list' as const, at: listAt },
      { kind: 'h' as const, at: hAt },
      { kind: 'p' as const, at: pAt },
    ].filter((c) => c.at >= 0);
    if (!candidates.length) break;
    candidates.sort((a, b) => a.at - b.at);
    const next = candidates[0];

    if (next.kind === 'table') {
      const end = findBalancedEnd(xml, next.at, 'table');
      if (end < 0) {
        i = next.at + 1;
        continue;
      }
      const table = parseOdtTable(xml.slice(next.at, end));
      if (table) blocks.push(table);
      i = end;
    } else if (next.kind === 'list') {
      const end = findBalancedEnd(xml, next.at, 'list');
      if (end < 0) {
        i = next.at + 1;
        continue;
      }
      const list = parseOdtList(xml.slice(next.at, end), styleMap);
      if (list) blocks.push(list);
      i = end;
    } else if (next.kind === 'h') {
      const end = findBalancedEnd(xml, next.at, 'h');
      if (end < 0) {
        i = next.at + 1;
        continue;
      }
      blocks.push(parseOdtHeading(xml.slice(next.at, end)));
      i = end;
    } else {
      const end = findBalancedEnd(xml, next.at, 'p');
      if (end < 0) {
        i = next.at + 1;
        continue;
      }
      const para = parseOdtParagraph(xml.slice(next.at, end));
      if (para) blocks.push(para);
      i = end;
    }
  }
  return blocks;
}

function indexOfTag(xml: string, qname: string, from: number): number {
  let i = from;
  const needle = `<${qname}`;
  while (i < xml.length) {
    const at = xml.indexOf(needle, i);
    if (at < 0) return -1;
    const after = xml.charAt(at + needle.length);
    if (!after || /[\s/>]/.test(after)) return at;
    i = at + needle.length;
  }
  return -1;
}

function parseOdtBody(xml: string): CanonicalFlowBlock[] {
  const styleMap = parseListStyleOrdered(xml);
  // Prefer office:text body; fall back to whole content.xml.
  const body =
    /<office:text\b[^>]*>([\s\S]*)<\/office:text>/i.exec(xml)?.[1] ?? xml;
  return parseOdtFragment(body, styleMap);
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
