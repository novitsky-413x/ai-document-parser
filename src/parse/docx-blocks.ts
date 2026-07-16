import { DocumentError } from '../utils/errors.js';
import type { ContentBlock, DocxStructured, InlineRun } from '../types.js';
import { flattenBlocksToParagraphs } from '../convert/blocks.js';
import type JSZip from 'jszip';

export type DocxPackageParts = {
  documentXml: string;
  stylesXml?: string;
  numberingXml?: string;
  relsXml?: string;
  footnotesXml?: string;
  endnotesXml?: string;
};

type CharStyleFlags = { bold?: boolean; italic?: boolean; strike?: boolean };

type NumberingIndex = {
  /** numId -> abstractNumId */
  numToAbstract: Map<string, string>;
  /** abstractNumId -> ilvl -> numFmt */
  abstractLevels: Map<string, Map<number, string>>;
};

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function attr(el: string, name: string): string | undefined {
  const re = new RegExp(`${name}="([^"]*)"`, 'i');
  const m = re.exec(el);
  if (m) return m[1];
  const re2 = new RegExp(`${name}='([^']*)'`, 'i');
  return re2.exec(el)?.[1];
}

function extractBodyInner(xml: string): string {
  const m = /<w:body[^>]*>([\s\S]*)<\/w:body>/i.exec(xml);
  return m ? m[1] : xml;
}

/** Find matching close for a start tag at `start`, handling nesting of the same local name. */
function findBalancedEnd(xml: string, start: number, localName: string): number {
  const openRe = new RegExp(`<w:${localName}\\b`, 'g');
  const closeRe = new RegExp(`</w:${localName}>`, 'g');
  openRe.lastIndex = start;
  const first = openRe.exec(xml);
  if (!first || first.index !== start) return -1;

  // Self-closing empty element: <w:p .../>
  const gt = xml.indexOf('>', first.index);
  if (gt < 0) return -1;
  if (xml.charAt(gt - 1) === '/') return gt + 1;

  let depth = 1;
  let i = gt + 1;
  while (depth > 0 && i < xml.length) {
    openRe.lastIndex = i;
    closeRe.lastIndex = i;
    const nextOpen = openRe.exec(xml);
    const nextClose = closeRe.exec(xml);
    if (!nextClose) return -1;
    if (nextOpen && nextOpen.index < nextClose.index) {
      const openGt = xml.indexOf('>', nextOpen.index);
      if (openGt < 0) return -1;
      // Nested self-closing does not increase depth.
      if (xml.charAt(openGt - 1) === '/') {
        i = openGt + 1;
        continue;
      }
      depth++;
      i = openGt + 1;
    } else {
      depth--;
      i = nextClose.index + nextClose[0].length;
      if (depth === 0) return i;
    }
  }
  return -1;
}

function iterBodyParts(body: string): Array<{ kind: 'p' | 'tbl'; xml: string }> {
  const parts: Array<{ kind: 'p' | 'tbl'; xml: string }> = [];
  let i = 0;
  while (i < body.length) {
    const tblAt = body.indexOf('<w:tbl', i);
    const pAt = body.indexOf('<w:p', i);
    if (tblAt < 0 && pAt < 0) break;
    let kind: 'p' | 'tbl';
    let start: number;
    if (tblAt >= 0 && (pAt < 0 || tblAt < pAt)) {
      // '<w:tbl' must be followed by space, /, or > (not tblGrid / tblPr…)
      const after = body.charAt(tblAt + 6);
      if (after && !/[\s/>]/.test(after)) {
        i = tblAt + 6;
        continue;
      }
      kind = 'tbl';
      start = tblAt;
    } else {
      const after = body.charAt(pAt + 4);
      if (after && !/[\s/>]/.test(after)) {
        i = pAt + 4;
        continue;
      }
      kind = 'p';
      start = pAt;
    }
    const end = findBalancedEnd(body, start, kind);
    if (end < 0) {
      // Skip malformed tag instead of aborting the rest of the document.
      i = start + 1;
      continue;
    }
    parts.push({ kind, xml: body.slice(start, end) });
    i = end;
  }
  return parts;
}

function parseRels(relsXml: string | undefined): Map<string, string> {
  const map = new Map<string, string>();
  if (!relsXml) return map;
  const re = /<Relationship\b[^>]*>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(relsXml)) !== null) {
    const el = m[0];
    const id = attr(el, 'Id');
    const target = attr(el, 'Target');
    if (id && target) map.set(id, target);
  }
  return map;
}

function parseCharStyles(stylesXml: string | undefined): Map<string, CharStyleFlags> {
  const map = new Map<string, CharStyleFlags>();
  if (!stylesXml) return map;

  const styleRe = /<w:style\b[^>]*>[\s\S]*?<\/w:style>/g;
  let sm: RegExpExecArray | null;
  while ((sm = styleRe.exec(stylesXml)) !== null) {
    const block = sm[0];
    const open = block.match(/<w:style\b[^>]*>/)?.[0] ?? '';
    const type = attr(open, 'w:type');
    if (type && type !== 'character' && type !== 'paragraph') {
      // still allow character + paragraph for rPr/pPr basedOn chains
    }
    const styleId =
      attr(open, 'w:styleId') ??
      /<w:name\b[^>]*w:val="([^"]+)"/.exec(block)?.[1];
    if (!styleId) continue;

    const rPr = /<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/.exec(block)?.[0] ?? '';
    const flags: CharStyleFlags = {};
    if (isOnOffEnabled(rPr, 'b') || isOnOffEnabled(rPr, 'bCs')) flags.bold = true;
    if (isOnOffEnabled(rPr, 'i') || isOnOffEnabled(rPr, 'iCs')) flags.italic = true;
    if (isOnOffEnabled(rPr, 'strike') || isOnOffEnabled(rPr, 'dstrike')) {
      flags.strike = true;
    }

    // Well-known Word character style names even if rPr empty in basedOn leaf
    if (/^Strong$/i.test(styleId)) flags.bold = true;
    if (/^(Emphasis|SubtleEmphasis|IntenseEmphasis)$/i.test(styleId)) flags.italic = true;
    if (/IntenseEmphasis/i.test(styleId)) flags.bold = true;

    if (flags.bold || flags.italic || flags.strike) map.set(styleId, flags);
  }

  // Resolve basedOn once
  const basedOn = new Map<string, string>();
  const styleRe2 = /<w:style\b[^>]*>[\s\S]*?<\/w:style>/g;
  let sm2: RegExpExecArray | null;
  while ((sm2 = styleRe2.exec(stylesXml)) !== null) {
    const block = sm2[0];
    const open = block.match(/<w:style\b[^>]*>/)?.[0] ?? '';
    const styleId = attr(open, 'w:styleId');
    const base = /<w:basedOn\b[^>]*w:val="([^"]+)"/.exec(block)?.[1];
    if (styleId && base) basedOn.set(styleId, base);
  }

  const resolve = (id: string, seen = new Set<string>()): CharStyleFlags => {
    if (seen.has(id)) return map.get(id) ?? {};
    seen.add(id);
    const own = map.get(id) ?? {};
    const parentId = basedOn.get(id);
    if (!parentId) return own;
    const parent = resolve(parentId, seen);
    return {
      bold: own.bold || parent.bold,
      italic: own.italic || parent.italic,
      strike: own.strike || parent.strike,
    };
  };

  for (const id of new Set([...map.keys(), ...basedOn.keys()])) {
    const resolved = resolve(id);
    if (resolved.bold || resolved.italic || resolved.strike) map.set(id, resolved);
  }

  return map;
}

function parseNumbering(numberingXml: string | undefined): NumberingIndex {
  const numToAbstract = new Map<string, string>();
  const abstractLevels = new Map<string, Map<number, string>>();
  if (!numberingXml) return { numToAbstract, abstractLevels };

  const absRe = /<w:abstractNum\b[^>]*>[\s\S]*?<\/w:abstractNum>/g;
  let am: RegExpExecArray | null;
  while ((am = absRe.exec(numberingXml)) !== null) {
    const block = am[0];
    const open = block.match(/<w:abstractNum\b[^>]*>/)?.[0] ?? '';
    const absId = attr(open, 'w:abstractNumId');
    if (absId == null) continue;
    const levels = new Map<number, string>();
    const lvlRe = /<w:lvl\b[^>]*>[\s\S]*?<\/w:lvl>/g;
    let lm: RegExpExecArray | null;
    while ((lm = lvlRe.exec(block)) !== null) {
      const lvlBlock = lm[0];
      const lvlOpen = lvlBlock.match(/<w:lvl\b[^>]*>/)?.[0] ?? '';
      const ilvl = Number(attr(lvlOpen, 'w:ilvl') ?? '0');
      const fmt =
        /<w:numFmt\b[^>]*w:val="([^"]+)"/.exec(lvlBlock)?.[1] ??
        /<w:numFmt\b[^>]*w:val='([^']+)'/.exec(lvlBlock)?.[1] ??
        'bullet';
      levels.set(ilvl, fmt);
    }
    abstractLevels.set(absId, levels);
  }

  const numRe = /<w:num\b[^>]*>[\s\S]*?<\/w:num>/g;
  let nm: RegExpExecArray | null;
  while ((nm = numRe.exec(numberingXml)) !== null) {
    const block = nm[0];
    const open = block.match(/<w:num\b[^>]*>/)?.[0] ?? '';
    const numId = attr(open, 'w:numId');
    const absId =
      /<w:abstractNumId\b[^>]*w:val="([^"]+)"/.exec(block)?.[1] ??
      /<w:abstractNumId\b[^>]*w:val='([^']+)'/.exec(block)?.[1];
    if (numId != null && absId != null) numToAbstract.set(numId, absId);
  }

  return { numToAbstract, abstractLevels };
}

function isOrderedList(
  paraXml: string,
  numbering: NumberingIndex
): boolean {
  const numId =
    /<w:numId\b[^>]*w:val="([^"]+)"/.exec(paraXml)?.[1] ??
    /<w:numId\b[^>]*w:val='([^']+)'/.exec(paraXml)?.[1];
  if (numId == null) return false;
  const absId = numbering.numToAbstract.get(numId);
  if (absId == null) return false;
  const ilvl = listLevel(paraXml);
  const fmt = numbering.abstractLevels.get(absId)?.get(ilvl) ?? 'bullet';
  return fmt !== 'bullet' && fmt !== 'none';
}

function extractRunTexts(inner: string): string {
  const texts: string[] = [];
  // Preserve tabs (TOC page-number separator) and text runs in document order.
  const tokenRe = /<w:tab\b[^/]*\/>|<w:tab\b[^>]*>\s*<\/w:tab>|<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
  let tm: RegExpExecArray | null;
  while ((tm = tokenRe.exec(inner)) !== null) {
    if (tm[0].startsWith('<w:tab')) texts.push('\t');
    else texts.push(decodeXmlEntities(tm[1] ?? ''));
  }
  return texts.join('');
}

function isOnOffEnabled(rPr: string, tag: string): boolean {
  // Word may disable with w:val="0" / "false" even when the element is present.
  const re = new RegExp(`<w:${tag}\\b([^>]*)>`, 'i');
  const m = re.exec(rPr);
  if (!m) return false;
  const attrs = m[1];
  const val =
    /w:val="([^"]*)"/i.exec(attrs)?.[1] ?? /w:val='([^']*)'/i.exec(attrs)?.[1];
  if (val != null) {
    const v = val.toLowerCase();
    return v !== '0' && v !== 'false' && v !== 'off';
  }
  return true;
}

function flagsFromRPr(
  inner: string,
  charStyles: Map<string, CharStyleFlags>
): CharStyleFlags {
  const rPr = /<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/.exec(inner)?.[0] ?? '';
  const styleId =
    /<w:rStyle\b[^>]*w:val="([^"]+)"/.exec(rPr)?.[1] ??
    /<w:rStyle\b[^>]*w:val='([^']+)'/.exec(rPr)?.[1];
  const fromStyle = styleId ? charStyles.get(styleId) ?? {} : {};

  const bold =
    fromStyle.bold || isOnOffEnabled(rPr, 'b') || isOnOffEnabled(rPr, 'bCs');
  const italic =
    fromStyle.italic || isOnOffEnabled(rPr, 'i') || isOnOffEnabled(rPr, 'iCs');
  const strike =
    fromStyle.strike ||
    isOnOffEnabled(rPr, 'strike') ||
    isOnOffEnabled(rPr, 'dstrike');

  return {
    ...(bold ? { bold: true } : {}),
    ...(italic ? { italic: true } : {}),
    ...(strike ? { strike: true } : {}),
  };
}

function parseRunElement(
  runXml: string,
  charStyles: Map<string, CharStyleFlags>,
  href?: string
): InlineRun | null {
  // Footnote / endnote references inside a run
  const fn =
    /<w:footnoteReference\b[^>]*w:id="([^"]+)"/.exec(runXml)?.[1] ??
    /<w:footnoteReference\b[^>]*w:id='([^']+)'/.exec(runXml)?.[1];
  if (fn != null && fn !== '0' && fn !== '-1') {
    return { text: `[^fn${fn}]` };
  }
  const en =
    /<w:endnoteReference\b[^>]*w:id="([^"]+)"/.exec(runXml)?.[1] ??
    /<w:endnoteReference\b[^>]*w:id='([^']+)'/.exec(runXml)?.[1];
  if (en != null && en !== '0' && en !== '-1') {
    return { text: `[^en${en}]` };
  }

  const text = extractRunTexts(runXml);
  if (!text) return null;
  const flags = flagsFromRPr(runXml, charStyles);
  return {
    text,
    ...flags,
    ...(href ? { href } : {}),
  };
}

/**
 * Walk paragraph children in order: hyperlinks and runs (skip nested runs already inside hyperlink).
 */
function parseParagraphRuns(
  paraXml: string,
  charStyles: Map<string, CharStyleFlags>,
  rels: Map<string, string>
): InlineRun[] {
  const runs: InlineRun[] = [];
  // Strip pPr to avoid matching inside props
  const withoutPPr = paraXml.replace(/<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/, '');

  const tokenRe = /<w:hyperlink\b[^>]*>[\s\S]*?<\/w:hyperlink>|<w:r\b[^>]*>[\s\S]*?<\/w:r>/g;
  let m: RegExpExecArray | null;
  while ((m = tokenRe.exec(withoutPPr)) !== null) {
    const chunk = m[0];
    if (chunk.startsWith('<w:hyperlink')) {
      const open = chunk.match(/<w:hyperlink\b[^>]*>/)?.[0] ?? '';
      const rId = attr(open, 'r:id');
      const anchor = attr(open, 'w:anchor');
      let href: string | undefined;
      if (rId && rels.has(rId)) href = rels.get(rId);
      else if (anchor) href = `#${anchor}`;

      const innerRunRe = /<w:r\b[^>]*>[\s\S]*?<\/w:r>/g;
      let rm: RegExpExecArray | null;
      while ((rm = innerRunRe.exec(chunk)) !== null) {
        const run = parseRunElement(rm[0], charStyles, href);
        if (run) runs.push(run);
      }
    } else {
      const run = parseRunElement(chunk, charStyles);
      if (run) runs.push(run);
    }
  }

  if (runs.length === 0) {
    const text = extractRunTexts(withoutPPr);
    // Also strip tab-only TOC leftovers
    const cleaned = text.replace(/\t/g, ' ').replace(/\s+/g, ' ').trim();
    if (cleaned) runs.push({ text: cleaned });
  }
  return runs;
}

function paragraphStyleId(paraXml: string): string | undefined {
  return (
    /<w:pStyle\b[^>]*w:val="([^"]+)"/.exec(paraXml)?.[1] ??
    /<w:pStyle\b[^>]*w:val='([^']+)'/.exec(paraXml)?.[1]
  );
}

function headingLevel(paraXml: string): 1 | 2 | 3 | 4 | 5 | 6 | null {
  const style = paragraphStyleId(paraXml);
  if (style) {
    if (/^Title$/i.test(style)) return 1;
    const hm = /^Heading\s*([1-6])$/i.exec(style) || /^heading\s*([1-6])$/i.exec(style);
    if (hm) return Number(hm[1]) as 1 | 2 | 3 | 4 | 5 | 6;
  }
  const outline = /<w:outlineLvl\b[^>]*w:val="(\d+)"/.exec(paraXml);
  if (outline) {
    const lvl = Number(outline[1]) + 1;
    if (lvl >= 1 && lvl <= 6) return lvl as 1 | 2 | 3 | 4 | 5 | 6;
  }
  return null;
}

function tocLevel(paraXml: string): number | null {
  const style = paragraphStyleId(paraXml);
  if (!style) return null;
  if (/^TOCHeading$/i.test(style)) return 0;
  const m = /^TOC([1-9])$/i.exec(style);
  if (m) return Number(m[1]) - 1;
  return null;
}

function isListItem(paraXml: string): boolean {
  return /<w:numPr[\s\S]*?<\/w:numPr>/.test(paraXml) || /<w:numPr\b[^/]*\/>/.test(paraXml);
}

function listLevel(paraXml: string): number {
  const m = /<w:ilvl\b[^>]*w:val="(\d+)"/.exec(paraXml);
  return m ? Number(m[1]) : 0;
}

function isDropCap(paraXml: string): boolean {
  return /w:dropCap="/.test(paraXml) || /w:dropCap='/.test(paraXml);
}

/** Compact nested table for embedding in a cell (no `|` — keeps parent GFM table stable). */
function compactTableText(rows: string[][]): string {
  return rows
    .map((r) => r.map((c) => c.replace(/\s+/g, ' ').trim()).filter(Boolean).join(' · '))
    .filter(Boolean)
    .join(' / ');
}

function parseTableRows(tblXml: string): string[][] {
  const rows: string[][] = [];
  // Walk top-level rows only via balanced scan inside tbl
  const inner = tblXml.replace(/^<w:tbl\b[^>]*>/, '').replace(/<\/w:tbl>$/, '');
  let i = 0;
  while (i < inner.length) {
    const trAt = inner.indexOf('<w:tr', i);
    if (trAt < 0) break;
    const after = inner.charAt(trAt + 5);
    if (after && !/[\s/>]/.test(after)) {
      i = trAt + 5;
      continue;
    }
    const trEnd = findBalancedEnd(inner, trAt, 'tr');
    if (trEnd < 0) {
      i = trAt + 1;
      continue;
    }
    const trXml = inner.slice(trAt, trEnd);
    const cells: string[] = [];
    let j = 0;
    const trInner = trXml;
    while (j < trInner.length) {
      const tcAt = trInner.indexOf('<w:tc', j);
      if (tcAt < 0) break;
      const afterTc = trInner.charAt(tcAt + 5);
      if (afterTc && !/[\s/>]/.test(afterTc)) {
        j = tcAt + 5;
        continue;
      }
      const tcEnd = findBalancedEnd(trInner, tcAt, 'tc');
      if (tcEnd < 0) {
        j = tcAt + 1;
        continue;
      }
      cells.push(parseCellText(trInner.slice(tcAt, tcEnd)));
      j = tcEnd;
    }
    if (cells.length) rows.push(cells);
    i = trEnd;
  }
  return rows;
}

function parseCellText(tcXml: string): string {
  const parts: string[] = [];
  // Remove tcPr
  let rest = tcXml.replace(/<w:tcPr\b[^>]*>[\s\S]*?<\/w:tcPr>/, '');
  let i = 0;
  while (i < rest.length) {
    const tblAt = rest.indexOf('<w:tbl', i);
    const pAt = rest.indexOf('<w:p', i);
    if (tblAt < 0 && pAt < 0) break;
    if (tblAt >= 0 && (pAt < 0 || tblAt < pAt)) {
      const after = rest.charAt(tblAt + 6);
      if (after && !/[\s/>]/.test(after)) {
        i = tblAt + 6;
        continue;
      }
      const end = findBalancedEnd(rest, tblAt, 'tbl');
      if (end < 0) {
        i = tblAt + 1;
        continue;
      }
      const nestedRows = parseTableRows(rest.slice(tblAt, end));
      const compact = compactTableText(nestedRows);
      if (compact) parts.push(compact);
      i = end;
    } else {
      const after = rest.charAt(pAt + 4);
      if (after && !/[\s/>]/.test(after)) {
        i = pAt + 4;
        continue;
      }
      const end = findBalancedEnd(rest, pAt, 'p');
      if (end < 0) {
        i = pAt + 1;
        continue;
      }
      const text = extractRunTexts(rest.slice(pAt, end)).replace(/\t/g, ' ').trim();
      if (text) parts.push(text);
      i = end;
    }
  }
  return parts.join(' ').trim();
}

function parseTable(tblXml: string): ContentBlock {
  return { type: 'table', rows: parseTableRows(tblXml) };
}

type ParseCtx = {
  charStyles: Map<string, CharStyleFlags>;
  numbering: NumberingIndex;
  rels: Map<string, string>;
};

function paraToBlock(paraXml: string, ctx: ParseCtx): ContentBlock | null {
  const runs = parseParagraphRuns(paraXml, ctx.charStyles, ctx.rels);
  // TOC: take text before tab (page number after)
  const tocLvl = tocLevel(paraXml);
  if (tocLvl != null) {
    const plain = runs
      .map((r) => r.text)
      .join('')
      .split('\t')[0]
      .replace(/\s+/g, ' ')
      .trim();
    if (!plain) return null;
    if (/^TOCHeading$/i.test(paragraphStyleId(paraXml) ?? '')) {
      return { type: 'heading', level: 2, runs: [{ text: plain }] };
    }
    return {
      type: 'list_item',
      level: tocLvl,
      runs: [{ text: plain }],
    };
  }

  const plain = runs.map((r) => r.text).join('').trim();
  if (!plain && !isListItem(paraXml)) return null;

  const level = headingLevel(paraXml);
  if (level) {
    return { type: 'heading', level, runs: runs.length ? runs : [{ text: plain }] };
  }
  if (isListItem(paraXml)) {
    const ordered = isOrderedList(paraXml, ctx.numbering);
    return {
      type: 'list_item',
      level: listLevel(paraXml),
      ...(ordered ? { ordered: true } : {}),
      runs: runs.length ? runs : [{ text: plain || ' ' }],
    };
  }
  if (!plain) return null;
  return { type: 'paragraph', runs };
}

function parseNotesXml(
  xml: string | undefined,
  kind: 'footnote' | 'endnote'
): Array<{ id: string; text: string }> {
  if (!xml) return [];
  const tag = kind === 'footnote' ? 'footnote' : 'endnote';
  // Do not let self-closing separator tags swallow the next real note.
  const re = new RegExp(
    `<w:${tag}\\b(?![^>]*/>)[^>]*>([\\s\\S]*?)<\\/w:${tag}>`,
    'g'
  );
  const out: Array<{ id: string; text: string }> = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(xml)) !== null) {
    const open = m[0].slice(0, m[0].indexOf('>') + 1);
    const type = attr(open, 'w:type');
    if (type === 'separator' || type === 'continuationSeparator') continue;
    const id = attr(open, 'w:id');
    if (id == null || id === '-1' || id === '0') continue;
    const text = extractRunTexts(m[1]).replace(/\s+/g, ' ').trim();
    if (text) out.push({ id, text });
  }
  return out;
}

/** Parse DOCX package parts into structural blocks. */
export function parseDocxBlocksFromPackage(parts: DocxPackageParts): ContentBlock[] {
  const ctx: ParseCtx = {
    charStyles: parseCharStyles(parts.stylesXml),
    numbering: parseNumbering(parts.numberingXml),
    rels: parseRels(parts.relsXml),
  };

  const body = extractBodyInner(parts.documentXml);
  const blocks: ContentBlock[] = [];
  let dropCapBuf = '';

  for (const part of iterBodyParts(body)) {
    if (part.kind === 'tbl') {
      dropCapBuf = '';
      const table = parseTable(part.xml);
      if (table.type === 'table' && table.rows.length) blocks.push(table);
      continue;
    }

    if (isDropCap(part.xml)) {
      const letter = extractRunTexts(part.xml).trim();
      dropCapBuf += letter;
      continue;
    }

    const block = paraToBlock(part.xml, ctx);
    if (!block) continue;

    if (dropCapBuf && (block.type === 'paragraph' || block.type === 'list_item')) {
      const runs = [...block.runs];
      if (runs.length) {
        runs[0] = { ...runs[0], text: dropCapBuf + runs[0].text };
      } else {
        runs.push({ text: dropCapBuf });
      }
      dropCapBuf = '';
      blocks.push({ ...block, runs });
    } else {
      dropCapBuf = '';
      blocks.push(block);
    }
  }

  const footnotes = parseNotesXml(parts.footnotesXml, 'footnote');
  const endnotes = parseNotesXml(parts.endnotesXml, 'endnote');
  if (footnotes.length || endnotes.length) {
    for (const fn of footnotes) {
      blocks.push({
        type: 'paragraph',
        runs: [{ text: `[^fn${fn.id}]: ${fn.text}` }],
      });
    }
    for (const en of endnotes) {
      blocks.push({
        type: 'paragraph',
        runs: [{ text: `[^en${en.id}]: ${en.text}` }],
      });
    }
  }

  return blocks;
}

/** Parse word/document.xml only (no styles/numbering/links/notes). */
export function parseDocxBlocksFromXml(xml: string): ContentBlock[] {
  return parseDocxBlocksFromPackage({ documentXml: xml });
}

export function docxStructuredFromPackage(parts: DocxPackageParts): DocxStructured {
  const blocks = parseDocxBlocksFromPackage(parts);
  return {
    format: 'docx',
    blocks,
    paragraphs: flattenBlocksToParagraphs(blocks),
  };
}

export function docxStructuredFromXml(xml: string): DocxStructured {
  return docxStructuredFromPackage({ documentXml: xml });
}

export async function docxStructuredFromZip(zip: JSZip): Promise<DocxStructured> {
  const docFile = zip.file('word/document.xml');
  if (!docFile) {
    throw new DocumentError('PARSE_FAILED', 'Invalid docx: missing word/document.xml');
  }
  const documentXml = await docFile.async('string');
  const load = async (path: string) => {
    const f = zip.file(path);
    return f ? f.async('string') : undefined;
  };
  return docxStructuredFromPackage({
    documentXml,
    stylesXml: await load('word/styles.xml'),
    numberingXml: await load('word/numbering.xml'),
    relsXml: await load('word/_rels/document.xml.rels'),
    footnotesXml: await load('word/footnotes.xml'),
    endnotesXml: await load('word/endnotes.xml'),
  });
}
