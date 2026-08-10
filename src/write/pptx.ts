import JSZip from 'jszip';
import type { CanonicalFlowBlock, CanonicalInline } from '../ir/types.js';
import type { DocumentWarning, PptxThemeId } from '../types.js';
import { escapeXml } from '../utils/xml.js';

export type { PptxThemeId };

type SlideLine =
  | { type: 'paragraph'; runs: CanonicalInline[] }
  | { type: 'bullet'; runs: CanonicalInline[] }
  | { type: 'numbered'; runs: CanonicalInline[] }
  | { type: 'table'; rows: string[][] };

type SlideContent = { title: CanonicalInline[]; body: SlideLine[] };

export type PptxWriteOptions = {
  /** Visual theme preset (colors + accent bar). Default: office. */
  theme?: PptxThemeId;
};

export type PptxWriteOutput = {
  buffer: Buffer;
  warnings: DocumentWarning[];
};

/** Standard 16:9 slide size in EMUs (13.333" × 7.5"). */
const SLIDE_CX = 12192000;
const SLIDE_CY = 6858000;

type ThemePreset = {
  id: PptxThemeId;
  name: string;
  dk1: string;
  lt1: string;
  dk2: string;
  lt2: string;
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
  hlink: string;
  folHlink: string;
  majorFont: string;
  minorFont: string;
  /** Solid slide background (srgb). */
  bg: string;
  /** Title text color. */
  titleColor: string;
  /** Body text color. */
  bodyColor: string;
  /** Table header fill. */
  tableHeaderFill: string;
  /** Table header text. */
  tableHeaderText: string;
};

const THEMES: Record<PptxThemeId, ThemePreset> = {
  office: {
    id: 'office',
    name: 'Office',
    dk1: '000000',
    lt1: 'FFFFFF',
    dk2: '44546A',
    lt2: 'E7E6E6',
    accent1: '4472C4',
    accent2: 'ED7D31',
    accent3: 'A5A5A5',
    accent4: 'FFC000',
    accent5: '5B9BD5',
    accent6: '70AD47',
    hlink: '0563C1',
    folHlink: '954F72',
    majorFont: 'Calibri Light',
    minorFont: 'Calibri',
    bg: 'FFFFFF',
    titleColor: '2F5496',
    bodyColor: '333333',
    tableHeaderFill: '4472C4',
    tableHeaderText: 'FFFFFF',
  },
  ocean: {
    id: 'ocean',
    name: 'Ocean',
    dk1: '0B1F2A',
    lt1: 'F4FAFC',
    dk2: '1B3A4B',
    lt2: 'D7E8EF',
    accent1: '0E7490',
    accent2: 'F59E0B',
    accent3: '64748B',
    accent4: '14B8A6',
    accent5: '38BDF8',
    accent6: '22C55E',
    hlink: '0369A1',
    folHlink: '7C3AED',
    majorFont: 'Calibri Light',
    minorFont: 'Calibri',
    bg: 'F4FAFC',
    titleColor: '0E7490',
    bodyColor: '1B3A4B',
    tableHeaderFill: '0E7490',
    tableHeaderText: 'FFFFFF',
  },
  slate: {
    id: 'slate',
    name: 'Slate',
    dk1: 'E8EEF5',
    lt1: '0F172A',
    dk2: 'CBD5E1',
    lt2: '1E293B',
    accent1: '38BDF8',
    accent2: 'F472B6',
    accent3: '94A3B8',
    accent4: 'FBBF24',
    accent5: '818CF8',
    accent6: '34D399',
    hlink: '7DD3FC',
    folHlink: 'C084FC',
    majorFont: 'Calibri Light',
    minorFont: 'Calibri',
    bg: '0F172A',
    titleColor: 'F8FAFC',
    bodyColor: 'E2E8F0',
    tableHeaderFill: '1E293B',
    tableHeaderText: 'F8FAFC',
  },
};

function resolveTheme(id?: PptxThemeId): ThemePreset {
  if (id && THEMES[id]) return THEMES[id];
  return THEMES.office;
}

function runsToPlain(runs: CanonicalInline[]): string {
  return runs.map((r) => r.text).join('');
}

function runsNeedLinkWarning(runs: CanonicalInline[]): boolean {
  return runs.some(
    (r) =>
      Boolean(r.href) ||
      Boolean(r.subscript) ||
      Boolean(r.superscript)
  );
}

function pushWarning(
  warnings: DocumentWarning[],
  seen: Set<string>,
  code: string,
  message: string
): void {
  if (seen.has(code)) return;
  seen.add(code);
  warnings.push({ code, message });
}

function noteInlineLimits(
  runs: CanonicalInline[],
  warnings: DocumentWarning[],
  seen: Set<string>
): void {
  if (runsNeedLinkWarning(runs)) {
    pushWarning(
      warnings,
      seen,
      'PPTX_INLINE_FLATTENED',
      'PPTX_INLINE_FLATTENED: hyperlinks/sub/sup are styled as text (no clickable relationship); bold/italic/strike/underline/code are preserved'
    );
  }
}

/** Insert a space when joining adjacent list-item paragraphs. */
function joinRunsWithGap(
  target: CanonicalInline[],
  next: CanonicalInline[]
): void {
  if (!next.length) return;
  if (target.length) {
    const prevText = runsToPlain(target);
    const nextText = runsToPlain(next);
    if (
      prevText.length &&
      nextText.trim().length &&
      !/\s$/.test(prevText) &&
      !/^\s/.test(nextText)
    ) {
      target.push({ text: ' ' });
    }
  }
  for (const r of next) {
    if (r.text) target.push(r);
  }
}

function appendList(
  list: Extract<CanonicalFlowBlock, { type: 'list' }>,
  body: SlideLine[],
  warnings: DocumentWarning[],
  seen: Set<string>
): void {
  for (const item of list.items) {
    const runs: CanonicalInline[] = [];
    const nested: Extract<CanonicalFlowBlock, { type: 'list' }>[] = [];
    for (const sub of item) {
      if (sub.type === 'paragraph' || sub.type === 'heading') {
        noteInlineLimits(sub.runs, warnings, seen);
        joinRunsWithGap(runs, sub.runs);
      } else if (sub.type === 'list') {
        nested.push(sub);
      }
    }
    if (runsToPlain(runs).trim()) {
      body.push(
        list.ordered ? { type: 'numbered', runs } : { type: 'bullet', runs }
      );
    }
    for (const sub of nested) appendList(sub, body, warnings, seen);
  }
}

function appendBlockLines(
  block: CanonicalFlowBlock,
  body: SlideLine[],
  warnings: DocumentWarning[],
  seen: Set<string>
): void {
  if (block.type === 'heading' || block.type === 'paragraph') {
    noteInlineLimits(block.runs, warnings, seen);
    if (runsToPlain(block.runs).trim()) {
      body.push({ type: 'paragraph', runs: block.runs });
    }
    return;
  }
  if (block.type === 'code') {
    pushWarning(
      warnings,
      seen,
      'PPTX_CODE_AS_PARAGRAPHS',
      'PPTX_CODE_AS_PARAGRAPHS: fenced/code blocks are emitted as ordinary paragraphs in PPTX'
    );
    for (const line of block.text.replace(/\r\n/g, '\n').split('\n')) {
      if (line.length) {
        body.push({
          type: 'paragraph',
          runs: [{ text: line, code: true }],
        });
      }
    }
    return;
  }
  if (block.type === 'blockquote') {
    pushWarning(
      warnings,
      seen,
      'PPTX_BLOCKQUOTE_UNWRAPPED',
      'PPTX_BLOCKQUOTE_UNWRAPPED: blockquotes are unwrapped to ordinary paragraphs in PPTX (quote semantics lost)'
    );
    for (const inner of block.blocks) appendBlockLines(inner, body, warnings, seen);
    return;
  }
  if (block.type === 'list') {
    appendList(block, body, warnings, seen);
    return;
  }
  if (block.type === 'table') {
    for (const row of block.rows) {
      for (const cell of row) {
        if (cell.runs?.length) noteInlineLimits(cell.runs, warnings, seen);
      }
    }
    body.push({
      type: 'table',
      rows: block.rows.map((row) => row.map((cell) => cell.text)),
    });
  }
}

/** Split flow blocks into slides: `#` / `##` start a new slide. */
export function flowBlocksToSlides(
  blocks: CanonicalFlowBlock[],
  warnings: DocumentWarning[] = [],
  seen: Set<string> = new Set()
): SlideContent[] {
  const slides: SlideContent[] = [];
  let current: SlideContent | null = null;

  const flush = () => {
    if (current) slides.push(current);
    current = null;
  };

  for (const block of blocks) {
    if (block.type === 'heading' && block.level <= 2) {
      flush();
      noteInlineLimits(block.runs, warnings, seen);
      const title = block.runs.length
        ? block.runs
        : [{ text: 'Slide' }];
      current = { title, body: [] };
      continue;
    }
    if (!current) {
      current = { title: [{ text: 'Slide' }], body: [] };
    }
    appendBlockLines(block, current.body, warnings, seen);
  }
  flush();
  return slides.length ? slides : [{ title: [{ text: 'Slide' }], body: [] }];
}

function tSpace(text: string): string {
  return text.startsWith(' ') || text.endsWith(' ') || text.includes('  ')
    ? ' xml:space="preserve"'
    : '';
}

function runXml(
  run: CanonicalInline,
  theme: ThemePreset,
  opts?: { sizePt?: number; color?: string; forceBold?: boolean }
): string {
  const sizePt = opts?.sizePt ?? 18;
  const sz = Math.round(sizePt * 100);
  const color = opts?.color ?? theme.bodyColor;
  const attrs: string[] = [`lang="en-US"`, `sz="${sz}"`, `dirty="0"`];
  if (opts?.forceBold || run.bold) attrs.push('b="1"');
  if (run.italic) attrs.push('i="1"');
  if (run.strike) attrs.push('strike="sngStrike"');
  if (run.underline || run.href) attrs.push('u="sng"');
  const typeface = run.code ? 'Consolas' : theme.minorFont;
  const solid = run.href ? theme.hlink : run.code ? 'B45309' : color;
  return `<a:r><a:rPr ${attrs.join(' ')}><a:solidFill><a:srgbClr val="${solid}"/></a:solidFill><a:latin typeface="${escapeXml(typeface)}"/></a:rPr><a:t${tSpace(run.text)}>${escapeXml(run.text)}</a:t></a:r>`;
}

function runsXml(
  runs: CanonicalInline[],
  theme: ThemePreset,
  opts?: { sizePt?: number; color?: string; forceBold?: boolean }
): string {
  const list = runs.length ? runs : [{ text: '' }];
  return list.map((r) => runXml(r, theme, opts)).join('');
}

/** DrawingML paragraph spacing after; `pts` is in points (→ hundredths). */
function spcAftXml(pts?: number): string {
  if (pts == null || pts <= 0) return '';
  return `<a:spcAft><a:spcPts val="${Math.round(pts * 100)}"/></a:spcAft>`;
}

function paraXml(
  runs: CanonicalInline[],
  theme: ThemePreset,
  opts?: {
    bullet?: boolean;
    numbered?: boolean;
    num?: number;
    sizePt?: number;
    color?: string;
    forceBold?: boolean;
    /** Space after paragraph, in points. */
    spcAft?: number;
  }
): string {
  const spacing = spcAftXml(opts?.spcAft);
  let pPr = spacing ? `<a:pPr>${spacing}</a:pPr>` : '<a:pPr/>';
  if (opts?.bullet) {
    pPr = `<a:pPr marL="342900" indent="-342900">${spacing}<a:buFont typeface="Arial"/><a:buClr><a:srgbClr val="${theme.accent1}"/></a:buClr><a:buChar char="•"/></a:pPr>`;
  } else if (opts?.numbered) {
    pPr = `<a:pPr marL="342900" indent="-342900">${spacing}<a:buFont typeface="Arial"/><a:buClr><a:srgbClr val="${theme.accent1}"/></a:buClr><a:buAutoNum type="arabicPeriod" startAt="${opts.num ?? 1}"/></a:pPr>`;
  }
  return `<a:p>${pPr}${runsXml(runs, theme, {
    sizePt: opts?.sizePt,
    color: opts?.color,
    forceBold: opts?.forceBold,
  })}<a:endParaRPr lang="en-US"/></a:p>`;
}

function tableXml(rows: string[][], theme: ThemePreset): string {
  const colCount = Math.max(1, ...rows.map((r) => r.length));
  const colWidth = Math.floor(8229600 / colCount);
  const grid = Array.from(
    { length: colCount },
    () => `<a:gridCol w="${colWidth}"/>`
  ).join('');
  const trs = rows
    .map((row, rowIndex) => {
      const isHeader = rowIndex === 0;
      const cells = Array.from({ length: colCount }, (_, i) => {
        const text = row[i] ?? '';
        const fill = isHeader
          ? `<a:solidFill><a:srgbClr val="${theme.tableHeaderFill}"/></a:solidFill>`
          : rowIndex % 2 === 0
            ? `<a:solidFill><a:srgbClr val="${theme.lt2}"/></a:solidFill>`
            : '<a:noFill/>';
        const color = isHeader ? theme.tableHeaderText : theme.bodyColor;
        return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>${runXml(
          { text, bold: isHeader },
          theme,
          { sizePt: 12, color, forceBold: isHeader }
        )}<a:endParaRPr lang="en-US"/></a:p></a:txBody><a:tcPr>${fill}</a:tcPr></a:tc>`;
      }).join('');
      return `<a:tr h="370840">${cells}</a:tr>`;
    })
    .join('');
  return `<a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>${grid}</a:tblGrid>${trs}</a:tbl>`;
}

function estimateTableCy(rows: string[][]): number {
  const rowCount = Math.max(1, rows.length);
  return Math.min(4000000, Math.max(370840, rowCount * 370840 + 80000));
}

/** Rough body height from paragraph count (~18pt line + spacing). */
function estimateBodyCy(paraCount: number): number {
  return Math.max(200000, paraCount * 360000 + 120000);
}

function graphicFrameXml(
  rows: string[][],
  id: number,
  y: number,
  theme: ThemePreset,
  cy?: number
): string {
  const height = cy ?? estimateTableCy(rows);
  return `<p:graphicFrame>
      <p:nvGraphicFramePr>
        <p:cNvPr id="${id}" name="Table ${id}"/>
        <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
        <p:nvPr/>
      </p:nvGraphicFramePr>
      <p:xfrm>
        <a:off x="685800" y="${y}"/>
        <a:ext cx="8229600" cy="${height}"/>
      </p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          ${tableXml(rows, theme)}
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>`;
}

function accentBarXml(theme: ThemePreset): string {
  return `<p:sp>
      <p:nvSpPr>
        <p:cNvPr id="3" name="Accent Bar"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="228600" cy="${SLIDE_CY}"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="${theme.accent1}"/></a:solidFill>
        <a:ln><a:noFill/></a:ln>
      </p:spPr>
    </p:sp>`;
}

function textShapeXml(
  paras: string[],
  id: number,
  x: number,
  y: number,
  cx: number,
  cy: number,
  name: string
): string {
  return `<p:sp>
      <p:nvSpPr>
        <p:cNvPr id="${id}" name="${escapeXml(name)}"/>
        <p:cNvSpPr txBox="1"/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="${x}" y="${y}"/>
          <a:ext cx="${cx}" cy="${cy}"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap="square" rtlCol="0" anchor="t">
          <a:normAutofit/>
        </a:bodyPr>
        <a:lstStyle/>
        ${paras.join('\n        ')}
      </p:txBody>
    </p:sp>`;
}

function slideXml(slide: SlideContent, theme: ThemePreset): string {
  const titlePara = paraXml(slide.title, theme, {
    sizePt: 32,
    color: theme.titleColor,
    forceBold: true,
    spcAft: 12,
  });
  const bodyParas: string[] = [];
  const tables: string[][][] = [];
  let numberAt = 1;
  for (const line of slide.body) {
    if (line.type === 'paragraph') {
      bodyParas.push(
        paraXml(line.runs, theme, {
          sizePt: 18,
          color: theme.bodyColor,
          spcAft: 8,
        })
      );
      numberAt = 1;
    } else if (line.type === 'bullet') {
      bodyParas.push(
        paraXml(line.runs, theme, {
          bullet: true,
          sizePt: 18,
          color: theme.bodyColor,
          spcAft: 6,
        })
      );
      numberAt = 1;
    } else if (line.type === 'numbered') {
      bodyParas.push(
        paraXml(line.runs, theme, {
          numbered: true,
          num: numberAt++,
          sizePt: 18,
          color: theme.bodyColor,
          spcAft: 6,
        })
      );
    } else if (line.type === 'table') {
      tables.push(line.rows);
    }
  }

  const contentX = 685800;
  const contentCx = 11000000;
  const titleY = 274320;
  const titleCy = 900000;
  const bodyY = titleY + titleCy + 80000;
  const contentBottom = SLIDE_CY - 274320;
  const gap = 120000;
  const tableCys = tables.map(estimateTableCy);
  const tablesBlock =
    tableCys.reduce((sum, h) => sum + h, 0) +
    Math.max(0, tables.length - 1) * gap +
    (tables.length && bodyParas.length ? gap : 0);

  let bodyCy = 0;
  if (bodyParas.length) {
    const available = Math.max(200000, contentBottom - bodyY);
    if (!tables.length) {
      bodyCy = available;
    } else {
      const desired = estimateBodyCy(bodyParas.length);
      const maxForBody = Math.max(200000, available - tablesBlock);
      bodyCy = Math.min(desired, maxForBody);
    }
  }

  const titleShape = textShapeXml(
    [titlePara],
    2,
    contentX,
    titleY,
    contentCx,
    titleCy,
    'Title'
  );
  const bodyShape = bodyParas.length
    ? textShapeXml(bodyParas, 4, contentX, bodyY, contentCx, bodyCy, 'Body')
    : '';

  let tableY =
    bodyY + (bodyParas.length ? bodyCy : 0) + (tables.length && bodyParas.length ? gap : 0);
  const tableShapes = tables
    .map((rows, index) => {
      const maxCy = Math.max(370840, contentBottom - tableY);
      const cy = Math.min(tableCys[index]!, maxCy);
      const frame = graphicFrameXml(rows, 10 + index, tableY, theme, cy);
      tableY += cy + gap;
      return frame;
    })
    .join('\n    ');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="${theme.bg}"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="${SLIDE_CX}" cy="${SLIDE_CY}"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="${SLIDE_CX}" cy="${SLIDE_CY}"/>
        </a:xfrm>
      </p:grpSpPr>
      ${accentBarXml(theme)}
      ${titleShape}
      ${bodyShape}
      ${tableShapes}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`;
}

function themeXml(theme: ThemePreset): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${escapeXml(theme.name)}">
  <a:themeElements>
    <a:clrScheme name="${escapeXml(theme.name)}">
      <a:dk1><a:srgbClr val="${theme.dk1}"/></a:dk1>
      <a:lt1><a:srgbClr val="${theme.lt1}"/></a:lt1>
      <a:dk2><a:srgbClr val="${theme.dk2}"/></a:dk2>
      <a:lt2><a:srgbClr val="${theme.lt2}"/></a:lt2>
      <a:accent1><a:srgbClr val="${theme.accent1}"/></a:accent1>
      <a:accent2><a:srgbClr val="${theme.accent2}"/></a:accent2>
      <a:accent3><a:srgbClr val="${theme.accent3}"/></a:accent3>
      <a:accent4><a:srgbClr val="${theme.accent4}"/></a:accent4>
      <a:accent5><a:srgbClr val="${theme.accent5}"/></a:accent5>
      <a:accent6><a:srgbClr val="${theme.accent6}"/></a:accent6>
      <a:hlink><a:srgbClr val="${theme.hlink}"/></a:hlink>
      <a:folHlink><a:srgbClr val="${theme.folHlink}"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="${escapeXml(theme.name)}">
      <a:majorFont>
        <a:latin typeface="${escapeXml(theme.majorFont)}" pitchFamily="34" charset="0"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="${escapeXml(theme.minorFont)}" pitchFamily="34" charset="0"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="${escapeXml(theme.name)}">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1"><a:gsLst>
          <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
          <a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
          <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
        </a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill>
        <a:gradFill rotWithShape="1"><a:gsLst>
          <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="100000"/><a:shade val="100000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
          <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/><a:shade val="100000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
        </a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1"><a:gsLst>
          <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
          <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>
        </a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill>
        <a:gradFill rotWithShape="1"><a:gsLst>
          <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
          <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>
        </a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>`;
}

const SLIDE_LAYOUT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="${SLIDE_CX}" cy="${SLIDE_CY}"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="${SLIDE_CX}" cy="${SLIDE_CY}"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`;

function slideMasterXml(theme: ThemePreset): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="${theme.bg}"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="${SLIDE_CX}" cy="${SLIDE_CY}"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="${SLIDE_CX}" cy="${SLIDE_CY}"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>`;
}

/**
 * First-party PPTX writer (jszip + DrawingML).
 * Office-openable package + theme presets, accent bar, title/body slots,
 * preserved bold/italic/strike/underline/code runs, styled tables.
 */
export async function writePptxFromBlocks(
  blocks: CanonicalFlowBlock[],
  options: PptxWriteOptions = {}
): Promise<PptxWriteOutput> {
  const warnings: DocumentWarning[] = [];
  const seen = new Set<string>();
  const theme = resolveTheme(options.theme);
  const slides = flowBlocksToSlides(blocks, warnings, seen);
  const zip = new JSZip();

  const slideOverrides = slides
    .map(
      (_, i) =>
        `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
    )
    .join('\n  ');

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  ${slideOverrides}
</Types>`
  );

  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`
  );

  const sldIds = slides
    .map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`)
    .join('');

  zip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>${sldIds}</p:sldIdLst>
  <p:sldSz cx="${SLIDE_CX}" cy="${SLIDE_CY}"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`
  );

  const slideRels = slides
    .map(
      (_, i) =>
        `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`
    )
    .join('\n  ');

  zip.file(
    'ppt/_rels/presentation.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  ${slideRels}
</Relationships>`
  );

  zip.file('ppt/theme/theme1.xml', themeXml(theme));
  zip.file('ppt/slideMasters/slideMaster1.xml', slideMasterXml(theme));
  zip.file(
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`
  );
  zip.file('ppt/slideLayouts/slideLayout1.xml', SLIDE_LAYOUT_XML);
  zip.file(
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`
  );

  slides.forEach((slide, i) => {
    const n = i + 1;
    zip.file(`ppt/slides/slide${n}.xml`, slideXml(slide, theme));
    zip.file(
      `ppt/slides/_rels/slide${n}.xml.rels`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`
    );
  });

  const buffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
  });
  return { buffer, warnings };
}
