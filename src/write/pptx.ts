import JSZip from 'jszip';
import type { CanonicalFlowBlock, CanonicalInline } from '../ir/types.js';
import type { DocumentWarning } from '../types.js';
import { escapeXml } from '../utils/xml.js';

type SlideLine =
  | { type: 'paragraph'; text: string }
  | { type: 'bullet'; text: string }
  | { type: 'numbered'; text: string }
  | { type: 'table'; rows: string[][] };

type SlideContent = { title: string; body: SlideLine[] };

export type PptxWriteOutput = {
  buffer: Buffer;
  warnings: DocumentWarning[];
};

/** Standard 16:9 slide size in EMUs (13.333" × 7.5"). */
const SLIDE_CX = 12192000;
const SLIDE_CY = 6858000;

function runsToPlain(runs: CanonicalInline[]): string {
  return runs.map((r) => r.text).join('');
}

function runsHaveRichFormatting(runs: CanonicalInline[]): boolean {
  return runs.some(
    (r) =>
      Boolean(r.bold) ||
      Boolean(r.italic) ||
      Boolean(r.strike) ||
      Boolean(r.code) ||
      Boolean(r.underline) ||
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

function noteInlineFlatten(
  runs: CanonicalInline[],
  warnings: DocumentWarning[],
  seen: Set<string>
): void {
  if (runsHaveRichFormatting(runs)) {
    pushWarning(
      warnings,
      seen,
      'PPTX_INLINE_FLATTENED',
      'PPTX_INLINE_FLATTENED: bold/italic/strike/link/code inline formatting is flattened to plain text in PPTX write'
    );
  }
}

function appendList(
  list: Extract<CanonicalFlowBlock, { type: 'list' }>,
  body: SlideLine[],
  warnings: DocumentWarning[],
  seen: Set<string>
): void {
  for (const item of list.items) {
    const texts: string[] = [];
    const nested: Extract<CanonicalFlowBlock, { type: 'list' }>[] = [];
    for (const sub of item) {
      if (sub.type === 'paragraph' || sub.type === 'heading') {
        noteInlineFlatten(sub.runs, warnings, seen);
        const t = runsToPlain(sub.runs).trim();
        if (t) texts.push(t);
      } else if (sub.type === 'list') {
        nested.push(sub);
      }
    }
    const text = texts.join(' ').trim();
    if (text) {
      body.push(
        list.ordered ? { type: 'numbered', text } : { type: 'bullet', text }
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
    noteInlineFlatten(block.runs, warnings, seen);
    const t = runsToPlain(block.runs).trim();
    if (t) body.push({ type: 'paragraph', text: t });
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
      if (line.length) body.push({ type: 'paragraph', text: line });
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
        if (cell.runs?.length) noteInlineFlatten(cell.runs, warnings, seen);
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
      noteInlineFlatten(block.runs, warnings, seen);
      current = { title: runsToPlain(block.runs).trim() || 'Slide', body: [] };
      continue;
    }
    if (!current) {
      current = { title: 'Slide', body: [] };
    }
    appendBlockLines(block, current.body, warnings, seen);
  }
  flush();
  return slides.length ? slides : [{ title: 'Slide', body: [] }];
}

function textRunXml(text: string, bold?: boolean): string {
  const rPr = bold
    ? '<a:rPr lang="en-US" b="1" dirty="0"/>'
    : '<a:rPr lang="en-US" dirty="0"/>';
  const space =
    text.startsWith(' ') || text.endsWith(' ') || text.includes('  ')
      ? ' xml:space="preserve"'
      : '';
  return `<a:r>${rPr}<a:t${space}>${escapeXml(text)}</a:t></a:r>`;
}

function paraXml(
  text: string,
  opts?: { bold?: boolean; bullet?: boolean; numbered?: boolean; num?: number }
): string {
  let pPr = '<a:pPr/>';
  if (opts?.bullet) {
    pPr = '<a:pPr marL="342900" indent="-342900"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>';
  } else if (opts?.numbered) {
    pPr = `<a:pPr marL="342900" indent="-342900"><a:buFont typeface="Arial"/><a:buAutoNum type="arabicPeriod" startAt="${opts.num ?? 1}"/></a:pPr>`;
  }
  return `<a:p>${pPr}${textRunXml(text, opts?.bold)}<a:endParaRPr lang="en-US"/></a:p>`;
}

function tableXml(rows: string[][]): string {
  const colCount = Math.max(1, ...rows.map((r) => r.length));
  const colWidth = Math.floor(8229600 / colCount);
  const grid = Array.from(
    { length: colCount },
    () => `<a:gridCol w="${colWidth}"/>`
  ).join('');
  const trs = rows
    .map((row) => {
      const cells = Array.from({ length: colCount }, (_, i) => {
        const text = row[i] ?? '';
        return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p>${textRunXml(text)}<a:endParaRPr lang="en-US"/></a:p></a:txBody><a:tcPr/></a:tc>`;
      }).join('');
      return `<a:tr h="370840">${cells}</a:tr>`;
    })
    .join('');
  return `<a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>${grid}</a:tblGrid>${trs}</a:tbl>`;
}

function graphicFrameXml(rows: string[][], id: number, y: number): string {
  const rowCount = Math.max(1, rows.length);
  const cy = Math.min(4000000, Math.max(370840, rowCount * 370840 + 80000));
  return `<p:graphicFrame>
      <p:nvGraphicFramePr>
        <p:cNvPr id="${id}" name="Table ${id}"/>
        <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
        <p:nvPr/>
      </p:nvGraphicFramePr>
      <p:xfrm>
        <a:off x="457200" y="${y}"/>
        <a:ext cx="8229600" cy="${cy}"/>
      </p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          ${tableXml(rows)}
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>`;
}

function textShapeXml(paras: string[], id: number, y: number, cy: number): string {
  return `<p:sp>
      <p:nvSpPr>
        <p:cNvPr id="${id}" name="Content ${id}"/>
        <p:cNvSpPr txBox="1"/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="457200" y="${y}"/>
          <a:ext cx="8229600" cy="${cy}"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap="square" rtlCol="0"/>
        <a:lstStyle/>
        ${paras.join('\n        ')}
      </p:txBody>
    </p:sp>`;
}

function slideXml(slide: SlideContent): string {
  const paras: string[] = [paraXml(slide.title || 'Slide', { bold: true })];
  const tables: string[][][] = [];
  let numberAt = 1;
  for (const line of slide.body) {
    if (line.type === 'paragraph') {
      paras.push(paraXml(line.text));
      numberAt = 1;
    } else if (line.type === 'bullet') {
      paras.push(paraXml(line.text, { bullet: true }));
      numberAt = 1;
    } else if (line.type === 'numbered') {
      paras.push(paraXml(line.text, { numbered: true, num: numberAt++ }));
    } else if (line.type === 'table') {
      tables.push(line.rows);
    }
  }

  const hasBodyParas = paras.length > 1;
  const textCy = tables.length
    ? hasBodyParas
      ? 2200000
      : 900000
    : 5600000;
  const textShape = textShapeXml(paras, 2, 274320, textCy);
  const tableStartY = 274320 + textCy + 120000;
  const tableShapes = tables
    .map((rows, index) =>
      graphicFrameXml(rows, 10 + index, tableStartY + index * 3200000)
    )
    .join('\n    ');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
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
      ${textShape}
      ${tableShapes}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`;
}

const THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light" pitchFamily="34" charset="0"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri" pitchFamily="34" charset="0"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
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

const SLIDE_MASTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef>
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

/**
 * Minimal first-party PPTX writer (jszip + DrawingML).
 * Emits an Office-openable package: theme + slideMaster + slideLayout + slide rels.
 * Lists → bulleted/numbered paragraphs; tables → DrawingML `a:tbl`.
 * Returns buffer plus degradation warnings (inline/blockquote/code flatten).
 */
export async function writePptxFromBlocks(
  blocks: CanonicalFlowBlock[]
): Promise<PptxWriteOutput> {
  const warnings: DocumentWarning[] = [];
  const seen = new Set<string>();
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

  zip.file('ppt/theme/theme1.xml', THEME_XML);
  zip.file('ppt/slideMasters/slideMaster1.xml', SLIDE_MASTER_XML);
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
    zip.file(`ppt/slides/slide${n}.xml`, slideXml(slide));
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
