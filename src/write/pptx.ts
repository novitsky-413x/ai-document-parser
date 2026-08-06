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

function paraXml(
  text: string,
  opts?: { bold?: boolean; bullet?: boolean; numbered?: boolean; num?: number }
): string {
  const rPr = opts?.bold ? '<a:rPr b="1"/>' : '';
  let pPr = '';
  if (opts?.bullet) {
    pPr = '<a:pPr marL="342900"><a:buChar char="•"/></a:pPr>';
  } else if (opts?.numbered) {
    pPr = `<a:pPr marL="342900"><a:buAutoNum type="arabicPeriod" startAt="${opts.num ?? 1}"/></a:pPr>`;
  }
  return `<a:p>${pPr}<a:r>${rPr}<a:t>${escapeXml(text)}</a:t></a:r></a:p>`;
}

function tableXml(rows: string[][]): string {
  const colCount = Math.max(1, ...rows.map((r) => r.length));
  const colWidth = Math.floor(8000000 / colCount);
  const grid = Array.from(
    { length: colCount },
    () => `<a:gridCol w="${colWidth}"/>`
  ).join('');
  const trs = rows
    .map((row) => {
      const cells = Array.from({ length: colCount }, (_, i) => {
        const text = row[i] ?? '';
        return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeXml(text)}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`;
      }).join('');
      return `<a:tr h="370840">${cells}</a:tr>`;
    })
    .join('');
  return `<a:tbl><a:tblPr/><a:tblGrid>${grid}</a:tblGrid>${trs}</a:tbl>`;
}

function graphicFrameXml(rows: string[][], id: number, y: number): string {
  return `<p:graphicFrame>
      <p:nvGraphicFramePr>
        <p:cNvPr id="${id}" name="Table ${id}"/>
        <p:cNvGraphicFramePr/>
        <p:nvPr/>
      </p:nvGraphicFramePr>
      <p:xfrm>
        <a:off x="457200" y="${y}"/>
        <a:ext cx="8229600" cy="3000000"/>
      </p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          ${tableXml(rows)}
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>`;
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

  const tableShapes = tables
    .map((rows, index) => graphicFrameXml(rows, 10 + index, 2000000 + index * 3200000))
    .join('\n    ');

  return `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Content"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
      <p:spPr/>
      <p:txBody>
        <a:bodyPr/><a:lstStyle/>
        ${paras.join('\n        ')}
      </p:txBody>
    </p:sp>
    ${tableShapes}
  </p:spTree></p:cSld>
</p:sld>`;
}

/**
 * Minimal first-party PPTX writer (jszip + DrawingML).
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

  const overrides = slides
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
  ${overrides}
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
    .map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`)
    .join('');
  zip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>${sldIds}</p:sldIdLst>
</p:presentation>`
  );

  const rels = slides
    .map(
      (_, i) =>
        `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`
    )
    .join('');
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${rels}
</Relationships>`
  );

  slides.forEach((slide, i) => {
    zip.file(`ppt/slides/slide${i + 1}.xml`, slideXml(slide));
  });

  const buffer = await zip.generateAsync({ type: 'nodebuffer' });
  return { buffer, warnings };
}
