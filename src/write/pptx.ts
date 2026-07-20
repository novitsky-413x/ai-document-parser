import JSZip from 'jszip';
import type { CanonicalFlowBlock, CanonicalInline } from '../ir/types.js';
import { escapeXml } from '../utils/xml.js';

type SlideContent = { title: string; body: string[] };

function runsToPlain(runs: CanonicalInline[]): string {
  return runs.map((r) => r.text).join('');
}

function blockPlainLines(block: CanonicalFlowBlock): string[] {
  if (block.type === 'heading' || block.type === 'paragraph') {
    const t = runsToPlain(block.runs).trim();
    return t ? [t] : [];
  }
  if (block.type === 'code') {
    return block.text
      .replace(/\r\n/g, '\n')
      .split('\n')
      .map((l) => l.trimEnd())
      .filter((l) => l.length > 0);
  }
  if (block.type === 'blockquote') {
    return block.blocks.flatMap(blockPlainLines);
  }
  if (block.type === 'list') {
    const lines: string[] = [];
    for (const item of block.items) {
      for (const sub of item) {
        lines.push(...blockPlainLines(sub));
      }
    }
    return lines;
  }
  if (block.type === 'table') {
    return block.rows.map((row) =>
      row.map((cell) => cell.text.trim()).filter(Boolean).join(' · ')
    );
  }
  return [];
}

/** Split flow blocks into slides: `#` / `##` start a new slide. */
export function flowBlocksToSlides(blocks: CanonicalFlowBlock[]): SlideContent[] {
  const slides: SlideContent[] = [];
  let current: SlideContent | null = null;

  const flush = () => {
    if (current) slides.push(current);
    current = null;
  };

  for (const block of blocks) {
    if (block.type === 'heading' && block.level <= 2) {
      flush();
      current = { title: runsToPlain(block.runs).trim() || 'Slide', body: [] };
      continue;
    }
    if (!current) {
      current = { title: 'Slide', body: [] };
    }
    current.body.push(...blockPlainLines(block));
  }
  flush();
  return slides.length ? slides : [{ title: 'Slide', body: [] }];
}

function slideXml(slide: SlideContent): string {
  const paras: string[] = [];
  paras.push(
    `<a:p><a:r><a:rPr b="1"/><a:t>${escapeXml(slide.title || 'Slide')}</a:t></a:r></a:p>`
  );
  for (const line of slide.body) {
    paras.push(`<a:p><a:r><a:t>${escapeXml(line)}</a:t></a:r></a:p>`);
  }
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
  </p:spTree></p:cSld>
</p:sld>`;
}

/**
 * Minimal first-party PPTX writer (jszip + DrawingML text).
 * Emits slides that parseDocument and editDocument surgical replace can use.
 */
export async function writePptxFromBlocks(
  blocks: CanonicalFlowBlock[]
): Promise<Buffer> {
  const slides = flowBlocksToSlides(blocks);
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

  return zip.generateAsync({ type: 'nodebuffer' });
}
