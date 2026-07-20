import JSZip from 'jszip';
import type { CanonicalFlowBlock, CanonicalInline } from '../ir/types.js';
import { escapeXml } from '../utils/xml.js';

function runsToOdfXml(runs: CanonicalInline[]): string {
  return runs
    .map((r) => {
      const text = escapeXml(r.text);
      if (r.bold && r.italic) return `<text:span text:style-name="ADPBI">${text}</text:span>`;
      if (r.bold) return `<text:span text:style-name="ADPB">${text}</text:span>`;
      if (r.italic) return `<text:span text:style-name="ADPI">${text}</text:span>`;
      return text;
    })
    .join('');
}

type OdfCtx = { tables: number };

function blockToOdfXml(block: CanonicalFlowBlock, ctx: OdfCtx): string {
  if (block.type === 'heading') {
    return `<text:h text:outline-level="${block.level}">${runsToOdfXml(block.runs)}</text:h>`;
  }
  if (block.type === 'paragraph') {
    return `<text:p>${runsToOdfXml(block.runs)}</text:p>`;
  }
  if (block.type === 'code') {
    return block.text
      .split('\n')
      .map((line) => `<text:p>${escapeXml(line)}</text:p>`)
      .join('');
  }
  if (block.type === 'thematic_break') {
    return '<text:p/>';
  }
  if (block.type === 'blockquote') {
    return block.blocks.map((inner) => blockToOdfXml(inner, ctx)).join('');
  }
  if (block.type === 'list') {
    const items = block.items
      .map((item) => {
        const inner = item
          .map((sub) => {
            if (sub.type === 'paragraph' || sub.type === 'heading') {
              return `<text:p>${runsToOdfXml(sub.runs)}</text:p>`;
            }
            if (sub.type === 'list') return blockToOdfXml(sub, ctx);
            return '';
          })
          .join('');
        return `<text:list-item>${inner}</text:list-item>`;
      })
      .join('');
    const style = block.ordered ? 'ADPNum' : 'ADPBullet';
    return `<text:list text:style-name="${style}">${items}</text:list>`;
  }
  if (block.type === 'table') {
    ctx.tables++;
    const cols = Math.max(1, ...block.rows.map((r) => r.length));
    const rows = block.rows
      .map(
        (row) =>
          `<table:table-row>${row
            .map(
              (cell) =>
                `<table:table-cell office:value-type="string"><text:p>${escapeXml(cell.text)}</text:p></table:table-cell>`
            )
            .join('')}</table:table-row>`
      )
      .join('');
    return `<table:table table:name="Table${ctx.tables}"><table:table-column table:number-columns-repeated="${cols}"/>${rows}</table:table>`;
  }
  return '';
}

const CONTENT_STYLES = `<office:automatic-styles>
  <text:list-style style:name="ADPBullet">
    <text:list-level-style-bullet text:level="1" text:bullet-char="•"/>
    <text:list-level-style-bullet text:level="2" text:bullet-char="◦"/>
    <text:list-level-style-bullet text:level="3" text:bullet-char="▪"/>
  </text:list-style>
  <text:list-style style:name="ADPNum">
    <text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."/>
    <text:list-level-style-number text:level="2" style:num-format="1" style:num-suffix="."/>
    <text:list-level-style-number text:level="3" style:num-format="1" style:num-suffix="."/>
  </text:list-style>
  <style:style style:name="ADPB" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>
  <style:style style:name="ADPI" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>
  <style:style style:name="ADPBI" style:family="text"><style:text-properties fo:font-weight="bold" fo:font-style="italic"/></style:style>
</office:automatic-styles>`;

const ODF_NS = `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"`;

export async function writeOdtFromBlocks(blocks: CanonicalFlowBlock[]): Promise<Buffer> {
  const ctx: OdfCtx = { tables: 0 };
  const body = blocks.map((b) => blockToOdfXml(b, ctx)).join('');
  const zip = new JSZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' });
  zip.file(
    'content.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${ODF_NS} office:version="1.2">
  ${CONTENT_STYLES}
  <office:body><office:text>${body}</office:text></office:body>
</office:document-content>`
  );
  zip.file(
    'styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles ${ODF_NS} office:version="1.2">
  <office:styles/>
</office:document-styles>`
  );
  zip.file(
    'meta.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta ${ODF_NS} office:version="1.2"><office:meta/></office:document-meta>`
  );
  zip.file(
    'META-INF/manifest.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
  <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.text" manifest:full-path="/"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml"/>
</manifest:manifest>`
  );
  return zip.generateAsync({ type: 'nodebuffer' });
}
