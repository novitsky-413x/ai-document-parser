import JSZip from 'jszip';
import type { CanonicalFlowBlock, CanonicalInline } from '../ir/types.js';
import { codeStyleId } from '../utils/code-style.js';
import { QUOTE_STYLE_ID } from '../utils/quote-style.js';
import { escapeXml } from '../utils/xml.js';
import { sanitizeHref } from '../utils/url.js';

function runsToOdfXml(runs: CanonicalInline[]): string {
  return runs
    .map((r) => {
      let text = escapeXml(r.text);
      if (r.code) text = `<text:span text:style-name="ADPCodeInline">${text}</text:span>`;
      if (r.bold && r.italic) text = `<text:span text:style-name="ADPBI">${text}</text:span>`;
      else if (r.bold) text = `<text:span text:style-name="ADPB">${text}</text:span>`;
      else if (r.italic) text = `<text:span text:style-name="ADPI">${text}</text:span>`;
      if (r.strike) text = `<text:span text:style-name="ADPS">${text}</text:span>`;
      if (r.href) {
        const safe = sanitizeHref(r.href);
        if (safe) {
          text = `<text:a xlink:href="${escapeXml(safe)}">${text}</text:a>`;
        }
      }
      return text;
    })
    .join('');
}

type OdfCtx = { tables: number; codeStyles: Set<string>; hasQuote: boolean };

function blockToOdfXml(block: CanonicalFlowBlock, ctx: OdfCtx): string {
  if (block.type === 'heading') {
    return `<text:h text:outline-level="${block.level}">${runsToOdfXml(block.runs)}</text:h>`;
  }
  if (block.type === 'paragraph') {
    return `<text:p>${runsToOdfXml(block.runs)}</text:p>`;
  }
  if (block.type === 'code') {
    const style = codeStyleId(block.language);
    ctx.codeStyles.add(style);
    const inner = block.text
      .replace(/\r\n/g, '\n')
      .split('\n')
      .map((line) => escapeXml(line))
      .join('<text:line-break/>');
    return `<text:p text:style-name="${escapeXml(style)}">${inner}</text:p>`;
  }
  if (block.type === 'thematic_break') {
    return '<text:p/>';
  }
  if (block.type === 'blockquote') {
    ctx.hasQuote = true;
    return block.blocks
      .map((inner) => {
        if (inner.type === 'paragraph' || inner.type === 'heading') {
          return `<text:p text:style-name="${QUOTE_STYLE_ID}">${runsToOdfXml(inner.runs)}</text:p>`;
        }
        return blockToOdfXml(inner, ctx);
      })
      .join('');
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

function codeParagraphStyleXml(styleName: string): string {
  return `<style:style style:name="${escapeXml(styleName)}" style:family="paragraph">
    <style:paragraph-properties fo:background-color="#f5f5f5"/>
    <style:text-properties style:font-name="Courier New" fo:font-family="Courier New" fo:font-size="10pt"/>
  </style:style>`;
}

function buildContentStyles(codeStyles: Set<string>, hasQuote: boolean): string {
  const codeDefs = [...codeStyles].map(codeParagraphStyleXml).join('\n  ');
  const quoteDef = hasQuote
    ? `<style:style style:name="${QUOTE_STYLE_ID}" style:family="paragraph">
    <style:paragraph-properties fo:margin-left="0.5in"/>
    <style:text-properties fo:font-style="italic"/>
  </style:style>`
    : '';
  return `<office:automatic-styles>
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
  <style:style style:name="ADPS" style:family="text"><style:text-properties style:text-line-through-style="solid"/></style:style>
  <style:style style:name="ADPCodeInline" style:family="text"><style:text-properties style:font-name="Courier New" fo:font-family="Courier New" fo:font-size="10pt"/></style:style>
  ${quoteDef}
  ${codeDefs}
</office:automatic-styles>`;
}

const ODF_NS = `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
 xmlns:xlink="http://www.w3.org/1999/xlink"`;

export async function writeOdtFromBlocks(blocks: CanonicalFlowBlock[]): Promise<Buffer> {
  const ctx: OdfCtx = { tables: 0, codeStyles: new Set(['ADPCode']), hasQuote: false };
  const body = blocks.map((b) => blockToOdfXml(b, ctx)).join('');
  const zip = new JSZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' });
  zip.file(
    'content.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${ODF_NS} office:version="1.2">
  ${buildContentStyles(ctx.codeStyles, ctx.hasQuote)}
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
