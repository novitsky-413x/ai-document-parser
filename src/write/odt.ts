import JSZip from 'jszip';
import type { CanonicalFlowBlock, CanonicalInline } from '../ir/types.js';
import type { BlockLayout, DocumentLayout } from '../types.js';
import { codeStyleId } from '../utils/code-style.js';
import { QUOTE_STYLE_ID } from '../utils/quote-style.js';
import { escapeXml } from '../utils/xml.js';
import { sanitizeHref } from '../utils/url.js';

function pt(n: number): string {
  return `${Math.round(n * 100) / 100}pt`;
}

/** Text-properties attributes for layout run props (size/font/color/spacing/underline). */
function runLayoutAttrs(r: CanonicalInline): string {
  const a: string[] = [];
  if (r.sizePt) a.push(`fo:font-size="${pt(r.sizePt)}"`);
  if (r.font && !r.code) a.push(`style:font-name="${escapeXml(r.font)}" fo:font-family="${escapeXml(r.font)}"`);
  if (r.color) a.push(`fo:color="#${escapeXml(r.color)}"`);
  if (r.letterSpacingPt) a.push(`fo:letter-spacing="${pt(r.letterSpacingPt)}"`);
  if (r.underline) a.push('style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"');
  return a.join(' ');
}

function textWithBreaks(text: string): string {
  return text
    .replace(/\r\n/g, '\n')
    .split('\n')
    .map((part) => escapeXml(part))
    .join('<text:line-break/>');
}

function runsToOdfXml(runs: CanonicalInline[], ctx?: OdfCtx): string {
  return runs
    .map((r) => {
      let text = textWithBreaks(r.text);
      // Layout span innermost: readers match the first closing span tag.
      const extra = ctx ? runLayoutAttrs(r) : '';
      if (extra && ctx) {
        text = `<text:span text:style-name="${ctx.styleFor('text', extra)}">${text}</text:span>`;
      }
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

const ODF_ALIGN = { left: 'start', center: 'center', right: 'end', justify: 'justify' } as const;

/** Paragraph-properties attributes for block layout. */
function paragraphLayoutAttrs(l: BlockLayout | undefined): string {
  if (!l) return '';
  const a: string[] = [];
  if (l.align) a.push(`fo:text-align="${ODF_ALIGN[l.align]}"`);
  if (l.spaceBeforePt != null) a.push(`fo:margin-top="${pt(l.spaceBeforePt)}"`);
  if (l.spaceAfterPt != null) a.push(`fo:margin-bottom="${pt(l.spaceAfterPt)}"`);
  if (l.lineHeight != null) a.push(`fo:line-height="${Math.round((l.lineHeight / 1.15) * 100)}%"`);
  if (l.indentLeftPt) a.push(`fo:margin-left="${pt(l.indentLeftPt)}"`);
  if (l.indentRightPt) a.push(`fo:margin-right="${pt(l.indentRightPt)}"`);
  if (l.firstLineIndentPt) a.push(`fo:text-indent="${pt(l.firstLineIndentPt)}"`);
  if (l.pageBreakBefore) a.push('fo:break-before="page"');
  return a.join(' ');
}

function paragraphStyleAttr(l: BlockLayout | undefined, ctx: OdfCtx, parent?: string): string {
  const attrs = paragraphLayoutAttrs(l);
  if (!attrs) return parent ? ` text:style-name="${escapeXml(parent)}"` : '';
  return ` text:style-name="${ctx.styleFor('paragraph', attrs, parent)}"`;
}

type OdfCtx = {
  tables: number;
  codeStyles: Set<string>;
  hasQuote: boolean;
  /** Automatic styles for layout (paragraph ADPPn / text ADPTn), deduped by attrs. */
  autoStyles: Map<string, { name: string; family: 'paragraph' | 'text'; attrs: string; parent?: string }>;
  styleFor: (family: 'paragraph' | 'text', attrs: string, parent?: string) => string;
};

function newOdfCtx(): OdfCtx {
  const ctx: OdfCtx = {
    tables: 0,
    codeStyles: new Set(['ADPCode']),
    hasQuote: false,
    autoStyles: new Map(),
    styleFor(family, attrs, parent) {
      const key = `${family}|${parent ?? ''}|${attrs}`;
      let st = ctx.autoStyles.get(key);
      if (!st) {
        const n = [...ctx.autoStyles.values()].filter((x) => x.family === family).length + 1;
        st = { name: `${family === 'paragraph' ? 'ADPP' : 'ADPT'}${n}`, family, attrs, parent };
        ctx.autoStyles.set(key, st);
      }
      return st.name;
    },
  };
  return ctx;
}

function autoStylesXml(ctx: OdfCtx): string {
  return [...ctx.autoStyles.values()]
    .map((st) =>
      st.family === 'paragraph'
        ? `<style:style style:name="${st.name}" style:family="paragraph"${st.parent ? ` style:parent-style-name="${escapeXml(st.parent)}"` : ''}><style:paragraph-properties ${st.attrs}/></style:style>`
        : `<style:style style:name="${st.name}" style:family="text"><style:text-properties ${st.attrs}/></style:style>`
    )
    .join('\n  ');
}

function blockToOdfXml(block: CanonicalFlowBlock, ctx: OdfCtx): string {
  if (block.type === 'heading') {
    return `<text:h${paragraphStyleAttr(block.layout, ctx)} text:outline-level="${block.level}">${runsToOdfXml(block.runs, ctx)}</text:h>`;
  }
  if (block.type === 'paragraph') {
    return `<text:p${paragraphStyleAttr(block.layout, ctx)}>${runsToOdfXml(block.runs, ctx)}</text:p>`;
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
          return `<text:p${paragraphStyleAttr(inner.layout, ctx, QUOTE_STYLE_ID)}>${runsToOdfXml(inner.runs, ctx)}</text:p>`;
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
              return `<text:p>${runsToOdfXml(sub.runs, ctx)}</text:p>`;
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

function buildContentStyles(codeStyles: Set<string>, hasQuote: boolean, ctx: OdfCtx): string {
  const autoDefs = autoStylesXml(ctx);
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
  ${autoDefs}
</office:automatic-styles>`;
}

const ODF_NS = `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
 xmlns:xlink="http://www.w3.org/1999/xlink"
 xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"`;

export type OdtWriteOptions = {
  /** Document defaults (font, size, color, line height) and page size/margins. */
  layout?: DocumentLayout;
};

function documentStylesXml(layout: DocumentLayout): string {
  const text: string[] = [];
  if (layout.font) text.push(`style:font-name="${escapeXml(layout.font)}" fo:font-family="${escapeXml(layout.font)}"`);
  if (layout.sizePt) text.push(`fo:font-size="${pt(layout.sizePt)}"`);
  if (layout.color) text.push(`fo:color="#${escapeXml(layout.color)}"`);
  const para =
    layout.lineHeight != null ? `<style:paragraph-properties fo:line-height="${Math.round((layout.lineHeight / 1.15) * 100)}%"/>` : '';
  const fontDecl = layout.font
    ? `<office:font-face-decls><style:font-face style:name="${escapeXml(layout.font)}" svg:font-family="${escapeXml(layout.font)}"/></office:font-face-decls>`
    : '';
  const defaults =
    text.length || para
      ? `<office:styles><style:default-style style:family="paragraph">${para}${text.length ? `<style:text-properties ${text.join(' ')}/>` : ''}</style:default-style></office:styles>`
      : '<office:styles/>';
  const p = layout.page;
  let pageXml = '';
  if (p) {
    const a: string[] = [];
    if (p.widthPt && p.heightPt) {
      a.push(`fo:page-width="${pt(p.widthPt)}" fo:page-height="${pt(p.heightPt)}"`);
      a.push(`style:print-orientation="${p.widthPt > p.heightPt ? 'landscape' : 'portrait'}"`);
    }
    if (p.marginTopPt != null) a.push(`fo:margin-top="${pt(p.marginTopPt)}"`);
    if (p.marginRightPt != null) a.push(`fo:margin-right="${pt(p.marginRightPt)}"`);
    if (p.marginBottomPt != null) a.push(`fo:margin-bottom="${pt(p.marginBottomPt)}"`);
    if (p.marginLeftPt != null) a.push(`fo:margin-left="${pt(p.marginLeftPt)}"`);
    if (a.length) {
      pageXml = `<office:automatic-styles><style:page-layout style:name="ADPPage"><style:page-layout-properties ${a.join(' ')}/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="ADPPage"/></office:master-styles>`;
    }
  }
  return `${fontDecl}${defaults}${pageXml}`;
}

export async function writeOdtFromBlocks(
  blocks: CanonicalFlowBlock[],
  options: OdtWriteOptions = {}
): Promise<Buffer> {
  const ctx = newOdfCtx();
  const body = blocks.map((b) => blockToOdfXml(b, ctx)).join('');
  const zip = new JSZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' });
  zip.file(
    'content.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${ODF_NS} office:version="1.2">
  ${options.layout?.font ? `<office:font-face-decls><style:font-face style:name="${escapeXml(options.layout.font)}" svg:font-family="${escapeXml(options.layout.font)}"/></office:font-face-decls>` : ''}
  ${buildContentStyles(ctx.codeStyles, ctx.hasQuote, ctx)}
  <office:body><office:text>${body}</office:text></office:body>
</office:document-content>`
  );
  zip.file(
    'styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles ${ODF_NS} office:version="1.2">
  ${documentStylesXml(options.layout ?? {})}
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
