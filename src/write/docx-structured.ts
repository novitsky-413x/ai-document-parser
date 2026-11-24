import {
  AlignmentType,
  BorderStyle,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  LevelFormat,
  LineRuleType,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import type { CanonicalFlowBlock, CanonicalInline, CanonicalTableCell } from '../ir/types.js';
import type { BlockLayout, DocumentLayout } from '../types.js';
import { codeStyleId } from '../utils/code-style.js';
import { QUOTE_STYLE_ID } from '../utils/quote-style.js';

const NUMBERING_REF = 'adp-ordered';
const MAX_LIST_LEVEL = 8;
const CODE_FONT = 'Consolas';
/** Usable page width in twips (~A4 / Letter with ~1" margins). */
const TABLE_WIDTH_DXA = 9026;

const TABLE_BORDER = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const TABLE_BORDERS = {
  top: TABLE_BORDER,
  bottom: TABLE_BORDER,
  left: TABLE_BORDER,
  right: TABLE_BORDER,
};

/** Points → twentieths of a point (twips). */
function twips(pt: number): number {
  return Math.round(pt * 20);
}

/**
 * CSS line-height is a multiple of the font size; Word "multiple" spacing is a
 * multiple of the font's natural line (~1.15 × size for common fonts).
 */
const CSS_NORMAL_LINE = 1.15;

function lineSpacing(lineHeight: number): { line: number; lineRule: (typeof LineRuleType)['AUTO'] } {
  return { line: Math.max(1, Math.round((lineHeight / CSS_NORMAL_LINE) * 240)), lineRule: LineRuleType.AUTO };
}

const ALIGN = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
} as const;

/** Paragraph options from block layout (alignment, spacing, indents, page break). */
function paragraphLayout(layout: BlockLayout | undefined): Record<string, unknown> {
  if (!layout) return {};
  const out: Record<string, unknown> = {};
  if (layout.align) out.alignment = ALIGN[layout.align];
  const spacing: Record<string, unknown> = {};
  if (layout.spaceBeforePt != null) spacing.before = twips(layout.spaceBeforePt);
  if (layout.spaceAfterPt != null) spacing.after = twips(layout.spaceAfterPt);
  if (layout.lineHeight != null) Object.assign(spacing, lineSpacing(layout.lineHeight));
  if (Object.keys(spacing).length) out.spacing = spacing;
  const indent: Record<string, number> = {};
  if (layout.indentLeftPt) indent.left = twips(layout.indentLeftPt);
  if (layout.indentRightPt) indent.right = twips(layout.indentRightPt);
  if (layout.firstLineIndentPt) {
    if (layout.firstLineIndentPt > 0) indent.firstLine = twips(layout.firstLineIndentPt);
    else indent.hanging = twips(-layout.firstLineIndentPt);
  }
  if (Object.keys(indent).length) out.indent = indent;
  if (layout.pageBreakBefore) out.pageBreakBefore = true;
  return out;
}

function textRunOptions(r: CanonicalInline): Record<string, unknown> {
  return {
    bold: r.bold,
    italics: r.italic,
    strike: r.strike,
    ...(r.underline ? { underline: {} } : {}),
    ...(r.sizePt ? { size: Math.round(r.sizePt * 2) } : {}),
    ...(r.code ? { font: CODE_FONT } : r.font ? { font: r.font } : {}),
    ...(r.color ? { color: r.color } : {}),
    ...(r.letterSpacingPt ? { characterSpacing: twips(r.letterSpacingPt) } : {}),
    ...(r.href ? { style: 'Hyperlink' } : {}),
  };
}

/** `\n` in run text is a line break (`w:br`), matching how DOCX is read. */
function textRuns(r: CanonicalInline): TextRun[] {
  const parts = r.text.replace(/\r\n/g, '\n').split('\n');
  return parts.map(
    (part, i) =>
      new TextRun({
        ...textRunOptions(r),
        text: part,
        ...(i > 0 ? { break: 1 } : {}),
      })
  );
}

function runsToChildren(runs: CanonicalInline[]): Array<TextRun | ExternalHyperlink> {
  return runs.flatMap((r): Array<TextRun | ExternalHyperlink> => {
    const runsOut = textRuns(r);
    if (r.href) {
      return [new ExternalHyperlink({ children: runsOut, link: r.href })];
    }
    return runsOut;
  });
}

function listToParagraphs(
  list: Extract<CanonicalFlowBlock, { type: 'list' }>,
  depth: number,
  out: Array<Paragraph | Table>
): void {
  const level = Math.min(depth, MAX_LIST_LEVEL);
  for (const item of list.items) {
    for (const sub of item) {
      if (sub.type === 'paragraph' || sub.type === 'heading') {
        out.push(
          new Paragraph({
            ...(list.ordered
              ? { numbering: { reference: NUMBERING_REF, level } }
              : { bullet: { level } }),
            children: runsToChildren(sub.runs),
          })
        );
      } else if (sub.type === 'list') {
        listToParagraphs(sub, depth + 1, out);
      }
    }
  }
}

export type DocxWriteOptions = {
  /** Document defaults (font, size, color, line height) and page size/margins. */
  layout?: DocumentLayout;
};

export async function writeDocxFromBlocks(
  blocks: CanonicalFlowBlock[],
  options: DocxWriteOptions = {}
): Promise<Buffer> {
  const docLayout = options.layout ?? {};
  const children: Array<Paragraph | Table> = [];

  for (const b of blocks) {
    if (b.type === 'heading') {
      const level = Math.min(6, Math.max(1, b.level));
      const map = [
        HeadingLevel.HEADING_1,
        HeadingLevel.HEADING_2,
        HeadingLevel.HEADING_3,
        HeadingLevel.HEADING_4,
        HeadingLevel.HEADING_5,
        HeadingLevel.HEADING_6,
      ] as const;
      children.push(
        new Paragraph({
          heading: map[level - 1],
          ...paragraphLayout(b.layout),
          children: runsToChildren(b.runs),
        })
      );
    } else if (b.type === 'paragraph') {
      children.push(
        new Paragraph({ ...paragraphLayout(b.layout), children: runsToChildren(b.runs) })
      );
    } else if (b.type === 'code') {
      children.push(
        new Paragraph({
          style: codeStyleId(b.language),
          children: codeLineRuns(b.text),
        })
      );
    } else if (b.type === 'thematic_break') {
      children.push(
        new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, space: 1 } },
        })
      );
    } else if (b.type === 'blockquote') {
      for (const inner of b.blocks) {
        if (inner.type === 'paragraph' || inner.type === 'heading') {
          children.push(
            new Paragraph({
              style: QUOTE_STYLE_ID,
              ...paragraphLayout(inner.layout),
              indent: {
                left: 720 + (inner.layout?.indentLeftPt ? twips(inner.layout.indentLeftPt) : 0),
              },
              children: runsToChildren(inner.runs),
            })
          );
        }
      }
    } else if (b.type === 'table') {
      children.push(tableFromCanonical(b));
    } else if (b.type === 'list') {
      listToParagraphs(b, 0, children);
    }
  }

  const codeStyles = collectCodeStyles(blocks);
  const defaultRun: Record<string, unknown> = {};
  if (docLayout.font) defaultRun.font = docLayout.font;
  if (docLayout.sizePt) defaultRun.size = Math.round(docLayout.sizePt * 2);
  if (docLayout.color) defaultRun.color = docLayout.color;
  const defaultParagraph: Record<string, unknown> = {};
  if (docLayout.lineHeight != null) defaultParagraph.spacing = lineSpacing(docLayout.lineHeight);
  const hasDefaults = Object.keys(defaultRun).length || Object.keys(defaultParagraph).length;

  const doc = new Document({
    styles: {
      ...(hasDefaults
        ? {
            default: {
              document: {
                ...(Object.keys(defaultRun).length ? { run: defaultRun } : {}),
                ...(Object.keys(defaultParagraph).length ? { paragraph: defaultParagraph } : {}),
              },
            },
          }
        : {}),
      // Cast: docx typings reject mixed style option shapes (code vs quote).
      paragraphStyles: [
        ...codeStyles.map((id) => ({
          id,
          name: id === 'ADPCode' ? 'ADP Code' : id.replace(/^ADPCode_/, 'ADP Code '),
          basedOn: 'Normal',
          next: 'Normal',
          quickStyle: true,
          run: { font: CODE_FONT, size: 20 },
          paragraph: { shading: { type: 'clear' as const, fill: 'F5F5F5' } },
        })),
        {
          id: QUOTE_STYLE_ID,
          name: 'ADP Quote',
          basedOn: 'Normal',
          next: 'Normal',
          quickStyle: true,
          paragraph: { indent: { left: 720 } },
        },
      ] as never,
    },
    numbering: {
      config: [
        {
          reference: NUMBERING_REF,
          levels: Array.from({ length: MAX_LIST_LEVEL + 1 }, (_, level) => ({
            level,
            format: LevelFormat.DECIMAL,
            text: `%${level + 1}.`,
            indent: { left: 720 * (level + 1), hanging: 360 },
          })),
        },
      ],
    },
    sections: [
      {
        ...sectionProperties(docLayout),
        children: children.length ? children : [new Paragraph(' ')],
      },
    ],
  });
  return Packer.toBuffer(doc);
}

function sectionProperties(layout: DocumentLayout): Record<string, unknown> {
  const p = layout.page;
  if (!p) return {};
  const page: Record<string, unknown> = {};
  if (p.widthPt && p.heightPt) {
    // docx swaps width/height itself when orientation is landscape.
    const landscape = p.widthPt > p.heightPt;
    page.size = {
      width: twips(Math.min(p.widthPt, p.heightPt)),
      height: twips(Math.max(p.widthPt, p.heightPt)),
      ...(landscape ? { orientation: 'landscape' } : {}),
    };
  }
  const margin: Record<string, number> = {};
  if (p.marginTopPt != null) margin.top = twips(p.marginTopPt);
  if (p.marginRightPt != null) margin.right = twips(p.marginRightPt);
  if (p.marginBottomPt != null) margin.bottom = twips(p.marginBottomPt);
  if (p.marginLeftPt != null) margin.left = twips(p.marginLeftPt);
  if (Object.keys(margin).length) page.margin = margin;
  return Object.keys(page).length ? { properties: { page } } : {};
}

function collectCodeStyles(blocks: CanonicalFlowBlock[]): string[] {
  const ids = new Set<string>(['ADPCode']);
  const walk = (list: CanonicalFlowBlock[]) => {
    for (const b of list) {
      if (b.type === 'code') ids.add(codeStyleId(b.language));
      else if (b.type === 'blockquote') walk(b.blocks);
      else if (b.type === 'list') {
        for (const item of b.items) walk(item);
      }
    }
  };
  walk(blocks);
  return [...ids];
}

function codeLineRuns(text: string): TextRun[] {
  const lines = text.replace(/\r\n/g, '\n').split('\n');
  const runs: TextRun[] = [];
  for (let i = 0; i < lines.length; i++) {
    if (i > 0) runs.push(new TextRun({ break: 1 }));
    runs.push(new TextRun({ text: lines[i], font: CODE_FONT }));
  }
  if (!runs.length) runs.push(new TextRun({ text: ' ', font: CODE_FONT }));
  return runs;
}

function tableFromCanonical(
  block: Extract<CanonicalFlowBlock, { type: 'table' }>
): Table {
  const colCount = Math.max(1, ...block.rows.map((r) => r.length));
  const colWidth = Math.max(1, Math.floor(TABLE_WIDTH_DXA / colCount));
  const columnWidths = Array.from({ length: colCount }, () => colWidth);

  return new Table({
    width: { size: TABLE_WIDTH_DXA, type: WidthType.DXA },
    columnWidths,
    rows: block.rows.map((row) => {
      const cells = padRow(row, colCount);
      return new TableRow({
        children: cells.map(
          (cell) =>
            new TableCell({
              borders: TABLE_BORDERS,
              width: { size: colWidth, type: WidthType.DXA },
              children: cellParagraphs(cell),
            })
        ),
      });
    }),
  });
}

function padRow(row: CanonicalTableCell[], colCount: number): CanonicalTableCell[] {
  const cells = row.slice(0, colCount);
  while (cells.length < colCount) cells.push({ text: '' });
  return cells;
}

function splitRunsOnNewlines(runs: CanonicalInline[]): CanonicalInline[][] {
  const lines: CanonicalInline[][] = [[]];
  for (const run of runs) {
    const parts = (run.text ?? '').replace(/\r\n/g, '\n').split('\n');
    parts.forEach((part, i) => {
      if (i > 0) lines.push([]);
      if (part || lines[lines.length - 1].length === 0) {
        lines[lines.length - 1].push({ ...run, text: part });
      }
    });
  }
  return lines.length ? lines : [[]];
}

function cellParagraphs(cell: CanonicalTableCell): Paragraph[] {
  if (cell.runs?.length) {
    return splitRunsOnNewlines(cell.runs).map(
      (runs) =>
        new Paragraph({
          children: runsToChildren(runs.length ? runs : [{ text: ' ' }]),
        })
    );
  }
  const lines = (cell.text || ' ').replace(/\r\n/g, '\n').split('\n');
  return lines.map(
    (line) =>
      new Paragraph({
        children: [new TextRun({ text: line.length ? line : ' ' })],
      })
  );
}
