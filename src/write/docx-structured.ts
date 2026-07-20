import {
  BorderStyle,
  Document,
  ExternalHyperlink,
  HeadingLevel,
  LevelFormat,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import type { CanonicalFlowBlock, CanonicalInline, CanonicalTableCell } from '../ir/types.js';
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

function runsToChildren(runs: CanonicalInline[]): Array<TextRun | ExternalHyperlink> {
  return runs.map((r) => {
    const text = new TextRun({
      text: r.text,
      bold: r.bold,
      italics: r.italic,
      strike: r.strike,
      ...(r.code ? { font: CODE_FONT } : {}),
      ...(r.href ? { style: 'Hyperlink' } : {}),
    });
    if (r.href) {
      return new ExternalHyperlink({ children: [text], link: r.href });
    }
    return text;
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

export async function writeDocxFromBlocks(
  blocks: CanonicalFlowBlock[]
): Promise<Buffer> {
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
          children: runsToChildren(b.runs),
        })
      );
    } else if (b.type === 'paragraph') {
      children.push(new Paragraph({ children: runsToChildren(b.runs) }));
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
          border: { bottom: { style: 'single', size: 6, space: 1 } },
        })
      );
    } else if (b.type === 'blockquote') {
      for (const inner of b.blocks) {
        if (inner.type === 'paragraph' || inner.type === 'heading') {
          children.push(
            new Paragraph({
              style: QUOTE_STYLE_ID,
              indent: { left: 720 },
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
  const doc = new Document({
    styles: {
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
    sections: [{ children: children.length ? children : [new Paragraph(' ')] }],
  });
  return Packer.toBuffer(doc);
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
              children: [
                new Paragraph({
                  children: cellChildren(cell),
                }),
              ],
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

function cellChildren(cell: CanonicalTableCell): Array<TextRun | ExternalHyperlink> {
  if (cell.runs?.length) return runsToChildren(cell.runs);
  return [new TextRun({ text: cell.text || ' ' })];
}
