import {
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
} from 'docx';
import type { CanonicalFlowBlock, CanonicalInline } from '../ir/types.js';

const NUMBERING_REF = 'adp-ordered';
const MAX_LIST_LEVEL = 8;

function runsToChildren(runs: CanonicalInline[]): Array<TextRun | ExternalHyperlink> {
  return runs.map((r) => {
    const text = new TextRun({
      text: r.text,
      bold: r.bold,
      italics: r.italic,
      strike: r.strike,
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
      for (const line of b.text.split('\n')) {
        children.push(
          new Paragraph({
            children: [new TextRun({ text: line, font: 'Consolas' })],
          })
        );
      }
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
              indent: { left: 720 },
              children: runsToChildren(inner.runs),
            })
          );
        }
      }
    } else if (b.type === 'table') {
      children.push(
        new Table({
          rows: b.rows.map(
            (row) =>
              new TableRow({
                children: row.map(
                  (cell) =>
                    new TableCell({
                      children: [
                        new Paragraph({
                          children: [new TextRun(cell.text || ' ')],
                        }),
                      ],
                    })
                ),
              })
          ),
        })
      );
    } else if (b.type === 'list') {
      listToParagraphs(b, 0, children);
    }
  }

  const doc = new Document({
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
