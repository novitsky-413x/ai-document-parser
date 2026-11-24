import type {
  ContentBlock,
  DocxStructured,
  InlineRun,
  PptxStructured,
  SpreadsheetStructured,
} from '../types.js';
import type {
  CanonicalDocument,
  CanonicalFlowBlock,
  CanonicalInline,
  CanonicalPresentation,
  CanonicalWorkbook,
  ParsedCanonical,
} from './types.js';

function runsToPlain(runs: CanonicalInline[] | InlineRun[]): string {
  return runs.map((r) => r.text).join('');
}

function flattenFlowBlocks(blocks: CanonicalFlowBlock[]): string[] {
  const out: string[] = [];
  for (const b of blocks) {
    if (b.type === 'heading' || b.type === 'paragraph') {
      const t = runsToPlain(b.runs).trim();
      if (t && !/^\[\^[\w-]+\]:/.test(t)) out.push(t);
    } else if (b.type === 'list') {
      for (const item of b.items) out.push(...flattenFlowBlocks(item));
    } else if (b.type === 'table') {
      for (const row of b.rows) {
        const line = row.map((c) => c.text.trim()).join(' | ');
        if (line.replace(/\s*\|\s*/g, '').trim()) out.push(line);
      }
    } else if (b.type === 'code') {
      const t = b.text.trim();
      if (t) out.push(t);
    } else if (b.type === 'blockquote') {
      out.push(...flattenFlowBlocks(b.blocks));
    }
  }
  return out;
}

function listToContentBlocks(
  list: Extract<CanonicalFlowBlock, { type: 'list' }>,
  depth: number
): ContentBlock[] {
  const out: ContentBlock[] = [];
  for (const item of list.items) {
    const runs: InlineRun[] = [];
    for (const x of item) {
      if (x.type === 'paragraph' || x.type === 'heading') {
        if (runs.length) runs.push({ text: ' ' });
        runs.push(...x.runs);
      }
    }
    if (runs.length) {
      out.push({ type: 'list_item', level: depth, ordered: list.ordered, runs });
    }
    for (const x of item) {
      if (x.type === 'list') out.push(...listToContentBlocks(x, depth + 1));
      else if (x.type !== 'paragraph' && x.type !== 'heading') {
        out.push(...flowBlockToContentBlock(x));
      }
    }
  }
  return out;
}

function flowBlockToContentBlock(b: CanonicalFlowBlock): ContentBlock[] {
  switch (b.type) {
    case 'heading':
      return [{ type: 'heading', level: b.level, runs: b.runs }];
    case 'paragraph':
      return [{ type: 'paragraph', runs: b.runs }];
    case 'code':
      return [{ type: 'code', text: b.text, language: b.language }];
    case 'thematic_break':
      return [{ type: 'thematic_break' }];
    case 'list':
      return listToContentBlocks(b, 0);
    case 'table':
      return [
        {
          type: 'table',
          rows: b.rows.map((row) => row.map((c) => c.text)),
        },
      ];
    case 'blockquote':
      return [{ type: 'blockquote', blocks: b.blocks.flatMap(flowBlockToContentBlock) }];
    default:
      return [];
  }
}

export function flowToContentBlocks(blocks: CanonicalFlowBlock[]): ContentBlock[] {
  return blocks.flatMap(flowBlockToContentBlock);
}

export function workbookToSpreadsheetStructured(
  doc: CanonicalWorkbook
): SpreadsheetStructured {
  const sheets: SpreadsheetStructured['sheets'] = {};
  for (const s of doc.sheets) {
    sheets[s.name] = s.rows;
  }
  return { format: doc.format as 'xlsx' | 'xls' | 'ods', sheets };
}

export function presentationToPptxStructured(
  doc: CanonicalPresentation
): PptxStructured {
  const slides: string[][] = [];
  const blocks: ContentBlock[] = [];
  for (const slide of doc.slides) {
    const texts: string[] = [];
    if (slide.title?.length) texts.push(runsToPlain(slide.title));
    for (const b of slide.body) {
      if (b.type === 'paragraph' || b.type === 'heading') {
        texts.push(runsToPlain(b.runs));
      }
    }
    slides.push(texts);
    blocks.push({
      type: 'heading',
      level: 2,
      runs: [{ text: `Slide ${slide.index + 1}` }],
    });
    if (slide.title?.length) {
      blocks.push({ type: 'heading', level: 3, runs: slide.title });
    }
    blocks.push(...flowToContentBlocks(slide.body));
  }
  return { format: 'pptx', slides, blocks };
}

export function flowToDocxStructured(
  doc: Extract<CanonicalDocument, { kind: 'flow' }>
): DocxStructured {
  const blocks = flowToContentBlocks(doc.blocks);
  return {
    format: 'docx',
    paragraphs: flattenFlowBlocks(doc.blocks),
    blocks,
  };
}

export function canonicalToPlainText(doc: CanonicalDocument): string {
  if (doc.kind === 'flow') return flattenFlowBlocks(doc.blocks).join('\n\n');
  if (doc.kind === 'workbook') {
    return doc.sheets
      .map((s) => {
        const csv = s.rows.map((r) => r.map((c) => String(c ?? '')).join(',')).join('\n');
        return `## ${s.name}\n${csv}`;
      })
      .join('\n\n');
  }
  return doc.slides
    .map((s, i) => {
      const lines = [
        s.title ? runsToPlain(s.title) : '',
        ...s.body
          .filter((b) => b.type === 'paragraph' || b.type === 'heading')
          .map((b) => runsToPlain(b.runs)),
      ].filter(Boolean);
      return `## Slide ${i + 1}\n${lines.join('\n')}`;
    })
    .join('\n\n');
}

export function canonicalToStructured(doc: CanonicalDocument): Record<string, unknown> {
  if (doc.kind === 'flow') {
    const blocks = flowToContentBlocks(doc.blocks);
    return {
      format: doc.format,
      paragraphs: flattenFlowBlocks(doc.blocks),
      blocks,
      ...(doc.raw != null ? { raw: doc.raw } : {}),
    };
  }
  if (doc.kind === 'workbook') {
    return workbookToSpreadsheetStructured(doc) as unknown as Record<string, unknown>;
  }
  return presentationToPptxStructured(doc) as unknown as Record<string, unknown>;
}

type FlatListItem = { level: number; ordered?: boolean; runs: InlineRun[] };

/** Build a (possibly nested) canonical list from flat list items with levels. */
function buildNestedList(
  items: FlatListItem[],
  depth: number
): Extract<CanonicalFlowBlock, { type: 'list' }> {
  const list: Extract<CanonicalFlowBlock, { type: 'list' }> = {
    type: 'list',
    ordered: items[0]?.ordered,
    items: [],
  };
  let i = 0;
  while (i < items.length) {
    if (items[i].level <= depth) {
      list.items.push([{ type: 'paragraph', runs: items[i].runs }]);
      i++;
    } else {
      const sub: FlatListItem[] = [];
      while (i < items.length && items[i].level > depth) {
        sub.push(items[i]);
        i++;
      }
      const nested = buildNestedList(sub, depth + 1);
      if (list.items.length) list.items[list.items.length - 1].push(nested);
      else list.items.push([nested]);
    }
  }
  return list;
}

export function contentBlocksToFlow(blocks: ContentBlock[]): CanonicalFlowBlock[] {
  const out: CanonicalFlowBlock[] = [];
  let listBuf: FlatListItem[] = [];
  let listOrdered: boolean | undefined;

  const flushList = () => {
    if (listBuf.length) {
      out.push(buildNestedList(listBuf, 0));
      listBuf = [];
      listOrdered = undefined;
    }
  };

  for (const b of blocks) {
    if (b.type === 'list_item') {
      // A new top-level item with a different ordered flag starts a new list;
      // nested items keep the surrounding list.
      if (listBuf.length && b.level === 0 && listOrdered !== b.ordered) {
        flushList();
      }
      if (!listBuf.length) listOrdered = b.ordered;
      listBuf.push({ level: Math.max(0, b.level), ordered: b.ordered, runs: b.runs });
      continue;
    }
    flushList();
    if (b.type === 'heading')
      out.push({
        type: 'heading',
        level: b.level,
        runs: b.runs,
        ...(b.layout ? { layout: b.layout } : {}),
      });
    else if (b.type === 'paragraph')
      out.push({ type: 'paragraph', runs: b.runs, ...(b.layout ? { layout: b.layout } : {}) });
    else if (b.type === 'table')
      out.push({
        type: 'table',
        headerRows: 1,
        rows: b.rows.map((row) => row.map((text) => ({ text }))),
      });
    else if (b.type === 'thematic_break') out.push({ type: 'thematic_break' });
    else if (b.type === 'code') out.push({ type: 'code', text: b.text, language: b.language });
    else if (b.type === 'blockquote')
      out.push({ type: 'blockquote', blocks: contentBlocksToFlow(b.blocks) });
  }
  flushList();
  return out;
}

export function buildParsedCanonical(
  doc: CanonicalDocument,
  warnings: ParsedCanonical['warnings'] = []
): ParsedCanonical {
  return {
    canonical: doc,
    warnings,
    structured: canonicalToStructured(doc),
    plainText: canonicalToPlainText(doc),
  };
}
