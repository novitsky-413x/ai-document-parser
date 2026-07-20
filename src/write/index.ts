import path from 'node:path';
import { DocumentError } from '../utils/errors.js';
import type {
  ContentBlock,
  DocumentWarning,
  WriteDocumentInput,
  WriteMode,
  WriteResult,
} from '../types.js';
import { writeDocx } from './docx.js';
import { writePdf } from './pdf.js';
import { writeSpreadsheet } from './spreadsheet.js';
import { writeDocxFromBlocks } from './docx-structured.js';
import { writeOdtFromBlocks } from './odt.js';
import { emitRtfFromFlow } from '../parse/rtf.js';
import { emitEpubFromFlow } from '../parse/epub.js';
import { contentBlocksToFlow } from '../ir/projections.js';
import { htmlToBlocks, markdownToBlocks } from '../parse/html-blocks.js';
import { ensureParentDir, writeFileAtomic } from '../utils/paths.js';

type FlowWriteType = 'docx' | 'odt' | 'rtf' | 'epub';

type ResolvedWrite = {
  blocks: ContentBlock[];
  mode: WriteMode;
  warnings: DocumentWarning[];
  paragraphCount?: number;
};

const PLAIN_MARKUP_HINT =
  'lines look like lists/tables/headings; use blocks or format:\'markdown\' (or \'html\') for structured emit';

/**
 * Heuristic: plain `content` that looks like markdown structure (agents often
 * put `- item` / `| a | b |` here by mistake).
 */
export function looksLikeIgnoredMarkup(lines: string[]): boolean {
  return lines.some(
    (line) =>
      /^\s{0,3}#{1,6}\s+\S/.test(line) ||
      /^\s*[-*+]\s+\S/.test(line) ||
      /^\s*\d+\.\s+\S/.test(line) ||
      /^\s*\|.+\|/.test(line)
  );
}

function countBlocks(
  blocks: ContentBlock[]
): Partial<Record<ContentBlock['type'], number>> {
  const counts: Partial<Record<ContentBlock['type'], number>> = {};
  for (const b of blocks) {
    counts[b.type] = (counts[b.type] ?? 0) + 1;
  }
  return counts;
}

function paragraphsFromContent(content: string | string[]): string[] {
  return Array.isArray(content) ? content : content.split(/\r?\n/);
}

/**
 * Resolve flow write input to ContentBlocks + mode.
 * Mutually exclusive: plain content | blocks | format+content.
 */
export function resolveFlowWriteInput(
  input: Extract<WriteDocumentInput, { type: FlowWriteType }>
): ResolvedWrite {
  const hasBlocks = 'blocks' in input && Array.isArray(input.blocks);
  const hasFormat = 'format' in input && input.format != null;
  const hasContent = 'content' in input && input.content !== undefined;

  if (hasBlocks && (hasFormat || hasContent)) {
    throw new DocumentError(
      'INVALID_ARGUMENT',
      'writeDocument: pass exactly one of content (plain), blocks, or format+content — not a mix'
    );
  }

  if (hasBlocks) {
    return {
      blocks: (input as { blocks: ContentBlock[] }).blocks,
      mode: 'blocks',
      warnings: [],
    };
  }

  if (hasFormat) {
    const fmt = (input as { format: 'markdown' | 'html'; content: string }).format;
    const content = (input as { content: string }).content;
    if (typeof content !== 'string') {
      throw new DocumentError(
        'INVALID_ARGUMENT',
        'writeDocument: format markdown/html requires content: string'
      );
    }
    if (fmt === 'markdown') {
      return {
        blocks: markdownToBlocks(content),
        mode: 'markdown',
        warnings: [],
      };
    }
    if (fmt === 'html') {
      return {
        blocks: htmlToBlocks(content),
        mode: 'html',
        warnings: [],
      };
    }
    throw new DocumentError(
      'INVALID_ARGUMENT',
      `writeDocument: unsupported format '${String(fmt)}' (use 'markdown' or 'html')`
    );
  }

  if (!hasContent) {
    throw new DocumentError(
      'INVALID_ARGUMENT',
      'writeDocument: provide content, blocks, or format+content'
    );
  }

  const paragraphs = paragraphsFromContent(
    (input as { content: string | string[] }).content
  );
  const warnings: DocumentWarning[] = [];
  if (looksLikeIgnoredMarkup(paragraphs)) {
    warnings.push({
      code: 'PLAIN_WRITE_IGNORED_MARKUP',
      message: `PLAIN_WRITE_IGNORED_MARKUP: ${PLAIN_MARKUP_HINT}`,
    });
  }

  return {
    blocks: paragraphs.map((p) => ({
      type: 'paragraph' as const,
      runs: [{ text: p }],
    })),
    mode: 'plain',
    warnings,
    paragraphCount: paragraphs.length,
  };
}

function writeResult(
  filePath: string,
  message: string,
  resolved: ResolvedWrite
): WriteResult {
  return {
    path: filePath,
    message,
    summary: {
      mode: resolved.mode,
      blockCounts: countBlocks(resolved.blocks),
      ...(resolved.paragraphCount !== undefined
        ? { paragraphCount: resolved.paragraphCount }
        : {}),
    },
    ...(resolved.warnings.length > 0 ? { warnings: resolved.warnings } : {}),
  };
}

/**
 * Create a new document from scratch.
 *
 * Flow formats (docx / odt / rtf / epub):
 * - `{ content }` — plain paragraphs only (markdown markers are literal text)
 * - `{ blocks }` — structured ContentBlocks
 * - `{ format: 'markdown' | 'html', content }` — parse markup then emit rich structure
 */
export async function writeDocument(
  filePath: string,
  input: WriteDocumentInput
): Promise<WriteResult> {
  const ext = path.extname(filePath).toLowerCase();

  if (input.type === 'docx') {
    if (ext && ext !== '.docx') {
      throw new DocumentError(
        'INVALID_ARGUMENT',
        `Expected .docx path for type docx, got ${ext}`
      );
    }
    const resolved = resolveFlowWriteInput(input);
    if (resolved.mode === 'plain') {
      const plainContent = (input as { content: string | string[] }).content;
      const result = await writeDocx(filePath, plainContent);
      return writeResult(filePath, result.message, resolved);
    }
    const flow = contentBlocksToFlow(resolved.blocks);
    const buf = await writeDocxFromBlocks(flow);
    await ensureParentDir(filePath);
    await writeFileAtomic(filePath, buf);
    return writeResult(
      filePath,
      `Created DOCX: ${filePath} (${resolved.mode} mode)`,
      resolved
    );
  }

  if (input.type === 'xlsx' || input.type === 'xls') {
    const expected = input.type === 'xls' ? '.xls' : '.xlsx';
    if (ext && ext !== expected) {
      throw new DocumentError(
        'INVALID_ARGUMENT',
        `Expected ${expected} path for type ${input.type}, got ${ext}`
      );
    }
    const result = await writeSpreadsheet(filePath, input.sheets);
    return {
      path: filePath,
      message: result.message,
      summary: { mode: 'plain' },
    };
  }

  if (input.type === 'pdf') {
    if (ext && ext !== '.pdf') {
      throw new DocumentError(
        'INVALID_ARGUMENT',
        `Expected .pdf path for type pdf, got ${ext}`
      );
    }
    const result = await writePdf(filePath, input.content, input.title);
    return {
      path: filePath,
      message: result.message,
      summary: { mode: 'plain' },
    };
  }

  if (input.type === 'odt' || input.type === 'rtf' || input.type === 'epub') {
    const expected = `.${input.type}`;
    if (ext && ext !== expected) {
      throw new DocumentError(
        'INVALID_ARGUMENT',
        `Expected ${expected} path for type ${input.type}, got ${ext}`
      );
    }
    const resolved = resolveFlowWriteInput(input);
    const flow = contentBlocksToFlow(resolved.blocks);
    await ensureParentDir(filePath);
    if (input.type === 'odt') {
      await writeFileAtomic(filePath, await writeOdtFromBlocks(flow));
      return writeResult(filePath, `Created ODT: ${filePath}`, resolved);
    }
    if (input.type === 'rtf') {
      await writeFileAtomic(filePath, emitRtfFromFlow(flow));
      return writeResult(filePath, `Created RTF: ${filePath}`, resolved);
    }
    const title =
      'title' in input && typeof input.title === 'string'
        ? input.title
        : 'Document';
    await writeFileAtomic(filePath, await emitEpubFromFlow(flow, title));
    return writeResult(filePath, `Created EPUB: ${filePath}`, resolved);
  }

  throw new DocumentError(
    'INVALID_ARGUMENT',
    `Unsupported write type: ${(input as { type: string }).type}`
  );
}

export { writeDocx } from './docx.js';
export { writePdf } from './pdf.js';
export { writeSpreadsheet } from './spreadsheet.js';
export { writeDocxFromBlocks } from './docx-structured.js';
export { writeOdtFromBlocks } from './odt.js';
