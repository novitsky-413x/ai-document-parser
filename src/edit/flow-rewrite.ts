import path from 'node:path';
import { DocumentError } from '../utils/errors.js';
import type { ContentBlock, EditFlowOp, InlineRun } from '../types.js';
import { parseDocument } from '../parse/router.js';
import { htmlToBlocks, markdownToBlocks } from '../parse/html-blocks.js';
import {
  emitHtmlFromBlocks,
  emitMarkdownFromBlocks,
} from '../convert/blocks.js';
import { writeDocument } from '../write/index.js';
import { ensureParentDir, writeFileAtomic } from '../utils/paths.js';
import { replaceVisibleTokenText } from './text-selection.js';

const FLOW_EXTENSIONS = new Set([
  '.docx',
  '.odt',
  '.rtf',
  '.epub',
  '.md',
  '.markdown',
  '.html',
  '.htm',
]);

function replaceRuns(
  runs: InlineRun[],
  oldString: string,
  newString: string
): { runs: InlineRun[]; count: number } {
  let offset = 0;
  const tokens = runs.map((run) => {
    const start = offset;
    offset += run.text.length;
    return { start, end: offset, decoded: run.text };
  });
  const result = replaceVisibleTokenText(
    tokens,
    oldString,
    newString,
    true
  );
  return {
    runs: runs.map((run, index) => ({ ...run, text: result.texts[index] })),
    count: result.count,
  };
}

function replaceInBlocks(
  blocks: ContentBlock[],
  oldString: string,
  newString: string
): { blocks: ContentBlock[]; count: number } {
  let total = 0;
  const next = blocks.map((block): ContentBlock => {
    if (
      block.type === 'heading' ||
      block.type === 'paragraph' ||
      block.type === 'list_item'
    ) {
      const result = replaceRuns(block.runs, oldString, newString);
      total += result.count;
      return { ...block, runs: result.runs };
    }
    if (block.type === 'table') {
      const rows = block.rows.map((row) =>
        row.map((cell) => {
          const result = replaceVisibleTokenText(
            [{ start: 0, end: cell.length, decoded: cell }],
            oldString,
            newString,
            true
          );
          total += result.count;
          return result.texts[0];
        })
      );
      return { ...block, rows };
    }
    if (block.type === 'code') {
      const result = replaceVisibleTokenText(
        [{ start: 0, end: block.text.length, decoded: block.text }],
        oldString,
        newString,
        true
      );
      total += result.count;
      return { ...block, text: result.texts[0] };
    }
    return { ...block };
  });
  return { blocks: next, count: total };
}

async function writeBlocksToOriginalFormat(
  filePath: string,
  blocks: ContentBlock[]
): Promise<void> {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.md' || ext === '.markdown') {
    await ensureParentDir(filePath);
    await writeFileAtomic(filePath, emitMarkdownFromBlocks(blocks));
    return;
  }
  if (ext === '.html' || ext === '.htm') {
    await ensureParentDir(filePath);
    await writeFileAtomic(filePath, emitHtmlFromBlocks(blocks));
    return;
  }
  if (ext === '.docx') {
    await writeDocument(filePath, { type: 'docx', blocks });
    return;
  }
  if (ext === '.odt') {
    await writeDocument(filePath, { type: 'odt', blocks });
    return;
  }
  if (ext === '.rtf') {
    await writeDocument(filePath, { type: 'rtf', blocks });
    return;
  }
  if (ext === '.epub') {
    await writeDocument(filePath, { type: 'epub', blocks });
    return;
  }
  throw new DocumentError(
    'INVALID_ARGUMENT',
    `Flow rewrite is not supported for extension ${ext || '(none)'}`
  );
}

async function parsedBlocks(filePath: string): Promise<ContentBlock[]> {
  const parsed = await parseDocument(filePath, { output: 'structured' });
  const blocks = (
    parsed.structured as { blocks?: ContentBlock[] } | undefined
  )?.blocks;
  if (!Array.isArray(blocks)) {
    throw new DocumentError(
      'INVALID_ARGUMENT',
      `Structured flow blocks are unavailable for ${parsed.format}`
    );
  }
  return blocks;
}

export async function editFlowDocument(
  filePath: string,
  op: EditFlowOp
): Promise<{ message: string; changes: number }> {
  const ext = path.extname(filePath).toLowerCase();
  if (!FLOW_EXTENSIONS.has(ext)) {
    throw new DocumentError(
      'INVALID_ARGUMENT',
      `Flow edit is not supported for extension ${ext || '(none)'}`
    );
  }

  if (op.op === 'rewrite') {
    const blocks =
      'blocks' in op
        ? op.blocks
        : op.format === 'markdown'
          ? markdownToBlocks(op.content)
          : htmlToBlocks(op.content);
    await writeBlocksToOriginalFormat(filePath, blocks);
    return {
      message: `Rewrote ${ext.slice(1)} from structured flow: ${filePath}`,
      changes: 1,
    };
  }

  if (!op.oldString) {
    throw new DocumentError('INVALID_ARGUMENT', 'oldString cannot be empty');
  }
  const result = replaceInBlocks(
    await parsedBlocks(filePath),
    op.oldString,
    op.newString
  );
  if (result.count === 0) {
    throw new DocumentError(
      'EDIT_TARGET_NOT_FOUND',
      'Text not found in structured flow. Re-read with output:"structured" and refine oldString.'
    );
  }
  if (result.count > 1 && !op.replaceAll) {
    throw new DocumentError(
      'MULTIPLE_MATCHES',
      `Found ${result.count} matches. Narrow oldString or set replaceAll: true.`
    );
  }

  await writeBlocksToOriginalFormat(filePath, result.blocks);
  return {
    message: `Replaced ${result.count} occurrence(s) by rewriting ${ext.slice(1)}: ${filePath}`,
    changes: result.count,
  };
}
