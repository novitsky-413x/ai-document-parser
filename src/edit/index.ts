import path from 'node:path';
import { DocumentError } from '../utils/errors.js';
import type { EditDocumentOp, EditResult } from '../types.js';
import {
  appendDocxParagraphs,
  editDocx,
} from './docx-ooxml.js';
import { updateSpreadsheetCells } from './spreadsheet.js';
import { editPptxText } from './pptx-ooxml.js';
import { editOdtText } from './odt-xml.js';
import { editEpubText } from './epub-xhtml.js';
import { editRtfText } from './rtf-tokenizer.js';
import { editFlowDocument } from './flow-rewrite.js';

/**
 * Edit an existing document using the requested surgical / patch / rewrite mode.
 */
export async function editDocument(
  filePath: string,
  op: EditDocumentOp
): Promise<EditResult> {
  const ext = path.extname(filePath).toLowerCase();

  if (op.type === 'docx') {
    if (ext && !['.docx'].includes(ext)) {
      throw new DocumentError(
        'INVALID_ARGUMENT',
        `Expected .docx file for docx edit, got ${ext || '(no extension)'}`
      );
    }
    if (op.op === 'replace') {
      const result = await editDocx(
        filePath,
        op.oldString,
        op.newString,
        op.replaceAll ?? false
      );
      return {
        path: filePath,
        message: result.message,
        changes: result.changes,
        mode: 'surgical',
      };
    }
    if (op.op === 'append') {
      const result = await appendDocxParagraphs(filePath, op.paragraphs);
      return {
        path: filePath,
        message: result.message,
        changes: result.changes,
        mode: 'surgical',
      };
    }
  }

  if (op.type === 'spreadsheet' && op.op === 'patchCells') {
    if (ext && !['.xlsx', '.xls', '.ods'].includes(ext)) {
      throw new DocumentError(
        'INVALID_ARGUMENT',
        `Expected .xlsx/.xls/.ods file for spreadsheet edit, got ${ext || '(no extension)'}`
      );
    }
    const result = await updateSpreadsheetCells(filePath, op.updates);
    return {
      path: filePath,
      message: result.message,
      changes: result.changes,
      mode: ext === '.ods' ? 'rewrite' : 'surgical',
    };
  }

  if (op.type === 'pptx' && op.op === 'replace') {
    if (ext && ext !== '.pptx') {
      throw new DocumentError(
        'INVALID_ARGUMENT',
        `Expected .pptx file for pptx edit, got ${ext || '(no extension)'}`
      );
    }
    const result = await editPptxText(
      filePath,
      op.oldString,
      op.newString,
      op.replaceAll ?? false,
      op.slide,
      op.includeNotes ?? false
    );
    return { path: filePath, ...result, mode: 'surgical' };
  }

  if (op.type === 'odt' && op.op === 'replace') {
    if (ext && ext !== '.odt') {
      throw new DocumentError(
        'INVALID_ARGUMENT',
        `Expected .odt file for odt edit, got ${ext || '(no extension)'}`
      );
    }
    const result = await editOdtText(
      filePath,
      op.oldString,
      op.newString,
      op.replaceAll ?? false
    );
    return { path: filePath, ...result, mode: 'surgical' };
  }

  if (op.type === 'epub' && op.op === 'replace') {
    if (ext && ext !== '.epub') {
      throw new DocumentError(
        'INVALID_ARGUMENT',
        `Expected .epub file for epub edit, got ${ext || '(no extension)'}`
      );
    }
    const result = await editEpubText(
      filePath,
      op.oldString,
      op.newString,
      op.replaceAll ?? false,
      op.chapter
    );
    return { path: filePath, ...result, mode: 'surgical' };
  }

  if (op.type === 'rtf' && op.op === 'replace') {
    if (ext && ext !== '.rtf') {
      throw new DocumentError(
        'INVALID_ARGUMENT',
        `Expected .rtf file for rtf edit, got ${ext || '(no extension)'}`
      );
    }
    const result = await editRtfText(
      filePath,
      op.oldString,
      op.newString,
      op.replaceAll ?? false
    );
    return { path: filePath, ...result, mode: 'surgical' };
  }

  if (op.type === 'flow') {
    const result = await editFlowDocument(filePath, op);
    return { path: filePath, ...result, mode: 'rewrite' };
  }

  throw new DocumentError(
    'INVALID_ARGUMENT',
    `Unsupported edit operation: ${JSON.stringify(op)}`
  );
}

export {
  editDocx,
  appendDocxParagraphs,
  readDocxStructured,
  escapeXml,
} from './docx-ooxml.js';
export {
  updateSpreadsheetCells,
  readSpreadsheetStructured,
  readSpreadsheetStructuredAsync,
  type CellUpdate,
  type SheetGrid,
} from './spreadsheet.js';
export { editPptxText } from './pptx-ooxml.js';
export { editOdtText } from './odt-xml.js';
export { editEpubText } from './epub-xhtml.js';
export { editRtfText, tokenizeRtfParagraphs } from './rtf-tokenizer.js';
export { editFlowDocument } from './flow-rewrite.js';
