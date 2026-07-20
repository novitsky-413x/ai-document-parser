import path from 'node:path';
import { DocumentError } from '../utils/errors.js';
import type { EditDocumentOp, EditResult } from '../types.js';
import {
  appendDocxParagraphs,
  editDocx,
} from './docx-ooxml.js';
import { updateSpreadsheetCells } from './spreadsheet.js';

/**
 * Surgically edit an existing docx or spreadsheet without full regeneration.
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
      return { path: filePath, message: result.message, changes: result.changes };
    }
    if (op.op === 'append') {
      const result = await appendDocxParagraphs(filePath, op.paragraphs);
      return { path: filePath, message: result.message, changes: result.changes };
    }
  }

  if (op.type === 'spreadsheet' && op.op === 'patchCells') {
    if (ext && !['.xlsx', '.xls'].includes(ext)) {
      throw new DocumentError(
        'INVALID_ARGUMENT',
        `Expected .xlsx/.xls file for spreadsheet edit, got ${ext || '(no extension)'}`
      );
    }
    const result = await updateSpreadsheetCells(filePath, op.updates);
    return { path: filePath, message: result.message, changes: result.changes };
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
