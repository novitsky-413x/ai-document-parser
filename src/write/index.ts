import path from 'node:path';
import { DocumentError } from '../utils/errors.js';
import type { WriteDocumentInput, WriteResult } from '../types.js';
import { writeDocx } from './docx.js';
import { writePdf } from './pdf.js';
import { writeSpreadsheet } from './spreadsheet.js';

/**
 * Create a new document from scratch (plain text / grid).
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
    const result = await writeDocx(filePath, input.content);
    return { path: filePath, message: result.message };
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
    return { path: filePath, message: result.message };
  }

  if (input.type === 'pdf') {
    if (ext && ext !== '.pdf') {
      throw new DocumentError(
        'INVALID_ARGUMENT',
        `Expected .pdf path for type pdf, got ${ext}`
      );
    }
    const result = await writePdf(filePath, input.content, input.title);
    return { path: filePath, message: result.message };
  }

  throw new DocumentError(
    'INVALID_ARGUMENT',
    `Unsupported write type: ${(input as { type: string }).type}`
  );
}

export { writeDocx } from './docx.js';
export { writePdf } from './pdf.js';
export { writeSpreadsheet } from './spreadsheet.js';
