/**
 * ai-document-parser — parse, convert, edit, and write office documents.
 *
 * Import from this package root (or subpaths).
 */

export { parseDocument, convertDocToDocx } from './parse/index.js';
export { convertDocument } from './convert/index.js';
export {
  editDocument,
  editDocx,
  appendDocxParagraphs,
  readDocxStructured,
  updateSpreadsheetCells,
  readSpreadsheetStructured,
} from './edit/index.js';
export { writeDocument, writeDocx, writePdf, writeSpreadsheet } from './write/index.js';
export { DocumentError } from './utils/errors.js';

export type {
  ParseOptions,
  ParseOutputMode,
  ConvertFormat,
  ConvertOptions,
  ConvertResult,
  DocumentInput,
  DocumentResult,
  DocumentWarning,
  DocxStructured,
  SpreadsheetStructured,
  LegacyDocStructured,
  PptxStructured,
  EditDocumentOp,
  EditResult,
  WriteDocumentInput,
  WriteResult,
  DocumentErrorCode,
} from './types.js';
