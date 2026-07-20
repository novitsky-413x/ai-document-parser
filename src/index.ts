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
  editPptxText,
  editOdtText,
  editEpubText,
  editRtfText,
  editFlowDocument,
} from './edit/index.js';
export {
  writeDocument,
  writeDocx,
  writePdf,
  writeSpreadsheet,
  writeDocxFromBlocks,
  writeOdtFromBlocks,
} from './write/index.js';
export { DocumentError } from './utils/errors.js';
export { getCapabilities } from './adapters/registry.js';

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
  ContentBlock,
  InlineRun,
  EditDocumentOp,
  EditDocxOp,
  EditSpreadsheetOp,
  EditPptxOp,
  EditOdtOp,
  EditEpubOp,
  EditRtfOp,
  EditFlowOp,
  EditResult,
  WriteDocumentInput,
  WriteMode,
  WriteMarkupFormat,
  WriteResult,
  DocumentErrorCode,
  DocumentAst,
} from './types.js';
export type {
  AdapterCapabilities,
  ConvertTarget,
  WriteCapabilityMode,
  EditCapabilityMode,
  EditCapabilityOperation,
} from './ir/types.js';
