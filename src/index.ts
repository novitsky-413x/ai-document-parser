/**
 * ai-document-parser — parse, convert, edit, and write office documents.
 *
 * Import from this package root (or subpaths).
 */

export { parseDocument, convertDocToDocx } from './parse/index.js';
export {
  resolveOcrRuntime,
  normalizeOcrOptions,
  tesseractOcrProvider,
  createTesseractOcrProvider,
  PDF_OCR_PAGE_CAP,
  DEFAULT_OCR_LANGUAGES,
} from './parse/pdf-ocr.js';
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
  writePptxFromBlocks,
} from './write/index.js';
export { DocumentError } from './utils/errors.js';
export { getCapabilities } from './adapters/registry.js';

export type {
  ParseOptions,
  ParseOutputMode,
  ConvertFormat,
  ConvertOptions,
  OcrOptions,
  OcrEnabledMode,
  OcrRuntime,
  OcrRuntimeSource,
  OcrProvider,
  OcrRecognizeRequest,
  OcrRecognizeResult,
  OcrWord,
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
  PptxThemeId,
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
