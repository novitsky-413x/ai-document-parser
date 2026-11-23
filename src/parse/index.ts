export { parseDocument, convertDocToDocx } from './router.js';
export {
  readLegacyDoc,
  readLegacyDocBuffer,
} from './legacy-doc.js';
export {
  readPptxStructured,
  readPptxStructuredFromBuffer,
} from './pptx.js';
export {
  resolveOcrRuntime,
  normalizeOcrOptions,
  tesseractOcrProvider,
  createTesseractOcrProvider,
  PDF_OCR_PAGE_CAP,
  DEFAULT_OCR_LANGUAGES,
} from './pdf-ocr.js';
