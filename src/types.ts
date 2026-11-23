/**
 * Shared types for ai-document-parser.
 */

export type ParseOutputMode = 'text' | 'structured' | 'ast' | 'chunks';

export type ConvertFormat =
  | 'md'
  | 'html'
  | 'csv'
  | 'rtf'
  | 'epub'
  | 'text'
  | 'chunks'
  | 'pdf'
  | 'docx'
  | 'odt'
  | 'ods'
  | 'odp';

export type DocumentInput = string | Buffer | Uint8Array | ArrayBuffer;

export type OcrEnabledMode = boolean | 'auto';

export type OcrRuntimeSource = 'options' | 'env' | 'bundled' | 'path';

export interface OcrWord {
  text: string;
  x: number;
  y: number;
  width: number;
  height: number;
  page: number;
}

export interface OcrRuntime {
  bin: string;
  libDir?: string;
  tessdataDir?: string;
  popplerBin?: string;
  bundled: boolean;
  source: OcrRuntimeSource;
}

export interface OcrRecognizeRequest {
  imagePath: string;
  page: number;
  languages: string[];
  psm: number;
  oem?: number;
  timeoutMs: number;
  runtime: OcrRuntime;
}

export interface OcrRecognizeResult {
  words: OcrWord[];
  warnings?: DocumentWarning[];
}

export interface OcrProvider {
  /** Recognize one raster image. The default provider shells out to Tesseract. */
  recognize(request: OcrRecognizeRequest): Promise<OcrRecognizeResult>;
}

export interface OcrOptions {
  /**
   * `'auto'` (default): OCR raster images and PDF pages that have no usable text layer.
   * `true`: OCR every PDF page (and images).
   * `false`: skip OCR entirely.
   */
  enabled?: OcrEnabledMode;
  /** Tesseract language codes, e.g. `['rus', 'eng']`. Default `['rus', 'eng']`. */
  languages?: string[];
  /** Path to the Tesseract binary. Overrides `ADP_TESSERACT`. */
  bin?: string;
  /** Directory containing `*.traineddata`. Overrides `ADP_TESSDATA` / `TESSDATA_PREFIX`. */
  tessdataDir?: string;
  /** Path to `pdftoppm` (Poppler). Overrides `ADP_PDFTOPPM`. */
  popplerBin?: string;
  /** Tesseract page segmentation mode. Default 6. */
  psm?: number;
  /** Tesseract OCR engine mode. */
  oem?: number;
  /** Rasterization DPI for scanned PDF pages. Default 300. */
  dpi?: number;
  /** Max PDF pages to OCR in one call. Default 80. */
  pageCap?: number;
  /** Timeout in milliseconds for a Tesseract recognition call. Default 120000. */
  timeoutMs?: number;
  /** Replace the default Tesseract CLI provider (e.g. tesseract.js or a cloud engine). */
  provider?: OcrProvider;
}

export interface ParseOptions {
  /** Default: 'text' */
  output?: ParseOutputMode;
  /**
   * `false` skips OCR. `true` is the same as `{ enabled: 'auto' }` (historical
   * default: images and PDF pages without a usable text layer). Pass an
   * object to set languages, paths, or a custom provider.
   */
  ocr?: boolean | OcrOptions;
  extractAttachments?: boolean;
  includeRawContent?: boolean;
}

export interface ConvertOptions {
  ocr?: boolean | OcrOptions;
  extractAttachments?: boolean;
  generator?: Record<string, unknown>;
  /** Sheet name to export when converting a workbook to CSV (default: first sheet). */
  sheet?: string;
}

export interface DocumentWarning {
  code: string;
  message: string;
}

export interface DocumentResult {
  /** Detected / resolved format (e.g. docx, xlsx, pdf) */
  format: string;
  /** Plain text when output is text, or stringified structured/chunks */
  text: string;
  charCount: number;
  /** Path when input was a file path */
  path?: string;
  /** Structured agent-friendly payload when output === 'structured' */
  structured?:
    | DocxStructured
    | SpreadsheetStructured
    | LegacyDocStructured
    | PptxStructured
    | Record<string, unknown>;
  /** Simplified document shape when output === 'ast' */
  ast?: DocumentAst;
  /** Chunks when output === 'chunks' */
  chunks?: unknown[];
  warnings?: DocumentWarning[];
}

export interface DocxStructured {
  format: 'docx';
  /** Flat plain-text paragraphs for surgical edit (replace). */
  paragraphs: string[];
  /** Structural blocks for convert → md/html. */
  blocks: ContentBlock[];
}

export interface SpreadsheetStructured {
  format: 'xlsx' | 'xls' | 'ods';
  sheets: Record<string, (string | number | boolean | null)[][]>;
}

export interface LegacyDocStructured {
  format: 'doc';
  paragraphs: string[];
  source: 'ai-document-parser-doc-binary';
}

export interface PptxStructured {
  format: 'pptx';
  /** Per-slide text blocks (DrawingML a:t runs grouped by a:p). */
  slides: string[][];
  /** Flat structural blocks for convert → md/html. */
  blocks: ContentBlock[];
}

/** Inline text run with optional emphasis / link. */
export type InlineRun = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  /** When set, emit as markdown/HTML link. */
  href?: string;
  strike?: boolean;
  /** Inline code span. */
  code?: boolean;
};

/** Shared structural model for convert → md/html. */
export type ContentBlock =
  | { type: 'heading'; level: 1 | 2 | 3 | 4 | 5 | 6; runs: InlineRun[] }
  | { type: 'paragraph'; runs: InlineRun[] }
  | { type: 'list_item'; level: number; ordered?: boolean; runs: InlineRun[] }
  | { type: 'table'; rows: string[][] }
  | { type: 'thematic_break' }
  | { type: 'code'; text: string; language?: string }
  | { type: 'blockquote'; blocks: ContentBlock[] };

export interface ConvertResult {
  value: string | Uint8Array | unknown;
  messages: DocumentWarning[];
  format: ConvertFormat;
}

export type EditDocxOp =
  | {
      type: 'docx';
      op: 'replace';
      oldString: string;
      newString: string;
      replaceAll?: boolean;
    }
  | {
      type: 'docx';
      op: 'append';
      paragraphs: string[];
    };

export type EditSpreadsheetOp = {
  type: 'spreadsheet';
  op: 'patchCells';
  updates: Array<{
    sheet?: string;
    cell: string;
    value: string | number | boolean | null;
  }>;
};

export type EditPptxOp = {
  type: 'pptx';
  op: 'replace';
  oldString: string;
  newString: string;
  replaceAll?: boolean;
  /** One-based slide number. Omit to search all slides. */
  slide?: number;
  /** Also search notes slide parts. Default: false. */
  includeNotes?: boolean;
};

export type EditOdtOp = {
  type: 'odt';
  op: 'replace';
  oldString: string;
  newString: string;
  replaceAll?: boolean;
};

export type EditEpubOp = {
  type: 'epub';
  op: 'replace';
  oldString: string;
  newString: string;
  replaceAll?: boolean;
  /** One-based spine index or resolved XHTML part path. */
  chapter?: number | string;
};

export type EditRtfOp = {
  type: 'rtf';
  op: 'replace';
  oldString: string;
  newString: string;
  replaceAll?: boolean;
};

export type EditFlowOp =
  | {
      type: 'flow';
      op: 'rewrite';
      format: WriteMarkupFormat;
      content: string;
    }
  | {
      type: 'flow';
      op: 'rewrite';
      blocks: ContentBlock[];
    }
  | {
      type: 'flow';
      op: 'replace';
      oldString: string;
      newString: string;
      replaceAll?: boolean;
    };

export type EditDocumentOp =
  | EditDocxOp
  | EditSpreadsheetOp
  | EditPptxOp
  | EditOdtOp
  | EditEpubOp
  | EditRtfOp
  | EditFlowOp;

export interface EditResult {
  path: string;
  message: string;
  changes: number;
  /** Whether the original package was patched or fully regenerated. */
  mode?: 'surgical' | 'rewrite';
}

/** How `writeDocument` interpreted the payload (for agents / hosts). */
export type WriteMode = 'plain' | 'blocks' | 'markdown' | 'html';

/** Markup formats accepted on rich write (`format` + `content`). */
export type WriteMarkupFormat = 'markdown' | 'html';

/** Visual theme presets for PPTX write. */
export type PptxThemeId = 'office' | 'ocean' | 'slate';

/**
 * Create-from-scratch input.
 *
 * For flow formats (docx / odt / rtf / epub), exactly one of:
 * - `content` only → **plain** paragraphs (one line / array entry = one paragraph; MD markers are literal)
 * - `blocks` → structured {@link ContentBlock}s
 * - `format: 'markdown' | 'html'` + `content` → parse markup then emit rich OOXML / ODF / RTF / EPUB / PPTX
 */
export type WriteDocumentInput =
  | { type: 'docx'; content: string | string[] }
  | { type: 'docx'; blocks: ContentBlock[] }
  | { type: 'docx'; format: WriteMarkupFormat; content: string }
  | {
      type: 'xlsx' | 'xls';
      sheets: Record<string, (string | number | boolean | null)[][]>;
    }
  | {
      type: 'pdf';
      content: string;
      title?: string;
    }
  | { type: 'odt'; content: string | string[] }
  | { type: 'odt'; blocks: ContentBlock[] }
  | { type: 'odt'; format: WriteMarkupFormat; content: string }
  | { type: 'rtf'; content: string | string[] }
  | { type: 'rtf'; blocks: ContentBlock[] }
  | { type: 'rtf'; format: WriteMarkupFormat; content: string }
  | { type: 'epub'; content: string | string[]; title?: string }
  | { type: 'epub'; blocks: ContentBlock[]; title?: string }
  | {
      type: 'epub';
      format: WriteMarkupFormat;
      content: string;
      title?: string;
    }
  | { type: 'pptx'; content: string | string[]; theme?: PptxThemeId }
  | { type: 'pptx'; blocks: ContentBlock[]; theme?: PptxThemeId }
  | { type: 'pptx'; format: WriteMarkupFormat; content: string; theme?: PptxThemeId };

export interface WriteResult {
  path: string;
  message: string;
  /** What was actually emitted (for agents / hosts). */
  summary?: {
    mode: WriteMode;
    blockCounts?: Partial<Record<ContentBlock['type'], number>>;
    paragraphCount?: number;
  };
  warnings?: DocumentWarning[];
}

export type DocumentAst =
  | { type: 'flow'; format: string; blocks: unknown[] }
  | { type: 'workbook'; format: string; sheets: unknown[] }
  | { type: 'presentation'; format: string; slides: unknown[] };

export type DocumentErrorCode =
  | 'UNSUPPORTED_FORMAT'
  | 'PARSE_FAILED'
  | 'CONVERT_FAILED'
  | 'EDIT_TARGET_NOT_FOUND'
  | 'MULTIPLE_MATCHES'
  | 'WRITE_FAILED'
  | 'LICENSE_PEER_MISSING'
  | 'LEGACY_DOC_LOW_FIDELITY'
  | 'INVALID_ARGUMENT'
  | 'IO_ERROR';
