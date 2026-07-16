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
  | 'pdf';

export type DocumentInput = string | Buffer | Uint8Array | ArrayBuffer;

export interface ParseOptions {
  /** Default: 'text' */
  output?: ParseOutputMode;
  /** Reserved; currently ignored (no OCR pipeline in core). Default: false */
  ocr?: boolean;
  extractAttachments?: boolean;
  includeRawContent?: boolean;
}

export interface ConvertOptions {
  ocr?: boolean;
  extractAttachments?: boolean;
  generator?: Record<string, unknown>;
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
  ast?: unknown;
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
  format: 'xlsx' | 'xls';
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
};

/** Shared structural model for convert → md/html. */
export type ContentBlock =
  | { type: 'heading'; level: 1 | 2 | 3 | 4 | 5 | 6; runs: InlineRun[] }
  | { type: 'paragraph'; runs: InlineRun[] }
  | { type: 'list_item'; level: number; ordered?: boolean; runs: InlineRun[] }
  | { type: 'table'; rows: string[][] }
  | { type: 'thematic_break' }
  | { type: 'code'; text: string; language?: string };

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

export type EditDocumentOp = EditDocxOp | EditSpreadsheetOp;

export interface EditResult {
  path: string;
  message: string;
  changes: number;
}

export type WriteDocumentInput =
  | {
      type: 'docx';
      content: string | string[];
    }
  | {
      type: 'xlsx' | 'xls';
      sheets: Record<string, (string | number | boolean | null)[][]>;
    }
  | {
      type: 'pdf';
      content: string;
      title?: string;
    };

export interface WriteResult {
  path: string;
  message: string;
}

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
