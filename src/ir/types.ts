import type { ContentBlock, DocumentWarning, InlineRun } from '../types.js';

/** Rich inline span in canonical flow documents. */
export type CanonicalInline = InlineRun & {
  code?: boolean;
  underline?: boolean;
  subscript?: boolean;
  superscript?: boolean;
};

export type CanonicalFlowBlock =
  | { type: 'heading'; level: 1 | 2 | 3 | 4 | 5 | 6; runs: CanonicalInline[] }
  | { type: 'paragraph'; runs: CanonicalInline[] }
  | { type: 'list'; ordered?: boolean; items: CanonicalFlowBlock[][] }
  | { type: 'table'; headerRows?: number; rows: CanonicalTableCell[][] }
  | { type: 'blockquote'; blocks: CanonicalFlowBlock[] }
  | { type: 'code'; text: string; language?: string }
  | { type: 'thematic_break' }
  | { type: 'page_break' }
  | { type: 'image'; src: string; alt?: string }
  | { type: 'footnote_def'; id: string; blocks: CanonicalFlowBlock[] };

export type CanonicalTableCell = {
  text: string;
  runs?: CanonicalInline[];
  colspan?: number;
  rowspan?: number;
};

export type CanonicalFlowDocument = {
  kind: 'flow';
  format: string;
  blocks: CanonicalFlowBlock[];
  raw?: string;
};

export type CanonicalCellValue = string | number | boolean | null;

export type CanonicalSheet = {
  name: string;
  rows: CanonicalCellValue[][];
  formulas?: Record<string, string>;
  merges?: Array<{ from: string; to: string }>;
};

export type CanonicalWorkbook = {
  kind: 'workbook';
  format: string;
  sheets: CanonicalSheet[];
};

export type CanonicalSlide = {
  index: number;
  title?: CanonicalInline[];
  body: CanonicalFlowBlock[];
  notes?: string;
};

export type CanonicalPresentation = {
  kind: 'presentation';
  format: string;
  slides: CanonicalSlide[];
};

export type CanonicalDocument =
  | CanonicalFlowDocument
  | CanonicalWorkbook
  | CanonicalPresentation;

export type DocumentAst =
  | { type: 'flow'; format: string; blocks: CanonicalFlowBlock[] }
  | { type: 'workbook'; format: string; sheets: CanonicalSheet[] }
  | { type: 'presentation'; format: string; slides: CanonicalSlide[] };

export type ParsedCanonical = {
  canonical: CanonicalDocument;
  warnings: DocumentWarning[];
  /** Legacy structured payload for backward compatibility. */
  structured: Record<string, unknown>;
  plainText: string;
};

export type DocumentKind = CanonicalDocument['kind'];

export type ConvertTarget =
  | 'md'
  | 'html'
  | 'csv'
  | 'text'
  | 'chunks'
  | 'docx'
  | 'odt'
  | 'rtf'
  | 'epub'
  | 'ods'
  | 'xlsx'
  | 'odp'
  | 'pptx';

export type AdapterCapabilities = {
  format: string;
  kind: DocumentKind;
  parse: boolean;
  convertTo: ConvertTarget[];
  edit: boolean;
  write: boolean;
};
