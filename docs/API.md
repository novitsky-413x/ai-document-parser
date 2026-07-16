# API Reference

## `parseDocument(input, options?)`

Parse a file path, `Buffer`, `Uint8Array`, or `ArrayBuffer`.

```ts
type ParseOptions = {
  output?: 'text' | 'structured' | 'ast' | 'chunks'; // default 'text'
  /** Reserved; currently ignored (no OCR pipeline in core). */
  ocr?: boolean;
  /** Reserved; currently ignored. */
  extractAttachments?: boolean;
  /** Reserved; currently ignored. */
  includeRawContent?: boolean;
};
```

Returns `DocumentResult` with `format`, `text`, `charCount`, optional `structured` / `ast` / `chunks` / `warnings`.

**Core formats:** `docx`, `xlsx`, `xls`, `pptx`, `pdf`, `md`, `html`, `csv`, legacy `doc`.

**Structured contracts (for edit / agent tools):**

- DOCX: `{ format: 'docx', paragraphs: string[] }`
- Spreadsheet: `{ format: 'xlsx'|'xls', sheets: { [name]: cell[][] } }`
- PPTX: `{ format: 'pptx', slides: string[][] }` — text blocks per slide (DrawingML)
- Legacy DOC: `{ format: 'doc', paragraphs: string[], source: 'ai-document-parser-doc-binary' }`
- MD/HTML: `{ format, paragraphs: string[] }`
- CSV: `{ format: 'csv', rows: string[][] }`

ZIP packages without a reliable extension are refined via package parts / `[Content_Types].xml` (docx vs xlsx vs pptx).

## `convertDocument(input, to, options?)`

One-step parse + generate. Always parses with `output: 'structured'` internally, then emits.

```ts
type ConvertOptions = {
  /** Reserved; currently ignored. */
  ocr?: boolean;
  /** Reserved; currently ignored. */
  extractAttachments?: boolean;
  /** Reserved; currently ignored. */
  generator?: Record<string, unknown>;
};
```

Supported `to`: `'md' | 'html' | 'csv' | 'text' | 'chunks'`

| Source | Typical emit |
|--------|----------------|
| docx / doc | md, html, text, chunks — **not** csv |
| xlsx / xls | csv, md, html, text, chunks |
| pptx | md (`## Slide N`), html, text, chunks — **not** csv |
| pdf | md, html, text, chunks |
| md / html / csv | md, html, text, csv, chunks |

Not supported as convert **targets**: `'rtf' | 'epub' | 'pdf'` (for PDF creation use `writeDocument`).

Input may be a path or buffer (same as `parseDocument`).

## `editDocument(path, op)`

Filesystem path only (in-place).

```ts
// DOCX replace
{ type: 'docx', op: 'replace', oldString, newString, replaceAll?: boolean }

// DOCX append
{ type: 'docx', op: 'append', paragraphs: string[] }

// Spreadsheet patch
{ type: 'spreadsheet', op: 'patchCells', updates: [{ sheet?, cell, value }] }
```

Throws `DocumentError` with codes:

- `EDIT_TARGET_NOT_FOUND`
- `MULTIPLE_MATCHES`
- `INVALID_ARGUMENT`

There is **no** surgical edit for pptx / pdf / md / html / csv.

## `writeDocument(path, input)`

```ts
{ type: 'docx', content: string | string[] }
{ type: 'xlsx' | 'xls', sheets: Record<string, cell[][]> }
{ type: 'pdf', content: string, title?: string }
```

## `DocumentError`

```ts
class DocumentError extends Error {
  code: DocumentErrorCode;
  cause?: unknown;
}
```

Common codes: `UNSUPPORTED_FORMAT`, `PARSE_FAILED`, `CONVERT_FAILED`, `EDIT_TARGET_NOT_FOUND`, `MULTIPLE_MATCHES`, `WRITE_FAILED`, `INVALID_ARGUMENT`, `IO_ERROR`, `LEGACY_DOC_LOW_FIDELITY`.

`LICENSE_PEER_MISSING` is reserved / unused in core (kept for type stability).

## `convertDocToDocx(docPath, docxPath?)`

Extract text from legacy `.doc` and write a new `.docx`.

## Types worth importing

```ts
import type {
  DocxStructured,
  SpreadsheetStructured,
  PptxStructured,
  LegacyDocStructured,
  DocumentResult,
  ConvertResult,
} from 'ai-document-parser';
```
