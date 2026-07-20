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

**Core formats:** `docx`, `xlsx`, `xls`, `pptx`, `pdf`, `md`, `html`, `csv`, legacy `doc`, `odt`, `ods`, `odp`, `rtf`, `epub`.

**Structured contracts (for edit / agent tools):**

- DOCX: `{ format: 'docx', paragraphs: string[], blocks: ContentBlock[] }` — use `paragraphs` for edit; `blocks` for rich convert
- Spreadsheet: `{ format: 'xlsx'|'xls', sheets: { [name]: cell[][] } }`
- PPTX: `{ format: 'pptx', slides: string[][], blocks: ContentBlock[] }`
- Legacy DOC: `{ format: 'doc', paragraphs: string[], source: 'ai-document-parser-doc-binary' }`
- MD/HTML: `{ format, paragraphs: string[], blocks?: ContentBlock[], raw?: string }` (md keeps `raw` for passthrough)
- CSV: `{ format: 'csv', rows: string[][] }`

`ContentBlock` covers headings, paragraphs (with bold/italic runs), list items, GFM tables, thematic breaks, and code fences.

ZIP packages without a reliable extension are refined via package parts / `[Content_Types].xml` (docx vs xlsx vs pptx).

## `convertDocument(input, to, options?)`

One-step parse + generate. Always parses with `output: 'structured'` internally, then emits.

```ts
type ConvertOptions = {
  /** Reserved; currently ignored. */
  ocr?: boolean;
  extractAttachments?: boolean;
  generator?: Record<string, unknown>;
  /** Sheet name to export when converting a workbook to CSV (default: first sheet). */
  sheet?: string;
};
```

For **md/html**, rich sources emit from `blocks` (structure preserved):

| Source | md/html quality |
|--------|-----------------|
| docx | headings, **bold**/*italic*/~~strike~~, ordered/bullet lists, **tables**, `[links](url)`, GFM footnotes; images omitted |
| xlsx / xls | sheet headings + GFM tables |
| pptx | `## Slide N`, title as `###`, tables when present |
| html | h1–h6, lists, tables, emphasis |
| csv | GFM table |
| md → md | **passthrough** of original file text |
| pdf / legacy doc | plain text only (no reliable structure) |

Supported `to`: `'md' | 'html' | 'csv' | 'text' | 'chunks' | 'rtf' | 'epub' | 'docx' | 'odt'`

| Source | Allowed targets |
|--------|----------------|
| docx / doc / odt / rtf / epub | md, html, text, chunks, docx, odt, rtf, epub |
| xlsx / xls / ods | csv, md, html, text, chunks |
| pptx / odp | md, html, text, chunks |
| pdf | md, html, text, chunks (plain) |
| md / html / csv | md, html, text, csv, chunks, docx, odt, rtf, epub |

Not supported as convert **targets**: `'pdf'` (for PDF creation use `writeDocument`).

Use `getCapabilities()` (also exposed as CLI `capabilities`) to inspect per-format parse/convert/edit/write support.

Extensionless paths and mislabeled binaries are detected from magic bytes / ZIP package parts when possible.

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
{ type: 'odt' | 'rtf' | 'epub', content: string | string[], title?: string }
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
  ContentBlock,
  InlineRun,
  DocumentResult,
  ConvertResult,
} from 'ai-document-parser';
```
