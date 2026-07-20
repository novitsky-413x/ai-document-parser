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

Not supported as convert **targets**: `'pdf'` (use `writeDocument` for plain-text PDF) and `'pptx'` (no PPTX writer; parse / convert-from / surgical replace only).

`parseDocument(..., { output: 'ast' })` for flow formats (including DOCX/PPTX) returns canonical flow blocks derived from structured `blocks` when available.

Extensionless paths and mislabeled binaries are detected from magic bytes / ZIP package parts when possible.

## `editDocument(path, op)`

Filesystem path only. `EditResult.mode` reports whether the original package
was patched (`surgical`) or regenerated (`rewrite`).

```ts
// DOCX replace
{ type: 'docx', op: 'replace', oldString, newString, replaceAll?: boolean }

// DOCX append
{ type: 'docx', op: 'append', paragraphs: string[] }

// XLSX / XLS / ODS cell patch (ODS reports mode:'rewrite')
{ type: 'spreadsheet', op: 'patchCells', updates: [{ sheet?, cell, value }] }

// PPTX DrawingML text replace; slide is one-based
{
  type: 'pptx',
  op: 'replace',
  oldString,
  newString,
  replaceAll?: boolean,
  slide?: number,
  includeNotes?: boolean
}

// ODT / RTF text replace
{ type: 'odt' | 'rtf', op: 'replace', oldString, newString, replaceAll?: boolean }

// EPUB spine XHTML replace; chapter is one-based or a resolved spine path
{
  type: 'epub',
  op: 'replace',
  oldString,
  newString,
  replaceAll?: boolean,
  chapter?: number | string
}

// Explicit regeneration for docx/odt/rtf/epub/md/html
{ type: 'flow', op: 'rewrite', format: 'markdown' | 'html', content: string }
{ type: 'flow', op: 'rewrite', blocks: ContentBlock[] }
{ type: 'flow', op: 'replace', oldString, newString, replaceAll?: boolean }
```

Throws `DocumentError` with codes:

- `EDIT_TARGET_NOT_FOUND`
- `MULTIPLE_MATCHES`
- `INVALID_ARGUMENT`

Surgical replace is bounded by semantic containers: it does not match across
paragraphs, slides, EPUB chapters, or RTF destinations. PPTX charts/SmartArt
and arbitrary existing PDF page text are not editable.

`getCapabilities(format)` exposes:

- `editModes`: `surgical` and/or `rewrite`;
- `editOps`: concrete operation names (`replace`, `append`, `patchCells`,
  `rewrite`) for agent tool schemas.

`patchCells` is an operation, not a fidelity claim: XLSX/XLS advertise
`editModes:['surgical']`; ODS advertises `editModes:['rewrite']`.

## `writeDocument(path, input)`

Create a new file from scratch. For **docx / odt / rtf / epub**, pick **exactly one** mode:

| Mode | Input shape | Behavior |
|------|-------------|----------|
| **plain** | `{ type, content: string \| string[] }` | One paragraph per line / array entry. Markdown markers (`- item`, `\| a \| b \|`) are **literal text**, not lists/tables. |
| **blocks** | `{ type, blocks: ContentBlock[] }` | Emit real headings, lists, tables, emphasis, links, code. |
| **markup** | `{ type, format: 'markdown' \| 'html', content: string }` | Parse MD/HTML → same structured emit as convert. |

```ts
// Plain paragraphs only
{ type: 'docx', content: 'Paragraph one\nParagraph two' }

// Preferred for LLM agents — rich DOCX/ODT/RTF/EPUB
{ type: 'docx', format: 'markdown', content: '# Title\n\n- item\n\n| A | B |\n|---|---|\n| 1 | 2 |\n' }

// Typed hosts
{ type: 'docx', blocks: [{ type: 'heading', level: 1, runs: [{ text: 'Title' }] }] }

{ type: 'xlsx' | 'xls', sheets: Record<string, cell[][]> }
{ type: 'pdf', content: string, title?: string }  // plain text PDF
{ type: 'odt' | 'rtf' | 'epub', /* same plain | blocks | format+content as docx */, title?: string }
```

`WriteResult` includes `summary.mode` (`'plain' | 'blocks' | 'markdown' | 'html'`) and optional `warnings` (e.g. `PLAIN_WRITE_IGNORED_MARKUP` when plain `content` looks like markdown structure).

`getCapabilities()` sets `writeModes: ['plain','blocks','markdown','html']` for docx/odt/rtf/epub (spreadsheets/PDF: `['plain']` only).

Advanced: `writeDocxFromBlocks` / `writeOdtFromBlocks` are exported for hosts that already hold canonical flow blocks.

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
