# API Reference

## `parseDocument(input, options?)`

Parse a file path, `Buffer`, `Uint8Array`, or `ArrayBuffer`.

```ts
type ParseOptions = {
  output?: 'text' | 'structured' | 'ast' | 'chunks'; // default 'text'
  /**
   * false skips OCR. true is the same as { enabled: 'auto' } (images + PDF
   * pages without a usable text layer). Pass an object for languages, paths,
   * or a custom OcrProvider. Default: auto with languages ['rus', 'eng'].
   */
  ocr?: boolean | OcrOptions;
  /** Reserved; currently ignored. */
  extractAttachments?: boolean;
  /** Reserved; currently ignored. */
  includeRawContent?: boolean;
};

type OcrOptions = {
  enabled?: boolean | 'auto';
  languages?: string[];
  bin?: string;
  tessdataDir?: string;
  popplerBin?: string;
  pdfRasterizer?: 'auto' | 'pdfium' | 'pdftoppm'; // default 'auto': PDFium (WASM), then pdftoppm
  psm?: number;
  oem?: number;
  dpi?: number;
  pageCap?: number;
  timeoutMs?: number;
  provider?: OcrProvider;
};
```

`resolveOcrRuntime({ languages, bin, tessdataDir, popplerBin })` is exported from the package root. Resolution order: explicit options → `ADP_TESSERACT` / `ADP_TESSDATA` / `TESSDATA_PREFIX` / `ADP_PDFTOPPM` → bundled `resources/ocr` (valid when the binary and every requested `*.traineddata` exist) → `tesseract` on PATH. Scanned PDF pages are rendered with the bundled PDFium WebAssembly build (`@hyzyla/pdfium`); `pdftoppm` (looked up the same way and next to the runtime) is only a fallback, reported with `PDF_OCR_RASTER_FALLBACK`. `ADP_PDFIUM_WASM` overrides the `pdfium.wasm` path. Missing languages yield `OCR_LANG_MISSING` and OCR is skipped; any text layer is kept. See the README OCR section.

Returns `DocumentResult` with `format`, `text`, `charCount`, optional `structured` / `ast` / `chunks` / `warnings`.

**Core formats:** `docx`, `xlsx`, `xls`, `pptx`, `pdf`, `md`, `html`, `csv`, legacy `doc`, `odt`, `ods`, `odp`, `rtf`, `epub`, raster `image` (`png`/`jpg`/`tiff`/`gif`/`webp`/`bmp`).

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
  /** Same `ocr` option as `parseDocument` (forwarded to the parse step). */
  ocr?: boolean | OcrOptions;
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
| pptx | `## Slide N`, title as `###`, `list_item` (from DrawingML bullets), tables when present; inline/blockquote/code may be flattened |
| html | h1–h6, lists, tables, emphasis |
| csv | GFM table |
| md → md | **passthrough** of original file text |
| pdf / legacy doc | plain text only (no reliable structure) |

Supported `to`: `'md' | 'html' | 'csv' | 'text' | 'chunks' | 'rtf' | 'epub' | 'docx' | 'odt'`

| Source | Allowed targets |
|--------|----------------|
| docx / doc / odt / rtf / epub | md, html, text, chunks, docx, odt, rtf, epub |
| xlsx / xls / ods | csv, md, html, text, chunks |
| pptx / odp | md, html, text, chunks (**not** docx/odt/rtf/epub — by design) |
| pdf | md, html, text, chunks (plain) |
| md / html / csv | md, html, text, csv, chunks, docx, odt, rtf, epub |

Not supported as convert **targets** (by design): `'pdf'` (use `writeDocument` for plain-text PDF) and `'pptx'` (create with `writeDocument({ type: 'pptx', … })`). Convert **from** pptx/odp never yields Office flow binaries (`pptx→docx` is rejected with `UNSUPPORTED_FORMAT`).

`parseDocument(..., { output: 'ast' })` for flow formats (including DOCX/PPTX) returns canonical flow blocks derived from structured `blocks` when available.

Multi-sheet CSV/MD: default is the first sheet; `ConvertResult.messages` may include `{ code: 'CSV_FIRST_SHEET', … }`. Pass `options.sheet` to select another sheet.

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

Create a new file from scratch. For **docx / odt / rtf / epub / pptx**, pick **exactly one** mode:

| Mode | Input shape | Behavior |
|------|-------------|----------|
| **plain** | `{ type, content: string \| string[] }` | One paragraph per line / array entry. Markdown markers (`- item`, `\| a \| b \|`) are **literal text**, not lists/tables. |
| **blocks** | `{ type, blocks: ContentBlock[] }` | Emit real headings, lists, tables, emphasis, links, code, blockquotes. |
| **markup** | `{ type, format: 'markdown' \| 'html', content: string }` | Parse MD/HTML → same structured emit as convert. |

### HTML layout (`format: 'html'`)

HTML writes keep visual layout, not only structure. Applied CSS: inline `style`, `<style>` rules with tag / `.class` / `#id` selectors joined by descendant or `>` combinators, legacy `align=`, `<center>`, `<font face color>`, and browser defaults for `h1`–`h6` / `p` (sizes, bold headings, 1em paragraph margins). `<head>` / `<title>` never become body text; `<br>` is a line break (`\n` in run text, `w:br` / `text:line-break` / `a:br`); `div` / `section` / `article` / `header` / `footer` / … are block containers whose style is inherited (margins go to the first/last child, CSS sibling-margin collapse is emulated).

| CSS | DOCX | ODT | PPTX | RTF / EPUB |
|-----|------|-----|------|------------|
| `text-align` | ✓ | ✓ | ✓ | — |
| `margin-*` / `padding-*` (spacing before/after, indents), `text-indent`, `line-height`, `page-break-*` / `break-*` | ✓ | ✓ | — | — |
| `font-size`, `font-family`, `color`, `letter-spacing`, `font-weight`, `font-style`, `text-decoration` | ✓ | ✓ | bold/italic/strike/underline | bold/italic |
| `body` font / size / color / line-height → document defaults | ✓ | ✓ | — | — |
| `@page { size; margin }` (+ `body` margin) → page size / margins | ✓ | ✓ | — | — |

Anything else (unsupported properties, pseudo-classes, `@media`, …) is skipped and reported once as `HTML_STYLE_IGNORED` in `WriteResult.warnings`, with the property / selector names. RTF / EPUB / PPTX add the same code when the HTML carried CSS they cannot represent.

The block model carries this as optional `layout` on `paragraph` / `heading` blocks (`BlockLayout`: `align`, `spaceBeforePt`, `spaceAfterPt`, `lineHeight`, `indentLeftPt`, `indentRightPt`, `firstLineIndentPt`, `pageBreakBefore`) and optional run props (`sizePt`, `font`, `color`, `letterSpacingPt`, `underline`). Blocks mode for docx/odt accepts `layout?: DocumentLayout` (`font`, `sizePt`, `color`, `lineHeight`, `page: { widthPt, heightPt, margin*Pt }`). Parsing / converting HTML files stays content-only (no `layout` in parse results).

Markdown `>` blockquotes round-trip for docx/odt/rtf/epub via durable quote styles (`ADPQuote` / Word `Quote`) or RTF `{\*\adpquote}`.

PPTX: `#` / `##` headings start a new slide; following blocks become body text (lists → DrawingML bullets; tables → `a:tbl`). Bold/italic/strike/underline/code runs are preserved. Optional `theme?: 'office' | 'ocean' | 'slate'` (default `office`). Convert **to** `pptx` remains unsupported. Lossy bits may set `WriteResult.warnings`: `PPTX_INLINE_FLATTENED` (hyperlinks/sub/sup styled as text, no clickable relationship), `PPTX_BLOCKQUOTE_UNWRAPPED`, `PPTX_CODE_AS_PARAGRAPHS`. Advanced `writePptxFromBlocks` returns `{ buffer, warnings }`.

```ts
// Plain paragraphs only
{ type: 'docx', content: 'Paragraph one\nParagraph two' }

// Preferred for LLM agents — rich DOCX/ODT/RTF/EPUB/PPTX
{ type: 'docx', format: 'markdown', content: '# Title\n\n> quote\n\n- item\n' }

// Layout-faithful documents (title pages, letters, forms) — CSS is honoured
{ type: 'docx', format: 'html', content: '<p style="text-align:center;font-size:22pt"><b>Title</b></p>' }

// Typed hosts
{ type: 'docx', blocks: [{ type: 'heading', level: 1, runs: [{ text: 'Title' }] }] }

{ type: 'xlsx' | 'xls', sheets: Record<string, cell[][]> }
{ type: 'pdf', content: string, title?: string }  // plain text PDF
{ type: 'odt' | 'rtf' | 'epub' | 'pptx', /* same plain | blocks | format+content as docx */, title?: string, theme?: 'office' | 'ocean' | 'slate' /* pptx only */ }
```

`WriteResult` includes `summary.mode` (`'plain' | 'blocks' | 'markdown' | 'html'`) and optional `warnings` (e.g. `PLAIN_WRITE_IGNORED_MARKUP` when plain `content` looks like markdown structure, `HTML_STYLE_IGNORED` when HTML CSS could not be applied).

`getCapabilities()` sets `writeModes: ['plain','blocks','markdown','html']` for docx/odt/rtf/epub/pptx (spreadsheets/PDF: `['plain']` only).

Advanced: `writeDocxFromBlocks` / `writeOdtFromBlocks` / `writePptxFromBlocks` are exported for hosts that already hold canonical flow blocks.

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
  OcrOptions,
  OcrProvider,
  OcrRuntime,
} from 'ai-document-parser';
```
