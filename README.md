# ai-document-parser

TypeScript library for **parse / convert / edit / write** of office documents — built for AI agents and Node apps.

- **MIT** license
- Surgical **DOCX** / spreadsheet edit without full regeneration
- Permissive dependencies only (see [LICENSES.md](LICENSES.md))

## Install

```bash
npm install ai-document-parser
```

## Requirements

The package is **ESM-only** (`"type": "module"`). Import it with `import { … } from 'ai-document-parser'` — CommonJS `require()` is not supported. Node.js **>= 22.13.0** is required.


## Quick start

```ts
import { parseDocument, convertDocument, editDocument, writeDocument } from 'ai-document-parser';

// Parse
const result = await parseDocument('report.docx', { output: 'text' });
// output: 'text' | 'structured' | 'ast' | 'chunks'

// Convert preserves structure where possible (headings, tables, emphasis)
const { value } = await convertDocument('report.docx', 'md');
const slides = await convertDocument('deck.pptx', 'md'); // ## Slide N …

// Surgical edit
await editDocument('contract.docx', {
    type: 'docx',
    op: 'replace',
    oldString: 'Acme Corp',
    newString: 'Beta LLC',
});

await editDocument('budget.xlsx', {
    type: 'spreadsheet',
    op: 'patchCells',
    updates: [{ sheet: 'Q1', cell: 'B2', value: 1500 }],
});

// Create — plain paragraphs (markdown markers are literal text)
await writeDocument('memo.docx', {
    type: 'docx',
    content: 'Paragraph one\nParagraph two',
});

// Create — rich DOCX from markdown (lists, tables, headings, emphasis)
await writeDocument('report.docx', {
    type: 'docx',
    format: 'markdown',
    content: '# Title\n\n- item\n\n| A | B |\n|---|---|\n| 1 | 2 |\n',
});

await writeDocument('out.pdf', {
    type: 'pdf',
    content: 'Текст с кириллицей',
    title: 'Отчёт',
});
```

## Agent workflow

1. `parseDocument(path, { output: 'structured' })` — paragraphs / sheet grids / pptx slides
2. Inspect `getCapabilities(format).editModes` and choose `surgical`, `patchCells`, or explicit `rewrite`
3. `editDocument(...)` — targeted or structured changes
4. Re-read structured output to verify

PDF remains read/convert-only for writing (plain-text `writeDocument`). Scanned PDFs and raster images can be read via OCR when a Tesseract runtime is present — see [OCR](#ocr-scanned-pdfs-and-images).

See [docs/AGENT_GUIDE.md](docs/AGENT_GUIDE.md).

## Format matrix (v0.3)

| Format          | Parse     | Convert                                   | Edit operation (fidelity)      | Write                |
| --------------- | --------- | ----------------------------------------- | ------------------------------ | -------------------- |
| docx            | yes       | md/html/text/chunks/docx/odt/rtf/epub     | surgical replace/append; rewrite | plain / blocks / md  |
| xlsx / xls      | yes       | csv/md/html/text/chunks                   | patchCells (surgical)          | yes                  |
| pptx            | yes       | md/html/text/chunks                       | surgical replace              | plain / blocks / md  |
| pdf             | text + OCR of empty pages | md/html/text/chunks              | —                              | plain text           |
| images (png/jpg/tiff/gif/webp/bmp) | OCR text | md/html/text/chunks                | —                              | —                    |
| md / html       | yes       | md/html/text/csv/chunks/docx/odt/rtf/epub | rewrite                        | via flow rewrite     |
| csv             | yes       | md/html/text/csv/chunks/docx/odt/rtf/epub | —                              | —                    |
| legacy .doc     | heuristic | md/html/text/chunks/docx/odt/rtf/epub     | —                              | via convertDocToDocx |
| odt             | yes       | md/html/text/chunks/docx/odt/rtf/epub     | surgical replace; rewrite     | plain / blocks / md  |
| ods             | yes       | csv/md/html/text/chunks                   | patchCells (rewrite)           | —                    |
| odp             | yes       | md/html/text/chunks                       | —                              | —                    |
| rtf / epub      | yes       | md/html/text/chunks/docx/odt/rtf/epub     | surgical replace; rewrite     | plain / blocks / md  |
| legacy .ppt     | —         | —                                         | —                              | —                    |

Notes:

- Convert → **md/html** preserves structure for DOCX (headings, tables, bold/italic/strike, ordered lists, hyperlinks, footnotes; images omitted), HTML, CSV (GFM table), PPTX slides, XLSX sheets. MD→md is passthrough. PDF / legacy `.doc` stay plain text (PDF may include OCR text).
- Convert target `pdf` is **not** supported (create PDF with `writeDocument`).
- Convert target `pptx` / `pdf` is **not** supported by design (create with `writeDocument`). Convert **from** pptx/odp is only `md|html|text|chunks` — **pptx→docx** is unsupported (not a host bug).
- **DOCX** and **PPTX** cannot convert **to csv** (use a spreadsheet source). Multi-sheet workbooks export the first sheet to CSV unless `options.sheet` is passed (`CSV_FIRST_SHEET` message when defaulting).
- **Write** for docx/odt/rtf/epub/pptx: `{ content }` is plain paragraphs only; use `{ format: 'markdown', content }` or `{ blocks }` for headings/lists/tables/blockquotes. PPTX: `#`/`##` start a new slide; lists/tables round-trip structurally (`list_item` / `table` on parse). Inline/blockquote/code on PPTX are lossy and emit `PPTX_*` warnings on `WriteResult.warnings`.
- Markdown `>` blockquotes round-trip for docx/odt/rtf/epub.
- **Edit fidelity:** `EditResult.mode` is `surgical` when package markup is patched in place and `rewrite` when content is regenerated. ODS `patchCells` currently reports `rewrite`.
- `getCapabilities()` (CLI: `capabilities`) reports the exact per-format matrix.


## OCR (scanned PDFs and images)

`parseDocument` already runs OCR when a Tesseract runtime is available:

- Raster images (`png`, `jpg`, `tiff`, `gif`, `webp`, `bmp`) are always OCR'd unless you pass `ocr: false`.
- PDF pages are read with pdf.js first. Pages with almost no text (scans) are OCR'd automatically (`enabled: 'auto'`). Pass `{ enabled: true }` to OCR every page, or `ocr: false` to skip OCR.

The npm package does **not** ship the Tesseract binary or language data. Provide a runtime in one of these ways (first match wins):

1. Explicit options: `{ bin, tessdataDir, popplerBin }`
2. Environment: `ADP_TESSERACT`, `ADP_TESSDATA` or `TESSDATA_PREFIX`, `ADP_PDFTOPPM`
3. A bundled folder at `resources/ocr` (Electron hosts can pack this as `extraResources` so `process.resourcesPath/ocr` resolves)
4. `tesseract` (and for scanned PDFs, `pdftoppm`) on `PATH`

Bundled layout:

```
resources/ocr/bin/tesseract(.exe)
resources/ocr/tessdata/eng.traineddata
resources/ocr/tessdata/<lang>.traineddata
resources/ocr/bin/pdftoppm(.exe)   (optional)
resources/ocr/lib/                 (optional)
```

A bundled runtime is used only when the binary exists **and** every requested language has a `*.traineddata` file. Fetch official tessdata (and optionally copy a system binary) with:

```bash
node scripts/fetch-ocr-runtime.mjs --dry-run
node scripts/fetch-ocr-runtime.mjs --langs rus,eng
node scripts/fetch-ocr-runtime.mjs --langs deu,eng --set best
```

`--set` may be `fast` (default), `best`, or `standard`. Any Tesseract language id works — URLs are built from the official GitHub tessdata templates. Scanned-PDF OCR also needs Poppler (`pdftoppm`); it is not downloaded by the script. Image OCR works with Tesseract alone.

```ts
import { parseDocument, convertDocument, resolveOcrRuntime } from 'ai-document-parser';
import type { OcrOptions, OcrProvider } from 'ai-document-parser';

// Historical default (rus+eng, auto on empty PDF pages and images)
await parseDocument('scan.pdf');

await parseDocument('scan.pdf', { ocr: false });
await parseDocument('scan.de.pdf', {
  ocr: {
    languages: ['deu', 'eng'],
    tessdataDir: '/usr/share/tesseract-ocr/5/tessdata',
  },
});

// Custom engine (tesseract.js, a cloud API, …)
const provider: OcrProvider = {
  async recognize({ imagePath, page }) {
    const text = await myEngine(imagePath);
    return { words: text ? [{ text, x: 0, y: 0, width: 1, height: 10, page }] : [] };
  },
};
await parseDocument('scan.png', { ocr: { languages: ['eng'], provider } });

const runtime = resolveOcrRuntime({ languages: ['deu', 'eng'] });
```

Warnings such as `OCR_LANG_MISSING`, `PDF_OCR_UNAVAILABLE`, or `IMAGE_OCR_USED` are returned on `DocumentResult.warnings` (and on `ConvertResult.messages` when converting).

## Subpath exports

```ts
import { parseDocument } from 'ai-document-parser/parse';
import { convertDocument } from 'ai-document-parser/convert';
import { editDocument } from 'ai-document-parser/edit';
import { writeDocument } from 'ai-document-parser/write';
```

## CLI

```bash
npx ai-document-parser read file.docx --output text
npx ai-document-parser read file.docx --output structured
npx ai-document-parser read scan.pdf --ocr-lang rus,eng
npx ai-document-parser read scan.pdf --no-ocr
npx ai-document-parser convert file.docx --to md -o out.md
npx ai-document-parser convert deck.pptx --to md -o slides.md
npx ai-document-parser edit file.docx --replace "old" "new"
npx ai-document-parser edit file.docx --replace "dup" "new" --all
npx ai-document-parser write-docx out.docx --content "Hello"
npx ai-document-parser write-xlsx out.xlsx --sheet Data --csv "a,b\n1,2"
npx ai-document-parser write-pdf out.pdf --content "Текст" --title "Отчёт"
npx ai-document-parser capabilities --format docx
```

## License

MIT — see [LICENSE](LICENSE) and [LICENSES.md](LICENSES.md).
