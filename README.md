# ai-document-parser

TypeScript library for **parse / convert / edit / write** of office documents — built for AI agents and Node apps.

- **MIT** license
- Surgical **DOCX** / spreadsheet edit without full regeneration
- Permissive dependencies only (see [LICENSES.md](LICENSES.md))

## Install

```bash
npm install ai-document-parser
```

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

// Create
await writeDocument('memo.docx', {
  type: 'docx',
  content: 'Paragraph one\nParagraph two',
});

await writeDocument('out.pdf', {
  type: 'pdf',
  content: 'Текст с кириллицей',
  title: 'Отчёт',
});
```

## Agent workflow

1. `parseDocument(path, { output: 'structured' })` — paragraphs / sheet grids / pptx slides
2. `editDocument(...)` — targeted changes (**docx** / **xlsx** only)
3. Re-read structured output to verify

For read-only formats (pptx, pdf, md, …) prefer `convertDocument` → `md` / `text` / `chunks`.

See [docs/AGENT_GUIDE.md](docs/AGENT_GUIDE.md).

## Format matrix (v0.1)

| Format | Parse | Convert | Edit | Write |
|--------|-------|---------|------|-------|
| docx | yes | md/html/text/chunks/docx/odt/rtf/epub | replace/append | yes |
| xlsx / xls | yes | csv/md/html/text/chunks | cell patch | yes |
| pptx | yes | md/html/text/chunks | — | — |
| pdf | text | md/html/text/chunks | — | plain text |
| md / html / csv | yes | md/html/text/csv/chunks/docx/odt/rtf/epub | — | — |
| legacy .doc | heuristic | md/html/text/chunks/docx/odt/rtf/epub | — | via convertDocToDocx |
| odt | yes | md/html/text/chunks/docx/odt/rtf/epub | — | yes |
| ods / odp | yes | csv (ods) / md/html/text/chunks | — | — |
| rtf / epub | yes | md/html/text/chunks/docx/odt/rtf/epub | — | yes |
| legacy .ppt | — | — | — | — |

Notes:

- Convert → **md/html** preserves structure for DOCX (headings, tables, bold/italic/strike, ordered lists, hyperlinks, footnotes; images omitted), HTML, CSV (GFM table), PPTX slides, XLSX sheets. MD→md is passthrough. PDF / legacy `.doc` stay plain text.
- Convert target `pdf` is **not** supported (create PDF with `writeDocument`).
- **DOCX** and **PPTX** cannot convert **to csv** (use a spreadsheet source). Multi-sheet workbooks export the first sheet to CSV unless `options.sheet` is passed.
- `getCapabilities()` (CLI: `capabilities`) reports the exact per-format matrix.

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
