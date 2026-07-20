# Agent guide

Recommended workflow for AI coding agents (WATCHER, Cursor tools, RAG pipelines).

## Read → edit → verify (DOCX / XLSX)

1. **Read structured** (not raw AST) so edit targets match what you will search for:

```ts
const { structured } = await parseDocument(path, { output: 'structured' });
```

2. **Edit** with exact strings / cells from that structured output:

```ts
// DOCX — use structured.paragraphs[*] as oldString
await editDocument(path, {
  type: 'docx',
  op: 'replace',
  oldString: structured.paragraphs[0],
  newString: '…',
  replaceAll: false,
});
```

3. **Verify** by reading structured again.

## Spreadsheets

- Prefer `output: 'structured'` to get 2D grids.
- Patch with A1-style addresses: `{ cell: 'B2', value: 42, sheet: 'Q1' }`.
- Creating a new workbook: `writeDocument(path, { type: 'xlsx', sheets })`.
- Convert: `convertDocument(path, 'csv' | 'md' | …)`.

## Presentations (PPTX)

- Parse: `{ format: 'pptx', slides: string[][] }` — each slide is an array of text blocks.
- Convert: `convertDocument(path, 'md')` yields `## Slide 1` … sections.
- **No** surgical edit / write for PPTX in core — extract or convert only.

## Convert for agents

Use `convertDocument` when you need a single artifact for RAG or downstream tools:

| Goal | Call |
|------|------|
| Markdown for LLM context | `convertDocument(path, 'md')` — DOCX keeps headings/tables/emphasis/lists/links/footnotes (images omitted) |

| Embedding chunks | `convertDocument(path, 'chunks')` or `parseDocument(..., { output: 'chunks' })` |
| Spreadsheet dump | `convertDocument(xlsxPath, 'csv')` or `'md'` for GFM tables |

Do **not** convert DOCX/PPTX to `csv`. The only unsupported convert target is `pdf` (use `writeDocument` for plain-text PDF).  
Flow formats (docx/odt/rtf/epub/md/html) also convert to binary `docx` / `odt` / `rtf` / `epub`.  
PDF / legacy `.doc` → md is plain text only (no reliable document structure).  
Multi-sheet workbooks: pass `options.sheet` to `convertDocument(path, 'csv', { sheet: 'Q2' })`.

## When to use AST / chunks

- `ast` — advanced transformation, custom generators, RAG with structure.
- `chunks` — embedding / retrieval pipelines.
- Do **not** invent `oldString` from AST text nodes for surgical DOCX edit; use `structured.paragraphs`.

## Legacy `.doc`

Extraction is heuristic (OLE CFB string scrape). Prefer `convertDocToDocx` then edit the `.docx`.

## Errors

Catch `DocumentError` and branch on `error.code`:

| Code | Meaning |
|------|---------|
| `EDIT_TARGET_NOT_FOUND` | `oldString` missing — re-read structured |
| `MULTIPLE_MATCHES` | Set `replaceAll: true` or narrow string |
| `UNSUPPORTED_FORMAT` | Wrong extension / not a core format / bad convert target |
| `INVALID_ARGUMENT` | Bad edit payload |
| `PARSE_FAILED` / `CONVERT_FAILED` | Corrupt file or emit failure |

## Truncation

This library returns full content. Agent hosts (like WATCHER) may truncate tool replies — re-read with `structured` or page through large sheets by sheet name / pptx slides by index.
