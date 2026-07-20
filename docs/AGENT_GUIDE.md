# Agent guide

Recommended workflow for AI coding agents (WATCHER, Cursor tools, RAG pipelines).

## Read → choose fidelity → edit → verify

1. **Read structured** (not raw AST) so edit targets match what you will search for:

```ts
const { structured } = await parseDocument(path, { output: 'structured' });
```

2. **Choose an advertised edit mode**:

```ts
const capability = getCapabilities(format)[0];
// capability.editModes: surgical | rewrite
// capability.editOps: replace | append | patchCells | rewrite
```

- `surgical` preserves unrelated package markup/assets;
- `patchCells` is a targeted spreadsheet operation;
- XLSX/XLS advertise surgical mode; ODS advertises rewrite mode;
- `rewrite` intentionally regenerates flow content and can lose unsupported layout.

3. **Edit** with exact strings / cells from structured output:

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

4. **Verify** by reading structured again.

## Spreadsheets

- Prefer `output: 'structured'` to get 2D grids.
- Patch with A1-style addresses: `{ cell: 'B2', value: 42, sheet: 'Q1' }`.
- The same `patchCells` operation accepts `.ods`; its `EditResult.mode` is `rewrite`.
- Creating a new workbook: `writeDocument(path, { type: 'xlsx', sheets })`.
- Convert: `convertDocument(path, 'csv' | 'md' | …)`.

## Create rich DOCX / ODT / RTF / EPUB

**Do not** put markdown into `{ content }` and expect Word lists/tables — that path is **plain paragraphs only** (`- item` stays literal text). Check `getCapabilities('docx').writeModes`.

**Preferred for LLM agents** — markdown (or HTML) on write:

```ts
await writeDocument('report.docx', {
  type: 'docx',
  format: 'markdown',
  content: `# Отчёт

Paragraph with **bold** and *italic*.

- Хлеб
- Молоко

| Квартал | Выручка |
|---------|---------|
| Q1      | 1200000 |
`,
});
```

Same shape works with `type: 'odt' | 'rtf' | 'epub'`.

**Typed hosts** can pass `blocks: ContentBlock[]` instead of `format` + `content`.

If `WriteResult.warnings` contains `PLAIN_WRITE_IGNORED_MARKUP`, re-issue the call with `format: 'markdown'` or `blocks`.

Use `convertDocument` when transforming an **existing** file (e.g. md→docx already on disk), not as the primary create-from-scratch path.

## Presentations (PPTX)

- Parse: `{ format: 'pptx', slides: string[][] }` — each slide is an array of text blocks.
- Convert: `convertDocument(path, 'md')` yields `## Slide 1` … sections.
- Surgical text replace preserves slides/layout/media:

```ts
await editDocument(path, {
  type: 'pptx',
  op: 'replace',
  oldString: structured.slides[0][0],
  newString: 'Updated title',
  slide: 1,
});
```

Use `includeNotes: true` only when notes are intended edit targets. Charts,
SmartArt, adding shapes, and adding slides are not supported.

## ODT / RTF / EPUB edits

- `type:'odt'|'rtf'|'epub', op:'replace'` performs surgical text replacement.
- EPUB accepts `chapter` (one-based spine index or resolved XHTML path).
- RTF v1 edits body paragraphs and intentionally excludes fields/tables/binary destinations.
- For structural changes, use explicit flow rewrite:

```ts
await editDocument(path, {
  type: 'flow',
  op: 'rewrite',
  format: 'markdown',
  content: '# Replacement document\n\n- structured item\n',
});
```

Never present rewrite as surgical. Check `EditResult.mode`.

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
- Do **not** invent `oldString` from AST text nodes. Use structured paragraphs,
  slides, or spine content from a fresh read.

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
