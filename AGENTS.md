# AGENTS.md

Guidance for AI coding agents working on **ai-document-parser**.

This file is for contributors changing this repo. For agents *calling* the library as a tool, see [docs/AGENT_GUIDE.md](docs/AGENT_GUIDE.md) and [docs/API.md](docs/API.md).

## Project overview

TypeScript ESM library (Node ≥22.13, MIT) that **parses, converts, surgically edits, and writes** office documents for AI agents and Node apps.

All parse/convert/edit/write code under `src/` is first-party. Public facade: `parseDocument` · `convertDocument` · `editDocument` · `writeDocument` (+ `DocumentError`, `convertDocToDocx`).

## Clean-room rules

1. Implement parsers/converters from **public format specs** and **public API behavior** (tests + `docs/API.md`) — do not paste historical third-party engine sources into the tree.
2. Prefer extending `src/edit`, `src/write`, `src/parse`, `src/convert` with small focused modules.
3. New formats land only behind the public facade; update the README format matrix in the same change.
4. Prefer small focused modules; do not add abandoned `xlsx` or `docxtemplater` (deny-listed in `check-licenses`).

## Layout

```
src/
  parse/          parseDocument router, detect, formats, pdf, pptx, docx-blocks, html-blocks, text-formats, legacy-doc
  convert/        convertDocument + blocks emit helpers
  edit/           surgical DOCX OOXML + spreadsheet patch
  write/          create docx / xlsx / plain pdf / odt / pptx
  utils/          DocumentError, path helpers, text escape helpers
  index.ts        public barrel
  cli.ts          bin: ai-document-parser
  types.ts        shared public types
test/integration/ roundtrip tests (import dist/)
scripts/          check-licenses, copy-assets, generate-scratch-fixtures, smoke-scratch
docs/             API.md, AGENT_GUIDE.md
fonts/            DejaVuSans for Cyrillic PDF write
```

Subpath exports: `ai-document-parser/parse|convert|edit|write`.

`scratch/` is gitignored — regenerate locally before smoke.

## Architecture

| Concern | Implementation |
|---------|----------------|
| Parse core formats | `src/parse/router.ts` (+ pdf / pptx / text-formats / docx-ooxml / spreadsheet) |
| Convert | `src/convert/converter.ts` + `emit.ts` (driven by parse structured result) |
| Surgical edit | `src/edit/*` |
| Write | `src/write/*` |
| Legacy `.doc` | `src/parse/legacy-doc.ts` |

Core formats: **docx, xlsx, xls, pptx, pdf, md, html, csv, legacy doc, odt, ods, odp, rtf, epub**. Unsupported (for now): legacy `.ppt`, rich HTML→PDF convert.

## Commands

```bash
npm install
npm run typecheck
npm run build
npm test
npm run test:smoke   # needs scratch/ fixtures first
npm run check-licenses
npm audit
```

Local scratch fixtures:

```bash
node scripts/generate-scratch-fixtures.mjs
npm run test:smoke
```

## TypeScript / ESM conventions

- `"type": "module"`, `module`/`moduleResolution`: **NodeNext**
- Relative imports **must** use `.js` extensions
- Root `tsconfig.json` is **`noEmit: true`** (includes `src/` + `test/` for IDE)
- Emit via **`tsconfig.build.json`** (`include: src` only) — used by `npm run build`
- `test/tsconfig.json` also exists for focused test typecheck (`types: ["node"]`)
- Do **not** point the IDE at `tsconfig.build.json` as the default project
- Domain errors: throw `DocumentError`; wrap unknowns via `wrapError`

## Testing

- Runner: **node:test** + **tsx**; tests import **`../../dist/...`**
- Primary suite: `test/integration/roundtrip.test.ts` (write/parse/edit + **`convertDocument matrix`**)
- Optional full smoke: `scripts/smoke-scratch.mjs` against `scratch/` (includes pptx)
- After source changes run `npm test` (rebuilds first)

## License policy

Read [LICENSES.md](LICENSES.md). Do not add non-permissive deps, abandoned `xlsx`, or `docxtemplater`.

## Code change guidelines

- Prefer small, focused diffs
- Wire new behavior through the public facade and update docs/README matrix together
- DOCX replace: use **structured paragraphs** as `oldString` sources
- Keep `jszip` for OOXML (DOCX edit + PPTX parse); `@e965/xlsx` for sheets; `docx` / `jspdf` for writers; `pdfjs-dist` for PDF text

## Key references

| File | Why |
|------|-----|
| [README.md](README.md) | Format matrix, quick start |
| [docs/API.md](docs/API.md) | Public API contracts |
| [docs/AGENT_GUIDE.md](docs/AGENT_GUIDE.md) | Read → edit → verify |
| [LICENSES.md](LICENSES.md) | License inventory + policy |
| [src/types.ts](src/types.ts) | Error codes and structured shapes |
| [src/parse/formats.ts](src/parse/formats.ts) | Clean-room reminder + core format list |
