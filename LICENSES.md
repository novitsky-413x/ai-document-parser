# Third-party licenses

This package is licensed under the **MIT License** (see [LICENSE](LICENSE)).

## Bundled assets

| Component | Upstream | License | Notes |
|-----------|----------|---------|-------|
| DejaVu Sans | DejaVu Fonts | SIL Open Font License | Bundled in `fonts/` for Cyrillic PDF output |

## Direct dependencies (runtime)

| Package | License | Role |
|---------|---------|------|
| `pdfjs-dist` | Apache-2.0 | PDF text extraction |
| `@hyzyla/pdfium` | MIT (bundled PDFium WebAssembly: BSD-3-Clause / Apache-2.0) | Render scanned PDF pages for OCR |
| `jszip` | MIT OR GPL-3.0-or-later (used as MIT) | OOXML ZIP: DOCX surgical edit, PPTX parse, ZIP format refine |
| `docx` | MIT | Create DOCX from scratch |
| `jspdf` | MIT | Plain-text PDF write |
| `cfb` | Apache-2.0 | Legacy `.doc` OLE |
| `@e965/xlsx` | Apache-2.0 | XLSX / XLS read, write, cell patch |

## Policy

- Direct dependencies must be permissive (MIT / Apache-2.0 / BSD / ISC / dual including MIT).
- CI runs `npm run check-licenses` (allowlist + deny-list: abandoned `xlsx`, `docxtemplater`).
- Install scripts: `esbuild` allowed (tsx); `core-js` denied (optional jspdf polyfill noise).
