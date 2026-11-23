# OCR runtime (not committed)

`ai-document-parser` implements Tesseract OCR for scanned PDFs (after the pdf.js text layer) and for raster images. This directory is where a bundled runtime lives:

```
bin/tesseract(.exe)
tessdata/<lang>.traineddata
bin/pdftoppm(.exe)   (optional, used for scanned PDFs)
lib/                 (optional; LD_LIBRARY_PATH / DYLD_LIBRARY_PATH / PATH)
```

Binaries and `*.traineddata` files are gitignored — do not commit them. Fetch official language data (any Tesseract language id) and optionally copy a system binary with:

```
node scripts/fetch-ocr-runtime.mjs --dry-run
node scripts/fetch-ocr-runtime.mjs --langs rus,eng
node scripts/fetch-ocr-runtime.mjs --langs deu,eng --set best
```

`--set` is `fast` (default), `best`, or `standard`. `--lang-url` may supply a custom `{lang}` URL template. `--bin download --bin-url` expects a raw tesseract binary, not an archive.

The bundled runtime is accepted when `bin/tesseract` exists and **every language requested by the caller** has a traineddata file (the default request is `rus` and `eng`). Hosts such as WATCHER can pack this folder as `extraResources` (`resources/ocr` → `ocr`) so `process.resourcesPath/ocr` resolves at runtime.

Resolution order used by `resolveOcrRuntime()`:

1. `ocr.bin` / `ocr.tessdataDir` / `ocr.popplerBin`
2. `ADP_TESSERACT`, `ADP_TESSDATA` or `TESSDATA_PREFIX`, `ADP_PDFTOPPM`
3. this bundled folder
4. `tesseract` and `pdftoppm` on PATH

If the runtime is missing, the parser reports `PDF_OCR_UNAVAILABLE` / `IMAGE_OCR_UNAVAILABLE` and still returns any text layer that pdf.js already extracted. Scanned-PDF OCR needs `pdftoppm` (Poppler); it is not downloaded here. Image OCR works with Tesseract alone.

Callers can also pass a custom `OcrProvider` and skip the CLI runtime entirely.
