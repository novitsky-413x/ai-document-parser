import { readFile, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { deflateSync } from 'node:zlib';

/**
 * PDF page rasterization for OCR.
 *
 * Primary path: PDFium compiled to WebAssembly (`@hyzyla/pdfium`, MIT wrapper;
 * PDFium itself is BSD-3-Clause / Apache-2.0). Pure JS + WASM, no native binary,
 * so scanned-PDF OCR works out of the box wherever Tesseract is available.
 * The `pdftoppm` (Poppler) path in `pdf-ocr.ts` stays as a fallback.
 */

/** Upper bound for one rendered page (gray, 1 byte/pixel) to keep WASM memory sane. */
export const PDFIUM_MAX_PAGE_PIXELS = 60_000_000;

/** One opened PDF that can render pages to grayscale PNG files. */
export interface PdfPageRasterizer {
  /** Render 1-based `page` at `dpi` into `outPath` (PNG). Returns the written path. */
  renderPage(page: number, dpi: number, outPath: string): Promise<string>;
  close(): void;
}

/** Opens a PDF for rasterization. Throws when the engine or the document cannot be loaded. */
export type PdfRasterizerFactory = (pdfPath: string) => Promise<PdfPageRasterizer>;

interface PdfiumPageLike {
  render(options: {
    scale: number;
    colorSpace: 'Gray';
    render: 'bitmap';
    renderFormFields?: boolean;
  }): Promise<{ width: number; height: number; data: Uint8Array }>;
  getOriginalSize(): { originalWidth: number; originalHeight: number };
  close(): void;
}

interface PdfiumDocumentLike {
  getPageCount(): number;
  getPage(index: number): PdfiumPageLike;
  destroy(): void;
}

interface PdfiumLibraryLike {
  loadDocument(buff: Uint8Array, password?: string): Promise<PdfiumDocumentLike>;
}

let libraryPromise: Promise<PdfiumLibraryLike> | undefined;

/**
 * Locate `pdfium.wasm`. `ADP_PDFIUM_WASM` overrides; otherwise resolve it from the
 * installed `@hyzyla/pdfium` package. Reading the bytes ourselves (instead of letting
 * the package fetch a URL next to its bundle) keeps loading working inside Electron
 * asar archives and when the host bundles its main process.
 */
export function resolvePdfiumWasmPath(env: NodeJS.ProcessEnv = process.env): string {
  const override = (env.ADP_PDFIUM_WASM || '').trim();
  if (override) return override;
  const require = createRequire(import.meta.url);
  return require.resolve('@hyzyla/pdfium/pdfium.wasm');
}

async function loadPdfiumLibrary(): Promise<PdfiumLibraryLike> {
  const wasmPath = resolvePdfiumWasmPath();
  const wasmBinary = await readFile(wasmPath);
  const mod = (await import('@hyzyla/pdfium')) as unknown as {
    PDFiumLibrary: { init(options: { wasmBinary: Uint8Array }): Promise<PdfiumLibraryLike> };
  };
  return mod.PDFiumLibrary.init({ wasmBinary: new Uint8Array(wasmBinary) });
}

/** Shared PDFium instance (WASM init costs ~100 ms; reuse it across calls). */
export function getPdfiumLibrary(): Promise<PdfiumLibraryLike> {
  if (!libraryPromise) {
    libraryPromise = loadPdfiumLibrary().catch((err) => {
      libraryPromise = undefined;
      throw err;
    });
  }
  return libraryPromise;
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buf: Uint8Array): number {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function pngChunk(type: string, data: Uint8Array): Buffer {
  const out = Buffer.alloc(12 + data.length);
  out.writeUInt32BE(data.length, 0);
  out.write(type, 4, 'ascii');
  Buffer.from(data.buffer, data.byteOffset, data.byteLength).copy(out, 8);
  out.writeUInt32BE(crc32(out.subarray(4, 8 + data.length)), 8 + data.length);
  return out;
}

/** Encode an 8-bit grayscale bitmap (row stride may exceed width) as PNG. */
export function encodeGrayPng(width: number, height: number, data: Uint8Array): Buffer {
  if (width <= 0 || height <= 0) throw new Error('empty bitmap');
  const stride = Math.floor(data.length / height);
  if (stride < width) throw new Error('bitmap is smaller than ' + width + 'x' + height);
  const raw = Buffer.alloc((width + 1) * height);
  for (let y = 0; y < height; y++) {
    const dst = y * (width + 1);
    raw[dst] = 0; // filter: none
    raw.set(data.subarray(y * stride, y * stride + width), dst + 1);
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 0; // grayscale
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', deflateSync(raw, { level: 3 })),
    pngChunk('IEND', new Uint8Array(0)),
  ]);
}

/** Scale (PDF points → pixels) for `dpi`, reduced if the page would exceed the pixel cap. */
export function pdfiumScaleFor(dpi: number, widthPt: number, heightPt: number): number {
  let scale = dpi / 72;
  const pixels = widthPt * scale * heightPt * scale;
  if (pixels > PDFIUM_MAX_PAGE_PIXELS && pixels > 0) {
    scale = scale * Math.sqrt(PDFIUM_MAX_PAGE_PIXELS / pixels);
  }
  return scale;
}

/** Default factory: open `pdfPath` with PDFium (WASM). */
export const openPdfiumRasterizer: PdfRasterizerFactory = async (pdfPath) => {
  const library = await getPdfiumLibrary();
  const bytes = await readFile(pdfPath);
  const doc = await library.loadDocument(new Uint8Array(bytes));
  let closed = false;
  return {
    async renderPage(page, dpi, outPath) {
      const count = doc.getPageCount();
      if (page < 1 || page > count) {
        throw new Error('PDFium: page ' + page + ' is out of range (1-' + count + ')');
      }
      const pdfPage = doc.getPage(page - 1);
      try {
        const size = pdfPage.getOriginalSize();
        const scale = pdfiumScaleFor(dpi, size.originalWidth, size.originalHeight);
        const bitmap = await pdfPage.render({
          scale,
          colorSpace: 'Gray',
          render: 'bitmap',
          renderFormFields: false,
        });
        await writeFile(outPath, encodeGrayPng(bitmap.width, bitmap.height, bitmap.data));
        return outPath;
      } finally {
        pdfPage.close();
      }
    },
    close() {
      if (closed) return;
      closed = true;
      doc.destroy();
    },
  };
};
