import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { after, before, test } from 'node:test';
import { inflateSync } from 'node:zlib';
import { jsPDF } from 'jspdf';
import type { OcrRuntime } from '../../src/types.js';
import { parseDocument } from '../../src/parse/index.js';
import { normalizeOcrOptions } from '../../src/parse/ocr-options.js';
import { ocrPdfPages } from '../../src/parse/pdf-ocr.js';
import { resolveOcrRuntime } from '../../src/parse/ocr-runtime.js';
import {
  encodeGrayPng,
  openPdfiumRasterizer,
  PDFIUM_MAX_PAGE_PIXELS,
  pdfiumScaleFor,
  resolvePdfiumWasmPath,
  type PdfPageRasterizer,
} from '../../src/parse/pdf-raster.js';
import { writePdf } from '../../src/write/pdf.js';

const PNG_SIGNATURE = '89504e470d0a1a0a';

let dir = '';
let textPdf = '';
let scanPdf = '';

function fakeRuntime(): OcrRuntime {
  return { bin: 'tesseract', popplerBin: 'pdftoppm', tessdataDir: '/tessdata', bundled: false, source: 'path' };
}

function tsvWord(text: string): string {
  return [
    'level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\twidth\theight\tconf\ttext',
    `5\t1\t1\t1\t1\t1\t10\t20\t30\t12\t99\t${text}`,
  ].join('\n');
}

function pngSize(png: Buffer): { width: number; height: number } {
  return { width: png.readUInt32BE(16), height: png.readUInt32BE(20) };
}

/** Real Tesseract is optional: set ADP_OCR_ROOT / ADP_TESSERACT or put tesseract on PATH. */
function tesseractAvailable(): boolean {
  try {
    const runtime = resolveOcrRuntime();
    execFileSync(runtime.bin, ['--version'], { stdio: 'ignore', timeout: 20_000 });
    return true;
  } catch {
    return false;
  }
}

/** Build an image-only PDF: render a text PDF with PDFium, then embed only the bitmap. */
async function makeScannedPdf(source: string, target: string): Promise<void> {
  const raster = await openPdfiumRasterizer(source);
  const pagePng = path.join(dir, 'scan-page.png');
  try {
    await raster.renderPage(1, 200, pagePng);
  } finally {
    raster.close();
  }
  const png = await readFile(pagePng);
  const doc = new jsPDF({ unit: 'mm', format: 'a4' });
  doc.addImage(new Uint8Array(png), 'PNG', 0, 0, 210, 297);
  await writeFile(target, Buffer.from(doc.output('arraybuffer')));
}

before(async () => {
  dir = await mkdtemp(path.join(os.tmpdir(), 'adp-scanned-pdf-'));
  textPdf = path.join(dir, 'text.pdf');
  scanPdf = path.join(dir, 'scan.pdf');
  await writePdf(textPdf, 'WATCHER scanned invoice 12345\n\nПроверка распознавания сканов', 'Scan');
  await makeScannedPdf(textPdf, scanPdf);
});

after(async () => {
  if (dir) await rm(dir, { recursive: true, force: true });
});

test('pdfium wasm resolves from the installed package and honors ADP_PDFIUM_WASM', () => {
  const wasm = resolvePdfiumWasmPath({});
  assert.ok(existsSync(wasm), wasm);
  assert.match(wasm, /pdfium\.wasm$/);
  assert.equal(resolvePdfiumWasmPath({ ADP_PDFIUM_WASM: ' /opt/pdfium.wasm ' }), '/opt/pdfium.wasm');
});

test('encodeGrayPng writes a valid grayscale PNG and handles row stride', () => {
  const width = 3;
  const height = 2;
  const stride = 4;
  const data = new Uint8Array([1, 2, 3, 99, 4, 5, 6, 99]);
  const png = encodeGrayPng(width, height, data);
  assert.equal(png.subarray(0, 8).toString('hex'), PNG_SIGNATURE);
  assert.deepEqual(pngSize(png), { width, height });
  assert.equal(png[24], 8);
  assert.equal(png[25], 0);
  const idatLen = png.readUInt32BE(33);
  assert.equal(png.toString('ascii', 37, 41), 'IDAT');
  const raw = inflateSync(png.subarray(41, 41 + idatLen));
  assert.deepEqual([...raw], [0, 1, 2, 3, 0, 4, 5, 6]);
  assert.equal(stride, data.length / height);
  assert.throws(() => encodeGrayPng(5, 2, data), /smaller/);
});

test('pdfium scale follows dpi and is capped for huge pages', () => {
  assert.equal(pdfiumScaleFor(144, 595, 842), 2);
  const scale = pdfiumScaleFor(2400, 14_400, 14_400);
  const pixels = Math.pow(14_400 * scale, 2);
  assert.ok(pixels <= PDFIUM_MAX_PAGE_PIXELS * 1.001, String(pixels));
});

test('pdfium rasterizes a PDF page at the requested dpi', async () => {
  const raster = await openPdfiumRasterizer(textPdf);
  try {
    const out = path.join(dir, 'page-150.png');
    await raster.renderPage(1, 150, out);
    const png = await readFile(out);
    assert.equal(png.subarray(0, 8).toString('hex'), PNG_SIGNATURE);
    const size = pngSize(png);
    // A4 = 595x842 pt -> ~1240x1754 px at 150 dpi.
    assert.ok(Math.abs(size.width - 1240) <= 2, String(size.width));
    assert.ok(Math.abs(size.height - 1754) <= 2, String(size.height));
    await assert.rejects(raster.renderPage(2, 150, path.join(dir, 'nope.png')), /out of range/);
  } finally {
    raster.close();
  }
});

test('scanned pdf fixture has no text layer', async () => {
  const parsed = await parseDocument(scanPdf, { ocr: false });
  assert.doesNotMatch(parsed.text, /WATCHER|12345|Проверка/);
});

test('scanned pdf pages go to tesseract as pdfium-rendered PNGs without pdftoppm', async () => {
  const commands: string[] = [];
  const seen: string[] = [];
  const result = await ocrPdfPages(scanPdf, [1], { languages: ['eng'], dpi: 100 }, {
    resolveRuntime: () => fakeRuntime(),
    run: async (cmd, args) => {
      commands.push(String(cmd));
      if (args.includes('--list-langs')) return { stdout: 'eng\n', stderr: '' };
      const png = await readFile(args[0]);
      seen.push(png.subarray(0, 8).toString('hex'));
      return { stdout: tsvWord('Rendered'), stderr: '' };
    },
  });
  assert.ok(!commands.some((cmd) => cmd.includes('pdftoppm')), commands.join(','));
  assert.deepEqual(seen, [PNG_SIGNATURE]);
  assert.deepEqual(result.usedPages, [1]);
  assert.deepEqual(result.warnings, []);
});

test('scanned pdf OCR falls back to pdftoppm when pdfium is unavailable', async () => {
  const created: string[] = [];
  const pdftoppmCalls: string[][] = [];
  const result = await ocrPdfPages('/tmp/doc.pdf', [1], { languages: ['eng'] }, {
    openPdfRasterizer: null,
    resolveRuntime: () => fakeRuntime(),
    existsSync: (filePath) => created.includes(filePath),
    run: async (cmd, args) => {
      if (args.includes('--list-langs')) return { stdout: 'eng\n', stderr: '' };
      if (String(cmd).includes('pdftoppm')) {
        pdftoppmCalls.push(args);
        const img = `${args[args.length - 1]}.png`;
        await writeFile(img, 'png');
        created.push(img);
        return { stdout: '', stderr: '' };
      }
      return { stdout: tsvWord('Poppler'), stderr: '' };
    },
  });
  assert.equal(pdftoppmCalls.length, 1);
  assert.deepEqual(result.words.map((w) => w.text), ['Poppler']);
  const fallback = result.warnings.find((w) => w.code === 'PDF_OCR_RASTER_FALLBACK');
  assert.ok(fallback, JSON.stringify(result.warnings));
  assert.match(fallback.message, /PDFium rasterizer is not available/);
});

test('scanned pdf OCR falls back to pdftoppm only for pages pdfium fails to render', async () => {
  const created: string[] = [];
  const pdftoppmPages: string[] = [];
  let closed = 0;
  const flaky: PdfPageRasterizer = {
    async renderPage(page, _dpi, outPath) {
      if (page === 2) throw new Error('bad page tree');
      await writeFile(outPath, 'png');
      return outPath;
    },
    close() {
      closed += 1;
    },
  };
  const result = await ocrPdfPages('/tmp/doc.pdf', [1, 2, 3], { languages: ['eng'] }, {
    openPdfRasterizer: async () => flaky,
    resolveRuntime: () => fakeRuntime(),
    existsSync: (filePath) => created.includes(filePath),
    run: async (cmd, args) => {
      if (args.includes('--list-langs')) return { stdout: 'eng\n', stderr: '' };
      if (String(cmd).includes('pdftoppm')) {
        pdftoppmPages.push(args[args.indexOf('-f') + 1]);
        const img = `${args[args.length - 1]}.png`;
        await writeFile(img, 'png');
        created.push(img);
        return { stdout: '', stderr: '' };
      }
      return { stdout: tsvWord(path.basename(args[0])), stderr: '' };
    },
  });
  assert.deepEqual(pdftoppmPages, ['2']);
  assert.deepEqual(result.usedPages, [1, 2, 3]);
  assert.equal(closed, 1);
  const fallback = result.warnings.filter((w) => w.code === 'PDF_OCR_RASTER_FALLBACK');
  assert.equal(fallback.length, 1);
  assert.match(fallback[0].message, /page 2: bad page tree/);
});

test('scanned pdf OCR reports both renderers when pdfium and pdftoppm are missing', async () => {
  const result = await ocrPdfPages('/tmp/doc.pdf', [1], { languages: ['eng'] }, {
    openPdfRasterizer: async () => {
      throw new Error('wasm load failed');
    },
    resolveRuntime: () => fakeRuntime(),
    run: async (_cmd, args) => {
      if (args.includes('--list-langs')) return { stdout: 'eng\n', stderr: '' };
      const err = new Error('spawn pdftoppm ENOENT');
      (err as NodeJS.ErrnoException).code = 'ENOENT';
      throw err;
    },
  });
  assert.equal(result.words.length, 0);
  assert.equal(result.warnings.length, 1);
  assert.equal(result.warnings[0].code, 'PDF_OCR_UNAVAILABLE');
  assert.match(result.warnings[0].message, /pdftoppm was not found/);
  assert.match(result.warnings[0].message, /wasm load failed/);
});

test('pdfRasterizer pdfium does not fall back and pdftoppm skips pdfium', async () => {
  let spawned = 0;
  const strict = await ocrPdfPages('/tmp/doc.pdf', [1], { languages: ['eng'], pdfRasterizer: 'pdfium' }, {
    openPdfRasterizer: async () => {
      throw new Error('not a PDF');
    },
    resolveRuntime: () => fakeRuntime(),
    run: async (cmd, args) => {
      if (!args.includes('--list-langs')) spawned += 1;
      return { stdout: 'eng\n', stderr: '' };
    },
  });
  assert.equal(spawned, 0);
  assert.equal(strict.warnings[0]?.code, 'PDF_OCR_UNAVAILABLE');
  assert.match(strict.warnings[0]?.message ?? '', /PDFium .*not a PDF/);

  let opened = 0;
  const created: string[] = [];
  const poppler = await ocrPdfPages('/tmp/doc.pdf', [1], { languages: ['eng'], pdfRasterizer: 'pdftoppm' }, {
    openPdfRasterizer: async () => {
      opened += 1;
      throw new Error('should not open');
    },
    resolveRuntime: () => fakeRuntime(),
    existsSync: (filePath) => created.includes(filePath),
    run: async (cmd, args) => {
      if (args.includes('--list-langs')) return { stdout: 'eng\n', stderr: '' };
      if (String(cmd).includes('pdftoppm')) {
        const img = `${args[args.length - 1]}.png`;
        await writeFile(img, 'png');
        created.push(img);
        return { stdout: '', stderr: '' };
      }
      return { stdout: tsvWord('P'), stderr: '' };
    },
  });
  assert.equal(opened, 0);
  assert.deepEqual(poppler.usedPages, [1]);
  assert.ok(!poppler.warnings.some((w) => w.code === 'PDF_OCR_RASTER_FALLBACK'));
});

test('pdfRasterizer option defaults to auto and ignores unknown values', () => {
  assert.equal(normalizeOcrOptions(undefined).pdfRasterizer, 'auto');
  assert.equal(normalizeOcrOptions({ pdfRasterizer: 'pdftoppm' }).pdfRasterizer, 'pdftoppm');
  assert.equal(
    normalizeOcrOptions({ pdfRasterizer: 'ghostscript' as unknown as 'auto' }).pdfRasterizer,
    'auto'
  );
});

test('scanned pdf with Russian and English text is recognized end to end without pdftoppm', async (t) => {
  if (!tesseractAvailable()) {
    t.skip('Tesseract not available (set ADP_OCR_ROOT or ADP_TESSERACT)');
    return;
  }
  const started = Date.now();
  const parsed = await parseDocument(scanPdf, {
    ocr: { languages: ['rus', 'eng'], popplerBin: path.join(dir, 'no-such-pdftoppm.exe') },
  });
  const seconds = (Date.now() - started) / 1000;
  t.diagnostic('scanned page OCR seconds: ' + seconds.toFixed(2));
  assert.match(parsed.text, /WATCHER/);
  assert.match(parsed.text, /12345/);
  assert.match(parsed.text, /Проверка распознавания сканов/);
  const codes = (parsed.warnings ?? []).map((w) => w.code);
  assert.ok(codes.includes('PDF_OCR_USED'), codes.join(','));
  assert.ok(!codes.includes('PDF_OCR_RASTER_FALLBACK'), codes.join(','));
  assert.ok(!codes.includes('PDF_OCR_UNAVAILABLE'), codes.join(','));
});
