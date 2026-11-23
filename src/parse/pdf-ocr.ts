import { execFile } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdtemp, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { promisify } from 'node:util';
import type {
  DocumentWarning,
  OcrOptions,
  OcrProvider,
  OcrRecognizeRequest,
  OcrRecognizeResult,
  OcrRuntime,
  OcrWord,
} from '../types.js';
import type { PdfWord } from './pdf-structure.js';
import {
  DEFAULT_PDFTOPPM_TIMEOUT_MS,
  languageFlag,
  normalizeOcrOptions,
  PDF_OCR_PAGE_CAP,
  type ResolvedOcrOptions,
} from './ocr-options.js';
import {
  hasAllOcrLanguages,
  listInstalledOcrLanguages,
  missingOcrLanguages,
  ocrLangMissingWarning,
  ocrProcessEnv,
  resolveOcrRuntime,
  type OcrCommandRunner,
  type ResolveOcrRuntimeOptions,
} from './ocr-runtime.js';

export { PDF_OCR_PAGE_CAP, DEFAULT_OCR_LANGUAGES, normalizeOcrOptions } from './ocr-options.js';
export { resolveOcrRuntime, ocrProcessEnv, listInstalledOcrLanguages } from './ocr-runtime.js';
export type { OcrRuntime } from '../types.js';

const execFileAsync = promisify(execFile);

export interface OcrPagesResult {
  words: PdfWord[];
  warnings: DocumentWarning[];
  usedPages: number[];
}

/** Test / host seam for spawning Tesseract and pdftoppm. */
export interface OcrEngineHooks {
  run?: OcrCommandRunner;
  existsSync?: (filePath: string) => boolean;
  resolveRuntime?: (input?: ResolveOcrRuntimeOptions) => OcrRuntime;
}

const defaultRun: OcrCommandRunner = async (cmd, args, env, timeout) => {
  const out = await execFileAsync(cmd, args, {
    encoding: 'utf8',
    env,
    timeout,
    maxBuffer: 16 * 1024 * 1024,
  });
  if (typeof out === 'string') return { stdout: out, stderr: '' };
  return {
    stdout: typeof out.stdout === 'string' ? out.stdout : '',
    stderr: typeof out.stderr === 'string' ? out.stderr : '',
  };
};

function uniqueSorted(pages: number[]): number[] {
  const sorted = pages.slice().sort((a, b) => a - b);
  const out: number[] = [];
  for (const page of sorted) {
    if (out[out.length - 1] !== page) out.push(page);
  }
  return out;
}

function tsvToWords(tsv: string, page: number): PdfWord[] {
  const rows = tsv.split(/\n/);
  const words: PdfWord[] = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row || (i === 0 && row.indexOf('level') === 0)) continue;
    const cols = row.split('\t');
    if (cols.length < 12 || cols[0] !== '5') continue;
    const text = cols.slice(11).join('\t').trim();
    if (!text) continue;
    const left = Number(cols[6]);
    const top = Number(cols[7]);
    const width = Number(cols[8]);
    const height = Number(cols[9]);
    if (!isFinite(left) || !isFinite(top)) continue;
    words.push({
      text,
      x: left,
      y: -top,
      width: width > 0 ? width : 1,
      height: height > 0 ? height : 10,
      page,
    });
  }
  return words;
}

function isMissingBinary(err: unknown): boolean {
  const code = (err as NodeJS.ErrnoException).code;
  const msg = err instanceof Error ? err.message : String(err);
  return code === 'ENOENT' || msg.indexOf('ENOENT') !== -1;
}

function errorMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

function mentions(err: unknown, token: string): boolean {
  return errorMessage(err).toLowerCase().indexOf(token.toLowerCase()) !== -1;
}

export function tesseractRecognizeArgs(request: OcrRecognizeRequest): string[] {
  const args = [request.imagePath, 'stdout', '-l', languageFlag(request.languages), '--psm', String(request.psm)];
  if (request.oem != null) args.push('--oem', String(request.oem));
  if (request.runtime.tessdataDir) args.push('--tessdata-dir', request.runtime.tessdataDir);
  args.push('tsv');
  return args;
}

export function createTesseractOcrProvider(hooks: OcrEngineHooks = {}): OcrProvider {
  const run = hooks.run ?? defaultRun;
  return {
    async recognize(request: OcrRecognizeRequest): Promise<OcrRecognizeResult> {
      const env = ocrProcessEnv(request.runtime);
      const out = await run(request.runtime.bin, tesseractRecognizeArgs(request), env, request.timeoutMs);
      return { words: tsvToWords(out.stdout, request.page) };
    },
  };
}

export const tesseractOcrProvider: OcrProvider = createTesseractOcrProvider();

async function checkRequestedLanguages(
  runtime: OcrRuntime,
  options: ResolvedOcrOptions,
  hooks: OcrEngineHooks
): Promise<DocumentWarning | undefined> {
  const run = hooks.run ?? defaultRun;
  const exists = hooks.existsSync ?? existsSync;
  const listed = await listInstalledOcrLanguages(runtime, run, options.timeoutMs);
  let installed = listed.languages;
  if (!installed.length && runtime.tessdataDir) {
    installed = options.languages.filter((lang) =>
      hasAllOcrLanguages(runtime.tessdataDir!, [lang], exists)
    );
  }
  if (installed.length) {
    const missing = missingOcrLanguages(options.languages, installed);
    if (missing.length) return ocrLangMissingWarning(missing, installed);
    return undefined;
  }
  return listed.warning;
}

async function rasterizePdfPage(
  pdfPath: string,
  page: number,
  tmp: string,
  options: ResolvedOcrOptions,
  runtime: OcrRuntime,
  hooks: OcrEngineHooks
): Promise<string> {
  const run = hooks.run ?? defaultRun;
  const exists = hooks.existsSync ?? existsSync;
  const poppler = runtime.popplerBin || (process.platform === 'win32' ? 'pdftoppm.exe' : 'pdftoppm');
  const base = path.join(tmp, 'page-' + page);
  await run(
    poppler,
    ['-png', '-r', String(options.dpi), '-f', String(page), '-l', String(page), '-singlefile', pdfPath, base],
    ocrProcessEnv(runtime),
    DEFAULT_PDFTOPPM_TIMEOUT_MS
  );
  const img = base + '.png';
  if (!exists(img)) {
    throw new Error('pdftoppm did not create an image for page ' + page);
  }
  return img;
}

export async function ocrPdfPages(
  pdfPath: string,
  pages: number[],
  ocr?: boolean | OcrOptions,
  hooks: OcrEngineHooks = {}
): Promise<OcrPagesResult> {
  const warnings: DocumentWarning[] = [];
  const words: PdfWord[] = [];
  const usedPages: number[] = [];
  const options = normalizeOcrOptions(ocr);
  if (options.enabled === false || !pages.length) return { words, warnings, usedPages };

  const resolve = hooks.resolveRuntime ?? resolveOcrRuntime;
  const runtime = resolve({
    languages: options.languages,
    bin: options.bin,
    tessdataDir: options.tessdataDir,
    popplerBin: options.popplerBin,
    existsSync: hooks.existsSync,
  });

  const provider = options.provider ?? createTesseractOcrProvider(hooks);
  const usingCustom = Boolean(options.provider);

  if (!usingCustom) {
    const langWarning = await checkRequestedLanguages(runtime, options, hooks);
    if (langWarning) {
      warnings.push(langWarning);
      if (langWarning.code === 'OCR_LANG_MISSING') return { words, warnings, usedPages };
    }
  }

  const sorted = uniqueSorted(pages);
  let todo = sorted;
  if (sorted.length > options.pageCap) {
    todo = sorted.slice(0, options.pageCap);
    warnings.push({
      code: 'PDF_OCR_TRUNCATED',
      message:
        'OCR ran only on the first ' +
        options.pageCap +
        ' of ' +
        sorted.length +
        ' pages that need recognition.',
    });
  }

  const tmp = await mkdtemp(path.join(os.tmpdir(), 'adp-ocr-'));
  try {
    for (const page of todo) {
      try {
        const img = await rasterizePdfPage(pdfPath, page, tmp, options, runtime, hooks);
        const recognized = await provider.recognize({
          imagePath: img,
          page,
          languages: options.languages,
          psm: options.psm,
          oem: options.oem,
          timeoutMs: options.timeoutMs,
          runtime,
        });
        if (recognized.warnings) {
          for (const warning of recognized.warnings) warnings.push(warning);
        }
        const pageWords = recognized.words as PdfWord[];
        if (pageWords.length) {
          for (const word of pageWords) words.push({ ...word, page: word.page || page });
          usedPages.push(page);
        }
      } catch (err) {
        if (isMissingBinary(err)) {
          const missingPoppler = mentions(err, 'pdftoppm');
          warnings.push({
            code: 'PDF_OCR_UNAVAILABLE',
            message: missingPoppler
              ? 'pdftoppm was not found (Poppler). Scanned-PDF OCR needs it; Poppler is not bundled. Image OCR works with Tesseract alone. Set popplerBin, ADP_PDFTOPPM, or install poppler-utils.'
              : 'OCR is unavailable: no bundled resources/ocr (run `node scripts/fetch-ocr-runtime.mjs`) and no tesseract on PATH. Set bin / ADP_TESSERACT to use a custom binary.',
          });
          // A missing runtime cannot recover on later pages.
          return { words, warnings, usedPages };
        }
        warnings.push({
          code: 'PDF_OCR_FAILED',
          message: 'Tesseract failed to recognize page ' + page + ': ' + errorMessage(err).slice(0, 300),
        });
        // Keep going so a single bad page does not drop the rest.
      }
    }
  } finally {
    await rm(tmp, { recursive: true, force: true });
  }
  return { words, warnings, usedPages };
}

/** Recognize one raster image (png/jpeg/tiff/gif/webp/bmp). */
export async function ocrImageFile(
  imagePath: string,
  ocr?: boolean | OcrOptions,
  hooks: OcrEngineHooks = {}
): Promise<{
  words: PdfWord[];
  warnings: DocumentWarning[];
}> {
  const warnings: DocumentWarning[] = [];
  const options = normalizeOcrOptions(ocr);
  if (options.enabled === false) return { words: [], warnings };

  const resolve = hooks.resolveRuntime ?? resolveOcrRuntime;
  const runtime = resolve({
    languages: options.languages,
    bin: options.bin,
    tessdataDir: options.tessdataDir,
    popplerBin: options.popplerBin,
    existsSync: hooks.existsSync,
  });

  const provider = options.provider ?? createTesseractOcrProvider(hooks);
  if (!options.provider) {
    const langWarning = await checkRequestedLanguages(runtime, options, hooks);
    if (langWarning) {
      warnings.push(langWarning);
      if (langWarning.code === 'OCR_LANG_MISSING') return { words: [], warnings };
    }
  }

  try {
    const recognized = await provider.recognize({
      imagePath,
      page: 1,
      languages: options.languages,
      psm: options.psm,
      oem: options.oem,
      timeoutMs: options.timeoutMs,
      runtime,
    });
    if (recognized.warnings) {
      for (const warning of recognized.warnings) warnings.push(warning);
    }
    return { words: recognized.words as PdfWord[], warnings };
  } catch (err) {
    if (isMissingBinary(err)) {
      warnings.push({
        code: 'IMAGE_OCR_UNAVAILABLE',
        message:
          'OCR is unavailable: no bundled resources/ocr (run `node scripts/fetch-ocr-runtime.mjs`) and no tesseract on PATH. Set bin / ADP_TESSERACT to use a custom binary.',
      });
      return { words: [], warnings };
    }
    warnings.push({
      code: 'IMAGE_OCR_FAILED',
      message: 'Tesseract failed to recognize the image: ' + errorMessage(err).slice(0, 300),
    });
    return { words: [], warnings };
  }
}

export function ocrUsedMessage(
  kind: 'pdf-auto' | 'pdf-force' | 'image',
  languages: readonly string[],
  dpi: number,
  pages?: number[]
): DocumentWarning {
  const langs = languageFlag(languages);
  if (kind === 'image') {
    return {
      code: 'IMAGE_OCR_USED',
      message: 'Image text recognized with Tesseract (' + langs + ').',
    };
  }
  const prefix =
    kind === 'pdf-force'
      ? 'Pages were recognized with Tesseract ('
      : 'Pages without a text layer were recognized with Tesseract (';
  return {
    code: 'PDF_OCR_USED',
    message: prefix + langs + ', ' + dpi + ' dpi): ' + (pages ?? []).join(', ') + '.',
  };
}

// Re-export word type alias used by hosts that imported PdfWord-shaped OCR output.
export type { OcrWord };
