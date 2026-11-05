import { execFile } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdtemp, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import type { DocumentWarning } from '../types.js';
import type { PdfWord } from './pdf-structure.js';

const moduleDir = path.dirname(fileURLToPath(import.meta.url));

const execFileAsync = promisify(execFile);

export const PDF_OCR_PAGE_CAP = 80;

export interface OcrRuntime {
  bin: string;
  libDir?: string;
  tessdataDir?: string;
  bundled: boolean;
}

export interface OcrPagesResult {
  words: PdfWord[];
  warnings: DocumentWarning[];
  usedPages: number[];
}

function ocrRootCandidates(): string[] {
  const out: string[] = [];
  const resourcesPath = (process as NodeJS.Process & { resourcesPath?: string }).resourcesPath;
  const resources = typeof resourcesPath === 'string' ? resourcesPath : '';
  if (resources) out.push(path.join(resources, 'ocr'));
  let dir = moduleDir;
  for (let i = 0; i < 8; i++) {
    out.push(path.join(dir, 'resources', 'ocr'));
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  out.push(path.join(process.cwd(), 'resources', 'ocr'));
  return out;
}

export function resolveOcrRuntime(): OcrRuntime {
  const seen = new Set<string>();
  for (const root of ocrRootCandidates()) {
    if (seen.has(root)) continue;
    seen.add(root);
    const binName = process.platform === 'win32' ? 'tesseract.exe' : 'tesseract';
    const bin = path.join(root, 'bin', binName);
    const tess = path.join(root, 'tessdata');
    if (existsSync(bin) && existsSync(path.join(tess, 'rus.traineddata'))) {
      const lib = path.join(root, 'lib');
      return {
        bin,
        libDir: existsSync(lib) ? lib : undefined,
        tessdataDir: tess,
        bundled: true,
      };
    }
  }
  return {
    bin: process.platform === 'win32' ? 'tesseract.exe' : 'tesseract',
    bundled: false,
  };
}

function ocrEnv(runtime: OcrRuntime): NodeJS.ProcessEnv {
  const env = Object.assign({}, process.env);
  if (runtime.libDir) {
    const prev = env.LD_LIBRARY_PATH || '';
    env.LD_LIBRARY_PATH = prev ? runtime.libDir + ':' + prev : runtime.libDir;
  }
  return env;
}

function run(
  cmd: string,
  args: string[],
  env: NodeJS.ProcessEnv,
  timeout: number
): Promise<string> {
  return execFileAsync(cmd, args, {
    encoding: 'utf8',
    env,
    timeout,
    maxBuffer: 16 * 1024 * 1024,
  }).then((out) => {
    if (typeof out === 'string') return out;
    if (out && typeof out.stdout === 'string') return out.stdout;
    return '';
  });
}

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

function unavailable(message: string): DocumentWarning {
  return { code: 'PDF_OCR_UNAVAILABLE', message };
}

export async function ocrPdfPages(pdfPath: string, pages: number[]): Promise<OcrPagesResult> {
  const warnings: DocumentWarning[] = [];
  const words: PdfWord[] = [];
  const usedPages: number[] = [];
  if (!pages.length) return { words, warnings, usedPages };

  const runtime = resolveOcrRuntime();
  const sorted = uniqueSorted(pages);
  let todo = sorted;
  if (sorted.length > PDF_OCR_PAGE_CAP) {
    todo = sorted.slice(0, PDF_OCR_PAGE_CAP);
    warnings.push({
      code: 'PDF_OCR_TRUNCATED',
      message:
        'OCR выполнен только для ' +
        PDF_OCR_PAGE_CAP +
        ' из ' +
        sorted.length +
        ' страниц без текстового слоя.',
    });
  }

  const tmp = await mkdtemp(path.join(os.tmpdir(), 'adp-ocr-'));
  try {
    let failed = false;
    for (const page of todo) {
      const base = path.join(tmp, 'page-' + page);
      try {
        await run(
          'pdftoppm',
          ['-png', '-r', '300', '-f', String(page), '-l', String(page), '-singlefile', pdfPath, base],
          process.env,
          60000
        );
        const img = base + '.png';
        if (!existsSync(img)) {
          throw new Error('pdftoppm не создал изображение страницы ' + page);
        }
        const args = [img, 'stdout', '-l', 'rus+eng', '--psm', '6'];
        if (runtime.tessdataDir) args.push('--tessdata-dir', runtime.tessdataDir);
        args.push('tsv');
        const tsv = await run(runtime.bin, args, ocrEnv(runtime), 120000);
        const pageWords = tsvToWords(tsv, page);
        if (pageWords.length) {
          for (const word of pageWords) words.push(word);
          usedPages.push(page);
        }
      } catch (err) {
        const code = (err as NodeJS.ErrnoException).code;
        const msg = err instanceof Error ? err.message : String(err);
        if (code === 'ENOENT' || msg.indexOf('ENOENT') !== -1) {
          const missing =
            msg.indexOf('pdftoppm') !== -1
              ? 'Не найден pdftoppm (пакет poppler-utils). Растеризация страниц для OCR от него зависит, в сборку Poppler не входит.'
              : 'OCR недоступен: нет resources/ocr (scripts/fetch-ocr-runtime.sh) и нет tesseract в PATH.';
          warnings.push(unavailable(missing));
          return { words, warnings, usedPages };
        }
        if (!failed) {
          failed = true;
          warnings.push({
            code: 'PDF_OCR_FAILED',
            message: 'Tesseract не смог распознать страницу ' + page + ': ' + msg.slice(0, 300),
          });
        }
      }
    }
  } finally {
    await rm(tmp, { recursive: true, force: true });
  }
  return { words, warnings, usedPages };
}

/** Recognize one raster image (png/jpeg/tiff/gif/webp/bmp) with the bundled Tesseract. */
export async function ocrImageFile(imagePath: string): Promise<{
  words: PdfWord[];
  warnings: DocumentWarning[];
}> {
  const warnings: DocumentWarning[] = [];
  const runtime = resolveOcrRuntime();
  try {
    const args = [imagePath, 'stdout', '-l', 'rus+eng', '--psm', '6'];
    if (runtime.tessdataDir) args.push('--tessdata-dir', runtime.tessdataDir);
    args.push('tsv');
    const tsv = await run(runtime.bin, args, ocrEnv(runtime), 120000);
    return { words: tsvToWords(tsv, 1), warnings };
  } catch (err) {
    const code = (err as NodeJS.ErrnoException).code;
    const msg = err instanceof Error ? err.message : String(err);
    if (code === 'ENOENT' || msg.indexOf('ENOENT') !== -1) {
      warnings.push(
        unavailable(
          'OCR недоступен: нет resources/ocr (scripts/fetch-ocr-runtime.sh) и нет tesseract в PATH.'
        )
      );
      warnings[warnings.length - 1] = {
        code: 'IMAGE_OCR_UNAVAILABLE',
        message: warnings[warnings.length - 1].message,
      };
      return { words: [], warnings };
    }
    warnings.push({
      code: 'IMAGE_OCR_FAILED',
      message: 'Tesseract не смог распознать изображение: ' + msg.slice(0, 300),
    });
    return { words: [], warnings };
  }
}
