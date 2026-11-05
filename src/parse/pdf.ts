import {
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  statSync,
  writeFileSync,
} from 'node:fs';
import { createRequire } from 'node:module';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { DocumentError, wrapError } from '../utils/errors.js';
import { assertMaxPdfPages } from '../utils/limits.js';
import { readFileLimited } from '../utils/zip.js';
import type { ContentBlock, DocumentWarning } from '../types.js';
import { emitMarkdownFromBlocks, flattenBlocksToParagraphs } from '../convert/blocks.js';
import { ensurePdfRuntimePolyfills } from './pdf-runtime-polyfills.js';
import { ocrPdfPages } from './pdf-ocr.js';
import {
  countLetters,
  linesToBlocks,
  PDF_EMPTY_PAGE_LETTERS,
  wordsToLines,
  type PdfWord,
} from './pdf-structure.js';

const moduleDir = path.dirname(fileURLToPath(import.meta.url));
const moduleFile = fileURLToPath(import.meta.url);

function asarUnpackedTwin(p: string): string {
  return p
    .replace(`${path.sep}app.asar${path.sep}`, `${path.sep}app.asar.unpacked${path.sep}`)
    .replace(/[/\\]app\.asar$/, `${path.sep}app.asar.unpacked`);
}

function pdfjsRootCandidates(): string[] {
  const out: string[] = [];
  const add = (root: string | undefined | null) => {
    if (!root) return;
    out.push(path.join(root, 'node_modules', 'pdfjs-dist'));
  };
  const resourcesPath = (process as NodeJS.Process & { resourcesPath?: string }).resourcesPath;
  const resources = typeof resourcesPath === 'string' ? resourcesPath : '';
  if (resources) {
    add(path.join(resources, 'app.asar.unpacked'));
    add(path.join(resources, 'app.asar'));
    add(resources);
  }
  for (const start of [moduleDir, asarUnpackedTwin(moduleDir)]) {
    let dir = start;
    for (let i = 0; i < 10; i++) {
      add(dir);
      const parent = path.dirname(dir);
      if (parent === dir) break;
      dir = parent;
    }
  }
  add(process.cwd());
  return [...new Set(out)];
}

/** `app.asar/...` is visible to Electron fs but not to import()/stat. */
function isAsarArchivePath(p: string): boolean {
  return p.split(/[/\\]/).includes('app.asar');
}

function isPdfjsRoot(dir: string): boolean {
  if (!dir || isAsarArchivePath(dir)) return false;
  return (
    existsSync(path.join(dir, 'package.json')) &&
    existsSync(path.join(dir, 'legacy', 'build', 'pdf.mjs'))
  );
}

function copyOutOfAsar(src: string, dest: string): void {
  mkdirSync(dest, { recursive: true });
  for (const name of readdirSync(src)) {
    const from = path.join(src, name);
    const to = path.join(dest, name);
    if (statSync(from).isDirectory()) copyOutOfAsar(from, to);
    else writeFileSync(to, readFileSync(from));
  }
}

/** Electron can read pdfjs inside asar, but dynamic import() cannot. */
function materializePdfjsFromAsar(asarRoot: string): string {
  let version = 'unknown';
  try {
    const pkg = JSON.parse(readFileSync(path.join(asarRoot, 'package.json'), 'utf8')) as {
      version?: string;
    };
    if (pkg.version) version = pkg.version;
  } catch {
    // keep unknown
  }
  const dest = path.join(os.tmpdir(), `adp-pdfjs-${version}`);
  const probe = path.join(dest, 'legacy', 'build', 'pdf.mjs');
  if (!existsSync(probe)) copyOutOfAsar(asarRoot, dest);
  if (!existsSync(probe)) {
    throw new DocumentError(
      'PARSE_FAILED',
      'pdfjs-dist is inside app.asar and could not be copied out for import()'
    );
  }
  return dest;
}

/** Resolve pdfjs-dist from asar.unpacked or materialize it out of app.asar. */
export function resolvePdfjsRoot(): string {
  for (const candidate of pdfjsRootCandidates()) {
    if (isPdfjsRoot(candidate)) return candidate;
  }
  const resourcesPath = (process as NodeJS.Process & { resourcesPath?: string }).resourcesPath;
  const resources = typeof resourcesPath === 'string' ? resourcesPath : '';
  const requireFrom = existsSync(
    path.join(resources, 'app.asar.unpacked', 'package.json')
  )
    ? path.join(resources, 'app.asar.unpacked', 'package.json')
    : existsSync(path.join(process.cwd(), 'package.json'))
      ? path.join(process.cwd(), 'package.json')
      : moduleFile;
  const nodeRequire = createRequire(requireFrom);
  const resolved = path.dirname(nodeRequire.resolve('pdfjs-dist/package.json'));
  if (!isAsarArchivePath(resolved)) return resolved;
  return materializePdfjsFromAsar(resolved);
}

function pdfjsAssetUrl(root: string, subdir: string): string {
  return pathToFileURL(path.join(root, subdir) + path.sep).href;
}

export interface PdfExtractResult {
  text: string;
  format: 'pdf';
  blocks: ContentBlock[];
  paragraphs: string[];
  warnings: DocumentWarning[];
}

interface PdfTextItem {
  str?: string;
  transform?: number[];
  width?: number;
  height?: number;
}

function itemHeight(item: PdfTextItem): number {
  if (typeof item.height === 'number' && item.height > 0) return item.height;
  const tr = item.transform;
  if (tr && tr.length >= 4) {
    const h = Math.hypot(tr[2], tr[3]);
    if (h > 0) return h;
  }
  return 10;
}

function itemsToWords(items: PdfTextItem[], page: number): PdfWord[] {
  const words: PdfWord[] = [];
  for (const item of items) {
    const text = typeof item.str === 'string' ? item.str.trim() : '';
    if (!text) continue;
    const tr = item.transform;
    const width =
      typeof item.width === 'number' && item.width > 0 ? item.width : Math.max(4, text.length * 4);
    words.push({
      text,
      x: tr && typeof tr[4] === 'number' ? tr[4] : 0,
      y: tr && typeof tr[5] === 'number' ? tr[5] : 0,
      width,
      height: itemHeight(item),
      page,
    });
  }
  return words;
}

/**
 * Text layer via pdfjs, OCR (bundled or system Tesseract) only for pages with almost no letters,
 * then geometry → ContentBlock markdown.
 */
export async function extractPdfText(input: string | Buffer): Promise<PdfExtractResult> {
  ensurePdfRuntimePolyfills();
  let pdf: {
    numPages: number;
    getPage: (n: number) => Promise<{ getTextContent: () => Promise<{ items: unknown[] }> }>;
    destroy?: () => void;
  } | null = null;
  try {
    const data =
      typeof input === 'string'
        ? new Uint8Array(await readFileLimited(input))
        : new Uint8Array(input);

    const pdfjsRoot = resolvePdfjsRoot();
    const pdfjs = await import(
      pathToFileURL(path.join(pdfjsRoot, 'legacy', 'build', 'pdf.mjs')).href
    );
    if (pdfjs.GlobalWorkerOptions) {
      pdfjs.GlobalWorkerOptions.workerSrc = pathToFileURL(
        path.join(pdfjsRoot, 'legacy', 'build', 'pdf.worker.mjs')
      ).href;
    }
    const loadingTask = pdfjs.getDocument({
      data,
      useSystemFonts: false,
      isEvalSupported: false,
      useWasm: false,
      verbosity: 0,
      cMapPacked: true,
      cMapUrl: pdfjsAssetUrl(pdfjsRoot, 'cmaps'),
      standardFontDataUrl: pdfjsAssetUrl(pdfjsRoot, 'standard_fonts'),
      disableWorker: true,
    } as Parameters<typeof pdfjs.getDocument>[0]);
    const loaded = await loadingTask.promise;
    pdf = loaded;
    assertMaxPdfPages(loaded.numPages);

    const pageWords = new Map<number, PdfWord[]>();
    const thinPages: number[] = [];
    for (let pageNum = 1; pageNum <= loaded.numPages; pageNum++) {
      const page = await loaded.getPage(pageNum);
      const content = await page.getTextContent();
      const words = itemsToWords(content.items as PdfTextItem[], pageNum);
      pageWords.set(pageNum, words);
      let letters = 0;
      for (const word of words) letters += countLetters(word.text);
      if (letters < PDF_EMPTY_PAGE_LETTERS) thinPages.push(pageNum);
    }

    const warnings: DocumentWarning[] = [];
    if (thinPages.length) {
      let pdfPath = typeof input === 'string' ? input : '';
      let tmpDir: string | undefined;
      try {
        if (!pdfPath) {
          tmpDir = await mkdtemp(path.join(os.tmpdir(), 'adp-pdf-ocr-'));
          pdfPath = path.join(tmpDir, 'input.pdf');
          await writeFile(pdfPath, input);
        }
        const ocr = await ocrPdfPages(pdfPath, thinPages);
        for (const warning of ocr.warnings) warnings.push(warning);
        if (ocr.usedPages.length) {
          const byPage = new Map<number, PdfWord[]>();
          for (const word of ocr.words) {
            const list = byPage.get(word.page);
            if (list) list.push(word);
            else byPage.set(word.page, [word]);
          }
          for (const pageNum of ocr.usedPages) {
            pageWords.set(pageNum, byPage.get(pageNum) || []);
          }
          warnings.push({
            code: 'PDF_OCR_USED',
            message:
              'Страницы без текстового слоя распознаны Tesseract (rus+eng, 300 dpi): ' +
              ocr.usedPages.join(', ') +
              '.',
          });
        }
      } finally {
        if (tmpDir) await rm(tmpDir, { recursive: true, force: true });
      }
    }

    const words: PdfWord[] = [];
    const pageNums = Array.from(pageWords.keys()).sort((a, b) => a - b);
    for (const pageNum of pageNums) {
      const list = pageWords.get(pageNum) || [];
      for (const word of list) words.push(word);
    }
    const structured = linesToBlocks(wordsToLines(words));
    for (const warning of structured.warnings) warnings.push(warning);
    const text = emitMarkdownFromBlocks(structured.blocks).trim();
    return {
      text,
      format: 'pdf',
      blocks: structured.blocks,
      paragraphs: flattenBlocksToParagraphs(structured.blocks),
      warnings,
    };
  } catch (err) {
    if (err instanceof DocumentError) throw err;
    throw wrapError('PARSE_FAILED', 'Failed to parse PDF', err);
  } finally {
    pdf?.destroy?.();
  }
}
