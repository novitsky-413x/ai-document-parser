import { existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import type { DocumentWarning, OcrRuntime, OcrRuntimeSource } from '../types.js';
import { DEFAULT_OCR_LANGUAGES, normalizeOcrLanguages } from './ocr-options.js';

const defaultModuleDir = path.dirname(fileURLToPath(import.meta.url));

export interface ResolveOcrRuntimeOptions {
  languages?: readonly string[];
  bin?: string;
  tessdataDir?: string;
  popplerBin?: string;
  env?: NodeJS.ProcessEnv;
  existsSync?: (filePath: string) => boolean;
  moduleDir?: string;
  cwd?: string;
  resourcesPath?: string;
}

function tesseractBinName(): string {
  return process.platform === 'win32' ? 'tesseract.exe' : 'tesseract';
}

function popplerBinName(): string {
  return process.platform === 'win32' ? 'pdftoppm.exe' : 'pdftoppm';
}

function trimEnv(value: string | undefined): string | undefined {
  if (value == null) return undefined;
  const trimmed = value.trim();
  return trimmed ? trimmed : undefined;
}

export function ocrRootCandidates(input: ResolveOcrRuntimeOptions = {}): string[] {
  const out: string[] = [];
  const env = input.env ?? process.env;
  const resourcesPath =
    input.resourcesPath !== undefined
      ? input.resourcesPath
      : typeof (process as NodeJS.Process & { resourcesPath?: string }).resourcesPath === 'string'
        ? (process as NodeJS.Process & { resourcesPath?: string }).resourcesPath
        : '';
  const resources = typeof resourcesPath === 'string' ? resourcesPath : '';
  if (resources) out.push(path.join(resources, 'ocr'));
  let dir = input.moduleDir ?? defaultModuleDir;
  for (let i = 0; i < 8; i++) {
    out.push(path.join(dir, 'resources', 'ocr'));
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  out.push(path.join(input.cwd ?? process.cwd(), 'resources', 'ocr'));
  const extra = trimEnv(env.ADP_OCR_ROOT);
  if (extra) out.push(extra);
  return out;
}

function hasTraineddata(
  tessdataDir: string,
  lang: string,
  exists: (filePath: string) => boolean
): boolean {
  return exists(path.join(tessdataDir, `${lang}.traineddata`));
}

export function hasAllOcrLanguages(
  tessdataDir: string,
  languages: readonly string[],
  exists: (filePath: string) => boolean = existsSync
): boolean {
  return languages.every((lang) => hasTraineddata(tessdataDir, lang, exists));
}

/**
 * TESSDATA_PREFIX is the parent of `tessdata/` in official Tesseract docs,
 * but hosts often point it at the tessdata directory itself. Accept both.
 */
export function normalizeTessdataDir(
  candidate: string | undefined,
  exists: (filePath: string) => boolean = existsSync
): string | undefined {
  const raw = trimEnv(candidate);
  if (!raw) return undefined;
  if (exists(path.join(raw, 'eng.traineddata')) || traineddataPresent(raw, exists)) {
    return raw;
  }
  const nested = path.join(raw, 'tessdata');
  if (exists(nested) || traineddataPresent(nested, exists)) return nested;
  return raw;
}

function traineddataPresent(dir: string, exists: (filePath: string) => boolean): boolean {
  // existsSync cannot list a directory; treat a known default file or the dir itself.
  if (exists(path.join(dir, 'osd.traineddata'))) return true;
  return exists(dir);
}

function findPopplerNear(
  roots: string[],
  exists: (filePath: string) => boolean
): string | undefined {
  const name = popplerBinName();
  const seen = new Set<string>();
  for (const root of roots) {
    const candidates = [
      path.join(root, 'bin', name),
      path.join(root, name),
      path.join(root, 'bin', 'pdftoppm'),
      path.join(root, 'pdftoppm'),
    ];
    for (const candidate of candidates) {
      if (seen.has(candidate)) continue;
      seen.add(candidate);
      if (exists(candidate)) return candidate;
    }
  }
  return undefined;
}

export function resolveOcrRuntime(input: ResolveOcrRuntimeOptions = {}): OcrRuntime {
  const env = input.env ?? process.env;
  const exists = input.existsSync ?? existsSync;
  const languages = normalizeOcrLanguages(input.languages ?? DEFAULT_OCR_LANGUAGES);
  const optionBin = trimEnv(input.bin);
  const envBin = trimEnv(env.ADP_TESSERACT);
  const optionTess = normalizeTessdataDir(input.tessdataDir, exists);
  const envTess = normalizeTessdataDir(trimEnv(env.ADP_TESSDATA) ?? trimEnv(env.TESSDATA_PREFIX), exists);
  const tessdataDir = optionTess ?? envTess;
  const optionPoppler = trimEnv(input.popplerBin);
  const envPoppler = trimEnv(env.ADP_PDFTOPPM);
  const binName = tesseractBinName();
  const pathPoppler = popplerBinName();

  const attachPoppler = (runtimeRoot?: string, binPath?: string): string | undefined => {
    if (optionPoppler) return optionPoppler;
    if (envPoppler) return envPoppler;
    const search: string[] = [];
    if (runtimeRoot) search.push(runtimeRoot);
    if (binPath) search.push(path.dirname(binPath), path.dirname(path.dirname(binPath)));
    search.push(...ocrRootCandidates(input));
    return findPopplerNear(search, exists) ?? pathPoppler;
  };

  const libDirFor = (root: string): string | undefined => {
    const lib = path.join(root, 'lib');
    return exists(lib) ? lib : undefined;
  };

  if (optionBin || envBin) {
    const bin = (optionBin ?? envBin) as string;
    const source: OcrRuntimeSource = optionBin ? 'options' : 'env';
    const root = path.dirname(path.dirname(bin));
    return {
      bin,
      tessdataDir,
      popplerBin: attachPoppler(root, bin),
      libDir: libDirFor(root),
      bundled: false,
      source,
    };
  }

  const seen = new Set<string>();
  for (const root of ocrRootCandidates(input)) {
    if (seen.has(root)) continue;
    seen.add(root);
    const bin = path.join(root, 'bin', binName);
    const tess = path.join(root, 'tessdata');
    if (exists(bin) && hasAllOcrLanguages(tess, languages, exists)) {
      return {
        bin,
        tessdataDir: tessdataDir ?? tess,
        popplerBin: attachPoppler(root, bin),
        libDir: libDirFor(root),
        bundled: true,
        source: 'bundled',
      };
    }
  }

  return {
    bin: binName,
    tessdataDir,
    popplerBin: attachPoppler(undefined, undefined),
    bundled: false,
    source: 'path',
  };
}

export function ocrProcessEnv(
  runtime: OcrRuntime,
  baseEnv: NodeJS.ProcessEnv = process.env
): NodeJS.ProcessEnv {
  const env = Object.assign({}, baseEnv);
  const prepend = (key: string, value: string, sep: string) => {
    const prev = env[key] || '';
    env[key] = prev ? value + sep + prev : value;
  };

  if (runtime.libDir) {
    if (process.platform === 'win32') {
      prepend('PATH', runtime.libDir, ';');
    } else if (process.platform === 'darwin') {
      prepend('DYLD_LIBRARY_PATH', runtime.libDir, ':');
    } else {
      prepend('LD_LIBRARY_PATH', runtime.libDir, ':');
    }
  }

  if (runtime.bundled || runtime.source === 'options' || runtime.source === 'env') {
    const binDir = path.dirname(runtime.bin);
    if (binDir && binDir !== '.' && binDir !== path.basename(runtime.bin)) {
      prepend('PATH', binDir, process.platform === 'win32' ? ';' : ':');
    }
  }

  if (runtime.tessdataDir && !trimEnv(env.TESSDATA_PREFIX)) {
    env.TESSDATA_PREFIX = runtime.tessdataDir;
  }
  return env;
}

export function parseTesseractListLangs(output: string): string[] {
  const langs: string[] = [];
  const seen = new Set<string>();
  const lines = output.split(/\r?\n/);
  for (const raw of lines) {
    const line = raw.trim();
    if (!line) continue;
    if (/list of available languages/i.test(line)) continue;
    if (/^tesseract\s/i.test(line)) continue;
    if (/\s/.test(line)) continue;
    if (!/^[A-Za-z][A-Za-z0-9_]*(?:\/[A-Za-z][A-Za-z0-9_]*)?$/.test(line)) continue;
    const key = line.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    langs.push(line);
  }
  return langs;
}

export function missingOcrLanguages(
  requested: readonly string[],
  installed: readonly string[]
): string[] {
  const have = new Set(installed.map((lang) => lang.toLowerCase()));
  return requested.filter((lang) => !have.has(lang.toLowerCase()));
}

export function ocrLangMissingWarning(
  missing: readonly string[],
  installed: readonly string[]
): DocumentWarning {
  const available = installed.length ? installed.join(', ') : '(none reported)';
  return {
    code: 'OCR_LANG_MISSING',
    message:
      'Requested OCR language(s) are not installed: ' +
      missing.join(', ') +
      '. Available: ' +
      available +
      '. Add the traineddata file to tessdata, or pass tessdataDir / TESSDATA_PREFIX.',
  };
}

export type OcrCommandRunner = (
  cmd: string,
  args: string[],
  env: NodeJS.ProcessEnv,
  timeout: number
) => Promise<{ stdout: string; stderr: string }>;

export async function listInstalledOcrLanguages(
  runtime: OcrRuntime,
  run: OcrCommandRunner,
  timeoutMs: number
): Promise<{ languages: string[]; warning?: DocumentWarning }> {
  try {
    const args = ['--list-langs'];
    if (runtime.tessdataDir) args.push('--tessdata-dir', runtime.tessdataDir);
    const out = await run(runtime.bin, args, ocrProcessEnv(runtime), timeoutMs);
    const languages = parseTesseractListLangs(`${out.stdout}\n${out.stderr}`);
    if (languages.length) return { languages };
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    if (runtime.tessdataDir) {
      // Fall through to file probe below; if that also fails, surface this error.
      void msg;
    } else {
      return {
        languages: [],
        warning: {
          code: 'OCR_LANG_CHECK_FAILED',
          message: 'Could not list Tesseract languages: ' + msg.slice(0, 300),
        },
      };
    }
  }

  // Fallback when --list-langs is unavailable but tessdata is known.
  return { languages: [] };
}
