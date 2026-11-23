#!/usr/bin/env node
/**
 * Fetch a Tesseract runtime into resources/ocr (binaries are gitignored).
 *
 * Layout expected by src/parse/pdf-ocr.ts:
 *   <dest>/bin/tesseract(.exe)
 *   <dest>/tessdata/<lang>.traineddata
 *   <dest>/lib/   (optional; LD_LIBRARY_PATH / DYLD_LIBRARY_PATH / PATH)
 *
 * This script is not run by npm test. Defaults:
 *   --langs rus,eng   official tessdata_fast (any ISO / Tesseract lang id)
 *   --set fast        tessdata_fast | tessdata_best | tessdata (standard)
 *   --bin copy        copy a system tesseract into dest/bin when one is on PATH
 *
 * --bin download accepts a raw binary or an archive (zip, tar.*, 7z, NSIS
 * installer, 7-Zip SFX). Archives are extracted and the tesseract binary plus
 * adjacent libraries are copied into dest/bin. NSIS / 7z extraction needs a
 * full 7-Zip (not 7zr); pass --seven-zip-* pins or have `7z` on PATH.
 *
 * There is no official portable zip for every OS, so --bin download is opt-in
 * and requires --bin-url (and --bin-sha256 in production). Use --dry-run first
 * and check printed sizes before downloading anything large.
 *
 * Scanned-PDF OCR also needs pdftoppm (Poppler). Poppler is GPL, so this
 * script does not download it unless you pass --poppler-url explicitly.
 * Image OCR works with Tesseract alone.
 */
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, copyFileSync, chmodSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  assertSha256,
  copyRuntimeBinaries,
  detectArchiveKind,
  downloadVerified,
  ensureSevenZip,
  extractArchive,
  formatBytes,
  parseLangSha256,
  sha256File,
} from './ocr-archive.mjs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const DEFAULT_DEST = path.join(ROOT, 'resources', 'ocr');

export const TESSDATA_SETS = {
  fast: 'https://github.com/tesseract-ocr/tessdata_fast/raw/refs/heads/main/{lang}.traineddata',
  best: 'https://github.com/tesseract-ocr/tessdata_best/raw/refs/heads/main/{lang}.traineddata',
  standard: 'https://github.com/tesseract-ocr/tessdata/raw/refs/heads/main/{lang}.traineddata',
};

const LANG_ID = /^[A-Za-z][A-Za-z0-9_]*(?:\/[A-Za-z][A-Za-z0-9_]*)?$/;

export function assertLangId(lang) {
  if (!LANG_ID.test(lang) || lang.includes('..')) {
    throw new Error(`Invalid OCR language id '${lang}'`);
  }
}

export function tessdataUrl(lang, set = 'fast', template) {
  assertLangId(lang);
  const tmpl = template || TESSDATA_SETS[set];
  if (!tmpl) {
    throw new Error(`Unknown tessdata set '${set}' (use fast, best, or standard)`);
  }
  return tmpl.replaceAll('{lang}', lang);
}

/** @deprecated Prefer tessdataUrl(lang, 'fast'). Kept for existing callers/tests. */
export const TESSDATA_FAST = {
  get eng() {
    return tessdataUrl('eng', 'fast');
  },
  get rus() {
    return tessdataUrl('rus', 'fast');
  },
};

const HELP = `Usage: node scripts/fetch-ocr-runtime.mjs [options]

  --dest <dir>         Target directory (default: ${DEFAULT_DEST})
  --langs <list>       Comma-separated tessdata languages (default: rus,eng)
  --set fast|best|standard
                       Official tessdata variant (default: fast)
  --lang-url <tmpl>    URL template with {lang} (overrides --set)
  --lang-sha256 <map>  Comma-separated lang=hex (or lang:hex) SHA-256 checks
  --bin copy|none|download
                       copy: copy tesseract from PATH into dest/bin (default)
                       none: only fetch tessdata
                       download: fetch --bin-url (raw binary or archive)
  --bin-url <url>      Archive (zip/7z/tar/NSIS/SFX) or raw tesseract binary
  --bin-sha256 <hex>   SHA-256 check for --bin-url
  --bin-raw            Treat --bin-url as a raw binary even if it looks like an archive
  --cache-dir <dir>    Cache for installer / 7-Zip bootstrap downloads
  --seven-zip-url <url>       7zr.exe (Windows) bootstrap
  --seven-zip-sha256 <hex>
  --seven-zip-sfx-url <url>   Full 7-Zip SFX (Windows, extracted with 7zr)
  --seven-zip-sfx-sha256 <hex>
  --seven-zip-tarball-url <url>   7-Zip .tar.xz (Linux)
  --seven-zip-tarball-sha256 <hex>
  --poppler-url <url>  Optional Poppler/pdftoppm archive (GPL — opt-in only)
  --poppler-sha256 <hex>
  --dry-run            Print the plan (and optional HEAD sizes with --probe-sizes)
  --probe-sizes        HEAD tessdata (and --bin-url) to report Content-Length
  --force              Overwrite existing files
  --help               Show this help
`;

export function parseArgs(argv) {
  const out = {
    dest: DEFAULT_DEST,
    langs: ['rus', 'eng'],
    set: 'fast',
    langUrl: undefined,
    langSha256: {},
    bin: 'copy',
    binUrl: undefined,
    binSha256: undefined,
    binRaw: false,
    cacheDir: undefined,
    sevenZipUrl: undefined,
    sevenZipSha256: undefined,
    sevenZipSfxUrl: undefined,
    sevenZipSfxSha256: undefined,
    sevenZipTarballUrl: undefined,
    sevenZipTarballSha256: undefined,
    popplerUrl: undefined,
    popplerSha256: undefined,
    dryRun: false,
    probeSizes: false,
    force: false,
    help: false,
  };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--help' || a === '-h') out.help = true;
    else if (a === '--dry-run') out.dryRun = true;
    else if (a === '--probe-sizes') out.probeSizes = true;
    else if (a === '--force') out.force = true;
    else if (a === '--bin-raw') out.binRaw = true;
    else if (a === '--dest') out.dest = path.resolve(argv[++i] ?? '');
    else if (a === '--langs') {
      out.langs = String(argv[++i] ?? '')
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean);
    } else if (a === '--set') out.set = String(argv[++i] ?? 'fast');
    else if (a === '--lang-url') out.langUrl = argv[++i];
    else if (a === '--lang-sha256') out.langSha256 = parseLangSha256(argv[++i]);
    else if (a === '--bin') out.bin = String(argv[++i] ?? 'copy');
    else if (a === '--bin-url') out.binUrl = argv[++i];
    else if (a === '--bin-sha256') out.binSha256 = String(argv[++i] ?? '').toLowerCase();
    else if (a === '--cache-dir') out.cacheDir = path.resolve(argv[++i] ?? '');
    else if (a === '--seven-zip-url') out.sevenZipUrl = argv[++i];
    else if (a === '--seven-zip-sha256') out.sevenZipSha256 = String(argv[++i] ?? '').toLowerCase();
    else if (a === '--seven-zip-sfx-url') out.sevenZipSfxUrl = argv[++i];
    else if (a === '--seven-zip-sfx-sha256') out.sevenZipSfxSha256 = String(argv[++i] ?? '').toLowerCase();
    else if (a === '--seven-zip-tarball-url') out.sevenZipTarballUrl = argv[++i];
    else if (a === '--seven-zip-tarball-sha256') out.sevenZipTarballSha256 = String(argv[++i] ?? '').toLowerCase();
    else if (a === '--poppler-url') out.popplerUrl = argv[++i];
    else if (a === '--poppler-sha256') out.popplerSha256 = String(argv[++i] ?? '').toLowerCase();
    else throw new Error(`Unknown argument: ${a}`);
  }
  if (!['fast', 'best', 'standard'].includes(out.set)) {
    throw new Error(`--set must be fast, best, or standard (got ${out.set})`);
  }
  if (!['copy', 'none', 'download'].includes(out.bin)) {
    throw new Error(`--bin must be copy, none, or download (got ${out.bin})`);
  }
  if (out.bin === 'download' && !out.binUrl) {
    throw new Error('--bin download requires --bin-url (no single official portable zip for every OS)');
  }
  return out;
}

function whichTesseract() {
  const bin = process.platform === 'win32' ? 'tesseract.exe' : 'tesseract';
  const r = spawnSync(process.platform === 'win32' ? 'where' : 'which', [bin], {
    encoding: 'utf8',
    windowsHide: true,
  });
  const line = (r.stdout || '')
    .split(/\r?\n/)
    .map((s) => s.trim())
    .find((s) => s && existsSync(s));
  return line;
}

async function headSize(url) {
  const res = await fetch(url, { method: 'HEAD', redirect: 'follow' });
  if (!res.ok) return { url, ok: false, status: res.status, bytes: null };
  const len = res.headers.get('content-length');
  return { url, ok: true, status: res.status, bytes: len ? Number(len) : null };
}

export function plan(opts) {
  const dest = opts.dest;
  const binName = process.platform === 'win32' ? 'tesseract.exe' : 'tesseract';
  const langs = opts.langs.map((lang) => {
    const url = tessdataUrl(lang, opts.set ?? 'fast', opts.langUrl);
    return {
      lang,
      url,
      dest: path.join(dest, 'tessdata', `${lang}.traineddata`),
      sha256: opts.langSha256?.[lang],
    };
  });
  const systemBin = opts.bin === 'copy' ? whichTesseract() : undefined;
  return {
    dest,
    langs,
    binMode: opts.bin,
    binName,
    binDest: path.join(dest, 'bin', binName),
    systemBin: systemBin ?? null,
    binUrl: opts.binUrl ?? null,
    binRaw: Boolean(opts.binRaw),
    cacheDir: opts.cacheDir || path.join(os.tmpdir(), 'adp-ocr-fetch-cache'),
    popplerUrl: opts.popplerUrl ?? null,
  };
}

function sevenZipPins(opts) {
  const pins = {};
  if (opts.sevenZipUrl) pins.sevenZr = { url: opts.sevenZipUrl, sha256: opts.sevenZipSha256 };
  if (opts.sevenZipSfxUrl) pins.sfx = { url: opts.sevenZipSfxUrl, sha256: opts.sevenZipSfxSha256 };
  if (opts.sevenZipTarballUrl) pins.tarball = { url: opts.sevenZipTarballUrl, sha256: opts.sevenZipTarballSha256 };
  return Object.keys(pins).length ? pins : undefined;
}

async function installFromDownload(opts, p) {
  const cacheDir = p.cacheDir;
  const filename = path.basename(new URL(p.binUrl).pathname) || (process.platform === 'win32' ? 'tesseract-download.exe' : 'tesseract-download');
  const archivePath = path.join(cacheDir, filename);
  const result = await downloadVerified(p.binUrl, archivePath, { sha256: opts.binSha256, force: opts.force });
  console.log(`${result.skipped ? 'reused' : 'wrote'} ${archivePath} (${formatBytes(result.bytes)})`);

  if (opts.binRaw) {
    mkdirSync(path.dirname(p.binDest), { recursive: true });
    copyFileSync(archivePath, p.binDest);
    if (process.platform !== 'win32') {
      try { chmodSync(p.binDest, 0o755); } catch { /* ignore */ }
    }
    console.log(`copied raw binary -> ${p.binDest}`);
    return;
  }

  const kind = detectArchiveKind(archivePath);
  if (kind === 'raw') {
    mkdirSync(path.dirname(p.binDest), { recursive: true });
    copyFileSync(archivePath, p.binDest);
    if (process.platform !== 'win32') {
      try { chmodSync(p.binDest, 0o755); } catch { /* ignore */ }
    }
    console.log(`copied raw binary -> ${p.binDest}`);
    return;
  }

  let sevenZipBin;
  if (kind === '7z' || kind === 'exe' || kind === 'nsis') {
    sevenZipBin = await ensureSevenZip({ cacheDir, sevenZip: sevenZipPins(opts) });
    if (kind === 'exe') {
      const { sevenZipCanOpen } = await import('./ocr-archive.mjs');
      if (!sevenZipCanOpen(sevenZipBin, archivePath)) {
        mkdirSync(path.dirname(p.binDest), { recursive: true });
        copyFileSync(archivePath, p.binDest);
        console.log(`copied raw executable -> ${p.binDest}`);
        return;
      }
    }
  }

  const extractDir = path.join(cacheDir, 'extract', filename.replace(/[^\w.-]+/g, '_'));
  console.log(`extracting ${kind} archive with ${sevenZipBin ? '7-Zip' : kind}…`);
  await extractArchive(archivePath, extractDir, { sevenZipBin, kind });
  const installed = copyRuntimeBinaries(extractDir, p.dest);
  console.log(`installed ${installed}`);
}

async function installPoppler(opts, p) {
  if (!opts.popplerUrl) return;
  console.warn(
    'WARNING: Poppler is licensed under the GPL. Bundling it has copyleft obligations. Image OCR works without it.'
  );
  const cacheDir = p.cacheDir;
  const filename = path.basename(new URL(opts.popplerUrl).pathname) || 'poppler-download';
  const archivePath = path.join(cacheDir, filename);
  const result = await downloadVerified(opts.popplerUrl, archivePath, { sha256: opts.popplerSha256, force: opts.force });
  console.log(`${result.skipped ? 'reused' : 'wrote'} ${archivePath} (${formatBytes(result.bytes)})`);
  const extractDir = path.join(cacheDir, 'extract', filename.replace(/[^\w.-]+/g, '_'));
  const kind = detectArchiveKind(archivePath);
  let sevenZipBin;
  if (kind === '7z' || kind === 'exe' || kind === 'nsis') {
    sevenZipBin = await ensureSevenZip({ cacheDir, sevenZip: sevenZipPins(opts) });
  }
  await extractArchive(archivePath, extractDir, { sevenZipBin, kind });
  const destBin = path.join(p.dest, 'bin');
  mkdirSync(destBin, { recursive: true });
  const { walkFiles } = await import('./ocr-archive.mjs');
  const names = process.platform === 'win32' ? ['pdftoppm.exe'] : ['pdftoppm'];
  const found = walkFiles(extractDir).find((f) => names.includes(path.basename(f).toLowerCase()) || names.includes(path.basename(f)));
  const match = walkFiles(extractDir).find((f) => names.includes(path.basename(f).toLowerCase()));
  if (!match) throw new Error(`pdftoppm not found in Poppler archive ${archivePath}`);
  const destName = process.platform === 'win32' ? 'pdftoppm.exe' : 'pdftoppm';
  copyFileSync(match, path.join(destBin, destName));
  const srcDir = path.dirname(match);
  const { readdirSync } = await import('node:fs');
  for (const name of readdirSync(srcDir)) {
    if (path.extname(name).toLowerCase() === '.dll') {
      copyFileSync(path.join(srcDir, name), path.join(destBin, name));
    }
  }
  console.log(`installed ${path.join(destBin, destName)}`);
}

export async function main(argv = process.argv.slice(2)) {
  const opts = parseArgs(argv);
  if (opts.help) {
    process.stdout.write(HELP);
    return 0;
  }
  const p = plan(opts);
  console.log(`OCR runtime dest: ${p.dest}`);
  console.log(`  tessdata set: ${opts.set}${opts.langUrl ? ` (template ${opts.langUrl})` : ''}`);
  for (const lang of p.langs) {
    console.log(`  tessdata ${lang.lang}: ${lang.url}`);
    console.log(`    -> ${lang.dest}${lang.sha256 ? ` (sha256 ${lang.sha256.slice(0, 12)}…)` : ''}`);
  }
  if (p.binMode === 'copy') {
    console.log(`  bin copy: ${p.systemBin ?? '(tesseract not found on PATH — install it or use --bin none / --bin download --bin-url …)'}`);
    console.log(`    -> ${p.binDest}`);
  } else if (p.binMode === 'download') {
    console.log(`  bin download: ${p.binUrl}${p.binRaw ? ' (raw)' : ' (archive or raw, auto-detected)'}`);
    console.log(`    -> ${p.binDest}`);
  } else {
    console.log('  bin: skipped (--bin none)');
  }
  if (p.popplerUrl) console.log(`  poppler (GPL, opt-in): ${p.popplerUrl}`);
  else console.log('  poppler/pdftoppm: not bundled (needed only for scanned-PDF OCR; GPL)');

  if (opts.probeSizes) {
    const probes = [...p.langs.map((l) => l.url)];
    if (p.binUrl) probes.push(p.binUrl);
    if (p.popplerUrl) probes.push(p.popplerUrl);
    let total = 0;
    for (const url of probes) {
      const info = await headSize(url);
      const size = formatBytes(info.bytes);
      if (info.bytes) total += info.bytes;
      console.log(`  size ${url}: ${info.ok ? size : `HTTP ${info.status}`}`);
    }
    console.log(`  tessdata(+bin/poppler url) total: ${formatBytes(total || null)}`);
  }

  if (opts.dryRun) {
    console.log('dry-run: nothing written');
    return 0;
  }

  mkdirSync(path.join(p.dest, 'tessdata'), { recursive: true });
  mkdirSync(path.join(p.dest, 'bin'), { recursive: true });
  for (const lang of p.langs) {
    const result = await downloadVerified(lang.url, lang.dest, { sha256: lang.sha256, force: opts.force });
    console.log(`${result.skipped ? 'skip existing' : 'wrote'} ${lang.dest} (${formatBytes(result.bytes)})`);
  }
  if (p.binMode === 'copy') {
    if (!p.systemBin) {
      console.error('No tesseract on PATH; tessdata was fetched. Install Tesseract and re-run with --bin copy, or pass --bin download --bin-url <portable archive>.');
      return 2;
    }
    if (existsSync(p.binDest) && !opts.force) {
      console.log(`skip existing ${p.binDest}`);
    } else {
      copyFileSync(p.systemBin, p.binDest);
      if (process.platform !== 'win32') {
        try { chmodSync(p.binDest, 0o755); } catch { /* ignore */ }
      }
      console.log(`copied ${p.systemBin} -> ${p.binDest}`);
    }
  } else if (p.binMode === 'download' && p.binUrl) {
    await installFromDownload(opts, p);
  }
  if (opts.popplerUrl) await installPoppler(opts, p);
  return 0;
}

const isMain = process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isMain) {
  main().then(
    (code) => process.exit(code),
    (err) => {
      console.error(err instanceof Error ? err.message : err);
      process.exit(1);
    }
  );
}

export { HELP, DEFAULT_DEST, sha256File, assertSha256 };

