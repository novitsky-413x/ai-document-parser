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
 * There is no official portable zip for every OS, so --bin download is opt-in
 * and requires --bin-url (and optional --bin-sha256). Use --dry-run first and
 * check printed sizes before downloading anything large.
 *
 * Scanned-PDF OCR also needs pdftoppm (Poppler). This script does not bundle
 * Poppler; image OCR works with Tesseract alone. Install poppler-utils (Linux)
 * or a Poppler build (Windows) if you need PDF page rasterization.
 */
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, copyFileSync, writeFileSync, chmodSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

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
  --bin copy|none|download
                       copy: copy tesseract from PATH into dest/bin (default)
                       none: only fetch tessdata
                       download: fetch --bin-url into dest/bin
  --bin-url <url>      Archive or binary to place as dest/bin/tesseract(.exe)
  --bin-sha256 <hex>   Optional SHA-256 check for --bin-url
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
    bin: 'copy',
    binUrl: undefined,
    binSha256: undefined,
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
    else if (a === '--dest') out.dest = path.resolve(argv[++i] ?? '');
    else if (a === '--langs') {
      out.langs = String(argv[++i] ?? '')
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean);
    } else if (a === '--set') out.set = String(argv[++i] ?? 'fast');
    else if (a === '--lang-url') out.langUrl = argv[++i];
    else if (a === '--bin') out.bin = String(argv[++i] ?? 'copy');
    else if (a === '--bin-url') out.binUrl = argv[++i];
    else if (a === '--bin-sha256') out.binSha256 = String(argv[++i] ?? '').toLowerCase();
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

function formatBytes(n) {
  if (n == null || !Number.isFinite(n)) return 'unknown size';
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KiB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MiB`;
}

export function plan(opts) {
  const dest = opts.dest;
  const binName = process.platform === 'win32' ? 'tesseract.exe' : 'tesseract';
  const langs = opts.langs.map((lang) => {
    const url = tessdataUrl(lang, opts.set ?? 'fast', opts.langUrl);
    return { lang, url, dest: path.join(dest, 'tessdata', `${lang}.traineddata`) };
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
  };
}

async function download(url, dest, sha256, force) {
  if (existsSync(dest) && !force) {
    console.log(`skip existing ${dest} (use --force to overwrite)`);
    return;
  }
  const res = await fetch(url, { redirect: 'follow' });
  if (!res.ok) throw new Error(`GET ${url} failed: ${res.status}`);
  const buf = Buffer.from(await res.arrayBuffer());
  if (sha256) {
    const actual = createHash('sha256').update(buf).digest('hex');
    if (actual !== sha256) throw new Error(`SHA-256 mismatch for ${url}: ${actual} != ${sha256}`);
  }
  mkdirSync(path.dirname(dest), { recursive: true });
  writeFileSync(dest, buf);
  console.log(`wrote ${dest} (${formatBytes(buf.length)})`);
}

async function main(argv = process.argv.slice(2)) {
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
    console.log(`    -> ${lang.dest}`);
  }
  if (p.binMode === 'copy') {
    console.log(`  bin copy: ${p.systemBin ?? '(tesseract not found on PATH — install it or use --bin none / --bin download --bin-url …)'}`);
    console.log(`    -> ${p.binDest}`);
  } else if (p.binMode === 'download') {
    console.log(`  bin download: ${p.binUrl}`);
    console.log(`    -> ${p.binDest}`);
  } else {
    console.log('  bin: skipped (--bin none)');
  }
  console.log('  poppler/pdftoppm: not bundled (needed only for scanned-PDF OCR)');

  if (opts.probeSizes) {
    const probes = [...p.langs.map((l) => l.url)];
    if (p.binUrl) probes.push(p.binUrl);
    let total = 0;
    for (const url of probes) {
      const info = await headSize(url);
      const size = formatBytes(info.bytes);
      if (info.bytes) total += info.bytes;
      console.log(`  size ${url}: ${info.ok ? size : `HTTP ${info.status}`}`);
    }
    console.log(`  tessdata(+bin url) total: ${formatBytes(total || null)}`);
  }

  if (opts.dryRun) {
    console.log('dry-run: nothing written');
    return 0;
  }

  mkdirSync(path.join(p.dest, 'tessdata'), { recursive: true });
  mkdirSync(path.join(p.dest, 'bin'), { recursive: true });
  for (const lang of p.langs) {
    await download(lang.url, lang.dest, undefined, opts.force);
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
        try {
          chmodSync(p.binDest, 0o755);
        } catch {
          /* ignore */
        }
      }
      console.log(`copied ${p.systemBin} -> ${p.binDest}`);
    }
  } else if (p.binMode === 'download' && p.binUrl) {
    await download(p.binUrl, p.binDest, opts.binSha256, opts.force);
    if (process.platform !== 'win32') {
      try {
        chmodSync(p.binDest, 0o755);
      } catch {
        /* ignore */
      }
    }
  }
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

export { HELP, DEFAULT_DEST, main };
