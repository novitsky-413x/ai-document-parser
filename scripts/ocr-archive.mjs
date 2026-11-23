/**
 * Shared helpers for fetching and extracting an OCR runtime.
 * Used by ai-document-parser's fetch-ocr-runtime.mjs and by WATCHER's
 * wrapper (as a fallback when the installed parser is older than 0.3.111).
 *
 * Keep the two copies of this file in sync.
 */
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  existsSync,
  mkdirSync,
  readdirSync,
  statSync,
  copyFileSync,
  chmodSync,
  readFileSync,
  writeFileSync,
  rmSync,
} from 'node:fs';
import { open } from 'node:fs/promises';
import path from 'node:path';
import { inflateRawSync } from 'node:zlib';

export function formatBytes(n) {
  if (n == null || !Number.isFinite(n)) return 'unknown size';
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KiB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MiB`;
}

export function sha256Buffer(buf) {
  return createHash('sha256').update(buf).digest('hex');
}

export async function sha256File(file) {
  const hash = createHash('sha256');
  const fh = await open(file, 'r');
  try {
    const stream = fh.createReadStream();
    for await (const chunk of stream) hash.update(chunk);
  } finally {
    await fh.close();
  }
  return hash.digest('hex');
}

export function assertSha256(actual, expected, label) {
  const want = String(expected || '').trim().toLowerCase();
  const got = String(actual || '').trim().toLowerCase();
  if (!want) throw new Error(`SHA-256 missing for ${label}`);
  if (!/^[0-9a-f]{64}$/.test(want)) throw new Error(`Invalid SHA-256 for ${label}: ${expected}`);
  if (got !== want) throw new Error(`SHA-256 mismatch for ${label}: ${got} != ${want}`);
}

export async function downloadVerified(url, dest, { sha256, force = false } = {}) {
  if (existsSync(dest) && !force) {
    if (sha256) assertSha256(await sha256File(dest), sha256, url);
    return { dest, bytes: statSync(dest).size, skipped: true, sha256: sha256 ? String(sha256).toLowerCase() : await sha256File(dest) };
  }
  const res = await fetch(url, { redirect: 'follow' });
  if (!res.ok) throw new Error(`GET ${url} failed: ${res.status}`);
  const buf = Buffer.from(await res.arrayBuffer());
  const actual = sha256Buffer(buf);
  if (sha256) assertSha256(actual, sha256, url);
  mkdirSync(path.dirname(dest), { recursive: true });
  writeFileSync(dest, buf);
  return { dest, bytes: buf.length, skipped: false, sha256: actual };
}

export function readMagic(filePath, n = 8) {
  const fh = readFileSync(filePath);
  return fh.subarray(0, Math.min(n, fh.length));
}

export function detectArchiveKind(filePath, header) {
  const name = String(filePath || '').toLowerCase();
  const h = header || (existsSync(filePath) ? readMagic(filePath, 8) : Buffer.alloc(0));
  if (name.endsWith('.zip') || (h.length >= 2 && h[0] === 0x50 && h[1] === 0x4b)) return 'zip';
  if (name.endsWith('.7z') || (h.length >= 4 && h[0] === 0x37 && h[1] === 0x7a && h[2] === 0xbc && h[3] === 0xaf)) return '7z';
  if (/\.(tar\.xz|txz|tar\.gz|tgz|tar\.bz2|tbz2|tar)$/.test(name)) return 'tar';
  if (h.length >= 2 && h[0] === 0x4d && h[1] === 0x5a) return 'exe';
  if (name.endsWith('.exe')) return 'exe';
  return 'raw';
}

function safeJoin(outDir, entryName) {
  const normalized = String(entryName).replace(/\\/g, '/');
  if (path.isAbsolute(normalized) || normalized.split('/').some((p) => p === '..')) {
    throw new Error(`Refusing unsafe archive path: ${entryName}`);
  }
  const dest = path.resolve(outDir, normalized);
  const root = path.resolve(outDir) + path.sep;
  if (dest !== path.resolve(outDir) && !dest.startsWith(root)) {
    throw new Error(`Refusing unsafe archive path: ${entryName}`);
  }
  return dest;
}

function findEocd(buf) {
  const sig = 0x06054b50;
  const min = Math.max(0, buf.length - 22 - 65535);
  for (let i = buf.length - 22; i >= min; i--) {
    if (buf.readUInt32LE(i) === sig) return i;
  }
  throw new Error('ZIP end-of-central-directory not found');
}

/** Pure-JS unzip for store (0) and deflate (8). Rejects Zip64 and path traversal. */
export function extractZipBuffer(buf, outDir, { maxOutputBytes = 2 * 1024 * 1024 * 1024 } = {}) {
  if (!Buffer.isBuffer(buf)) buf = Buffer.from(buf);
  const eocd = findEocd(buf);
  const disk = buf.readUInt16LE(eocd + 4);
  const cdCount = buf.readUInt16LE(eocd + 10);
  const cdSize = buf.readUInt32LE(eocd + 12);
  const cdOff = buf.readUInt32LE(eocd + 16);
  if (disk !== 0 || cdCount === 0xffff || cdSize === 0xffffffff || cdOff === 0xffffffff) {
    throw new Error('Zip64 archives are not supported by the built-in extractor');
  }
  mkdirSync(outDir, { recursive: true });
  let offset = cdOff;
  let written = 0;
  const files = [];
  for (let i = 0; i < cdCount; i++) {
    if (buf.readUInt32LE(offset) !== 0x02014b50) throw new Error('Invalid ZIP central-directory signature');
    const method = buf.readUInt16LE(offset + 10);
    const compSize = buf.readUInt32LE(offset + 20);
    const uncompSize = buf.readUInt32LE(offset + 24);
    const nameLen = buf.readUInt16LE(offset + 28);
    const extraLen = buf.readUInt16LE(offset + 30);
    const commentLen = buf.readUInt16LE(offset + 32);
    const localOff = buf.readUInt32LE(offset + 42);
    const name = buf.subarray(offset + 46, offset + 46 + nameLen).toString('utf8');
    offset += 46 + nameLen + extraLen + commentLen;
    if (name.endsWith('/')) {
      mkdirSync(safeJoin(outDir, name), { recursive: true });
      continue;
    }
    if (buf.readUInt32LE(localOff) !== 0x04034b50) throw new Error(`Invalid ZIP local header for ${name}`);
    const localNameLen = buf.readUInt16LE(localOff + 26);
    const localExtraLen = buf.readUInt16LE(localOff + 28);
    const dataStart = localOff + 30 + localNameLen + localExtraLen;
    const compact = buf.subarray(dataStart, dataStart + compSize);
    let data;
    if (method === 0) data = compact;
    else if (method === 8) data = inflateRawSync(compact);
    else throw new Error(`Unsupported ZIP compression method ${method} for ${name}`);
    if (data.length !== uncompSize) throw new Error(`ZIP size mismatch for ${name}`);
    written += data.length;
    if (written > maxOutputBytes) throw new Error('ZIP output exceeds size limit');
    const dest = safeJoin(outDir, name);
    mkdirSync(path.dirname(dest), { recursive: true });
    writeFileSync(dest, data);
    files.push(name);
  }
  return { files, bytes: written };
}

export function extractZipFile(zipPath, outDir, opts) {
  return extractZipBuffer(readFileSync(zipPath), outDir, opts);
}

export function extractZip(zipPath, outDir, opts) {
  mkdirSync(outDir, { recursive: true });
  const tar = spawnSync('tar', ['-xf', zipPath, '-C', outDir], { encoding: 'utf8', windowsHide: true });
  if (tar.status === 0) return { method: 'tar' };
  const js = extractZipFile(zipPath, outDir, opts);
  return { method: 'js', ...js };
}

function whichAll(names) {
  const cmd = process.platform === 'win32' ? 'where' : 'which';
  for (const name of names) {
    const r = spawnSync(cmd, [name], { encoding: 'utf8', windowsHide: true });
    const line = (r.stdout || '')
      .split(/\r?\n/)
      .map((s) => s.trim())
      .find((s) => s && existsSync(s));
    if (line) return line;
  }
  return null;
}

export function findSevenZipOnPath() {
  return whichAll(process.platform === 'win32' ? ['7z.exe', '7zz.exe'] : ['7z', '7zz']);
}

export function runSevenZip(bin, args, opts = {}) {
  const cwd = opts.cwd || path.dirname(bin);
  const r = spawnSync(bin, args, {
    encoding: 'utf8',
    windowsHide: true,
    cwd,
    maxBuffer: opts.maxBuffer ?? 32 * 1024 * 1024,
  });
  if (r.status !== 0) {
    throw new Error(`7-Zip failed (${r.status}): ${(r.stderr || r.stdout || '').trim() || 'no output'}\n${bin} ${args.join(' ')}`);
  }
  return r;
}

export function sevenZipCanOpen(bin, archivePath) {
  const r = spawnSync(bin, ['l', archivePath], {
    encoding: 'utf8',
    windowsHide: true,
    cwd: path.dirname(bin),
    maxBuffer: 16 * 1024 * 1024,
  });
  const out = `${r.stdout || ''}\n${r.stderr || ''}`;
  if (r.status !== 0) return false;
  if (/Cannot open the file as archive/i.test(out)) return false;
  return /Type = |Listing archive:/i.test(out);
}

/**
 * Bootstraps a full 7-Zip (with NSIS support) into cacheDir when needed.
 * 7zr (reduced) cannot open NSIS installers, so Windows uses 7zr to extract
 * the official 7-Zip SFX, then runs 7z.exe from that tree.
 * sevenZip: { sevenZr?: {url, sha256}, sfx?: {url, sha256}, tarball?: {url, sha256} }
 */
export async function ensureSevenZip({ cacheDir, sevenZip } = {}) {
  const onPath = findSevenZipOnPath();
  if (onPath) return onPath;
  if (!cacheDir) throw new Error('No 7-Zip on PATH and no cacheDir to bootstrap one');
  mkdirSync(cacheDir, { recursive: true });

  if (process.platform === 'linux' && sevenZip?.tarball?.url) {
    const tarPath = path.join(cacheDir, path.basename(new URL(sevenZip.tarball.url).pathname) || '7z-linux.tar.xz');
    await downloadVerified(sevenZip.tarball.url, tarPath, { sha256: sevenZip.tarball.sha256 });
    const out = path.join(cacheDir, '7z-linux');
    mkdirSync(out, { recursive: true });
    const r = spawnSync('tar', ['-xf', tarPath, '-C', out], { encoding: 'utf8' });
    if (r.status !== 0) throw new Error(`Failed to extract 7-Zip tarball: ${(r.stderr || r.stdout || '').trim()}`);
    const found = findFileByNames(out, ['7zz', '7z']);
    if (!found) throw new Error('7zz not found in extracted 7-Zip tarball');
    try { chmodSync(found, 0o755); } catch { /* ignore */ }
    return found;
  }

  if (process.platform === 'win32' && sevenZip?.sevenZr?.url && sevenZip?.sfx?.url) {
    const zrPath = path.join(cacheDir, '7zr.exe');
    const sfxPath = path.join(cacheDir, path.basename(new URL(sevenZip.sfx.url).pathname) || '7z-x64.exe');
    await downloadVerified(sevenZip.sevenZr.url, zrPath, { sha256: sevenZip.sevenZr.sha256 });
    await downloadVerified(sevenZip.sfx.url, sfxPath, { sha256: sevenZip.sfx.sha256 });
    const full = path.join(cacheDir, '7z-full');
    const exe = path.join(full, '7z.exe');
    if (!existsSync(exe)) {
      mkdirSync(full, { recursive: true });
      const r = spawnSync(zrPath, ['x', sfxPath, `-o${full}`, '-y'], { encoding: 'utf8', windowsHide: true });
      if (r.status !== 0) throw new Error(`Failed to extract 7-Zip SFX: ${(r.stderr || r.stdout || '').trim()}`);
    }
    if (!existsSync(exe) || !existsSync(path.join(full, '7z.dll'))) {
      throw new Error('7z.exe / 7z.dll missing after SFX extraction');
    }
    return exe;
  }

  throw new Error(
    'No 7-Zip on PATH and no bootstrap URLs for this platform. Pass sevenZip pins (WATCHER manifest) or install 7-Zip.'
  );
}

export async function extractArchive(archivePath, outDir, { sevenZipBin, kind } = {}) {
  const header = readMagic(archivePath, 8);
  const k = kind || detectArchiveKind(archivePath, header);
  mkdirSync(outDir, { recursive: true });
  if (k === 'zip') return extractZip(archivePath, outDir);
  if (k === 'tar') {
    const r = spawnSync('tar', ['-xf', archivePath, '-C', outDir], { encoding: 'utf8', windowsHide: true });
    if (r.status !== 0) throw new Error(`tar failed: ${(r.stderr || r.stdout || '').trim()}`);
    return { method: 'tar' };
  }
  if (k === '7z' || k === 'exe' || k === 'nsis') {
    if (!sevenZipBin) throw new Error(`7-Zip is required to extract ${k} archive ${archivePath}`);
    if (k === 'exe' && !sevenZipCanOpen(sevenZipBin, archivePath)) {
      throw new Error(`File is a Windows executable but not an extractable archive: ${archivePath}`);
    }
    runSevenZip(sevenZipBin, ['x', archivePath, `-o${outDir}`, '-y']);
    return { method: '7z' };
  }
  throw new Error(`Cannot extract ${k} archive ${archivePath}`);
}

export function walkFiles(root, acc = []) {
  if (!existsSync(root)) return acc;
  for (const name of readdirSync(root, { withFileTypes: true })) {
    const p = path.join(root, name.name);
    if (name.name === '$PLUGINSDIR' || name.name === '.' || name.name === '..') continue;
    if (name.isDirectory()) walkFiles(p, acc);
    else acc.push(p);
  }
  return acc;
}

function findFileByNames(root, names) {
  const want = new Set(names.map((n) => n.toLowerCase()));
  const files = walkFiles(root);
  const matches = files.filter((f) => want.has(path.basename(f).toLowerCase()));
  matches.sort((a, b) => a.split(path.sep).length - b.split(path.sep).length || a.length - b.length);
  return matches[0] || null;
}

export function findTesseractBin(root) {
  const names = process.platform === 'win32' ? ['tesseract.exe'] : ['tesseract', 'tesseract.exe'];
  return findFileByNames(root, names);
}

const TRAINING_TOOL = /^(text2image|lstmtraining|lstmeval|cntraining|mftraining|classifier_tester|shapeclustering|unicharset_extractor|combine_lang_model|combine_tessdata|wordlist2dawg|dawg2wordlist|ambiguous_words|merge_unicharsets|set_unicharset_properties|winpath)(\.exe)?$/i;

export function copyRuntimeBinaries(extractedDir, destDir, { includeTrainingTools = false } = {}) {
  const bin = findTesseractBin(extractedDir);
  if (!bin) throw new Error(`tesseract binary not found under ${extractedDir}`);
  const destBin = path.join(destDir, 'bin');
  mkdirSync(destBin, { recursive: true });
  const destName = process.platform === 'win32' ? 'tesseract.exe' : 'tesseract';
  const destBinPath = path.join(destBin, destName);
  copyFileSync(bin, destBinPath);
  if (process.platform !== 'win32') {
    try { chmodSync(destBinPath, 0o755); } catch { /* ignore */ }
  }
  const srcDir = path.dirname(bin);
  for (const name of readdirSync(srcDir)) {
    if (!includeTrainingTools && TRAINING_TOOL.test(name)) continue;
    const ext = path.extname(name).toLowerCase();
    const isLib = ext === '.dll' || ext === '.so' || ext === '.dylib' || /\.so\.\d/.test(name);
    if (!isLib) continue;
    copyFileSync(path.join(srcDir, name), path.join(destBin, name));
  }
  const siblingLib = path.join(path.dirname(srcDir), 'lib');
  if (existsSync(siblingLib) && statSync(siblingLib).isDirectory()) {
    const destLib = path.join(destDir, 'lib');
    mkdirSync(destLib, { recursive: true });
    for (const name of readdirSync(siblingLib)) {
      const src = path.join(siblingLib, name);
      if (statSync(src).isFile()) copyFileSync(src, path.join(destLib, name));
    }
  }
  return destBinPath;
}

export function dirSize(root) {
  return walkFiles(root).reduce((n, f) => n + statSync(f).size, 0);
}

export function parseLangSha256(spec) {
  const map = {};
  if (!spec) return map;
  for (const part of String(spec).split(',').map((s) => s.trim()).filter(Boolean)) {
    const idx = part.search(/[=:]/);
    if (idx < 1) throw new Error(`Invalid --lang-sha256 entry '${part}' (expected lang=hex)`);
    map[part.slice(0, idx)] = part.slice(idx + 1).toLowerCase();
  }
  return map;
}

