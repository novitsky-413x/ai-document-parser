import assert from 'node:assert/strict';
import { mkdirSync, writeFileSync, readFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { test } from 'node:test';
import { deflateRawSync } from 'node:zlib';
import {
  assertSha256,
  copyRuntimeBinaries,
  detectArchiveKind,
  extractZipBuffer,
  findTesseractBin,
  formatBytes,
  parseLangSha256,
  sha256Buffer,
} from '../../scripts/ocr-archive.mjs';

test('assertSha256 accepts a matching digest and rejects a mismatch', () => {
  const digest = sha256Buffer(Buffer.from('hello'));
  assertSha256(digest, digest, 'hello');
  assert.throws(() => assertSha256(digest, '0'.repeat(64), 'hello'), /SHA-256 mismatch/);
  assert.throws(() => assertSha256(digest, 'not-a-hash', 'hello'), /Invalid SHA-256/);
});

test('formatBytes and parseLangSha256', () => {
  assert.equal(formatBytes(512), '512 B');
  assert.match(formatBytes(4113088), /KiB|MiB/);
  assert.deepEqual(parseLangSha256('rus=abc,eng:def'), { rus: 'abc', eng: 'def' });
  assert.throws(() => parseLangSha256('nocolon'), /Invalid --lang-sha256/);
});

test('detectArchiveKind uses extension and magic', () => {
  assert.equal(detectArchiveKind('foo.zip', Buffer.from([0x50, 0x4b, 0x03, 0x04])), 'zip');
  assert.equal(detectArchiveKind('foo.7z', Buffer.from([0x37, 0x7a, 0xbc, 0xaf, 0x27, 0x1c])), '7z');
  assert.equal(detectArchiveKind('setup.exe', Buffer.from('MZ')), 'exe');
  assert.equal(detectArchiveKind('tesseract', Buffer.from('ELF')), 'raw');
});

function buildStoreZip(entries: { name: string; data: Buffer }[]): Buffer {
  const localParts: Buffer[] = [];
  const centralParts: Buffer[] = [];
  let offset = 0;
  for (const entry of entries) {
    const name = Buffer.from(entry.name, 'utf8');
    const data = entry.data;
    const compact = deflateRawSync(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(8, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0, 12);
    local.writeUInt32LE(0, 14);
    local.writeUInt32LE(compact.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(name.length, 26);
    local.writeUInt16LE(0, 28);
    localParts.push(local, name, compact);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0, 8);
    central.writeUInt16LE(8, 10);
    central.writeUInt32LE(0, 12);
    central.writeUInt32LE(compact.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    centralParts.push(central, name);
    offset += local.length + name.length + compact.length;
  }
  const locals = Buffer.concat(localParts);
  const centrals = Buffer.concat(centralParts);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centrals.length, 12);
  eocd.writeUInt32LE(locals.length, 16);
  return Buffer.concat([locals, centrals, eocd]);
}

test('extractZipBuffer writes deflated files and rejects path traversal', () => {
  const dir = path.join(os.tmpdir(), `adp-zip-${process.pid}-${Date.now()}`);
  mkdirSync(dir, { recursive: true });
  const zip = buildStoreZip([
    { name: 'bin/tesseract.exe', data: Buffer.from('tess-bin') },
    { name: 'bin/libtesseract-5.dll', data: Buffer.from('dll') },
  ]);
  const out = path.join(dir, 'out');
  const result = extractZipBuffer(zip, out);
  assert.equal(readFileSync(path.join(out, 'bin', 'tesseract.exe'), 'utf8'), 'tess-bin');
  assert.ok(result.files.includes('bin/tesseract.exe'));
  const evil = buildStoreZip([{ name: '../escape.txt', data: Buffer.from('nope') }]);
  assert.throws(() => extractZipBuffer(evil, path.join(dir, 'safe')), /unsafe archive path/);
});

test('copyRuntimeBinaries keeps tesseract and DLLs and drops training tools', () => {
  const dir = path.join(os.tmpdir(), `adp-rt-${process.pid}-${Date.now()}`);
  const src = path.join(dir, 'extracted');
  const dest = path.join(dir, 'ocr');
  mkdirSync(src, { recursive: true });
  const binName = process.platform === 'win32' ? 'tesseract.exe' : 'tesseract';
  writeFileSync(path.join(src, binName), 'exe');
  writeFileSync(path.join(src, 'libtesseract-5.dll'), 'dll');
  writeFileSync(path.join(src, 'lstmtraining.exe'), 'tool');
  const installed = copyRuntimeBinaries(src, dest);
  assert.equal(findTesseractBin(src), path.join(src, binName));
  assert.ok(installed.endsWith(path.join('bin', binName)));
  assert.equal(readFileSync(path.join(dest, 'bin', binName), 'utf8'), 'exe');
  assert.equal(readFileSync(path.join(dest, 'bin', 'libtesseract-5.dll'), 'utf8'), 'dll');
  assert.equal(
    (() => {
      try {
        readFileSync(path.join(dest, 'bin', 'lstmtraining.exe'));
        return 'present';
      } catch {
        return 'absent';
      }
    })(),
    'absent'
  );
});

