import assert from 'node:assert/strict';
import path from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { parseArgs, plan, TESSDATA_FAST, tessdataUrl } from '../../scripts/fetch-ocr-runtime.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..');

test('parseArgs defaults to rus+eng tessdata and --bin copy', () => {
  const opts = parseArgs([]);
  assert.deepEqual(opts.langs, ['rus', 'eng']);
  assert.equal(opts.bin, 'copy');
  assert.equal(opts.dryRun, false);
  assert.equal(opts.dest, path.join(root, 'resources', 'ocr'));
});

test('parseArgs accepts dry-run dest and rejects download without a url', () => {
  const opts = parseArgs(['--dry-run', '--dest', '/tmp/ocr-x', '--langs', 'eng']);
  assert.equal(opts.dryRun, true);
  assert.equal(opts.dest, path.resolve('/tmp/ocr-x'));
  assert.deepEqual(opts.langs, ['eng']);
  assert.throws(() => parseArgs(['--bin', 'download']), /--bin-url/);
});

test('plan points tessdata_fast official URLs at dest/tessdata', () => {
  const p = plan(parseArgs(['--dest', '/tmp/ocr-plan', '--bin', 'none', '--langs', 'rus,eng']));
  assert.equal(p.langs.length, 2);
  assert.equal(p.langs[0].url, TESSDATA_FAST.rus);
  assert.equal(p.langs[1].url, TESSDATA_FAST.eng);
  assert.ok(p.langs[0].dest.endsWith(`${path.sep}tessdata${path.sep}rus.traineddata`));
  assert.equal(p.binMode, 'none');
  assert.match(TESSDATA_FAST.eng, /tesseract-ocr\/tessdata_fast/);
});

test('plan builds official tessdata URLs for any language and set', () => {
  const deu = plan(parseArgs(['--dest', '/tmp/ocr-deu', '--bin', 'none', '--langs', 'deu']));
  assert.equal(deu.langs.length, 1);
  assert.equal(deu.langs[0].lang, 'deu');
  assert.equal(deu.langs[0].url, tessdataUrl('deu', 'fast'));
  assert.match(deu.langs[0].url, /tesseract-ocr\/tessdata_fast\/.*deu\.traineddata/);
  const best = plan(parseArgs(['--dest', '/tmp/ocr-best', '--bin', 'none', '--langs', 'fra', '--set', 'best']));
  assert.equal(best.langs[0].url, tessdataUrl('fra', 'best'));
  assert.match(best.langs[0].url, /tessdata_best\/.*fra\.traineddata/);
  const custom = plan(
    parseArgs([
      '--dest',
      '/tmp/ocr-tmpl',
      '--bin',
      'none',
      '--langs',
      'chi_sim',
      '--lang-url',
      'https://example.test/{lang}.traineddata',
    ])
  );
  assert.equal(custom.langs[0].url, 'https://example.test/chi_sim.traineddata');
});

test('parseArgs rejects an unknown --set and plan rejects a bad language id', () => {
  assert.throws(() => parseArgs(['--set', 'turbo']), /fast, best, or standard/);
  assert.throws(
    () => plan(parseArgs(['--dest', '/tmp/x', '--bin', 'none', '--langs', '../eng'])),
    /Invalid OCR language id/
  );
});

test('parseArgs accepts archive verification flags', () => {
  const sha = 'ab'.repeat(32);
  const opts = parseArgs([
    '--bin',
    'download',
    '--bin-url',
    'https://example.test/tesseract-setup.exe',
    '--bin-sha256',
    sha,
    '--lang-sha256',
    `eng=${sha},rus=${sha}`,
    '--cache-dir',
    '/tmp/ocr-cache',
    '--seven-zip-url',
    'https://example.test/7zr.exe',
    '--seven-zip-sha256',
    sha,
  ]);
  assert.equal(opts.bin, 'download');
  assert.equal(opts.binUrl, 'https://example.test/tesseract-setup.exe');
  assert.equal(opts.binSha256, sha);
  assert.equal(opts.langSha256.eng, sha);
  assert.equal(opts.cacheDir, path.resolve('/tmp/ocr-cache'));
  assert.equal(opts.sevenZipUrl, 'https://example.test/7zr.exe');
});
