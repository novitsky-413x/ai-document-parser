import assert from 'node:assert/strict';
import { test } from 'node:test';
import {
  DEFAULT_OCR_DPI,
  DEFAULT_OCR_LANGUAGES,
  DEFAULT_OCR_PSM,
  DEFAULT_OCR_TIMEOUT_MS,
  normalizeOcrLanguages,
  normalizeOcrOptions,
  ocrOptionsFromCliFlags,
  PDF_OCR_PAGE_CAP,
} from '../../src/parse/ocr-options.js';

test('normalizeOcrOptions defaults match historical auto rus+eng behavior', () => {
  const a = normalizeOcrOptions();
  const b = normalizeOcrOptions(true);
  for (const opts of [a, b]) {
    assert.equal(opts.enabled, 'auto');
    assert.deepEqual(opts.languages, [...DEFAULT_OCR_LANGUAGES]);
    assert.equal(opts.psm, DEFAULT_OCR_PSM);
    assert.equal(opts.dpi, DEFAULT_OCR_DPI);
    assert.equal(opts.pageCap, PDF_OCR_PAGE_CAP);
    assert.equal(opts.timeoutMs, DEFAULT_OCR_TIMEOUT_MS);
    assert.equal(opts.oem, undefined);
    assert.equal(opts.provider, undefined);
  }
});

test('normalizeOcrOptions false and enabled false skip OCR', () => {
  assert.equal(normalizeOcrOptions(false).enabled, false);
  assert.equal(normalizeOcrOptions({ enabled: false, languages: ['deu'] }).enabled, false);
  assert.deepEqual(normalizeOcrOptions({ enabled: false, languages: ['deu'] }).languages, ['deu']);
});

test('normalizeOcrOptions enabled true forces every PDF page', () => {
  assert.equal(normalizeOcrOptions({ enabled: true }).enabled, true);
});

test('normalizeOcrLanguages dedupes and falls back to rus+eng', () => {
  assert.deepEqual(normalizeOcrLanguages(['deu', ' DEU ', 'eng', '']), ['deu', 'eng']);
  assert.deepEqual(normalizeOcrLanguages([]), ['rus', 'eng']);
  assert.deepEqual(normalizeOcrLanguages(['chi_sim']), ['chi_sim']);
});

test('normalizeOcrOptions clamps numeric fields and keeps paths', () => {
  const opts = normalizeOcrOptions({
    languages: ['eng'],
    bin: ' /opt/tesseract ',
    tessdataDir: '/data/tessdata',
    popplerBin: '/usr/bin/pdftoppm',
    psm: 4,
    oem: 1,
    dpi: 150,
    pageCap: 12,
    timeoutMs: 5000,
  });
  assert.equal(opts.bin, '/opt/tesseract');
  assert.equal(opts.tessdataDir, '/data/tessdata');
  assert.equal(opts.popplerBin, '/usr/bin/pdftoppm');
  assert.equal(opts.psm, 4);
  assert.equal(opts.oem, 1);
  assert.equal(opts.dpi, 150);
  assert.equal(opts.pageCap, 12);
  assert.equal(opts.timeoutMs, 5000);
  assert.equal(normalizeOcrOptions({ psm: -1 }).psm, DEFAULT_OCR_PSM);
  assert.equal(normalizeOcrOptions({ dpi: 0 }).dpi, DEFAULT_OCR_DPI);
});

test('ocrOptionsFromCliFlags maps --ocr --no-ocr and --ocr-lang', () => {
  assert.equal(ocrOptionsFromCliFlags({}), undefined);
  assert.equal(ocrOptionsFromCliFlags({ noOcr: true }), false);
  assert.deepEqual(ocrOptionsFromCliFlags({ ocr: true }), { enabled: 'auto' });
  assert.deepEqual(ocrOptionsFromCliFlags({ ocrLang: 'deu,eng' }), { languages: ['deu', 'eng'] });
  assert.deepEqual(ocrOptionsFromCliFlags({ ocr: true, ocrLang: 'fra' }), {
    enabled: 'auto',
    languages: ['fra'],
  });
  assert.throws(() => ocrOptionsFromCliFlags({ ocr: true, noOcr: true }), /either --ocr or --no-ocr/);
});
