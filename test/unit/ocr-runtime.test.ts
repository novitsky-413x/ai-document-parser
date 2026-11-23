import assert from 'node:assert/strict';
import path from 'node:path';
import { test } from 'node:test';
import {
  hasAllOcrLanguages,
  missingOcrLanguages,
  normalizeTessdataDir,
  ocrProcessEnv,
  parseTesseractListLangs,
  resolveOcrRuntime,
} from '../../src/parse/ocr-runtime.js';

const binName = process.platform === 'win32' ? 'tesseract.exe' : 'tesseract';
const popplerName = process.platform === 'win32' ? 'pdftoppm.exe' : 'pdftoppm';

function treeExists(files: Set<string>) {
  const norm = (p: string) => path.normalize(p);
  const set = new Set([...files].map(norm));
  return (p: string) => set.has(norm(p));
}

test('resolveOcrRuntime prefers options.bin over env, bundled, and PATH', () => {
  const bundled = path.join('/bundle', 'resources', 'ocr');
  const exists = treeExists(
    new Set([
      path.join(bundled, 'bin', binName),
      path.join(bundled, 'tessdata', 'rus.traineddata'),
      path.join(bundled, 'tessdata', 'eng.traineddata'),
      '/opt/custom/tesseract',
    ])
  );
  const runtime = resolveOcrRuntime({
    bin: '/opt/custom/tesseract',
    tessdataDir: '/opt/custom/tessdata',
    env: { ADP_TESSERACT: '/env/tesseract', TESSDATA_PREFIX: '/env/tessdata' },
    existsSync: exists,
    moduleDir: path.join('/bundle', 'dist', 'parse'),
    cwd: '/tmp',
    resourcesPath: '',
  });
  assert.equal(runtime.bin, '/opt/custom/tesseract');
  assert.equal(runtime.source, 'options');
  assert.equal(runtime.bundled, false);
  assert.equal(runtime.tessdataDir, '/opt/custom/tessdata');
});

test('resolveOcrRuntime uses ADP_TESSERACT and TESSDATA_PREFIX before bundled', () => {
  const bundled = path.join('/bundle', 'resources', 'ocr');
  const exists = treeExists(
    new Set([
      path.join(bundled, 'bin', binName),
      path.join(bundled, 'tessdata', 'rus.traineddata'),
      path.join(bundled, 'tessdata', 'eng.traineddata'),
      '/env/tesseract',
      path.join('/env/prefix', 'tessdata'),
    ])
  );
  const runtime = resolveOcrRuntime({
    env: { ADP_TESSERACT: '/env/tesseract', TESSDATA_PREFIX: '/env/prefix' },
    existsSync: exists,
    moduleDir: path.join('/bundle', 'dist', 'parse'),
    cwd: '/tmp',
    resourcesPath: '',
  });
  assert.equal(runtime.bin, '/env/tesseract');
  assert.equal(runtime.source, 'env');
  assert.equal(runtime.tessdataDir, path.join('/env/prefix', 'tessdata'));
});

test('bundled runtime is valid only when every requested language is present', () => {
  const root = path.join('/app', 'resources', 'ocr');
  const files = new Set([
    path.join(root, 'bin', binName),
    path.join(root, 'tessdata', 'eng.traineddata'),
    path.join(root, 'lib'),
    path.join(root, 'bin', popplerName),
  ]);
  const exists = treeExists(files);
  const input = {
    existsSync: exists,
    moduleDir: path.join('/app', 'dist', 'parse'),
    cwd: '/unrelated',
    resourcesPath: '',
    env: {},
  };
  const engOnly = resolveOcrRuntime({ ...input, languages: ['eng'] });
  assert.equal(engOnly.source, 'bundled');
  assert.equal(engOnly.bundled, true);
  assert.equal(engOnly.bin, path.join(root, 'bin', binName));
  assert.equal(engOnly.tessdataDir, path.join(root, 'tessdata'));
  assert.equal(engOnly.libDir, path.join(root, 'lib'));
  assert.equal(engOnly.popplerBin, path.join(root, 'bin', popplerName));

  const rusEng = resolveOcrRuntime({ ...input, languages: ['rus', 'eng'] });
  assert.equal(rusEng.source, 'path');
  assert.equal(rusEng.bundled, false);
  assert.equal(rusEng.bin, binName);
});

test('normalizeTessdataDir accepts either tessdata or its parent prefix', () => {
  const prefix = '/usr/share/tesseract-ocr/4.00';
  const tess = path.join(prefix, 'tessdata');
  const exists = treeExists(new Set([tess, path.join(tess, 'eng.traineddata')]));
  assert.equal(normalizeTessdataDir(prefix, exists), tess);
  assert.equal(normalizeTessdataDir(tess, exists), tess);
});

test('hasAllOcrLanguages and missingOcrLanguages compare requested set', () => {
  const tess = '/data/tessdata';
  const exists = treeExists(
    new Set([path.join(tess, 'eng.traineddata'), path.join(tess, 'deu.traineddata')])
  );
  assert.equal(hasAllOcrLanguages(tess, ['eng'], exists), true);
  assert.equal(hasAllOcrLanguages(tess, ['eng', 'rus'], exists), false);
  assert.deepEqual(missingOcrLanguages(['rus', 'eng'], ['eng', 'deu']), ['rus']);
});

test('parseTesseractListLangs skips the header and noise', () => {
  const langs = parseTesseractListLangs(`Tesseract Open Source OCR Engine v5.3.0
List of available languages in /tessdata/ (3):
eng
rus
chi_sim
`);
  assert.deepEqual(langs, ['eng', 'rus', 'chi_sim']);
});

test('ocrProcessEnv prepends platform library paths and bin dir for bundled runtimes', () => {
  const runtime = {
    bin: path.join('/ocr', 'bin', binName),
    libDir: path.join('/ocr', 'lib'),
    tessdataDir: path.join('/ocr', 'tessdata'),
    bundled: true,
    source: 'bundled' as const,
  };
  const env = ocrProcessEnv(runtime, { PATH: '/usr/bin', TESSDATA_PREFIX: undefined as unknown as string });
  delete (env as { TESSDATA_PREFIX?: string }).TESSDATA_PREFIX;
  const filled = ocrProcessEnv(runtime, { PATH: '/usr/bin' });
  assert.equal(filled.TESSDATA_PREFIX, path.join('/ocr', 'tessdata'));
  assert.ok(String(filled.PATH).includes(path.join('/ocr', 'bin')));
  if (process.platform === 'win32') {
    assert.ok(String(filled.PATH).includes(path.join('/ocr', 'lib')));
  } else if (process.platform === 'darwin') {
    assert.ok(String(filled.DYLD_LIBRARY_PATH).startsWith(path.join('/ocr', 'lib')));
  } else {
    assert.ok(String(filled.LD_LIBRARY_PATH).startsWith(path.join('/ocr', 'lib')));
  }
});
