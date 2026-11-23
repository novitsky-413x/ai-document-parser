import assert from 'node:assert/strict';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { test } from 'node:test';
import type { OcrProvider, OcrRuntime } from '../../src/types.js';
import { ocrImageFile, ocrPdfPages, tesseractRecognizeArgs } from '../../src/parse/pdf-ocr.js';
import { languageFlag } from '../../src/parse/ocr-options.js';

const tsv = [
  'level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\twidth\theight\tconf\ttext',
  '5\t1\t1\t1\t1\t1\t10\t20\t30\t12\t99\tHello',
].join('\n');

function tsvWord(text: string): string {
  return [
    'level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\twidth\theight\tconf\ttext',
    `5\t1\t1\t1\t1\t1\t10\t20\t30\t12\t99\t${text}`,
  ].join('\n');
}

function fakeRuntime(): OcrRuntime {
  return {
    bin: 'tesseract',
    popplerBin: 'pdftoppm',
    tessdataDir: '/tessdata',
    bundled: false,
    source: 'path',
  };
}

test('tesseractRecognizeArgs include languages, psm, oem and tessdata-dir', () => {
  const args = tesseractRecognizeArgs({
    imagePath: '/tmp/page.png',
    page: 1,
    languages: ['rus', 'eng'],
    psm: 6,
    oem: 1,
    timeoutMs: 1000,
    runtime: fakeRuntime(),
  });
  assert.deepEqual(args, [
    '/tmp/page.png',
    'stdout',
    '-l',
    'rus+eng',
    '--psm',
    '6',
    '--oem',
    '1',
    '--tessdata-dir',
    '/tessdata',
    'tsv',
  ]);
  assert.equal(languageFlag(['deu']), 'deu');
});

test('ocrImageFile honors ocr:false and does not spawn tesseract', async () => {
  let ran = 0;
  const result = await ocrImageFile('/tmp/scan.png', false, {
    run: async () => {
      ran += 1;
      return { stdout: tsv, stderr: '' };
    },
  });
  assert.equal(ran, 0);
  assert.deepEqual(result.words, []);
  assert.deepEqual(result.warnings, []);
});

test('ocrImageFile returns OCR_LANG_MISSING without recognizing', async () => {
  const calls: string[][] = [];
  const result = await ocrImageFile('/tmp/scan.png', { languages: ['deu'] }, {
    resolveRuntime: () => fakeRuntime(),
    run: async (_cmd, args) => {
      calls.push(args);
      if (args.includes('--list-langs')) {
        return { stdout: 'List of available languages (2):\neng\nrus\n', stderr: '' };
      }
      return { stdout: tsv, stderr: '' };
    },
  });
  assert.equal(result.words.length, 0);
  assert.equal(result.warnings[0]?.code, 'OCR_LANG_MISSING');
  assert.match(result.warnings[0]?.message ?? '', /deu/);
  assert.ok(result.warnings[0]?.message?.includes('eng'));
  assert.ok(calls.every((args) => args.includes('--list-langs')));
});

test('ocrImageFile uses a custom provider and skips tesseract', async () => {
  const provider: OcrProvider = {
    async recognize(request) {
      assert.equal(request.imagePath, '/tmp/scan.png');
      assert.deepEqual(request.languages, ['eng']);
      return {
        words: [{ text: 'Custom', x: 1, y: -2, width: 8, height: 10, page: 1 }],
      };
    },
  };
  let ran = 0;
  const result = await ocrImageFile('/tmp/scan.png', { languages: ['eng'], provider }, {
    run: async () => {
      ran += 1;
      return { stdout: '', stderr: '' };
    },
  });
  assert.equal(ran, 0);
  assert.equal(result.words[0]?.text, 'Custom');
  assert.deepEqual(result.warnings, []);
});

test('ocrPdfPages continues after a per-page failure', async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), 'adp-ocr-pipe-'));
  try {
    const created: string[] = [];
    const result = await ocrPdfPages('/tmp/doc.pdf', [1, 2, 3], { languages: ['eng'] }, {
      resolveRuntime: () => fakeRuntime(),
      existsSync: (filePath) => created.includes(filePath),
      run: async (cmd, args) => {
        if (args.includes('--list-langs')) {
          return { stdout: 'List of available languages (1):\neng\n', stderr: '' };
        }
        if (String(cmd).includes('pdftoppm') || args.includes('-png')) {
          const base = args[args.length - 1];
          const img = `${base}.png`;
          await writeFile(img, 'png');
          created.push(img);
          return { stdout: '', stderr: '' };
        }
        const img = args[0];
        if (img.includes('page-2')) throw new Error('blob died');
        const label = img.includes('page-1') ? 'One' : 'Three';
        return { stdout: tsvWord(label), stderr: '' };
      },
    });
    assert.deepEqual(
      result.words.map((word) => word.text),
      ['One', 'Three']
    );
    assert.deepEqual(result.usedPages, [1, 3]);
    const failed = result.warnings.filter((warning) => warning.code === 'PDF_OCR_FAILED');
    assert.equal(failed.length, 1);
    assert.match(failed[0].message, /page 2/);
    assert.match(failed[0].message, /blob died/);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('ocrPdfPages respects pageCap and custom dpi/psm', async () => {
  const created: string[] = [];
  const dpis: string[] = [];
  const result = await ocrPdfPages(
    '/tmp/doc.pdf',
    [1, 2, 3],
    { languages: ['eng'], pageCap: 2, dpi: 150, psm: 4 },
    {
      resolveRuntime: () => fakeRuntime(),
      existsSync: (filePath) => created.includes(filePath),
      run: async (cmd, args) => {
        if (args.includes('--list-langs')) {
          return { stdout: 'eng\n', stderr: '' };
        }
        if (String(cmd).includes('pdftoppm') || args.includes('-png')) {
          const rIdx = args.indexOf('-r');
          dpis.push(args[rIdx + 1]);
          const base = args[args.length - 1];
          const img = `${base}.png`;
          await writeFile(img, 'png');
          created.push(img);
          return { stdout: '', stderr: '' };
        }
        assert.ok(args.includes('--psm'));
        assert.equal(args[args.indexOf('--psm') + 1], '4');
        return { stdout: tsvWord('X'), stderr: '' };
      },
    }
  );
  assert.equal(result.usedPages.length, 2);
  assert.equal(result.warnings.some((warning) => warning.code === 'PDF_OCR_TRUNCATED'), true);
  assert.match(result.warnings.find((warning) => warning.code === 'PDF_OCR_TRUNCATED')?.message ?? '', /first 2 of 3/);
  assert.deepEqual(dpis, ['150', '150']);
});

test('ocrPdfPages reports PDF_OCR_UNAVAILABLE in English when pdftoppm is missing', async () => {
  const result = await ocrPdfPages('/tmp/doc.pdf', [1], { languages: ['eng'] }, {
    resolveRuntime: () => fakeRuntime(),
    run: async (_cmd, args) => {
      if (args.includes('--list-langs')) return { stdout: 'eng\n', stderr: '' };
      const err = new Error('spawn pdftoppm ENOENT');
      (err as NodeJS.ErrnoException).code = 'ENOENT';
      throw err;
    },
  });
  assert.equal(result.words.length, 0);
  assert.equal(result.warnings[0]?.code, 'PDF_OCR_UNAVAILABLE');
  assert.match(result.warnings[0]?.message ?? '', /pdftoppm was not found/);
  assert.ok(!/[А-Яа-яЁё]/.test(result.warnings[0]?.message ?? ''));
});
