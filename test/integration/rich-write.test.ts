import { mkdtempSync, rmSync, readFileSync, writeFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { after, before, describe, it } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {
  parseDocument,
  convertDocument,
  editDocument,
  writeDocument,
  getCapabilities,
  DocumentError,
  type ContentBlock,
  type DocxStructured,
} from '../../dist/index.js';

const FIXTURE_MD = readFileSync(
  path.join(
    path.dirname(fileURLToPath(import.meta.url)),
    '../fixtures/complex-roundtrip.md'
  ),
  'utf8'
);
const MARKER = 'E2E-COMPLEX-ROUNDTRIP-2026-08-06';

let dir: string;

before(() => {
  dir = mkdtempSync(path.join(tmpdir(), 'adp-rich-write-'));
});

after(() => {
  rmSync(dir, { recursive: true, force: true });
});

async function readDocumentXml(filePath: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(filePath));
  return zip.file('word/document.xml')!.async('string');
}

function blockTypes(blocks: ContentBlock[] | undefined): string[] {
  return (blocks ?? []).map((b) => b.type);
}

describe('rich writeDocument', () => {
  it('plain write keeps markdown-looking lines as paragraphs and warns', async () => {
    const file = path.join(dir, 'plain.docx');
    const result = await writeDocument(file, {
      type: 'docx',
      content: '# Title\n- Хлеб\n| Q1 | 1200 |',
    });
    assert.equal(result.summary?.mode, 'plain');
    assert.ok(
      result.warnings?.some((w) => w.code === 'PLAIN_WRITE_IGNORED_MARKUP'),
      'expected PLAIN_WRITE_IGNORED_MARKUP'
    );

    const parsed = await parseDocument(file, { output: 'structured' });
    const structured = parsed.structured as DocxStructured;
    assert.ok(structured.blocks.every((b) => b.type === 'paragraph'));
    assert.ok(structured.paragraphs.some((p) => p.includes('- Хлеб')));

    const xml = await readDocumentXml(file);
    assert.doesNotMatch(xml, /<w:tbl[\s>]/);
    assert.doesNotMatch(xml, /<w:numPr>/);
  });

  it('markdown write emits headings, lists, and tables', async () => {
    const file = path.join(dir, 'from-md.docx');
    const md = `# Отчёт

Paragraph with **bold**.

- Хлеб
- Молоко

1. Анализ

| Квартал | Выручка |
|---------|---------|
| Q1      | 1200000 |
`;
    const result = await writeDocument(file, {
      type: 'docx',
      format: 'markdown',
      content: md,
    });
    assert.equal(result.summary?.mode, 'markdown');
    assert.ok((result.summary?.blockCounts?.heading ?? 0) >= 1);
    assert.ok((result.summary?.blockCounts?.list_item ?? 0) >= 2);
    assert.ok((result.summary?.blockCounts?.table ?? 0) >= 1);

    const parsed = await parseDocument(file, { output: 'structured' });
    const types = blockTypes((parsed.structured as DocxStructured).blocks);
    assert.ok(types.includes('heading'));
    assert.ok(types.includes('list_item'));
    assert.ok(types.includes('table'));

    const xml = await readDocumentXml(file);
    assert.match(xml, /<w:tbl[\s>]/);
    assert.match(xml, /<w:numPr>/);
    assert.match(xml, /w:val="Heading1"|w:val="Title"|w:pStyle[^>]*Heading/i);
  });

  it('blocks write emits structured DOCX', async () => {
    const file = path.join(dir, 'from-blocks.docx');
    const blocks: ContentBlock[] = [
      { type: 'heading', level: 1, runs: [{ text: 'Тестовый документ' }] },
      {
        type: 'paragraph',
        runs: [
          { text: 'Обычный текст, ' },
          { text: 'жирный', bold: true },
        ],
      },
      { type: 'list_item', level: 0, ordered: false, runs: [{ text: 'Хлеб' }] },
      {
        type: 'table',
        rows: [
          ['Квартал', 'Выручка'],
          ['Q1', '1200000'],
        ],
      },
    ];
    const result = await writeDocument(file, { type: 'docx', blocks });
    assert.equal(result.summary?.mode, 'blocks');

    const parsed = await parseDocument(file, { output: 'structured' });
    const types = blockTypes((parsed.structured as DocxStructured).blocks);
    assert.ok(types.includes('heading'));
    assert.ok(types.includes('list_item'));
    assert.ok(types.includes('table'));

    const xml = await readDocumentXml(file);
    assert.match(xml, /<w:tbl[\s>]/);
  });

  it('markdown write and convertDocument(md→docx) produce equivalent block types', async () => {
    const md = `# Title

- a
- b

| A | B |
|---|---|
| 1 | 2 |
`;
    const mdPath = path.join(dir, 'parity.md');
    writeFileSync(mdPath, md, 'utf8');

    const written = path.join(dir, 'parity-write.docx');
    await writeDocument(written, { type: 'docx', format: 'markdown', content: md });

    const converted = path.join(dir, 'parity-convert.docx');
    const { value } = await convertDocument(mdPath, 'docx');
    writeFileSync(converted, Buffer.from(value as Uint8Array));

    const writeBlocks = blockTypes(
      ((await parseDocument(written, { output: 'structured' })).structured as DocxStructured)
        .blocks
    );
    const convertBlocks = blockTypes(
      ((await parseDocument(converted, { output: 'structured' })).structured as DocxStructured)
        .blocks
    );
    assert.deepEqual(writeBlocks, convertBlocks);
  });

  it('odt / rtf accept markdown mode', async () => {
    const md = '# H\n\n- item\n';
    const odt = path.join(dir, 'rich.odt');
    const rtf = path.join(dir, 'rich.rtf');
    const odtResult = await writeDocument(odt, {
      type: 'odt',
      format: 'markdown',
      content: md,
    });
    const rtfResult = await writeDocument(rtf, {
      type: 'rtf',
      format: 'markdown',
      content: md,
    });
    assert.equal(odtResult.summary?.mode, 'markdown');
    assert.equal(rtfResult.summary?.mode, 'markdown');

    const odtTypes = blockTypes(
      ((await parseDocument(odt, { output: 'structured' })).structured as { blocks?: ContentBlock[] })
        .blocks
    );
    assert.ok(odtTypes.includes('heading'));
    assert.ok(odtTypes.includes('list_item'));
  });

  it('docx AST from markdown write is non-empty', async () => {
    const file = path.join(dir, 'ast.docx');
    await writeDocument(file, {
      type: 'docx',
      format: 'markdown',
      content: '# T\n\n- a\n\n| A | B |\n|---|---|\n| 1 | 2 |\n',
    });
    const { ast } = await parseDocument(file, { output: 'ast' });
    assert.ok(ast);
    assert.equal(ast!.type, 'flow');
    if (ast!.type !== 'flow') throw new Error('expected flow');
    assert.ok(ast.blocks.length > 0);
    const types = ast.blocks.map((b) => (b as { type: string }).type);
    assert.ok(types.includes('heading'));
    assert.ok(types.includes('list'));
    assert.ok(types.includes('table'));
  });

  it('markdown blockquote round-trips for docx/odt/rtf', async () => {
    const md = '> Цитата: важный тезис\n';
    const mdPath = path.join(dir, 'quote.md');
    writeFileSync(mdPath, md, 'utf8');

    const mdParsed = await parseDocument(mdPath, { output: 'structured' });
    const mdBlocks = (mdParsed.structured as { blocks: ContentBlock[] }).blocks;
    assert.ok(mdBlocks.some((b) => b.type === 'blockquote'));

    for (const type of ['docx', 'odt', 'rtf'] as const) {
      const file = path.join(dir, `quote.${type}`);
      if (type === 'docx') {
        await writeDocument(file, { type: 'docx', format: 'markdown', content: md });
      } else if (type === 'odt') {
        await writeDocument(file, { type: 'odt', format: 'markdown', content: md });
      } else {
        await writeDocument(file, { type: 'rtf', format: 'markdown', content: md });
      }
      const parsed = await parseDocument(file, { output: 'structured' });
      const blocks = (parsed.structured as { blocks: ContentBlock[] }).blocks;
      assert.ok(
        blocks.some((b) => b.type === 'blockquote'),
        `${type} should preserve blockquote`
      );
      assert.match(parsed.text, /Цитата: важный тезис/);
      assert.equal(
        blocks.some(
          (b) =>
            b.type === 'paragraph' &&
            b.runs.some((r) => r.text.trimStart().startsWith('>'))
        ),
        false
      );
    }

    const docxPath = path.join(dir, 'quote.docx');
    const converted = await convertDocument(docxPath, 'md');
    assert.match(String(converted.value), /^>\s*Цитата:/m);
  });

  it('rejects mixing blocks with content', async () => {
    await assert.rejects(
      () =>
        writeDocument(path.join(dir, 'bad.docx'), {
          type: 'docx',
          blocks: [{ type: 'paragraph', runs: [{ text: 'x' }] }],
          content: 'y',
        } as never),
      (err: unknown) => err instanceof DocumentError && err.code === 'INVALID_ARGUMENT'
    );
  });

  it('getCapabilities reports writeModes for flow writers', () => {
    const docx = getCapabilities('docx')[0];
    assert.deepEqual(docx.writeModes, ['plain', 'blocks', 'markdown', 'html']);
    const pptx = getCapabilities('pptx')[0];
    assert.equal(pptx.write, true);
    assert.deepEqual(pptx.writeModes, ['plain', 'blocks', 'markdown', 'html']);
    const pdf = getCapabilities('pdf')[0];
    assert.deepEqual(pdf.writeModes, ['plain']);
    const xlsx = getCapabilities('xlsx')[0];
    assert.deepEqual(xlsx.writeModes, ['plain']);
  });

    it('pptx write emits Office-required package parts and shape geometry', async () => {
    const file = path.join(dir, 'package.pptx');
    await writeDocument(file, {
      type: 'pptx',
      format: 'markdown',
      content: '# Title\n\nBody line\n\n| A | B |\n|---|---|\n| 1 | 2 |\n',
    });
    const zip = await JSZip.loadAsync(readFileSync(file));
    const names = Object.keys(zip.files);
    for (const part of [
      '[Content_Types].xml',
      '_rels/.rels',
      'ppt/presentation.xml',
      'ppt/_rels/presentation.xml.rels',
      'ppt/theme/theme1.xml',
      'ppt/slideMasters/slideMaster1.xml',
      'ppt/slideMasters/_rels/slideMaster1.xml.rels',
      'ppt/slideLayouts/slideLayout1.xml',
      'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
      'ppt/slides/slide1.xml',
      'ppt/slides/_rels/slide1.xml.rels',
    ]) {
      assert.ok(names.includes(part), `missing package part: ${part}`);
    }
    const presentation = await zip.file('ppt/presentation.xml').async('string');
    assert.match(presentation, /sldMasterIdLst/);
    assert.match(presentation, /sldSz/);
    const slide = await zip.file('ppt/slides/slide1.xml').async('string');
    assert.match(slide, /<a:xfrm>/);
    assert.match(slide, /<a:prstGeom/);
    assert.match(slide, /graphicFrame/);
    const ct = await zip.file('[Content_Types].xml').async('string');
    assert.match(ct, /slideMaster\+xml/);
    assert.match(ct, /slideLayout\+xml/);
    assert.match(ct, /theme\+xml/);
  });

  it('pptx markdown write round-trips parse and surgical edit', async () => {
    const file = path.join(dir, 'deck.pptx');
    const result = await writeDocument(file, {
      type: 'pptx',
      format: 'markdown',
      content: '# Введение\n\nПервый слайд.\n\n## Итоги\n\nВторой слайд.\n',
    });
    assert.equal(result.summary?.mode, 'markdown');

    const parsed = await parseDocument(file, { output: 'structured' });
    assert.match(parsed.text, /Введение/);
    assert.match(parsed.text, /Итоги/);
    assert.match(parsed.text, /Первый слайд/);

    const edited = await editDocument(file, {
      type: 'pptx',
      op: 'replace',
      oldString: 'Первый слайд.',
      newString: 'Обновлённый слайд.',
      slide: 1,
    });
    assert.equal(edited.mode, 'surgical');
    assert.equal(edited.changes, 1);

    const again = await parseDocument(file, { output: 'structured' });
    assert.match(again.text, /Обновлённый слайд/);
    assert.doesNotMatch(again.text, /Первый слайд/);
  });

  it('pptx markdown write keeps tables and structural list_item blocks', async () => {
    const file = path.join(dir, 'structured.pptx');
    const result = await writeDocument(file, {
      type: 'pptx',
      format: 'markdown',
      content:
        '# Title\n\nParagraph with **bold** and a [link](https://example.com).\n\n> quote me\n\n```ts\nconst x = 1;\n```\n\n- item one\n- item two\n\n| A | B |\n|---|---|\n| 1 | 2 |\n',
    });
    assert.ok(
      result.warnings?.some((w) => w.code === 'PPTX_INLINE_FLATTENED'),
      'expected PPTX_INLINE_FLATTENED'
    );
    assert.ok(
      result.warnings?.some((w) => w.code === 'PPTX_BLOCKQUOTE_UNWRAPPED'),
      'expected PPTX_BLOCKQUOTE_UNWRAPPED'
    );
    assert.ok(
      result.warnings?.some((w) => w.code === 'PPTX_CODE_AS_PARAGRAPHS'),
      'expected PPTX_CODE_AS_PARAGRAPHS'
    );

    const parsed = await parseDocument(file, { output: 'structured' });
    const structured = parsed.structured as {
      slides: string[][];
      blocks: ContentBlock[];
    };
    assert.ok(structured.slides[0].includes('item one'));
    assert.ok(structured.slides[0].includes('item two'));
    assert.equal(
      structured.slides[0].some((t) => t.includes('·')),
      false,
      'table cells must not be flattened with ·'
    );
    const types = blockTypes(structured.blocks);
    assert.ok(
      types.includes('list_item'),
      `expected list_item in blocks, got ${types.join(',')}`
    );
    assert.ok(types.includes('table'));
    const listItems = structured.blocks.filter((b) => b.type === 'list_item');
    assert.ok(listItems.length >= 2);
    assert.ok(
      listItems.some((b) =>
        b.type === 'list_item' && b.runs.map((r) => r.text).join('').includes('item one')
      )
    );
    const table = structured.blocks.find((b) => b.type === 'table') as Extract<
      ContentBlock,
      { type: 'table' }
    >;
    assert.deepEqual(table.rows[0], ['A', 'B']);
    assert.deepEqual(table.rows[1], ['1', '2']);
  });
});

describe('complex-roundtrip fixture matrix', () => {
  let dir: string;

  before(() => {
    dir = mkdtempSync(path.join(tmpdir(), 'adp-complex-rt-'));
  });

  after(() => {
    rmSync(dir, { recursive: true, force: true });
  });

  for (const type of ['docx', 'odt', 'rtf'] as const) {
    it(`flow ${type}: markdown write → parse keeps structure + marker`, async () => {
      const file = path.join(dir, `complex.${type}`);
      const markup = { format: 'markdown' as const, content: FIXTURE_MD };
      if (type === 'docx') await writeDocument(file, { type: 'docx', ...markup });
      else if (type === 'odt') await writeDocument(file, { type: 'odt', ...markup });
      else await writeDocument(file, { type: 'rtf', ...markup });
      const parsed = await parseDocument(file, { output: 'structured' });
      assert.match(parsed.text, new RegExp(MARKER));
      const blocks = (parsed.structured as { blocks?: ContentBlock[] }).blocks;
      const types = blockTypes(blocks);
      assert.ok(types.includes('list_item'), `list_item missing in ${type}: ${types}`);
      assert.ok(types.includes('table'), `table missing in ${type}`);
      assert.ok(
        types.includes('blockquote') || parsed.text.includes('Important blockquote'),
        `blockquote missing in ${type}`
      );
      assert.ok(
        types.includes('code') || parsed.text.includes('export function hello'),
        `code missing in ${type}`
      );

      if (type === 'docx' || type === 'odt') {
        const md = await convertDocument(file, 'md');
        assert.match(String(md.value), new RegExp(MARKER));
        assert.match(String(md.value), /Alpha/);
      }
    });
  }

  it('html/text convert from markdown source keeps marker', async () => {
    const mdPath = path.join(dir, 'source.md');
    writeFileSync(mdPath, FIXTURE_MD);
    const html = await convertDocument(mdPath, 'html');
    assert.match(String(html.value), new RegExp(MARKER));
    const text = await convertDocument(mdPath, 'text');
    assert.match(String(text.value), new RegExp(MARKER));
  });

  it('pptx: structural lists + table; convert md ok; pptx→docx rejected', async () => {
    const file = path.join(dir, 'complex.pptx');
    const written = await writeDocument(file, {
      type: 'pptx',
      format: 'markdown',
      content: FIXTURE_MD,
    });
    assert.ok(written.warnings?.some((w) => w.code.startsWith('PPTX_')));

    const parsed = await parseDocument(file, { output: 'structured' });
    const blocks = (parsed.structured as { blocks?: ContentBlock[] }).blocks;
    const types = blockTypes(blocks);
    assert.ok(types.includes('list_item'), `pptx list_item missing: ${types}`);
    assert.ok(types.includes('table'), `pptx table missing: ${types}`);
    assert.match(parsed.text, new RegExp(MARKER));

    const md = await convertDocument(file, 'md');
    assert.match(String(md.value), /Bullet one on slide 2|Alpha|item/i);

    await assert.rejects(
      () => convertDocument(file, 'docx'),
      (err: unknown) =>
        err instanceof DocumentError && err.code === 'UNSUPPORTED_FORMAT'
    );
  });
});
