import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
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
});
