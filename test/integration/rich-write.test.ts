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
    const pdf = getCapabilities('pdf')[0];
    assert.deepEqual(pdf.writeModes, ['plain']);
    const xlsx = getCapabilities('xlsx')[0];
    assert.deepEqual(xlsx.writeModes, ['plain']);
  });
});
