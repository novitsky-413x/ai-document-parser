import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
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
  readDocxStructured,
  DocumentError,
} from '../../dist/index.js';
import { escapeXml } from '../../dist/edit/docx-ooxml.js';

async function writeMinimalPptx(
  filePath: string,
  slides: string[][]
): Promise<void> {
  const zip = new JSZip();
  const overrides = slides
    .map(
      (_, i) =>
        `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
    )
    .join('\n');
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  ${overrides}
</Types>`
  );
  zip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst/>
</p:presentation>`
  );
  slides.forEach((texts, i) => {
    const body = texts
      .map(
        (t) =>
          `<a:p><a:r><a:t>${escapeXml(t)}</a:t></a:r></a:p>`
      )
      .join('');
    zip.file(
      `ppt/slides/slide${i + 1}.xml`,
      `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>${body}</p:spTree></p:cSld>
</p:sld>`
    );
  });
  const buf = await zip.generateAsync({ type: 'nodebuffer' });
  writeFileSync(filePath, buf);
}

describe('DocumentError', () => {
  it('carries code', () => {
    const err = new DocumentError('UNSUPPORTED_FORMAT', 'nope');
    assert.equal(err.code, 'UNSUPPORTED_FORMAT');
    assert.equal(err.name, 'DocumentError');
  });
});

describe('escapeXml', () => {
  it('escapes special characters', () => {
    assert.equal(escapeXml(`a<b>&"c"`), 'a&lt;b&gt;&amp;&quot;c&quot;');
  });
});

describe('write + parse + edit', () => {
  let dir: string;

  before(() => {
    dir = mkdtempSync(path.join(tmpdir(), 'adp-'));
  });

  after(() => {
    rmSync(dir, { recursive: true, force: true });
  });

  it('writes and reads docx text', async () => {
    const file = path.join(dir, 'memo.docx');
    await writeDocument(file, {
      type: 'docx',
      content: 'Hello world\nSecond line',
    });
    const result = await parseDocument(file, { output: 'text' });
    assert.match(result.text, /Hello world/);
    assert.match(result.text, /Second line/);
  });

  it('reads docx structured paragraphs', async () => {
    const file = path.join(dir, 'memo2.docx');
    await writeDocument(file, { type: 'docx', content: ['Alpha', 'Beta'] });
    const structured = await readDocxStructured(file);
    assert.deepEqual(structured.paragraphs, ['Alpha', 'Beta']);
  });

  it('replaces text in docx', async () => {
    const file = path.join(dir, 'contract.docx');
    await writeDocument(file, {
      type: 'docx',
      content: 'Acme Corp signed the deal.',
    });
    await editDocument(file, {
      type: 'docx',
      op: 'replace',
      oldString: 'Acme Corp',
      newString: 'Beta LLC',
    });
    const structured = await readDocxStructured(file);
    const joined = structured.paragraphs.join(' ');
    assert.match(joined, /Beta LLC/);
    assert.doesNotMatch(joined, /Acme Corp/);
  });

  it('writes and patches xlsx', async () => {
    const file = path.join(dir, 'budget.xlsx');
    await writeDocument(file, {
      type: 'xlsx',
      sheets: { Q1: [['Item', 'Amount'], ['Rent', 1500]] },
    });
    const parsed = await parseDocument(file, { output: 'structured' });
    assert.ok(parsed.structured);
    const sheets = (
      parsed.structured as { sheets: Record<string, unknown[][]> }
    ).sheets;
    assert.equal(sheets.Q1[0][0], 'Item');
    assert.equal(sheets.Q1[1][1], 1500);

    await editDocument(file, {
      type: 'spreadsheet',
      op: 'patchCells',
      updates: [{ sheet: 'Q1', cell: 'B2', value: 99 }],
    });
    const again = await parseDocument(file, { output: 'structured' });
    const sheets2 = (
      again.structured as { sheets: Record<string, unknown[][]> }
    ).sheets;
    assert.equal(sheets2.Q1[1][1], 99);
  });

  it('writes plain pdf and parses text', async () => {
    const file = path.join(dir, 'out.pdf');
    await writeDocument(file, {
      type: 'pdf',
      content: 'Текст с кириллицей и Latin',
      title: 'Отчёт',
    });
    const result = await parseDocument(file, { output: 'text' });
    assert.ok(result.text.length > 0);
    assert.equal(result.format, 'pdf');
  });

  it('converts docx to markdown and text', async () => {
    const file = path.join(dir, 'note.docx');
    await writeDocument(file, {
      type: 'docx',
      content: 'Heading line\nBody text here',
    });
    const md = await convertDocument(file, 'md');
    assert.equal(md.format, 'md');
    assert.match(String(md.value), /Heading line/);

    const text = await convertDocument(file, 'text');
    assert.match(String(text.value), /Body text here/);
  });

  it('parses and converts markdown', async () => {
    const file = path.join(dir, 'note.md');
    writeFileSync(file, '# Title\n\nHello **world**\n', 'utf8');
    const parsed = await parseDocument(file, { output: 'text' });
    assert.match(parsed.text, /Hello/);
    const html = await convertDocument(file, 'html');
    assert.match(String(html.value), /Hello/);
  });

  it('parses html and csv', async () => {
    const htmlFile = path.join(dir, 'page.html');
    writeFileSync(htmlFile, '<html><body><p>Alpha</p><p>Beta</p></body></html>', 'utf8');
    const html = await parseDocument(htmlFile, { output: 'text' });
    assert.match(html.text, /Alpha/);
    assert.match(html.text, /Beta/);

    const csvFile = path.join(dir, 'data.csv');
    writeFileSync(csvFile, 'a,b\n1,2\n', 'utf8');
    const csv = await parseDocument(csvFile, { output: 'structured' });
    const rows = (csv.structured as { rows: string[][] }).rows;
    assert.deepEqual(rows[0], ['a', 'b']);
    assert.deepEqual(rows[1], ['1', '2']);
  });

  it('converts xlsx to csv and md', async () => {
    const file = path.join(dir, 'table.xlsx');
    await writeDocument(file, {
      type: 'xlsx',
      sheets: { S: [['A', 'B'], ['1', '2']] },
    });
    const csv = await convertDocument(file, 'csv');
    assert.match(String(csv.value), /A,B/);
    const md = await convertDocument(file, 'md');
    assert.match(String(md.value), /## S/);
  });

  it('rejects unsupported formats', async () => {
    const file = path.join(dir, 'notes.odt');
    writeFileSync(file, 'PK\x03\x04not-really', 'utf8');
    await assert.rejects(
      () => parseDocument(file, { output: 'text' }),
      (err: unknown) =>
        err instanceof DocumentError && err.code === 'UNSUPPORTED_FORMAT'
    );
  });

  it('parses and converts pptx', async () => {
    const file = path.join(dir, 'deck.pptx');
    await writeMinimalPptx(file, [
      ['Title Slide', 'Subtitle RU: привет'],
      ['Second slide body'],
    ]);
    const structured = await parseDocument(file, { output: 'structured' });
    assert.equal(structured.format, 'pptx');
    const slides = (structured.structured as { slides: string[][] }).slides;
    assert.equal(slides.length, 2);
    assert.ok(slides[0].some((t) => t.includes('Title Slide')));
    assert.ok(slides[0].some((t) => t.includes('привет')));

    const md = await convertDocument(file, 'md');
    assert.match(String(md.value), /## Slide 1/);
    assert.match(String(md.value), /Title Slide/);
    assert.match(String(md.value), /## Slide 2/);

    const text = await convertDocument(file, 'text');
    assert.match(String(text.value), /Second slide body/);
  });

  it('appends paragraphs to docx', async () => {
    const file = path.join(dir, 'append.docx');
    await writeDocument(file, { type: 'docx', content: ['First'] });
    await editDocument(file, {
      type: 'docx',
      op: 'append',
      paragraphs: ['APPENDED_LINE'],
    });
    const structured = await readDocxStructured(file);
    assert.ok(structured.paragraphs.some((p) => p.includes('APPENDED_LINE')));
  });

  it('rejects multiple replace matches without replaceAll', async () => {
    const file = path.join(dir, 'multi.docx');
    await writeDocument(file, {
      type: 'docx',
      content: ['ООО Ромашка one', 'ООО Ромашка two'],
    });
    await assert.rejects(
      () =>
        editDocument(file, {
          type: 'docx',
          op: 'replace',
          oldString: 'ООО Ромашка',
          newString: 'ООО Василёк',
        }),
      (err: unknown) =>
        err instanceof DocumentError && err.code === 'MULTIPLE_MATCHES'
    );
    await editDocument(file, {
      type: 'docx',
      op: 'replace',
      oldString: 'ООО Ромашка',
      newString: 'ООО Василёк',
      replaceAll: true,
    });
    const structured = await readDocxStructured(file);
    const joined = structured.paragraphs.join('\n');
    assert.equal((joined.match(/ООО Василёк/g) || []).length, 2);
    assert.doesNotMatch(joined, /ООО Ромашка/);
  });

  it('rejects convert to pdf', async () => {
    const file = path.join(dir, 'plain.docx');
    await writeDocument(file, { type: 'docx', content: 'x' });
    await assert.rejects(
      () => convertDocument(file, 'pdf'),
      (err: unknown) =>
        err instanceof DocumentError && err.code === 'UNSUPPORTED_FORMAT'
    );
  });
});

describe('convertDocument matrix', () => {
  let dir: string;

  before(() => {
    dir = mkdtempSync(path.join(tmpdir(), 'adp-convert-'));
  });

  after(() => {
    rmSync(dir, { recursive: true, force: true });
  });

  it('converts docx to md/html/text/chunks and rejects csv', async () => {
    const file = path.join(dir, 'c-docx.docx');
    await writeDocument(file, {
      type: 'docx',
      content: ['Alpha line', 'Beta <tag> & more'],
    });

    const md = await convertDocument(file, 'md');
    assert.equal(md.format, 'md');
    assert.match(String(md.value), /Alpha line/);
    assert.match(String(md.value), /Beta/);

    const html = await convertDocument(file, 'html');
    assert.equal(html.format, 'html');
    assert.match(String(html.value), /<p>Alpha line<\/p>/);
    assert.match(String(html.value), /&lt;tag&gt;/);
    assert.match(String(html.value), /&amp;/);

    const text = await convertDocument(file, 'text');
    assert.equal(text.format, 'text');
    assert.match(String(text.value), /Alpha line/);

    const chunks = await convertDocument(file, 'chunks');
    assert.equal(chunks.format, 'chunks');
    assert.ok(Array.isArray(chunks.value));
    assert.ok((chunks.value as string[]).some((c) => /Alpha/.test(c)));

    await assert.rejects(
      () => convertDocument(file, 'csv'),
      (err: unknown) =>
        err instanceof DocumentError && err.code === 'UNSUPPORTED_FORMAT'
    );
  });

  it('converts xlsx to csv/md/html/text/chunks', async () => {
    const file = path.join(dir, 'c-sheet.xlsx');
    await writeDocument(file, {
      type: 'xlsx',
      sheets: {
        Sales: [
          ['Item', 'Qty'],
          ['Widget, A', 3],
        ],
      },
    });

    const csv = await convertDocument(file, 'csv');
    assert.equal(csv.format, 'csv');
    assert.match(String(csv.value), /Item,Qty/);
    assert.match(String(csv.value), /"Widget, A",3/);

    const md = await convertDocument(file, 'md');
    assert.match(String(md.value), /## Sales/);
    assert.match(String(md.value), /Widget/);

    const html = await convertDocument(file, 'html');
    assert.match(String(html.value), /Sales|Widget|Qty/);

    const text = await convertDocument(file, 'text');
    assert.match(String(text.value), /## Sales/);

    const chunks = await convertDocument(file, 'chunks');
    assert.ok(Array.isArray(chunks.value));
    assert.ok((chunks.value as string[]).length > 0);
  });

  it('converts pptx to md/html/text/chunks and rejects csv', async () => {
    const file = path.join(dir, 'c-deck.pptx');
    await writeMinimalPptx(file, [['Deck Title'], ['Body slide']]);

    const md = await convertDocument(file, 'md');
    assert.match(String(md.value), /## Slide 1/);
    assert.match(String(md.value), /Deck Title/);

    const html = await convertDocument(file, 'html');
    assert.match(String(html.value), /<h2>Slide 1<\/h2>/);
    assert.match(String(html.value), /Deck Title/);

    const text = await convertDocument(file, 'text');
    assert.match(String(text.value), /Body slide/);

    const chunks = await convertDocument(file, 'chunks');
    assert.ok(Array.isArray(chunks.value));

    await assert.rejects(
      () => convertDocument(file, 'csv'),
      (err: unknown) =>
        err instanceof DocumentError && err.code === 'UNSUPPORTED_FORMAT'
    );
  });

  it('converts pdf to md/html/text/chunks', async () => {
    const file = path.join(dir, 'c-pdf.pdf');
    await writeDocument(file, {
      type: 'pdf',
      content: 'PDF convert probe Cyrillic: проверка',
      title: 'Convert PDF',
    });

    const md = await convertDocument(file, 'md');
    assert.equal(md.format, 'md');
    assert.ok(String(md.value).length > 0);

    const html = await convertDocument(file, 'html');
    assert.match(String(html.value), /<p>|<pre>/);

    const text = await convertDocument(file, 'text');
    assert.ok(String(text.value).length > 0);

    const chunks = await convertDocument(file, 'chunks');
    assert.ok(Array.isArray(chunks.value));
  });

  it('converts markdown and html and csv between text targets', async () => {
    const mdFile = path.join(dir, 'c-note.md');
    writeFileSync(mdFile, '# H\n\nParagraph one.\n', 'utf8');
    const mdToHtml = await convertDocument(mdFile, 'html');
    assert.match(String(mdToHtml.value), /Paragraph one/);
    const mdToText = await convertDocument(mdFile, 'text');
    assert.match(String(mdToText.value), /Paragraph one/);

    const htmlFile = path.join(dir, 'c-page.html');
    writeFileSync(
      htmlFile,
      '<html><body><p>HtmlPara</p><p>Second</p></body></html>',
      'utf8'
    );
    const htmlToMd = await convertDocument(htmlFile, 'md');
    assert.match(String(htmlToMd.value), /HtmlPara/);
    const htmlToText = await convertDocument(htmlFile, 'text');
    assert.match(String(htmlToText.value), /Second/);

    const csvFile = path.join(dir, 'c-data.csv');
    writeFileSync(csvFile, 'name,city\n"A, B",Moscow\n', 'utf8');
    const csvRound = await convertDocument(csvFile, 'csv');
    assert.match(String(csvRound.value), /name,city/);
    const csvToMd = await convertDocument(csvFile, 'md');
    assert.match(String(csvToMd.value), /Moscow|name/);
  });

  it('converts docx from buffer input', async () => {
    const file = path.join(dir, 'c-buf.docx');
    await writeDocument(file, { type: 'docx', content: ['From buffer'] });
    const { readFileSync } = await import('node:fs');
    const buf = readFileSync(file);
    const md = await convertDocument(buf, 'md');
    assert.match(String(md.value), /From buffer/);
  });

  it('rejects unsupported convert targets rtf/epub/pdf', async () => {
    const file = path.join(dir, 'c-reject.docx');
    await writeDocument(file, { type: 'docx', content: ['x'] });
    for (const to of ['pdf', 'rtf', 'epub'] as const) {
      await assert.rejects(
        () => convertDocument(file, to),
        (err: unknown) =>
          err instanceof DocumentError && err.code === 'UNSUPPORTED_FORMAT'
      );
    }
  });
});
