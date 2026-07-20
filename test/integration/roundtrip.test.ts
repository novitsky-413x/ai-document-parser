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
    const file = path.join(dir, 'notes.ppt');
    writeFileSync(file, 'not-a-real-ppt', 'utf8');
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

  it('rejects unsupported convert target pdf', async () => {
    const file = path.join(dir, 'c-reject.docx');
    await writeDocument(file, { type: 'docx', content: ['x'] });
    await assert.rejects(
      () => convertDocument(file, 'pdf'),
      (err: unknown) =>
        err instanceof DocumentError && err.code === 'UNSUPPORTED_FORMAT'
    );
  });

  it('converts docx to rtf and epub', async () => {
    const file = path.join(dir, 'c-rich.docx');
    await writeDocument(file, { type: 'docx', content: ['Hello **world**'] });
    const rtf = await convertDocument(file, 'rtf');
    assert.equal(rtf.format, 'rtf');
    assert.match(String(rtf.value), /Hello/);
    const epub = await convertDocument(file, 'epub');
    assert.equal(epub.format, 'epub');
    assert.ok(Buffer.isBuffer(epub.value) || epub.value instanceof Uint8Array);
  });
});

describe('structural convert → md', () => {
  let dir: string;

  before(() => {
    dir = mkdtempSync(path.join(tmpdir(), 'adp-struct-'));
  });

  after(() => {
    rmSync(dir, { recursive: true, force: true });
  });

  it('docx preserves headings, bold, table, and list', async () => {
    const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, BorderStyle } =
      await import('docx');
    const border = { style: BorderStyle.SINGLE, size: 1, color: '000000' };
    const borders = { top: border, bottom: border, left: border, right: border };
    const doc = new Document({
      sections: [
        {
          children: [
            new Paragraph({
              heading: HeadingLevel.HEADING_1,
              children: [new TextRun('Report Title')],
            }),
            new Paragraph({
              children: [
                new TextRun('Hello '),
                new TextRun({ text: 'bold', bold: true }),
                new TextRun(' world'),
              ],
            }),
            new Paragraph({
              bullet: { level: 0 },
              children: [new TextRun('Bullet one')],
            }),
            new Table({
              width: { size: 5000, type: WidthType.DXA },
              rows: [
                new TableRow({
                  children: ['A', 'B'].map(
                    (t) =>
                      new TableCell({
                        borders,
                        width: { size: 2500, type: WidthType.DXA },
                        children: [new Paragraph({ children: [new TextRun(t)] })],
                      })
                  ),
                }),
                new TableRow({
                  children: ['1', '2'].map(
                    (t) =>
                      new TableCell({
                        borders,
                        width: { size: 2500, type: WidthType.DXA },
                        children: [new Paragraph({ children: [new TextRun(t)] })],
                      })
                  ),
                }),
              ],
            }),
          ],
        },
      ],
    });
    const file = path.join(dir, 'rich.docx');
    writeFileSync(file, await Packer.toBuffer(doc));

    const md = await convertDocument(file, 'md');
    const value = String(md.value);
    assert.match(value, /^# Report Title/m);
    assert.match(value, /\*\*bold\*\*/);
    assert.match(value, /\|\s*A\s*\|\s*B\s*\|/);
    assert.match(value, /\|\s*1\s*\|\s*2\s*\|/);
    assert.match(value, /^-\s+Bullet one/m);

    const structured = await parseDocument(file, { output: 'structured' });
    const s = structured.structured as { blocks: unknown[]; paragraphs: string[] };
    assert.ok(Array.isArray(s.blocks) && s.blocks.length >= 3);
    assert.ok(s.paragraphs.some((p) => p.includes('Report Title')));

    await editDocument(file, {
      type: 'docx',
      op: 'replace',
      oldString: 'Report Title',
      newString: 'Updated Title',
    });
    const again = await parseDocument(file, { output: 'structured' });
    assert.ok(
      (again.structured as { paragraphs: string[] }).paragraphs.some((p) =>
        p.includes('Updated Title')
      )
    );
  });

  it('html preserves h2, table, and list in md', async () => {
    const file = path.join(dir, 'struct.html');
    writeFileSync(
      file,
      `<html><body>
        <h2>Section</h2>
        <ul><li>Item A</li><li>Item B</li></ul>
        <table><tr><th>X</th><th>Y</th></tr><tr><td>3</td><td>4</td></tr></table>
      </body></html>`,
      'utf8'
    );
    const md = String((await convertDocument(file, 'md')).value);
    assert.match(md, /^## Section/m);
    assert.match(md, /^-\s+Item A/m);
    assert.match(md, /\|\s*X\s*\|\s*Y\s*\|/);
  });

  it('csv converts to GFM table', async () => {
    const file = path.join(dir, 'struct.csv');
    writeFileSync(file, 'col1,col2\nx,y\n', 'utf8');
    const md = String((await convertDocument(file, 'md')).value);
    assert.match(md, /\|\s*col1\s*\|\s*col2\s*\|/);
    assert.match(md, /\|\s*---/);
    assert.match(md, /\|\s*x\s*\|\s*y\s*\|/);
  });

  it('md to md is passthrough', async () => {
    const file = path.join(dir, 'keep.md');
    const src = '# Keep\n\n- a\n- b\n\n| c | d |\n| --- | --- |\n| 1 | 2 |\n';
    writeFileSync(file, src, 'utf8');
    const md = String((await convertDocument(file, 'md')).value);
    assert.match(md, /# Keep/);
    assert.match(md, /- a/);
    assert.match(md, /\| c \|/);
  });

  it('pptx emits slide headings in md', async () => {
    const file = path.join(dir, 'struct.pptx');
    await writeMinimalPptx(file, [['Title One'], ['Body two']]);
    const md = String((await convertDocument(file, 'md')).value);
    assert.match(md, /## Slide 1/);
    assert.match(md, /Title One/);
    assert.match(md, /## Slide 2/);
  });

  it('docx fidelity: escape, styles, lists, links, notes, dropcap, nested table', async () => {
    const { parseDocxBlocksFromPackage } = await import('../../src/parse/docx-blocks.js');
    const { emitMarkdownFromBlocks } = await import('../../src/convert/blocks.js');

    const stylesXml = `<?xml version="1.0"?>
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:type="character" w:styleId="Strong"><w:rPr><w:b/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="Emphasis"><w:rPr><w:i/></w:rPr></w:style>
      </w:styles>`;

    const numberingXml = `<?xml version="1.0"?>
      <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:abstractNum w:abstractNumId="0">
          <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>
        </w:abstractNum>
        <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
      </w:numbering>`;

    const relsXml = `<?xml version="1.0"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId9" Target="http://example.com/download" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"/>
      </Relationships>`;

    const footnotesXml = `<?xml version="1.0"?>
      <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:footnote w:type="separator" w:id="-1"/>
        <w:footnote w:id="1"><w:p><w:r><w:t>A footnote body.</w:t></w:r></w:p></w:footnote>
      </w:footnotes>`;

    const documentXml = `<?xml version="1.0"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <w:body>
          <w:p><w:r><w:t>Word (2007) and newer.</w:t></w:r></w:p>
          <w:p>
            <w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>bold</w:t></w:r>
            <w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>-italic</w:t></w:r>
          </w:p>
          <w:p>
            <w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>StrongText</w:t></w:r>
            <w:r><w:t> </w:t></w:r>
            <w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>EmphText</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
            <w:r><w:t>First numbered</w:t></w:r>
          </w:p>
          <w:p>
            <w:r><w:t>See </w:t></w:r>
            <w:hyperlink r:id="rId9"><w:r><w:t>download page</w:t></w:r></w:hyperlink>
            <w:r><w:t>.</w:t></w:r>
            <w:r><w:footnoteReference w:id="1"/></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:framePr w:dropCap="drop" w:lines="3"/></w:pPr>
            <w:r><w:t>D</w:t></w:r>
          </w:p>
          <w:p><w:r><w:t>rop caps continue here.</w:t></w:r></w:p>
          <w:tbl>
            <w:tr>
              <w:tc>
                <w:p><w:r><w:t>Outer</w:t></w:r></w:p>
                <w:tbl>
                  <w:tr>
                    <w:tc><w:p><w:r><w:t>One</w:t></w:r></w:p></w:tc>
                    <w:tc><w:p><w:r><w:t>Two</w:t></w:r></w:p></w:tc>
                  </w:tr>
                  <w:tr>
                    <w:tc><w:p><w:r><w:t>Three</w:t></w:r></w:p></w:tc>
                    <w:tc><w:p><w:r><w:t>Four</w:t></w:r></w:p></w:tc>
                  </w:tr>
                </w:tbl>
              </w:tc>
              <w:tc><w:p><w:r><w:t>Right</w:t></w:r></w:p></w:tc>
            </w:tr>
          </w:tbl>
          <w:p>
            <w:pPr><w:pStyle w:val="TOC1"/></w:pPr>
            <w:r><w:t>Chapter One</w:t></w:r>
            <w:r><w:tab/></w:r>
            <w:r><w:t>5</w:t></w:r>
          </w:p>
        </w:body>
      </w:document>`;

    const blocks = parseDocxBlocksFromPackage({
      documentXml,
      stylesXml,
      numberingXml,
      relsXml,
      footnotesXml,
    });
    const md = emitMarkdownFromBlocks(blocks);

    assert.match(md, /Word \(2007\) and newer/);
    assert.doesNotMatch(md, /\\\(/);
    assert.match(md, /\*\*_bold-italic_\*\*/);
    assert.doesNotMatch(md, /\*\*\*\*bold/);
    assert.doesNotMatch(md, /\u200b/);
    assert.match(md, /\*\*StrongText\*\*/);
    assert.match(md, /_EmphText_/);
    assert.match(md, /^1\.\s+First numbered/m);
    assert.match(md, /\[download page\]\(http:\/\/example\.com\/download\)/);
    assert.match(md, /\[\^fn1\]/);
    assert.match(md, /\[\^fn1\]:\s*A footnote body/);
    assert.match(md, /Drop caps continue here/);
    assert.doesNotMatch(md, /^D$/m);
    assert.match(md, /One\s*·\s*Two/);
    assert.doesNotMatch(md, /OneThree/);
    assert.doesNotMatch(md, /One Three \\ \|/);
    assert.match(md, /^-\s+Chapter One$/m);
    assert.doesNotMatch(md, /Chapter One5/);
  });

  it('emits compact markdown tables without empty columns', async () => {
    const { emitMarkdownTable } = await import('../../src/convert/blocks.js');
    const md = emitMarkdownTable([
      ['December 2007', '', '', '', '', '', ''],
      ['Sun', '', 'Mon', '', 'Tue', '', 'Wed'],
      ['1', '', '2', '', '3', '', '4'],
      ['', '', '', '', '', '', ''],
    ]);
    assert.match(md, /December 2007/);
    assert.match(md, /\|\s*Sun\s*\|\s*Mon\s*\|\s*Tue\s*\|\s*Wed\s*\|/);
    assert.match(md, /\|\s*1\s*\|\s*2\s*\|\s*3\s*\|\s*4\s*\|/);
    assert.equal(md.split('\n').length, 4); // header + sep + 2 data rows (empty row dropped)
  });

  it('ignores w:b w:val=0 and keeps footnote defs out of edit paragraphs', async () => {
    const { parseDocxBlocksFromPackage, docxStructuredFromPackage } = await import(
      '../../src/parse/docx-blocks.js'
    );
    const documentXml = `<?xml version="1.0"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>not-bold</w:t></w:r>
            <w:r><w:rPr><w:b/></w:rPr><w:t>yes-bold</w:t></w:r>
            <w:r><w:footnoteReference w:id="1"/></w:r>
          </w:p>
        </w:body>
      </w:document>`;
    const footnotesXml = `<?xml version="1.0"?>
      <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:footnote w:id="1"><w:p><w:r><w:t>note body</w:t></w:r></w:p></w:footnote>
      </w:footnotes>`;
    const { emitMarkdownFromBlocks } = await import('../../src/convert/blocks.js');
    const blocks = parseDocxBlocksFromPackage({ documentXml, footnotesXml });
    const md = emitMarkdownFromBlocks(blocks);
    assert.match(md, /not-bold/);
    assert.doesNotMatch(md, /\*\*not-bold\*\*/);
    assert.match(md, /\*\*yes-bold\*\*/);
    assert.match(md, /\[\^fn1\]:\s*note body/);
    const structured = docxStructuredFromPackage({ documentXml, footnotesXml });
    assert.ok(structured.paragraphs.every((p) => !p.startsWith('[^fn')));
  });
});
