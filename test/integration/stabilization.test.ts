import { execFile } from 'node:child_process';
import { mkdtempSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { promisify } from 'node:util';
import { after, before, describe, it } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import * as XLSX from '@e965/xlsx';
import {
  parseDocument,
  convertDocument,
  editDocument,
  writeDocument,
  getCapabilities,
  DocumentError,
} from '../../dist/index.js';

const execFileAsync = promisify(execFile);
const CLI = path.resolve('dist/cli.js');

let dir: string;

before(() => {
  dir = mkdtempSync(path.join(tmpdir(), 'adp-stab-'));
});

after(() => {
  rmSync(dir, { recursive: true, force: true });
});

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

async function writeMinimalDocx(filePath: string, bodyXml: string): Promise<void> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
  );
  zip.file(
    'word/document.xml',
    `<?xml version="1.0"?><w:document ${W_NS}><w:body>${bodyXml}</w:body></w:document>`
  );
  writeFileSync(filePath, await zip.generateAsync({ type: 'nodebuffer' }));
}

async function readDocumentXml(filePath: string): Promise<string> {
  const { readFile } = await import('node:fs/promises');
  const zip = await JSZip.loadAsync(await readFile(filePath));
  return zip.file('word/document.xml')!.async('string');
}

describe('ODT parse order and duplication', () => {
  it('keeps document order and does not duplicate table cell paragraphs', async () => {
    const file = path.join(dir, 'order.odt');
    const zip = new JSZip();
    zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' });
    zip.file(
      'content.xml',
      `<?xml version="1.0"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
<office:body><office:text>
<text:h text:outline-level="1">Title</text:h>
<text:p>Intro</text:p>
<table:table><table:table-row><table:table-cell><text:p>CellA</text:p></table:table-cell></table:table-row></table:table>
<text:p>After</text:p>
</office:text></office:body></office:document-content>`
    );
    writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer' }));

    const md = String((await convertDocument(file, 'md')).value);
    const cellMatches = md.match(/CellA/g) ?? [];
    assert.equal(cellMatches.length, 1, 'table cell text must appear exactly once');
    assert.ok(md.indexOf('Title') < md.indexOf('Intro'), 'heading before intro');
    assert.ok(md.indexOf('Intro') < md.indexOf('CellA'), 'intro before table');
    assert.ok(md.indexOf('CellA') < md.indexOf('After'), 'table before trailing paragraph');
  });
});

describe('ODS repeated cells and typed values', () => {
  it('expands number-columns-repeated and reads office:value', async () => {
    const file = path.join(dir, 'grid.ods');
    const zip = new JSZip();
    zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet', {
      compression: 'STORE',
    });
    zip.file(
      'content.xml',
      `<?xml version="1.0"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
<office:body><office:spreadsheet>
<table:table table:name="Data">
<table:table-row>
<table:table-cell office:value-type="float" office:value="5"/>
<table:table-cell table:number-columns-repeated="2"><text:p>x</text:p></table:table-cell>
</table:table-row>
</table:table>
</office:spreadsheet></office:body></office:document-content>`
    );
    writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer' }));

    const parsed = await parseDocument(file, { output: 'structured' });
    const sheets = (parsed.structured as { sheets: Record<string, unknown[][]> }).sheets;
    assert.deepEqual(sheets.Data[0], [5, 'x', 'x']);
  });
});

describe('EPUB manifest attribute order', () => {
  it('parses OPF items regardless of attribute order', async () => {
    const file = path.join(dir, 'book.epub');
    const zip = new JSZip();
    zip.file('mimetype', 'application/epub+zip', { compression: 'STORE' });
    zip.file(
      'META-INF/container.xml',
      `<?xml version="1.0"?><container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles></container>`
    );
    zip.file(
      'OEBPS/content.opf',
      `<?xml version="1.0"?><package xmlns="http://www.idpf.org/2007/opf" version="3.0">
<manifest><item href="chapter.xhtml" media-type="application/xhtml+xml" id="ch"/></manifest>
<spine><itemref idref="ch"/></spine></package>`
    );
    zip.file(
      'OEBPS/chapter.xhtml',
      `<html xmlns="http://www.w3.org/1999/xhtml"><body><p>Hello EPUB order</p></body></html>`
    );
    writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer' }));

    const parsed = await parseDocument(file);
    assert.match(parsed.text, /Hello EPUB order/);
  });

  it('round-trips writeDocument epub through parseDocument', async () => {
    const file = path.join(dir, 'written.epub');
    await writeDocument(file, { type: 'epub', content: 'Alpha\nBeta', title: 'T' });
    const parsed = await parseDocument(file);
    assert.match(parsed.text, /Alpha/);
    assert.match(parsed.text, /Beta/);
  });
});

describe('RTF escapes across style toggles', () => {
  it('decodes hex escapes buffered before a style change', async () => {
    const file = path.join(dir, 'esc.rtf');
    writeFileSync(file, "{\\rtf1\\ansi caf\\'e9 \\b bold\\b0 tail\\par}", 'utf8');
    const parsed = await parseDocument(file);
    assert.match(parsed.text, /café/);
    assert.match(parsed.text, /bold/);
  });
});

describe('DOCX surgical edit', () => {
  it('replaces a string spanning multiple runs without duplication', async () => {
    const file = path.join(dir, 'multirun.docx');
    await writeMinimalDocx(
      file,
      `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Ac</w:t></w:r><w:r><w:t>me Corp</w:t></w:r></w:p>`
    );
    await editDocument(file, {
      type: 'docx',
      op: 'replace',
      oldString: 'Acme',
      newString: 'Beta',
    });
    const xml = await readDocumentXml(file);
    const betas = xml.match(/Beta/g) ?? [];
    assert.equal(betas.length, 1, 'newString must be inserted exactly once');
    assert.match(xml, /<w:b\/>/, 'run properties preserved');
    const parsed = await parseDocument(file, { output: 'structured' });
    const paragraphs = (parsed.structured as { paragraphs: string[] }).paragraphs;
    assert.equal(paragraphs[0], 'Beta Corp');
  });

  it('preserves hyperlink wrappers when replacing inside them', async () => {
    const file = path.join(dir, 'link.docx');
    await writeMinimalDocx(
      file,
      `<w:p><w:hyperlink r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:r><w:t>Click here</w:t></w:r></w:hyperlink></w:p>`
    );
    await editDocument(file, {
      type: 'docx',
      op: 'replace',
      oldString: 'Click',
      newString: 'Tap',
    });
    const xml = await readDocumentXml(file);
    assert.match(xml, /<w:hyperlink /, 'hyperlink element preserved');
    assert.match(xml, /Tap here/);
  });

  it('appends paragraphs before w:sectPr', async () => {
    const file = path.join(dir, 'sect.docx');
    await writeMinimalDocx(
      file,
      `<w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`
    );
    await editDocument(file, { type: 'docx', op: 'append', paragraphs: ['Appended'] });
    const xml = await readDocumentXml(file);
    assert.ok(
      xml.indexOf('Appended') < xml.indexOf('<w:sectPr'),
      'appended content must precede sectPr'
    );
  });

  it('leaves no temp files after edits', () => {
    const leftovers = readdirSync(dir).filter((n) => n.endsWith('.tmp'));
    assert.deepEqual(leftovers, []);
  });
});

describe('DOCX w:br parsing', () => {
  it('turns line breaks into newlines in paragraphs', async () => {
    const file = path.join(dir, 'br.docx');
    await writeMinimalDocx(
      file,
      `<w:p><w:r><w:t>Line1</w:t><w:br/><w:t>Line2</w:t></w:r></w:p>`
    );
    const parsed = await parseDocument(file, { output: 'structured' });
    const paragraphs = (parsed.structured as { paragraphs: string[] }).paragraphs;
    assert.equal(paragraphs[0], 'Line1\nLine2');
  });
});

describe('spreadsheet edit safety', () => {
  it('sets a formula when value starts with =', async () => {
    const file = path.join(dir, 'formula.xlsx');
    await writeDocument(file, { type: 'xlsx', sheets: { Data: [[1, 2]] } });
    await editDocument(file, {
      type: 'spreadsheet',
      op: 'patchCells',
      updates: [{ sheet: 'Data', cell: 'C1', value: '=A1+B1' }],
    });
    const wb = XLSX.read(await (await import('node:fs/promises')).readFile(file), {
      type: 'buffer',
    });
    const cell = wb.Sheets.Data.C1 as { f?: string };
    assert.equal(cell.f, 'A1+B1');
  });

  it('rejects invalid sheet names on write', async () => {
    const file = path.join(dir, 'bad.xlsx');
    await assert.rejects(
      writeDocument(file, { type: 'xlsx', sheets: { 'Bad[name]': [[1]] } }),
      (err: unknown) => err instanceof DocumentError && err.code === 'INVALID_ARGUMENT'
    );
  });
});

describe('multi-sheet CSV export', () => {
  it('honors options.sheet and rejects unknown sheets', async () => {
    const file = path.join(dir, 'multi.xlsx');
    await writeDocument(file, {
      type: 'xlsx',
      sheets: { First: [['a']], Second: [['b']] },
    });
    const second = await convertDocument(file, 'csv', { sheet: 'Second' });
    assert.equal(String(second.value).trim(), 'b');
    await assert.rejects(
      convertDocument(file, 'csv', { sheet: 'Nope' }),
      (err: unknown) => err instanceof DocumentError && err.code === 'INVALID_ARGUMENT'
    );
  });
});

describe('capabilities conformance', () => {
  it('every advertised convert target actually converts', async () => {
    const fixtures: Record<string, string> = {};

    const docx = path.join(dir, 'cap.docx');
    await writeDocument(docx, { type: 'docx', content: 'Hello world' });
    fixtures.docx = docx;

    const xlsx = path.join(dir, 'cap.xlsx');
    await writeDocument(xlsx, { type: 'xlsx', sheets: { S1: [['a', 'b']] } });
    fixtures.xlsx = xlsx;

    const pdf = path.join(dir, 'cap.pdf');
    await writeDocument(pdf, { type: 'pdf', content: 'Hello pdf' });
    fixtures.pdf = pdf;

    const odt = path.join(dir, 'cap.odt');
    await writeDocument(odt, { type: 'odt', content: 'Hello odt' });
    fixtures.odt = odt;

    const rtf = path.join(dir, 'cap.rtf');
    await writeDocument(rtf, { type: 'rtf', content: 'Hello rtf' });
    fixtures.rtf = rtf;

    const epub = path.join(dir, 'cap.epub');
    await writeDocument(epub, { type: 'epub', content: 'Hello epub' });
    fixtures.epub = epub;

    const md = path.join(dir, 'cap.md');
    writeFileSync(md, '# Head\n\nBody text', 'utf8');
    fixtures.md = md;

    const html = path.join(dir, 'cap.html');
    writeFileSync(html, '<h1>Head</h1><p>Body</p>', 'utf8');
    fixtures.html = html;

    const csv = path.join(dir, 'cap.csv');
    writeFileSync(csv, 'a,b\n1,2\n', 'utf8');
    fixtures.csv = csv;

    for (const cap of getCapabilities()) {
      const fixture = fixtures[cap.format];
      if (!fixture) continue;
      for (const target of cap.convertTo) {
        const result = await convertDocument(fixture, target);
        assert.ok(
          result.value !== undefined && result.value !== null,
          `${cap.format} → ${target} produced no value`
        );
      }
    }
  });

  it('rejects removed phantom targets at runtime', async () => {
    const xlsx = path.join(dir, 'phantom.xlsx');
    await writeDocument(xlsx, { type: 'xlsx', sheets: { S1: [[1]] } });
    await assert.rejects(
      convertDocument(xlsx, 'ods'),
      (err: unknown) => err instanceof DocumentError && err.code === 'UNSUPPORTED_FORMAT'
    );
  });
});

describe('nested list conversion', () => {
  it('md nested list converts to properly nested html', async () => {
    const file = path.join(dir, 'nested.md');
    writeFileSync(file, '- A\n  - B\n- C\n', 'utf8');
    const html = String((await convertDocument(file, 'html')).value);
    assert.match(html, /<li>A\s*<ul>\s*<li>B\s*<\/li>\s*<\/ul>\s*<\/li>/);
  });
});

describe('CLI', () => {
  it('accepts flags before the positional file', async () => {
    const file = path.join(dir, 'cli.md');
    writeFileSync(file, '# CLI test\n\ncontent line', 'utf8');
    const { stdout } = await execFileAsync(process.execPath, [CLI, 'convert', '--to', 'md', file]);
    assert.match(stdout, /CLI test/);
  });

  it('implements write-xlsx', async () => {
    const out = path.join(dir, 'cli.xlsx');
    await execFileAsync(process.execPath, [
      CLI,
      'write-xlsx',
      out,
      '--sheet',
      'Data',
      '--csv',
      'a,b\n1,2',
    ]);
    const parsed = await parseDocument(out, { output: 'structured' });
    const sheets = (parsed.structured as { sheets: Record<string, unknown[][]> }).sheets;
    assert.deepEqual(sheets.Data[0], ['a', 'b']);
  });

  it('exits non-zero with a clear error for unknown commands', async () => {
    await assert.rejects(
      execFileAsync(process.execPath, [CLI, 'frobnicate']),
      (err: unknown) => (err as { code?: number }).code === 1
    );
  });
});

describe('BOM handling end-to-end', () => {
  it('parses markdown buffers with a BOM', async () => {
    const parsed = await parseDocument(Buffer.from('\uFEFF# BOM Title\n\nBody', 'utf8'));
    assert.match(parsed.text, /BOM Title/);
  });
});
