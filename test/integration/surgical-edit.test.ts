import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { after, before, describe, it } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import * as XLSX from '@e965/xlsx';
import {
  DocumentError,
  editDocument,
  getCapabilities,
  parseDocument,
  writeDocument,
} from '../../dist/index.js';

let dir: string;

before(() => {
  dir = mkdtempSync(path.join(tmpdir(), 'adp-surgical-'));
});

after(() => {
  rmSync(dir, { recursive: true, force: true });
});

async function saveZip(filePath: string, zip: JSZip): Promise<void> {
  writeFileSync(filePath, await zip.generateAsync({ type: 'nodebuffer' }));
}

async function entryBytes(
  filePath: string,
  entry: string
): Promise<Buffer> {
  const zip = await JSZip.loadAsync(readFileSync(filePath));
  const part = zip.file(entry);
  assert.ok(part, `missing ZIP entry ${entry}`);
  return Buffer.from(await part.async('uint8array'));
}

async function makePptx(filePath: string): Promise<void> {
  const zip = new JSZip();
  zip.file(
    'ppt/presentation.xml',
    '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId1"/></p:sldIdLst></p:presentation>'
  );
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>'
  );
  const slide = (first: string, second: string) =>
    `<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r data-test="x > y"><a:rPr b="1"/><a:t>${first}</a:t></a:r><a:r><a:t>${second}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`;
  zip.file('ppt/slides/slide1.xml', slide('Hello ', 'world'));
  zip.file('ppt/slides/slide2.xml', slide('Hello ', 'world'));
  zip.file(
    'ppt/notesSlides/notesSlide1.xml',
    slide('Private ', 'note')
  );
  zip.file('ppt/media/image.bin', Buffer.from([1, 2, 3, 4]));
  await saveZip(filePath, zip);
}

async function makeOdt(filePath: string): Promise<void> {
  const zip = new JSZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text', {
    compression: 'STORE',
  });
  zip.file(
    'content.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:text>
    <text:p text:style-name="Body > Test"><text:span text:style-name="Strong">Hello</text:span><text:s/>world</text:p>
    <text:p>Other paragraph</text:p>
  </office:text></office:body>
</office:document-content>`
  );
  zip.file('styles.xml', '<styles>preserve-me</styles>');
  zip.file('Pictures/preserve.bin', Buffer.from([7, 8, 9]));
  await saveZip(filePath, zip);
}

async function makeEpub(filePath: string): Promise<void> {
  const zip = new JSZip();
  zip.file('mimetype', 'application/epub+zip', { compression: 'STORE' });
  zip.file(
    'META-INF/container.xml',
    `<?xml version="1.0"?><container xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles></container>`
  );
  zip.file(
    'OEBPS/content.opf',
    `<?xml version="1.0"?><package xmlns="http://www.idpf.org/2007/opf" version="3.0"><manifest><item id="one" href="one.xhtml" media-type="application/xhtml+xml"/><item id="two" href="two.xhtml" media-type="application/xhtml+xml"/></manifest><spine><itemref idref="one"/><itemref idref="two"/></spine></package>`
  );
  zip.file(
    'OEBPS/one.xhtml',
    `<html xmlns="http://www.w3.org/1999/xhtml"><body><p data-label="a > b"><strong>Hello</strong> world<svg xmlns="http://www.w3.org/2000/svg"><svg><text>Hello world</text></svg><text>Hello world</text></svg></p></body></html>`
  );
  zip.file(
    'OEBPS/two.xhtml',
    `<html xmlns="http://www.w3.org/1999/xhtml"><body><p>Hello world</p></body></html>`
  );
  zip.file('OEBPS/image.bin', Buffer.from([4, 3, 2, 1]));
  await saveZip(filePath, zip);
}

describe('clean-room surgical edit', () => {
  it('replaces PPTX text across runs with slide scope and preserves other parts', async () => {
    const file = path.join(dir, 'slides.pptx');
    await makePptx(file);
    const slide1Before = await entryBytes(file, 'ppt/slides/slide1.xml');
    const mediaBefore = await entryBytes(file, 'ppt/media/image.bin');

    const result = await editDocument(file, {
      type: 'pptx',
      op: 'replace',
      oldString: 'Hello world',
      newString: 'Привет мир',
      slide: 1,
    });
    assert.equal(result.mode, 'surgical');
    assert.equal(result.changes, 1);

    const logicalFirst = (await entryBytes(file, 'ppt/slides/slide2.xml')).toString();
    assert.match(logicalFirst, /Привет мир/);
    assert.match(logicalFirst, /<a:rPr b="1"\/>/);
    assert.deepEqual(await entryBytes(file, 'ppt/slides/slide1.xml'), slide1Before);
    assert.deepEqual(await entryBytes(file, 'ppt/media/image.bin'), mediaBefore);

    const parsed = await parseDocument(file, { output: 'structured' });
    assert.match(parsed.text, /Привет мир/);
  });

  it('detects multiple PPTX matches before writing', async () => {
    const file = path.join(dir, 'slides-multiple.pptx');
    await makePptx(file);
    const before = readFileSync(file);
    await assert.rejects(
      editDocument(file, {
        type: 'pptx',
        op: 'replace',
        oldString: 'Hello world',
        newString: 'x',
      }),
      (error: unknown) =>
        error instanceof DocumentError && error.code === 'MULTIPLE_MATCHES'
    );
    assert.deepEqual(readFileSync(file), before);
  });

  it('replaces ODT text across span and repeated-space tokens', async () => {
    const file = path.join(dir, 'text.odt');
    await makeOdt(file);
    const stylesBefore = await entryBytes(file, 'styles.xml');
    const pictureBefore = await entryBytes(file, 'Pictures/preserve.bin');

    const result = await editDocument(file, {
      type: 'odt',
      op: 'replace',
      oldString: 'Hello world',
      newString: 'Привет мир',
    });
    assert.equal(result.mode, 'surgical');
    const content = (await entryBytes(file, 'content.xml')).toString();
    assert.match(content, /text:style-name="Strong"/);
    assert.match(content, /Привет/);
    assert.deepEqual(await entryBytes(file, 'styles.xml'), stylesBefore);
    assert.deepEqual(await entryBytes(file, 'Pictures/preserve.bin'), pictureBefore);

    const parsed = await parseDocument(file, { output: 'structured' });
    assert.match(parsed.text, /Привет мир/);
  });

  it('ODT Cyrillic markdown-write replace leaves no text:s debris', async () => {
    const file = path.join(dir, 'odt-cyr-edit.odt');
    const newString = 'Конец ODT-документа. (изменено через edit_document)';
    await writeDocument(file, {
      type: 'odt',
      format: 'markdown',
      content: 'Конец документа.\n',
    });
    const result = await editDocument(file, {
      type: 'odt',
      op: 'replace',
      oldString: 'Конец документа.',
      newString,
    });
    assert.equal(result.mode, 'surgical');
    assert.equal(result.changes, 1);

    const parsed = await parseDocument(file, { output: 'structured' });
    const blob = JSON.stringify(parsed.structured);
    assert.equal(blob.includes('text:s'), false);
    assert.equal(blob.includes('text:tab'), false);
    assert.match(parsed.text, new RegExp(newString.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')));
  });

  it('ODT replaceAll:false throws MULTIPLE_MATCHES without mutating', async () => {
    const file = path.join(dir, 'odt-multi.odt');
    await writeDocument(file, {
      type: 'odt',
      format: 'markdown',
      content: 'alpha\n\nalpha\n',
    });
    const before = readFileSync(file);
    await assert.rejects(
      editDocument(file, {
        type: 'odt',
        op: 'replace',
        oldString: 'alpha',
        newString: 'beta',
      }),
      (error: unknown) =>
        error instanceof DocumentError && error.code === 'MULTIPLE_MATCHES'
    );
    assert.deepEqual(readFileSync(file), before);
  });

  it('patches ODS cells and reports package rewrite', async () => {
    const file = path.join(dir, 'sheet.ods');
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([
      ['Name', 'Value'],
      ['A', 1],
    ]);
    XLSX.utils.book_append_sheet(wb, ws, 'Data');
    writeFileSync(
      file,
      XLSX.write(wb, { type: 'buffer', bookType: 'ods' }) as Buffer
    );

    const result = await editDocument(file, {
      type: 'spreadsheet',
      op: 'patchCells',
      updates: [
        { sheet: 'Data', cell: 'B2', value: 42 },
        { sheet: 'Data', cell: 'C2', value: '=B2*2' },
        { sheet: 'Data', cell: 'D2', value: '=IF(B2="A1",1,0)' },
      ],
    });
    assert.equal(result.mode, 'rewrite');
    assert.equal(result.changes, 3);

    const edited = XLSX.read(readFileSync(file), { type: 'buffer' });
    assert.equal(edited.Sheets.Data.B2.v, 42);
    assert.equal(edited.Sheets.Data.C2.f, 'B2*2');
    assert.equal(edited.Sheets.Data.D2.f, 'IF(B2="A1",1,0)');
  });

  it('rewrites markdown flow and preserves inline structure', async () => {
    const file = path.join(dir, 'flow.md');
    writeFileSync(file, '# Title\n\nHello **bold world**.\n', 'utf8');
    const result = await editDocument(file, {
      type: 'flow',
      op: 'replace',
      oldString: 'Hello bold world',
      newString: 'Updated text',
    });
    assert.equal(result.mode, 'rewrite');
    assert.match(readFileSync(file, 'utf8'), /Updated text\./);
  });

  it('edits one EPUB spine chapter and preserves assets', async () => {
    const file = path.join(dir, 'book.epub');
    await makeEpub(file);
    const secondBefore = await entryBytes(file, 'OEBPS/two.xhtml');
    const imageBefore = await entryBytes(file, 'OEBPS/image.bin');

    const result = await editDocument(file, {
      type: 'epub',
      op: 'replace',
      oldString: 'Hello world',
      newString: 'Edited chapter',
      chapter: 1,
    });
    assert.equal(result.mode, 'surgical');
    assert.match(
      (await entryBytes(file, 'OEBPS/one.xhtml')).toString(),
      /Edited chapter/
    );
    assert.deepEqual(await entryBytes(file, 'OEBPS/two.xhtml'), secondBefore);
    assert.deepEqual(await entryBytes(file, 'OEBPS/image.bin'), imageBefore);
  });

  it('replaces RTF text across formatting and Unicode tokens', async () => {
    const file = path.join(dir, 'text.rtf');
    writeFileSync(
      file,
      String.raw`{\rtf1\ansi\ansicpg1252\uc1{\fonttbl{\f0 Arial;}}\pard Hello \b world\b0  \u8212? test\par }`,
      'utf8'
    );
    const result = await editDocument(file, {
      type: 'rtf',
      op: 'replace',
      oldString: 'Hello world — test',
      newString: 'Привет мир',
    });
    assert.equal(result.mode, 'surgical');
    const raw = readFileSync(file, 'utf8');
    assert.match(raw, /\\fonttbl/);
    assert.match(raw, /\\u1055\?/);
    const parsed = await parseDocument(file, { output: 'structured' });
    assert.match(parsed.text, /Привет мир/);
  });

  it('preserves unrelated raw ANSI bytes during RTF edit', async () => {
    const file = path.join(dir, 'ansi.rtf');
    writeFileSync(
      file,
      Buffer.from(
        String.raw`{\rtf1\ansi\ansicpg1252\uc1\pard caf` +
          String.fromCharCode(0xe9) +
          String.raw` Target\par }`,
        'latin1'
      )
    );
    await editDocument(file, {
      type: 'rtf',
      op: 'replace',
      oldString: 'Target',
      newString: 'Done',
    });
    const bytes = readFileSync(file);
    assert.ok(bytes.includes(0xe9));
    assert.equal(bytes.includes(Buffer.from([0xef, 0xbf, 0xbd])), false);
  });

  it('advertises precise edit modes and operations', () => {
    const pptx = getCapabilities('pptx')[0];
    assert.equal(pptx.edit, true);
    assert.deepEqual(pptx.editModes, ['surgical']);
    assert.deepEqual(pptx.editOps, ['replace']);

    const odt = getCapabilities('odt')[0];
    assert.deepEqual(odt.editModes, ['surgical', 'rewrite']);

    const ods = getCapabilities('ods')[0];
    assert.deepEqual(ods.editModes, ['rewrite']);
    assert.deepEqual(ods.editOps, ['patchCells']);

    const pdf = getCapabilities('pdf')[0];
    assert.equal(pdf.edit, false);
    assert.equal(pdf.editModes, undefined);
  });
});
