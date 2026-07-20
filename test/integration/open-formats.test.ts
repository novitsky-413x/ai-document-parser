import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { after, before, describe, it } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { parseDocument, convertDocument } from '../../dist/index.js';

async function writeMinimalOdt(filePath: string, paragraphs: string[]): Promise<void> {
  const body = paragraphs
    .map((p) => `<text:p>${p.replace(/&/g, '&amp;')}</text:p>`)
    .join('');
  const zip = new JSZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' });
  zip.file(
    'content.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body><office:text>${body}</office:text></office:body>
</office:document-content>`
  );
  writeFileSync(filePath, await zip.generateAsync({ type: 'nodebuffer' }));
}

describe('open formats', () => {
  let dir: string;

  before(() => {
    dir = mkdtempSync(path.join(tmpdir(), 'adp-open-'));
  });

  after(() => {
    rmSync(dir, { recursive: true, force: true });
  });

  it('parses ODT and converts to md', async () => {
    const file = path.join(dir, 'sample.odt');
    await writeMinimalOdt(file, ['Привет ODT', 'Second paragraph']);
    const parsed = await parseDocument(file, { output: 'structured' });
    assert.equal(parsed.format, 'odt');
    const md = await convertDocument(file, 'md');
    assert.match(String(md.value), /Привет ODT/);
  });

  it('parses RTF and converts to md', async () => {
    const file = path.join(dir, 'sample.rtf');
    writeFileSync(file, '{\\rtf1\\ansi Hello RTF\\par World\\par}', 'utf8');
    const parsed = await parseDocument(file, { output: 'text' });
    assert.match(parsed.text, /Hello RTF/);
    const md = await convertDocument(file, 'md');
    assert.match(String(md.value), /World/);
  });
});
