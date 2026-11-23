import assert from 'node:assert/strict';
import { test } from 'node:test';
import { htmlToBlocks } from '../../src/parse/html-blocks.js';
import { writeDocxFromBlocks } from '../../src/write/docx-structured.js';
import { contentBlocksToFlow } from '../../src/ir/projections.js';

test('html table cells keep <br> and block breaks as newlines', () => {
  const blocks = htmlToBlocks(
    '<table><tr><td>УТВЕРЖДЁН<br>решением собрания</td><td><p>строка 1</p><p>строка 2</p></td></tr></table>'
  );
  const table = blocks.find((b) => b.type === 'table');
  assert.ok(table && table.type === 'table');
  assert.equal(table.rows[0][0], 'УТВЕРЖДЁН\nрешением собрания');
  assert.equal(table.rows[0][1], 'строка 1\n\nстрока 2');
});

test('write_docx emits a separate paragraph per line inside a table cell', async () => {
  const blocks = htmlToBlocks(
    '<table><tr><td>УТВЕРЖДЁН<br/>решением собрания</td></tr></table>'
  );
  const buf = await writeDocxFromBlocks(contentBlocksToFlow(blocks));
  const zip = await import('jszip').then((m) => m.default.loadAsync(buf));
  const doc = await zip.file('word/document.xml')?.async('text');
  assert.ok(doc, 'document.xml missing');
  const cell = /<w:tc>[\s\S]*?<\/w:tc>/.exec(doc)?.[0] ?? '';
  const paragraphs = cell.match(/<w:p[\s>]/g) ?? [];
  assert.ok(paragraphs.length >= 2, `expected >=2 cell paragraphs, got ${paragraphs.length}: ${cell}`);
  assert.match(cell, /УТВЕРЖДЁН/);
  assert.match(cell, /решением собрания/);
  // A paragraph boundary sits between the two lines (not one collapsed run).
  assert.match(cell, /УТВЕРЖДЁН[\s\S]*<\/w:p>[\s\S]*решением собрания/);
});
