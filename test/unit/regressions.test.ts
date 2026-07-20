import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { parseHtml } from '../../dist/parse/text-formats.js';
import { markdownToBlocks } from '../../dist/parse/html-blocks.js';
import { emitMarkdownFromBlocks, emitHtmlFromBlocks } from '../../dist/convert/blocks.js';
import { chunkText } from '../../dist/convert/emit.js';
import { detectFormatFromBuffer } from '../../dist/parse/detect.js';
import { contentBlocksToFlow, flowToContentBlocks } from '../../dist/ir/projections.js';
import { getCapabilities } from '../../dist/index.js';

describe('entity decoding', () => {
  it('does not double-decode &amp;lt; in parseHtml', () => {
    const parsed = parseHtml('<p>&amp;lt;b&amp;gt;</p>');
    assert.equal(parsed.text, '&lt;b&gt;');
  });
});

describe('markdown inline code', () => {
  it('parses `code` as code run, not bold+italic', () => {
    const blocks = markdownToBlocks('Use `npm test` now');
    assert.equal(blocks[0].type, 'paragraph');
    const runs = (blocks[0] as { runs: Array<Record<string, unknown>> }).runs;
    const code = runs.find((r) => r.code === true);
    assert.ok(code, 'expected a code run');
    assert.equal(code!.text, 'npm test');
    assert.notEqual(code!.bold, true);
  });

  it('emits code runs back as backticks in md and <code> in html', () => {
    const blocks = markdownToBlocks('Use `npm test` now');
    assert.match(emitMarkdownFromBlocks(blocks), /`npm test`/);
    assert.match(emitHtmlFromBlocks(blocks), /<code>npm test<\/code>/);
  });
});

describe('markdown emit escaping', () => {
  it('escapes block starters at paragraph start', () => {
    const md = emitMarkdownFromBlocks([
      { type: 'paragraph', runs: [{ text: '# not a heading' }] },
      { type: 'paragraph', runs: [{ text: '1. not a list' }] },
    ]);
    assert.match(md, /\\# not a heading/);
    assert.match(md, /1\\. not a list/);
  });

  it('grows code fence when body contains backticks', () => {
    const md = emitMarkdownFromBlocks([
      { type: 'code', text: 'fence ```inside``` body' },
    ]);
    assert.match(md, /^````/m);
  });
});

describe('nested lists', () => {
  it('parses md list levels from indentation', () => {
    const blocks = markdownToBlocks('- A\n  - B\n- C');
    const items = blocks.filter((b) => b.type === 'list_item') as Array<{ level: number }>;
    assert.deepEqual(items.map((i) => i.level), [0, 1, 0]);
  });

  it('emits nested html lists inside the parent <li>', () => {
    const html = emitHtmlFromBlocks([
      { type: 'list_item', level: 0, runs: [{ text: 'A' }] },
      { type: 'list_item', level: 1, runs: [{ text: 'B' }] },
      { type: 'list_item', level: 0, runs: [{ text: 'C' }] },
    ]);
    assert.match(html, /<li>A\s*<ul>\s*<li>B\s*<\/li>\s*<\/ul>\s*<\/li>/);
    assert.match(html, /<li>C\s*<\/li>/);
  });

  it('round-trips list levels through the canonical flow projection', () => {
    const original = [
      { type: 'list_item' as const, level: 0, runs: [{ text: 'A', bold: true }] },
      { type: 'list_item' as const, level: 1, runs: [{ text: 'B' }] },
      { type: 'list_item' as const, level: 0, runs: [{ text: 'C' }] },
    ];
    const back = flowToContentBlocks(contentBlocksToFlow(original)).filter(
      (b) => b.type === 'list_item'
    ) as Array<{ level: number; runs: Array<{ text: string; bold?: boolean }> }>;
    assert.deepEqual(back.map((i) => i.level), [0, 1, 0]);
    assert.deepEqual(back.map((i) => i.runs[0].text), ['A', 'B', 'C']);
    assert.equal(back[0].runs[0].bold, true, 'run formatting survives projection');
  });
});

describe('chunkText', () => {
  const long = Array.from({ length: 50 }, (_, i) => `word${i}`).join(' ');

  it('terminates when overlap >= maxChars', () => {
    const chunks = chunkText(long, undefined, 50, 50);
    assert.ok(chunks.length > 0 && chunks.length < 1000);
  });

  it('clamps invalid params', () => {
    const chunks = chunkText(long, undefined, 0, -5);
    assert.ok(chunks.length > 0);
  });
});

describe('CFB detection', () => {
  const magic = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);

  it('detects xls via Workbook stream', () => {
    const buf = Buffer.concat([magic, Buffer.alloc(64), Buffer.from('Workbook', 'utf16le')]);
    assert.equal(detectFormatFromBuffer(buf), 'xls');
  });

  it('detects doc via WordDocument stream', () => {
    const buf = Buffer.concat([magic, Buffer.alloc(64), Buffer.from('WordDocument', 'utf16le')]);
    assert.equal(detectFormatFromBuffer(buf), 'doc');
  });
});

describe('BOM handling', () => {
  it('detects rtf behind a UTF-8 BOM', () => {
    const buf = Buffer.from('\uFEFF{\\rtf1\\ansi Hi\\par}', 'utf8');
    assert.equal(detectFormatFromBuffer(buf), 'rtf');
  });
});

describe('capabilities registry', () => {
  it('has no duplicate formats', () => {
    const names = getCapabilities().map((c) => c.format);
    assert.equal(new Set(names).size, names.length);
  });

  it('does not advertise phantom workbook/presentation binary targets', () => {
    const byFormat = new Map(getCapabilities().map((c) => [c.format, c.convertTo]));
    assert.ok(!byFormat.get('xlsx')!.includes('ods'));
    assert.ok(!byFormat.get('ods')!.includes('xlsx'));
    assert.ok(!byFormat.get('pptx')!.includes('odp'));
    assert.ok(!byFormat.get('odp')!.includes('pptx'));
  });
});
