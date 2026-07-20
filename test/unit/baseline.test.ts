import { describe, it } from 'node:test';
import assert from 'node:assert/strict';
import { parseCsv } from '../../dist/parse/text-formats.js';

describe('parseCsv multiline', () => {
  it('preserves quoted newlines', () => {
    const csv = `name,note
"Line1
Line2",ok`;
    const parsed = parseCsv(csv);
    assert.equal(parsed.rows.length, 2);
    assert.equal(parsed.rows[1][0], 'Line1\nLine2');
    assert.equal(parsed.rows[1][1], 'ok');
  });
});

describe('detect raw markdown string', () => {
  it('treats single-line heading as content', async () => {
    const { parseDocument } = await import('../../dist/index.js');
    const result = await parseDocument('# Title only');
    assert.match(result.text, /Title only/);
  });
});
