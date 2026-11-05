import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import {
  countLetters,
  linesToBlocks,
  PDF_EMPTY_PAGE_LETTERS,
  wordsToLines,
  type PdfWord,
} from '../../src/parse/pdf-structure.js';

function word(
  text: string,
  x: number,
  y: number,
  page: number,
  width?: number,
  height?: number
): PdfWord {
  const h = height == null ? 12 : height;
  return {
    text,
    x,
    y,
    width: width == null ? Math.max(12, text.length * h * 0.45) : width,
    height: h,
    page,
  };
}

describe('pdf structure', () => {
  it('keeps the short hygiene text layer above the scan threshold', () => {
    assert.ok(countLetters('Привет HYGIENE-PDF-OK') >= PDF_EMPTY_PAGE_LETTERS);
  });

  it('splits a heading glued to the first numbered clause', () => {
    const lines = wordsToLines([
      word('Глава', 72, 700, 1),
      word('1.', 120, 700, 1),
      word('Общие', 150, 700, 1),
      word('положения', 210, 700, 1),
      word('1.1.', 300, 700, 1),
      word('Первый', 350, 700, 1),
      word('пункт', 420, 700, 1),
    ]);
    const { blocks } = linesToBlocks(lines);
    assert.equal(blocks.length, 2);
    assert.equal(blocks[0].type, 'heading');
    assert.equal(blocks[1].type, 'paragraph');
    if (blocks[1].type !== 'paragraph') return;
    assert.equal(blocks[1].runs[0].bold, true);
    assert.match(blocks[1].runs.map((run) => run.text).join(''), /1\.1\./);
    assert.match(blocks[1].runs.map((run) => run.text).join(''), /Первый пункт/);
  });

  it('puts each numbered clause in its own paragraph', () => {
    const lines = wordsToLines([
      word('1.1.', 72, 700, 1),
      word('Альфа', 120, 700, 1),
      word('1.2.', 72, 670, 1),
      word('Бета', 120, 670, 1),
    ]);
    const { blocks } = linesToBlocks(lines);
    assert.equal(blocks.length, 2);
    assert.equal(blocks[0].type, 'paragraph');
    assert.equal(blocks[1].type, 'paragraph');
  });

  it('joins a hyphenated word and drops a bare page number', () => {
    const lines = wordsToLines([
      word('28', 300, 800, 1, 16, 10),
      word('докумен-', 72, 700, 1),
      word('том', 72, 680, 1),
    ]);
    const { blocks } = linesToBlocks(lines);
    assert.equal(blocks.length, 1);
    assert.equal(blocks[0].type, 'paragraph');
    if (blocks[0].type !== 'paragraph') return;
    assert.equal(blocks[0].runs[0].text, 'документом');
  });

  it('replaces Latin No with numero on a Cyrillic line', () => {
    const lines = wordsToLines([word('Счет No 12 открыт', 72, 700, 1)]);
    const { blocks } = linesToBlocks(lines);
    assert.equal(blocks[0].type, 'paragraph');
    if (blocks[0].type !== 'paragraph') return;
    assert.ok(blocks[0].runs[0].text.indexOf('\u2116') !== -1);
  });

  it('drops a short header repeated on three pages', () => {
    const lines = wordsToLines([
      word('Банк России', 72, 800, 1),
      word('Текст один', 72, 700, 1),
      word('Банк России', 72, 800, 2),
      word('Текст два', 72, 700, 2),
      word('Банк России', 72, 800, 3),
      word('Текст три', 72, 700, 3),
    ]);
    const { blocks } = linesToBlocks(lines);
    const text = blocks
      .map((block) => (block.type === 'paragraph' ? block.runs.map((run) => run.text).join('') : ''))
      .join('\n');
    assert.equal(text.indexOf('Банк России'), -1);
    assert.match(text, /Текст один/);
    assert.match(text, /Текст три/);
  });

  it('builds a table when several lines share wide column gaps', () => {
    const lines = wordsToLines([
      word('А', 0, 700, 1, 20, 12),
      word('Б', 200, 700, 1, 20, 12),
      word('В', 400, 700, 1, 20, 12),
      word('1', 0, 680, 1, 20, 12),
      word('2', 200, 680, 1, 20, 12),
      word('3', 400, 680, 1, 20, 12),
      word('4', 0, 660, 1, 20, 12),
      word('5', 200, 660, 1, 20, 12),
      word('6', 400, 660, 1, 20, 12),
    ]);
    const { blocks, warnings } = linesToBlocks(lines);
    assert.equal(blocks.length, 1);
    assert.equal(blocks[0].type, 'table');
    if (blocks[0].type !== 'table') return;
    assert.deepEqual(blocks[0].rows[0], ['А', 'Б', 'В']);
    assert.ok(warnings.some((warning) => warning.code === 'PDF_TABLE_HEURISTIC'));
  });

  it('joins wrapped lines on a normal leading and breaks on a larger gap', () => {
    const lines = wordsToLines([
      word('Первая', 72, 700, 1),
      word('строка', 150, 700, 1),
      word('вторая', 72, 686, 1),
      word('строка', 150, 686, 1),
      word('третья', 72, 672, 1),
      word('строка', 150, 672, 1),
      word('Новый', 72, 600, 1),
      word('абзац', 150, 600, 1),
    ]);
    const { blocks } = linesToBlocks(lines);
    assert.equal(blocks.length, 2);
    assert.equal(blocks[0].type, 'paragraph');
    assert.equal(blocks[1].type, 'paragraph');
    if (blocks[0].type !== 'paragraph' || blocks[1].type !== 'paragraph') return;
    const first = blocks[0].runs.map((run) => run.text).join('');
    const second = blocks[1].runs.map((run) => run.text).join('');
    assert.equal(first, 'Первая строка вторая строка третья строка');
    assert.equal(second, 'Новый абзац');
  });

  it('does not emit a lone dash as its own paragraph', () => {
    const lines = wordsToLines([
      word('|', 72, 740, 1, 8, 12),
      word('России,', 72, 700, 1),
      word('организаций', 160, 700, 1),
      word('юридических', 72, 686, 1),
      word('лиц,', 200, 686, 1),
      word('\u2014', 72, 672, 1),
      word('других', 72, 658, 1),
      word('подразделений', 150, 658, 1),
      word('Банка', 72, 644, 1),
      word('России', 150, 644, 1),
    ]);
    const { blocks } = linesToBlocks(lines);
    assert.equal(blocks.length, 2);
    assert.equal(blocks[0].type, 'paragraph');
    assert.equal(blocks[1].type, 'list_item');
    if (blocks[0].type !== 'paragraph' || blocks[1].type !== 'list_item') return;
    const para = blocks[0].runs.map((run) => run.text).join('');
    const item = blocks[1].runs.map((run) => run.text).join('');
    assert.equal(para, 'России, организаций юридических лиц,');
    assert.equal(item, 'других подразделений Банка России');
    assert.equal(para.indexOf('|'), -1);
  });
});
