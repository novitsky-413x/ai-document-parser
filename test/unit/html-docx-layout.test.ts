import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { after, test } from 'node:test';
import JSZip from 'jszip';
import { htmlToBlocks, htmlToDocument } from '../../src/parse/html-blocks.js';
import { writeDocument } from '../../src/write/index.js';
import { parseDocument } from '../../src/index.js';
import type { WriteDocumentInput } from '../../src/types.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const fixture = (name: string) => readFileSync(path.join(here, '..', 'fixtures', name), 'utf8');
const dir = mkdtempSync(path.join(tmpdir(), 'adp-html-layout-'));
after(() => rmSync(dir, { recursive: true, force: true }));

async function zipText(file: string, entry: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(file));
  const text = await zip.file(entry)?.async('text');
  assert.ok(text, `${entry} missing`);
  return text;
}

function paragraphs(documentXml: string): string[] {
  return documentXml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? [];
}

function paraText(p: string): string {
  return (p.match(/<w:t[^>]*>[^<]*<\/w:t>|<w:br\/>/g) ?? [])
    .map((t) => (t === '<w:br/>' ? '\n' : t.replace(/<[^>]+>/g, '')))
    .join('');
}

test('html title page keeps alignment, spacing, sizes and line breaks in DOCX', async () => {
  const file = path.join(dir, 'title-inline.docx');
  const res = await writeDocument(file, {
    type: 'docx',
    format: 'html',
    content: fixture('title-page-inline-style.html'),
  });
  assert.equal(res.warnings, undefined);
  const doc = await zipText(file, 'word/document.xml');
  const ps = paragraphs(doc).filter((p) => paraText(p).trim());
  // <title> from <head> is not body text.
  assert.equal(ps.length, 4);
  assert.ok(!doc.includes('Устав ООО'), 'head <title> leaked into the body');

  const [approval, title, subtitle, place] = ps;
  assert.match(approval, /<w:jc w:val="right"\/>/);
  assert.equal(
    paraText(approval),
    'УТВЕРЖДЁН\nрешением общего собрания учредителей\nОбщества с ограниченной ответственностью\n«Эксперт Плюс Москва»\nПротокол № 1 от 28 апреля 2025 г.'
  );
  assert.match(approval, /<w:sz w:val="24"\/>/);

  assert.match(title, /<w:jc w:val="center"\/>/);
  assert.match(title, /<w:sz w:val="44"\/>/);
  assert.match(title, /<w:spacing w:val="110"\/>/, 'letter-spacing 0.25em of 22pt = 5.5pt');
  assert.match(title, /w:before="4160"/, 'margin-top 220pt minus the previous 12pt margin');
  assert.match(subtitle, /<w:jc w:val="center"\/>/);
  assert.match(subtitle, /<w:b\/>/);
  assert.match(place, /<w:jc w:val="center"\/>/);
  assert.equal(paraText(place), 'д. Городище\n2025 год');

  const styles = await zipText(file, 'word/styles.xml');
  assert.match(styles, /w:ascii="Times New Roman"/);
  assert.match(styles, /<w:sz w:val="28"\/>/);
  assert.match(styles, /<w:color w:val="111111"\/>/);

  // Line breaks survive a read back.
  const parsed = await parseDocument(file);
  assert.match(parsed.text, /д\. Городище\n2025 год/);
  assert.doesNotMatch(parsed.text, /Устав ООО/);
});

test('html stylesheet rules, div containers and @page margins reach DOCX', async () => {
  const file = path.join(dir, 'title-sheet.docx');
  const res = await writeDocument(file, {
    type: 'docx',
    format: 'html',
    content: fixture('title-page-stylesheet.html'),
  });
  assert.equal(res.warnings, undefined);
  assert.deepEqual(res.summary?.blockCounts, { paragraph: 4, heading: 1 });
  const doc = await zipText(file, 'word/document.xml');
  const ps = paragraphs(doc).filter((p) => paraText(p).trim());
  assert.equal(paraText(ps[0]), 'УТВЕРЖДЁН');
  assert.match(ps[0], /<w:jc w:val="right"\/>/, '.approval text-align inherited into the nested div');
  assert.match(ps[1], /<w:jc w:val="right"\/>/);
  assert.match(ps[2], /<w:pStyle w:val="Heading1"\/>/);
  assert.match(ps[2], /<w:jc w:val="center"\/>/);
  assert.match(ps[2], /w:before="4082"/, '.mid margin-top 72mm goes to its first child');
  // @page A4, margin 20mm 20mm 20mm 25mm
  assert.match(doc, /<w:pgSz w:w="11906" w:h="16838"/);
  assert.match(doc, /w:top="1134"/);
  assert.match(doc, /w:left="1417"/);
});

test('html <br>, <head>/<title>, div and section parse into blocks', () => {
  const blocks = htmlToBlocks(
    '<html><head><title>Doc title</title><style>p{color:red}</style></head><body>' +
      '<section><div>first line<br>second   line</div><div><b>bold</b> tail</div></section>' +
      '<p>a <span>b</span>\n c</p></body></html>'
  );
  assert.deepEqual(blocks, [
    { type: 'paragraph', runs: [{ text: 'first line\nsecond line' }] },
    { type: 'paragraph', runs: [{ text: 'bold', bold: true }, { text: ' tail' }] },
    { type: 'paragraph', runs: [{ text: 'a b c' }] },
  ]);
});

test('html parse path stays content-only (no layout props)', () => {
  const blocks = htmlToBlocks(
    '<p style="text-align:center;font-size:20pt;margin-top:3em"><span style="font-family:Arial;color:#c00">x</span></p>'
  );
  assert.deepEqual(blocks, [{ type: 'paragraph', runs: [{ text: 'x' }] }]);
});

test('html layout: class/descendant/child selectors, units and margin collapse', () => {
  const doc = htmlToDocument(
    '<style>body{font:13pt/1.5 Georgia, serif} .box p{text-indent:1cm} .box > p.r{text-align:right} #big{font-size:150%}</style>' +
      '<div class="box"><p>one</p><p class="r" style="margin-top:24pt">two <span id="big">big</span></p></div>' +
      '<p style="page-break-before:always">three</p>'
  );
  assert.deepEqual(doc.layout, { font: 'Georgia', sizePt: 13, lineHeight: 1.5 });
  assert.equal(doc.warnings.length, 0);
  const [one, two, three] = doc.blocks;
  assert.ok(one.type === 'paragraph' && two.type === 'paragraph' && three.type === 'paragraph');
  assert.equal(one.layout?.firstLineIndentPt, 28.35);
  assert.equal(two.layout?.align, 'right');
  // Sibling margins collapse: max(13 after, 24 before) → 24 total gap → 11 before.
  assert.equal(one.layout?.spaceAfterPt, 13);
  assert.equal(two.layout?.spaceBeforePt, 11);
  assert.deepEqual(two.runs[1], { text: 'big', sizePt: 19.5 });
  assert.equal(three.layout?.pageBreakBefore, true);
});

test('HTML_STYLE_IGNORED lists CSS that could not be applied', async () => {
  const file = path.join(dir, 'ignored.docx');
  const res = await writeDocument(file, {
    type: 'docx',
    format: 'html',
    content:
      '<style>p:first-child{color:red} @media print{p{color:blue}} .c{position:absolute;text-align:center}</style>' +
      '<p class="c" style="background:#eee">x</p>',
  });
  const w = res.warnings?.find((x) => x.code === 'HTML_STYLE_IGNORED');
  assert.ok(w, 'expected HTML_STYLE_IGNORED');
  assert.match(w.message, /background/);
  assert.match(w.message, /position/);
  assert.match(w.message, /p:first-child/);
  assert.match(w.message, /@media/);
  assert.doesNotMatch(w.message, /text-align/);
  const doc = await zipText(file, 'word/document.xml');
  assert.match(doc, /<w:jc w:val="center"\/>/);
});

test('HTML_STYLE_IGNORED warns for formats without layout only when CSS was given', async () => {
  const styled = '<p style="text-align:center">x<br>y</p>';
  for (const type of ['rtf', 'epub', 'pptx'] as const) {
    const res = await writeDocument(path.join(dir, `styled.${type}`), {
      type,
      format: 'html',
      content: styled,
    } as WriteDocumentInput);
    assert.ok(
      res.warnings?.some((w) => w.code === 'HTML_STYLE_IGNORED'),
      `${type} should warn`
    );
    const plain = await writeDocument(path.join(dir, `plain.${type}`), {
      type,
      format: 'html',
      content: '<h1>T</h1><p>x</p>',
    } as WriteDocumentInput);
    assert.ok(
      !plain.warnings?.some((w) => w.code === 'HTML_STYLE_IGNORED'),
      `${type} should not warn without CSS`
    );
  }
});

test('ODT write keeps paragraph alignment, run size, line breaks and page setup', async () => {
  const file = path.join(dir, 'title.odt');
  await writeDocument(file, {
    type: 'odt',
    format: 'html',
    content: fixture('title-page-stylesheet.html'),
  });
  const content = await zipText(file, 'content.xml');
  assert.match(content, /fo:text-align="end"/);
  assert.match(content, /fo:text-align="center"/);
  assert.match(content, /fo:font-size="22pt"/);
  assert.match(content, /д\. Городище<text:line-break\/>2025 год/);
  const styles = await zipText(file, 'styles.xml');
  assert.match(styles, /fo:page-width="595.28pt"/);
  assert.match(styles, /fo:margin-left="70.87pt"/);
  assert.match(styles, /fo:font-family="Times New Roman"/);
  const parsed = await parseDocument(file);
  assert.match(parsed.text, /д\. Городище\n2025 год/);
});

test('PPTX write keeps paragraph alignment and line breaks', async () => {
  const file = path.join(dir, 'aligned.pptx');
  await writeDocument(file, {
    type: 'pptx',
    format: 'html',
    content: '<h1>Title</h1><p style="text-align:center">a<br>b</p>',
  });
  const slide = await zipText(file, 'ppt/slides/slide1.xml');
  assert.match(slide, /<a:pPr algn="ctr">/);
  assert.match(slide, /<a:t>a<\/a:t><\/a:r><a:br>/);
});

test('DOCX blocks mode accepts document layout', async () => {
  const file = path.join(dir, 'blocks-layout.docx');
  await writeDocument(file, {
    type: 'docx',
    blocks: [
      {
        type: 'paragraph',
        runs: [{ text: 'Hi', sizePt: 18, font: 'Arial', color: 'FF0000' }],
        layout: { align: 'justify', firstLineIndentPt: -18, indentLeftPt: 36 },
      },
    ],
    layout: { font: 'Georgia', page: { widthPt: 842, heightPt: 595, marginTopPt: 36 } },
  });
  const doc = await zipText(file, 'word/document.xml');
  assert.match(doc, /<w:jc w:val="both"\/>/);
  assert.match(doc, /w:hanging="360"/);
  assert.match(doc, /<w:sz w:val="36"\/>/);
  assert.match(doc, /w:orient="landscape"/);
  assert.match(doc, /<w:pgSz w:w="16840" w:h="11900"/);
});
