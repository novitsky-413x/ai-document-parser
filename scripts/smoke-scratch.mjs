/**
 * Full smoke against scratch/ fixtures.
 * Run: node scripts/smoke-scratch.mjs
 */
import { copyFileSync, mkdirSync, readFileSync, existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import {
  parseDocument,
  convertDocument,
  editDocument,
  writeDocument,
  convertDocToDocx,
  DocumentError,
} from '../dist/index.js';

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const scratch = path.join(root, 'scratch');
const out = path.join(scratch, 'out');
mkdirSync(out, { recursive: true });

const results = [];

function ok(name, detail = '') {
  results.push({ status: 'PASS', name, detail });
  console.log(`PASS  ${name}${detail ? ' — ' + detail : ''}`);
}

function fail(name, err) {
  const detail = err instanceof Error ? `${err.message}${err.code ? ` [${err.code}]` : ''}` : String(err);
  results.push({ status: 'FAIL', name, detail });
  console.error(`FAIL  ${name} — ${detail}`);
}

function expect(name, cond, detail = '') {
  if (cond) ok(name, detail);
  else fail(name, new Error(detail || 'assertion failed'));
}

function cli(...args) {
  const r = spawnSync(process.execPath, [path.join(root, 'dist', 'cli.js'), ...args], {
    cwd: root,
    encoding: 'utf8',
    maxBuffer: 10 * 1024 * 1024,
  });
  return r;
}

function cliOk(name, args, check) {
  const r = cli(...args);
  if (r.status !== 0) {
    fail(name, new Error((r.stderr || r.stdout || 'exit ' + r.status).trim()));
    return null;
  }
  try {
    check?.(r.stdout || '');
    ok(name);
  } catch (e) {
    fail(name, e);
  }
  return r;
}

console.log('\n=== CLI read ===\n');

cliOk('cli read report.docx text', ['read', 'scratch/report.docx', '--output', 'text'], (s) => {
  if (!s.includes('Acme Corp')) throw new Error('missing Acme Corp');
  if (!s.includes('ООО Ромашка')) throw new Error('missing ООО Ромашка');
});

cliOk('cli read report.docx structured', ['read', 'scratch/report.docx', '--output', 'structured'], (s) => {
  const j = JSON.parse(s);
  if (!Array.isArray(j.paragraphs) || j.paragraphs.length < 5) {
    throw new Error('expected paragraphs array');
  }
});

cliOk('cli read budget.xlsx structured', ['read', 'scratch/budget.xlsx', '--output', 'structured'], (s) => {
  const j = JSON.parse(s);
  if (!j.sheets?.Q1) throw new Error('missing Q1 sheet');
});

cliOk('cli read sample.pdf text', ['read', 'scratch/sample.pdf', '--output', 'text'], (s) => {
  if (!/кириллиц/i.test(s) && !s.includes('страница')) throw new Error('cyrillic/page missing');
});

cliOk('cli read notes.md', ['read', 'scratch/notes.md'], (s) => {
  if (!s.includes('Заголовок')) throw new Error('md content missing');
});

cliOk('cli read page.html', ['read', 'scratch/page.html'], (s) => {
  if (!s.includes('Страница HTML') && !s.includes('абзац')) throw new Error('html content missing');
});

cliOk('cli read data.csv structured', ['read', 'scratch/data.csv', '--output', 'structured'], (s) => {
  const j = JSON.parse(s);
  if (!j.rows || j.rows.length < 3) throw new Error('csv rows missing');
});

cliOk('cli read legacy.doc', ['read', 'scratch/legacy.doc'], (s) => {
  if (!s.includes('Legacy DOC')) throw new Error('legacy text missing');
});

cliOk('cli read sample.pptx structured', ['read', 'scratch/sample.pptx', '--output', 'structured'], (s) => {
  const j = JSON.parse(s);
  if (!Array.isArray(j.slides) || j.slides.length < 2) {
    throw new Error('expected pptx slides');
  }
});

{
  const r = cli('read', 'scratch/unsupported.odt');
  expect(
    'cli read unsupported.odt rejects',
    r.status !== 0 && /UNSUPPORTED|Unsupported/i.test(r.stderr + r.stdout),
    (r.stderr || r.stdout).trim().slice(0, 80)
  );
}

console.log('\n=== CLI convert ===\n');

for (const [file, to, outName, needle] of [
  ['scratch/report.docx', 'md', 'report.md', 'Acme'],
  ['scratch/report.docx', 'html', 'report.html', 'Acme'],
  ['scratch/budget.xlsx', 'csv', 'budget.csv', 'Revenue'],
  ['scratch/budget.xlsx', 'md', 'budget.md', 'Q1'],
  ['scratch/sample.pptx', 'md', 'sample-pptx.md', 'Slide 1'],
  ['scratch/sample.pdf', 'md', 'sample.md', 'PDF'],
  ['scratch/notes.md', 'html', 'notes.html', 'Заголовок'],
  ['scratch/data.csv', 'md', 'data.md', 'Москва'],
]) {
  const dest = path.join('scratch', 'out', outName);
  cliOk(`cli convert ${path.basename(file)} → ${to}`, ['convert', file, '--to', to, '-o', dest], () => {
    const body = readFileSync(path.join(root, dest), 'utf8');
    if (!body.includes(needle)) throw new Error(`output missing "${needle}"`);
  });
}

console.log('\n=== CLI edit (docx) ===\n');

copyFileSync(path.join(scratch, 'report.docx'), path.join(out, 'edit-unique.docx'));
cliOk(
  'cli edit unique replace',
  ['edit', 'scratch/out/edit-unique.docx', '--replace', 'Acme Corp', 'Beta LLC'],
  (s) => {
    if (!/replacement/i.test(s) && !/updated/i.test(s)) throw new Error('unexpected message: ' + s);
  }
);
{
  const r = await parseDocument(path.join(out, 'edit-unique.docx'), { output: 'structured' });
  const joined = r.structured.paragraphs.join('\n');
  expect('cli edit unique verified', joined.includes('Beta LLC') && !joined.includes('Acme Corp'));
}

copyFileSync(path.join(scratch, 'report.docx'), path.join(out, 'edit-multi.docx'));
{
  const r = cli('edit', 'scratch/out/edit-multi.docx', '--replace', 'ООО Ромашка', 'ООО Василёк');
  expect(
    'cli edit multi without --all → MULTIPLE_MATCHES',
    r.status !== 0 && /MULTIPLE|multiple|matches/i.test(r.stderr + r.stdout),
    (r.stderr || r.stdout).trim().slice(0, 100)
  );
}
cliOk(
  'cli edit multi with --all',
  ['edit', 'scratch/out/edit-multi.docx', '--replace', 'ООО Ромашка', 'ООО Василёк', '--all'],
  () => {}
);
{
  const r = await parseDocument(path.join(out, 'edit-multi.docx'), { output: 'structured' });
  const joined = r.structured.paragraphs.join('\n');
  const count = (joined.match(/ООО Василёк/g) || []).length;
  expect('cli edit --all verified', count >= 2 && !joined.includes('ООО Ромашка'), `count=${count}`);
}

console.log('\n=== API structured → edit → verify ===\n');

try {
  copyFileSync(path.join(scratch, 'report.docx'), path.join(out, 'api-replace.docx'));
  const { structured } = await parseDocument(path.join(out, 'api-replace.docx'), {
    output: 'structured',
  });
  const target = structured.paragraphs.find((p) => p.includes('Acme Corp'));
  if (!target) throw new Error('no paragraph with Acme Corp');
  await editDocument(path.join(out, 'api-replace.docx'), {
    type: 'docx',
    op: 'replace',
    oldString: target,
    newString: target.replace('Acme Corp', 'Gamma Inc'),
  });
  const again = await parseDocument(path.join(out, 'api-replace.docx'), { output: 'structured' });
  const joined = again.structured.paragraphs.join('\n');
  expect(
    'api docx replace from structured',
    joined.includes('Gamma Inc') && !joined.includes('Acme Corp'),
    `paragraphs=${again.structured.paragraphs.length}`
  );
} catch (e) {
  fail('api docx replace from structured', e);
}

try {
  copyFileSync(path.join(scratch, 'report.docx'), path.join(out, 'api-append.docx'));
  await editDocument(path.join(out, 'api-append.docx'), {
    type: 'docx',
    op: 'append',
    paragraphs: ['APPENDED_SMOKE_LINE'],
  });
  const r = await parseDocument(path.join(out, 'api-append.docx'), { output: 'structured' });
  expect(
    'api docx append',
    r.structured.paragraphs.some((p) => p.includes('APPENDED_SMOKE_LINE'))
  );
} catch (e) {
  fail('api docx append', e);
}

try {
  copyFileSync(path.join(scratch, 'budget.xlsx'), path.join(out, 'api-patch.xlsx'));
  await editDocument(path.join(out, 'api-patch.xlsx'), {
    type: 'spreadsheet',
    op: 'patchCells',
    updates: [
      { sheet: 'Q1', cell: 'B2', value: 1500 },
      { sheet: 'Сводка', cell: 'B3', value: 'утверждён' },
    ],
  });
  const r = await parseDocument(path.join(out, 'api-patch.xlsx'), { output: 'structured' });
  const q1 = r.structured.sheets.Q1;
  const sv = r.structured.sheets['Сводка'];
  // B2 is row1 col1 (0-indexed: [1][1])
  const b2 = q1?.[1]?.[1];
  const status = sv?.[2]?.[1];
  expect(
    'api spreadsheet patchCells',
    Number(b2) === 1500 && String(status).includes('утвержд'),
    `B2=${b2} status=${status}`
  );
} catch (e) {
  fail('api spreadsheet patchCells', e);
}

console.log('\n=== API write → parse ===\n');

try {
  await writeDocument(path.join(out, 'written.docx'), {
    type: 'docx',
    content: ['Строка один', 'Line two'],
  });
  const r = await parseDocument(path.join(out, 'written.docx'), { output: 'structured' });
  expect(
    'write+parse docx',
    r.structured.paragraphs.includes('Строка один') &&
      r.structured.paragraphs.includes('Line two')
  );
} catch (e) {
  fail('write+parse docx', e);
}

try {
  await writeDocument(path.join(out, 'written.xlsx'), {
    type: 'xlsx',
    sheets: { Sheet1: [['A', 'B'], [1, 2]] },
  });
  const r = await parseDocument(path.join(out, 'written.xlsx'), { output: 'structured' });
  expect('write+parse xlsx', r.structured.sheets.Sheet1?.[0]?.[0] === 'A');
} catch (e) {
  fail('write+parse xlsx', e);
}

try {
  await writeDocument(path.join(out, 'written.pdf'), {
    type: 'pdf',
    content: 'Созданный PDF smoke\nSecond line Latin',
    title: 'Smoke PDF',
  });
  const r = await parseDocument(path.join(out, 'written.pdf'), { output: 'text' });
  expect(
    'write+parse pdf',
    r.text.includes('Созданный') || r.text.includes('Smoke'),
    r.text.slice(0, 60)
  );
} catch (e) {
  fail('write+parse pdf', e);
}

console.log('\n=== convertDocToDocx + chunks/ast ===\n');

try {
  const dest = path.join(out, 'from-legacy.docx');
  await convertDocToDocx(path.join(scratch, 'legacy.doc'), dest);
  expect('convertDocToDocx file exists', existsSync(dest));
  const r = await parseDocument(dest, { output: 'text' });
  expect('convertDocToDocx readable', r.text.includes('Legacy DOC'), r.text.slice(0, 80));
} catch (e) {
  fail('convertDocToDocx', e);
}

try {
  const r = await parseDocument(path.join(scratch, 'report.docx'), { output: 'chunks' });
  expect('parse chunks', Array.isArray(r.chunks) && r.chunks.length > 0, `n=${r.chunks?.length}`);
} catch (e) {
  fail('parse chunks', e);
}

try {
  const r = await parseDocument(path.join(scratch, 'report.docx'), { output: 'ast' });
  expect('parse ast', r.ast && typeof r.ast === 'object');
} catch (e) {
  fail('parse ast', e);
}

try {
  await convertDocument(path.join(scratch, 'report.docx'), 'pdf');
  fail('convert to pdf should reject', new Error('did not throw'));
} catch (e) {
  expect(
    'convert to pdf → UNSUPPORTED',
    e instanceof DocumentError && e.code === 'UNSUPPORTED_FORMAT',
    e.message?.slice(0, 80)
  );
}

try {
  const r = await parseDocument(path.join(scratch, 'legacy.xls'), { output: 'structured' });
  expect('parse legacy.xls', r.format === 'xls' && !!r.structured.sheets.Q1);
} catch (e) {
  fail('parse legacy.xls', e);
}

// Image/table survival after edit: re-read paragraphs still include table cells
try {
  const r = await parseDocument(path.join(out, 'edit-unique.docx'), { output: 'structured' });
  const joined = r.structured.paragraphs.join(' | ');
  expect(
    'edit preserves table cell text',
    joined.includes('Выручка') && joined.includes('Показатель'),
    'table headers present after replace'
  );
} catch (e) {
  fail('edit preserves table cell text', e);
}

console.log('\n=== SUMMARY ===\n');
const passed = results.filter((x) => x.status === 'PASS').length;
const failed = results.filter((x) => x.status === 'FAIL');
console.log(`${passed} passed, ${failed.length} failed, ${results.length} total`);
if (failed.length) {
  console.log('\nFailures:');
  for (const f of failed) console.log(`  - ${f.name}: ${f.detail}`);
  process.exit(1);
}
console.log('\nOpen for visual check:');
console.log('  scratch/out/edit-unique.docx   (Acme → Beta)');
console.log('  scratch/out/edit-multi.docx    (Ромашка → Василёк ×2)');
console.log('  scratch/out/api-patch.xlsx     (B2=1500, статус)');
console.log('  scratch/out/written.docx|xlsx|pdf');
process.exit(0);
