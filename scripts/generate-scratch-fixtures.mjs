/**
 * Generate scratch/ smoke fixtures for manual + agent verification.
 * Run: node scripts/generate-scratch-fixtures.mjs
 */
import { copyFileSync, mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  AlignmentType,
  Document,
  HeadingLevel,
  ImageRun,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
  BorderStyle,
} from 'docx';
import * as XLSX from '@e965/xlsx';
import { jsPDF } from 'jspdf';
import JSZip from 'jszip';
import * as CFB from 'cfb';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(__dirname, '..');
const outDir = path.join(root, 'scratch');

mkdirSync(outDir, { recursive: true });
mkdirSync(path.join(outDir, 'out'), { recursive: true });

/** Valid 1×1 PNG (scaled in DOCX). */
function tinyPng() {
  return Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64'
  );
}

async function writeReportDocx(filePath) {
  const png = tinyPng();
  const border = {
    style: BorderStyle.SINGLE,
    size: 8,
    color: '999999',
  };
  const borders = { top: border, bottom: border, left: border, right: border };

  const table = new Table({
    width: { size: 9000, type: WidthType.DXA },
    rows: [
      new TableRow({
        children: ['Показатель', 'План', 'Факт'].map(
          (t) =>
            new TableCell({
              borders,
              width: { size: 3000, type: WidthType.DXA },
              children: [
                new Paragraph({
                  children: [new TextRun({ text: t, bold: true })],
                }),
              ],
            })
        ),
      }),
      ...[
        ['Выручка', '1000', '950'],
        ['Расходы', '400', '420'],
      ].map(
        (row) =>
          new TableRow({
            children: row.map(
              (t) =>
                new TableCell({
                  borders,
                  width: { size: 3000, type: WidthType.DXA },
                  children: [new Paragraph({ children: [new TextRun(t)] })],
                })
            ),
          })
      ),
    ],
  });

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            heading: HeadingLevel.HEADING_1,
            children: [new TextRun('Квартальный отчёт / Quarterly Report')],
          }),
          new Paragraph({
            children: [
              new TextRun('Контрагент: Acme Corp'),
              new TextRun(' — уникальная фраза для replace (ровно один раз).'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun('Договор подписан стороной '),
              new TextRun({ text: 'ООО Ромашка', bold: true }),
              new TextRun(' (первое вхождение).'),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({ text: 'ООО Ромашка', italics: true }),
              new TextRun(
                ' подтвердила поставку (второе вхождение — для MULTIPLE_MATCHES / --all).'
              ),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun(
                'English paragraph: The budget was approved on 2026-03-15. Cyrillic mix: Москва / Moscow.'
              ),
            ],
          }),
          new Paragraph({
            children: [new TextRun({ text: 'Ключевые пункты:', bold: true })],
          }),
          new Paragraph({
            bullet: { level: 0 },
            children: [new TextRun('Проверить structured.paragraphs перед edit')],
          }),
          new Paragraph({
            bullet: { level: 0 },
            children: [new TextRun('Не ломать таблицу и картинку при replace')],
          }),
          new Paragraph({
            children: [new TextRun('Таблица ниже (3×3):')],
          }),
          table,
          new Paragraph({
            children: [new TextRun('Встроенная картинка (PNG):')],
          }),
          new Paragraph({
            alignment: AlignmentType.LEFT,
            children: [
              new ImageRun({
                type: 'png',
                data: png,
                transformation: { width: 64, height: 64 },
                altText: {
                  title: 'fixture',
                  description: 'teal square smoke fixture',
                  name: 'fixture.png',
                },
              }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun(
                'Конец документа. После smoke-edit этот абзац должен остаться читаемым.'
              ),
            ],
          }),
        ],
      },
    ],
  });

  const buf = await Packer.toBuffer(doc);
  writeFileSync(filePath, buf);
}

function writeBudgetXlsx(filePath, bookType = 'xlsx') {
  const wb = XLSX.utils.book_new();

  const wsQ1 = XLSX.utils.aoa_to_sheet([
    ['Revenue', 'План', 'Факт', 'Сумма'],
    [1000, 1200, 950, null],
    ['Операционные расходы', 400, 420, null],
    ['Дата отчёта', '2026-03-31', '', ''],
    ['', '', '', ''],
    ['Примечание', 'Кириллица и numbers в одной книге', '', ''],
  ]);
  wsQ1['D2'] = { t: 'n', f: 'B2+C2' };
  wsQ1['D3'] = { t: 'n', f: 'B3+C3' };
  wsQ1['!ref'] = 'A1:D6';
  XLSX.utils.book_append_sheet(wb, wsQ1, 'Q1');

  const summary = XLSX.utils.aoa_to_sheet([
    ['Лист', 'Значение'],
    ['Сводка', 'Итого по Q1'],
    ['Статус', 'черновик'],
    ['Ответственный', 'Иванов И.И.'],
  ]);
  XLSX.utils.book_append_sheet(wb, summary, 'Сводка');

  const data = XLSX.write(wb, { type: 'buffer', bookType });
  writeFileSync(filePath, data);
}

function registerFont(doc) {
  const fontFile = 'DejaVuSans.ttf';
  const fontPath = path.join(root, 'fonts', fontFile);
  const fontData = readFileSync(fontPath).toString('binary');
  doc.addFileToVFS(fontFile, fontData);
  doc.addFont(fontFile, 'DejaVuSans', 'normal');
  doc.setFont('DejaVuSans');
}

function writeSamplePdf(filePath) {
  const doc = new jsPDF({ unit: 'mm', format: 'a4' });
  registerFont(doc);
  const margin = 15;
  const width = doc.internal.pageSize.getWidth() - margin * 2;

  doc.setFontSize(14);
  doc.text('Sample PDF — страница 1', margin, margin);
  doc.setFontSize(11);
  const p1 = doc.splitTextToSize(
    'Это текстовый PDF с кириллицей и Latin text. Страница одна из двух. ' +
      'Используется для parseDocument / convertDocument (text, md, html, chunks). ' +
      'Edit для PDF не поддерживается.',
    width
  );
  doc.text(p1, margin, margin + 12);

  doc.addPage();
  registerFont(doc);
  doc.setFontSize(14);
  doc.text('Sample PDF — страница 2', margin, margin);
  doc.setFontSize(11);
  const p2 = doc.splitTextToSize(
    'Second page. Markers: BUDGET_OK, Москва, Acme Corp (reference only). ' +
      'Table-like lines:\nItem A — 10\nItem B — 20\nItem C — 30',
    width
  );
  doc.text(p2, margin, margin + 12);

  writeFileSync(filePath, Buffer.from(doc.output('arraybuffer')));
}

function writeNotesMd(filePath) {
  writeFileSync(
    filePath,
    `# Заголовок notes.md

Абзац один с **bold** и обычным текстом.

## Секция

- пункт A
- пункт B

Параграф с \`кодом\` и кириллицей: проверка parse/convert.

English line for mixed-language smoke.
`,
    'utf8'
  );
}

function writePageHtml(filePath) {
  writeFileSync(
    filePath,
    `<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="utf-8" />
  <title>Smoke HTML</title>
</head>
<body>
  <h1>Страница HTML</h1>
  <p>Первый абзац с кириллицей и Latin.</p>
  <p>Второй абзац для structured.paragraphs.</p>
  <ul>
    <li>Элемент списка 1</li>
    <li>Элемент списка 2</li>
  </ul>
  <table>
    <tr><th>Колонка</th><th>Значение</th></tr>
    <tr><td>Город</td><td>Москва</td></tr>
    <tr><td>Status</td><td>ok</td></tr>
  </table>
</body>
</html>
`,
    'utf8'
  );
}

function writeDataCsv(filePath) {
  writeFileSync(
    filePath,
    `id,city,note,amount
1,"Москва, Россия","обычная строка",100
2,Berlin,"field with ""quotes""",200
3,Tokyo,"line1
line2 inside quotes",300
4,Paris,simple,400
5,Lisbon,кириллица ok,500
`,
    'utf8'
  );
}

/** Minimal OLE CFB with WordDocument stream (UTF-16LE text) for heuristic .doc parse. */
function writeLegacyDoc(filePath) {
  const text =
    'Legacy DOC smoke fixture.\n' +
    'Кириллица: проверка heuristic extract.\n' +
    'Acme Corp mentioned once.\n' +
    'End of WordDocument stream.';
  const utf16 = Buffer.alloc(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    utf16.writeUInt16LE(text.charCodeAt(i), i * 2);
  }

  const cfb = CFB.utils.cfb_new();
  CFB.utils.cfb_add(cfb, 'WordDocument', utf16);
  const out = CFB.write(cfb, { type: 'buffer' });
  writeFileSync(filePath, Buffer.from(out));
}

async function writeSamplePptx(filePath) {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`
  );
  zip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst/>
</p:presentation>`
  );
  zip.file(
    'ppt/slides/slide1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <a:p><a:r><a:t>PPTX Smoke Title</a:t></a:r></a:p>
    <a:p><a:r><a:t>Кириллица на слайде 1</a:t></a:r></a:p>
  </p:spTree></p:cSld>
</p:sld>`
  );
  zip.file(
    'ppt/slides/slide2.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <a:p><a:r><a:t>Slide two body for convert</a:t></a:r></a:p>
  </p:spTree></p:cSld>
</p:sld>`
  );
  writeFileSync(filePath, await zip.generateAsync({ type: 'nodebuffer' }));
}

async function writeSampleOdt(filePath) {
  const zip = new JSZip();
  zip.file(
    'mimetype',
    'application/vnd.oasis.opendocument.text',
    { compression: 'STORE' }
  );
  zip.file(
    'content.xml',
    `<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:h text:outline-level="1">ODT Sample</text:h><text:p>Абзац на русском</text:p><text:p>English paragraph</text:p></office:text></office:body></office:document-content>`
  );
  writeFileSync(filePath, await zip.generateAsync({ type: 'nodebuffer' }));
}

function writeUnsupportedPpt(filePath) {
  // Legacy .ppt is intentionally unsupported; any bytes behind the extension do.
  writeFileSync(filePath, Buffer.from('not a real ppt'));
}

function writeReadme(filePath) {
  writeFileSync(
    filePath,
    `# scratch fixtures

Generated by \`node scripts/generate-scratch-fixtures.mjs\`.

| File | Purpose |
|------|---------|
| report.docx | Rich DOCX: RU/EN, unique + duplicate phrases, table, bullets, PNG |
| report-copy.docx | Working copy for in-place edit |
| budget.xlsx | Sheets Q1 + Сводка, numbers, formulas D2/D3 |
| legacy.xls | Same idea, Excel 97–2003 |
| sample.pptx | 2 slides, RU/EN text (parse/convert only) |
| sample.pdf | 2 pages, Cyrillic+Latin text layer |
| notes.md / page.html / data.csv | Text-format parse/convert |
| legacy.doc | Minimal CFB WordDocument for heuristic parse |
| sample.odt | Minimal ODT: heading + RU/EN paragraphs |
| unsupported.ppt | Expect UNSUPPORTED_FORMAT (legacy .ppt) |
| out/ | Empty folder for convert/write outputs |

Open report.docx / budget.xlsx / sample.pptx / sample.pdf and confirm they look sane before the smoke run.
`,
    'utf8'
  );
}

const reportPath = path.join(outDir, 'report.docx');
await writeReportDocx(reportPath);
copyFileSync(reportPath, path.join(outDir, 'report-copy.docx'));
writeBudgetXlsx(path.join(outDir, 'budget.xlsx'), 'xlsx');
writeBudgetXlsx(path.join(outDir, 'legacy.xls'), 'xls');
await writeSamplePptx(path.join(outDir, 'sample.pptx'));
writeSamplePdf(path.join(outDir, 'sample.pdf'));
writeNotesMd(path.join(outDir, 'notes.md'));
writePageHtml(path.join(outDir, 'page.html'));
writeDataCsv(path.join(outDir, 'data.csv'));
writeLegacyDoc(path.join(outDir, 'legacy.doc'));
await writeSampleOdt(path.join(outDir, 'sample.odt'));
writeUnsupportedPpt(path.join(outDir, 'unsupported.ppt'));
writeReadme(path.join(outDir, 'README.md'));

console.log('Generated fixtures in', outDir);
for (const name of [
  'report.docx',
  'report-copy.docx',
  'budget.xlsx',
  'legacy.xls',
  'sample.pptx',
  'sample.pdf',
  'notes.md',
  'page.html',
  'data.csv',
  'legacy.doc',
  'sample.odt',
  'unsupported.ppt',
  'README.md',
]) {
  const p = path.join(outDir, name);
  console.log(`  ${name.padEnd(22)} ${statSync(p).size} bytes`);
}
