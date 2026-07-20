import { readFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import * as XLSX from '@e965/xlsx';
import JSZip from 'jszip';
import { DocumentError } from '../utils/errors.js';
import { assertMaxSpreadsheetCells } from '../utils/limits.js';
import { ensureParentDir, writeFileAtomic } from '../utils/paths.js';
import { csvEscapeCell } from '../utils/text.js';
import { decodeXmlEntities, escapeXml } from '../utils/xml.js';
import type { SpreadsheetStructured } from '../types.js';

export type SheetGrid = Record<string, (string | number | boolean | null)[][]>;

export type CellUpdate = {
  sheet?: string;
  cell: string;
  value: string | number | boolean | null;
};

function bookTypeFor(filePath: string): 'xlsx' | 'biff8' | 'ods' {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.xls') return 'biff8';
  if (ext === '.ods') return 'ods';
  return 'xlsx';
}

function formatFor(filePath: string): 'xlsx' | 'xls' | 'ods' {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.xls') return 'xls';
  if (ext === '.ods') return 'ods';
  return 'xlsx';
}

function coerceValue(v: unknown): string | number | boolean | null {
  if (v === null || v === undefined) return null;
  if (typeof v === 'boolean' || typeof v === 'number') return v;
  const s = String(v);
  if (s === 'true') return true;
  if (s === 'false') return false;
  const n = Number(s);
  if (s !== '' && !Number.isNaN(n) && String(n) === s) return n;
  return s;
}

function cellObject(value: string | number | boolean | null): XLSX.CellObject {
  if (value === null) return { t: 'z' };
  if (typeof value === 'boolean') return { t: 'b', v: value };
  if (typeof value === 'number') return { t: 'n', v: value };
  return { t: 's', v: value };
}

type OdsFormulaUpdate = {
  sheet: string;
  address: string;
  formula: string;
};

function toOpenFormula(formula: string): string {
  if (formula.includes('!')) {
    throw new DocumentError(
      'INVALID_ARGUMENT',
      'ODS formula updates with cross-sheet references are not supported'
    );
  }
  const convertOutsideString = (segment: string): string => {
    const ranges: string[] = [];
    let converted = segment.replace(
      /(\$?[A-Za-z]{1,3}\$?\d+):(\$?[A-Za-z]{1,3}\$?\d+)/g,
      (_match, from: string, to: string) => {
        const marker = `__ADP_RANGE_${ranges.length}__`;
        ranges.push(`[.${from.toUpperCase()}:.${to.toUpperCase()}]`);
        return marker;
      }
    );
    converted = converted.replace(
      /\b(\$?[A-Za-z]{1,3}\$?\d+)\b/g,
      (match, address: string, offset: number, source: string) =>
        source[offset + match.length] === '('
          ? match
          : `[.${address.toUpperCase()}]`
    );
    ranges.forEach((range, index) => {
      converted = converted.replace(`__ADP_RANGE_${index}__`, range);
    });
    return converted;
  };

  return formula
    .split(/("(?:[^"]|"")*")/)
    .map((part) =>
      part.startsWith('"') && part.endsWith('"')
        ? part
        : convertOutsideString(part)
    )
    .join('');
}

function patchOdsCellFormula(
  contentXml: string,
  update: OdsFormulaUpdate
): string {
  const repeatCount = (element: string, attribute: string): number => {
    const value = new RegExp(`\\b${attribute}="(\\d+)"`).exec(element)?.[1];
    return value ? Math.max(1, Number(value)) : 1;
  };
  const withRepeat = (
    element: string,
    attribute: string,
    count: number
  ): string => {
    const expression = new RegExp(`\\s+${attribute}="\\d+"`);
    if (count <= 1) return element.replace(expression, '');
    if (expression.test(element)) {
      return element.replace(expression, ` ${attribute}="${count}"`);
    }
    return element.replace(/\/?>/, (end) => ` ${attribute}="${count}"${end}`);
  };
  const formulaCell = (cellXml: string): string => {
    if (/^<table:covered-table-cell\b/.test(cellXml)) {
      throw new DocumentError(
        'INVALID_ARGUMENT',
        `Cannot assign formula to covered merged cell ${update.sheet}!${update.address}`
      );
    }
    const open = /^<table:table-cell\b[^>]*\/?>/.exec(cellXml)?.[0];
    if (!open) {
      throw new DocumentError(
        'WRITE_FAILED',
        `Could not locate ODS cell ${update.sheet}!${update.address}`
      );
    }
    const retained = open
      .replace(/\/?>$/, '')
      .replace(/\s+table:number-columns-repeated="\d+"/g, '')
      .replace(
        /\s+(?:table:formula|office:string-value|office:value-type|office:value|calcext:value-type)="[^"]*"/g,
        ''
      );
    const formula = escapeXml(toOpenFormula(update.formula));
    return `${retained} table:formula="of:=${formula}" office:value-type="float" office:value="0"><text:p>0</text:p></table:table-cell>`;
  };
  const patchRow = (rowXml: string, targetColumn: number): string => {
    const cells = [
      ...rowXml.matchAll(
        /<table:(?:table-|covered-table-)?cell\b[^>]*\/>|<table:(?:table-|covered-table-)?cell\b[\s\S]*?<\/table:(?:table-)?cell>/g
      ),
    ];
    let logicalColumn = 0;
    for (const cell of cells) {
      const repeat = repeatCount(cell[0], 'table:number-columns-repeated');
      if (
        targetColumn < logicalColumn ||
        targetColumn >= logicalColumn + repeat
      ) {
        logicalColumn += repeat;
        continue;
      }
      const before = targetColumn - logicalColumn;
      const after = repeat - before - 1;
      const pieces: string[] = [];
      if (before > 0) {
        pieces.push(
          withRepeat(cell[0], 'table:number-columns-repeated', before)
        );
      }
      pieces.push(
        formulaCell(
          withRepeat(cell[0], 'table:number-columns-repeated', 1)
        )
      );
      if (after > 0) {
        pieces.push(
          withRepeat(cell[0], 'table:number-columns-repeated', after)
        );
      }
      const at = cell.index ?? 0;
      return (
        rowXml.slice(0, at) +
        pieces.join('') +
        rowXml.slice(at + cell[0].length)
      );
    }
    throw new DocumentError(
      'WRITE_FAILED',
      `Could not locate ODS column for ${update.sheet}!${update.address}`
    );
  };

  const sheetRegex =
    /<table:table\b[^>]*table:name="([^"]*)"[\s\S]*?<\/table:table>/g;
  let sheetMatch: RegExpExecArray | null;
  while ((sheetMatch = sheetRegex.exec(contentXml)) !== null) {
    if (decodeXmlEntities(sheetMatch[1]) !== update.sheet) continue;
    const sheetXml = sheetMatch[0];
    const point = XLSX.utils.decode_cell(update.address);
    const rows = [
      ...sheetXml.matchAll(
        /<table:table-row\b[^>]*\/>|<table:table-row\b[\s\S]*?<\/table:table-row>/g
      ),
    ];
    let logicalRow = 0;
    for (const row of rows) {
      const repeat = repeatCount(row[0], 'table:number-rows-repeated');
      if (point.r < logicalRow || point.r >= logicalRow + repeat) {
        logicalRow += repeat;
        continue;
      }
      const before = point.r - logicalRow;
      const after = repeat - before - 1;
      const pieces: string[] = [];
      if (before > 0) {
        pieces.push(withRepeat(row[0], 'table:number-rows-repeated', before));
      }
      const oneRow = withRepeat(row[0], 'table:number-rows-repeated', 1);
      pieces.push(patchRow(oneRow, point.c));
      if (after > 0) {
        pieces.push(withRepeat(row[0], 'table:number-rows-repeated', after));
      }
      const rowAt = row.index ?? 0;
      const patchedSheet =
        sheetXml.slice(0, rowAt) +
        pieces.join('') +
        sheetXml.slice(rowAt + row[0].length);
      return (
        contentXml.slice(0, sheetMatch.index) +
        patchedSheet +
        contentXml.slice(sheetMatch.index + sheetXml.length)
      );
    }
    break;
  }
  throw new DocumentError(
    'WRITE_FAILED',
    `Could not encode ODS formula at ${update.sheet}!${update.address}`
  );
}

async function patchOdsFormulas(
  data: Buffer,
  updates: OdsFormulaUpdate[]
): Promise<Buffer> {
  if (!updates.length) return data;
  const zip = await JSZip.loadAsync(data);
  const part = zip.file('content.xml');
  if (!part) throw new DocumentError('WRITE_FAILED', 'Generated ODS is missing content.xml');
  let xml = await part.async('string');
  for (const update of updates) xml = patchOdsCellFormula(xml, update);
  zip.file('content.xml', xml);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function writeWorkbook(
  filePath: string,
  wb: XLSX.WorkBook,
  odsFormulas: OdsFormulaUpdate[] = []
): Promise<void> {
  const bookType = bookTypeFor(filePath);
  let data = XLSX.write(wb, { type: 'buffer', bookType }) as Buffer;
  if (bookType === 'ods') data = await patchOdsFormulas(data, odsFormulas);
  await writeFileAtomic(filePath, data);
}

function assertValidSheetName(name: string): void {
  if (!name || name.length > 31 || /[\\/?*[\]:]/.test(name)) {
    throw new DocumentError(
      'INVALID_ARGUMENT',
      `Invalid sheet name "${name}": must be 1-31 chars without \\ / ? * [ ] :`
    );
  }
}

/** Expand worksheet !ref so a patched A1 address is included. */
function ensureCellInRange(ws: XLSX.WorkSheet, address: string): void {
  const cell = XLSX.utils.decode_cell(address);
  const range = ws['!ref']
    ? XLSX.utils.decode_range(ws['!ref'])
    : { s: { c: cell.c, r: cell.r }, e: { c: cell.c, r: cell.r } };
  if (cell.c < range.s.c) range.s.c = cell.c;
  if (cell.r < range.s.r) range.s.r = cell.r;
  if (cell.c > range.e.c) range.e.c = cell.c;
  if (cell.r > range.e.r) range.e.r = cell.r;
  ws['!ref'] = XLSX.utils.encode_range(range);
}

function sheetToGrid(ws: XLSX.WorkSheet): (string | number | boolean | null)[][] {
  return XLSX.utils.sheet_to_json<(string | number | boolean | null)[]>(ws, {
    header: 1,
    defval: null,
  }) as (string | number | boolean | null)[][];
}

export async function writeSpreadsheet(
  filePath: string,
  sheets: SheetGrid
): Promise<{ message: string }> {
  await ensureParentDir(filePath);
  const wb = XLSX.utils.book_new();
  for (const [name, rows] of Object.entries(sheets)) {
    assertValidSheetName(name);
    const ws = XLSX.utils.aoa_to_sheet(
      rows.map((row) => row.map((c) => (c === null ? '' : coerceValue(c))))
    );
    XLSX.utils.book_append_sheet(wb, ws, name);
  }
  await writeWorkbook(filePath, wb);
  const bookType = bookTypeFor(filePath);
  const label = bookType === 'biff8' ? 'XLS' : bookType.toUpperCase();
  return { message: `Created ${label}: ${filePath}` };
}

export async function updateSpreadsheetCells(
  filePath: string,
  updates: CellUpdate[]
): Promise<{ message: string; changes: number }> {
  const buf = await readFile(filePath);
  const wb = XLSX.read(buf, { type: 'buffer' });
  let count = 0;
  const odsFormulas: OdsFormulaUpdate[] = [];
  for (const u of updates) {
    const sheetName = u.sheet ?? wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    if (!ws) {
      throw new DocumentError('INVALID_ARGUMENT', `Sheet not found: ${sheetName}`);
    }
    const addr = u.cell.trim().toUpperCase();
    if (!/^[A-Z]+\d+$/.test(addr)) {
      throw new DocumentError('INVALID_ARGUMENT', `Invalid cell address: ${u.cell}`);
    }
    ensureCellInRange(ws, addr);
    const prev = ws[addr] as XLSX.CellObject | undefined;
    if (typeof u.value === 'string' && u.value.startsWith('=')) {
      // Explicit formula update.
      ws[addr] = { t: 'n', f: u.value.slice(1) };
      if (bookTypeFor(filePath) === 'ods') {
        odsFormulas.push({
          sheet: sheetName,
          address: addr,
          formula: u.value.slice(1),
        });
      }
    } else {
      const next = cellObject(coerceValue(u.value));
      // Literal value replaces any formula but keeps number format / style.
      if (prev?.z != null) next.z = prev.z;
      if (prev?.s != null) next.s = prev.s;
      ws[addr] = next;
    }
    count++;
  }
  await writeWorkbook(filePath, wb, odsFormulas);
  return { message: `Updated ${count} cell(s) in ${filePath}`, changes: count };
}

export function readSpreadsheetStructured(filePath: string): SpreadsheetStructured {
  const buf = readFileSync(filePath);
  return readSpreadsheetStructuredFromBuffer(buf, formatFor(filePath));
}

export async function readSpreadsheetStructuredAsync(
  filePath: string
): Promise<SpreadsheetStructured> {
  const buf = await readFile(filePath);
  return readSpreadsheetStructuredFromBuffer(buf, formatFor(filePath));
}

export function readSpreadsheetStructuredFromBuffer(
  buf: Buffer,
  format: 'xlsx' | 'xls' | 'ods' = 'xlsx'
): SpreadsheetStructured {
  const wb = XLSX.read(buf, { type: 'buffer' });
  const sheets: SheetGrid = {};
  for (const name of wb.SheetNames) {
    const grid = sheetToGrid(wb.Sheets[name]);
    const cells = grid.reduce((n, row) => n + row.length, 0);
    assertMaxSpreadsheetCells(cells);
    sheets[name] = grid;
  }
  return { format, sheets };
}

export function spreadsheetToText(structured: SpreadsheetStructured): string {
  const parts: string[] = [];
  for (const [name, rows] of Object.entries(structured.sheets)) {
    const csv = rows
      .map((row) => row.map((c) => csvEscapeCell(c)).join(','))
      .join('\n');
    parts.push(`## ${name}\n${csv}`);
  }
  return parts.join('\n\n');
}
