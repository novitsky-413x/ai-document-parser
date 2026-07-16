import { readFileSync } from 'node:fs';
import { readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import * as XLSX from '@e965/xlsx';
import { DocumentError } from '../utils/errors.js';
import { ensureParentDir } from '../utils/paths.js';
import { csvEscapeCell } from '../utils/text.js';
import type { SpreadsheetStructured } from '../types.js';

export type SheetGrid = Record<string, (string | number | boolean | null)[][]>;

export type CellUpdate = {
  sheet?: string;
  cell: string;
  value: string | number | boolean | null;
};

function bookTypeFor(filePath: string): 'xlsx' | 'biff8' {
  return path.extname(filePath).toLowerCase() === '.xls' ? 'biff8' : 'xlsx';
}

function formatFor(filePath: string): 'xlsx' | 'xls' {
  return path.extname(filePath).toLowerCase() === '.xls' ? 'xls' : 'xlsx';
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

async function writeWorkbook(filePath: string, wb: XLSX.WorkBook): Promise<void> {
  const bookType = bookTypeFor(filePath);
  const data = XLSX.write(wb, { type: 'buffer', bookType }) as Buffer;
  await writeFile(filePath, data);
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
    const ws = XLSX.utils.aoa_to_sheet(
      rows.map((row) => row.map((c) => (c === null ? '' : coerceValue(c))))
    );
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  }
  await writeWorkbook(filePath, wb);
  const label = bookTypeFor(filePath) === 'biff8' ? 'XLS' : 'XLSX';
  return { message: `Created ${label}: ${filePath}` };
}

export async function updateSpreadsheetCells(
  filePath: string,
  updates: CellUpdate[]
): Promise<{ message: string; changes: number }> {
  const buf = await readFile(filePath);
  const wb = XLSX.read(buf, { type: 'buffer' });
  let count = 0;
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
    ws[addr] = cellObject(coerceValue(u.value));
    count++;
  }
  await writeWorkbook(filePath, wb);
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
  format: 'xlsx' | 'xls' = 'xlsx'
): SpreadsheetStructured {
  const wb = XLSX.read(buf, { type: 'buffer' });
  const sheets: SheetGrid = {};
  for (const name of wb.SheetNames) {
    sheets[name] = sheetToGrid(wb.Sheets[name]);
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
