import JSZip from 'jszip';
import { DocumentError } from '../utils/errors.js';
import { extOf } from '../utils/paths.js';
import { CORE_FORMATS_HELP, CORE_PARSE_FORMAT_SET } from './formats.js';

export type DetectedFormat =
  | 'docx'
  | 'xlsx'
  | 'xls'
  | 'pdf'
  | 'md'
  | 'html'
  | 'csv'
  | 'doc'
  | 'pptx'
  | 'unknown';

const EXT_MAP: Record<string, DetectedFormat> = {
  '.docx': 'docx',
  '.xlsx': 'xlsx',
  '.xls': 'xls',
  '.pptx': 'pptx',
  '.pdf': 'pdf',
  '.md': 'md',
  '.markdown': 'md',
  '.html': 'html',
  '.htm': 'html',
  '.csv': 'csv',
  '.doc': 'doc',
};

const UNSUPPORTED_EXT = new Set([
  '.ppt',
  '.odt',
  '.ods',
  '.odp',
  '.rtf',
  '.epub',
]);

export function formatFromExt(filePath: string | undefined): DetectedFormat | null {
  if (!filePath) return null;
  const ext = extOf(filePath);
  if (UNSUPPORTED_EXT.has(ext)) {
    throw new DocumentError(
      'UNSUPPORTED_FORMAT',
      `Unsupported format "${ext}". Core formats: ${CORE_FORMATS_HELP}.`
    );
  }
  return EXT_MAP[ext] ?? null;
}

function detectTextFormat(sample: string): DetectedFormat {
  if (/^\s*</.test(sample)) return 'html';
  if (sample.includes(',') && !sample.includes('<')) return 'csv';
  return 'md';
}

export function isZipBuffer(buf: Buffer): boolean {
  return buf.length >= 2 && buf[0] === 0x50 && buf[1] === 0x4b;
}

/**
 * Detect format from magic bytes, then UTF-8 text heuristics for buffers
 * without a binary signature.
 * ZIP containers return a provisional `'docx'`; callers with a buffer should
 * follow up with `refineZipFormat` when the extension is unknown.
 */
export function detectFormatFromBuffer(buf: Buffer): DetectedFormat | null {
  if (buf.length >= 4 && buf.slice(0, 4).toString('ascii') === '%PDF') return 'pdf';
  if (isZipBuffer(buf)) {
    return 'docx';
  }
  if (
    buf.length >= 8 &&
    buf[0] === 0xd0 &&
    buf[1] === 0xcf &&
    buf[2] === 0x11 &&
    buf[3] === 0xe0
  ) {
    return 'doc';
  }
  if (buf.length === 0) return null;
  return detectTextFormat(buf.slice(0, 200).toString('utf8'));
}

/**
 * Distinguish OOXML ZIP packages (docx / xlsx / pptx) via package parts
 * and `[Content_Types].xml`.
 */
export async function refineZipFormat(buf: Buffer): Promise<DetectedFormat> {
  try {
    const zip = await JSZip.loadAsync(buf);
    if (zip.file('ppt/presentation.xml') || hasSlidePart(zip)) return 'pptx';
    if (zip.file('xl/workbook.xml')) return 'xlsx';
    if (zip.file('word/document.xml')) return 'docx';

    const ct = zip.file('[Content_Types].xml');
    if (ct) {
      const xml = await ct.async('string');
      if (/presentationml/i.test(xml) || /\/ppt\//i.test(xml)) return 'pptx';
      if (/spreadsheetml/i.test(xml) || /\/xl\//i.test(xml)) return 'xlsx';
      if (/wordprocessingml/i.test(xml) || /\/word\//i.test(xml)) return 'docx';
    }
  } catch {
    /* fall through */
  }
  return 'docx';
}

function hasSlidePart(zip: JSZip): boolean {
  return Object.keys(zip.files).some((n) =>
    /^ppt\/slides\/slide\d+\.xml$/i.test(n)
  );
}

/** Detect format for a raw string that is not a filesystem path. */
export function detectFormatFromText(text: string): DetectedFormat {
  return detectTextFormat(text.slice(0, 200));
}

export function assertSupportedParse(format: string): void {
  if (!CORE_PARSE_FORMAT_SET.has(format)) {
    throw new DocumentError(
      'UNSUPPORTED_FORMAT',
      `Unsupported format "${format}". Core formats: ${CORE_FORMATS_HELP}.`
    );
  }
}
