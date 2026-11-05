import type JSZip from 'jszip';
import { DocumentError } from '../utils/errors.js';
import { extOf } from '../utils/paths.js';
import { stripBom } from '../utils/text.js';
import { loadZipBuffer } from '../utils/zip.js';
import { CORE_FORMATS_HELP, CORE_PARSE_FORMAT_SET } from './formats.js';
import { detectOdfKind } from './odf.js';

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
  | 'odt'
  | 'ods'
  | 'odp'
  | 'rtf'
  | 'epub'
  | 'image'
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
  '.odt': 'odt',
  '.ods': 'ods',
  '.odp': 'odp',
  '.rtf': 'rtf',
  '.epub': 'epub',
  '.fb2': 'html',
  '.png': 'image',
  '.jpg': 'image',
  '.jpeg': 'image',
  '.tif': 'image',
  '.tiff': 'image',
  '.gif': 'image',
  '.webp': 'image',
  '.bmp': 'image',
};

const UNSUPPORTED_EXT = new Set(['.ppt']);

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
  const trimmed = sample.trimStart();
  if (/^\s*</.test(trimmed)) return 'html';
  if (/^#{1,6}\s/.test(trimmed)) return 'md';
  if (/^[\w-]+,[\w-]+(?:,[\w-]+)+\s*$/m.test(sample.slice(0, 400)) && !trimmed.startsWith('<'))
    return 'csv';
  return 'md';
}

export function isZipBuffer(buf: Buffer): boolean {
  return buf.length >= 2 && buf[0] === 0x50 && buf[1] === 0x4b;
}

/** OLE CFB directory entries store stream names as UTF-16LE. */
function cfbHasStream(buf: Buffer, streamName: string): boolean {
  return buf.includes(Buffer.from(streamName, 'utf16le'));
}

/**
 * Detect format from magic bytes, then UTF-8 text heuristics for buffers
 * without a binary signature.
 * ZIP containers return a provisional `'docx'`; callers with a buffer should
 * follow up with `refineZipFormat` when the extension is unknown.
 */
function isImageBuffer(buf: Buffer): boolean {
  if (
    buf.length >= 8 &&
    buf[0] === 0x89 &&
    buf[1] === 0x50 &&
    buf[2] === 0x4e &&
    buf[3] === 0x47
  ) {
    return true;
  }
  if (buf.length >= 3 && buf[0] === 0xff && buf[1] === 0xd8 && buf[2] === 0xff) return true;
  if (buf.length >= 6) {
    const gif = buf.slice(0, 6).toString('ascii');
    if (gif === 'GIF87a' || gif === 'GIF89a') return true;
  }
  if (
    buf.length >= 4 &&
    buf[0] === 0x49 &&
    buf[1] === 0x49 &&
    buf[2] === 0x2a &&
    buf[3] === 0x00
  ) {
    return true;
  }
  if (
    buf.length >= 4 &&
    buf[0] === 0x4d &&
    buf[1] === 0x4d &&
    buf[2] === 0x00 &&
    buf[3] === 0x2a
  ) {
    return true;
  }
  if (
    buf.length >= 12 &&
    buf.slice(0, 4).toString('ascii') === 'RIFF' &&
    buf.slice(8, 12).toString('ascii') === 'WEBP'
  ) {
    return true;
  }
  return false;
}

export function detectFormatFromBuffer(buf: Buffer): DetectedFormat | null {
  if (buf.length >= 4 && buf.slice(0, 4).toString('ascii') === '%PDF') return 'pdf';
  if (isImageBuffer(buf)) return 'image';
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
    // Same magic for legacy Word and Excel; distinguish by CFB stream names.
    if (cfbHasStream(buf, 'WordDocument')) return 'doc';
    if (cfbHasStream(buf, 'Workbook') || cfbHasStream(buf, 'Book')) return 'xls';
    return 'doc';
  }
  if (buf.length === 0) return null;
  const sample = stripBom(buf.slice(0, 200).toString('utf8'));
  if (/^\s*\{\\rtf/i.test(sample)) return 'rtf';
  return detectTextFormat(sample);
}

/**
 * Distinguish OOXML ZIP packages (docx / xlsx / pptx) via package parts
 * and `[Content_Types].xml`.
 */
export async function refineZipFormat(buf: Buffer): Promise<DetectedFormat> {
  try {
    const zip = await loadZipBuffer(buf);
    const odf = await detectOdfKind(buf);
    if (odf) return odf;
    if (zip.file('ppt/presentation.xml') || hasSlidePart(zip)) return 'pptx';
    if (zip.file('xl/workbook.xml')) return 'xlsx';
    if (zip.file('word/document.xml')) return 'docx';
    if (zip.file('META-INF/container.xml')) return 'epub';

    const ct = zip.file('[Content_Types].xml');
    if (ct) {
      const xml = await ct.async('string');
      if (/epub\+zip/i.test(xml)) return 'epub';
      if (/presentationml/i.test(xml) || /\/ppt\//i.test(xml)) return 'pptx';
      if (/spreadsheetml/i.test(xml) || /\/xl\//i.test(xml)) return 'xlsx';
      if (/wordprocessingml/i.test(xml) || /\/word\//i.test(xml)) return 'docx';
      if (/opendocument\.text/i.test(xml)) return 'odt';
      if (/opendocument\.spreadsheet/i.test(xml)) return 'ods';
      if (/opendocument\.presentation/i.test(xml)) return 'odp';
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
  const sample = stripBom(text).slice(0, 200);
  if (/^\s*\{\\rtf/i.test(sample)) return 'rtf';
  return detectTextFormat(sample);
}

export function assertSupportedParse(format: string): void {
  if (!CORE_PARSE_FORMAT_SET.has(format)) {
    throw new DocumentError(
      'UNSUPPORTED_FORMAT',
      `Unsupported format "${format}". Core formats: ${CORE_FORMATS_HELP}.`
    );
  }
}
