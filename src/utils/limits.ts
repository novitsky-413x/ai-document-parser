import { DocumentError } from './errors.js';

/** Default resource limits for untrusted document inputs. */
export const DEFAULT_LIMITS = {
  maxFileBytes: 64 * 1024 * 1024,
  maxZipEntries: 4096,
  maxZipExpandedBytes: 256 * 1024 * 1024,
  maxPdfPages: 2000,
  maxSpreadsheetCells: 2_000_000,
} as const;

export type ResourceLimits = typeof DEFAULT_LIMITS;

export function assertMaxBytes(size: number, limit = DEFAULT_LIMITS.maxFileBytes): void {
  if (size > limit) {
    throw new DocumentError(
      'PARSE_FAILED',
      `Input exceeds maximum size (${size} > ${limit} bytes)`
    );
  }
}

export function assertMaxPdfPages(pages: number, limit = DEFAULT_LIMITS.maxPdfPages): void {
  if (pages > limit) {
    throw new DocumentError(
      'PARSE_FAILED',
      `PDF exceeds maximum page count (${pages} > ${limit})`
    );
  }
}

export function assertMaxSpreadsheetCells(
  cells: number,
  limit = DEFAULT_LIMITS.maxSpreadsheetCells
): void {
  if (cells > limit) {
    throw new DocumentError(
      'PARSE_FAILED',
      `Spreadsheet exceeds maximum cell count (${cells} > ${limit})`
    );
  }
}
