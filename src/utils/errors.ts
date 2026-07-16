import type { DocumentErrorCode } from '../types.js';

export class DocumentError extends Error {
  readonly code: DocumentErrorCode;
  readonly cause?: unknown;

  constructor(code: DocumentErrorCode, message: string, cause?: unknown) {
    super(message);
    this.name = 'DocumentError';
    this.code = code;
    this.cause = cause;
  }
}

export function wrapError(
  code: DocumentErrorCode,
  message: string,
  err: unknown
): DocumentError {
  if (err instanceof DocumentError) return err;
  const detail = err instanceof Error ? err.message : String(err);
  return new DocumentError(code, `${message}: ${detail}`, err);
}
