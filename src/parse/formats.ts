/**
 * Clean-room rules for contributors (also summarized in AGENTS.md):
 *
 * 1. Implement parsers/converters from public format specs and public API behavior
 *    (tests + docs/API.md) — do not copy code from historical derived engines.
 * 2. Prefer extending src/edit, src/write, src/parse over large opaque backends.
 * 3. New format support lands behind parseDocument / convertDocument only.
 */

/** Normalized format ids returned by detect / parseDocument. */
export const CORE_PARSE_FORMATS = [
  'docx',
  'xlsx',
  'xls',
  'pdf',
  'md',
  'html',
  'csv',
  'doc',
  'pptx',
  'odt',
  'ods',
  'odp',
  'rtf',
  'epub',
] as const;

export type CoreParseFormat = (typeof CORE_PARSE_FORMATS)[number];

export const CORE_PARSE_FORMAT_SET = new Set<string>(CORE_PARSE_FORMATS);

export const CORE_FORMATS_HELP =
  'docx, xlsx, xls, pptx, pdf, md, html, csv, legacy doc, odt, ods, odp, rtf, epub';
