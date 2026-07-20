import { DocumentError } from '../utils/errors.js';
import type { AdapterCapabilities, ConvertTarget } from '../ir/types.js';
import { CORE_PARSE_FORMATS } from '../parse/formats.js';

const CONVERT_MATRIX: Record<string, ConvertTarget[]> = {
  docx: ['md', 'html', 'text', 'chunks', 'docx', 'odt', 'rtf', 'epub'],
  doc: ['md', 'html', 'text', 'chunks', 'docx', 'odt', 'rtf', 'epub'],
  md: ['md', 'html', 'text', 'chunks', 'csv', 'docx', 'odt', 'rtf', 'epub'],
  html: ['md', 'html', 'text', 'chunks', 'csv', 'docx', 'odt', 'rtf', 'epub'],
  csv: ['md', 'html', 'text', 'chunks', 'csv', 'docx', 'odt', 'rtf', 'epub'],
  pdf: ['md', 'html', 'text', 'chunks'],
  xlsx: ['csv', 'md', 'html', 'text', 'chunks'],
  xls: ['csv', 'md', 'html', 'text', 'chunks'],
  ods: ['csv', 'md', 'html', 'text', 'chunks'],
  pptx: ['md', 'html', 'text', 'chunks'],
  odp: ['md', 'html', 'text', 'chunks'],
  odt: ['md', 'html', 'text', 'chunks', 'docx', 'odt', 'rtf', 'epub'],
  rtf: ['md', 'html', 'text', 'chunks', 'docx', 'odt', 'rtf', 'epub'],
  epub: ['md', 'html', 'text', 'chunks', 'docx', 'odt', 'rtf', 'epub'],
};

const KIND_MAP: Record<string, AdapterCapabilities['kind']> = {
  docx: 'flow',
  doc: 'flow',
  md: 'flow',
  html: 'flow',
  csv: 'flow',
  pdf: 'flow',
  odt: 'flow',
  rtf: 'flow',
  epub: 'flow',
  xlsx: 'workbook',
  xls: 'workbook',
  ods: 'workbook',
  pptx: 'presentation',
  odp: 'presentation',
};

const EDIT_FORMATS = new Set(['docx', 'xlsx', 'xls']);
const WRITE_FORMATS = new Set(['docx', 'xlsx', 'xls', 'pdf', 'odt', 'rtf', 'epub']);

export function getCapabilities(filter?: string): AdapterCapabilities[] {
  const formats = [...CORE_PARSE_FORMATS];
  const list = filter ? formats.filter((f) => f === filter) : formats;
  return list.map((format) => ({
    format,
    kind: KIND_MAP[format] ?? 'flow',
    parse: true,
    convertTo: CONVERT_MATRIX[format] ?? ['text', 'chunks'],
    edit: EDIT_FORMATS.has(format),
    write: WRITE_FORMATS.has(format),
  }));
}

export function allowedConvertTargets(sourceFormat: string): ConvertTarget[] {
  return CONVERT_MATRIX[sourceFormat] ?? ['text', 'chunks'];
}

export function assertConvertAllowed(sourceFormat: string, target: ConvertTarget): void {
  const allowed = allowedConvertTargets(sourceFormat);
  if (!allowed.includes(target)) {
    throw new DocumentError(
      'UNSUPPORTED_FORMAT',
      `Convert from ${sourceFormat} to ${target} is not supported. Allowed: ${allowed.join(', ')}`
    );
  }
}

export function semanticConvertAllowed(
  sourceKind: AdapterCapabilities['kind'],
  target: ConvertTarget
): boolean {
  if (target === 'text' || target === 'chunks') return true;
  if (sourceKind === 'workbook') return ['csv', 'md', 'html', 'text', 'chunks'].includes(target);
  if (sourceKind === 'presentation') return ['md', 'html', 'text', 'chunks'].includes(target);
  return ['md', 'html', 'text', 'chunks', 'docx', 'odt', 'rtf', 'epub'].includes(target);
}
