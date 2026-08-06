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

const WRITE_FORMATS = new Set(['docx', 'xlsx', 'xls', 'pdf', 'odt', 'rtf', 'epub', 'pptx']);
/** Rich create (blocks / markdown / html) for flow / presentation writers. */
const RICH_WRITE_FORMATS = new Set(['docx', 'odt', 'rtf', 'epub', 'pptx']);
const EDIT_CAPABILITIES: Record<
  string,
  Pick<AdapterCapabilities, 'editModes' | 'editOps'>
> = {
  docx: {
    editModes: ['surgical', 'rewrite'],
    editOps: ['replace', 'append', 'rewrite'],
  },
  xlsx: { editModes: ['surgical'], editOps: ['patchCells'] },
  xls: { editModes: ['surgical'], editOps: ['patchCells'] },
  ods: { editModes: ['rewrite'], editOps: ['patchCells'] },
  pptx: { editModes: ['surgical'], editOps: ['replace'] },
  odt: {
    editModes: ['surgical', 'rewrite'],
    editOps: ['replace', 'rewrite'],
  },
  rtf: {
    editModes: ['surgical', 'rewrite'],
    editOps: ['replace', 'rewrite'],
  },
  epub: {
    editModes: ['surgical', 'rewrite'],
    editOps: ['replace', 'rewrite'],
  },
  md: { editModes: ['rewrite'], editOps: ['replace', 'rewrite'] },
  html: { editModes: ['rewrite'], editOps: ['replace', 'rewrite'] },
};

export function getCapabilities(filter?: string): AdapterCapabilities[] {
  const formats = [...CORE_PARSE_FORMATS];
  const list = filter ? formats.filter((f) => f === filter) : formats;
  return list.map((format) => {
    const write = WRITE_FORMATS.has(format);
    const writeModes = !write
      ? undefined
      : RICH_WRITE_FORMATS.has(format)
        ? (['plain', 'blocks', 'markdown', 'html'] as const)
        : (['plain'] as const);
    const editCapabilities = EDIT_CAPABILITIES[format];
    return {
      format,
      kind: KIND_MAP[format] ?? 'flow',
      parse: true,
      convertTo: CONVERT_MATRIX[format] ?? ['text', 'chunks'],
      edit: editCapabilities !== undefined,
      ...(editCapabilities ?? {}),
      write,
      ...(writeModes ? { writeModes: [...writeModes] } : {}),
    };
  });
}

export function allowedConvertTargets(sourceFormat: string): ConvertTarget[] {
  return CONVERT_MATRIX[sourceFormat] ?? ['text', 'chunks'];
}

export function assertConvertAllowed(sourceFormat: string, target: ConvertTarget): void {
  const allowed = allowedConvertTargets(sourceFormat);
  if (!allowed.includes(target)) {
    const presentationHint =
      sourceFormat === 'pptx' || sourceFormat === 'odp'
        ? ' Presentation→docx/odt/rtf/epub is unsupported by design.'
        : '';
    throw new DocumentError(
      'UNSUPPORTED_FORMAT',
      `Convert from ${sourceFormat} to ${target} is not supported. Allowed: ${allowed.join(', ')}.${presentationHint}`
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
