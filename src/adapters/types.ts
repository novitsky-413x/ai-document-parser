import type { DocumentInput, DocumentResult, ParseOptions } from '../types.js';
import type { ConvertOptions, ConvertResult } from '../types.js';
import type { AdapterCapabilities, CanonicalDocument, ConvertTarget, ParsedCanonical } from '../ir/types.js';

export type AdapterContext = {
  filePath?: string;
  buf: Buffer | null;
  rawText?: string;
  options: ParseOptions;
};

export type FormatAdapter = {
  capabilities: AdapterCapabilities;
  parse(ctx: AdapterContext): Promise<ParsedCanonical>;
  canConvertTo(target: ConvertTarget): boolean;
};

export type ConvertContext = {
  parsed: ParsedCanonical;
  target: ConvertTarget;
  options: ConvertOptions;
};

export type DocumentResultBuilder = (
  parsed: ParsedCanonical,
  output: NonNullable<ParseOptions['output']>
) => DocumentResult;
