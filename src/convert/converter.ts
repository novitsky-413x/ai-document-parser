import { DocumentError, wrapError } from '../utils/errors.js';
import type {
  ConvertFormat,
  ConvertOptions,
  ConvertResult,
  DocumentInput,
  DocumentResult,
  DocxStructured,
  PptxStructured,
  SpreadsheetStructured,
  ContentBlock,
} from '../types.js';
import { parseDocument } from '../parse/router.js';
import {
  emitCsvFromSheet,
  emitHtmlFromParagraphs,
  emitHtmlFromBlocks,
  emitMarkdownFromParagraphs,
  emitMarkdownFromBlocks,
  emitMarkdownFromText,
  emitHtmlFromText,
  emitText,
  emitTextFromSpreadsheet,
  chunkText,
  csvEscapeCell,
  blocksFromSpreadsheet,
  blocksFromCsvRows,
} from './emit.js';
import { pptxToText } from '../parse/pptx.js';
import { isPathLike } from '../utils/paths.js';
import { readFile } from 'node:fs/promises';
import { allowedConvertTargets } from '../adapters/registry.js';
import { contentBlocksToFlow } from '../ir/projections.js';
import { emitRtfFromFlow } from '../parse/rtf.js';
import { emitEpubFromFlow } from '../parse/epub.js';
import { writeDocxFromBlocks } from '../write/docx-structured.js';
import { writeOdtFromBlocks } from '../write/odt.js';

const FORMAT_MAP: Record<string, ConvertFormat | 'text'> = {
  md: 'md',
  markdown: 'md',
  html: 'html',
  csv: 'csv',
  text: 'text',
  txt: 'text',
  chunks: 'chunks',
  pdf: 'pdf',
  rtf: 'rtf',
  epub: 'epub',
  docx: 'docx',
  odt: 'odt',
  ods: 'ods',
  odp: 'odp',
};

export async function convertDocument(
  input: DocumentInput,
  to: ConvertFormat | string,
  options: ConvertOptions = {}
): Promise<ConvertResult> {
  const rawTarget = String(to).toLowerCase();
  const dest = FORMAT_MAP[rawTarget];
  if (!dest) {
    if (rawTarget === 'pptx' || rawTarget === 'odp') {
      throw new DocumentError(
        'UNSUPPORTED_FORMAT',
        `Convert target '${to}' is not supported by design. Create presentations with writeDocument({ type: 'pptx', … }); convert from pptx/odp only to md|html|text|chunks.`
      );
    }
    throw new DocumentError(
      'UNSUPPORTED_FORMAT',
      `Unsupported convert target: ${to}`
    );
  }

  try {
    if (dest === 'pdf') {
      throw new DocumentError(
        'UNSUPPORTED_FORMAT',
        'Convert to PDF is not supported by design. Use writeDocument({ type: \'pdf\', … }) for plain-text PDF creation — not convertDocument.'
      );
    }

    if (dest === 'md' && typeof input === 'string' && isPathLike(input)) {
      const ext = input.toLowerCase();
      if (ext.endsWith('.md') || ext.endsWith('.markdown')) {
        const raw = (await readFile(input, 'utf8')).replace(/\r\n/g, '\n');
        return { value: raw, messages: [], format: 'md' };
      }
    }

    const parsed = await parseDocument(input, { output: 'structured' });
    const allowed = allowedConvertTargets(parsed.format);
    if (!allowed.includes(dest as typeof allowed[number])) {
      const presentationHint =
        parsed.format === 'pptx' || parsed.format === 'odp'
          ? ` Presentation→flow Office (e.g. ${parsed.format}→docx) is unsupported by design — convert to md|html|text|chunks, or recreate flow docs with writeDocument.`
          : '';
      throw new DocumentError(
        'UNSUPPORTED_FORMAT',
        `Cannot convert ${parsed.format} to ${dest}. Allowed: ${allowed.join(', ')}.${presentationHint}`
      );
    }

    if (dest === 'md' && parsed.format === 'md') {
      const raw = (parsed.structured as { raw?: string } | undefined)?.raw;
      if (typeof raw === 'string') {
        return { value: raw.replace(/\r\n/g, '\n'), messages: parsed.warnings ?? [], format: 'md' };
      }
    }

    return await emitFromParsed(parsed, dest, options);
  } catch (err) {
    if (err instanceof DocumentError) throw err;
    throw wrapError('CONVERT_FAILED', `Failed to convert to ${to}`, err);
  }
}

function paragraphsOf(parsed: DocumentResult): string[] {
  const paras = (parsed.structured as { paragraphs?: string[] } | undefined)
    ?.paragraphs;
  if (Array.isArray(paras)) return paras;
  return [];
}

function blocksOf(parsed: DocumentResult): ContentBlock[] | undefined {
  const blocks = (parsed.structured as { blocks?: ContentBlock[] } | undefined)
    ?.blocks;
  return Array.isArray(blocks) ? blocks : undefined;
}

async function emitBinaryFlow(
  blocks: ContentBlock[],
  dest: 'rtf' | 'epub' | 'docx' | 'odt',
  messages: ConvertResult['messages']
): Promise<ConvertResult> {
  const flow = contentBlocksToFlow(blocks);
  if (dest === 'rtf') {
    return { value: emitRtfFromFlow(flow), messages, format: 'rtf' };
  }
  if (dest === 'epub') {
    return { value: await emitEpubFromFlow(flow), messages, format: 'epub' };
  }
  if (dest === 'docx') {
    return { value: await writeDocxFromBlocks(flow), messages, format: 'docx' };
  }
  return { value: await writeOdtFromBlocks(flow), messages, format: 'odt' };
}

function emitRich(
  blocks: ContentBlock[],
  dest: ConvertFormat | 'text',
  messages: ConvertResult['messages'],
  plainFallback: string
): ConvertResult | Promise<ConvertResult> {
  if (dest === 'md') {
    return { value: emitMarkdownFromBlocks(blocks), messages, format: 'md' };
  }
  if (dest === 'html') {
    return { value: emitHtmlFromBlocks(blocks), messages, format: 'html' };
  }
  if (dest === 'text') {
    return { value: emitText(plainFallback), messages, format: 'text' };
  }
  if (dest === 'chunks') {
    return { value: chunkText(plainFallback, blocks), messages, format: 'chunks' };
  }
  if (dest === 'csv') {
    const csv = plainFallback
      .split('\n')
      .map((line) => csvEscapeCell(line))
      .join('\n');
    return { value: csv, messages, format: 'csv' };
  }
  if (dest === 'rtf' || dest === 'epub' || dest === 'docx' || dest === 'odt') {
    return emitBinaryFlow(blocks, dest, messages);
  }
  throw new DocumentError('UNSUPPORTED_FORMAT', `Unsupported convert target: ${dest}`);
}

async function emitFromParsed(
  parsed: DocumentResult,
  dest: ConvertFormat | 'text',
  options: ConvertOptions = {}
): Promise<ConvertResult> {
  const messages = parsed.warnings ?? [];
  const format = parsed.format;

  if (format === 'xlsx' || format === 'xls' || format === 'ods') {
    const structured = parsed.structured as SpreadsheetStructured;
    if (dest === 'csv') {
      const sheetNames = Object.keys(structured.sheets);
      if (options.sheet != null && !sheetNames.includes(options.sheet)) {
        throw new DocumentError(
          'INVALID_ARGUMENT',
          `Sheet "${options.sheet}" not found. Available: ${sheetNames.join(', ')}`
        );
      }
      const sheet = options.sheet ?? sheetNames[0];
      if (options.sheet == null && sheetNames.length > 1) {
        messages.push({
          code: 'CSV_FIRST_SHEET',
          message: `CSV export uses first sheet "${sheetNames[0]}". Pass options.sheet to export another sheet.`,
        });
      }
      return { value: emitCsvFromSheet(structured, sheet), messages, format: 'csv' };
    }
    const blocks = blocksFromSpreadsheet(structured.sheets);
    const plain = emitTextFromSpreadsheet(structured);
    const rich = emitRich(blocks, dest, messages, plain);
    return rich instanceof Promise ? rich : rich;
  }

  if (format === 'csv') {
    const rows =
      (parsed.structured as { rows?: string[][] } | undefined)?.rows ?? [];
    const csvText = rows.map((row) => row.map(csvEscapeCell).join(',')).join('\n');
    if (dest === 'csv') {
      return { value: csvText, messages, format: 'csv' };
    }
    const blocks = blocksFromCsvRows(rows);
    if (dest === 'md' || dest === 'html' || dest === 'rtf' || dest === 'epub' || dest === 'docx' || dest === 'odt') {
      const rich = emitRich(blocks, dest, messages, csvText);
      return rich instanceof Promise ? rich : rich;
    }
    return emitFromPlain(csvText, dest, messages);
  }

  if (format === 'docx' || format === 'odt' || format === 'rtf' || format === 'epub') {
    const structured = parsed.structured as DocxStructured;
    const blocks = structured.blocks ?? [];
    const plain = (structured.paragraphs ?? []).join('\n\n');
    if (blocks.length) {
      const rich = emitRich(blocks, dest, messages, plain);
      return rich instanceof Promise ? rich : rich;
    }
    return emitFromParagraphs(structured.paragraphs ?? paragraphsOf(parsed), dest, messages);
  }

  if (format === 'doc') {
    return emitFromParagraphs(paragraphsOf(parsed), dest, messages);
  }

  if (format === 'pptx' || format === 'odp') {
    const structured = parsed.structured as PptxStructured;
    const plain = pptxToText(structured);
    if (structured.blocks?.length) {
      const rich = emitRich(structured.blocks, dest, messages, plain);
      return rich instanceof Promise ? rich : rich;
    }
    return emitFromParagraphs(
      structured.slides.flatMap((s, i) => [`Slide ${i + 1}`, ...s]),
      dest,
      messages
    );
  }

  if (format === 'html' || format === 'md') {
    const blocks = blocksOf(parsed);
    const plain = paragraphsOf(parsed).join('\n\n') || parsed.text;
    if (blocks?.length) {
      if (format === 'md' && dest === 'md') {
        const raw = (parsed.structured as { raw?: string })?.raw;
        if (raw) return { value: raw.replace(/\r\n/g, '\n'), messages, format: 'md' };
      }
      const rich = emitRich(blocks, dest, messages, plain);
      return rich instanceof Promise ? rich : rich;
    }
  }

  if (format === 'pdf' || format === 'image') {
    const blocks = blocksOf(parsed);
    const plain = paragraphsOf(parsed).join('\n\n') || parsed.text;
    if (blocks && blocks.length) {
      const rich = emitRich(blocks, dest, messages, plain);
      return rich instanceof Promise ? rich : rich;
    }
  }

  const paragraphs = paragraphsOf(parsed);
  if (paragraphs.length > 0) {
    return emitFromParagraphs(paragraphs, dest, messages);
  }

  return emitFromPlain(parsed.text || '', dest, messages);
}

function emitFromParagraphs(
  paragraphs: string[],
  dest: ConvertFormat | 'text',
  messages: ConvertResult['messages']
): ConvertResult | Promise<ConvertResult> {
  if (dest === 'md') {
    return {
      value: emitMarkdownFromParagraphs(paragraphs),
      messages,
      format: 'md',
    };
  }
  if (dest === 'html') {
    return {
      value: emitHtmlFromParagraphs(paragraphs),
      messages,
      format: 'html',
    };
  }
  if (dest === 'text') {
    return {
      value: emitText(paragraphs.join('\n')),
      messages,
      format: 'text',
    };
  }
  if (dest === 'chunks') {
    return {
      value: chunkText(paragraphs.join('\n\n')),
      messages,
      format: 'chunks',
    };
  }
  if (dest === 'csv') {
    const csv = paragraphs.map((line) => csvEscapeCell(line)).join('\n');
    return { value: csv, messages, format: 'csv' };
  }
  if (dest === 'rtf' || dest === 'epub' || dest === 'docx' || dest === 'odt') {
    const blocks: ContentBlock[] = paragraphs.map((p) => ({
      type: 'paragraph',
      runs: [{ text: p }],
    }));
    return emitBinaryFlow(blocks, dest, messages);
  }
  throw new DocumentError('UNSUPPORTED_FORMAT', `Unsupported convert target: ${dest}`);
}

function emitFromPlain(
  text: string,
  dest: ConvertFormat | 'text',
  messages: ConvertResult['messages']
): ConvertResult {
  if (dest === 'text') {
    return { value: emitText(text), messages, format: 'text' };
  }
  if (dest === 'md') {
    return { value: emitMarkdownFromText(text), messages, format: 'md' };
  }
  if (dest === 'html') {
    return { value: emitHtmlFromText(text), messages, format: 'html' };
  }
  if (dest === 'chunks') {
    return { value: chunkText(text), messages, format: 'chunks' };
  }
  if (dest === 'csv') {
    const csv = text
      .split('\n')
      .map((line) => csvEscapeCell(line))
      .join('\n');
    return { value: csv, messages, format: 'csv' };
  }
  throw new DocumentError('UNSUPPORTED_FORMAT', `Unsupported convert target: ${dest}`);
}
