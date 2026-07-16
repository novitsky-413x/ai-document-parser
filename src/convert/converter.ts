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
};

/**
 * Convert a document to another format in one step (parse + generate).
 * Driven by parseDocument structured result — works for paths and buffers alike.
 */
export async function convertDocument(
  input: DocumentInput,
  to: ConvertFormat | string,
  options: ConvertOptions = {}
): Promise<ConvertResult> {
  void options;
  const dest = FORMAT_MAP[String(to).toLowerCase()];
  if (!dest) {
    throw new DocumentError(
      'UNSUPPORTED_FORMAT',
      `Unsupported convert target: ${to}`
    );
  }

  try {
    if (dest === 'rtf' || dest === 'epub') {
      throw new DocumentError(
        'UNSUPPORTED_FORMAT',
        `Convert target "${dest}" is not supported. Use md, html, csv, text, or chunks.`
      );
    }

    if (dest === 'pdf') {
      throw new DocumentError(
        'UNSUPPORTED_FORMAT',
        'Convert to PDF is not supported in core. Use writeDocument for plain-text PDF creation.'
      );
    }

    // Markdown → md: passthrough original source text
    if (dest === 'md' && typeof input === 'string' && isPathLike(input)) {
      const ext = input.toLowerCase();
      if (ext.endsWith('.md') || ext.endsWith('.markdown')) {
        const raw = (await readFile(input, 'utf8')).replace(/\r\n/g, '\n');
        return { value: raw, messages: [], format: 'md' };
      }
    }

    const parsed = await parseDocument(input, { output: 'structured' });

    // MD structured may carry raw for passthrough when input was buffer/string content
    if (dest === 'md' && parsed.format === 'md') {
      const raw = (parsed.structured as { raw?: string } | undefined)?.raw;
      if (typeof raw === 'string') {
        return { value: raw.replace(/\r\n/g, '\n'), messages: parsed.warnings ?? [], format: 'md' };
      }
    }

    return emitFromParsed(parsed, dest);
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

function emitRich(
  blocks: ContentBlock[],
  dest: ConvertFormat | 'text',
  messages: ConvertResult['messages'],
  plainFallback: string
): ConvertResult {
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
    return { value: chunkText(plainFallback), messages, format: 'chunks' };
  }
  if (dest === 'csv') {
    const csv = plainFallback
      .split('\n')
      .map((line) => csvEscapeCell(line))
      .join('\n');
    return { value: csv, messages, format: 'csv' };
  }
  throw new DocumentError('UNSUPPORTED_FORMAT', `Unsupported convert target: ${dest}`);
}

function emitFromParsed(
  parsed: DocumentResult,
  dest: ConvertFormat | 'text'
): ConvertResult {
  const messages = parsed.warnings ?? [];
  const format = parsed.format;

  if (format === 'xlsx' || format === 'xls') {
    const structured = parsed.structured as SpreadsheetStructured;
    if (dest === 'csv') {
      return { value: emitCsvFromSheet(structured), messages, format: 'csv' };
    }
    const blocks = blocksFromSpreadsheet(structured.sheets);
    const plain = emitTextFromSpreadsheet(structured);
    return emitRich(blocks, dest, messages, plain);
  }

  if (format === 'csv') {
    const rows =
      (parsed.structured as { rows?: string[][] } | undefined)?.rows ?? [];
    const csvText = rows.map((row) => row.map(csvEscapeCell).join(',')).join('\n');
    if (dest === 'csv') {
      return { value: csvText, messages, format: 'csv' };
    }
    const blocks = blocksFromCsvRows(rows);
    if (dest === 'md' || dest === 'html') {
      return emitRich(blocks, dest, messages, csvText);
    }
    return emitFromPlain(csvText, dest, messages);
  }

  if (format === 'docx') {
    if (dest === 'csv') {
      throw new DocumentError(
        'UNSUPPORTED_FORMAT',
        'Cannot convert DOCX/DOC to CSV; use a spreadsheet source.'
      );
    }
    const structured = parsed.structured as DocxStructured;
    const blocks = structured.blocks ?? [];
    const plain = (structured.paragraphs ?? []).join('\n\n');
    if (blocks.length && (dest === 'md' || dest === 'html')) {
      return emitRich(blocks, dest, messages, plain);
    }
    return emitFromParagraphs(structured.paragraphs ?? paragraphsOf(parsed), dest, messages);
  }

  if (format === 'doc') {
    if (dest === 'csv') {
      throw new DocumentError(
        'UNSUPPORTED_FORMAT',
        'Cannot convert DOCX/DOC to CSV; use a spreadsheet source.'
      );
    }
    return emitFromParagraphs(paragraphsOf(parsed), dest, messages);
  }

  if (format === 'pptx') {
    if (dest === 'csv') {
      throw new DocumentError(
        'UNSUPPORTED_FORMAT',
        'Cannot convert PPTX to CSV; use md, html, text, or chunks.'
      );
    }
    const structured = parsed.structured as PptxStructured;
    const plain = pptxToText(structured);
    if (structured.blocks?.length && (dest === 'md' || dest === 'html')) {
      return emitRich(structured.blocks, dest, messages, plain);
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
    if (blocks?.length && (dest === 'md' || dest === 'html')) {
      // md→md already handled via raw passthrough above when possible
      if (format === 'md' && dest === 'md') {
        const raw = (parsed.structured as { raw?: string })?.raw;
        if (raw) return { value: raw.replace(/\r\n/g, '\n'), messages, format: 'md' };
      }
      return emitRich(blocks, dest, messages, plain);
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
): ConvertResult {
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
