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
} from '../types.js';
import { parseDocument } from '../parse/router.js';
import {
  emitCsvFromSheet,
  emitHtmlFromParagraphs,
  emitHtmlFromPptx,
  emitHtmlFromText,
  emitMarkdownFromParagraphs,
  emitMarkdownFromPptx,
  emitMarkdownFromSpreadsheet,
  emitMarkdownFromText,
  emitText,
  emitTextFromSpreadsheet,
  chunkText,
  csvEscapeCell,
} from './emit.js';
import { pptxToText } from '../parse/pptx.js';

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

    const parsed = await parseDocument(input, { output: 'structured' });
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
    if (dest === 'md') {
      return {
        value: emitMarkdownFromSpreadsheet(structured),
        messages,
        format: 'md',
      };
    }
    if (dest === 'text') {
      return {
        value: emitTextFromSpreadsheet(structured),
        messages,
        format: 'text',
      };
    }
    if (dest === 'html') {
      return {
        value: emitHtmlFromText(emitTextFromSpreadsheet(structured)),
        messages,
        format: 'html',
      };
    }
    if (dest === 'chunks') {
      return {
        value: chunkText(emitTextFromSpreadsheet(structured)),
        messages,
        format: 'chunks',
      };
    }
  }

  if (format === 'csv') {
    const rows =
      (parsed.structured as { rows?: string[][] } | undefined)?.rows ?? [];
    const csvText = rows.map((row) => row.map(csvEscapeCell).join(',')).join('\n');
    if (dest === 'csv') {
      return { value: csvText, messages, format: 'csv' };
    }
    return emitFromPlain(csvText, dest, messages);
  }

  if (format === 'docx' || format === 'doc') {
    if (dest === 'csv') {
      throw new DocumentError(
        'UNSUPPORTED_FORMAT',
        'Cannot convert DOCX/DOC to CSV; use a spreadsheet source.'
      );
    }
    const paragraphs =
      (parsed.structured as DocxStructured | undefined)?.paragraphs ??
      paragraphsOf(parsed);
    return emitFromParagraphs(paragraphs, dest, messages);
  }

  if (format === 'pptx') {
    if (dest === 'csv') {
      throw new DocumentError(
        'UNSUPPORTED_FORMAT',
        'Cannot convert PPTX to CSV; use md, html, text, or chunks.'
      );
    }
    const slides =
      (parsed.structured as PptxStructured | undefined)?.slides ?? [];
    if (dest === 'md') {
      return {
        value: emitMarkdownFromPptx(slides),
        messages,
        format: 'md',
      };
    }
    if (dest === 'html') {
      return {
        value: emitHtmlFromPptx(slides),
        messages,
        format: 'html',
      };
    }
    if (dest === 'text') {
      return {
        value: emitText(pptxToText({ format: 'pptx', slides })),
        messages,
        format: 'text',
      };
    }
    if (dest === 'chunks') {
      return {
        value: chunkText(pptxToText({ format: 'pptx', slides })),
        messages,
        format: 'chunks',
      };
    }
  }

  // md / html / pdf / other paragraph-shaped structured
  const paragraphs = paragraphsOf(parsed);
  if (paragraphs.length > 0) {
    return emitFromParagraphs(paragraphs, dest, messages);
  }

  return emitFromPlain('', dest, messages);
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
