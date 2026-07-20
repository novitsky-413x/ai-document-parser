import { readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import { DocumentError, wrapError } from '../utils/errors.js';
import { extOf, isPathLike } from '../utils/paths.js';
import { stripBom } from '../utils/text.js';
import { readFileLimited } from '../utils/zip.js';
import type {
  DocumentInput,
  DocumentResult,
  ParseOptions,
} from '../types.js';
import { readLegacyDoc, readLegacyDocBuffer } from './legacy-doc.js';
import {
  readDocxStructured,
  readDocxStructuredFromBuffer,
} from '../edit/docx-ooxml.js';
import {
  readSpreadsheetStructuredAsync,
  readSpreadsheetStructuredFromBuffer,
  spreadsheetToText,
} from '../edit/spreadsheet.js';
import {
  assertSupportedParse,
  detectFormatFromBuffer,
  detectFormatFromText,
  formatFromExt,
  isZipBuffer,
  refineZipFormat,
  type DetectedFormat,
} from './detect.js';
import { extractPdfText } from './pdf.js';
import {
  pptxToParagraphs,
  pptxToText,
  readPptxStructured,
  readPptxStructuredFromBuffer,
} from './pptx.js';
import {
  parseCsv,
  parseHtml,
  parseMarkdown,
  readTextFile,
} from './text-formats.js';
import { htmlToStructured, markdownToBlocks } from './html-blocks.js';
import { flattenBlocksToParagraphs } from '../convert/blocks.js';
import { chunkText } from '../convert/emit.js';
import { parseOdfBuffer } from './odf.js';
import { parseRtf } from './rtf.js';
import { parseEpub } from './epub.js';
import { buildParsedCanonical } from '../ir/projections.js';
import { contentBlocksToFlow } from '../ir/projections.js';

function toBuffer(input: DocumentInput): Buffer | null {
  if (typeof input === 'string') return null;
  if (Buffer.isBuffer(input)) return input;
  if (input instanceof ArrayBuffer) return Buffer.from(input);
  if (input instanceof Uint8Array) return Buffer.from(input);
  return null;
}

async function resolveInput(
  input: DocumentInput
): Promise<{ format: DetectedFormat; buf: Buffer | null; filePath: string | undefined }> {
  let buf = toBuffer(input);
  const filePath = typeof input === 'string' && isPathLike(input) ? input : undefined;

  if (filePath && !buf) {
    buf = await readFileLimited(filePath);
  }

  const format = await resolveFormat(input, buf, filePath);
  return { format, buf, filePath };
}

const BINARY_FORMATS = new Set(['docx', 'xlsx', 'xls', 'pptx', 'pdf', 'doc', 'odt', 'ods', 'odp', 'epub']);

async function resolveFormat(
  input: DocumentInput,
  buf: Buffer | null,
  filePath: string | undefined
): Promise<DetectedFormat> {
  const fromExt = formatFromExt(filePath);

  if (buf) {
    const guessed = detectFormatFromBuffer(buf);
    if (guessed === 'docx' && isZipBuffer(buf)) {
      const refined = await refineZipFormat(buf);
      if (fromExt && fromExt !== 'unknown' && BINARY_FORMATS.has(fromExt) && fromExt !== refined) {
        return refined;
      }
      return refined;
    }
    if (guessed && BINARY_FORMATS.has(guessed)) {
      if (fromExt === 'xls' && guessed === 'doc') return 'xls';
      if (fromExt === 'doc' && guessed === 'doc') return 'doc';
      if (fromExt && fromExt !== 'unknown' && fromExt !== guessed) return guessed;
      return guessed;
    }
  }

  if (fromExt && fromExt !== 'unknown') return fromExt;

  if (buf) {
    const guessed = detectFormatFromBuffer(buf);
    if (guessed) return guessed;
  }

  if (typeof input === 'string' && !isPathLike(input)) {
    return detectFormatFromText(input);
  }

  throw new DocumentError(
    'UNSUPPORTED_FORMAT',
    'Could not detect document format. Pass a path with a known extension or binary buffer.'
  );
}

/**
 * Parse a document into text, structured JSON (agent edit path), AST, or chunks.
 */
export async function parseDocument(
  input: DocumentInput,
  options: ParseOptions = {}
): Promise<DocumentResult> {
  const output = options.output ?? 'text';

  try {
    const { format, buf, filePath } = await resolveInput(input);
    assertSupportedParse(format);

    // Legacy .doc
    if (format === 'doc') {
      const legacy = filePath
        ? await readLegacyDoc(filePath)
        : await readLegacyDocBuffer(buf!);
      if (output === 'structured') {
        const structured = {
          format: 'doc' as const,
          paragraphs: legacy.paragraphs,
          source: legacy.source,
        };
        const text = JSON.stringify(structured, null, 2);
        return {
          format: 'doc',
          text,
          charCount: text.length,
          path: filePath,
          structured,
          warnings: [
            {
              code: 'LEGACY_DOC_LOW_FIDELITY',
              message:
                'Legacy .doc extraction is heuristic; formatting is not preserved.',
            },
          ],
        };
      }
      return finish(format, legacy.text, filePath, output, {
        paragraphs: legacy.paragraphs,
      });
    }

    // DOCX
    if (format === 'docx') {
      try {
        const structured = filePath
          ? await readDocxStructured(filePath)
          : await readDocxStructuredFromBuffer(buf!);
        const text = structured.paragraphs.join('\n');
        if (output === 'structured') {
          const payload = JSON.stringify(structured, null, 2);
          return {
            format: 'docx',
            text: payload,
            charCount: payload.length,
            path: filePath,
            structured,
          };
        }
        return finish('docx', text, filePath, output, {
          paragraphs: structured.paragraphs,
          blocks: structured.blocks,
        });
      } catch (err) {
        // ZIP might be xlsx/pptx without a reliable extension
        if (buf && isZipBuffer(buf)) {
          try {
            const refined = await refineZipFormat(buf);
            if (refined === 'pptx') {
              const structured = await readPptxStructuredFromBuffer(buf);
              return pptxResult(structured, filePath, output);
            }
            if (refined === 'xlsx') {
              const structured = readSpreadsheetStructuredFromBuffer(buf, 'xlsx');
              return spreadsheetResult(structured, filePath, output);
            }
          } catch {
            /* fall through */
          }
          try {
            const structured = readSpreadsheetStructuredFromBuffer(buf, 'xlsx');
            return spreadsheetResult(structured, filePath, output);
          } catch {
            /* fall through */
          }
        }
        throw err;
      }
    }

    // PPTX
    if (format === 'pptx') {
      const structured = filePath
        ? await readPptxStructured(filePath)
        : await readPptxStructuredFromBuffer(buf!);
      return pptxResult(structured, filePath, output);
    }

    // Spreadsheets
    if (format === 'xlsx' || format === 'xls') {
      const structured = filePath
        ? await readSpreadsheetStructuredAsync(filePath)
        : readSpreadsheetStructuredFromBuffer(buf!, format);
      return spreadsheetResult(structured, filePath, output);
    }

    // PDF
    if (format === 'pdf') {
      const { text } = await extractPdfText(filePath ?? buf!);
      return finish('pdf', text, filePath, output, { paragraphs: text.split(/\n+/) });
    }

    // MD / HTML / CSV
    if (format === 'md' || format === 'html' || format === 'csv') {
      const raw = stripBom(
        filePath != null
          ? await readTextFile(filePath)
          : buf
            ? buf.toString('utf8')
            : typeof input === 'string'
              ? input
              : ''
      );
      if (format === 'md') {
        const parsed = parseMarkdown(raw);
        const blocks = markdownToBlocks(parsed.text);
        const flatParas = flattenBlocksToParagraphs(blocks);
        const paragraphs = flatParas.length > 0 ? flatParas : parsed.paragraphs;
        if (output === 'structured') {
          const structured = {
            format: 'md' as const,
            paragraphs,
            blocks,
            raw: parsed.text,
          };
          const payload = JSON.stringify(structured, null, 2);
          return {
            format: 'md',
            text: payload,
            charCount: payload.length,
            path: filePath,
            structured,
          };
        }
        return finish('md', parsed.text, filePath, output, {
          paragraphs,
          blocks,
          raw: parsed.text,
        });
      }
      if (format === 'html') {
        const rich = htmlToStructured(raw);
        const plain = parseHtml(raw);
        if (output === 'structured') {
          const structured = {
            format: 'html' as const,
            paragraphs: rich.paragraphs.length ? rich.paragraphs : plain.paragraphs,
            blocks: rich.blocks,
          };
          const payload = JSON.stringify(structured, null, 2);
          return {
            format: 'html',
            text: payload,
            charCount: payload.length,
            path: filePath,
            structured,
          };
        }
        return finish('html', plain.text, filePath, output, {
          paragraphs: rich.paragraphs.length ? rich.paragraphs : plain.paragraphs,
          blocks: rich.blocks,
        });
      }
      const parsed = parseCsv(raw);
      if (output === 'structured') {
        const structured = { format: 'csv' as const, rows: parsed.rows };
        const payload = JSON.stringify(structured, null, 2);
        return {
          format: 'csv',
          text: payload,
          charCount: payload.length,
          path: filePath,
          structured,
        };
      }
      return finish('csv', parsed.text, filePath, output, { rows: parsed.rows });
    }

    // OpenDocument / RTF / EPUB
    if (format === 'odt' || format === 'ods' || format === 'odp') {
      const parsedCanon = await parseOdfBuffer(buf!, format);
      return canonicalResult(parsedCanon, filePath, output);
    }

    if (format === 'rtf') {
      const raw = stripBom(
        filePath != null
          ? await readTextFile(filePath)
          : buf
            ? buf.toString('utf8')
            : typeof input === 'string'
              ? input
              : ''
      );
      const parsedCanon = parseRtf(raw);
      return canonicalResult(parsedCanon, filePath, output);
    }

    if (format === 'epub') {
      const parsedCanon = await parseEpub(buf!);
      return canonicalResult(parsedCanon, filePath, output);
    }

    throw new DocumentError('UNSUPPORTED_FORMAT', `Unsupported format: ${format}`);
  } catch (err) {
    if (err instanceof DocumentError) throw err;
    throw wrapError('PARSE_FAILED', 'Failed to parse document', err);
  }
}

function canonicalResult(
  parsed: ReturnType<typeof buildParsedCanonical>,
  filePath: string | undefined,
  output: string
): DocumentResult {
  const format = parsed.canonical.format;
  if (output === 'structured') {
    const payload = JSON.stringify(parsed.structured, null, 2);
    return {
      format,
      text: payload,
      charCount: payload.length,
      path: filePath,
      structured: parsed.structured,
      warnings: parsed.warnings,
    };
  }
  if (output === 'ast') {
    const ast: import('../types.js').DocumentAst =
      parsed.canonical.kind === 'flow'
        ? { type: 'flow', format, blocks: parsed.canonical.blocks }
        : parsed.canonical.kind === 'workbook'
          ? { type: 'workbook', format, sheets: parsed.canonical.sheets }
          : { type: 'presentation', format, slides: parsed.canonical.slides };
    return {
      format,
      text: '',
      charCount: 0,
      path: filePath,
      ast,
      warnings: parsed.warnings,
    };
  }
  if (output === 'chunks') {
    const chunks = chunkText(parsed.plainText);
    const payload = JSON.stringify(chunks, null, 2);
    return {
      format,
      text: payload,
      charCount: payload.length,
      path: filePath,
      chunks,
      warnings: parsed.warnings,
    };
  }
  return {
    format,
    text: parsed.plainText,
    charCount: parsed.plainText.length,
    path: filePath,
    warnings: parsed.warnings,
  };
}

function pptxResult(
  structured: Awaited<ReturnType<typeof readPptxStructured>>,
  filePath: string | undefined,
  output: string
): DocumentResult {
  if (output === 'structured') {
    const payload = JSON.stringify(structured, null, 2);
    return {
      format: 'pptx',
      text: payload,
      charCount: payload.length,
      path: filePath,
      structured,
    };
  }
  const text = pptxToText(structured);
  return finish('pptx', text, filePath, output, {
    slides: structured.slides,
    paragraphs: pptxToParagraphs(structured),
    blocks: structured.blocks,
  });
}

function spreadsheetResult(
  structured: Awaited<ReturnType<typeof readSpreadsheetStructuredAsync>>,
  filePath: string | undefined,
  output: string
): DocumentResult {
  if (output === 'structured') {
    const payload = JSON.stringify(structured, null, 2);
    return {
      format: structured.format,
      text: payload,
      charCount: payload.length,
      path: filePath,
      structured,
    };
  }
  const text = spreadsheetToText(structured);
  return finish(structured.format, text, filePath, output, {
    sheets: structured.sheets,
  });
}

function finish(
  format: string,
  text: string,
  filePath: string | undefined,
  output: string,
  shape: Record<string, unknown>
): DocumentResult {
  if (output === 'ast') {
    return {
      format,
      text: '',
      charCount: 0,
      path: filePath,
      ast: {
        type: 'flow',
        format,
        blocks: contentBlocksToFlow(
          (shape.blocks as import('../types.js').ContentBlock[] | undefined) ?? []
        ),
      },
    };
  }
  if (output === 'chunks') {
    const chunks = chunkText(text);
    const payload = JSON.stringify(chunks, null, 2);
    return {
      format,
      text: payload,
      charCount: payload.length,
      path: filePath,
      chunks,
    };
  }
  // structured already handled by callers when needed
  if (output === 'structured') {
    const structured = { format, ...shape };
    const payload = JSON.stringify(structured, null, 2);
    return {
      format,
      text: payload,
      charCount: payload.length,
      path: filePath,
      structured,
    };
  }
  return {
    format,
    text,
    charCount: text.length,
    path: filePath,
  };
}

export { convertDocToDocx } from './legacy-doc.js';
