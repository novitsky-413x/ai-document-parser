import { DocumentError, wrapError } from '../utils/errors.js';
import type { CanonicalFlowBlock, CanonicalInline, ParsedCanonical } from '../ir/types.js';
import { buildParsedCanonical } from '../ir/projections.js';

function decodeRtfEscapes(text: string): string {
  return text
    .replace(/\\'([0-9a-fA-F]{2})/g, (_, hex) =>
      String.fromCharCode(parseInt(hex, 16))
    )
    .replace(/\\par\b/g, '\n')
    .replace(/\\tab\b/g, '\t')
    .replace(/\\line\b/g, '\n')
    .replace(/\\emdash\b/g, '—')
    .replace(/\\endash\b/g, '–')
    .replace(/\\bullet\b/g, '•')
    .replace(/\\~ /g, ' ')
    .replace(/\\-/g, '')
    .replace(/\\[a-z]+\d* ?/gi, '')
    .replace(/[{}]/g, '');
}

function parseRtfInline(segment: string): CanonicalInline[] {
  const runs: CanonicalInline[] = [];
  let bold = false;
  let italic = false;
  let strike = false;
  const parts = segment.split(/(\\b\b|\\b0|\\i\b|\\i0|\\strike\b|\\strike0)/);
  let buf = '';
  const flush = () => {
    if (!buf) return;
    const text = decodeRtfEscapes(buf);
    buf = '';
    if (text) runs.push({ text, bold, italic, strike });
  };
  for (const part of parts) {
    if (part === '\\b') {
      flush();
      bold = true;
    } else if (part === '\\b0') {
      flush();
      bold = false;
    } else if (part === '\\i') {
      flush();
      italic = true;
    } else if (part === '\\i0') {
      flush();
      italic = false;
    } else if (part === '\\strike') {
      flush();
      strike = true;
    } else if (part === '\\strike0') {
      flush();
      strike = false;
    } else {
      buf += part;
    }
  }
  flush();
  return runs.filter((r) => r.text);
}

export function parseRtf(text: string): ParsedCanonical {
  try {
    const normalized = text.replace(/\r\n/g, '\n');
    if (!/^\{\\rtf/i.test(normalized.trimStart())) {
      throw new DocumentError('PARSE_FAILED', 'Invalid RTF: missing {\\rtf header');
    }
    const body = normalized.replace(/^\{\\rtf[\s\S]*?\\viewkind\d+/i, '');
    const paragraphs = body.split(/\\par\b/).map((p) => p.trim()).filter(Boolean);
    const blocks: CanonicalFlowBlock[] = paragraphs.map((p) => ({
      type: 'paragraph',
      runs: parseRtfInline(p),
    }));
    return buildParsedCanonical(
      { kind: 'flow', format: 'rtf', blocks, raw: text },
      [{ code: 'RTF_BASIC_PARSE', message: 'RTF parsed with basic control-word extraction.' }]
    );
  } catch (err) {
    if (err instanceof DocumentError) throw err;
    throw wrapError('PARSE_FAILED', 'Failed to parse RTF', err);
  }
}

export function emitRtfFromFlow(blocks: CanonicalFlowBlock[]): string {
  const parts = ['{\\rtf1\\ansi'];
  for (const b of blocks) {
    if (b.type === 'heading') {
      parts.push('\\b ');
      parts.push(escapeRtf(runsToPlain(b.runs)));
      parts.push('\\b0\\par ');
    } else if (b.type === 'paragraph') {
      for (const r of b.runs) {
        if (r.bold) parts.push('\\b ');
        if (r.italic) parts.push('\\i ');
        if (r.strike) parts.push('\\strike ');
        parts.push(escapeRtf(r.text));
        if (r.strike) parts.push('\\strike0 ');
        if (r.italic) parts.push('\\i0 ');
        if (r.bold) parts.push('\\b0 ');
      }
      parts.push('\\par ');
    } else if (b.type === 'code') {
      parts.push(escapeRtf(b.text));
      parts.push('\\par ');
    }
  }
  parts.push('}');
  return parts.join('');
}

function runsToPlain(runs: CanonicalInline[]): string {
  return runs.map((r) => r.text).join('');
}

function escapeRtf(s: string): string {
  return s.replace(/\\/g, '\\\\').replace(/\{/g, '\\{').replace(/\}/g, '\\}');
}
