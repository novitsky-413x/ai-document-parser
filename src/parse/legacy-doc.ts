import path from 'node:path';
import * as CFB from 'cfb';
import { DocumentError } from '../utils/errors.js';
import { readFileLimited } from '../utils/zip.js';
import { writeDocx } from '../write/docx.js';

const MIN_STRING_LEN = 3;

/** Drop FIB noise (CJK/Hangul runs) while keeping Cyrillic and Latin prose. */
function looksLikeProse(s: string): boolean {
  if (/^[\d][\d\s.,]{3,}$/.test(s)) return true;
  if (!/[A-Za-z\u0400-\u04FF]{2,}/.test(s)) return false;
  if (/[\u1100-\u11FF\u3130-\u318F\uAC00-\uD7AF\u4E00-\u9FFF]/.test(s)) return false;
  let ok = 0;
  let other = 0;
  for (const ch of s) {
    if (/[A-Za-z\u0400-\u04FF0-9\s.,:;!?()\u00AB\u00BB"'\u2014\-/%\u2116+]/.test(ch))
      ok++;
    else other++;
  }
  return other <= ok;
}

function extractUtf16Strings(buf: Buffer, minChars = MIN_STRING_LEN): string[] {
  const out: string[] = [];
  let i = 0;
  while (i < buf.length - 1) {
    const chars: string[] = [];
    const start = i;
    while (i < buf.length - 1) {
      const code = buf.readUInt16LE(i);
      const ok =
        code >= 0x20 &&
        code !== 0x7f &&
        (code < 0xd800 || code > 0xdfff) &&
        code !== 0xfffe &&
        code !== 0xffff;
      if (!ok) break;
      chars.push(String.fromCharCode(code));
      i += 2;
    }
    if (chars.length >= minChars) {
      const s = chars.join('').trim();
      if (s && looksLikeProse(s)) out.push(s);
    }
    if (i === start) i += 2;
  }
  return out;
}

function extractAsciiStrings(buf: Buffer, minLen = MIN_STRING_LEN): string[] {
  const out: string[] = [];
  const re = /[\x20-\x7e\u00a0-\u024f]{4,}/g;
  const latin = buf.toString('latin1');
  let m: RegExpExecArray | null;
  while ((m = re.exec(latin)) !== null) {
    const s = m[0].trim();
    if (s.length >= minLen && !/^W\0o\0r\0d/.test(s)) out.push(s);
  }
  return out;
}

function dedupeOrdered(parts: string[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const p of parts) {
    const key = p.replace(/\s+/g, ' ').trim();
    if (key.length < MIN_STRING_LEN || seen.has(key)) continue;
    seen.add(key);
    result.push(key);
  }
  return result;
}

function extractTextFromWordDocument(data: Buffer): string {
  const parts = dedupeOrdered([
    ...extractUtf16Strings(data),
    ...extractAsciiStrings(data),
  ]);
  if (parts.length === 0) {
    throw new DocumentError(
      'PARSE_FAILED',
      'Could not extract text from .doc (empty or encrypted file)'
    );
  }
  return parts.join('\n');
}

function parseCfbBuffer(buf: Buffer): Buffer {
  const cfb = CFB.read(new Uint8Array(buf), { type: 'array' });
  const entry = CFB.find(cfb, 'WordDocument');
  if (!entry?.content) {
    throw new DocumentError(
      'UNSUPPORTED_FORMAT',
      'File is not a Word 97–2003 .doc (missing WordDocument stream)'
    );
  }
  const content = entry.content;
  return Buffer.isBuffer(content)
    ? content
    : Buffer.from(content as ArrayLike<number>);
}

export async function readLegacyDocBuffer(buf: Buffer): Promise<{
  text: string;
  paragraphs: string[];
  source: 'ai-document-parser-doc-binary';
}> {
  const data = parseCfbBuffer(buf);
  const text = extractTextFromWordDocument(data);
  const paragraphs = text.split(/\n+/).filter((p) => p.trim().length > 0);
  return { text, paragraphs, source: 'ai-document-parser-doc-binary' };
}

export async function readLegacyDoc(filePath: string): Promise<{
  text: string;
  paragraphs: string[];
  source: 'ai-document-parser-doc-binary';
}> {
  const buf = await readFileLimited(filePath);
  return readLegacyDocBuffer(buf);
}

export async function convertDocToDocx(
  docPath: string,
  docxPath?: string
): Promise<{ docxPath: string; paragraphCount: number }> {
  const { text } = await readLegacyDoc(docPath);
  const out =
    docxPath ??
    path.join(
      path.dirname(docPath),
      `${path.basename(docPath, path.extname(docPath))}.docx`
    );
  await writeDocx(out, text);
  return {
    docxPath: out,
    paragraphCount: text.split(/\n+/).filter(Boolean).length,
  };
}
