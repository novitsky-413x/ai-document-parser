import JSZip from 'jszip';
import { DocumentError } from '../utils/errors.js';
import { decodeXmlEntities, escapeXml } from '../utils/xml.js';
import { loadZipBuffer, readFileLimited } from '../utils/zip.js';
import { writeFileAtomic } from '../utils/paths.js';
import type { DocxStructured } from '../types.js';
import { docxStructuredFromZip } from '../parse/docx-blocks.js';

const DOC_PARTS: Array<string | RegExp> = [
  'word/document.xml',
  /^word\/header\d*\.xml$/,
  /^word\/footer\d*\.xml$/,
];

export { escapeXml } from '../utils/xml.js';

function isDocPart(name: string): boolean {
  return DOC_PARTS.some((p) =>
    typeof p === 'string' ? name === p : p.test(name)
  );
}

async function loadZip(filePath: string): Promise<JSZip> {
  const buf = await readFileLimited(filePath);
  return loadZipBuffer(buf);
}

/** One <w:t> element in a paragraph: its span within the paragraph XML plus decoded text. */
type TextToken = { start: number; end: number; decoded: string };

function extractTextTokens(paraXml: string): TextToken[] {
  const tokens: TextToken[] = [];
  const wtRe = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:t(?:\s[^>]*)?\/>/g;
  let m: RegExpExecArray | null;
  while ((m = wtRe.exec(paraXml)) !== null) {
    tokens.push({
      start: m.index,
      end: m.index + m[0].length,
      decoded: m[1] != null ? decodeXmlEntities(m[1]) : '',
    });
  }
  return tokens;
}

/**
 * Apply one replacement over decoded token texts. The insert lands only in the
 * first overlapping token; the matched span is cleared from the rest.
 */
function applyReplacement(
  texts: string[],
  start: number,
  end: number,
  insert: string
): string[] {
  let pos = 0;
  let inserted = false;
  return texts.map((decoded) => {
    const tokStart = pos;
    const tokEnd = pos + decoded.length;
    pos = tokEnd;
    if (tokEnd <= start || tokStart >= end) return decoded;
    const localStart = Math.max(0, start - tokStart);
    const localEnd = Math.min(decoded.length, end - tokStart);
    const before = decoded.slice(0, localStart);
    const after = decoded.slice(localEnd);
    const middle = inserted ? '' : insert;
    inserted = true;
    return before + middle + after;
  });
}

function renderTextToken(decoded: string): string {
  if (!decoded) return '<w:t/>';
  return `<w:t xml:space="preserve">${escapeXml(decoded)}</w:t>`;
}

function replaceInParagraphXml(
  paraXml: string,
  oldString: string,
  newString: string,
  replaceAll: boolean
): { xml: string; count: number } {
  const tokens = extractTextTokens(paraXml);
  if (!tokens.length) return { xml: paraXml, count: 0 };

  const combined = tokens.map((t) => t.decoded).join('');
  if (!combined.includes(oldString)) return { xml: paraXml, count: 0 };

  let texts = tokens.map((t) => t.decoded);
  let count = 0;
  let searchFrom = 0;

  if (replaceAll) {
    while (true) {
      const text = texts.join('');
      const idx = text.indexOf(oldString, searchFrom);
      if (idx === -1) break;
      texts = applyReplacement(texts, idx, idx + oldString.length, newString);
      count++;
      searchFrom = idx + newString.length;
    }
  } else {
    const idx = combined.indexOf(oldString);
    if (idx === -1) return { xml: paraXml, count: 0 };
    texts = applyReplacement(texts, idx, idx + oldString.length, newString);
    count = 1;
  }

  // Splice changed <w:t> elements back in place (right to left keeps offsets
  // valid) so hyperlinks, fields, bookmarks and run properties are untouched.
  let xml = paraXml;
  for (let i = tokens.length - 1; i >= 0; i--) {
    if (texts[i] === tokens[i].decoded) continue;
    xml = xml.slice(0, tokens[i].start) + renderTextToken(texts[i]) + xml.slice(tokens[i].end);
  }
  return { xml, count };
}

function replaceInDocumentXml(
  xml: string,
  oldString: string,
  newString: string,
  replaceAll: boolean
): { xml: string; count: number } {
  const paraRegex = /<w:p[\s\S]*?<\/w:p>/g;
  let total = 0;
  const result = xml.replace(paraRegex, (para) => {
    const { xml: next, count } = replaceInParagraphXml(
      para,
      oldString,
      newString,
      replaceAll
    );
    total += count;
    return next;
  });
  return { xml: result, count: total };
}

export async function editDocx(
  filePath: string,
  oldString: string,
  newString: string,
  replaceAll: boolean
): Promise<{ message: string; changes: number }> {
  if (!oldString) {
    throw new DocumentError('INVALID_ARGUMENT', 'oldString cannot be empty');
  }

  const zip = await loadZip(filePath);
  let total = 0;

  for (const name of Object.keys(zip.files)) {
    if (!isDocPart(name) || zip.files[name].dir) continue;
    const xml = await zip.file(name)!.async('string');
    const { xml: next, count } = replaceInDocumentXml(
      xml,
      oldString,
      newString,
      replaceAll
    );
    total += count;
    zip.file(name, next);
  }

  if (total === 0) {
    throw new DocumentError(
      'EDIT_TARGET_NOT_FOUND',
      'Text not found in docx. Parse with output:"structured" and refine oldString.'
    );
  }
  if (total > 1 && !replaceAll) {
    throw new DocumentError(
      'MULTIPLE_MATCHES',
      `Found ${total} matches. Narrow oldString or set replaceAll: true.`
    );
  }

  const out = await zip.generateAsync({ type: 'nodebuffer' });
  await writeFileAtomic(filePath, out);
  return {
    message: `DOCX updated: ${filePath} (${total} replacement(s); formatting preserved)`,
    changes: total,
  };
}

export async function appendDocxParagraphs(
  filePath: string,
  paragraphs: string[]
): Promise<{ message: string; changes: number }> {
  const zip = await loadZip(filePath);
  const docFile = zip.file('word/document.xml');
  if (!docFile) {
    throw new DocumentError('PARSE_FAILED', 'Invalid docx: missing word/document.xml');
  }

  let xml = await docFile.async('string');
  const blocks = paragraphs
    .map(
      (p) =>
        `<w:p><w:r><w:t xml:space="preserve">${escapeXml(p || ' ')}</w:t></w:r></w:p>`
    )
    .join('');

  if (!xml.includes('</w:body>')) {
    throw new DocumentError('PARSE_FAILED', 'Could not find w:body in document.xml');
  }
  // sectPr must remain the last child of w:body.
  const sectPrIdx = xml.search(/<w:sectPr[\s>]/);
  if (sectPrIdx >= 0) {
    xml = xml.slice(0, sectPrIdx) + blocks + xml.slice(sectPrIdx);
  } else {
    xml = xml.replace('</w:body>', `${blocks}</w:body>`);
  }

  zip.file('word/document.xml', xml);
  const out = await zip.generateAsync({ type: 'nodebuffer' });
  await writeFileAtomic(filePath, out);
  return {
    message: `DOCX: appended ${paragraphs.length} paragraph(s) to ${filePath}`,
    changes: paragraphs.length,
  };
}

export async function readDocxStructured(
  filePath: string
): Promise<DocxStructured> {
  const zip = await loadZip(filePath);
  if (!zip.file('word/document.xml')) {
    throw new DocumentError('PARSE_FAILED', 'Invalid docx: missing word/document.xml');
  }
  return docxStructuredFromZip(zip);
}

export async function readDocxStructuredFromBuffer(
  buf: Buffer
): Promise<DocxStructured> {
  const zip = await loadZipBuffer(buf);
  if (!zip.file('word/document.xml')) {
    throw new DocumentError('PARSE_FAILED', 'Invalid docx: missing word/document.xml');
  }
  return docxStructuredFromZip(zip);
}
