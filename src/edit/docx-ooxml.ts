import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { DocumentError } from '../utils/errors.js';
import type { DocxStructured } from '../types.js';

const DOC_PARTS: Array<string | RegExp> = [
  'word/document.xml',
  /^word\/header\d*\.xml$/,
  /^word\/footer\d*\.xml$/,
];

export function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function decodeXmlEntities(s: string): string {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function paragraphTextsFromXml(xml: string): string[] {
  const paragraphs: string[] = [];
  const paraRegex = /<w:p[\s\S]*?<\/w:p>/g;
  let pm: RegExpExecArray | null;
  while ((pm = paraRegex.exec(xml)) !== null) {
    const wtRegex = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
    const parts: string[] = [];
    let tm: RegExpExecArray | null;
    while ((tm = wtRegex.exec(pm[0])) !== null) {
      parts.push(decodeXmlEntities(tm[1]));
    }
    const text = parts.join('').trim();
    if (text) paragraphs.push(text);
  }
  return paragraphs;
}

function isDocPart(name: string): boolean {
  return DOC_PARTS.some((p) =>
    typeof p === 'string' ? name === p : p.test(name)
  );
}

async function loadZip(filePath: string): Promise<JSZip> {
  const buf = await readFile(filePath);
  return JSZip.loadAsync(buf);
}

function replaceInParagraphXml(
  paraXml: string,
  oldString: string,
  newString: string,
  replaceAll: boolean
): { xml: string; count: number } {
  const wtRegex = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
  const runs: { full: string; text: string }[] = [];
  let m: RegExpExecArray | null;
  while ((m = wtRegex.exec(paraXml)) !== null) {
    runs.push({ full: m[0], text: m[1] });
  }
  if (runs.length === 0) {
    return { xml: paraXml, count: 0 };
  }

  const combined = runs.map((r) => r.text).join('');
  if (!combined.includes(oldString)) {
    return { xml: paraXml, count: 0 };
  }

  let count = 0;
  let updatedCombined = combined;
  if (replaceAll) {
    const parts = combined.split(oldString);
    count = parts.length - 1;
    updatedCombined = parts.join(newString);
  } else {
    const idx = combined.indexOf(oldString);
    if (idx === -1) return { xml: paraXml, count: 0 };
    count = 1;
    updatedCombined =
      combined.slice(0, idx) +
      newString +
      combined.slice(idx + oldString.length);
  }

  const escaped = escapeXml(updatedCombined);
  const newRun = `<w:r><w:t xml:space="preserve">${escaped}</w:t></w:r>`;
  const pPrMatch = paraXml.match(/<w:pPr[\s\S]*?<\/w:pPr>/);
  const pPr = pPrMatch ? pPrMatch[0] : '';
  return { xml: `<w:p>${pPr}${newRun}</w:p>`, count };
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

  if (total === 0 && xml.includes(oldString)) {
    if (replaceAll) {
      const parts = xml.split(oldString);
      if (parts.length > 1) {
        return { xml: parts.join(escapeXml(newString)), count: parts.length - 1 };
      }
    }
    const idx = xml.indexOf(oldString);
    if (idx !== -1) {
      return {
        xml:
          xml.slice(0, idx) +
          escapeXml(newString) +
          xml.slice(idx + oldString.length),
        count: 1,
      };
    }
  }

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

  const { writeFile } = await import('node:fs/promises');
  const out = await zip.generateAsync({ type: 'nodebuffer' });
  await writeFile(filePath, out);
  return {
    message: `DOCX updated: ${filePath} (${total} replacement(s); formatting partially preserved)`,
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
  xml = xml.replace('</w:body>', `${blocks}</w:body>`);

  zip.file('word/document.xml', xml);
  const { writeFile } = await import('node:fs/promises');
  const out = await zip.generateAsync({ type: 'nodebuffer' });
  await writeFile(filePath, out);
  return {
    message: `DOCX: appended ${paragraphs.length} paragraph(s) to ${filePath}`,
    changes: paragraphs.length,
  };
}

export async function readDocxStructured(
  filePath: string
): Promise<DocxStructured> {
  const zip = await loadZip(filePath);
  const docFile = zip.file('word/document.xml');
  if (!docFile) {
    throw new DocumentError('PARSE_FAILED', 'Invalid docx: missing word/document.xml');
  }

  const xml = await docFile.async('string');
  return { format: 'docx', paragraphs: paragraphTextsFromXml(xml) };
}

export async function readDocxStructuredFromBuffer(
  buf: Buffer
): Promise<DocxStructured> {
  const zip = await JSZip.loadAsync(buf);
  const docFile = zip.file('word/document.xml');
  if (!docFile) {
    throw new DocumentError('PARSE_FAILED', 'Invalid docx: missing word/document.xml');
  }
  const xml = await docFile.async('string');
  return { format: 'docx', paragraphs: paragraphTextsFromXml(xml) };
}
