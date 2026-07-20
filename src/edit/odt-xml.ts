import { DocumentError } from '../utils/errors.js';
import { decodeXmlEntities, escapeXml } from '../utils/xml.js';
import { loadZipBuffer, readFileLimited } from '../utils/zip.js';
import { writeFileAtomic } from '../utils/paths.js';
import {
  replaceVisibleTokenText,
  type VisibleTextToken,
} from './text-selection.js';
import { scanXml, xmlElementRanges } from './xml-scan.js';

function repeatedSpaceCount(tag: string): number {
  const raw =
    /\btext:c\s*=\s*"(\d+)"/i.exec(tag)?.[1] ??
    /\btext:c\s*=\s*'(\d+)'/i.exec(tag)?.[1];
  return raw ? Math.max(1, Number(raw)) : 1;
}

function visibleTokens(containerXml: string): VisibleTextToken[] {
  const tokens: VisibleTextToken[] = [];
  for (const piece of scanXml(containerXml)) {
    const raw = piece.raw;
    let decoded: string | null = null;
    if (piece.kind === 'text') {
      decoded = decodeXmlEntities(raw);
    } else if (!piece.closing && piece.name === 'text:s') {
      decoded = ' '.repeat(repeatedSpaceCount(raw));
    } else if (!piece.closing && piece.name === 'text:tab') {
      decoded = '\t';
    } else if (!piece.closing && piece.name === 'text:line-break') {
      decoded = '\n';
    }
    if (decoded !== null) {
      tokens.push({
        start: piece.start,
        end: piece.end,
        decoded,
      });
    }
  }
  return tokens;
}

function renderOdtText(text: string): string {
  let out = '';
  let plain = '';
  const flush = () => {
    if (!plain) return;
    out += escapeXml(plain);
    plain = '';
  };
  for (const char of text) {
    if (char === ' ') {
      flush();
      out += '<text:s/>';
    } else if (char === '\t') {
      flush();
      out += '<text:tab/>';
    } else if (char === '\n') {
      flush();
      out += '<text:line-break/>';
    } else {
      plain += char;
    }
  }
  flush();
  return out;
}

function replaceContainer(
  xml: string,
  oldString: string,
  newString: string,
  replaceAll: boolean
): { xml: string; count: number } {
  const tokens = visibleTokens(xml);
  if (!tokens.length) return { xml, count: 0 };
  const result = replaceVisibleTokenText(
    tokens,
    oldString,
    newString,
    replaceAll
  );
  if (!result.count) return { xml, count: 0 };

  let next = xml;
  for (let index = tokens.length - 1; index >= 0; index--) {
    if (result.texts[index] === tokens[index].decoded) continue;
    next =
      next.slice(0, tokens[index].start) +
      renderOdtText(result.texts[index]) +
      next.slice(tokens[index].end);
  }
  return { xml: next, count: result.count };
}

function replaceContentXml(
  xml: string,
  oldString: string,
  newString: string
): { xml: string; count: number } {
  let count = 0;
  let next = xml;
  const ranges = xmlElementRanges(xml, new Set(['text:p', 'text:h']));
  for (let index = ranges.length - 1; index >= 0; index--) {
    const range = ranges[index];
    const result = replaceContainer(
      xml.slice(range.start, range.end),
      oldString,
      newString,
      true
    );
    count += result.count;
    if (result.count) {
      next =
        next.slice(0, range.start) +
        result.xml +
        next.slice(range.end);
    }
  }
  return { xml: next, count };
}

export async function editOdtText(
  filePath: string,
  oldString: string,
  newString: string,
  replaceAll = false
): Promise<{ message: string; changes: number }> {
  if (!oldString) {
    throw new DocumentError('INVALID_ARGUMENT', 'oldString cannot be empty');
  }

  const zip = await loadZipBuffer(await readFileLimited(filePath));
  const part = zip.file('content.xml');
  if (!part) {
    throw new DocumentError('PARSE_FAILED', 'Invalid ODT: missing content.xml');
  }

  const xml = await part.async('string');
  const result = replaceContentXml(xml, oldString, newString);
  if (result.count === 0) {
    throw new DocumentError(
      'EDIT_TARGET_NOT_FOUND',
      'Text not found in odt. Parse with output:"structured" and refine oldString.'
    );
  }
  if (result.count > 1 && !replaceAll) {
    throw new DocumentError(
      'MULTIPLE_MATCHES',
      `Found ${result.count} matches. Narrow oldString or set replaceAll: true.`
    );
  }

  zip.file('content.xml', result.xml);
  await writeFileAtomic(filePath, await zip.generateAsync({ type: 'nodebuffer' }));
  return {
    message: `Replaced ${result.count} occurrence(s) in ODT: ${filePath}`,
    changes: result.count,
  };
}
