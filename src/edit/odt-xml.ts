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

/**
 * Serialize visible text for ODT.
 * Single U+0020 spaces stay as character data (ODF collapses only sequences);
 * runs of 2+ spaces / tabs / line breaks use explicit ODF elements.
 */
function renderOdtText(text: string): string {
  let out = '';
  let plain = '';
  const flush = () => {
    if (!plain) return;
    out += escapeXml(plain);
    plain = '';
  };
  for (let index = 0; index < text.length; ) {
    const char = text[index];
    if (char === ' ') {
      let count = 0;
      while (index < text.length && text[index] === ' ') {
        count++;
        index++;
      }
      if (count === 1) {
        plain += ' ';
      } else {
        flush();
        out += `<text:s text:c="${count}"/>`;
      }
      continue;
    }
    if (char === '\t') {
      flush();
      out += '<text:tab/>';
      index++;
      continue;
    }
    if (char === '\n') {
      flush();
      out += '<text:line-break/>';
      index++;
      continue;
    }
    plain += char;
    index++;
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

  // Splice only changed leaves. Whitespace leaves (`<text:s/>`, tabs, breaks)
  // become empty strings and drop out; text leaves get renderOdtText so single
  // spaces stay as character data (preserving surrounding spans/styles).
  let next = xml;
  for (let index = tokens.length - 1; index >= 0; index--) {
    if (result.texts[index] === tokens[index].decoded) continue;
    const replacement =
      result.texts[index].length === 0
        ? ''
        : renderOdtText(result.texts[index]);
    next =
      next.slice(0, tokens[index].start) +
      replacement +
      next.slice(tokens[index].end);
  }
  return { xml: next, count: result.count };
}

function replaceContentXml(
  xml: string,
  oldString: string,
  newString: string,
  replaceAll: boolean
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
      // Count every match inside the paragraph; editOdtText enforces replaceAll.
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
  // When replaceAll is false, callers must not write if count > 1. Still avoid
  // applying more than one paragraph mutation into `next` for the dry run that
  // only needs the count — keep mutated XML only when a single match exists or
  // replaceAll was requested.
  if (!replaceAll && count > 1) {
    return { xml, count };
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
  const result = replaceContentXml(xml, oldString, newString, replaceAll);
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
