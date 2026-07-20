import { DocumentError } from '../utils/errors.js';
import { decodeXmlEntities, escapeXml } from '../utils/xml.js';
import { loadZipBuffer, readFileLimited } from '../utils/zip.js';
import { writeFileAtomic } from '../utils/paths.js';
import {
  replaceVisibleTokenText,
  type VisibleTextToken,
} from './text-selection.js';
import { scanXml, xmlElementRanges } from './xml-scan.js';

function attr(tag: string, name: string): string | undefined {
  const double = new RegExp(`\\b${name}\\s*=\\s*"([^"]*)"`, 'i').exec(tag)?.[1];
  if (double !== undefined) return double;
  return new RegExp(`\\b${name}\\s*=\\s*'([^']*)'`, 'i').exec(tag)?.[1];
}

function resolveHref(baseDir: string, href: string): string {
  let decoded = href.split('#')[0];
  try {
    decoded = decodeURIComponent(decoded);
  } catch {
    // Keep the literal package path.
  }
  const parts: string[] = [];
  for (const part of `${baseDir}${decoded}`.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') parts.pop();
    else parts.push(part);
  }
  return parts.join('/');
}

function decodeXhtmlText(text: string): string {
  return decodeXmlEntities(text.replace(/&nbsp;/gi, '\u00a0'));
}

function visibleTokens(containerXml: string): VisibleTextToken[] {
  const tokens: VisibleTextToken[] = [];
  const skipped: string[] = [];
  for (const piece of scanXml(containerXml)) {
    const raw = piece.raw;
    if (piece.kind === 'tag') {
      const localName = piece.name?.includes(':')
        ? piece.name.slice(piece.name.lastIndexOf(':') + 1)
        : piece.name;
      if (piece.closing && skipped.at(-1) === localName) {
        skipped.pop();
        continue;
      }
      if (
        !piece.closing &&
        !piece.selfClosing &&
        (localName === 'script' || localName === 'style' || localName === 'svg')
      ) {
        skipped.push(localName);
        continue;
      }
      if (!skipped.length && !piece.closing && localName === 'br') {
        tokens.push({
          start: piece.start,
          end: piece.end,
          decoded: '\n',
        });
      }
      continue;
    }
    if (!skipped.length) {
      tokens.push({
        start: piece.start,
        end: piece.end,
        decoded: decodeXhtmlText(raw),
      });
    }
  }
  return tokens;
}

function renderXhtmlText(text: string): string {
  return text
    .split('\n')
    .map((part) => escapeXml(part))
    .join('<br/>');
}

function replaceContainer(
  xml: string,
  oldString: string,
  newString: string
): { xml: string; count: number } {
  const tokens = visibleTokens(xml);
  if (!tokens.length) return { xml, count: 0 };
  const result = replaceVisibleTokenText(tokens, oldString, newString, true);
  if (!result.count) return { xml, count: 0 };

  let next = xml;
  for (let index = tokens.length - 1; index >= 0; index--) {
    if (result.texts[index] === tokens[index].decoded) continue;
    next =
      next.slice(0, tokens[index].start) +
      renderXhtmlText(result.texts[index]) +
      next.slice(tokens[index].end);
  }
  return { xml: next, count: result.count };
}

function replaceXhtml(
  xml: string,
  oldString: string,
  newString: string
): { xml: string; count: number } {
  let count = 0;
  let next = xml;
  const allRanges = xmlElementRanges(
    xml,
    new Set(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'li', 'td', 'th'])
  );
  // Avoid processing overlapping containers (for example nested list items or
  // a paragraph inside a table cell) twice. The innermost semantic block owns
  // its visible text.
  const ranges = allRanges.filter(
    (range) =>
      !allRanges.some(
        (other) =>
          other !== range &&
          other.start > range.start &&
          other.end < range.end
      )
  );
  for (let index = ranges.length - 1; index >= 0; index--) {
    const range = ranges[index];
    const result = replaceContainer(
      xml.slice(range.start, range.end),
      oldString,
      newString
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

async function spineParts(
  zip: Awaited<ReturnType<typeof loadZipBuffer>>
): Promise<string[]> {
  const containerPart = zip.file('META-INF/container.xml');
  if (!containerPart) {
    throw new DocumentError('PARSE_FAILED', 'Invalid EPUB: missing container.xml');
  }
  const containerXml = await containerPart.async('string');
  const rootTag = /<(?:\w+:)?rootfile\b[^>]*>/i.exec(containerXml)?.[0];
  const rootfile = rootTag ? attr(rootTag, 'full-path') : undefined;
  if (!rootfile) {
    throw new DocumentError('PARSE_FAILED', 'Invalid EPUB: missing rootfile');
  }

  const opfPart = zip.file(rootfile);
  if (!opfPart) throw new DocumentError('PARSE_FAILED', 'Invalid EPUB: missing OPF');
  const opf = await opfPart.async('string');
  const opfDir = rootfile.includes('/')
    ? rootfile.slice(0, rootfile.lastIndexOf('/') + 1)
    : '';

  const manifest = new Map<
    string,
    { path: string; mediaType: string; fallback?: string }
  >();
  for (const match of opf.matchAll(/<(?:\w+:)?item\b[^>]*\/?>/gi)) {
    const id = attr(match[0], 'id');
    const href = attr(match[0], 'href');
    const mediaType = attr(match[0], 'media-type') ?? '';
    const fallback = attr(match[0], 'fallback');
    if (id && href) {
      manifest.set(id, {
        path: resolveHref(opfDir, href),
        mediaType,
        ...(fallback ? { fallback } : {}),
      });
    }
  }

  const spine: string[] = [];
  for (const match of opf.matchAll(/<(?:\w+:)?itemref\b[^>]*\/?>/gi)) {
    let idref = attr(match[0], 'idref');
    const seen = new Set<string>();
    while (idref && !seen.has(idref)) {
      seen.add(idref);
      const item = manifest.get(idref);
      if (!item) break;
      if (item.mediaType.toLowerCase() === 'application/xhtml+xml') {
        spine.push(item.path);
        break;
      }
      idref = item.fallback;
    }
  }
  return spine;
}

export async function editEpubText(
  filePath: string,
  oldString: string,
  newString: string,
  replaceAll = false,
  chapter?: number | string
): Promise<{ message: string; changes: number }> {
  if (!oldString) {
    throw new DocumentError('INVALID_ARGUMENT', 'oldString cannot be empty');
  }

  const zip = await loadZipBuffer(await readFileLimited(filePath));
  const spine = await spineParts(zip);
  let candidates = spine;
  if (typeof chapter === 'number') {
    if (!Number.isInteger(chapter) || chapter < 1 || chapter > spine.length) {
      throw new DocumentError(
        'INVALID_ARGUMENT',
        `Chapter must be between 1 and ${spine.length}`
      );
    }
    candidates = [spine[chapter - 1]];
  } else if (typeof chapter === 'string') {
    if (!spine.includes(chapter)) {
      throw new DocumentError(
        'INVALID_ARGUMENT',
        `Chapter is not in the EPUB spine: ${chapter}`
      );
    }
    candidates = [chapter];
  }

  let total = 0;
  const changed = new Map<string, string>();
  for (const name of candidates) {
    const part = zip.file(name);
    if (!part) continue;
    const result = replaceXhtml(await part.async('string'), oldString, newString);
    total += result.count;
    if (result.count) changed.set(name, result.xml);
  }

  if (total === 0) {
    throw new DocumentError(
      'EDIT_TARGET_NOT_FOUND',
      'Text not found in EPUB spine content. Refine oldString or chapter.'
    );
  }
  if (total > 1 && !replaceAll) {
    throw new DocumentError(
      'MULTIPLE_MATCHES',
      `Found ${total} matches. Narrow oldString, set chapter, or set replaceAll: true.`
    );
  }

  for (const [name, xml] of changed) zip.file(name, xml);
  await writeFileAtomic(filePath, await zip.generateAsync({ type: 'nodebuffer' }));
  return {
    message: `Replaced ${total} occurrence(s) in EPUB: ${filePath}`,
    changes: total,
  };
}
