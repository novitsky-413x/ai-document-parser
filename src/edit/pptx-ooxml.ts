import { DocumentError } from '../utils/errors.js';
import { decodeXmlEntities, escapeXml } from '../utils/xml.js';
import { loadZipBuffer, readFileLimited } from '../utils/zip.js';
import { writeFileAtomic } from '../utils/paths.js';
import {
  replaceVisibleTokenText,
  type VisibleTextToken,
} from './text-selection.js';
import type JSZip from 'jszip';
import { scanXml, xmlElementRanges } from './xml-scan.js';

function slideNumber(name: string): number | null {
  const match = /^ppt\/slides\/slide(\d+)\.xml$/i.exec(name);
  return match ? Number(match[1]) : null;
}

function attr(tag: string, name: string): string | undefined {
  return (
    new RegExp(`\\b${name.replace(':', '\\:')}\\s*=\\s*"([^"]*)"`, 'i').exec(tag)?.[1] ??
    new RegExp(`\\b${name.replace(':', '\\:')}\\s*=\\s*'([^']*)'`, 'i').exec(tag)?.[1]
  );
}

function resolvePart(basePart: string, target: string): string {
  const baseDir = basePart.slice(0, basePart.lastIndexOf('/') + 1);
  const combined = target.startsWith('/')
    ? target.slice(1)
    : `${baseDir}${target}`;
  const parts: string[] = [];
  for (const segment of combined.replace(/\\/g, '/').split('/')) {
    if (!segment || segment === '.') continue;
    if (segment === '..') parts.pop();
    else parts.push(segment);
  }
  return parts.join('/');
}

function relationships(xml: string): Map<string, { target: string; type: string }> {
  const map = new Map<string, { target: string; type: string }>();
  for (const match of xml.matchAll(/<(?:\w+:)?Relationship\b[^>]*\/?>/gi)) {
    const id = attr(match[0], 'Id');
    const target = attr(match[0], 'Target');
    const type = attr(match[0], 'Type') ?? '';
    if (id && target) map.set(id, { target, type });
  }
  return map;
}

function fallbackSlidePaths(zip: JSZip): string[] {
  return Object.keys(zip.files)
    .filter((name) => slideNumber(name) !== null && !zip.files[name].dir)
    .sort(
      (left, right) =>
        (slideNumber(left) ?? 0) - (slideNumber(right) ?? 0)
    );
}

async function orderedSlidePaths(zip: JSZip): Promise<string[]> {
  const presentation = zip.file('ppt/presentation.xml');
  const rels = zip.file('ppt/_rels/presentation.xml.rels');
  if (!presentation || !rels) return fallbackSlidePaths(zip);

  const presentationXml = await presentation.async('string');
  const relMap = relationships(await rels.async('string'));
  const ordered: string[] = [];
  for (const match of presentationXml.matchAll(/<p:sldId\b[^>]*\/?>/gi)) {
    const relationshipId = attr(match[0], 'r:id');
    const relation = relationshipId ? relMap.get(relationshipId) : undefined;
    if (!relation) continue;
    const part = resolvePart('ppt/presentation.xml', relation.target);
    if (zip.file(part)) ordered.push(part);
  }
  return ordered.length ? ordered : fallbackSlidePaths(zip);
}

async function notesPartForSlide(
  zip: JSZip,
  slidePart: string
): Promise<string | undefined> {
  const slash = slidePart.lastIndexOf('/');
  const relsPart = `${slidePart.slice(0, slash + 1)}_rels/${slidePart.slice(slash + 1)}.rels`;
  const rels = zip.file(relsPart);
  if (!rels) return undefined;
  for (const relation of relationships(await rels.async('string')).values()) {
    if (!/\/notesSlide$/i.test(relation.type)) continue;
    const part = resolvePart(slidePart, relation.target);
    if (zip.file(part)) return part;
  }
  return undefined;
}

function textTokens(paragraphXml: string): VisibleTextToken[] {
  const tokens: VisibleTextToken[] = [];
  for (const range of xmlElementRanges(paragraphXml, new Set(['a:t']))) {
    const element = paragraphXml.slice(range.start, range.end);
    const decoded = scanXml(element)
      .filter((piece) => piece.kind === 'text')
      .map((piece) => decodeXmlEntities(piece.raw))
      .join('');
    tokens.push({
      start: range.start,
      end: range.end,
      decoded,
    });
  }
  for (const piece of scanXml(paragraphXml)) {
    if (
      piece.kind === 'tag' &&
      piece.name === 'a:t' &&
      piece.selfClosing
    ) {
      tokens.push({
        start: piece.start,
        end: piece.end,
        decoded: '',
      });
    }
  }
  return tokens.sort((left, right) => left.start - right.start);
}

function renderText(text: string): string {
  if (!text) return '<a:t/>';
  return `<a:t xml:space="preserve">${escapeXml(text)}</a:t>`;
}

function replaceParagraph(
  paragraphXml: string,
  oldString: string,
  newString: string,
  replaceAll: boolean
): { xml: string; count: number } {
  const tokens = textTokens(paragraphXml);
  if (!tokens.length) return { xml: paragraphXml, count: 0 };

  const result = replaceVisibleTokenText(
    tokens,
    oldString,
    newString,
    replaceAll
  );
  if (!result.count) return { xml: paragraphXml, count: 0 };

  let xml = paragraphXml;
  for (let index = tokens.length - 1; index >= 0; index--) {
    if (result.texts[index] === tokens[index].decoded) continue;
    xml =
      xml.slice(0, tokens[index].start) +
      renderText(result.texts[index]) +
      xml.slice(tokens[index].end);
  }
  return { xml, count: result.count };
}

function replacePart(
  xml: string,
  oldString: string,
  newString: string
): { xml: string; count: number } {
  let count = 0;
  let next = xml;
  const ranges = xmlElementRanges(xml, new Set(['a:p']));
  for (let index = ranges.length - 1; index >= 0; index--) {
    const range = ranges[index];
    const result = replaceParagraph(
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

export async function editPptxText(
  filePath: string,
  oldString: string,
  newString: string,
  replaceAll = false,
  slide?: number,
  includeNotes = false
): Promise<{ message: string; changes: number }> {
  if (!oldString) {
    throw new DocumentError('INVALID_ARGUMENT', 'oldString cannot be empty');
  }
  if (
    slide !== undefined &&
    (!Number.isInteger(slide) || slide < 1)
  ) {
    throw new DocumentError(
      'INVALID_ARGUMENT',
      'slide must be a positive one-based integer'
    );
  }

  const zip = await loadZipBuffer(await readFileLimited(filePath));
  const orderedSlides = await orderedSlidePaths(zip);
  if (slide !== undefined && slide > orderedSlides.length) {
    throw new DocumentError('INVALID_ARGUMENT', `Slide not found: ${slide}`);
  }
  const selectedSlides =
    slide === undefined ? orderedSlides : [orderedSlides[slide - 1]];
  const candidateSet = new Set(selectedSlides);
  if (includeNotes) {
    for (const slidePart of selectedSlides) {
      const notesPart = await notesPartForSlide(zip, slidePart);
      if (notesPart) candidateSet.add(notesPart);
    }
  }
  const candidates = [...candidateSet];

  let total = 0;
  const changed = new Map<string, string>();
  for (const name of candidates) {
    const xml = await zip.file(name)!.async('string');
    const result = replacePart(xml, oldString, newString);
    total += result.count;
    if (result.count) changed.set(name, result.xml);
  }

  if (total === 0) {
    throw new DocumentError(
      'EDIT_TARGET_NOT_FOUND',
      'Text not found in pptx. Parse with output:"structured" and refine oldString or slide.'
    );
  }
  if (total > 1 && !replaceAll) {
    throw new DocumentError(
      'MULTIPLE_MATCHES',
      `Found ${total} matches. Narrow oldString, set slide, or set replaceAll: true.`
    );
  }

  for (const [name, xml] of changed) zip.file(name, xml);
  await writeFileAtomic(filePath, await zip.generateAsync({ type: 'nodebuffer' }));
  return {
    message: `Replaced ${total} occurrence(s) in PPTX: ${filePath}`,
    changes: total,
  };
}
