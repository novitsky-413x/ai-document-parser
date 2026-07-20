import JSZip from 'jszip';
import { DocumentError, wrapError } from '../utils/errors.js';
import { loadZipBuffer } from '../utils/zip.js';
import { htmlToBlocks } from './html-blocks.js';
import type { ParsedCanonical } from '../ir/types.js';
import { buildParsedCanonical, contentBlocksToFlow } from '../ir/projections.js';

function resolveEpubHref(opfDir: string, href: string): string {
  let decoded = href;
  try {
    decoded = decodeURIComponent(href);
  } catch {
    // keep raw href if percent-decoding fails
  }
  const parts: string[] = [];
  for (const seg of (opfDir + decoded).split('/')) {
    if (!seg || seg === '.') continue;
    if (seg === '..') parts.pop();
    else parts.push(seg);
  }
  return parts.join('/');
}

export async function parseEpub(buf: Buffer): Promise<ParsedCanonical> {
  try {
    const zip = await loadZipBuffer(buf);
    const container = zip.file('META-INF/container.xml');
    if (!container) throw new DocumentError('PARSE_FAILED', 'Invalid EPUB: missing container.xml');
    const containerXml = await container.async('string');
    const rootfile = /full-path="([^"]+)"/i.exec(containerXml)?.[1];
    if (!rootfile) throw new DocumentError('PARSE_FAILED', 'Invalid EPUB: missing rootfile');

    const opfFile = zip.file(rootfile);
    if (!opfFile) throw new DocumentError('PARSE_FAILED', 'Invalid EPUB: missing OPF');
    const opf = await opfFile.async('string');
    const opfDir = rootfile.includes('/') ? rootfile.slice(0, rootfile.lastIndexOf('/') + 1) : '';

    const manifest = new Map<string, string>();
    const itemRe = /<item\b[^>]*\/?>/gi;
    let im: RegExpExecArray | null;
    while ((im = itemRe.exec(opf)) !== null) {
      const tag = im[0];
      const id = /\bid="([^"]+)"/i.exec(tag)?.[1];
      const href = /\bhref="([^"]+)"/i.exec(tag)?.[1];
      if (id && href) manifest.set(id, resolveEpubHref(opfDir, href));
    }

    const spine: string[] = [];
    const itemrefRe = /<itemref\b[^>]*idref="([^"]+)"[^>]*\/?>/gi;
    let sr: RegExpExecArray | null;
    while ((sr = itemrefRe.exec(opf)) !== null) {
      const href = manifest.get(sr[1]);
      if (href) spine.push(href);
    }

    const blocks: import('../ir/types.js').CanonicalFlowBlock[] = [];
    for (const href of spine) {
      const part = zip.file(href);
      if (!part) continue;
      const html = await part.async('string');
      blocks.push(...contentBlocksToFlow(htmlToBlocks(html)));
    }

    return buildParsedCanonical(
      { kind: 'flow', format: 'epub', blocks },
      [{ code: 'EPUB_SPINE_PARSE', message: 'EPUB parsed from spine XHTML content.' }]
    );
  } catch (err) {
    if (err instanceof DocumentError) throw err;
    throw wrapError('PARSE_FAILED', 'Failed to parse EPUB', err);
  }
}

export async function emitEpubFromFlow(
  blocks: import('../ir/types.js').CanonicalFlowBlock[],
  title = 'Document'
): Promise<Buffer> {
  const { emitHtmlFromBlocks } = await import('../convert/blocks.js');
  const { flowToContentBlocks } = await import('../ir/projections.js');
  const htmlBody = emitHtmlFromBlocks(flowToContentBlocks(blocks))
    .replace(/^[\s\S]*?<body>\s*/i, '')
    .replace(/\s*<\/body>[\s\S]*$/i, '');

  const zip = new JSZip();
  zip.file(
    'mimetype',
    'application/epub+zip',
    { compression: 'STORE' }
  );
  zip.file(
    'META-INF/container.xml',
    `<?xml version="1.0"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>`
  );
  zip.file(
    'OEBPS/content.opf',
    `<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="uid">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:title>${escapeXml(title)}</dc:title>
    <dc:identifier id="uid">urn:uuid:ai-document-parser</dc:identifier>
  </metadata>
  <manifest>
    <item id="chapter" href="chapter.xhtml" media-type="application/xhtml+xml"/>
    <item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>
  </manifest>
  <spine>
    <itemref idref="chapter"/>
  </spine>
</package>`
  );
  zip.file(
    'OEBPS/chapter.xhtml',
    `<?xml version="1.0" encoding="UTF-8"?>
<html xmlns="http://www.w3.org/1999/xhtml"><head><title>${escapeXml(title)}</title></head><body>${htmlBody}</body></html>`
  );
  zip.file(
    'OEBPS/nav.xhtml',
    `<?xml version="1.0" encoding="UTF-8"?>
<html xmlns="http://www.w3.org/1999/xhtml"><head><title>Nav</title></head><body><nav><ol><li><a href="chapter.xhtml">${escapeXml(title)}</a></li></ol></nav></body></html>`
  );
  return zip.generateAsync({ type: 'nodebuffer' });
}

function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
