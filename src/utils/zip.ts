import JSZip from 'jszip';
import { DocumentError } from './errors.js';
import { DEFAULT_LIMITS, assertMaxBytes, type ResourceLimits } from './limits.js';

export async function loadZipBuffer(
  buf: Buffer,
  limits: ResourceLimits = DEFAULT_LIMITS
): Promise<JSZip> {
  assertMaxBytes(buf.length, limits.maxFileBytes);
  const zip = await JSZip.loadAsync(buf);
  const entries = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  if (entries.length > limits.maxZipEntries) {
    throw new DocumentError(
      'PARSE_FAILED',
      `ZIP exceeds maximum entry count (${entries.length} > ${limits.maxZipEntries})`
    );
  }
  let expanded = 0;
  for (const name of entries) {
    const file = zip.file(name);
    if (!file) continue;
    const meta = (file as unknown as { _data?: { uncompressedSize?: number } })._data;
    const size = meta?.uncompressedSize;
    if (typeof size !== 'number' || size < 0) {
      // Fail closed: unknown uncompressed size must not bypass bomb accounting.
      throw new DocumentError(
        'PARSE_FAILED',
        `ZIP entry "${name}" has unknown uncompressed size`
      );
    }
    expanded += size;
    if (expanded > limits.maxZipExpandedBytes) {
      throw new DocumentError(
        'PARSE_FAILED',
        `ZIP expanded size exceeds limit (${expanded} > ${limits.maxZipExpandedBytes} bytes)`
      );
    }
  }
  return zip;
}

export async function readFileLimited(
  path: string,
  limits: ResourceLimits = DEFAULT_LIMITS
): Promise<Buffer> {
  const { readFile, stat } = await import('node:fs/promises');
  const st = await stat(path);
  assertMaxBytes(st.size, limits.maxFileBytes);
  return readFile(path);
}
