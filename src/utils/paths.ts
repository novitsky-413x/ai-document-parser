import path from 'node:path';

export function extOf(input: string): string {
  return path.extname(input).toLowerCase();
}

export function isPathLike(input: unknown): input is string {
  return typeof input === 'string' && !looksLikeRawContent(input);
}

/** Heuristic: short strings without path separators may still be paths (e.g. "a.docx"). */
function looksLikeRawContent(s: string): boolean {
  if (s.includes('\n') || s.includes('\r')) return true;
  if (/^#{1,6}\s/.test(s)) return true;
  if (/^\s*<\w+/i.test(s)) return true;
  if (/^\s*[-*+]\s+/.test(s)) return true;
  if (/^\s*\d+\.\s+/.test(s)) return true;
  if (s.length > 512 && !/[\\/]/.test(s) && !/\.\w{1,8}$/.test(s)) return true;
  return false;
}

export async function ensureParentDir(filePath: string): Promise<void> {
  const { mkdir } = await import('node:fs/promises');
  await mkdir(path.dirname(filePath), { recursive: true });
}

/**
 * Write via a sibling temp file + rename so a crash mid-write cannot leave a
 * truncated/corrupt document at the target path.
 */
export async function writeFileAtomic(
  filePath: string,
  data: string | Uint8Array
): Promise<void> {
  const { writeFile, rename, rm } = await import('node:fs/promises');
  const tmp = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  try {
    await writeFile(tmp, data);
    await rename(tmp, filePath);
  } catch (err) {
    await rm(tmp, { force: true }).catch(() => undefined);
    throw err;
  }
}
