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
  if (s.length > 512 && !/[\\/]/.test(s) && !/\.\w{1,8}$/.test(s)) return true;
  return false;
}

export async function ensureParentDir(filePath: string): Promise<void> {
  const { mkdir } = await import('node:fs/promises');
  await mkdir(path.dirname(filePath), { recursive: true });
}
