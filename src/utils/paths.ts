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

export function isFileLockError(err: unknown): boolean {
  const code = (err as NodeJS.ErrnoException | undefined)?.code;
  return code === 'EPERM' || code === 'EACCES' || code === 'EBUSY';
}

export class FileLockedError extends Error {
  code: string;
  override cause?: unknown;
  constructor(filePath: string, cause?: unknown) {
    const code = (cause as NodeJS.ErrnoException | undefined)?.code ?? 'EPERM';
    super(
      `Cannot replace "${filePath}": the file is probably open in another program (${code}). Close it and retry, or write to a different path.`
    );
    this.name = 'FileLockedError';
    this.code = code;
    this.cause = cause;
  }
}

const DEFAULT_BACKOFF_MS = [50, 150, 400, 800, 1600];

export type WriteFileAtomicOptions = {
  backoffMs?: number[];
  sleep?: (ms: number) => Promise<void>;
  rename?: (from: string, to: string) => Promise<void>;
};

/**
 * Write via a sibling temp file + rename so a crash mid-write cannot leave a
 * truncated/corrupt document at the target path. On EPERM/EACCES/EBUSY (typical
 * when the target is open in Word), retry with backoff, then throw FileLockedError.
 */
export async function writeFileAtomic(
  filePath: string,
  data: string | Uint8Array,
  options?: WriteFileAtomicOptions
): Promise<void> {
  const { writeFile, rename, rm } = await import('node:fs/promises');
  const backoff = options?.backoffMs ?? DEFAULT_BACKOFF_MS;
  const sleep = options?.sleep ?? ((ms: number) => new Promise((r) => setTimeout(r, ms)));
  const doRename = options?.rename ?? rename;
  const tmp = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  try {
    await writeFile(tmp, data);
    let lastErr: unknown;
    for (let i = 0; i <= backoff.length; i++) {
      try {
        await doRename(tmp, filePath);
        return;
      } catch (err) {
        lastErr = err;
        if (!isFileLockError(err) || i === backoff.length) break;
        await sleep(backoff[i]!);
      }
    }
    if (isFileLockError(lastErr)) throw new FileLockedError(filePath, lastErr);
    throw lastErr;
  } catch (err) {
    await rm(tmp, { force: true }).catch(() => undefined);
    throw err;
  }
}
