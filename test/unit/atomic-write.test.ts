import assert from 'node:assert/strict';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { test } from 'node:test';
import { FileLockedError, isFileLockError, writeFileAtomic } from '../../src/utils/paths.js';

test('isFileLockError recognizes EPERM/EACCES/EBUSY', () => {
  assert.equal(isFileLockError({ code: 'EPERM' }), true);
  assert.equal(isFileLockError({ code: 'EACCES' }), true);
  assert.equal(isFileLockError({ code: 'EBUSY' }), true);
  assert.equal(isFileLockError({ code: 'ENOENT' }), false);
  assert.equal(isFileLockError('nope'), false);
});

test('writeFileAtomic writes via a temp file and leaves the final content', async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), 'adp-atomic-'));
  try {
    const target = path.join(dir, 'out.docx');
    await writeFileAtomic(target, 'hello');
    assert.equal(await readFile(target, 'utf8'), 'hello');
    await writeFileAtomic(target, 'world');
    assert.equal(await readFile(target, 'utf8'), 'world');
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('writeFileAtomic retries a locked rename then succeeds', async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), 'adp-atomic-'));
  try {
    const target = path.join(dir, 'locked.docx');
    await writeFile(target, 'old');
    let attempts = 0;
    const delays: number[] = [];
    await writeFileAtomic(target, 'new', {
      backoffMs: [1, 1, 1],
      sleep: async (ms) => {
        delays.push(ms);
      },
      rename: async (from, to) => {
        attempts += 1;
        if (attempts < 3) {
          const err = Object.assign(new Error('EPERM, rename'), { code: 'EPERM' });
          throw err;
        }
        const { rename } = await import('node:fs/promises');
        await rename(from, to);
      },
    });
    assert.equal(attempts, 3);
    assert.deepEqual(delays, [1, 1]);
    assert.equal(await readFile(target, 'utf8'), 'new');
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test('writeFileAtomic surfaces a clear FileLockedError after retries', async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), 'adp-atomic-'));
  try {
    const target = path.join(dir, 'still-locked.docx');
    await assert.rejects(
      () =>
        writeFileAtomic(target, 'x', {
          backoffMs: [1],
          sleep: async () => undefined,
          rename: async () => {
            throw Object.assign(new Error('busy'), { code: 'EBUSY' });
          },
        }),
      (err: unknown) => {
        assert.ok(err instanceof FileLockedError);
        assert.match(err.message, /probably open in another program/);
        assert.match(err.message, /different path/);
        assert.match(err.message, /still-locked\.docx/);
        assert.equal(err.code, 'EBUSY');
        return true;
      }
    );
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
