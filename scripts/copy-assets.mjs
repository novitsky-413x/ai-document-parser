#!/usr/bin/env node
import { copyFileSync, mkdirSync, existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const fontsSrc = path.join(root, 'fonts');
const fontsDest = path.join(root, 'dist', 'fonts');

if (existsSync(fontsSrc)) {
  mkdirSync(fontsDest, { recursive: true });
  for (const name of ['DejaVuSans.ttf', 'LICENSE-DejaVu.txt']) {
    const from = path.join(fontsSrc, name);
    if (existsSync(from)) {
      copyFileSync(from, path.join(fontsDest, name));
    }
  }
  // also keep at package root fonts/ (already there for npm files)
  console.log('Copied fonts to dist/fonts');
}
