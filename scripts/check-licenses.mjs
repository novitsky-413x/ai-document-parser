#!/usr/bin/env node
/**
 * Fail CI if production dependency licenses are outside the allowlist.
 * Uses `npm ls --parseable` paths (no license-checker dependency).
 */
import { execSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { existsSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
const pkg = require('../package.json');

const ALLOW = new Set([
  'MIT',
  'MIT*',
  'MIT/X11',
  'Apache-2.0',
  'Apache 2.0',
  'BSD-2-Clause',
  'BSD-3-Clause',
  'ISC',
  '0BSD',
  'BlueOak-1.0.0',
  'CC0-1.0',
  'Unlicense',
  'MIT OR GPL-3.0-or-later',
  '(MIT OR GPL-3.0-or-later)',
  'MIT OR GPL-3.0',
  '(MIT AND Zlib)',
  'Python-2.0',
  'MPL-2.0',
  'Zlib',
]);

/**
 * npm package names that must never appear in the install tree:
 * - xlsx — abandoned community SheetJS build; we use @e965/xlsx instead
 * - docxtemplater — restrictive / commercial ecosystem
 */
const DENY_PACKAGES = new Set(['xlsx', 'docxtemplater']);

/** Packages that omit license field but are known-permissive. */
const KNOWN_LICENSE = new Map([
  // buffers@0.1.1 — http://github.com/substack/node-bufferjs (public domain / MIT-ish)
  ['buffers', 'MIT'],
]);

function normalize(license) {
  if (!license) return '';
  if (Array.isArray(license)) return license.map(normalize).join(' OR ');
  if (typeof license === 'object' && license.type) return String(license.type).trim();
  return String(license).trim();
}

function isAllowed(license) {
  const lic = normalize(license);
  if (!lic) return false;
  if (ALLOW.has(lic)) return true;
  if (/\bMIT\b/i.test(lic) && /\bOR\b|\bAND\b|\/X11/i.test(lic)) return true;
  if (/\bApache-2\.0\b/i.test(lic) && /\bOR\b|\bAND\b/i.test(lic)) return true;
  if (/\bZlib\b/i.test(lic)) return true;
  if (/^MIT\*?$/i.test(lic) || /^BSD/.test(lic)) return true;
  if (/^Custom:/i.test(lic) && /github\.com\/substack/i.test(lic)) return true;
  return false;
}

function readPkgJson(dir) {
  const p = join(dir, 'package.json');
  if (!existsSync(p)) return null;
  try {
    return JSON.parse(readFileSync(p, 'utf8'));
  } catch {
    return null;
  }
}

function npmParseablePaths() {
  let out = '';
  try {
    out = execSync('npm ls --omit=dev --all --parseable', {
      cwd: root,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
      maxBuffer: 20 * 1024 * 1024,
    });
  } catch (err) {
    // npm ls may exit non-zero on peer warnings; stdout still usable
    out = err && typeof err === 'object' && 'stdout' in err ? String(err.stdout) : '';
    if (!out.trim()) {
      console.error('npm ls failed:', err instanceof Error ? err.message : err);
      process.exit(1);
    }
  }
  return out
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l && l !== root);
}

for (const field of ['dependencies', 'optionalDependencies']) {
  const deps = pkg[field];
  if (!deps || typeof deps !== 'object') continue;
  for (const name of Object.keys(deps)) {
    if (DENY_PACKAGES.has(name)) {
      console.error(`Denied package in package.json ${field}: ${name}`);
      process.exit(1);
    }
  }
}

const paths = npmParseablePaths();
const seen = new Set();
const problems = [];
let scanned = 0;

for (const dir of paths) {
  const meta = readPkgJson(dir);
  if (!meta?.name) continue;
  const version = meta.version || '?';
  const key = `${meta.name}@${version}`;
  if (seen.has(key)) continue;
  seen.add(key);
  scanned++;

  if (meta.name === pkg.name) continue;
  if (DENY_PACKAGES.has(meta.name)) {
    problems.push(`${key}: denied package`);
    continue;
  }

  let license = meta.license ?? meta.licenses;
  if (!normalize(license) && KNOWN_LICENSE.has(meta.name)) {
    license = KNOWN_LICENSE.get(meta.name);
  }

  if (!isAllowed(license)) {
    problems.push(`${key}: license "${normalize(license) || '(missing)'}" not in allowlist`);
  }
}

if (problems.length) {
  console.error('License check failed:');
  for (const p of problems) console.error(' -', p);
  process.exit(1);
}

console.log(`License check OK (${scanned} production packages scanned)`);
