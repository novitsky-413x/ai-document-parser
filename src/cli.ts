#!/usr/bin/env node
/**
 * CLI for ai-document-parser. Run without arguments for usage.
 */
import { readFile, writeFile } from 'node:fs/promises';
import {
  parseDocument,
  convertDocument,
  editDocument,
  writeDocument,
  DocumentError,
} from './index.js';
import { getCapabilities } from './adapters/registry.js';
import { parseCsv } from './parse/text-formats.js';
import { ocrOptionsFromCliFlags } from './parse/ocr-options.js';
import type { DocumentWarning, OcrOptions, ParseOutputMode } from './types.js';

/** How many values each flag consumes (Infinity: rest of argv). */
const FLAG_ARITY: Record<string, number> = {
  '--output': 1,
  '--to': 1,
  '-o': 1,
  '--format': 1,
  '--content': 1,
  '--title': 1,
  '--sheet': 1,
  '--csv': 1,
  '--replace': 2,
  '--patch': 3,
  '--all': 0,
  '--append': Infinity,
  '--ocr': 0,
  '--no-ocr': 0,
  '--ocr-lang': 1,
};

/** First token after the command that is not a flag or a flag value. */
function findPositional(args: string[]): string | undefined {
  for (let i = 1; i < args.length; i++) {
    const a = args[i];
    if (a in FLAG_ARITY) {
      const n = FLAG_ARITY[a];
      if (!Number.isFinite(n)) return undefined;
      i += n;
      continue;
    }
    if (a.startsWith('--')) continue;
    return a;
  }
  return undefined;
}

function flagValue(args: string[], flag: string): string | undefined {
  const idx = args.indexOf(flag);
  if (idx < 0) return undefined;
  const value = args[idx + 1];
  if (value === undefined || (value.startsWith('--') && value in FLAG_ARITY)) {
    usage(`Missing value for ${flag}`);
  }
  return value;
}

async function resolveContent(raw: string): Promise<string> {
  if (raw === '-' || raw === '@-') {
    const chunks: Buffer[] = [];
    for await (const chunk of process.stdin) {
      chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    }
    return Buffer.concat(chunks).toString('utf8');
  }
  if (raw.startsWith('@')) return readFile(raw.slice(1), 'utf8');
  return raw;
}

function usage(reason?: string): never {
  if (reason) console.error(`Error: ${reason}\n`);
  console.error(`ai-document-parser <command> ...

Commands:
  read <file> [--output text|structured|ast|chunks] [--ocr|--no-ocr] [--ocr-lang rus,eng]
  convert <file> --to <format> [-o <outfile>] [--ocr|--no-ocr] [--ocr-lang rus,eng]
  edit <docx> --replace <old> <new> [--all]
  edit <docx> --append <paragraph> [<paragraph>...]
  edit <xlsx> --patch <sheet> <cell> <value>
  write-docx <file> --content <text|@file|->
  write-xlsx <file> --sheet <name> --csv <text|@file|->
  write-pdf <file> --content <text|@file|-> [--title <title>]
  capabilities [--format <id>]
`);
  process.exit(1);
}


function cliOcrOption(args: string[]): boolean | OcrOptions | undefined {
  return ocrOptionsFromCliFlags({
    ocr: args.includes('--ocr'),
    noOcr: args.includes('--no-ocr'),
    ocrLang: flagValue(args, '--ocr-lang'),
  });
}

function printWarnings(warnings?: DocumentWarning[]): void {
  if (!warnings?.length) return;
  for (const warning of warnings) {
    console.error(`${warning.code}: ${warning.message}`);
  }
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const cmd = args[0];
  if (!cmd) usage();

  if (cmd === 'read') {
    const file = findPositional(args);
    if (!file) usage('read requires a file');
    const output = (flagValue(args, '--output') ?? 'text') as ParseOutputMode;
    const result = await parseDocument(file, { output, ocr: cliOcrOption(args) });
    if (output === 'ast') {
      console.log(JSON.stringify(result.ast, null, 2));
    } else {
      console.log(result.text);
    }
    printWarnings(result.warnings);
    return;
  }

  if (cmd === 'convert') {
    const file = findPositional(args);
    const to = flagValue(args, '--to');
    if (!file || !to) usage('convert requires a file and --to <format>');
    const result = await convertDocument(file, to, { ocr: cliOcrOption(args) });
    const outPath = flagValue(args, '-o');
    const serialized = serializeConvertValue(result.value);
    if (outPath) {
      await writeFile(outPath, serialized);
      console.log(`Wrote ${outPath}`);
    } else {
      if (typeof serialized === 'string') console.log(serialized);
      else process.stdout.write(serialized);
    }
    printWarnings(result.messages as DocumentWarning[] | undefined);
    return;
  }

  if (cmd === 'capabilities') {
    const filter = flagValue(args, '--format');
    console.log(JSON.stringify(getCapabilities(filter), null, 2));
    return;
  }

  if (cmd === 'edit') {
    const file = findPositional(args);
    if (!file) usage('edit requires a file');
    const rIdx = args.indexOf('--replace');
    const aIdx = args.indexOf('--append');
    const pIdx = args.indexOf('--patch');

    if (rIdx >= 0) {
      const oldString = args[rIdx + 1];
      const newString = args[rIdx + 2];
      if (oldString === undefined || newString === undefined) {
        usage('--replace requires <old> and <new>');
      }
      const replaceAll = args.includes('--all');
      const result = await editDocument(file, {
        type: 'docx',
        op: 'replace',
        oldString,
        newString,
        replaceAll,
      });
      console.log(result.message);
      return;
    }

    if (aIdx >= 0) {
      const paragraphs = args.slice(aIdx + 1);
      if (!paragraphs.length) usage('--append requires at least one paragraph');
      const result = await editDocument(file, {
        type: 'docx',
        op: 'append',
        paragraphs,
      });
      console.log(result.message);
      return;
    }

    if (pIdx >= 0) {
      const sheet = args[pIdx + 1];
      const cell = args[pIdx + 2];
      const value = args[pIdx + 3];
      if (sheet === undefined || cell === undefined || value === undefined) {
        usage('--patch requires <sheet> <cell> <value>');
      }
      const result = await editDocument(file, {
        type: 'spreadsheet',
        op: 'patchCells',
        updates: [{ sheet, cell, value }],
      });
      console.log(result.message);
      return;
    }

    usage('edit requires --replace, --append or --patch');
  }

  if (cmd === 'write-docx') {
    const file = findPositional(args);
    const rawContent = flagValue(args, '--content');
    if (!file || rawContent === undefined) usage('write-docx requires a file and --content');
    const content = await resolveContent(rawContent);
    const result = await writeDocument(file, { type: 'docx', content });
    console.log(result.message);
    return;
  }

  if (cmd === 'write-xlsx') {
    const file = findPositional(args);
    const sheet = flagValue(args, '--sheet') ?? 'Sheet1';
    const rawCsv = flagValue(args, '--csv');
    if (!file || rawCsv === undefined) usage('write-xlsx requires a file and --csv');
    const csvText = await resolveContent(rawCsv);
    const { rows } = parseCsv(csvText);
    const result = await writeDocument(file, {
      type: 'xlsx',
      sheets: { [sheet]: rows },
    });
    console.log(result.message);
    return;
  }

  if (cmd === 'write-pdf') {
    const file = findPositional(args);
    const rawContent = flagValue(args, '--content');
    if (!file || rawContent === undefined) usage('write-pdf requires a file and --content');
    const content = await resolveContent(rawContent);
    const title = flagValue(args, '--title');
    const result = await writeDocument(file, { type: 'pdf', content, title });
    console.log(result.message);
    return;
  }

  usage(`Unknown command: ${cmd}`);
}

function serializeConvertValue(value: unknown): string | Buffer {
  if (typeof value === 'string') return value;
  if (Array.isArray(value)) return JSON.stringify(value, null, 2);
  if (value instanceof Uint8Array) return Buffer.from(value);
  return JSON.stringify(value, null, 2);
}

main().catch((err) => {
  if (err instanceof DocumentError) {
    console.error(`${err.code}: ${err.message}`);
  } else {
    console.error(err instanceof Error ? err.message : err);
  }
  process.exit(1);
});
