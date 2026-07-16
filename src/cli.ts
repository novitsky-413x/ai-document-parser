#!/usr/bin/env node
/**
 * Minimal CLI for ai-document-parser
 * Usage:
 *   ai-document-parser read <file> [--output text|structured|ast|chunks]
 *   ai-document-parser convert <file> --to md|html|csv|text|chunks [-o out]
 *   ai-document-parser edit <file> --replace <old> <new> [--all]
 *   ai-document-parser write-docx <file> --content <text>
 */
import { readFile, writeFile } from 'node:fs/promises';
import {
  parseDocument,
  convertDocument,
  editDocument,
  writeDocument,
} from './index.js';

function usage(): never {
  console.error(`ai-document-parser <command> ...

Commands:
  read <file> [--output text|structured|ast|chunks]
  convert <file> --to <format> [-o <outfile>]
  edit <docx> --replace <old> <new> [--all]
  write-docx <file> --content <text>
  write-pdf <file> --content <text> [--title <title>]
`);
  process.exit(1);
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const cmd = args[0];
  if (!cmd) usage();

  if (cmd === 'read') {
    const file = args[1];
    if (!file) usage();
    const outIdx = args.indexOf('--output');
    const output = (
      outIdx >= 0 ? args[outIdx + 1] : 'text'
    ) as 'text' | 'structured' | 'ast' | 'chunks';
    const result = await parseDocument(file, { output });
    if (output === 'ast') {
      console.log(JSON.stringify(result.ast, null, 2));
    } else {
      console.log(result.text);
    }
    return;
  }

  if (cmd === 'convert') {
    const file = args[1];
    const toIdx = args.indexOf('--to');
    const oIdx = args.indexOf('-o');
    if (!file || toIdx < 0) usage();
    const to = args[toIdx + 1];
    const result = await convertDocument(file, to);
    const outPath = oIdx >= 0 ? args[oIdx + 1] : undefined;
    if (outPath) {
      const data =
        typeof result.value === 'string'
          ? result.value
          : Buffer.from(result.value as Uint8Array);
      await writeFile(outPath, data);
      console.log(`Wrote ${outPath}`);
    } else {
      if (typeof result.value === 'string') console.log(result.value);
      else process.stdout.write(Buffer.from(result.value as Uint8Array));
    }
    return;
  }

  if (cmd === 'edit') {
    const file = args[1];
    const rIdx = args.indexOf('--replace');
    if (!file || rIdx < 0) usage();
    const oldString = args[rIdx + 1];
    const newString = args[rIdx + 2];
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

  if (cmd === 'write-docx') {
    const file = args[1];
    const cIdx = args.indexOf('--content');
    if (!file || cIdx < 0) usage();
    let content = args[cIdx + 1];
    if (content === '-' || content === '@-') {
      const chunks: Buffer[] = [];
      for await (const chunk of process.stdin) {
        chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
      }
      content = Buffer.concat(chunks).toString('utf8');
    } else if (content?.startsWith('@')) {
      content = await readFile(content.slice(1), 'utf8');
    }
    const result = await writeDocument(file, { type: 'docx', content });
    console.log(result.message);
    return;
  }

  if (cmd === 'write-pdf') {
    const file = args[1];
    const cIdx = args.indexOf('--content');
    const tIdx = args.indexOf('--title');
    if (!file || cIdx < 0) usage();
    const content = args[cIdx + 1];
    const title = tIdx >= 0 ? args[tIdx + 1] : undefined;
    const result = await writeDocument(file, { type: 'pdf', content, title });
    console.log(result.message);
    return;
  }

  usage();
}

main().catch((err) => {
  console.error(err instanceof Error ? err.message : err);
  process.exit(1);
});
