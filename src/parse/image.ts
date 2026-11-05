import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import type { ContentBlock, DocumentWarning } from '../types.js';
import { emitMarkdownFromBlocks, flattenBlocksToParagraphs } from '../convert/blocks.js';
import { ocrImageFile } from './pdf-ocr.js';
import { linesToBlocks, wordsToLines } from './pdf-structure.js';

export interface ImageExtractResult {
  text: string;
  format: 'image';
  blocks: ContentBlock[];
  paragraphs: string[];
  warnings: DocumentWarning[];
}

/** OCR a raster image with the same bundled Tesseract runtime as scanned PDFs. */
export async function extractImageText(input: string | Buffer): Promise<ImageExtractResult> {
  let imagePath = typeof input === 'string' ? input : '';
  let tmpDir: string | undefined;
  try {
    if (!imagePath) {
      tmpDir = await mkdtemp(path.join(os.tmpdir(), 'adp-img-ocr-'));
      imagePath = path.join(tmpDir, 'input.bin');
      await writeFile(imagePath, input);
    }
    const ocr = await ocrImageFile(imagePath);
    const warnings: DocumentWarning[] = [];
    for (const warning of ocr.warnings) warnings.push(warning);
    const blocked = warnings.some(
      (warning) =>
        warning.code === 'IMAGE_OCR_UNAVAILABLE' || warning.code === 'IMAGE_OCR_FAILED'
    );
    if (ocr.words.length) {
      warnings.push({
        code: 'IMAGE_OCR_USED',
        message: 'Текст с изображения распознан Tesseract (rus+eng).',
      });
    } else if (!blocked) {
      warnings.push({
        code: 'IMAGE_OCR_EMPTY',
        message: 'Tesseract не нашёл текста на изображении.',
      });
    }
    const structured = linesToBlocks(wordsToLines(ocr.words));
    for (const warning of structured.warnings) warnings.push(warning);
    const text = emitMarkdownFromBlocks(structured.blocks).trim();
    return {
      text,
      format: 'image',
      blocks: structured.blocks,
      paragraphs: flattenBlocksToParagraphs(structured.blocks),
      warnings,
    };
  } finally {
    if (tmpDir) await rm(tmpDir, { recursive: true, force: true });
  }
}
