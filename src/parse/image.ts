import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import type { ContentBlock, DocumentWarning, OcrOptions } from '../types.js';
import { emitMarkdownFromBlocks, flattenBlocksToParagraphs } from '../convert/blocks.js';
import { ocrImageFile, ocrUsedMessage } from './pdf-ocr.js';
import { normalizeOcrOptions } from './ocr-options.js';
import { linesToBlocks, wordsToLines } from './pdf-structure.js';

export interface ImageExtractResult {
  text: string;
  format: 'image';
  blocks: ContentBlock[];
  paragraphs: string[];
  warnings: DocumentWarning[];
}

/** OCR a raster image with the same Tesseract runtime as scanned PDFs. */
export async function extractImageText(
  input: string | Buffer,
  ocr?: boolean | OcrOptions
): Promise<ImageExtractResult> {
  const options = normalizeOcrOptions(ocr);
  let imagePath = typeof input === 'string' ? input : '';
  let tmpDir: string | undefined;
  try {
    if (options.enabled === false) {
      return {
        text: '',
        format: 'image',
        blocks: [],
        paragraphs: [],
        warnings: [],
      };
    }
    if (!imagePath) {
      tmpDir = await mkdtemp(path.join(os.tmpdir(), 'adp-img-ocr-'));
      imagePath = path.join(tmpDir, 'input.bin');
      await writeFile(imagePath, input);
    }
    const recognized = await ocrImageFile(imagePath, ocr);
    const warnings: DocumentWarning[] = [];
    for (const warning of recognized.warnings) warnings.push(warning);
    const blocked = warnings.some(
      (warning) =>
        warning.code === 'IMAGE_OCR_UNAVAILABLE' ||
        warning.code === 'IMAGE_OCR_FAILED' ||
        warning.code === 'OCR_LANG_MISSING'
    );
    if (recognized.words.length) {
      warnings.push(ocrUsedMessage('image', options.languages, options.dpi));
    } else if (!blocked) {
      warnings.push({
        code: 'IMAGE_OCR_EMPTY',
        message: 'Tesseract found no text in the image.',
      });
    }
    const structured = linesToBlocks(wordsToLines(recognized.words));
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
