import type { OcrOptions, OcrProvider } from '../types.js';

export const DEFAULT_OCR_LANGUAGES: readonly string[] = ['rus', 'eng'];
export const DEFAULT_OCR_PSM = 6;
export const DEFAULT_OCR_DPI = 300;
export const PDF_OCR_PAGE_CAP = 80;
export const DEFAULT_OCR_TIMEOUT_MS = 120_000;
export const DEFAULT_PDFTOPPM_TIMEOUT_MS = 60_000;

export type OcrEnabledResolved = 'auto' | true | false;

export interface ResolvedOcrOptions {
  enabled: OcrEnabledResolved;
  languages: string[];
  bin?: string;
  tessdataDir?: string;
  popplerBin?: string;
  psm: number;
  oem?: number;
  dpi: number;
  pageCap: number;
  timeoutMs: number;
  provider?: OcrProvider;
}

export function languageFlag(languages: readonly string[]): string {
  return languages.join('+');
}

export function normalizeOcrLanguages(languages?: readonly string[]): string[] {
  const out: string[] = [];
  const seen = new Set<string>();
  const source = languages ?? DEFAULT_OCR_LANGUAGES;
  for (const raw of source) {
    const lang = String(raw).trim();
    if (!lang) continue;
    const key = lang.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(lang);
  }
  return out.length ? out : [...DEFAULT_OCR_LANGUAGES];
}

function finiteInt(value: unknown, fallback: number, min: number, max?: number): number {
  const n = typeof value === 'number' ? value : Number(value);
  if (!Number.isFinite(n)) return fallback;
  const i = Math.floor(n);
  if (i < min) return fallback;
  if (max != null && i > max) return fallback;
  return i;
}

function emptyToUndef(value?: string): string | undefined {
  if (value == null) return undefined;
  const trimmed = value.trim();
  return trimmed ? trimmed : undefined;
}

function baseResolved(
  partial: Partial<ResolvedOcrOptions> & { enabled: OcrEnabledResolved }
): ResolvedOcrOptions {
  return {
    enabled: partial.enabled,
    languages: partial.languages ?? [...DEFAULT_OCR_LANGUAGES],
    bin: partial.bin,
    tessdataDir: partial.tessdataDir,
    popplerBin: partial.popplerBin,
    psm: partial.psm ?? DEFAULT_OCR_PSM,
    oem: partial.oem,
    dpi: partial.dpi ?? DEFAULT_OCR_DPI,
    pageCap: partial.pageCap ?? PDF_OCR_PAGE_CAP,
    timeoutMs: partial.timeoutMs ?? DEFAULT_OCR_TIMEOUT_MS,
    provider: partial.provider,
  };
}

/**
 * Normalize `parseDocument` / `convertDocument` `ocr` into a resolved options object.
 * Defaults match historical behavior: auto OCR, rus+eng, PSM 6, 300 dpi, 80-page cap.
 */
export function normalizeOcrOptions(ocr?: boolean | OcrOptions): ResolvedOcrOptions {
  if (ocr === false) return baseResolved({ enabled: false });
  if (ocr === true || ocr == null) return baseResolved({ enabled: 'auto' });

  let enabled: OcrEnabledResolved = 'auto';
  if (ocr.enabled === false) enabled = false;
  else if (ocr.enabled === true) enabled = true;

  return baseResolved({
    enabled,
    languages: normalizeOcrLanguages(ocr.languages),
    bin: emptyToUndef(ocr.bin),
    tessdataDir: emptyToUndef(ocr.tessdataDir),
    popplerBin: emptyToUndef(ocr.popplerBin),
    psm: ocr.psm == null ? DEFAULT_OCR_PSM : finiteInt(ocr.psm, DEFAULT_OCR_PSM, 0, 13),
    oem: ocr.oem == null ? undefined : finiteInt(ocr.oem, 3, 0, 3),
    dpi: ocr.dpi == null ? DEFAULT_OCR_DPI : finiteInt(ocr.dpi, DEFAULT_OCR_DPI, 1, 2400),
    pageCap: ocr.pageCap == null ? PDF_OCR_PAGE_CAP : finiteInt(ocr.pageCap, PDF_OCR_PAGE_CAP, 1),
    timeoutMs:
      ocr.timeoutMs == null
        ? DEFAULT_OCR_TIMEOUT_MS
        : finiteInt(ocr.timeoutMs, DEFAULT_OCR_TIMEOUT_MS, 1),
    provider: ocr.provider,
  });
}

/** Map CLI flags to the `ocr` option. `undefined` means package defaults. */
export function ocrOptionsFromCliFlags(flags: {
  ocr?: boolean;
  noOcr?: boolean;
  ocrLang?: string;
}): boolean | OcrOptions | undefined {
  if (flags.noOcr && flags.ocr) {
    throw new Error('Use either --ocr or --no-ocr, not both');
  }
  if (flags.noOcr) return false;
  const languages = flags.ocrLang
    ? flags.ocrLang.split(',').map((s) => s.trim()).filter(Boolean)
    : undefined;
  if (!flags.ocr && !languages) return undefined;
  const out: OcrOptions = {};
  if (flags.ocr) out.enabled = 'auto';
  if (languages) out.languages = languages;
  return out;
}
