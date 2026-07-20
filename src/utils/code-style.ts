/** Shared paragraph-style marker for fenced code round-trips (DOCX / ODT / RTF). */

export const CODE_STYLE_PREFIX = 'ADPCode';

export function codeStyleId(language?: string): string {
  const lang = (language ?? '').replace(/[^A-Za-z0-9_+-]/g, '').slice(0, 32);
  return lang ? `${CODE_STYLE_PREFIX}_${lang}` : CODE_STYLE_PREFIX;
}

export function parseCodeStyleId(
  styleId: string | undefined | null
): { isCode: boolean; language?: string } {
  if (!styleId) return { isCode: false };
  const m = new RegExp(`^${CODE_STYLE_PREFIX}(?:_(.+))?$`, 'i').exec(styleId.trim());
  if (!m) return { isCode: false };
  return { isCode: true, ...(m[1] ? { language: m[1] } : {}) };
}

/** Common monospace fonts we emit (and accept as a fallback signal). */
export function isMonospaceFont(name: string): boolean {
  return /^(Consolas|Courier New|Courier|Menlo|Monaco|Lucida Console|Liberation Mono|DejaVu Sans Mono)$/i.test(
    name.trim()
  );
}
