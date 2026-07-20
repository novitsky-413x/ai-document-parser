/** Escape text for HTML element content / attributes. */
export function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Strip a leading UTF-8 BOM. */
export function stripBom(s: string): string {
  return s.startsWith('\uFEFF') ? s.slice(1) : s;
}

/** Escape a single CSV cell (RFC-style quotes when needed). */
export function csvEscapeCell(value: string | number | boolean | null | undefined): string {
  if (value === null || value === undefined) return '';
  const s = String(value);
  return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}
