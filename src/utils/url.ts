/** Allow only safe URL schemes in HTML/MD emit; relative paths are kept. */
export function sanitizeHref(href: string): string | undefined {
  const trimmed = href.trim();
  if (!trimmed) return undefined;
  const lower = trimmed.toLowerCase();
  if (/^(javascript|data|vbscript):/i.test(lower)) return undefined;
  // Absolute URI with a scheme: only allow known-safe ones.
  if (/^[a-z][a-z0-9+.-]*:/i.test(trimmed)) {
    if (
      lower.startsWith('http://') ||
      lower.startsWith('https://') ||
      lower.startsWith('mailto:')
    ) {
      return trimmed;
    }
    return undefined;
  }
  // Relative paths, fragments, and rooted paths.
  return trimmed;
}
