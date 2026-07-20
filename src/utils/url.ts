/** Allow only safe URL schemes in HTML emit. */
export function sanitizeHref(href: string): string | undefined {
  const trimmed = href.trim();
  if (!trimmed) return undefined;
  const lower = trimmed.toLowerCase();
  if (
    lower.startsWith('http://') ||
    lower.startsWith('https://') ||
    lower.startsWith('mailto:') ||
    lower.startsWith('#') ||
    lower.startsWith('/') ||
    lower.startsWith('./') ||
    lower.startsWith('../')
  ) {
    return trimmed;
  }
  return undefined;
}
