/** Shared paragraph-style marker for blockquote round-trips (DOCX / ODT / RTF). */

export const QUOTE_STYLE_ID = 'ADPQuote';

/** Word built-in quote styles we accept on parse (in addition to ADPQuote). */
const WORD_QUOTE_STYLES = /^(ADPQuote|Quote|IntenseQuote)$/i;

export function isQuoteStyleId(styleId: string | undefined | null): boolean {
  if (!styleId) return false;
  return WORD_QUOTE_STYLES.test(styleId.trim());
}

/** Merge consecutive `{ type: 'blockquote' }` blocks into one. */
export function coalesceAdjacentBlockquotes<
  T extends { type: string; blocks?: unknown[] },
>(blocks: T[]): T[] {
  const out: T[] = [];
  for (const block of blocks) {
    const prev = out[out.length - 1];
    if (
      block.type === 'blockquote' &&
      prev?.type === 'blockquote' &&
      Array.isArray(prev.blocks) &&
      Array.isArray(block.blocks)
    ) {
      prev.blocks = [...prev.blocks, ...block.blocks];
      continue;
    }
    out.push(block);
  }
  return out;
}
