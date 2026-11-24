/**
 * Minimal CSS support for HTML → flow writes (DOCX/ODT layout).
 *
 * Deliberately small: inline `style`, `<style>` rules with tag / .class / #id
 * compounds joined by descendant or child combinators, `@page { size; margin }`.
 * Anything else is reported back so writers can emit HTML_STYLE_IGNORED.
 */
import type { TextAlign } from '../types.js';

export type ElementInfo = { tag: string; id?: string; classes: string[] };

type Compound = { tag?: string; id?: string; classes: string[] };
type Selector = { parts: Compound[]; combinators: Array<' ' | '>'>; specificity: number };

export type CssRule = { selector: Selector; order: number; decls: Record<string, string> };

export type CssSheet = {
  rules: CssRule[];
  page: Record<string, string>;
  /** Selectors / at-rules we could not apply (reported once). */
  ignored: string[];
};

/** Inherited (text) style. Sizes in points. */
export type TextStyle = {
  sizePt: number;
  font?: string;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  letterSpacingPt?: number;
  align?: TextAlign;
  lineHeight?: number;
  textIndentPt?: number;
};

/** Non-inherited box style of a block element. */
export type BoxStyle = {
  marginTopPt?: number;
  marginBottomPt?: number;
  marginLeftPt?: number;
  marginRightPt?: number;
  pageBreakBefore?: boolean;
  pageBreakAfter?: boolean;
};

export const ROOT_FONT_PT = 12;

export function emptySheet(): CssSheet {
  return { rules: [], page: {}, ignored: [] };
}

export function parseDeclarations(styleText: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const part of styleText.replace(/\/\*[\s\S]*?\*\//g, '').split(';')) {
    const at = part.indexOf(':');
    if (at < 0) continue;
    const prop = part.slice(0, at).trim().toLowerCase();
    const value = part
      .slice(at + 1)
      .replace(/!important/gi, '')
      .trim();
    if (prop && value) out[prop] = value;
  }
  return out;
}

function parseCompound(text: string): Compound | null {
  const m = /^(\*|[a-zA-Z][\w-]*)?((?:[.#][\w-]+)*)$/.exec(text);
  if (!m || (!m[1] && !m[2])) return null;
  const c: Compound = { classes: [] };
  if (m[1] && m[1] !== '*') c.tag = m[1].toLowerCase();
  const re = /([.#])([\w-]+)/g;
  let p: RegExpExecArray | null;
  while ((p = re.exec(m[2])) !== null) {
    if (p[1] === '#') c.id = p[2];
    else c.classes.push(p[2]);
  }
  return c;
}

function parseSelector(text: string): Selector | null {
  const tokens = text
    .trim()
    .replace(/\s*>\s*/g, ' > ')
    .split(/\s+/)
    .filter(Boolean);
  const parts: Compound[] = [];
  const combinators: Array<' ' | '>'> = [];
  let pending: ' ' | '>' = ' ';
  for (const t of tokens) {
    if (t === '>') {
      pending = '>';
      continue;
    }
    const c = parseCompound(t);
    if (!c) return null;
    if (parts.length) combinators.push(pending);
    parts.push(c);
    pending = ' ';
  }
  if (!parts.length) return null;
  let specificity = 0;
  for (const c of parts) {
    if (c.id) specificity += 100;
    specificity += c.classes.length * 10;
    if (c.tag) specificity += 1;
  }
  return { parts, combinators, specificity };
}

/** Parse `<style>` text into simple rules + `@page` declarations. */
export function parseCss(cssText: string, sheet: CssSheet = emptySheet()): CssSheet {
  let css = cssText.replace(/\/\*[\s\S]*?\*\//g, '').replace(/<!--|-->/g, '');
  // Drop nested at-rule blocks (@media, @font-face, @supports …) but keep @page.
  for (;;) {
    const m = /@(?!page\b)([\w-]+)[^{;]*\{/i.exec(css);
    if (!m) break;
    let depth = 0;
    let end = m.index;
    for (let i = m.index; i < css.length; i++) {
      if (css[i] === '{') depth++;
      else if (css[i] === '}') {
        depth--;
        if (depth === 0) {
          end = i + 1;
          break;
        }
      }
    }
    if (end <= m.index) end = css.length;
    noteIgnored(sheet, `@${m[1].toLowerCase()}`);
    css = css.slice(0, m.index) + css.slice(end);
  }
  css = css.replace(/@[\w-]+[^{;]*;/g, (at) => {
    noteIgnored(sheet, at.trim().split(/\s+/)[0]);
    return '';
  });
  const ruleRe = /([^{}]+)\{([^{}]*)\}/g;
  let m: RegExpExecArray | null;
  while ((m = ruleRe.exec(css)) !== null) {
    const selText = m[1].trim();
    const decls = parseDeclarations(m[2]);
    if (/^@page\b/i.test(selText)) {
      Object.assign(sheet.page, decls);
      continue;
    }
    for (const one of selText.split(',')) {
      const sel = parseSelector(one);
      if (!sel) {
        noteIgnored(sheet, one.trim());
        continue;
      }
      sheet.rules.push({ selector: sel, order: sheet.rules.length, decls });
    }
  }
  return sheet;
}

function noteIgnored(sheet: CssSheet, what: string): void {
  if (what && !sheet.ignored.includes(what)) sheet.ignored.push(what);
}

function compoundMatches(c: Compound, el: ElementInfo): boolean {
  if (c.tag && c.tag !== el.tag) return false;
  if (c.id && c.id !== el.id) return false;
  return c.classes.every((cls) => el.classes.includes(cls));
}

/** `ancestors` is outermost-first and excludes `el`. */
function selectorMatches(sel: Selector, el: ElementInfo, ancestors: ElementInfo[]): boolean {
  const last = sel.parts.length - 1;
  if (!compoundMatches(sel.parts[last], el)) return false;
  let ai = ancestors.length - 1;
  for (let pi = last - 1; pi >= 0; pi--) {
    const comb = sel.combinators[pi];
    if (comb === '>') {
      if (ai < 0 || !compoundMatches(sel.parts[pi], ancestors[ai])) return false;
      ai--;
    } else {
      while (ai >= 0 && !compoundMatches(sel.parts[pi], ancestors[ai])) ai--;
      if (ai < 0) return false;
      ai--;
    }
  }
  return true;
}

/** Cascaded declarations for an element: sheet rules by specificity/order, then inline style. */
export function cascade(
  sheet: CssSheet,
  el: ElementInfo,
  ancestors: ElementInfo[],
  inlineStyle: string | undefined
): Record<string, string> {
  const matched = sheet.rules
    .filter((r) => selectorMatches(r.selector, el, ancestors))
    .sort((a, b) => a.selector.specificity - b.selector.specificity || a.order - b.order);
  const out: Record<string, string> = {};
  for (const r of matched) Object.assign(out, expandShorthands(r.decls));
  if (inlineStyle) Object.assign(out, expandShorthands(parseDeclarations(inlineStyle)));
  return out;
}

function splitBoxValues(v: string): [string, string, string, string] | null {
  const p = v.trim().split(/\s+/);
  if (p.length === 1) return [p[0], p[0], p[0], p[0]];
  if (p.length === 2) return [p[0], p[1], p[0], p[1]];
  if (p.length === 3) return [p[0], p[1], p[2], p[1]];
  if (p.length === 4) return [p[0], p[1], p[2], p[3]];
  return null;
}

export function expandShorthands(decls: Record<string, string>): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(decls)) {
    if (k === 'margin' || k === 'padding') {
      const box = splitBoxValues(v);
      if (box) {
        out[`${k}-top`] = box[0];
        out[`${k}-right`] = box[1];
        out[`${k}-bottom`] = box[2];
        out[`${k}-left`] = box[3];
        continue;
      }
    }
    if (k === 'font') {
      // font: [style] [weight] size[/line-height] family
      const m = /^(.*?)([\d.]+(?:pt|px|em|rem|%|mm|cm|in)|xx-small|x-small|small|medium|large|x-large|xx-large)(?:\s*\/\s*([^\s]+))?\s+(.+)$/i.exec(
        v.trim()
      );
      if (m) {
        for (const w of m[1].trim().split(/\s+/).filter(Boolean)) {
          if (/^(italic|oblique)$/i.test(w)) out['font-style'] = w;
          else if (/^(bold|bolder|lighter|\d{3})$/i.test(w)) out['font-weight'] = w;
        }
        out['font-size'] = m[2];
        if (m[3]) out['line-height'] = m[3];
        out['font-family'] = m[4];
        continue;
      }
    }
    out[k] = v;
  }
  return out;
}

const SIZE_KEYWORDS: Record<string, number> = {
  'xx-small': 7,
  'x-small': 7.5,
  small: 10,
  medium: 12,
  large: 13.5,
  'x-large': 18,
  'xx-large': 24,
};

/** CSS length → points. `basePt` resolves em / %; `rootPt` resolves rem. */
export function lengthToPt(value: string, basePt: number, rootPt = ROOT_FONT_PT): number | undefined {
  const v = value.trim().toLowerCase();
  if (v === '0') return 0;
  const m = /^(-?[\d.]+)(pt|px|mm|cm|in|pc|em|rem|%|q)?$/.exec(v);
  if (!m) return undefined;
  const n = Number(m[1]);
  if (!Number.isFinite(n)) return undefined;
  switch (m[2]) {
    case 'pt':
      return n;
    case 'px':
      return n * 0.75;
    case 'mm':
      return (n * 72) / 25.4;
    case 'cm':
      return (n * 72) / 2.54;
    case 'q':
      return (n * 72) / 101.6;
    case 'in':
      return n * 72;
    case 'pc':
      return n * 12;
    case 'em':
      return n * basePt;
    case 'rem':
      return n * rootPt;
    case '%':
      return (n * basePt) / 100;
    default:
      return n === 0 ? 0 : undefined;
  }
}

function fontSizeToPt(value: string, parentPt: number): number | undefined {
  const v = value.trim().toLowerCase();
  if (SIZE_KEYWORDS[v] != null) return SIZE_KEYWORDS[v];
  if (v === 'smaller') return parentPt / 1.2;
  if (v === 'larger') return parentPt * 1.2;
  return lengthToPt(v, parentPt);
}

const GENERIC_FONTS: Record<string, string> = {
  serif: 'Times New Roman',
  'sans-serif': 'Arial',
  monospace: 'Courier New',
  'system-ui': 'Arial',
  cursive: 'Comic Sans MS',
};

export function firstFontFamily(value: string): string | undefined {
  const first = value.split(',')[0]?.trim().replace(/^["']|["']$/g, '').trim();
  if (!first) return undefined;
  return GENERIC_FONTS[first.toLowerCase()] ?? first;
}

const NAMED_COLORS: Record<string, string> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  gray: '808080',
  grey: '808080',
  darkgray: 'A9A9A9',
  darkgrey: 'A9A9A9',
  silver: 'C0C0C0',
  navy: '000080',
  maroon: '800000',
  purple: '800080',
  teal: '008080',
  olive: '808000',
  orange: 'FFA500',
  yellow: 'FFFF00',
};

export function parseColor(value: string): string | undefined {
  const v = value.trim().toLowerCase();
  if (NAMED_COLORS[v]) return NAMED_COLORS[v];
  let m = /^#([0-9a-f]{3})$/.exec(v);
  if (m) return m[1].split('').map((c) => c + c).join('').toUpperCase();
  m = /^#([0-9a-f]{6})$/.exec(v);
  if (m) return m[1].toUpperCase();
  m = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)/.exec(v);
  if (m) {
    return [m[1], m[2], m[3]]
      .map((n) => Math.min(255, Number(n)).toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase();
  }
  return undefined;
}

function parseAlign(value: string): TextAlign | undefined {
  const v = value.trim().toLowerCase();
  if (v === 'left' || v === 'start') return 'left';
  if (v === 'right' || v === 'end') return 'right';
  if (v === 'center' || v === '-webkit-center') return 'center';
  if (v === 'justify') return 'justify';
  return undefined;
}

function isPageBreak(v: string): boolean {
  return /^(always|page|left|right|recto|verso)$/i.test(v.trim());
}

/** Properties that are accepted but have no flow-document meaning (never warned). */
const SILENT_PROPS = new Set([
  'display',
  'box-sizing',
  'font-variant-ligatures',
  '-webkit-print-color-adjust',
  'print-color-adjust',
  'page-break-inside',
  'break-inside',
  'orphans',
  'widows',
]);

export type ComputedStyle = { text: TextStyle; box: BoxStyle };

/**
 * Apply cascaded declarations on top of the parent's inherited text style.
 * Unsupported properties are pushed to `ignored` (property names).
 */
export function computeStyle(
  decls: Record<string, string>,
  parent: TextStyle,
  ignored: Set<string>
): ComputedStyle {
  const text: TextStyle = { ...parent };
  const box: BoxStyle = {};
  // font-size first: em-based values below resolve against it.
  if (decls['font-size'] != null) {
    const pt = fontSizeToPt(decls['font-size'], parent.sizePt);
    if (pt != null && pt > 0) text.sizePt = pt;
    else ignored.add('font-size');
  }
  const em = text.sizePt;
  for (const [prop, value] of Object.entries(decls)) {
    const v = value.trim().toLowerCase();
    switch (prop) {
      case 'font-size':
        break;
      case 'font-family': {
        const f = firstFontFamily(value);
        if (f) text.font = f;
        break;
      }
      case 'font-weight':
        if (/^(bold|bolder)$/.test(v) || (/^\d+$/.test(v) && Number(v) >= 600)) text.bold = true;
        else if (/^(normal|lighter)$/.test(v) || (/^\d+$/.test(v) && Number(v) < 600)) text.bold = false;
        else ignored.add(prop);
        break;
      case 'font-style':
        if (/^(italic|oblique)/.test(v)) text.italic = true;
        else if (v === 'normal') text.italic = false;
        else ignored.add(prop);
        break;
      case 'text-decoration':
      case 'text-decoration-line':
        if (v === 'none') {
          text.underline = false;
          text.strike = false;
        } else {
          if (v.includes('underline')) text.underline = true;
          if (v.includes('line-through')) text.strike = true;
          if (!/underline|line-through/.test(v)) ignored.add(prop);
        }
        break;
      case 'color': {
        const c = parseColor(value);
        if (c) text.color = c;
        else ignored.add(prop);
        break;
      }
      case 'letter-spacing': {
        if (v === 'normal') {
          text.letterSpacingPt = undefined;
          break;
        }
        const pt = lengthToPt(v, em);
        if (pt != null) text.letterSpacingPt = pt;
        else ignored.add(prop);
        break;
      }
      case 'text-align': {
        const a = parseAlign(v);
        if (a) text.align = a;
        else ignored.add(prop);
        break;
      }
      case 'line-height': {
        if (v === 'normal') {
          text.lineHeight = undefined;
          break;
        }
        if (/^[\d.]+$/.test(v)) {
          text.lineHeight = Number(v);
          break;
        }
        if (/%$/.test(v)) {
          text.lineHeight = Number(v.slice(0, -1)) / 100;
          break;
        }
        const pt = lengthToPt(v, em);
        if (pt != null && em > 0) text.lineHeight = pt / em;
        else ignored.add(prop);
        break;
      }
      case 'text-indent': {
        const pt = lengthToPt(v, em);
        if (pt != null) text.textIndentPt = pt;
        else ignored.add(prop);
        break;
      }
      case 'margin-top':
      case 'margin-bottom':
      case 'margin-left':
      case 'margin-right':
      case 'padding-left':
      case 'padding-right':
      case 'padding-top':
      case 'padding-bottom': {
        if (v === 'auto') break;
        const pt = lengthToPt(v, em);
        if (pt == null) {
          ignored.add(prop);
          break;
        }
        const side = prop.split('-')[1];
        const key = (`margin${side[0].toUpperCase()}${side.slice(1)}Pt`) as keyof BoxStyle;
        (box as Record<string, number>)[key] = ((box[key] as number | undefined) ?? 0) + pt;
        break;
      }
      case 'page-break-before':
      case 'break-before':
        if (isPageBreak(v)) box.pageBreakBefore = true;
        break;
      case 'page-break-after':
      case 'break-after':
        if (isPageBreak(v)) box.pageBreakAfter = true;
        break;
      default:
        if (!SILENT_PROPS.has(prop)) ignored.add(prop);
    }
  }
  return { text, box };
}

/** Named page sizes for `@page { size }` (points). */
const PAGE_SIZES: Record<string, [number, number]> = {
  a3: [841.89, 1190.55],
  a4: [595.28, 841.89],
  a5: [419.53, 595.28],
  b5: [498.9, 708.66],
  letter: [612, 792],
  legal: [612, 1008],
};

export function parsePageSize(
  value: string
): { widthPt: number; heightPt: number } | undefined {
  const parts = value.trim().toLowerCase().split(/\s+/);
  let landscape = false;
  let dims: [number, number] | undefined;
  const lengths: number[] = [];
  for (const p of parts) {
    if (p === 'landscape') landscape = true;
    else if (p === 'portrait') landscape = false;
    else if (PAGE_SIZES[p]) dims = PAGE_SIZES[p];
    else {
      const pt = lengthToPt(p, ROOT_FONT_PT);
      if (pt != null && pt > 0) lengths.push(pt);
    }
  }
  if (!dims && lengths.length) dims = [lengths[0], lengths[1] ?? lengths[0]];
  if (!dims) return undefined;
  const [w, h] = dims;
  return landscape ? { widthPt: Math.max(w, h), heightPt: Math.min(w, h) } : { widthPt: w, heightPt: h };
}
