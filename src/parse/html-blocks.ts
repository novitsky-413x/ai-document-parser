import type {
  BlockLayout,
  ContentBlock,
  DocumentLayout,
  DocumentWarning,
  InlineRun,
} from '../types.js';
import { flattenBlocksToParagraphs } from '../convert/blocks.js';
import {
  cascade,
  computeStyle,
  emptySheet,
  lengthToPt,
  parseCss,
  parsePageSize,
  ROOT_FONT_PT,
  type BoxStyle,
  type ComputedStyle,
  type CssSheet,
  type ElementInfo,
  type TextStyle,
} from './html-css.js';

function decodeEntities(s: string): string {
  return s
    .replace(/&nbsp;/gi, ' ')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&apos;/gi, "'")
    .replace(/&laquo;/gi, '«')
    .replace(/&raquo;/gi, '»')
    .replace(/&mdash;/gi, '—')
    .replace(/&ndash;/gi, '–')
    .replace(/&hellip;/gi, '…')
    .replace(/&#x([0-9a-f]+);/gi, (_, h: string) => safeCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d: string) => safeCodePoint(Number(d)))
    .replace(/&amp;/gi, '&');
}

function safeCodePoint(cp: number): string {
  if (!Number.isFinite(cp) || cp <= 0 || cp > 0x10ffff) return '';
  return cp === 0xa0 ? ' ' : String.fromCodePoint(cp);
}

/** Find the end index (exclusive) of a balanced HTML element starting at `start`. */
function findBalancedHtmlEnd(html: string, start: number): number {
  const open = /^<([A-Za-z][\w:-]*)\b[^>]*>/.exec(html.slice(start));
  if (!open) return -1;
  const tag = open[1].toLowerCase();
  if (/\/>$/.test(open[0]) || /^(br|hr|img|meta|link|input)$/i.test(tag)) {
    return start + open[0].length;
  }
  const openRe = new RegExp(`<${tag}\\b[^>]*>`, 'gi');
  const closeRe = new RegExp(`</${tag}\\s*>`, 'gi');
  let depth = 0;
  let i = start;
  while (i < html.length) {
    openRe.lastIndex = i;
    closeRe.lastIndex = i;
    const om = openRe.exec(html);
    const cm = closeRe.exec(html);
    const oAt = om ? om.index : -1;
    const cAt = cm ? cm.index : -1;
    if (oAt < 0 && cAt < 0) return -1;
    if (oAt >= 0 && (cAt < 0 || oAt < cAt)) {
      const selfClosing = /\/>$/.test(om![0]);
      if (!selfClosing) depth++;
      i = oAt + om![0].length;
    } else {
      depth--;
      i = cAt + cm![0].length;
      if (depth === 0) return i;
    }
  }
  return -1;
}

/** Block containers: their children are parsed as blocks with inherited style. */
const CONTAINER_TAGS = new Set([
  'div',
  'section',
  'article',
  'main',
  'header',
  'footer',
  'aside',
  'nav',
  'center',
  'figure',
  'figcaption',
  'address',
  'form',
  'fieldset',
  'dl',
  'dt',
  'dd',
]);

function indexOfNextBlockTag(html: string, from: number): number {
  const re =
    /<(h[1-6]|p|ul|ol|table|pre|hr|blockquote|div|section|article|main|header|footer|aside|nav|center|figure|figcaption|address|form|fieldset|dl|dt|dd)(?=[\s/>])/gi;
  re.lastIndex = from;
  const m = re.exec(html);
  return m ? m.index : -1;
}

type OpenTag = ElementInfo & { style?: string; attrs: Record<string, string> };

function parseOpenTag(tagText: string): OpenTag {
  const m = /^<\s*([A-Za-z][\w:-]*)([\s\S]*?)\/?\s*>$/.exec(tagText);
  const tag = (m?.[1] ?? '').toLowerCase();
  const attrs: Record<string, string> = {};
  const attrRe = /([^\s=/>]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+)))?/g;
  let a: RegExpExecArray | null;
  const rest = m?.[2] ?? '';
  while ((a = attrRe.exec(rest)) !== null) {
    attrs[a[1].toLowerCase()] = decodeEntities(a[2] ?? a[3] ?? a[4] ?? '');
  }
  return {
    tag,
    id: attrs.id || undefined,
    classes: (attrs.class ?? '').split(/\s+/).filter(Boolean),
    style: attrs.style,
    attrs,
  };
}

type HtmlCtx = {
  sheet: CssSheet;
  /** Unsupported CSS property names seen while computing styles. */
  ignored: Set<string>;
  /** Carry layout (sizes, fonts, spacing, UA defaults) into blocks. */
  layout: boolean;
  /** Document base text style: run props equal to it are not repeated. */
  base: TextStyle;
};

type Scope = { style: TextStyle; ancestors: ElementInfo[] };

function defaultCtx(): HtmlCtx {
  return {
    sheet: emptySheet(),
    ignored: new Set(),
    layout: false,
    base: { sizePt: ROOT_FONT_PT },
  };
}

function rootScope(): Scope {
  return { style: { sizePt: ROOT_FONT_PT }, ancestors: [] };
}

/** Browser (UA) defaults applied under author CSS in layout mode. */
const UA_DECLS: Record<string, Record<string, string>> = {
  h1: { 'font-size': '2em', 'margin-top': '0.67em', 'margin-bottom': '0.67em' },
  h2: { 'font-size': '1.5em', 'margin-top': '0.83em', 'margin-bottom': '0.83em' },
  h3: { 'font-size': '1.17em', 'margin-top': '1em', 'margin-bottom': '1em' },
  h4: { 'font-size': '1em', 'margin-top': '1.33em', 'margin-bottom': '1.33em' },
  h5: { 'font-size': '0.83em', 'margin-top': '1.67em', 'margin-bottom': '1.67em' },
  h6: { 'font-size': '0.67em', 'margin-top': '2.33em', 'margin-bottom': '2.33em' },
  p: { 'margin-top': '1em', 'margin-bottom': '1em' },
  blockquote: { 'margin-top': '1em', 'margin-bottom': '1em' },
  small: { 'font-size': 'smaller' },
  big: { 'font-size': 'larger' },
};

/** Tag semantics that set inherited text props before author CSS. */
function applyTagSemantics(tag: string, style: TextStyle, layout: boolean): TextStyle {
  const out = { ...style };
  if (tag === 'b' || tag === 'strong' || (layout && /^h[1-6]$/.test(tag))) out.bold = true;
  if (tag === 'i' || tag === 'em' || tag === 'cite' || tag === 'var' || tag === 'dfn')
    out.italic = true;
  if (tag === 's' || tag === 'strike' || tag === 'del') out.strike = true;
  if (tag === 'u' || tag === 'ins') out.underline = true;
  if (tag === 'center') out.align = 'center';
  return out;
}

function computeElement(el: OpenTag, ctx: HtmlCtx, scope: Scope): ComputedStyle {
  const decls: Record<string, string> = {
    ...(ctx.layout ? (UA_DECLS[el.tag] ?? {}) : {}),
  };
  const align = el.attrs.align;
  if (align) decls['text-align'] = align;
  if (el.tag === 'font') {
    if (el.attrs.face) decls['font-family'] = el.attrs.face;
    if (el.attrs.color) decls.color = el.attrs.color;
  }
  Object.assign(decls, cascade(ctx.sheet, el, scope.ancestors, el.style));
  const parent = applyTagSemantics(el.tag, scope.style, ctx.layout);
  return computeStyle(decls, parent, ctx.ignored);
}

type InlineState = { style: TextStyle; href?: string; code?: boolean; tag: string; info?: OpenTag };

const VOID_TAGS = /^(br|hr|img|meta|link|input|wbr|col|area|source)$/;

function runFromState(text: string, st: InlineState, ctx: HtmlCtx): InlineRun {
  const r: InlineRun = { text };
  const s = st.style;
  if (s.bold) r.bold = true;
  if (s.italic) r.italic = true;
  if (s.strike) r.strike = true;
  if (s.underline) r.underline = true;
  if (st.code) r.code = true;
  if (st.href) r.href = st.href;
  if (ctx.layout) {
    if (Math.abs(s.sizePt - ctx.base.sizePt) > 0.01) r.sizePt = round2(s.sizePt);
    if (s.font && s.font !== ctx.base.font) r.font = s.font;
    if (s.color && s.color !== ctx.base.color) r.color = s.color;
    if (s.letterSpacingPt) r.letterSpacingPt = round2(s.letterSpacingPt);
  }
  return r;
}

function round2(n: number): number {
  return Math.round(n * 100) / 100;
}

function sameFormat(a: InlineRun, b: InlineRun): boolean {
  const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
  keys.delete('text');
  for (const k of keys) {
    if ((a as Record<string, unknown>)[k] !== (b as Record<string, unknown>)[k]) return false;
  }
  return true;
}

/** HTML whitespace collapsing + `<br>` (`\n`) handling + merge of equal runs. */
function finalizeRuns(raw: InlineRun[]): InlineRun[] {
  const out: InlineRun[] = [];
  let lineStart = true;
  let endsSpace = false;
  const trimTail = () => {
    while (out.length) {
      const last = out[out.length - 1];
      last.text = last.text.replace(/ +$/, '');
      if (last.text) break;
      out.pop();
    }
  };
  for (const r of raw) {
    let t = r.text.replace(/ +\n/g, '\n').replace(/\n +/g, '\n');
    if (lineStart || endsSpace) t = t.replace(/^ +/, '');
    if (!t) continue;
    if (t.startsWith('\n')) trimTail();
    out.push({ ...r, text: t });
    lineStart = t.endsWith('\n');
    endsSpace = t.endsWith(' ');
  }
  // Trailing spaces / a trailing <br> do not render.
  while (out.length) {
    const last = out[out.length - 1];
    last.text = last.text.replace(/[ \n]+$/, '');
    if (last.text) break;
    out.pop();
  }
  const merged: InlineRun[] = [];
  for (const r of out) {
    const prev = merged[merged.length - 1];
    if (prev && sameFormat(prev, r)) prev.text += r.text;
    else merged.push(r);
  }
  return merged;
}

function parseInline(html: string, ctx: HtmlCtx = defaultCtx(), scope: Scope = rootScope()): InlineRun[] {
  const raw: InlineRun[] = [];
  const stack: InlineState[] = [{ style: scope.style, tag: '#root' }];
  const ancestorsOf = () => [
    ...scope.ancestors,
    ...stack.slice(1).map((s) => s.info!).filter(Boolean),
  ];
  const re = /<!--[\s\S]*?-->|<\/\s*([A-Za-z][\w:-]*)\s*>|<([A-Za-z][\w:-]*)\b[^>]*>|([^<]+)|</g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(html)) !== null) {
    const top = stack[stack.length - 1];
    if (m[3] != null || m[0] === '<') {
      const text = decodeEntities(m[3] ?? '<').replace(/[ \t\r\n\f]+/g, ' ');
      if (text) raw.push(runFromState(text, top, ctx));
      continue;
    }
    if (m[1] != null) {
      const tag = m[1].toLowerCase();
      for (let i = stack.length - 1; i > 0; i--) {
        if (stack[i].tag === tag) {
          stack.length = i;
          break;
        }
      }
      continue;
    }
    if (m[2] != null) {
      const info = parseOpenTag(m[0]);
      if (info.tag === 'br') {
        raw.push(runFromState('\n', top, ctx));
        continue;
      }
      if (VOID_TAGS.test(info.tag) || /\/>$/.test(m[0])) continue;
      const comp = computeElement(info, ctx, { style: top.style, ancestors: ancestorsOf() });
      const next: InlineState = {
        style: comp.text,
        tag: info.tag,
        info,
        href: top.href,
        code: top.code,
      };
      if (info.tag === 'a' && info.attrs.href) next.href = info.attrs.href;
      if (/^(code|kbd|samp|tt)$/.test(info.tag)) next.code = true;
      stack.push(next);
    }
  }
  const runs = finalizeRuns(raw);
  if (!runs.length) {
    const plain = decodeEntities(html.replace(/<[^>]+>/g, '')).replace(/\s+/g, ' ').trim();
    if (plain) runs.push({ text: plain });
  }
  return runs;
}

/** Preserve <br> and block boundaries as newlines inside a table cell. */
function cellPlainText(html: string): string {
  const withBreaks = html
    .replace(/<\s*br\s*\/?\s*>/gi, '\n')
    .replace(/<\/\s*(p|div|li|h[1-6]|tr)\s*>/gi, '\n')
    .replace(/<\s*(p|div|li|h[1-6])(\s[^>]*)?>/gi, '\n');
  return decodeEntities(withBreaks.replace(/<[^>]+>/g, ''))
    .replace(/\r\n/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function parseTableHtml(tableHtml: string): ContentBlock {
  const rows: string[][] = [];
  const trRe = /<tr\b[^>]*>([\s\S]*?)<\/tr>/gi;
  let tr: RegExpExecArray | null;
  while ((tr = trRe.exec(tableHtml)) !== null) {
    const cells: string[] = [];
    const cellRe = /<(td|th)\b[^>]*>([\s\S]*?)<\/\1>/gi;
    let c: RegExpExecArray | null;
    while ((c = cellRe.exec(tr[1])) !== null) {
      cells.push(cellPlainText(c[2]));
    }
    if (cells.length) rows.push(cells);
  }
  return { type: 'table', rows };
}

function extractCodeLanguage(preOpen: string, codeOpen: string): string | undefined {
  const fromClass =
    /(?:^|\s)class=["'][^"']*\blanguage-([A-Za-z0-9_+-]+)/i.exec(codeOpen)?.[1] ??
    /(?:^|\s)class=["'][^"']*\blanguage-([A-Za-z0-9_+-]+)/i.exec(preOpen)?.[1];
  return fromClass || undefined;
}

function parseListHtml(
  listHtml: string,
  depth: number,
  ordered: boolean,
  ctx: HtmlCtx = defaultCtx(),
  scope: Scope = rootScope()
): ContentBlock[] {
  const blocks: ContentBlock[] = [];
  const inner = listHtml.replace(/^<(ul|ol)\b[^>]*>/i, '').replace(/<\/(ul|ol)\s*>$/i, '');
  let i = 0;
  while (i < inner.length) {
    const liAt = inner.toLowerCase().indexOf('<li', i);
    if (liAt < 0) break;
    const after = inner.charAt(liAt + 3);
    if (after && !/[\s/>]/.test(after)) {
      i = liAt + 3;
      continue;
    }
    const end = findBalancedHtmlEnd(inner, liAt);
    if (end < 0) {
      i = liAt + 3;
      continue;
    }
    const itemXml = inner.slice(liAt, end);
    const openEnd = itemXml.indexOf('>') + 1;
    let body = itemXml.slice(openEnd).replace(/<\/li\s*>$/i, '');

    // Nested lists inside this item.
    const nestedBlocks: ContentBlock[] = [];
    const nestedRe = /<(ul|ol)\b/gi;
    let nm: RegExpExecArray | null;
    let cursor = 0;
    const textParts: string[] = [];
    while ((nm = nestedRe.exec(body)) !== null) {
      textParts.push(body.slice(cursor, nm.index));
      const nEnd = findBalancedHtmlEnd(body, nm.index);
      if (nEnd < 0) break;
      const nestedXml = body.slice(nm.index, nEnd);
      const isOl = /^<ol\b/i.test(nestedXml);
      nestedBlocks.push(...parseListHtml(nestedXml, depth + 1, isOl, ctx, scope));
      cursor = nEnd;
      nestedRe.lastIndex = nEnd;
    }
    textParts.push(body.slice(cursor));
    const liInfo = parseOpenTag(itemXml.slice(0, openEnd));
    const liStyle = computeElement(liInfo, ctx, scope).text;
    const runs = parseInline(textParts.join(''), ctx, {
      style: liStyle,
      ancestors: [...scope.ancestors, liInfo],
    });
    if (runsToText(runs).trim() || !nestedBlocks.length) {
      blocks.push({
        type: 'list_item',
        level: depth,
        ...(ordered ? { ordered: true } : {}),
        runs: runs.length ? runs : [{ text: ' ' }],
      });
    }
    blocks.push(...nestedBlocks);
    i = end;
  }
  return blocks;
}

function layoutFor(comp: ComputedStyle, ctx: HtmlCtx): BlockLayout | undefined {
  if (!ctx.layout) return undefined;
  const l: BlockLayout = {};
  const t = comp.text;
  const b = comp.box;
  if (t.align && t.align !== 'left') l.align = t.align;
  if (b.marginTopPt) l.spaceBeforePt = round2(Math.max(0, b.marginTopPt));
  if (b.marginBottomPt) l.spaceAfterPt = round2(Math.max(0, b.marginBottomPt));
  if (t.lineHeight != null && t.lineHeight !== ctx.base.lineHeight) l.lineHeight = round2(t.lineHeight);
  if (b.marginLeftPt) l.indentLeftPt = round2(b.marginLeftPt);
  if (b.marginRightPt) l.indentRightPt = round2(b.marginRightPt);
  if (t.textIndentPt) l.firstLineIndentPt = round2(t.textIndentPt);
  if (b.pageBreakBefore) l.pageBreakBefore = true;
  return Object.keys(l).length ? l : undefined;
}

type LayoutBlock = Extract<ContentBlock, { type: 'paragraph' | 'heading' }>;

function isLayoutBlock(b: ContentBlock | undefined): b is LayoutBlock {
  return b?.type === 'paragraph' || b?.type === 'heading';
}

function mergeLayout(b: LayoutBlock, patch: BlockLayout): void {
  const l: BlockLayout = { ...(b.layout ?? {}) };
  for (const [k, v] of Object.entries(patch) as Array<[keyof BlockLayout, never]>) {
    if (v == null) continue;
    if (k === 'spaceBeforePt' || k === 'spaceAfterPt') {
      // Parent/child vertical margins collapse (CSS): keep the larger.
      l[k] = Math.max(l[k] ?? 0, v);
    } else if (k === 'indentLeftPt' || k === 'indentRightPt') {
      l[k] = round2((l[k] ?? 0) + (v as number));
    } else if (l[k] == null) {
      (l as Record<string, unknown>)[k] = v;
    }
  }
  b.layout = l;
}

/** Container box (div/section/…): margins go to first/last child, indents to all. */
function applyContainerBox(blocks: ContentBlock[], box: BoxStyle, ctx: HtmlCtx): void {
  if (!ctx.layout || !blocks.length) return;
  const first = blocks[0];
  const last = blocks[blocks.length - 1];
  if (isLayoutBlock(first)) {
    const patch: BlockLayout = {};
    if (box.marginTopPt && box.marginTopPt > 0) patch.spaceBeforePt = round2(box.marginTopPt);
    if (box.pageBreakBefore) patch.pageBreakBefore = true;
    if (Object.keys(patch).length) mergeLayout(first, patch);
  }
  if (isLayoutBlock(last) && box.marginBottomPt && box.marginBottomPt > 0) {
    mergeLayout(last, { spaceAfterPt: round2(box.marginBottomPt) });
  }
  if (box.marginLeftPt || box.marginRightPt) {
    for (const b of blocks) {
      if (!isLayoutBlock(b)) continue;
      mergeLayout(b, {
        ...(box.marginLeftPt ? { indentLeftPt: box.marginLeftPt } : {}),
        ...(box.marginRightPt ? { indentRightPt: box.marginRightPt } : {}),
      });
    }
  }
}

/**
 * Sibling vertical margins collapse in CSS (gap = max) but add up in Word/ODF
 * (gap = after + before). Reduce `spaceBefore` so the rendered gap matches.
 */
function collapseSiblingMargins(blocks: ContentBlock[]): void {
  for (let i = 1; i < blocks.length; i++) {
    const prev = blocks[i - 1];
    const cur = blocks[i];
    if (cur.type === 'blockquote') collapseSiblingMargins(cur.blocks);
    if (!isLayoutBlock(prev) || !isLayoutBlock(cur)) continue;
    const after = prev.layout?.spaceAfterPt ?? 0;
    const before = cur.layout?.spaceBeforePt;
    if (!after || before == null) continue;
    const next = round2(Math.max(0, before - after));
    const l = { ...cur.layout! };
    if (next > 0) l.spaceBeforePt = next;
    else delete l.spaceBeforePt;
    cur.layout = Object.keys(l).length ? l : undefined;
    if (!cur.layout) delete cur.layout;
  }
  if (blocks[0]?.type === 'blockquote') collapseSiblingMargins(blocks[0].blocks);
}

function parseFlow(html: string, ctx: HtmlCtx, scope: Scope): ContentBlock[] {
  const blocks: ContentBlock[] = [];
  let pendingBreak = false;
  const push = (...items: ContentBlock[]) => {
    if (pendingBreak && items.length && isLayoutBlock(items[0])) {
      mergeLayout(items[0], { pageBreakBefore: true });
      pendingBreak = false;
    }
    blocks.push(...items);
  };
  const pushLoose = (chunk: string) => {
    if (!/[^\s]/.test(chunk.replace(/<[^>]+>/g, '').replace(/&nbsp;/gi, ' '))) return;
    const runs = parseInline(chunk, ctx, scope);
    if (!runsToText(runs).trim()) return;
    const layout = layoutFor({ text: scope.style, box: {} }, ctx);
    push({ type: 'paragraph', runs, ...(layout ? { layout } : {}) });
  };

  let i = 0;
  while (i < html.length) {
    const at = indexOfNextBlockTag(html, i);
    if (at < 0) {
      pushLoose(html.slice(i));
      break;
    }
    if (at > i) pushLoose(html.slice(i, at));

    const open = /^<([A-Za-z][\w:-]*)\b[^>]*>/i.exec(html.slice(at));
    if (!open) {
      i = at + 1;
      continue;
    }
    const info = parseOpenTag(open[0]);
    const tag = info.tag;

    if (tag === 'hr') {
      push({ type: 'thematic_break' });
      i = at + open[0].length;
      continue;
    }

    const end = findBalancedHtmlEnd(html, at);
    if (end < 0) {
      i = at + open[0].length;
      continue;
    }
    const el = html.slice(at, end);
    const inner = el.slice(open[0].length).replace(new RegExp(`</${tag}\\s*>$`, 'i'), '');
    const comp = computeElement(info, ctx, scope);
    const child: Scope = { style: comp.text, ancestors: [...scope.ancestors, info] };

    if (/^h[1-6]$/.test(tag)) {
      const level = Number(tag[1]) as 1 | 2 | 3 | 4 | 5 | 6;
      const layout = layoutFor(comp, ctx);
      push({ type: 'heading', level, runs: parseInline(inner, ctx, child), ...(layout ? { layout } : {}) });
    } else if (tag === 'p') {
      const runs = parseInline(inner, ctx, child);
      const layout = layoutFor(comp, ctx);
      if (runsToText(runs).trim()) push({ type: 'paragraph', runs, ...(layout ? { layout } : {}) });
      else if (layout?.pageBreakBefore) pendingBreak = true;
    } else if (CONTAINER_TAGS.has(tag)) {
      const nested = parseFlow(inner, ctx, child);
      applyContainerBox(nested, comp.box, ctx);
      if (nested.length) push(...nested);
      else if (comp.box.pageBreakBefore) pendingBreak = true;
    } else if (tag === 'table') {
      const table = parseTableHtml(el);
      if (table.type === 'table' && table.rows.length) push(table);
    } else if (tag === 'pre') {
      const preOpen = /^<pre\b[^>]*>/i.exec(el)?.[0] ?? '';
      const codeOpen = /<code\b[^>]*>/i.exec(el)?.[0] ?? '';
      const language = extractCodeLanguage(preOpen, codeOpen);
      const code = inner.replace(/^<code\b[^>]*>/i, '').replace(/<\/code\s*>$/i, '');
      push({
        type: 'code',
        text: decodeEntities(code.replace(/<[^>]+>/g, '')),
        ...(language ? { language } : {}),
      });
    } else if (tag === 'ul' || tag === 'ol') {
      push(...parseListHtml(el, 0, tag === 'ol', ctx, scope));
    } else if (tag === 'blockquote') {
      const nested = parseFlow(inner, ctx, child);
      applyContainerBox(nested, { ...comp.box, marginLeftPt: undefined }, ctx);
      if (nested.length) push({ type: 'blockquote', blocks: nested });
    }
    if (comp.box.pageBreakAfter) pendingBreak = true;

    i = end;
  }
  return blocks;
}

function extractStyleSheets(html: string, sheet: CssSheet): string {
  return html.replace(/<style\b[^>]*>([\s\S]*?)<\/style\s*>/gi, (_, css: string) => {
    parseCss(css, sheet);
    return '';
  });
}

function normalizeHtml(html: string): string {
  return html
    .replace(/\r\n/g, '\n')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<!DOCTYPE[^>]*>/gi, '')
    .replace(/<\?xml[^>]*\?>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<binary\b[^>]*>[\s\S]*?<\/binary>/gi, '')
    .replace(/<v\b[^>]*>/gi, '<p>')
    .replace(/<\/v\s*>/gi, '</p>')
    .replace(/<empty-line\b[^>]*\/?>/gi, '<br/>');
}

export type HtmlDocument = {
  blocks: ContentBlock[];
  /** Document defaults (body font/size/color/line-height, @page size/margins). */
  layout: DocumentLayout;
  /** HTML_STYLE_IGNORED when CSS could not be represented. */
  warnings: DocumentWarning[];
};

const IGNORED_LIST_MAX = 12;

function listForWarning(items: string[]): string {
  const shown = items.slice(0, IGNORED_LIST_MAX).join(', ');
  return items.length > IGNORED_LIST_MAX ? `${shown}, … (+${items.length - IGNORED_LIST_MAX})` : shown;
}

/**
 * HTML → blocks with layout for writers (DOCX/ODT honour it): paragraph
 * alignment / spacing / indents / line height / page breaks, run size / font /
 * color / letter-spacing, body defaults and `@page` size + margins.
 * Inline `style`, `<style>` rules (tag, .class, #id, descendant/child) and
 * browser defaults for headings/paragraphs are applied.
 */
export function htmlToDocument(html: string): HtmlDocument {
  return parseHtmlDocument(html, true);
}

function parseHtmlDocument(html: string, layoutMode: boolean): HtmlDocument {
  const sheet = emptySheet();
  let normalized = extractStyleSheets(normalizeHtml(html), sheet);
  normalized = normalized
    .replace(/<head\b[^>]*>[\s\S]*?<\/head\s*>/gi, '')
    .replace(/<title\b[^>]*>[\s\S]*?<\/title\s*>/gi, '')
    .replace(/<(meta|link|base)\b[^>]*>/gi, '');

  const ctx: HtmlCtx = {
    sheet,
    ignored: new Set(),
    layout: layoutMode,
    base: { sizePt: ROOT_FONT_PT },
  };

  const htmlOpen = /<html\b[^>]*>/i.exec(normalized);
  const htmlInfo = parseOpenTag(htmlOpen?.[0] ?? '<html>');
  const bodyMatch = /<body\b[^>]*>/i.exec(normalized);
  const bodyInfo = parseOpenTag(bodyMatch?.[0] ?? '<body>');
  let body = normalized;
  if (bodyMatch) {
    body = normalized.slice(bodyMatch.index + bodyMatch[0].length).replace(/<\/body\s*>[\s\S]*$/i, '');
  }
  body = body.replace(/<\/?html\b[^>]*>/gi, '');

  const htmlComp = computeElement(htmlInfo, ctx, rootScope());
  const bodyComp = computeElement(bodyInfo, ctx, {
    style: htmlComp.text,
    ancestors: [htmlInfo],
  });
  const bt = bodyComp.text;
  ctx.base = { ...bt };

  const layout: DocumentLayout = {};
  if (layoutMode) {
    if (bt.font) layout.font = bt.font;
    layout.sizePt = round2(bt.sizePt);
    if (bt.color) layout.color = bt.color;
    if (bt.lineHeight != null) layout.lineHeight = round2(bt.lineHeight);
    const page = pageLayout(sheet, bodyComp.box);
    if (page) layout.page = page;
  }

  const blocks = parseFlow(body, ctx, {
    style: bt,
    ancestors: [htmlInfo, bodyInfo],
  });
  if (layoutMode) {
    // Body-level text-align etc. already flows to blocks through inheritance.
    collapseSiblingMargins(blocks);
  }

  const warnings: DocumentWarning[] = [];
  const props = [...ctx.ignored].sort();
  if (layoutMode && (props.length || sheet.ignored.length)) {
    const parts: string[] = [];
    if (props.length) parts.push(`CSS properties not applied: ${listForWarning(props)}`);
    if (sheet.ignored.length) parts.push(`selectors/at-rules not applied: ${listForWarning(sheet.ignored)}`);
    warnings.push({
      code: 'HTML_STYLE_IGNORED',
      message: `HTML_STYLE_IGNORED: ${parts.join('; ')}`,
    });
  }
  return { blocks, layout, warnings };
}

function pageLayout(sheet: CssSheet, bodyBox: BoxStyle): DocumentLayout['page'] | undefined {
  const page: NonNullable<DocumentLayout['page']> = {};
  const decls = sheet.page;
  if (decls.size) {
    const size = parsePageSize(decls.size);
    if (size) Object.assign(page, size);
  }
  const box: Record<'top' | 'right' | 'bottom' | 'left', number | undefined> = {
    top: undefined,
    right: undefined,
    bottom: undefined,
    left: undefined,
  };
  const margin = decls.margin?.trim().split(/\s+/);
  if (margin?.length) {
    const vals = margin.length === 1 ? [margin[0], margin[0], margin[0], margin[0]]
      : margin.length === 2 ? [margin[0], margin[1], margin[0], margin[1]]
      : margin.length === 3 ? [margin[0], margin[1], margin[2], margin[1]]
      : margin.slice(0, 4);
    (['top', 'right', 'bottom', 'left'] as const).forEach((side, idx) => {
      box[side] = lengthToPt(vals[idx], ROOT_FONT_PT);
    });
  }
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const v = decls[`margin-${side}`];
    if (v) box[side] = lengthToPt(v, ROOT_FONT_PT);
  }
  const bodyMargins = {
    top: bodyBox.marginTopPt,
    right: bodyBox.marginRightPt,
    bottom: bodyBox.marginBottomPt,
    left: bodyBox.marginLeftPt,
  };
  const anyPage = Object.values(box).some((v) => v != null);
  const anyBody = Object.values(bodyMargins).some((v) => v != null && v > 0);
  if (anyPage || anyBody) {
    // Browser print: page margin (Chromium default ≈ 1cm) + body margin.
    const DEFAULT_PAGE_MARGIN = 28.35;
    for (const side of ['top', 'right', 'bottom', 'left'] as const) {
      const total = (box[side] ?? DEFAULT_PAGE_MARGIN) + Math.max(0, bodyMargins[side] ?? 0);
      const key = `margin${side[0].toUpperCase()}${side.slice(1)}Pt` as
        | 'marginTopPt'
        | 'marginRightPt'
        | 'marginBottomPt'
        | 'marginLeftPt';
      page[key] = round2(total);
    }
  }
  return Object.keys(page).length ? page : undefined;
}

/** Remove writer-only layout so parse/convert results stay content-only. */
function stripLayout(blocks: ContentBlock[]): ContentBlock[] {
  return blocks.map((b) => {
    if (b.type === 'blockquote') return { ...b, blocks: stripLayout(b.blocks) };
    if (b.type === 'paragraph' || b.type === 'heading' || b.type === 'list_item') {
      const { layout: _l, ...rest } = b as LayoutBlock;
      return {
        ...rest,
        runs: b.runs.map(({ sizePt: _s, font: _f, color: _c, letterSpacingPt: _ls, ...r }) => r),
      } as ContentBlock;
    }
    return b;
  });
}

/**
 * Convert HTML snippet into ContentBlocks (structural convert path).
 * Content only: no layout (see {@link htmlToDocument} for writes).
 */
export function htmlToBlocks(html: string): ContentBlock[] {
  return stripLayout(parseHtmlDocument(html, false).blocks);
}

function runsToText(runs: InlineRun[]): string {
  return runs.map((r) => r.text).join('');
}

export function htmlToStructured(html: string): {
  format: 'html';
  paragraphs: string[];
  blocks: ContentBlock[];
  text: string;
} {
  const blocks = htmlToBlocks(html);
  const paragraphs = flattenBlocksToParagraphs(blocks);
  return {
    format: 'html',
    blocks,
    paragraphs,
    text: paragraphs.join('\n\n'),
  };
}

/** Lightweight MD → blocks (headings, lists, GFM tables, paragraphs). */
export function markdownToBlocks(md: string): ContentBlock[] {
  const lines = md.replace(/\r\n/g, '\n').split('\n');
  const blocks: ContentBlock[] = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (/^\s*$/.test(line)) {
      i++;
      continue;
    }
    const hm = /^(#{1,6})\s+(.*)$/.exec(line);
    if (hm) {
      blocks.push({
        type: 'heading',
        level: hm[1].length as 1 | 2 | 3 | 4 | 5 | 6,
        runs: parseMarkdownInline(hm[2]),
      });
      i++;
      continue;
    }
    if (/^---+$/.test(line.trim())) {
      blocks.push({ type: 'thematic_break' });
      i++;
      continue;
    }
    if (/^\s*>\s?/.test(line)) {
      const quoteLines: string[] = [];
      while (i < lines.length && /^\s*>\s?/.test(lines[i])) {
        quoteLines.push(lines[i].replace(/^\s*>\s?/, ''));
        i++;
      }
      const nested = markdownToBlocks(quoteLines.join('\n'));
      blocks.push({
        type: 'blockquote',
        blocks: nested.length
          ? nested
          : [{ type: 'paragraph', runs: parseMarkdownInline(quoteLines.join(' ')) }],
      });
      continue;
    }
    if (/^\s*(?:[-*+]|\d+\.)\s+/.test(line)) {
      while (i < lines.length && /^\s*(?:[-*+]|\d+\.)\s+/.test(lines[i])) {
        const m = /^(\s*)([-*+]|\d+\.)\s+(.*)$/.exec(lines[i])!;
        const level = Math.floor(m[1].replace(/\t/g, '  ').length / 2);
        const ordered = /^\d+\.$/.test(m[2]);
        blocks.push({
          type: 'list_item',
          level,
          ...(ordered ? { ordered: true } : {}),
          runs: parseMarkdownInline(m[3]),
        });
        i++;
      }
      continue;
    }
    if (/^\s*\|/.test(line) && i + 1 < lines.length && /^\s*\|?\s*:?-+:?/.test(lines[i + 1])) {
      const rows: string[][] = [];
      while (i < lines.length && /^\s*\|/.test(lines[i])) {
        if (/^\s*\|?\s*:?-+:?/.test(lines[i])) {
          i++;
          continue;
        }
        const cells = lines[i]
          .replace(/^\s*\|/, '')
          .replace(/\|\s*$/, '')
          .split('|')
          .map((c) => c.trim());
        rows.push(cells);
        i++;
      }
      if (rows.length) blocks.push({ type: 'table', rows });
      continue;
    }
    if (/^```/.test(line)) {
      const lang = line.replace(/^```/, '').trim();
      i++;
      const buf: string[] = [];
      while (i < lines.length && !/^```/.test(lines[i])) {
        buf.push(lines[i]);
        i++;
      }
      if (i < lines.length) i++;
      blocks.push({ type: 'code', text: buf.join('\n'), language: lang || undefined });
      continue;
    }
    const para: string[] = [line];
    i++;
    while (
      i < lines.length &&
      !/^\s*$/.test(lines[i]) &&
      !/^(#{1,6})\s+|^\s*>|^\s*[-*+]\s+|^\s*\d+\.\s+|^\s*\||^```/.test(lines[i])
    ) {
      para.push(lines[i]);
      i++;
    }
    blocks.push({ type: 'paragraph', runs: parseMarkdownInline(para.join(' ')) });
  }
  return blocks;
}

/**
 * Parse inline markdown emphasis/links/code.
 * Supports both *italic* / _italic_ and ~~strike~~ so emit↔parse stay aligned.
 */
export function parseMarkdownInline(text: string): InlineRun[] {
  const runs: InlineRun[] = [];
  // Order: code, links, bold+italic combos, bold, strike, italic, plain.
  const re =
    /(`([^`]+)`)|(\[([^\]]+)\]\(([^)]+)\))|(\*\*_([^*]+)_\*\*)|(\*\*([^*]+)\*\*)|(~~([^~]+)~~)|(\*([^*]+)\*)|((?<!\w)_([^_]+)_(?!\w))|([^*`[~_\]]+|[_*~\[\]])/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    if (m[2]) runs.push({ text: m[2], code: true });
    else if (m[4] && m[5]) runs.push({ text: m[4], href: m[5] });
    else if (m[7]) runs.push({ text: m[7], bold: true, italic: true });
    else if (m[9]) runs.push({ text: m[9], bold: true });
    else if (m[11]) runs.push({ text: m[11], strike: true });
    else if (m[13]) runs.push({ text: m[13], italic: true });
    else if (m[15]) runs.push({ text: m[15], italic: true });
    else if (m[16]) runs.push({ text: m[16] });
  }
  if (!runs.length && text) runs.push({ text });
  return mergeAdjacentPlain(runs);
}

function mergeAdjacentPlain(runs: InlineRun[]): InlineRun[] {
  if (runs.length < 2) return runs;
  const out: InlineRun[] = [{ ...runs[0] }];
  for (let i = 1; i < runs.length; i++) {
    const prev = out[out.length - 1];
    const cur = runs[i];
    const plain =
      !prev.bold &&
      !prev.italic &&
      !prev.strike &&
      !prev.code &&
      !prev.href &&
      !cur.bold &&
      !cur.italic &&
      !cur.strike &&
      !cur.code &&
      !cur.href;
    if (plain) prev.text += cur.text;
    else out.push({ ...cur });
  }
  return out;
}
