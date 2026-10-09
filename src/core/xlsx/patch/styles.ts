/**
 * styles.xml model for the package patcher (owner: patcher P).
 *  - Excel "format same as left/above" for inserted cells: the neighbour's xf, but a border side survives only
 *    when BOTH neighbours (before / after the insertion point) have the identical side (verified with Excel:
 *    tests/golden/com/probe-insert-format.ps1).
 *  - Premerge: import an xf of another package (numFmt, font, fill, border, alignment, protection) into this
 *    styles.xml, deduplicated; theme colours are converted to RGB when the two themes differ.
 * Records are appended; existing records are never rewritten, so every existing s="n" keeps its meaning.
 */
import { JA_BUILTIN_NUMFMT } from '../../display/numfmt';
import { dateModeOf } from '../styles';
import { DEFAULT_INDEXED, applyTint, parseThemeColors } from '../theme';
import { findAll, findFirst, getAttr, setAttr } from './xml';

const SECTIONS = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs'] as const;
type Section = (typeof SECTIONS)[number];
const CHILD: Record<Section, string> = {
  numFmts: 'numFmt', fonts: 'font', fills: 'fill', borders: 'border', cellStyleXfs: 'xf', cellXfs: 'xf', cellStyles: 'cellStyle', dxfs: 'dxf',
};

/** Built-in number formats whose meaning does not depend on the locale settings of the book. */
export const BUILTIN_NUMFMT_MAX = 163;

export class Styles {
  xml: string;
  private items: Partial<Record<Section, string[]>> = {};
  private added: Partial<Record<Section, string[]>> = {};
  private index = new Map<string, number>();
  private numFmtCodes = new Map<number, string>();
  private numFmtByCode = new Map<string, number>();
  /** ids of every <numFmt> in the file, empty codes included: a new custom format never reuses one */
  private usedNumFmtIds = new Set<number>();
  private insertCache = new Map<string, number>();
  changed = false;

  constructor(xml: string | null) {
    this.xml = xml && /<styleSheet\b/.test(xml) ? xml : DEFAULT_STYLES;
    for (const s of SECTIONS) {
      const m = findFirst(this.xml, s);
      this.items[s] = m && !m.selfClosing ? findAll(m.inner, CHILD[s]).map((e) => e.text) : [];
      this.added[s] = [];
    }
    for (const nf of this.items.numFmts!) {
      const id = Number(getAttr(nf, 'numFmtId'));
      const code = getAttr(nf, 'formatCode') ?? '';
      this.usedNumFmtIds.add(id);
      if (code.trim() === '') continue; // an empty formatCode defines nothing (the reader ignores it too)
      this.numFmtCodes.set(id, code);
      if (!this.numFmtByCode.has(code)) this.numFmtByCode.set(code, id);
    }
  }

  list(s: Section): string[] {
    return this.items[s]!.concat(this.added[s]!);
  }

  get(s: Section, i: number): string | undefined {
    const a = this.items[s]!;
    return i < a.length ? a[i] : this.added[s]![i - a.length];
  }

  count(s: Section): number {
    return this.items[s]!.length + this.added[s]!.length;
  }

  /** index of an equal record (canonical compare) or append it */
  add(s: Section, el: string): number {
    const key = s + '\u0000' + canon(el);
    const ex = this.index.get(key);
    if (ex !== undefined) return ex;
    const all = this.list(s);
    for (let i = 0; i < all.length; i++) {
      const k = s + '\u0000' + canon(all[i]);
      if (!this.index.has(k)) this.index.set(k, i);
    }
    const again = this.index.get(key);
    if (again !== undefined) return again;
    const i = this.count(s);
    this.added[s]!.push(el);
    this.index.set(key, i);
    this.changed = true;
    return i;
  }

  numFmtCode(id: number): string | undefined {
    return this.numFmtCodes.get(id);
  }

  /**
   * What the number format of xf `i` shows, by the reader's rules (styles.ts dateModeOf): 'date' (a calendar date, maybe with a
   * time), 'time' (time of day only), 'elapsed' ([h]:mm durations) or 'none'. The 1904 <-> 1900 serial conversion of the reader
   * shifts exactly the 'date' values and the 'time' values of 1 and more, so every writer here asks the same question.
   */
  dateMode(i: number): 'none' | 'date' | 'time' | 'elapsed' {
    const id = Number(getAttr(openTag(this.xf(i)), 'numFmtId') ?? 0);
    const code = this.numFmtCodes.get(id) ?? JA_BUILTIN_NUMFMT[id] ?? 'General';
    return dateModeOf(code);
  }

  /**
   * The legacy indexed palette (colours 0..63, 'RRGGBB') this styles.xml resolves `indexed="n"` colours with: its
   * <colors><indexedColors> entries over Excel's default palette (same rule as the reader, theme.ts).
   */
  effectivePalette(): string[] {
    const m = /<indexedColors\b[^>]*>([\s\S]*?)<\/indexedColors>/.exec(this.xml);
    const custom = m ? [...m[1].matchAll(/<rgbColor\b[^>]*\srgb="([0-9A-Fa-f]{6,8})"/g)].map((x) => x[1].slice(-6).toUpperCase()) : [];
    return Array.from({ length: 64 }, (_v, i) => custom[i] ?? DEFAULT_INDEXED[i]);
  }

  /** id for a custom number format code (existing id or a new one >= 164) */
  numFmtId(code: string): number {
    const ex = this.numFmtByCode.get(code);
    if (ex !== undefined) return ex;
    let id = 164;
    for (const k of this.numFmtCodes.keys()) if (k >= id) id = k + 1;
    while (this.usedNumFmtIds.has(id)) id++;
    this.usedNumFmtIds.add(id);
    this.numFmtCodes.set(id, code);
    this.numFmtByCode.set(code, id);
    this.added.numFmts!.push(`<numFmt numFmtId="${id}" formatCode="${escA(code)}"/>`);
    this.changed = true;
    return id;
  }

  xf(i: number): string {
    return this.get('cellXfs', i) ?? this.get('cellXfs', 0) ?? '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
  }

  /**
   * CellWrite.numFmt: xf `i` with number format `code` (Excel sets Range.NumberFormat when a date / % / currency is
   * typed into a General cell). Built-in ja-JP codes keep their built-in id (like Excel writes them), other codes
   * reuse or create a custom id >= 164. An identical existing xf is reused.
   */
  withNumFmt(i: number, code: string): number {
    const c = code.trim();
    if (!c) return i;
    const key = 'nf:' + i + '\u0000' + c;
    const hit = this.insertCache.get(key);
    if (hit !== undefined) return hit;
    const builtin = BUILTIN_BY_CODE.get(c.toLowerCase() === 'general' ? 'General' : c);
    const id = builtin ?? this.numFmtId(c);
    const el = this.xf(i);
    const open = openTag(el);
    let res = i;
    if (Number(getAttr(open, 'numFmtId') ?? 0) !== id || (i > 0 && !this.get('cellXfs', i))) {
      let o = setAttr(open, 'numFmtId', String(id));
      o = setAttr(o, 'applyNumberFormat', id ? '1' : null);
      res = this.add('cellXfs', o + el.slice(open.length));
    }
    this.insertCache.set(key, res);
    return res;
  }

  /**
   * xf `i` with <alignment horizontal="..."> (other alignment attributes kept: vertical, wrapText, rotation ...;
   * indent only survives with left / right / distributed, as in Excel). Used by the 'unmerge' op
   * (「選択範囲内で中央」 = centerContinuous).
   */
  withHorizontal(i: number, horizontal: string): number {
    return this.withAlignment(i, 'ha:' + i + '\u0000' + horizontal, (aopen) => {
      const na = setAttr(aopen, 'horizontal', horizontal);
      return /^(left|right|distributed)$/.test(horizontal) ? na : setAttr(na, 'indent', null);
    });
  }

  /** true when xf `i` has <alignment wrapText="1"> (the cell wraps its text at the cell edge / at line breaks). */
  hasWrapText(i: number): boolean {
    const v = getAttr(/<alignment\b[^>]*>/.exec(this.xf(i))?.[0] ?? '', 'wrapText');
    return v === '1' || v === 'true';
  }

  /**
   * xf `i` with <alignment wrapText="1"> (everything else of the xf kept). Excel does this on its own when a text that
   * contains a line feed is written into a cell (CellWrite, see Engine.setCell). Nothing to change = `i` itself.
   */
  withWrapText(i: number): number {
    if (this.hasWrapText(i)) return i;
    return this.withAlignment(i, 'wt:' + i, (aopen) => setAttr(aopen, 'wrapText', '1'));
  }

  /** xf `i` with its <alignment> start tag rewritten by `edit` (created when the xf has none; applyAlignment set). */
  private withAlignment(i: number, key: string, edit: (alignmentOpenTag: string) => string): number {
    const hit = this.insertCache.get(key);
    if (hit !== undefined) return hit;
    const el = this.xf(i);
    const open = openTag(el);
    const inner = el.endsWith('/>') && el === open ? '' : el.slice(open.length).replace(/<\/xf>\s*$/, '');
    const am = /<alignment\b[^>]*?(?:\/>|>[\s\S]*?<\/alignment>)/.exec(inner);
    let align = am ? am[0] : '<alignment/>';
    const aopen = /^<alignment\b[^>]*?(\/?)>/.exec(align)![0];
    align = edit(aopen) + align.slice(aopen.length);
    const nInner = am ? inner.slice(0, am.index) + align + inner.slice(am.index + am[0].length) : align + inner;
    let o = setAttr(open.endsWith('/>') ? open.slice(0, -2).replace(/\s+$/, '') + '>' : open, 'applyAlignment', '1');
    if (o.endsWith('/>')) o = o.slice(0, -2) + '>';
    const res = this.add('cellXfs', `${o}${nInner}</xf>`);
    this.insertCache.set(key, res);
    return res;
  }

  /** true when xf `i` carries quotePrefix (the apostrophe Excel shows in the formula bar). */
  hasQuotePrefix(i: number): boolean {
    const v = getAttr(openTag(this.xf(i)), 'quotePrefix');
    return v === '1' || v === 'true';
  }

  /**
   * CellWrite.cell.quote: xf `i` with quotePrefix="1" (`on`) or without it. Everything else of the xf (number format
   * '@' included) is kept; an identical existing xf is reused. Nothing to change = `i` itself.
   */
  withQuotePrefix(i: number, on: boolean): number {
    if (this.hasQuotePrefix(i) === on) return i;
    const key = 'qp:' + i + '\u0000' + (on ? 1 : 0);
    const hit = this.insertCache.get(key);
    if (hit !== undefined) return hit;
    const el = this.xf(i);
    const open = openTag(el);
    const res = this.add('cellXfs', setAttr(open, 'quotePrefix', on ? '1' : null) + el.slice(open.length));
    this.insertCache.set(key, res);
    return res;
  }

  /**
   * What an 'unmerge' spreads from the top-left cell's xf `tl` over the other cells of the former area: its FILL (when
   * it has one) and each border side that is set. null = nothing to spread.
   */
  mergeLook(tl: number): MergeLook | null {
    const open = openTag(this.xf(tl));
    const fillId = Number(getAttr(open, 'fillId') ?? 0);
    const fillEl = fillId ? this.get('fills', fillId) : undefined;
    const fill = fillEl && /<gradientFill\b|<patternFill\b[^>]*\spatternType="(?!none")/.test(fillEl) ? fillId : null;
    const border = this.borderOf(tl);
    const side = (s: string): string | null => {
      const el = new RegExp(`<${s}\\b[^>]*?(?:/>|>[\\s\\S]*?</${s}>)`).exec(border)?.[0];
      return el && /\sstyle="(?!none")/.test(el) ? el : null;
    };
    const sides = { top: side('top'), bottom: side('bottom'), left: side('left'), right: side('right') };
    if (fill === null && !sides.top && !sides.bottom && !sides.left && !sides.right) return null;
    return { fill, sides };
  }

  /**
   * xf `i` (a cell of a former merged area) with the area's top-left look: its fill, and the border sides the cell
   * lies on (`edges`: top = first row of the area, bottom = last row, left = first column, right = last column; a
   * side the top-left cell does not have leaves the cell's own side alone). Everything else of the xf is kept.
   */
  withMergeLook(i: number, look: MergeLook, edges: { top: boolean; bottom: boolean; left: boolean; right: boolean }): number {
    const key = 'ml:' + i + '\u0000' + look.fill + '\u0000' + (['top', 'bottom', 'left', 'right'] as const).map((s) => (edges[s] ? look.sides[s] ?? '' : '')).join('\u0001');
    const hit = this.insertCache.get(key);
    if (hit !== undefined) return hit;
    const el = this.xf(i);
    let open = openTag(el);
    let changed = false;
    if (look.fill !== null && Number(getAttr(open, 'fillId') ?? 0) !== look.fill) {
      open = setAttr(setAttr(open, 'fillId', String(look.fill)), 'applyFill', '1');
      changed = true;
    }
    let border = this.borderOf(i);
    for (const s of ['left', 'right', 'top', 'bottom'] as const) {
      const src = look.sides[s];
      if (!src || !edges[s]) continue;
      const nb = withBorderSide(border, s, src);
      if (canon(nb) !== canon(border)) { border = nb; changed = true; }
    }
    let res = i;
    if (changed) {
      const bid = this.add('borders', border);
      if (bid !== Number(getAttr(open, 'borderId') ?? 0)) open = setAttr(setAttr(open, 'borderId', String(bid)), 'applyBorder', bid ? '1' : null);
      res = this.add('cellXfs', open + el.slice(openTag(el).length));
    }
    this.insertCache.set(key, res);
    return res;
  }

  /** premerge: import font `id` of `src` (phonetic properties of shared strings refer to a font id). */
  importFont(src: Styles, id: number, themeRgb: string[] | null, indexedRgb: string[] | null = null): number | null {
    const f = src.get('fonts', id);
    if (f === undefined) return null;
    return this.add('fonts', convertColors(f, themeRgb, indexedRgb));
  }

  /** Font name / size (points) of xf `i` (font 0 = the workbook default font when i = 0). */
  fontOf(i: number): { name?: string; size?: number } {
    const id = Number(getAttr(openTag(this.xf(i)), 'fontId') ?? 0);
    const f = this.get('fonts', id) ?? this.get('fonts', 0) ?? '';
    const name = /<name\b[^>]*\sval="([^"]*)"/.exec(f)?.[1];
    const sz = Number(/<sz\b[^>]*\sval="([^"]*)"/.exec(f)?.[1]);
    return { name, size: Number.isFinite(sz) && sz > 0 ? sz : undefined };
  }

  borderOf(xfIndex: number): string {
    const id = Number(getAttr(openTag(this.xf(xfIndex)), 'borderId') ?? 0);
    return this.get('borders', id) ?? '<border/>';
  }

  /**
   * Style of a cell inserted between `before` (left / above) and `after` (right / below): the xf of `before`
   * with only the border sides both neighbours share. 0 = no style.
   */
  insertedStyle(before: number, after: number): number {
    if (!before) return 0;
    const key = before + ':' + after;
    const c = this.insertCache.get(key);
    if (c !== undefined) return c;
    const xa = this.xf(before);
    const ba = Number(getAttr(openTag(xa), 'borderId') ?? 0);
    const bb = after ? Number(getAttr(openTag(this.xf(after)), 'borderId') ?? 0) : 0;
    let res = before;
    if (ba !== bb) {
      const A = this.get('borders', ba) ?? '<border/>';
      const B = after ? this.get('borders', bb) ?? '<border/>' : '<border/>';
      const nb = intersectBorder(A, B);
      if (canon(nb) !== canon(A)) {
        const bid = this.add('borders', nb);
        let open = setAttr(openTag(xa), 'borderId', String(bid));
        if (bid === 0) open = setAttr(open, 'applyBorder', null);
        res = this.add('cellXfs', open + xa.slice(openTag(xa).length));
      }
    }
    this.insertCache.set(key, res);
    return res;
  }

  serialize(): string {
    if (!this.changed) return this.xml;
    let x = this.xml;
    for (const s of SECTIONS) {
      const add = this.added[s]!;
      if (!add.length) continue;
      const m = findFirst(x, s);
      const n = this.count(s);
      if (m && !m.selfClosing) {
        const open = /\scount="/.test(m.open) ? setAttr(m.open, 'count', String(n)) : m.open;
        x = x.slice(0, m.start) + open + m.inner + add.join('') + `</${s}>` + x.slice(m.end);
      } else {
        const el = `<${s} count="${n}">${add.join('')}</${s}>`;
        if (m) x = x.slice(0, m.start) + el + x.slice(m.end);
        else x = insertSection(x, s, el);
      }
    }
    return x;
  }

  // -----------------------------------------------------------------------------------------------
  // premerge: import an xf from another styles.xml
  // -----------------------------------------------------------------------------------------------

  /**
   * Import xf `i` of `src` into this styles.xml. `themeRgb`: the source theme colours ('RRGGBB', theme-index
   * order) when they differ from ours — theme colour references are then converted to RGB (with tint) and theme
   * font schemes are dropped so the imported cells look the same as in the source file.
   */
  importXf(src: Styles, i: number, themeRgb: string[] | null, cache: Map<number, number>, indexedRgb: string[] | null = null): number {
    const c = cache.get(i);
    if (c !== undefined) return c;
    const xf = src.xf(i);
    const open = openTag(xf);
    const attr = (n: string) => Number(getAttr(open, n) ?? 0);
    // number format
    let nf = attr('numFmtId');
    // a <numFmt> of the source that defines this id wins, whatever the id (write-excel-file numbers its custom formats from 100:
    // below 164, where the built-in ids live; the reader resolves them the same way, styles.ts numFmtCode)
    const custom = src.numFmtCode(nf);
    if (custom !== undefined) nf = this.numFmtId(custom);
    else if (nf > BUILTIN_NUMFMT_MAX) nf = 0;
    const conv = (el: string) => convertColors(el, themeRgb, indexedRgb);
    const font = this.add('fonts', conv(src.get('fonts', attr('fontId')) ?? src.get('fonts', 0) ?? '<font/>'));
    const fill = this.add('fills', conv(src.get('fills', attr('fillId')) ?? '<fill><patternFill patternType="none"/></fill>'));
    const border = this.add('borders', conv(src.get('borders', attr('borderId')) ?? '<border/>'));
    let o = '<xf';
    o += ` numFmtId="${nf}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0"`;
    if (getAttr(open, 'quotePrefix') === '1') o += ' quotePrefix="1"';
    if (nf) o += ' applyNumberFormat="1"';
    if (font) o += ' applyFont="1"';
    if (fill) o += ' applyFill="1"';
    if (border) o += ' applyBorder="1"';
    const inner = xf.slice(open.length).replace(/<\/xf>\s*$/, '');
    const align = /<alignment\b[^>]*\/>|<alignment\b[^>]*>[\s\S]*?<\/alignment>/.exec(inner)?.[0] ?? '';
    const prot = /<protection\b[^>]*\/>|<protection\b[^>]*>[\s\S]*?<\/protection>/.exec(inner)?.[0] ?? '';
    if (align) o += ' applyAlignment="1"';
    if (prot) o += ' applyProtection="1"';
    const el = align || prot ? `${o}>${align}${prot}</xf>` : `${o}/>`;
    const r = this.add('cellXfs', el);
    cache.set(i, r);
    return r;
  }
}

/** true when the xf's number format shows a date / time (the cell value is a date serial); a custom <numFmt> of ANY id counts (the reader's rule). */
export function isDateXf(st: Styles, xf: number): boolean {
  return st.dateMode(xf) !== 'none';
}

/** Date / time number-format code: y m d h s (or the Japanese era g / e) outside quotes, escapes and [...] tags. */
export function isDateCode(code: string): boolean {
  if (/^general$/i.test(code.trim())) return false;
  const first = code.split(';')[0];
  const bare = first
    .replace(/"[^"]*"/g, '')          // quoted literals
    .replace(/\\./g, '')              // escaped characters (\h)
    .replace(/_.|\*./g, '')           // padding / fill characters
    .replace(/\[(?!h+\]|m+\]|s+\])[^\]]*\]/gi, '') // [Red], [$-411] ... but keep elapsed [h] [mm] [ss]
    .replace(/[eE][+-]/g, '');        // scientific notation 0.00E+00
  return /[ymdhsge]/i.test(bare);
}

function openTag(el: string): string {
  return /^<[^>]*>/.exec(el)?.[0] ?? el;
}

function escA(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** attribute-order independent canonical form */
export function canon(x: string): string {
  return x.replace(/<([\w:]+)((?:\s+[\w:]+\s*=\s*"[^"]*")*)\s*(\/?)>/g, (_m, tag: string, attrs: string, sc: string) => {
    const list = [...attrs.matchAll(/([\w:]+)\s*=\s*"([^"]*)"/g)].map((a) => `${a[1]}="${a[2]}"`).sort();
    return `<${tag}${list.length ? ' ' + list.join(' ') : ''}${sc}>`;
  }).replace(/>\s+</g, '><').replace(/<(\w+)([^>]*)><\/\1>/g, '<$1$2/>');
}

/** What an 'unmerge' spreads over a former merged area: the top-left cell's fill id and its border side elements. */
export interface MergeLook {
  /** fillId of the top-left cell's fill (null = it has none: nothing to spread) */
  fill: number | null;
  /** the <top> / <bottom> / <left> / <right> elements of its border that have a style (null = side not set) */
  sides: { top: string | null; bottom: string | null; left: string | null; right: string | null };
}

/** Schema order of the children of <border>. */
const BORDER_ORDER = ['start', 'end', 'left', 'right', 'top', 'bottom', 'diagonal', 'vertical', 'horizontal'];

/** Border element `border` with its `side` child replaced by `el` (children stay in schema order). */
export function withBorderSide(border: string, side: string, el: string): string {
  let b = border.trim();
  if (/^<border\b[^>]*\/>$/.test(b)) b = b.replace(/\s*\/>$/, '><left/><right/><top/><bottom/><diagonal/></border>');
  const re = new RegExp(`<${side}\\b[^>]*?(?:/>|>[\\s\\S]*?</${side}>)`);
  if (re.test(b)) return b.replace(re, () => el);
  // the child is missing: insert before the first later child, else before </border>
  for (const next of BORDER_ORDER.slice(BORDER_ORDER.indexOf(side) + 1)) {
    const m = new RegExp(`<${next}\\b`).exec(b);
    if (m) return b.slice(0, m.index) + el + b.slice(m.index);
  }
  return b.replace(/<\/border>\s*$/, el + '</border>');
}

/** Border with only the sides present identically in both a and b. */
function intersectBorder(a: string, b: string): string {
  const side = (x: string, s: string) => new RegExp(`<${s}\\b[^>]*?(?:/>|>[\\s\\S]*?</${s}>)`).exec(x)?.[0];
  const has = (el: string | undefined) => !!el && /\sstyle="/.test(el) && !/\sstyle="none"/.test(el);
  const open = openTag(a).endsWith('/>') ? '<border>' : openTag(a).replace(/\s(diagonalUp|diagonalDown)="[^"]*"/g, '');
  let out = open;
  for (const s of ['left', 'right', 'top', 'bottom', 'diagonal']) {
    const sa = side(a, s), sb = side(b, s);
    const keep = has(sa) && has(sb) && canon(sa!) === canon(sb!);
    out += keep ? sa! : `<${s}/>`;
  }
  // diagonal flags only survive with a kept diagonal
  if (/<diagonal\b[^>]*style=/.test(out)) {
    const up = getAttr(openTag(a), 'diagonalUp'), down = getAttr(openTag(a), 'diagonalDown');
    let o = openTag(out);
    if (up) o = setAttr(o, 'diagonalUp', up);
    if (down) o = setAttr(o, 'diagonalDown', down);
    out = o + out.slice(openTag(out).length);
  }
  return out + '</border>';
}

/**
 * Built-in number formats as Japanese Excel shows them (JA_BUILTIN_NUMFMT) -> id, for codes that belong to exactly
 * one built-in id (Excel stores e.g. yyyy/m/d as numFmtId 14 and 0% as 9 instead of a custom format).
 */
const BUILTIN_BY_CODE: Map<string, number> = (() => {
  const seen = new Map<string, number[]>();
  for (const [k, v] of Object.entries(JA_BUILTIN_NUMFMT)) {
    const l = seen.get(v) ?? [];
    l.push(Number(k));
    seen.set(v, l);
  }
  const out = new Map<string, number>();
  for (const [code, ids] of seen) if (ids.length === 1) out.set(code, ids[0]);
  return out;
})();

/** Replace theme colour references (theme="n" [tint]) by rgb; drop font <scheme> (theme fonts). */
export function themeToRgb(el: string, theme: string[]): string {
  let x = el.replace(/<(color|fgColor|bgColor)\b([^>]*?)(\/?)>/g, (m0, tag: string, attrs: string, sc: string) => {
    const t = getAttr(attrs, 'theme');
    if (t == null) return m0;
    const base = theme[Number(t)];
    if (!base) return m0;
    const tint = Number(getAttr(attrs, 'tint') ?? 0);
    const rgb = tint ? applyTint(base, tint) : base;
    let a = setAttr(setAttr(attrs, 'theme', null), 'tint', null);
    a = setAttr(a, 'rgb', 'FF' + rgb.replace(/^#/, '').toUpperCase());
    return `<${tag}${a}${sc}>`;
  });
  x = x.replace(/<scheme\b[^>]*\/>/g, '');
  return x;
}

/**
 * Colour references of a source font / fill / border / rich-text run that would change meaning in the base package, as RGB:
 * theme colours (when the two themes differ) and indexed colours (when the two legacy palettes differ).
 */
export function convertColors(el: string, themeRgb: string[] | null, indexedRgb: string[] | null): string {
  let x = themeRgb ? themeToRgb(el, themeRgb) : el;
  if (indexedRgb) x = indexedToRgb(x, indexedRgb);
  return x;
}

/** `indexed="n"` (0..63) colours of an element as rgb, by `palette` (64 / 65 are the system colours and stay). */
export function indexedToRgb(el: string, palette: readonly string[]): string {
  if (el.indexOf('indexed') < 0) return el;
  return el.replace(/<(color|fgColor|bgColor)\b([^>]*?)(\/?)>/g, (m0, tag: string, attrs: string, sc: string) => {
    const n = getAttr(attrs, 'indexed');
    if (n == null) return m0;
    const i = Number(n);
    const rgb = Number.isInteger(i) && i >= 0 && i < palette.length ? palette[i] : undefined;
    if (!rgb) return m0;
    return `<${tag}${setAttr(setAttr(attrs, 'indexed', null), 'rgb', 'FF' + rgb.replace(/^#/, '').toUpperCase())}${sc}>`;
  });
}

/** Theme colours of a package theme part (theme-index order). */
export function themeColors(xml: string | null): string[] {
  return parseThemeColors(xml ?? null);
}

/** Section order of a styleSheet (for inserting a missing section). */
const ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];

function insertSection(x: string, s: string, el: string): string {
  const idx = ORDER.indexOf(s);
  for (let k = idx + 1; k < ORDER.length; k++) {
    const m = new RegExp(`<${ORDER[k]}\\b`).exec(x);
    if (m) return x.slice(0, m.index) + el + x.slice(m.index);
  }
  return x.replace(/<\/styleSheet>\s*$/, el + '</styleSheet>');
}

const DEFAULT_STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="游ゴシック"/><family val="2"/><charset val="128"/><scheme val="minor"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>';
