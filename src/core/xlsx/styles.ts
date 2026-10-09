/**
 * styles.xml reader (owner: reader R).
 *
 * Two layers:
 *  1. parseStyleSheet(xml) -> StyleSheet: the RAW records (numFmts, fonts, fills, borders, cellStyleXfs, cellXfs,
 *     cellStyles, dxfs, indexed palette), each with its verbatim XML (`xml`) and the parsed attributes. This is the
 *     writing side: the package patcher may reuse it to merge source styles into a base styles.xml (dedupe by
 *     `xml`, remap fontId/fillId/borderId/numFmtId, append) without re-serialising anything it did not change.
 *     Record order = file order, so array index = the id used in the file (fontId, fillId, borderId, xf index).
 *  2. buildDisplayStyles(sheet, palette) -> XStyle[] (index = cellXfs index) plus a per-xf "is date format" flag;
 *     colours resolved (theme + tint, indexed, rgb) to '#RRGGBB', built-in number format ids resolved to their
 *     ja-JP codes (JA_BUILTIN_NUMFMT).
 *
 * applyX semantics (verified against Excel via COM, tests/fixtures/xlsx/apply-probe.xlsx + excel-probes.json):
 * see effectiveXf().
 *
 * Pure TS, never throws on malformed input (missing parts fall back to Excel defaults).
 */
import { SSF } from 'xlsx';
import type { XBorderSide, XBorderStyle, XStyle } from '../types';
import { JA_BUILTIN_NUMFMT } from '../display/numfmt';
import { DEFAULT_INDEXED, DEFAULT_PALETTE, rawColorOf, resolveColor, type Palette, type RawColor } from './theme';
import { boolAttr, kid, kids, numAttr, parseXml, rootOf, decodeOoxmlEscapes, type XEl } from './xml';

// ---------------------------------------------------------------------------------------------
// Raw records
// ---------------------------------------------------------------------------------------------

export interface RawNumFmt {
  id: number;
  code: string;
  xml: string;
}

export interface RawFont {
  xml: string;
  bold?: boolean;
  italic?: boolean;
  /** 'single' | 'double' | 'singleAccounting' | 'doubleAccounting' | 'none' */
  underline?: string;
  strike?: boolean;
  /** points */
  size?: number;
  name?: string;
  color?: RawColor;
  family?: number;
  charset?: number;
  /** 'major' | 'minor' | 'none' */
  scheme?: string;
  /** 'superscript' | 'subscript' | 'baseline' */
  vertAlign?: string;
}

export interface RawFill {
  xml: string;
  /** patternFill patternType ('none', 'solid', 'gray125', ...); undefined for gradient fills */
  patternType?: string;
  fg?: RawColor;
  bg?: RawColor;
  /** gradient fill stops (position 0..1) */
  gradient?: Array<{ position: number; color?: RawColor }>;
}

export interface RawBorderSide {
  style?: string;
  color?: RawColor;
}

export interface RawBorder {
  xml: string;
  left?: RawBorderSide;
  right?: RawBorderSide;
  top?: RawBorderSide;
  bottom?: RawBorderSide;
  diagonal?: RawBorderSide;
  diagonalUp?: boolean;
  diagonalDown?: boolean;
}

export interface RawAlignment {
  horizontal?: string;
  vertical?: string;
  wrapText?: boolean;
  indent?: number;
  shrinkToFit?: boolean;
  textRotation?: number;
  readingOrder?: number;
}

export interface RawXf {
  xml: string;
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  /** parent cell style (cellStyleXfs index); cellXfs only */
  xfId?: number;
  /** applyX attributes as written (undefined = absent) */
  applyNumberFormat?: boolean;
  applyFont?: boolean;
  applyFill?: boolean;
  applyBorder?: boolean;
  applyAlignment?: boolean;
  applyProtection?: boolean;
  alignment?: RawAlignment;
  protection?: { locked?: boolean; hidden?: boolean };
  quotePrefix?: boolean;
}

export interface RawCellStyle {
  xml: string;
  name: string;
  xfId: number;
  builtinId?: number;
}

export interface StyleSheet {
  numFmts: RawNumFmt[];
  /** custom numFmtId -> format code */
  numFmtMap: Map<number, string>;
  fonts: RawFont[];
  fills: RawFill[];
  borders: RawBorder[];
  cellStyleXfs: RawXf[];
  cellXfs: RawXf[];
  cellStyles: RawCellStyle[];
  /** differential formats (CF / table styles), raw XML only */
  dxfs: Array<{ xml: string }>;
  /** <colors><indexedColors> override ('RRGGBB'), when present */
  indexedColors?: string[];
  /** element-name prefix used by the part ('' or e.g. 'x:') — needed when appending records */
  prefix: string;
}

/** Excel's default font when a package has no styles part / no fonts. */
const FALLBACK_FONT: RawFont = { xml: '', size: 11, name: '游ゴシック', color: { theme: 1 }, family: 2, charset: 128, scheme: 'minor' };

/**
 * Parse styles.xml into raw records. `xml` null/empty → an empty sheet with one default font / fills / border / xf.
 * Record arrays are never empty: index 0 always exists (Excel semantics for a missing record = defaults).
 */
export function parseStyleSheet(xml: string | null | undefined): StyleSheet {
  const ss: StyleSheet = {
    numFmts: [],
    numFmtMap: new Map(),
    fonts: [],
    fills: [],
    borders: [],
    cellStyleXfs: [],
    cellXfs: [],
    cellStyles: [],
    dxfs: [],
    prefix: '',
  };
  if (xml) {
    let root: XEl | undefined;
    try {
      root = rootOf(parseXml(xml));
    } catch {
      root = undefined;
    }
    if (root) {
      const i = root.name.indexOf(':');
      ss.prefix = i < 0 ? '' : root.name.slice(0, i + 1);
      const raw = (el: XEl) => xml.slice(el.start, el.end);
      for (const nf of kids(kid(root, 'numFmts'), 'numFmt')) {
        const id = numAttr(nf.attrs.numFmtId);
        if (id === undefined) continue;
        const code = decodeOoxmlEscapes(nf.attrs.formatCode ?? '');
        ss.numFmts.push({ id, code, xml: raw(nf) });
        // an empty formatCode (excel-export writes <numFmt numFmtId="0" formatCode=""/>) defines nothing: the id keeps its built-in meaning
        if (code.trim() !== '') ss.numFmtMap.set(id, code);
      }
      for (const f of kids(kid(root, 'fonts'), 'font')) ss.fonts.push(parseFont(f, raw(f)));
      for (const f of kids(kid(root, 'fills'), 'fill')) ss.fills.push(parseFill(f, raw(f)));
      for (const b of kids(kid(root, 'borders'), 'border')) ss.borders.push(parseBorder(b, raw(b)));
      for (const x of kids(kid(root, 'cellStyleXfs'), 'xf')) ss.cellStyleXfs.push(parseXf(x, raw(x), false));
      for (const x of kids(kid(root, 'cellXfs'), 'xf')) ss.cellXfs.push(parseXf(x, raw(x), true));
      for (const cs of kids(kid(root, 'cellStyles'), 'cellStyle')) {
        ss.cellStyles.push({ xml: raw(cs), name: cs.attrs.name ?? '', xfId: numAttr(cs.attrs.xfId) ?? 0, builtinId: numAttr(cs.attrs.builtinId) });
      }
      for (const d of kids(kid(root, 'dxfs'), 'dxf')) ss.dxfs.push({ xml: raw(d) });
      const idx = kid(kid(root, 'colors'), 'indexedColors');
      if (idx) {
        const list = kids(idx, 'rgbColor').map((c) => {
          const v = (c.attrs.rgb ?? '').trim();
          return /^[0-9A-Fa-f]{8}$/.test(v) ? v.slice(2).toUpperCase() : /^[0-9A-Fa-f]{6}$/.test(v) ? v.toUpperCase() : '000000';
        });
        if (list.length) ss.indexedColors = list;
      }
    }
  }
  if (!ss.fonts.length) ss.fonts.push(FALLBACK_FONT);
  if (!ss.fills.length) ss.fills.push({ xml: '', patternType: 'none' }, { xml: '', patternType: 'gray125' });
  if (!ss.borders.length) ss.borders.push({ xml: '' });
  if (!ss.cellStyleXfs.length) ss.cellStyleXfs.push({ xml: '', numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 });
  if (!ss.cellXfs.length) ss.cellXfs.push({ xml: '', numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0 });
  return ss;
}

/** Boolean font property element (<b/>, <b val="0"/>): present without val = true. */
function flag(el: XEl | undefined): boolean | undefined {
  if (!el) return undefined;
  const v = el.attrs.val;
  return v === undefined ? true : boolAttr(v);
}

/** Parse a <font> (styles.xml) or <rPr> (rich text run) element. Run properties use <rFont> for the name. */
export function parseFont(f: XEl, xml = ''): RawFont {
  const out: RawFont = { xml };
  for (const c of f.children) {
    switch (c.local) {
      case 'b':
        out.bold = flag(c);
        break;
      case 'i':
        out.italic = flag(c);
        break;
      case 'strike':
        out.strike = flag(c);
        break;
      case 'u':
        out.underline = c.attrs.val ?? 'single';
        break;
      case 'sz':
        out.size = numAttr(c.attrs.val);
        break;
      case 'name':
      case 'rFont':
        out.name = c.attrs.val;
        break;
      case 'color':
        out.color = rawColorOf(c);
        break;
      case 'family':
        out.family = numAttr(c.attrs.val);
        break;
      case 'charset':
        out.charset = numAttr(c.attrs.val);
        break;
      case 'scheme':
        out.scheme = c.attrs.val;
        break;
      case 'vertAlign':
        out.vertAlign = c.attrs.val;
        break;
      default:
        break;
    }
  }
  return out;
}

function parseFill(f: XEl, xml: string): RawFill {
  const pf = kid(f, 'patternFill');
  if (pf) {
    return { xml, patternType: pf.attrs.patternType ?? (kid(pf, 'fgColor') ? 'solid' : 'none'), fg: rawColorOf(kid(pf, 'fgColor')), bg: rawColorOf(kid(pf, 'bgColor')) };
  }
  const gf = kid(f, 'gradientFill');
  if (gf) {
    const stops = kids(gf, 'stop').map((s) => ({ position: numAttr(s.attrs.position) ?? 0, color: rawColorOf(kid(s, 'color')) }));
    stops.sort((a, b) => a.position - b.position);
    return { xml, gradient: stops };
  }
  return { xml, patternType: 'none' };
}

function parseSide(el: XEl | undefined): RawBorderSide | undefined {
  if (!el) return undefined;
  const style = el.attrs.style;
  if (!style || style === 'none') return undefined;
  return { style, color: rawColorOf(kid(el, 'color')) };
}

function parseBorder(b: XEl, xml: string): RawBorder {
  return {
    xml,
    // Strict / older writers may use start/end instead of left/right
    left: parseSide(kid(b, 'left') ?? kid(b, 'start')),
    right: parseSide(kid(b, 'right') ?? kid(b, 'end')),
    top: parseSide(kid(b, 'top')),
    bottom: parseSide(kid(b, 'bottom')),
    diagonal: parseSide(kid(b, 'diagonal')),
    diagonalUp: b.attrs.diagonalUp !== undefined ? boolAttr(b.attrs.diagonalUp) : undefined,
    diagonalDown: b.attrs.diagonalDown !== undefined ? boolAttr(b.attrs.diagonalDown) : undefined,
  };
}

function optBool(v: string | undefined): boolean | undefined {
  return v === undefined ? undefined : boolAttr(v);
}

function parseXf(x: XEl, xml: string, isCellXf: boolean): RawXf {
  const a = x.attrs;
  const xf: RawXf = {
    xml,
    numFmtId: numAttr(a.numFmtId) ?? 0,
    fontId: numAttr(a.fontId) ?? 0,
    fillId: numAttr(a.fillId) ?? 0,
    borderId: numAttr(a.borderId) ?? 0,
  };
  if (isCellXf) xf.xfId = numAttr(a.xfId) ?? 0;
  xf.applyNumberFormat = optBool(a.applyNumberFormat);
  xf.applyFont = optBool(a.applyFont);
  xf.applyFill = optBool(a.applyFill);
  xf.applyBorder = optBool(a.applyBorder);
  xf.applyAlignment = optBool(a.applyAlignment);
  xf.applyProtection = optBool(a.applyProtection);
  if (a.quotePrefix !== undefined) xf.quotePrefix = boolAttr(a.quotePrefix);
  const al = kid(x, 'alignment');
  if (al) {
    xf.alignment = {
      horizontal: al.attrs.horizontal,
      vertical: al.attrs.vertical,
      wrapText: optBool(al.attrs.wrapText),
      indent: numAttr(al.attrs.indent),
      shrinkToFit: optBool(al.attrs.shrinkToFit),
      textRotation: numAttr(al.attrs.textRotation),
      readingOrder: numAttr(al.attrs.readingOrder),
    };
  }
  const pr = kid(x, 'protection');
  if (pr) xf.protection = { locked: optBool(pr.attrs.locked), hidden: optBool(pr.attrs.hidden) };
  return xf;
}

// ---------------------------------------------------------------------------------------------
// Effective xf and display styles
// ---------------------------------------------------------------------------------------------

/** Format code of a numFmtId: custom codes first, then the ja-JP built-in table; unknown ids → 'General'. */
export function numFmtCode(ss: StyleSheet, id: number): string {
  return ss.numFmtMap.get(id) ?? JA_BUILTIN_NUMFMT[id] ?? 'General';
}

/**
 * The cellXf as Excel renders it. Excel uses the cellXf's OWN numFmtId / fontId / fillId / borderId / alignment
 * whatever the applyX attributes say (applyX only records whether the cell overrides its cell style when the
 * style is re-applied). Verified via COM with apply-probe.xlsx: applyFont="0" etc. with a parent style that has
 * a bold font / fill / border / date format still renders the cellXf's own records.
 */
export function effectiveXf(ss: StyleSheet, index: number): RawXf {
  return ss.cellXfs[index] ?? ss.cellXfs[0];
}

const BORDER_STYLES: ReadonlySet<string> = new Set([
  'thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot',
]);

const H_ALIGN: Record<string, XStyle['hAlign']> = {
  general: 'general', left: 'left', center: 'center', right: 'right', fill: 'fill', justify: 'justify', centerContinuous: 'centerContinuous', distributed: 'distributed',
};

function vAlignOf(v: string | undefined): XStyle['vAlign'] {
  switch (v) {
    case 'top':
      return 'top';
    case 'center':
    case 'distributed':
      return 'middle';
    case 'justify':
      return 'top';
    case 'bottom':
      return 'bottom';
    default:
      return undefined; // Excel default = bottom
  }
}

export function paletteOf(ss: StyleSheet, theme: readonly string[]): Palette {
  return { theme, indexed: ss.indexedColors ?? DEFAULT_INDEXED };
}

/** Font → the font part of an XStyle (used for cells and rich-text runs). */
export function fontStyle(f: RawFont | undefined, pal: Palette = DEFAULT_PALETTE): XStyle {
  const s: XStyle = {};
  if (!f) return s;
  if (f.bold) s.bold = true;
  if (f.italic) s.italic = true;
  if (f.underline && f.underline !== 'none') s.underline = true;
  if (f.strike) s.strike = true;
  if (f.size !== undefined) s.size = f.size;
  if (f.name) s.font = f.name;
  const color = resolveColor(f.color, pal);
  if (color) s.color = color;
  return s;
}

function fillColor(fill: RawFill | undefined, pal: Palette): string | undefined {
  if (!fill) return undefined;
  if (fill.gradient) return fill.gradient.length ? resolveColor(fill.gradient[0].color, pal) : undefined;
  const pt = fill.patternType;
  if (!pt || pt === 'none') return undefined;
  if (pt === 'solid') return resolveColor(fill.fg, pal) ?? (fill.fg?.auto ? '#000000' : undefined);
  // patterned fills: approximate with the background colour (what Excel reports as Interior.Color)
  return resolveColor(fill.bg, pal);
}

function borderSide(side: RawBorderSide | undefined, pal: Palette): XBorderSide | undefined {
  if (!side || !side.style || !BORDER_STYLES.has(side.style)) return undefined;
  const out: XBorderSide = { style: side.style as XBorderStyle };
  const c = resolveColor(side.color, pal);
  if (c) out.color = c;
  return out;
}

/** True when a format code shows a date and/or time ('General' / text formats are not dates). */
export function isDateCode(code: string): boolean {
  if (!code || code === 'General' || code === '@') return false;
  let r = DATE_CACHE.get(code);
  if (r !== undefined) return r;
  // strip literals ("..", \x), and [..] tags except elapsed-time [h] [m] [s]
  const bare = code
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/\[(?![hms]+\])[^\]]*\]/gi, '')
    .toLowerCase();
  if (/general/.test(bare) && !/[ymdhs]/.test(bare.replace(/general/g, ''))) r = false;
  else {
    let ssf = false;
    try {
      ssf = SSF.is_date(code);
    } catch {
      ssf = false;
    }
    // 和暦 (g/e) or ymdhs tokens outside literals; 'e' followed by +/- is scientific notation
    r = ssf || /[ymdhs]|g{1,3}|e(?![+-])/.test(bare.replace(/e[+-]/g, ''));
    // pure number formats such as 0.00E+00 or #,##0 never contain those letters after stripping
  }
  if (DATE_CACHE.size > 5000) DATE_CACHE.clear();
  DATE_CACHE.set(code, r);
  return r;
}
const DATE_CACHE = new Map<string, boolean>();

/**
 * What a date/time format shows: 'date' = a calendar date (maybe with time), 'time' = time of day only,
 * 'elapsed' = [h]/[m]/[s] durations, 'none' = not a date format. Used for 1904 → 1900 conversion: only values
 * that denote a calendar date are shifted (a 12:30 time or a 36:00 duration must keep its value).
 */
export function dateModeOf(code: string): 'none' | 'date' | 'time' | 'elapsed' {
  if (!isDateCode(code)) return 'none';
  const noLit = code.replace(/"[^"]*"/g, '').replace(/\\./g, '');
  if (/\[(h+|m+|s+)\]/i.test(noLit)) return 'elapsed';
  const bare = noLit
    .replace(/\[[^\]]*\]/g, '')
    .toLowerCase()
    .replace(/am\/pm|a\/p/g, '');
  if (/[ydg]|e(?![+-])|aaa/.test(bare)) return 'date';
  // 'm' is minutes next to h / s, else months
  const hasTime = /[hs]/.test(bare);
  if (/m/.test(bare) && !hasTime) return 'date';
  return 'time';
}

export interface DisplayStyles {
  styles: XStyle[];
  /** per cellXfs index: the number format is a date/time format */
  isDate: boolean[];
  /** per cellXfs index: date format with a calendar-date part ('date'), time-of-day only ('time') */
  dateMode: Array<'none' | 'date' | 'time' | 'elapsed'>;
}

/** Resolve every cellXf into a display XStyle (index = xf index). */
export function buildDisplayStyles(ss: StyleSheet, pal: Palette = DEFAULT_PALETTE): DisplayStyles {
  const fontCache = new Map<number, XStyle>();
  const styles: XStyle[] = [];
  const isDate: boolean[] = [];
  const dateMode: DisplayStyles['dateMode'] = [];
  for (let i = 0; i < ss.cellXfs.length; i++) {
    const xf = effectiveXf(ss, i);
    let fs = fontCache.get(xf.fontId);
    if (!fs) {
      fs = fontStyle(ss.fonts[xf.fontId] ?? ss.fonts[0], pal);
      fontCache.set(xf.fontId, fs);
    }
    const s: XStyle = { ...fs };
    const fill = fillColor(ss.fills[xf.fillId], pal);
    if (fill) s.fill = fill;
    const b = ss.borders[xf.borderId];
    if (b) {
      const border: NonNullable<XStyle['border']> = {};
      const top = borderSide(b.top, pal);
      const right = borderSide(b.right, pal);
      const bottom = borderSide(b.bottom, pal);
      const left = borderSide(b.left, pal);
      if (top) border.top = top;
      if (right) border.right = right;
      if (bottom) border.bottom = bottom;
      if (left) border.left = left;
      if (top || right || bottom || left) s.border = border;
    }
    const al = xf.alignment;
    if (al) {
      const h = al.horizontal ? H_ALIGN[al.horizontal] : undefined;
      if (h && h !== 'general') s.hAlign = h;
      const v = vAlignOf(al.vertical);
      if (v) s.vAlign = v;
      if (al.wrapText) s.wrap = true;
      if (al.indent) s.indent = al.indent;
    }
    const code = numFmtCode(ss, xf.numFmtId);
    if (code !== 'General') s.numFmt = code;
    styles.push(s);
    const mode = dateModeOf(code);
    isDate.push(mode !== 'none');
    dateMode.push(mode);
  }
  return { styles, isDate, dateMode };
}
