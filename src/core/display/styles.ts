/**
 * Display styles (owner: grid/display G). Pure: no React, no DOM.
 *
 *  - styleToCss(XStyle) -> CellCss (React-style camelCase declarations; the grid casts them to CSSProperties),
 *    cached per style id by makeStyleCache(doc.styles).
 *  - Excel colour helpers: resolveExcelColor ({argb} | {theme,tint} | {indexed} -> '#RRGGBB'), applyTint
 *    (Excel's HLS tint), OFFICE_2013_THEME / OFFICE_2023_THEME, INDEXED_COLORS.
 *  - Unit helpers: excelWidthToPx / pxToExcelWidth (ECMA-376 column width formulas), defaultColumnPx,
 *    ptToPx / pxToPt (row heights), MDW_BY_FONT (max digit width of common workbook default fonts).
 */
import type { XBorderSide, XStyle } from '../types';

/** A style declaration object with React camelCase property names (e.g. { backgroundColor: '#fff' }). */
export type CssDecl = Record<string, string | number | undefined>;

export interface CellCss {
  /** applied to the cell box */
  box: CssDecl;
  /** applied to an absolutely-positioned child that draws Excel borders (undefined = none) */
  border?: CssDecl;
  /** explicit horizontal alignment; undefined = Excel "General" (depends on the value type) */
  textAlign?: 'left' | 'center' | 'right' | 'justify';
  /** vertical alignment (Excel default: bottom) */
  vAlign: 'top' | 'middle' | 'bottom';
  /** Excel number format code used for display */
  numFmt?: string;
  wrap: boolean;
  /** CSS/canvas font shorthand of the cell text (used to measure text for autofit) */
  font: string;
}

/** Cell font stack when the workbook font is the Japanese default (游ゴシック) or unknown. */
export const DEFAULT_FONT_STACK = "'Yu Gothic UI', 'Meiryo UI', sans-serif";
export const DEFAULT_FONT_SIZE_PT = 11;
export const GRIDLINE_COLOR = '#e1e1e1';

const BORDER_CSS: Record<XBorderSide['style'], string> = {
  thin: '1px solid',
  hair: '1px dotted',
  dotted: '1px dotted',
  dashed: '1px dashed',
  dashDot: '1px dashed',
  dashDotDot: '1px dotted',
  medium: '2px solid',
  mediumDashed: '2px dashed',
  mediumDashDot: '2px dashed',
  mediumDashDotDot: '2px dotted',
  slantDashDot: '2px dashed',
  thick: '3px solid',
  double: '3px double',
};

function side(s?: XBorderSide): string | undefined {
  if (!s || !s.style) return undefined;
  return `${BORDER_CSS[s.style] ?? '1px solid'} ${s.color ?? '#000000'}`;
}

/**
 * Excel font name -> CSS font-family list. The Japanese defaults (游ゴシック / Yu Gothic, incl. Medium/Regular)
 * render with the screen-tuned 'Yu Gothic UI' (Chromium draws 'Yu Gothic' Regular too thin on Windows).
 */
export function fontFamilyFor(name: string | undefined): string {
  if (!name) return DEFAULT_FONT_STACK;
  const n = name.trim();
  if (!n) return DEFAULT_FONT_STACK;
  if (/^(游ゴシック|Yu ?Gothic)( (Medium|Regular|UI))?$/i.test(n)) return DEFAULT_FONT_STACK;
  const quoted = `'${n.replace(/'/g, '')}'`;
  return `${quoted}, ${DEFAULT_FONT_STACK}`;
}

function fontShorthand(s: XStyle | undefined): string {
  const italic = s?.italic ? 'italic ' : '';
  const weight = s?.bold ? '700 ' : '400 ';
  const size = s?.size && s.size > 0 ? s.size : DEFAULT_FONT_SIZE_PT;
  return `${italic}${weight}${size}pt ${fontFamilyFor(s?.font)}`;
}

const DEFAULT_CSS: CellCss = Object.freeze({ box: {}, vAlign: 'bottom', wrap: false, font: fontShorthand(undefined) }) as CellCss;

export function styleToCss(s: XStyle | undefined | null): CellCss {
  if (!s) return DEFAULT_CSS;
  const box: CssDecl = {};
  if (s.fill) box.backgroundColor = s.fill;
  if (s.color) box.color = s.color;
  if (s.bold) box.fontWeight = 700;
  if (s.italic) box.fontStyle = 'italic';
  if (s.underline || s.strike) box.textDecoration = [s.underline ? 'underline' : '', s.strike ? 'line-through' : ''].join(' ').trim();
  if (s.size && s.size > 0 && s.size !== DEFAULT_FONT_SIZE_PT) box.fontSize = `${s.size}pt`;
  const fam = fontFamilyFor(s.font);
  if (fam !== DEFAULT_FONT_STACK) box.fontFamily = fam;
  const wrap = !!s.wrap;
  if (wrap) {
    box.whiteSpace = 'pre-wrap';
    box.wordBreak = 'break-all'; // Japanese text has no spaces; Excel breaks anywhere
    box.lineHeight = 1.25;
  }
  const vAlign: CellCss['vAlign'] = s.vAlign === 'top' ? 'top' : s.vAlign === 'middle' ? 'middle' : 'bottom';
  if (vAlign === 'top') box.alignItems = 'flex-start';
  else if (vAlign === 'middle') box.alignItems = 'center';
  // bottom is the Excel default -> .g-cell uses align-items:flex-end
  let textAlign: CellCss['textAlign'];
  switch (s.hAlign) {
    case 'left':
    case 'fill':
      textAlign = 'left';
      break;
    case 'center':
    case 'centerContinuous':
    case 'distributed':
      textAlign = 'center';
      break;
    case 'right':
      textAlign = 'right';
      break;
    case 'justify':
      textAlign = 'justify';
      break;
    default:
      textAlign = undefined;
  }
  if (s.indent && s.indent > 0) {
    const pad = 3 + Math.min(s.indent, 15) * 9;
    if (textAlign === 'right') box.paddingRight = pad;
    else box.paddingLeft = pad;
  }
  let border: CssDecl | undefined;
  const b = s.border;
  if (b && (b.top || b.right || b.bottom || b.left)) {
    border = {};
    const t = side(b.top);
    const r = side(b.right);
    const bo = side(b.bottom);
    const l = side(b.left);
    if (t) border.borderTop = t;
    if (r) border.borderRight = r;
    if (bo) border.borderBottom = bo;
    if (l) border.borderLeft = l;
    if (!t && !r && !bo && !l) border = undefined;
  }
  const out: CellCss = { box, vAlign, wrap, font: fontShorthand(s) };
  if (border) out.border = border;
  if (textAlign) out.textAlign = textAlign;
  if (s.numFmt !== undefined && s.numFmt !== '') out.numFmt = s.numFmt;
  return out;
}

/** Style lookup cached per xf index. Unknown / out-of-range ids get the default style. */
export function makeStyleCache(styles: readonly (XStyle | undefined)[]): (id: number | undefined) => CellCss {
  const cache = new Map<number, CellCss>();
  return (id: number | undefined): CellCss => {
    const key = id ?? 0;
    let c = cache.get(key);
    if (!c) {
      c = styleToCss(key >= 0 && key < styles.length ? styles[key] : undefined);
      cache.set(key, c);
    }
    return c;
  };
}

// ---------------------------------------------------------------------------------------------
// Excel colours -> #RRGGBB (for readers that get {theme, tint} / {indexed} / ARGB colour references)

/** Office 2013-2022 default theme, in Excel's theme-index order (0 = lt1/bg1, 1 = dk1/tx1, 2 = lt2, 3 = dk2, 4-9 accents, 10 hlink, 11 folHlink). */
export const OFFICE_2013_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
/** Office 2023+ default theme (new files from Microsoft 365 since late 2023). Parse xl/theme/theme1.xml for the real one. */
export const OFFICE_2023_THEME = ['FFFFFF', '000000', 'E8E8E8', '0E2841', '156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E', '467886', '96607D'];

/** Legacy indexed palette (indexes 0-63; 64 = system foreground, 65 = system background). */
export const INDEXED_COLORS = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

export interface ExcelColor {
  argb?: string;
  rgb?: string;
  theme?: number;
  tint?: number;
  indexed?: number;
}

export function resolveExcelColor(c: ExcelColor | undefined | null, theme: readonly string[] = OFFICE_2013_THEME): string | undefined {
  if (!c) return undefined;
  let rgb: string | undefined;
  const raw = c.argb ?? c.rgb;
  if (raw) rgb = raw.length === 8 ? raw.slice(2) : raw;
  else if (c.theme !== undefined) rgb = theme[c.theme];
  else if (c.indexed !== undefined) rgb = c.indexed === 64 ? '000000' : c.indexed === 65 ? 'FFFFFF' : INDEXED_COLORS[c.indexed];
  if (!rgb || !/^[0-9a-f]{6}$/i.test(rgb)) return undefined;
  if (c.tint) rgb = applyTint(rgb, c.tint);
  return '#' + rgb.toUpperCase();
}

/** Excel's tint algorithm (HLS lightness, ECMA-376 18.8.19 / MS-OI29500). */
export function applyTint(rgb: string, tint: number): string {
  const r = parseInt(rgb.slice(0, 2), 16) / 255;
  const g = parseInt(rgb.slice(2, 4), 16) / 255;
  const b = parseInt(rgb.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  let s = 0;
  let l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const hue2rgb = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let R: number;
  let G: number;
  let B: number;
  if (s === 0) R = G = B = l;
  else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    R = hue2rgb(p, q, h + 1 / 3);
    G = hue2rgb(p, q, h);
    B = hue2rgb(p, q, h - 1 / 3);
  }
  const hex = (x: number) => Math.round(Math.max(0, Math.min(1, x)) * 255).toString(16).padStart(2, '0');
  return (hex(R) + hex(G) + hex(B)).toUpperCase();
}

// ---------------------------------------------------------------------------------------------
// Units

/**
 * Maximum digit width (px at 96 dpi, 100 % zoom) of common workbook default fonts at 11pt.
 * 游ゴシック / ＭＳ Ｐゴシック 11 -> 8 px (default column 8.38 = 72 px), Calibri 11 -> 7 px (8.43 = 64 px).
 */
export const MDW_BY_FONT: Record<string, number> = {
  '游ゴシック': 8,
  'Yu Gothic': 8,
  'ＭＳ Ｐゴシック': 8,
  'MS PGothic': 8,
  'ＭＳ ゴシック': 8,
  'MS Gothic': 8,
  'メイリオ': 8,
  Meiryo: 8,
  Calibri: 7,
  Arial: 7,
  'Aptos Narrow': 7,
};

/** Max digit width for a workbook default font (unknown fonts: 8 at 11pt, scaled by size). */
export function mdwFor(fontName: string | undefined, sizePt = 11): number {
  const base = (fontName && MDW_BY_FONT[fontName.trim()]) || 8;
  return Math.max(1, Math.round((base * sizePt) / 11));
}

/** Column width as stored in the file (<col width>, character units incl. padding) -> pixels. ECMA-376 18.3.1.13 */
export function excelWidthToPx(width: number, mdw = 8): number {
  if (!(width > 0)) return 0;
  return Math.trunc(((256 * width + Math.trunc(128 / mdw)) / 256) * mdw);
}

/** Pixels -> column width as stored in the file (inverse of excelWidthToPx; via the character count Excel shows). */
export function pxToExcelWidth(px: number, mdw = 8): number {
  if (!(px > 0)) return 0;
  const chars = Math.trunc(((px - 5) / mdw) * 100 + 0.5) / 100;
  if (chars <= 0) return Math.trunc((px / mdw) * 256) / 256;
  return Math.trunc(((chars * mdw + 5) / mdw) * 256) / 256;
}

/** Column width in the character count Excel's UI shows (e.g. 8.38) -> pixels. */
export function charsToPx(chars: number, mdw = 8): number {
  if (!(chars > 0)) return 0;
  return Math.trunc(chars * mdw + 5 + 0.5);
}

/** Default column width in pixels (sheetFormatPr baseColWidth, default 8 chars): padding rounded up to 8 px. */
export function defaultColumnPx(mdw = 8, baseColWidth = 8): number {
  return Math.ceil((baseColWidth * mdw + 5) / 8) * 8;
}

/** Row height points -> CSS pixels (96 dpi). */
export function ptToPx(pt: number): number {
  return Math.round((pt * 96) / 72);
}

/** CSS pixels -> points (96 dpi). */
export function pxToPt(px: number): number {
  return Math.round(((px * 72) / 96) * 100) / 100;
}
