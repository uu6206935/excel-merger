/**
 * Colour resolution for display (owner: reader R): theme palette from xl/theme/theme1.xml, the legacy indexed
 * palette (optionally overridden by <colors><indexedColors> in styles.xml) and Excel's tint algorithm.
 *
 * Theme index order in SpreadsheetML (NOT the clrScheme element order!):
 *   0 = lt1 (bg1), 1 = dk1 (tx1), 2 = lt2 (bg2), 3 = dk2 (tx2), 4..9 = accent1..6, 10 = hlink, 11 = folHlink
 * (clrScheme lists dk1, lt1, dk2, lt2 — the first two pairs are swapped). Verified against Excel via COM,
 * see tests/xlsx/read-styles.test.ts and tests/fixtures/xlsx/excel-colors.json.
 */
import { kid, parseXml, rootOf, type XEl } from './xml';

/** Office 2013-2022 default theme (fallback when a package has no theme part), in theme-index order. */
export const DEFAULT_THEME: readonly string[] = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];

/** Default legacy palette (indexed colours 0..63). 64 = system foreground (black), 65 = system background (white). */
export const DEFAULT_INDEXED: readonly string[] = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

const SCHEME_ORDER = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

const PRESET: Record<string, string> = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080',
};

/** Parse theme1.xml into the 12 theme colours (theme-index order, 'RRGGBB'). Missing entries fall back to DEFAULT_THEME. */
export function parseThemeColors(xml: string | null | undefined): string[] {
  const out = [...DEFAULT_THEME];
  if (!xml) return out;
  const root = rootOf(parseXml(xml));
  const scheme = findDeep(root, 'clrScheme');
  if (!scheme) return out;
  SCHEME_ORDER.forEach((name, idx) => {
    const el = kid(scheme, name);
    const rgb = el ? drawingColor(el) : undefined;
    if (rgb) out[idx] = rgb;
  });
  return out;
}

function findDeep(el: XEl | undefined, local: string): XEl | undefined {
  if (!el) return undefined;
  if (el.local === local) return el;
  for (const c of el.children) {
    const f = findDeep(c, local);
    if (f) return f;
  }
  return undefined;
}

/** DrawingML colour choice (<a:srgbClr val>, <a:sysClr lastClr>, <a:prstClr>) → 'RRGGBB'. */
function drawingColor(holder: XEl): string | undefined {
  for (const c of holder.children) {
    if (c.local === 'srgbClr' && c.attrs.val && /^[0-9A-Fa-f]{6}$/.test(c.attrs.val)) return c.attrs.val.toUpperCase();
    if (c.local === 'sysClr') {
      const last = c.attrs.lastClr;
      if (last && /^[0-9A-Fa-f]{6}$/.test(last)) return last.toUpperCase();
      if (c.attrs.val === 'windowText') return '000000';
      if (c.attrs.val === 'window') return 'FFFFFF';
    }
    if (c.local === 'prstClr' && c.attrs.val) return PRESET[c.attrs.val.toLowerCase()];
  }
  return undefined;
}

/** Raw colour reference as stored in styles.xml (<color>, <fgColor>, <bgColor>, rich-text run colours). */
export interface RawColor {
  /** 'AARRGGBB' or 'RRGGBB' as written */
  rgb?: string;
  theme?: number;
  tint?: number;
  indexed?: number;
  auto?: boolean;
}

export interface Palette {
  theme: readonly string[];
  indexed: readonly string[];
}

export const DEFAULT_PALETTE: Palette = { theme: DEFAULT_THEME, indexed: DEFAULT_INDEXED };

/**
 * Resolve a raw colour to '#RRGGBB'. Returns undefined for automatic / unresolvable colours (the caller picks the
 * context default: black text/borders, no fill). `systemIndexed`: what indexed 64/65 mean in this context.
 */
export function resolveColor(c: RawColor | undefined, pal: Palette = DEFAULT_PALETTE): string | undefined {
  if (!c) return undefined;
  let rgb: string | undefined;
  if (c.rgb !== undefined) {
    const v = c.rgb.trim();
    if (/^[0-9A-Fa-f]{8}$/.test(v)) rgb = v.slice(2);
    else if (/^[0-9A-Fa-f]{6}$/.test(v)) rgb = v;
  } else if (c.theme !== undefined) {
    rgb = pal.theme[c.theme];
  } else if (c.indexed !== undefined) {
    if (c.indexed === 64) rgb = '000000'; // system foreground
    else if (c.indexed === 65) rgb = 'FFFFFF'; // system background
    else rgb = pal.indexed[c.indexed] ?? DEFAULT_INDEXED[c.indexed];
  } else if (c.auto) {
    return undefined;
  }
  if (!rgb) return undefined;
  rgb = rgb.toUpperCase();
  if (c.tint) rgb = applyTint(rgb, c.tint);
  return '#' + rgb;
}

/** Parse a colour element's attributes into a RawColor. */
export function rawColorOf(el: XEl | undefined): RawColor | undefined {
  if (!el) return undefined;
  const a = el.attrs;
  const c: RawColor = {};
  if (a.rgb !== undefined) c.rgb = a.rgb;
  if (a.theme !== undefined && Number.isFinite(Number(a.theme))) c.theme = Number(a.theme);
  if (a.indexed !== undefined && Number.isFinite(Number(a.indexed))) c.indexed = Number(a.indexed);
  if (a.tint !== undefined && Number.isFinite(Number(a.tint)) && Number(a.tint) !== 0) c.tint = Number(a.tint);
  if (a.auto === '1' || a.auto === 'true') c.auto = true;
  return c;
}

// ---------------------------------------------------------------------------------------------
// Tint. ECMA-376 18.8.19 describes: convert to HLS, scale the luminance, convert back
//   tint < 0: L' = L * (1 + tint)
//   tint > 0: L' = L * (1 - tint) + (HLSMAX - HLSMAX * (1 - tint))
// Excel does this with the classic Windows integer HLS conversion (HLSMAX = 240, RGBMAX = 255, KB 29240) and
// truncates each product term to an integer. Floating-point or HLSMAX = 255 versions are off by one in ~80% of
// cases. This integer version reproduces all 161 tinted samples Excel reported via COM (12 theme colours x 13
// tints, fill and font; tests/fixtures/xlsx/excel-colors.json). A zero tint leaves the colour untouched.
// ---------------------------------------------------------------------------------------------

const HLSMAX = 240;
const RGBMAX = 255;

function idiv(a: number, b: number): number {
  return Math.trunc(a / b);
}

/** Apply an Excel tint to 'RRGGBB'; returns 'RRGGBB' (upper case). */
export function applyTint(rgb: string, tint: number): string {
  if (!tint || !Number.isFinite(tint)) return rgb.toUpperCase();
  const t = Math.max(-1, Math.min(1, tint));
  const R = parseInt(rgb.slice(0, 2), 16);
  const G = parseInt(rgb.slice(2, 4), 16);
  const B = parseInt(rgb.slice(4, 6), 16);
  if (!Number.isFinite(R + G + B)) return rgb.toUpperCase();
  // RGB -> HLS (integer)
  const cMax = Math.max(R, G, B);
  const cMin = Math.min(R, G, B);
  let L = idiv((cMax + cMin) * HLSMAX + RGBMAX, 2 * RGBMAX);
  let H = 0;
  let S = 0;
  if (cMax !== cMin) {
    const sum = cMax + cMin;
    const d = cMax - cMin;
    S = L <= HLSMAX / 2 ? idiv(d * HLSMAX + idiv(sum, 2), sum) : idiv(d * HLSMAX + idiv(2 * RGBMAX - sum, 2), 2 * RGBMAX - sum);
    const Rd = idiv((cMax - R) * (HLSMAX / 6) + idiv(d, 2), d);
    const Gd = idiv((cMax - G) * (HLSMAX / 6) + idiv(d, 2), d);
    const Bd = idiv((cMax - B) * (HLSMAX / 6) + idiv(d, 2), d);
    if (R === cMax) H = Bd - Gd;
    else if (G === cMax) H = HLSMAX / 3 + Rd - Bd;
    else H = (2 * HLSMAX) / 3 + Gd - Rd;
    if (H < 0) H += HLSMAX;
    if (H > HLSMAX) H -= HLSMAX;
  }
  // luminance
  if (t < 0) L = Math.trunc(L * (1 + t));
  else L = Math.trunc(L * (1 - t)) + (HLSMAX - Math.trunc(HLSMAX * (1 - t)));
  L = Math.max(0, Math.min(HLSMAX, L));
  // HLS -> RGB (integer)
  let r: number;
  let g: number;
  let b: number;
  if (S === 0) {
    r = g = b = idiv(L * RGBMAX + HLSMAX / 2, HLSMAX);
  } else {
    const m2 = L <= HLSMAX / 2 ? idiv(L * (HLSMAX + S) + HLSMAX / 2, HLSMAX) : L + S - idiv(L * S + HLSMAX / 2, HLSMAX);
    const m1 = 2 * L - m2;
    r = idiv(hueToRgb(m1, m2, H + HLSMAX / 3) * RGBMAX + HLSMAX / 2, HLSMAX);
    g = idiv(hueToRgb(m1, m2, H) * RGBMAX + HLSMAX / 2, HLSMAX);
    b = idiv(hueToRgb(m1, m2, H - HLSMAX / 3) * RGBMAX + HLSMAX / 2, HLSMAX);
  }
  return hex2(r) + hex2(g) + hex2(b);
}

function hueToRgb(n1: number, n2: number, hue: number): number {
  if (hue < 0) hue += HLSMAX;
  if (hue > HLSMAX) hue -= HLSMAX;
  if (hue < HLSMAX / 6) return n1 + idiv((n2 - n1) * hue + HLSMAX / 12, HLSMAX / 6);
  if (hue < HLSMAX / 2) return n2;
  if (hue < (HLSMAX * 2) / 3) return n1 + idiv((n2 - n1) * ((HLSMAX * 2) / 3 - hue) + HLSMAX / 12, HLSMAX / 6);
  return n1;
}

function hex2(v: number): string {
  const x = Math.max(0, Math.min(255, Math.round(v)));
  return x.toString(16).toUpperCase().padStart(2, '0');
}
