/**
 * Facts about a whole workbook package that premerge (imported columns) and the sheet copier (addsheets.ts) both need to decide how the
 * cells of one package must be converted to be shown identically in another: theme colours / font scheme, the legacy indexed palette
 * and the date system.
 */
import type { Engine } from './engine';
import { isRel } from './pkg';
import { themeColors, type Styles } from './styles';
import { getAttr, unesc } from './xml';

/** the theme part (xl/theme/theme1.xml) of a package, null when it has none */
export async function themeXml(eng: Engine): Promise<string | null> {
  const rel = (await eng.pkg.relTargets(eng.wbPath)).find((r) => isRel(r.type, 'theme'));
  return rel ? eng.pkg.read(rel.path) : null;
}

/** the <a:fontScheme> of a theme without its name (themes are "the same" when colours and font scheme agree) */
export function fontScheme(theme: string | null): string {
  return /<a:fontScheme\b[\s\S]*?<\/a:fontScheme>/.exec(theme ?? '')?.[0].replace(/\sname="[^"]*"/, '') ?? '';
}

/** the workbook stores dates in the 1904 system */
export function is1904(eng: Engine): boolean {
  return /<workbookPr\b[^>]*\sdate1904="(1|true)"/.test(eng.wbXml);
}

/** the two styles.xml resolve indexed colours (0..63) identically */
export function samePalette(a: Styles, b: Styles): boolean {
  const pa = a.effectivePalette(), pb = b.effectivePalette();
  return pa.every((c, i) => c === pb[i]);
}

/** theme colours AND font scheme agree: theme colour references of one package keep their meaning in the other */
export function sameTheme(a: string | null, b: string | null): boolean {
  return JSON.stringify(themeColors(a)) === JSON.stringify(themeColors(b)) && fontScheme(a) === fontScheme(b);
}

// ---------------------------------------------------------------------------------------------------------------------
// Transitional <-> Strict OOXML
// ---------------------------------------------------------------------------------------------------------------------

const NS_PAIRS: ReadonlyArray<readonly [strict: string, transitional: string]> = [
  ['http://purl.oclc.org/ooxml/spreadsheetml/main', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'],
  ['http://purl.oclc.org/ooxml/officeDocument/relationships', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'],
];

/** the workbook part is written in the Strict OOXML namespaces (purl.oclc.org) */
export function isStrict(wbXml: string): boolean {
  return wbXml.includes(NS_PAIRS[0][0]);
}

/** the namespace declarations (and relationship type URLs) of a part written for the other flavour of OOXML */
export function convertNamespaces(xml: string, toStrict: boolean): string {
  let x = xml;
  for (const [strict, trans] of NS_PAIRS) {
    const [from, to] = toStrict ? [trans, strict] : [strict, trans];
    if (x.includes(from)) x = x.split(from).join(to);
  }
  return x;
}

// ---------------------------------------------------------------------------------------------------------------------
// The default font of a workbook and the pixel width of its columns
// ---------------------------------------------------------------------------------------------------------------------

/**
 * The font a theme names for `which` ("minor" = body, "major" = headings) as Excel shows it in a Japanese-language Excel: the East Asian font
 * when the theme sets one, else the Japanese script font of the theme's font list (openpyxl / XlsxWriter / Excel 2007 stock themes list
 * "ＭＳ Ｐゴシック" there), else the Latin font.
 */
export function themeFont(theme: string | null, which: 'major' | 'minor'): string | undefined {
  const block = (which === 'minor' ? /<a:minorFont>([\s\S]*?)<\/a:minorFont>/ : /<a:majorFont>([\s\S]*?)<\/a:majorFont>/).exec(theme ?? '')?.[1];
  if (!block) return undefined;
  const face = (re: RegExp) => { const v = re.exec(block)?.[1]; return v ? unesc(v) : undefined; };
  return face(/<a:ea\b[^>]*\stypeface="([^"]+)"/) ?? face(/<a:font\b[^>]*\sscript="Jpan"[^>]*\stypeface="([^"]+)"/) ?? face(/<a:font\b[^>]*\stypeface="([^"]+)"[^>]*\sscript="Jpan"/) ?? face(/<a:latin\b[^>]*\stypeface="([^"]+)"/);
}

/** name and size of the font of cell format 0 (the workbook's default font); a theme (scheme) font is resolved through the theme */
export function normalFontOf(styles: Styles, theme: string | null): { name?: string; size?: number } {
  const fontId = Number(getAttr(/^<[^>]*>/.exec(styles.xf(0))?.[0] ?? '', 'fontId') ?? 0);
  const f = styles.get('fonts', fontId) ?? styles.get('fonts', 0) ?? '';
  const name = /<name\b[^>]*\sval="([^"]*)"/.exec(f)?.[1];
  const size = Number(/<sz\b[^>]*\sval="([^"]*)"/.exec(f)?.[1]);
  const scheme = /<scheme\b[^>]*\sval="(major|minor)"/.exec(f)?.[1] as 'major' | 'minor' | undefined;
  return { name: (scheme && themeFont(theme, scheme)) || (name ? unesc(name) : undefined), size: Number.isFinite(size) && size > 0 ? size : undefined };
}

/**
 * Pixels per digit (the "maximum digit width" Excel measures column widths in) of the default fonts whose value is known from Excel itself
 * (default column width 8.43 = 64 px with Calibri 11 / Arial 10, 8.38 = 72 px with the Japanese fonts at 11 pt); undefined for any other font:
 * widths are then left as they are.
 */
export function knownMdw(name: string | undefined, size: number | undefined): number | undefined {
  if (!name || size === undefined) return undefined;
  const n = name.normalize('NFKC').trim().toLowerCase().replace(/\s+/g, ' ');
  const s = Math.round(size * 2) / 2;
  if (n === 'calibri' && s === 11) return 7;
  if (n === 'arial' && s === 10) return 7;
  if (n === 'arial' && s === 11) return 8;
  if (s === 11 && ['游ゴシック', 'yu gothic', 'ms pgothic', 'ms gothic', 'メイリオ', 'meiryo', 'ｍｓ ｐゴシック', 'ｍｓ ゴシック'].includes(n)) return 8;
  if (s === 11 && n === 'ms pゴシック') return 8;
  if (s === 11 && n === 'ms ゴシック') return 8;
  return undefined;
}

