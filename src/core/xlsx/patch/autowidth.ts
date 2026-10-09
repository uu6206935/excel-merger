/**
 * Excel's automatic column widening for number formats (owner: patcher P).
 *
 * When a number format is applied to a number that no longer fits its column (it would show as #####), Excel widens
 * a column of STANDARD width (no customWidth, or an auto-fitted one: bestFit) to fit the formatted text; a column
 * whose width the user set is left alone. Measured with tests/golden/com/probe-numfmt-width.ps1 (Excel 365 ja-JP,
 * 游ゴシック 11 pt, Value2 then Range.NumberFormat — the order of goldens r-numfmt / r-numfmt-w): 2026/1/1 -> 8.63,
 * 2026/9/30 -> 9.63, 2026/10/10 -> 10.75, 1,234,567 -> 8.88, 123,456,789 -> 11, $1,234,567 -> 10, 12345.60% -> 9.88,
 * 2026/9/30 9:30 -> 14.13, 2026年9月30日 -> 14 (golden r-numfmt-w); 12:34:20, R8.9.30, 50 %, 9/30 fit the default
 * 8.38 and do not widen. Excel stores such a column as <col width="(px + 5) / MDW" bestFit="1" customWidth="1"/>.
 * (Through COM, setting the format BEFORE the value, or entering text with Range.Formula, does not widen.)
 *
 * The text is measured with a glyph-advance table of 游ゴシック 11 pt (px at 96 dpi, fitted to the probe: digits
 * 8.5 px, Excel adds ~4 px) scaled by the maximum digit width of the cell's font. It reproduces every probe value
 * above within 1 px; other fonts are approximations.
 */
import { mdwFor } from '../../display/styles';

/** advance widths (px) of 游ゴシック 11 pt for the characters number formats produce */
const ADV: Record<string, number> = {
  '/': 7, '-': 7, '−': 7, '.': 5, ',': 3.75, ':': 5, '%': 11, '$': 9, '¥': 9, '\\': 9, ' ': 5.5, '(': 5.5, ')': 5.5,
  '+': 8.5, '*': 7, '[': 5, ']': 5, '"': 6, "'": 3.5, '?': 7.5, '#': 9, '@': 13, '&': 10, '_': 7.5, '=': 8.5,
};
const DIGIT = 8.5;
const MARGIN = 4;

/** Width in px that Excel's autofit gives the text `text` in a font with maximum digit width `mdw` (8 = 游ゴシック 11). */
export function numberTextPx(text: string, mdw = 8): number {
  let w = 0;
  for (const ch of text) {
    const cp = ch.codePointAt(0) ?? 0;
    if (cp >= 0x30 && cp <= 0x39) w += DIGIT;
    else if (ADV[ch] !== undefined) w += ADV[ch];
    else if (cp >= 0x41 && cp <= 0x5a) w += 11; // A-Z
    else if (cp >= 0x61 && cp <= 0x7a) w += 9; // a-z
    else if (cp >= 0xff61 && cp <= 0xff9f) w += 7.5; // half-width katakana
    else if (cp >= 0x2e80) w += 16.17; // full-width (年 月 日 令和 ...): measured 2026年9月30日 -> 14 chars
    else w += 8;
  }
  return Math.floor((w + MARGIN) * (mdw / 8));
}

/** Max digit width of a font (name / size in points); unknown fonts are treated like 游ゴシック. */
export function fontMdw(name: string | undefined, sizePt: number | undefined): number {
  return mdwFor(name, sizePt && sizePt > 0 ? sizePt : 11);
}
