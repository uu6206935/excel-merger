/**
 * Excel-like interpretation of what the user typed into a cell (setCellInput).
 *
 *  ''                    -> clear (style kept)
 *  "'text"               -> text without the apostrophe (Excel's quote prefix)
 *  '=...'                -> formula (UI form -> FILE form via forms.toFileForm), no cached value (t 'z')
 *  TRUE / FALSE          -> boolean (half / full width, any case)
 *  #N/A, #REF! ...       -> error value
 *  numbers, dates, times -> Excel serials, WITH the number format Excel-ja gives the cell (see below)
 *  anything else         -> text
 *
 * Ground truth: every rule below was measured with Excel 365 (ja-JP) via COM (Range.FormulaLocal = typing) over
 * 382 inputs x up to 23 cell formats (4,554 cases); the results are the fixture tests/doc/fixtures/typed-excel.json that
 * tests/doc/typed.test.ts replays through parseInput. (9:30:15.5 is stored by Excel as built-in id 47, which readBook
 * spells 'mmss.0'; the doc writes the equivalent custom code mm:ss.0.)
 *
 * Input classes and the format Excel gives a General cell:
 *   plain 123 / 1.5 / (1000)       -> (none)
 *   thousands 1,000 / 1,000.5      -> #,##0 / #,##0.00        (groups after the first comma have >= 3 digits)
 *   percent 50% / 1.5% / %50       -> 0% / 0.00%              (value / 100)
 *   yen ¥1,000 / \1,000 / ¥1.5     -> "¥"#,##0;[Red]"¥"\-#,##0 (/ .00)
 *   dollar $1,000                  -> \$#,##0_);[Red]\(\$#,##0\) (/ .00)
 *   scientific 1e3                 -> 0.00E+00
 *   fraction 0 3/4 / 1 1/16        -> # ?/? / # ??/??         (a plain "3/4" is a fraction only in numeric cells)
 *   2 3/4%                         -> 0.00%                   (value 0.0275)
 *   dates 2026/10/20, 26-10-20     -> yyyy/m/d;  10/20, 10-20, 10月20日 -> m"月"d"日" (current year; 13/5 -> 5月13日)
 *         2026年10月20日 -> yyyy"年"m"月"d"日";  2026年10月 -> yyyy"年"m"月";  2026/10 -> mmm-yy
 *         R8.10.20, h30/1/5 -> [$-411]ge.m.d;  令和8年10月20日 -> [$-411]ggge"年"m"月"d"日" (元年: gannen code)
 *         20-Oct, Oct 20 -> d-mmm;  20-Oct-2026 / October 20, 2026 -> d-mmm-yy;  Oct-2026 -> mmm-yy
 *   times 9:30 -> h:mm; 9:30:15 -> h:mm:ss; 25:00 -> [h]:mm:ss; 9:30:15.5 -> mm:ss.0; 9:30 PM -> h:mm AM/PM;
 *         9時30分 -> h"時"mm"分" (+ss"秒");  2026/10/20 9:30 -> yyyy/m/d h:mm
 *   odd 9:60 / 2026/10/20 25:00    -> a plain number, and the cell format becomes General
 *
 * A cell that already has a number format keeps it when the input "fits" it (formatCategory / KEEPS):
 * number / currency / custom formats keep everything (so does an accounting format that shows a currency symbol);
 * percent keeps percent; date keeps dates; time keeps times ...; otherwise the class format replaces it (a date
 * typed into a 0% cell becomes yyyy/m/d).
 * ParsedInput.numFmt is set only when the cell's format changes.
 *
 * Deliberate deviations from Excel:
 *  - Integers with more than 15 digits stay text (Excel would round them: 1234567890123456789 -> ...800).
 *  - Columns that hold text-typed digit IDs ('00123') keep digit-only input as text (leading zeros survive).
 *  - Automatic percent entry (Excel UI option, on by default; not applied by COM): a plain number >= 1 typed into
 *    a percent cell is divided by 100 (50 -> 50%), a number < 1 is kept (0.5 -> 50%).
 * Text-formatted cells ('@') keep everything as text, formulas included (this IS Excel behaviour).
 *
 * Quote prefix (ParsedInput.quote -> CellWrite.cell.quote): a text result is flagged when Excel would re-parse it as
 * a number / date / time / percent / currency / boolean / error / formula on the next edit — the text the user typed
 * with a leading apostrophe ("'00123" -> "00123", flagged whatever follows, like Excel), digit-only input kept as text
 * in an ID column, and integers of more than 15 digits. Ordinary text ('abc', '03-1234-5678') and everything typed
 * into a text-formatted cell are not flagged (needsQuotePrefix is the predicate).
 */
import { isFormulaInput, toFileForm } from '../formula/forms';
import type { FormOptions } from '../formula/forms';
import { renderArea, tokenize } from '../formula/excelRef';
import type { CellType } from '../types';

export type ParsedInput =
  | { kind: 'clear' }
  /**
   * numFmt: the number format code the cell gets (only when it changes; 'General' = back to General).
   * quote: (text results only, set ONLY when true) the text must be stored with Excel's quote prefix (quotePrefix="1",
   * the apostrophe the formula bar shows) because Excel would otherwise re-parse it as a number / date / time /
   * percent / currency / boolean / error / formula when the cell is edited again — see needsQuotePrefix.
   */
  | { kind: 'value'; t: CellType; v: string | number | boolean; numFmt?: string; hint?: string; quote?: boolean }
  | { kind: 'formula'; f: string }
  | { kind: 'error'; reason: string };

export interface ParseContext {
  /** number format CODE of the target cell (undefined = General) */
  numFmt?: string;
  /** the column holds text-typed digit IDs ('00123') */
  textIds?: boolean;
  /** "today" for dates without a year (tests inject it) */
  now?: Date;
  /** table context for formulas ([@col] inside a table, bare table names) */
  formOptions?: FormOptions;
}

const MAX_TEXT = 32767;
const MAX_NUM = 9.99999999999999e307;
const MIN_NUM = 2.2250738585072e-308;

// ---------------------------------------------------------------------------------------------
// width folding
// ---------------------------------------------------------------------------------------------
/**
 * Full-width ASCII variants (！..～) -> ASCII, ideographic space -> space, ￥ -> ¥. Other dashes (−, ‐, –) are NOT
 * folded: Excel keeps "−１２３" (U+2212) as text.
 */
export function foldWidth(s: string): string {
  let out = '';
  for (const ch of s) {
    const c = ch.charCodeAt(0);
    if (c >= 0xff01 && c <= 0xff5e) out += String.fromCharCode(c - 0xfee0);
    else if (c === 0x3000) out += ' ';
    else if (ch === '￥') out += '¥';
    else out += ch;
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// number formats
// ---------------------------------------------------------------------------------------------
const bareFmt = (code: string): string =>
  code
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/_./g, '')
    .replace(/\*./g, '')
    .replace(/\[(?![hms]+\])[^\]]*\]/gi, '')
    .toLowerCase();

/** true when a number format code shows a date and/or time. */
export function isDateFormat(code: string | undefined): boolean {
  if (!code || code === 'General' || code === '@') return false;
  const bare = bareFmt(code).replace(/general/g, '');
  // 'e' followed by +/- is scientific notation; 'aaa' is the Japanese weekday
  return /[ymdhs]|g{1,3}|a{3,4}|e(?![+-])/.test(bare.replace(/e[+-]/g, ''));
}

/** '@' (or a first section built around '@', e.g. "ID:"@): everything typed stays text. */
export const isTextFormat = (code: string | undefined): boolean => !!code && bareFmt(code).split(';')[0].includes('@');

export const isPercentFormat = (code: string | undefined): boolean => !!code && bareFmt(code).includes('%');

export type FormatCategory = 'general' | 'text' | 'number' | 'percent' | 'scientific' | 'fraction' | 'accounting' | 'date' | 'time' | 'datetime';

/** What kind of format a cell has, as far as typing into it is concerned (measured: see the header). */
export function formatCategory(code: string | undefined): FormatCategory {
  if (!code || /^general$/i.test(code) || code === 'G/標準') return 'general';
  if (isTextFormat(code)) return 'text';
  if (isDateFormat(code)) {
    const b = bareFmt(code).split(';')[0].replace(/e[+-]/g, '');
    const time = /[hs]/.test(b);
    const date = /[ydge]|a{3}/.test(b) || (/m/.test(b) && !time);
    return date && time ? 'datetime' : time ? 'time' : 'date';
  }
  // a format showing a currency symbol ("¥"#,##0, _ "¥"* #,##0_ ...) keeps everything typed into it, like '0'
  if (/[¥$€£￥]|\[\$[^\]-]/.test(code)) return 'number';
  // accounting formats without a currency symbol fill with '*' (_ * #,##0_ ; ...)
  if (/\*/.test(code.replace(/"[^"]*"/g, '').replace(/\\./g, ''))) return 'accounting';
  const bare = bareFmt(code).split(';')[0];
  if (bare.includes('%')) return 'percent';
  if (/e[+-]/.test(bare)) return 'scientific';
  if (/[#0?]\s*\/\s*[#0-9?]/.test(bare)) return 'fraction';
  return 'number';
}

// ---------------------------------------------------------------------------------------------
// input classes
// ---------------------------------------------------------------------------------------------
export type InputClass = 'plain' | 'thousands' | 'percent' | 'currency' | 'scientific' | 'fraction' | 'date' | 'datetime' | 'time' | 'odd';

export interface Typed {
  value: number;
  cls: InputClass;
  /** the format Excel gives a General cell (null = none) */
  code: string | null;
  /** integer digits of a plain number (> 15 cannot be stored exactly) */
  digits?: number;
}

export const FMT = {
  thousands: '#,##0',
  thousands2: '#,##0.00',
  percent: '0%',
  percent2: '0.00%',
  yen: '"¥"#,##0;[Red]"¥"\\-#,##0',
  yen2: '"¥"#,##0.00;[Red]"¥"\\-#,##0.00',
  dollar: '\\$#,##0_);[Red]\\(\\$#,##0\\)',
  dollar2: '\\$#,##0.00_);[Red]\\(\\$#,##0.00\\)',
  scientific: '0.00E+00',
  fraction1: '# ?/?',
  fraction2: '# ??/??',
  ymd: 'yyyy/m/d',
  md: 'm"月"d"日"',
  ymdKanji: 'yyyy"年"m"月"d"日"',
  ymKanji: 'yyyy"年"m"月"',
  ym: 'mmm-yy',
  dmmm: 'd-mmm',
  dmmmyy: 'd-mmm-yy',
  eraShort: '[$-411]ge.m.d',
  eraLong: '[$-411]ggge"年"m"月"d"日"',
  eraGannen: '[$]ggge"年"m"月"d"日";@',
  hm: 'h:mm',
  hms: 'h:mm:ss',
  hOver: '[h]:mm:ss',
  msTenths: 'mm:ss.0',
  hmAmPm: 'h:mm AM/PM',
  hmsAmPm: 'h:mm:ss AM/PM',
  hmKanji: 'h"時"mm"分"',
  hmsKanji: 'h"時"mm"分"ss"秒"',
  datetime: 'yyyy/m/d h:mm',
  general: 'General',
} as const;

/** Input classes a cell format keeps (the class format is applied otherwise). */
const KEEPS: Record<Exclude<FormatCategory, 'general' | 'text' | 'number'>, ReadonlySet<InputClass>> = {
  percent: new Set(['plain', 'thousands', 'percent']),
  scientific: new Set(['plain', 'thousands', 'scientific']),
  fraction: new Set(['plain', 'thousands', 'fraction']),
  accounting: new Set(['plain', 'thousands']),
  date: new Set(['plain', 'thousands', 'date', 'datetime']),
  time: new Set(['plain', 'thousands', 'time']),
  datetime: new Set(['plain', 'thousands', 'date', 'datetime']),
};

/** In these cells "3/4" is the fraction 0.75, elsewhere the date 3月4日. */
const FRACTION_FIRST: ReadonlySet<FormatCategory> = new Set(['number', 'percent', 'scientific', 'fraction', 'accounting']);

// ---------------------------------------------------------------------------------------------
// numbers
// ---------------------------------------------------------------------------------------------
const reBody = /^(?:(\d+(?:,\d{3,})*)(\.\d*)?|(\.\d+))(?:[eE]([+-]?\d+))?$/;

/** Parse a (width-folded, trimmed) number with its input class; null when not a number. */
export function parseNumberInput(t: string): Typed | null {
  let s = t;
  let neg = false;
  let sign = '';
  let cur = '';
  let m = /^\((.*)\)$/.exec(s);
  if (m) {
    s = m[1].trim();
    neg = true;
  }
  for (let k = 0; k < 2; k++) {
    if (!sign && !neg && (s[0] === '+' || s[0] === '-')) {
      sign = s[0];
      s = s.slice(1).trimStart();
    } else if (!cur && (s[0] === '¥' || s[0] === '$' || s[0] === '€' || s[0] === '\\')) {
      cur = s[0] === '\\' ? '¥' : s[0];
      s = s.slice(1).trimStart();
    } else break;
  }
  if (cur && !neg && !sign && (m = /^\((.*)\)$/.exec(s))) {
    s = m[1].trim();
    neg = true;
  }
  let pct = false;
  if (s[0] === '%') {
    pct = true;
    s = s.slice(1).trimStart();
  } else if ((m = /^(.*?)\s*%$/.exec(s))) {
    pct = true;
    s = m[1];
  }
  if (pct && cur) return null; // ¥50% is text
  const b = reBody.exec(s);
  if (!b) return null;
  const [, int, frac, onlyFrac, exp] = b;
  const mant = int !== undefined ? int.replace(/,/g, '') + (frac ?? '') : onlyFrac;
  let v = Number(mant + (exp !== undefined ? 'e' + exp : ''));
  if (!Number.isFinite(v) || Math.abs(v) > MAX_NUM) return null;
  if (v !== 0 && Math.abs(v) < MIN_NUM) return null;
  if (v === 0 && /[1-9]/.test(mant)) return null; // underflow (1e-400) is text in Excel
  if (pct) v /= 100;
  if (neg || sign === '-') v = -v;
  if (Object.is(v, -0)) v = 0;
  const dot = frac !== undefined || onlyFrac !== undefined;
  const digits = int !== undefined && !dot && exp === undefined ? int.replace(/\D/g, '').replace(/^0+/, '').length : 0;
  if (exp !== undefined) return { value: v, cls: 'scientific', code: FMT.scientific };
  if (pct) return { value: v, cls: 'percent', code: dot ? FMT.percent2 : FMT.percent };
  if (cur === '¥') return { value: v, cls: 'currency', code: dot ? FMT.yen2 : FMT.yen };
  if (cur === '$') return { value: v, cls: 'currency', code: dot ? FMT.dollar2 : FMT.dollar };
  if (int !== undefined && int.includes(',')) return { value: v, cls: 'thousands', code: dot ? FMT.thousands2 : FMT.thousands };
  return { value: v, cls: 'plain', code: null, digits };
}

/** Kept for callers of the round-1 API: plain value of a number input. */
export function parseNumber(s: string): { value: number; percent: boolean; digits: number } | null {
  const r = parseNumberInput(s);
  return r ? { value: r.value, percent: r.cls === 'percent', digits: r.digits ?? 0 } : null;
}

const fractionCode = (den: number) => (den < 10 ? FMT.fraction1 : FMT.fraction2);

/** "1 1/2", "0 3/4", "-1 1/2" (any cell); "2 3/4%" is the percentage 2.75% (0.00%) */
function parseMixedFraction(t: string): Typed | null {
  const m = /^([+-]?)(\d+)\s+(\d+)\/(\d+)(\s*%)?$/.exec(t);
  if (!m) return null;
  const den = Number(m[4]);
  if (!(den >= 1)) return null;
  let v = Number(m[2]) + Number(m[3]) / den;
  if (m[1] === '-') v = -v;
  if (Object.is(v, -0)) v = 0;
  if (m[5] !== undefined) return { value: v / 100, cls: 'percent', code: FMT.percent2 };
  return { value: v, cls: 'fraction', code: fractionCode(den) };
}

/** "3/4", "-10/20" (numeric cells only; the numerator must be >= 1) */
function parseSimpleFraction(t: string): Typed | null {
  const m = /^([+-]?)(\d+)\/(\d+)$/.exec(t);
  if (!m) return null;
  const num = Number(m[2]);
  const den = Number(m[3]);
  if (!(num >= 1) || !(den >= 1)) return null;
  const v = (m[1] === '-' ? -num : num) / den;
  return { value: v, cls: 'fraction', code: fractionCode(den) };
}

// ---------------------------------------------------------------------------------------------
// dates
// ---------------------------------------------------------------------------------------------
const ERAS: Record<string, number> = {
  m: 1868, 明治: 1868,
  t: 1912, 大正: 1912,
  s: 1926, 昭和: 1926,
  h: 1989, 平成: 1989,
  r: 2019, 令和: 2019,
};

/**
 * Excel serial (1900 date system) of a calendar date; null when invalid. Like Excel, 1900-02-29 exists (60) and
 * serials before it are one less than the true day count.
 */
export function dateSerial(y: number, m: number, d: number): number | null {
  if (!Number.isInteger(y) || !Number.isInteger(m) || !Number.isInteger(d)) return null;
  if (y < 1900 || y > 9999 || m < 1 || m > 12 || d < 1 || d > 31) return null;
  if (y === 1900 && m === 2 && d === 29) return 60;
  const ms = Date.UTC(y, m - 1, d);
  const dt = new Date(ms);
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
  const serial = Math.round((ms - Date.UTC(1899, 11, 30)) / 86400000);
  return serial < 61 ? serial - 1 : serial;
}

/** two-digit years: 00-29 -> 2000s, 30-99 -> 1900s (Excel's default cutoff) */
const fullYear = (y: string): number => (y.length <= 2 ? (Number(y) < 30 ? 2000 : 1900) + Number(y) : Number(y));

const MONTHS: Record<string, number> = {
  jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12,
  january: 1, february: 2, march: 3, april: 4, june: 6, july: 7, august: 8, september: 9, october: 10, november: 11, december: 12,
};
const monthOf = (name: string): number | undefined => MONTHS[name.toLowerCase()];

interface TimeOf {
  frac: number;
  code: string;
  /** minutes / seconds >= 60: a plain number, format General */
  odd: boolean;
  /** hours >= 24 */
  over24: boolean;
}

function timeOf(h: number, mi: number, sec: number, code: string): TimeOf {
  return { frac: (h * 3600 + mi * 60 + sec) / 86400, code, odd: mi >= 60 || sec >= 60, over24: h >= 24 };
}

/** A time ("9:30", "9:30:15", "25:00", "9:30:15.5", "30:15.5", "9:30 PM", "9時30分"), or null. */
function parseTime(s: string): TimeOf | null {
  let m: RegExpExecArray | null;
  if ((m = /^(\d{1,2})時(\d{1,2})分(?:(\d{1,2})秒)?$/.exec(s))) {
    return timeOf(+m[1], +m[2], m[3] !== undefined ? +m[3] : 0, m[3] !== undefined ? FMT.hmsKanji : FMT.hmKanji);
  }
  if ((m = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s+([ap])m?$/i.exec(s))) {
    let h = +m[1];
    if (h > 12) return null;
    const pm = m[4].toLowerCase() === 'p';
    if (h === 12) h = pm ? 12 : 0;
    else if (pm) h += 12;
    return timeOf(h, +m[2], m[3] !== undefined ? +m[3] : 0, m[3] !== undefined ? FMT.hmsAmPm : FMT.hmAmPm);
  }
  if ((m = /^(\d{1,4}):(\d{1,2})(?::(\d{1,2})(\.\d+)?)?$/.exec(s))) {
    const h = +m[1];
    if (m[4] !== undefined) return timeOf(h, +m[2], Number(m[3] + m[4]), FMT.msTenths);
    const sec = m[3] !== undefined ? +m[3] : 0;
    return timeOf(h, +m[2], sec, h >= 24 ? FMT.hOver : m[3] !== undefined ? FMT.hms : FMT.hm);
  }
  // minutes:seconds.fraction
  if ((m = /^(\d{1,4}):(\d{1,2}\.\d+)$/.exec(s))) return timeOf(0, +m[1], Number(m[2]), FMT.msTenths);
  return null;
}

type DateOf = { serial: number; code: string };
const dateOut = (serial: number | null, code: string): DateOf | null => (serial === null ? null : { serial, code });

/** Numeric dates that may carry a time: y/m/d, yy/m/d, y-m-d, m/d, m-d. */
function parseNumericDate(s: string, now: Date): DateOf | null {
  let m: RegExpExecArray | null;
  if ((m = /^(\d{4}|\d{1,2})([/-])(\d{1,2})\2(\d{1,2})$/.exec(s))) return dateOut(dateSerial(fullYear(m[1]), +m[3], +m[4]), FMT.ymd);
  if ((m = /^(\d{1,2})[/-](\d{1,2})$/.exec(s))) {
    const y = now.getFullYear();
    // month first; 13/5 is read as 5月13日
    return dateOut(dateSerial(y, +m[1], +m[2]) ?? (+m[1] > 12 ? dateSerial(y, +m[2], +m[1]) : null), FMT.md);
  }
  return null;
}

function parseDateOnly(s: string, now: Date): DateOf | null {
  const num = parseNumericDate(s, now);
  if (num) return num;
  let m: RegExpExecArray | null;
  const y0 = now.getFullYear();
  if ((m = /^(\d{4})[/-](\d{1,2})$/.exec(s))) return dateOut(dateSerial(+m[1], +m[2], 1), FMT.ym);
  if ((m = /^(\d{4})年(\d{1,2})月(\d{1,2})日$/.exec(s))) return dateOut(dateSerial(+m[1], +m[2], +m[3]), FMT.ymdKanji);
  if ((m = /^(\d{4})年(\d{1,2})月$/.exec(s))) return dateOut(dateSerial(+m[1], +m[2], 1), FMT.ymKanji);
  if ((m = /^(\d{1,2})月(\d{1,2})日$/.exec(s))) return dateOut(dateSerial(y0, +m[1], +m[2]), FMT.md);
  // R8.10.20, R8/10/20, H30-1-5 (the era year may run past the era's end: H31.5.1 = 2019/5/1, like Excel)
  if ((m = /^([MTSHRmtshr])(\d{1,2})([/.-])(\d{1,2})\3(\d{1,2})$/.exec(s))) {
    const n = +m[2];
    return n >= 1 ? dateOut(dateSerial(ERAS[m[1].toLowerCase()] + n - 1, +m[4], +m[5]), FMT.eraShort) : null;
  }
  // 令和8年10月20日, 令和元年5月1日
  if ((m = /^(明治|大正|昭和|平成|令和)(\d{1,2}|元)年(\d{1,2})月(\d{1,2})日$/.exec(s))) {
    const gannen = m[2] === '元';
    const n = gannen ? 1 : +m[2];
    return n >= 1 ? dateOut(dateSerial(ERAS[m[1]] + n - 1, +m[3], +m[4]), gannen ? FMT.eraGannen : FMT.eraLong) : null;
  }
  // English month names: 20-Oct, Oct-20, Oct 20, Oct-2026, 20-Oct-2026, 20 Oct 2026, October 20, 2026
  if ((m = /^(\d{1,2})[-\s]([A-Za-z]{3,9})[-\s](\d{4}|\d{1,2})$/.exec(s))) {
    const mo = monthOf(m[2]);
    return mo ? dateOut(dateSerial(fullYear(m[3]), mo, +m[1]), FMT.dmmmyy) : null;
  }
  if ((m = /^([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{4})$/.exec(s))) {
    const mo = monthOf(m[1]);
    return mo ? dateOut(dateSerial(+m[3], mo, +m[2]), FMT.dmmmyy) : null;
  }
  if ((m = /^(\d{1,2})[-\s]([A-Za-z]{3,9})$/.exec(s))) {
    const mo = monthOf(m[2]);
    return mo ? dateOut(dateSerial(y0, mo, +m[1]), FMT.dmmm) : null;
  }
  if ((m = /^([A-Za-z]{3,9})[-\s](\d{4}|\d{1,2})$/.exec(s))) {
    const mo = monthOf(m[1]);
    if (!mo) return null;
    if (m[2].length === 4) return dateOut(dateSerial(+m[2], mo, 1), FMT.ym);
    return dateOut(dateSerial(y0, mo, +m[2]), FMT.dmmm) ?? dateOut(dateSerial(fullYear(m[2]), mo, 1), FMT.ym);
  }
  return null;
}

/** Date and/or time (right-trimmed, width-folded text) with its class; null when not one. */
export function parseDateTimeInput(s: string, now: Date): Typed | null {
  const t = s.replace(/\s+/g, ' ');
  const odd = (v: number): Typed => ({ value: v, cls: 'odd', code: FMT.general });
  const tm = parseTime(t);
  if (tm) return tm.odd ? odd(tm.frac) : { value: tm.frac, cls: 'time', code: tm.code };
  const sp = /^(\S+) (.+)$/.exec(t);
  if (sp) {
    const d = parseNumericDate(sp[1], now);
    const tt = d ? parseTime(sp[2]) : null;
    if (d && tt && !tt.code.includes('時')) {
      const v = d.serial + tt.frac;
      if (tt.odd || tt.over24) return odd(v);
      if (tt.code === FMT.msTenths) return { value: v, cls: 'time', code: FMT.msTenths };
      return { value: v, cls: 'datetime', code: FMT.datetime };
    }
  }
  const d = parseDateOnly(t, now);
  return d ? { value: d.serial, cls: 'date', code: d.code } : null;
}

/** Round-1 API: serial of a date and/or time (strong = has a year / era / 年月日). */
export function parseDateTime(s: string, now: Date): { serial: number; strong: boolean } | null {
  const r = parseDateTimeInput(s, now);
  return r ? { serial: r.value, strong: r.cls === 'date' || r.cls === 'datetime' } : null;
}

// ---------------------------------------------------------------------------------------------
// formulas
// ---------------------------------------------------------------------------------------------
/** Light syntax check so that nothing Excel would reject (and "repair" in a file) is ever written. */
export function formulaProblem(fileForm: string): string | null {
  if (!fileForm.trim()) return '数式が空です。';
  if (fileForm.length > 8192) return '数式が長すぎます（Excel の上限は 8,192 文字です）。';
  const toks = tokenize(fileForm);
  let depth = 0;
  let brace = 0;
  let last = '';
  for (const t of toks) {
    if (t.type === 'unknown') {
      if (t.text.startsWith('"')) return '数式の文字列（"…"）が閉じられていません。';
      return `数式に解釈できない部分があります: ${t.text.slice(0, 20)}`;
    }
    if (t.type === 'ws') continue;
    if (t.type === 'op') {
      if (t.text === '(') depth++;
      else if (t.text === ')') {
        depth--;
        if (depth < 0) return '数式のかっこ「)」が多すぎます。';
      } else if (t.text === '{') brace++;
      else if (t.text === '}') {
        brace--;
        if (brace < 0) return '数式のかっこ「}」が多すぎます。';
      }
    }
    last = t.type === 'op' ? t.text : t.type;
  }
  if (depth > 0) return '数式のかっこ「(」が閉じられていません。';
  if (brace > 0) return '数式のかっこ「{」が閉じられていません。';
  if (['+', '-', '*', '/', '^', '&', '=', '<', '>', '<=', '>=', '<>', ',', ':', '('].includes(last)) return '数式が途中で終わっています。';
  return null;
}

/**
 * Excel normalizes what the user typed: function names, cell references, TRUE/FALSE and error literals are
 * upper-cased (=sum(a1:b2) is stored as SUM(A1:B2)). Strings, names, sheet names, structured references and
 * LET/LAMBDA parameters keep their spelling.
 */
export function normalizeFormulaCase(f: string): string {
  const toks = tokenize(f);
  let changed = false;
  const out = toks.map((t) => {
    let x = t.text;
    if (t.type === 'func' && !t.prefix) {
      if (/^_xlpm\./i.test(x)) return x;
      const m = /^((?:_xlfn\.|_xlws\.)*)(.*)$/is.exec(x)!;
      x = m[1] + m[2].toUpperCase();
    } else if (t.type === 'ref' && t.area) {
      x = (t.prefix ? t.prefix.text : '') + renderArea(t.area);
    } else if (t.type === 'bool') x = x.toUpperCase();
    else if (t.type === 'error') {
      const pre = t.prefix ? t.prefix.text : '';
      x = pre + x.slice(pre.length).toUpperCase();
    }
    if (x !== t.text) changed = true;
    return x;
  });
  return changed ? out.join('') : f;
}

// ---------------------------------------------------------------------------------------------
// entry point
// ---------------------------------------------------------------------------------------------
/** Error values Excel accepts as typed input (#SPILL! / #CALC! stay text; #GETTING_DATA becomes #N/A). */
const ERRORS: Record<string, string> = {
  '#NULL!': '#NULL!', '#DIV/0!': '#DIV/0!', '#VALUE!': '#VALUE!', '#REF!': '#REF!', '#NAME?': '#NAME?', '#NUM!': '#NUM!',
  '#N/A': '#N/A', '#GETTING_DATA': '#N/A',
};

/** The class of a typed value for a cell of category `cat` (null = text). */
export function typedValue(folded: string, cat: FormatCategory, now: Date): Typed | null {
  const t = folded.trim();
  const num = parseNumberInput(t);
  if (num) return num;
  const mixed = parseMixedFraction(t);
  if (mixed) return mixed;
  if (FRACTION_FIRST.has(cat)) {
    const fr = parseSimpleFraction(t);
    if (fr) return fr;
  }
  // dates / times: trailing spaces are fine, leading ones make it text (measured)
  if (/^\s/.test(folded)) return null;
  return parseDateTimeInput(folded.trimEnd(), now);
}

const sameCode = (a: string | undefined, b: string | undefined) => (a ?? 'General').toLowerCase() === (b ?? 'General').toLowerCase();

/** A text result that the doc keeps as text on purpose: quote-prefixed when Excel would re-parse it (needsQuotePrefix). */
function textValue(text: string, ctx: ParseContext): { kind: 'value'; t: 's'; v: string; quote?: boolean } {
  const out: { kind: 'value'; t: 's'; v: string; quote?: boolean } = { kind: 'value', t: 's', v: text };
  if (needsQuotePrefix(text, ctx.numFmt, ctx.now)) out.quote = true;
  return out;
}

/**
 * The quote-prefix predicate: would Excel turn the TEXT `text` into something that is not text when it is typed into
 * (or the cell is edited again in) a cell of number format `numFmt`? True for what the typed-input rules above read
 * as a number / date / time / percent / currency / fraction / scientific, TRUE / FALSE, an error literal or a
 * formula ('=…'); false for ordinary text ('03-1234-5678' is not a date, 'abc', spaces). A text-formatted cell ('@')
 * keeps everything as text by itself, so it never needs the prefix. Uses exactly the helpers parseInput uses (same
 * width folding, same typedValue, same "today" for year-less dates), so it cannot drift from what typing does.
 */
export function needsQuotePrefix(text: string, numFmt?: string, now: Date = new Date()): boolean {
  if (text === '' || isTextFormat(numFmt)) return false;
  if (isFormulaInput(text)) return true;
  const folded = foldWidth(text);
  const trimmed = folded.trim();
  if (trimmed === '') return false;
  const upper = trimmed.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE' || ERRORS[upper]) return true;
  return typedValue(folded, formatCategory(numFmt), now) !== null;
}

export function parseInput(text: string, ctx: ParseContext = {}): ParsedInput {
  if (text === '') return { kind: 'clear' };
  if (text.length > MAX_TEXT) return { kind: 'error', reason: `入力が長すぎます（1 つのセルに入る文字数は ${MAX_TEXT.toLocaleString()} 文字までです）。` };
  // text-formatted cell: everything is text (Excel keeps even '=A1' as text there; an apostrophe is content too).
  // No quote prefix: the cell format keeps the text a text when the cell is edited again.
  if (isTextFormat(ctx.numFmt)) return { kind: 'value', t: 's', v: text };
  // a leading apostrophe is Excel's quote prefix: the text without it, stored with quotePrefix (whatever follows:
  // Excel keeps the prefix on 'abc too)
  if (text.startsWith("'")) return { kind: 'value', t: 's', v: text.slice(1), quote: true };
  if (isFormulaInput(text)) {
    let f: string;
    try {
      f = normalizeFormulaCase(toFileForm(text, ctx.formOptions));
    } catch {
      return { kind: 'error', reason: '数式を解釈できませんでした。入力内容を確認してください。' };
    }
    const why = formulaProblem(f);
    if (why) return { kind: 'error', reason: why };
    return { kind: 'formula', f };
  }
  const folded = foldWidth(text);
  const trimmed = folded.trim();
  if (trimmed === '') return { kind: 'value', t: 's', v: text };
  // IDs typed into an ID column keep their exact text (leading zeros): stored like '00123 typed with an apostrophe
  if (ctx.textIds && /^[0-9]+$/.test(trimmed)) return textValue(text, ctx);
  const upper = trimmed.toUpperCase();
  if (upper === 'TRUE') return { kind: 'value', t: 'b', v: true };
  if (upper === 'FALSE') return { kind: 'value', t: 'b', v: false };
  const err = ERRORS[upper];
  if (err) return { kind: 'value', t: 'e', v: err };

  const cat = formatCategory(ctx.numFmt);
  const typed = typedValue(folded, cat, ctx.now ?? new Date());
  if (!typed) return { kind: 'value', t: 's', v: text };
  if ((typed.digits ?? 0) > 15) {
    return { ...textValue(text, ctx), hint: '15 桁を超える数字は、Excel では正確に保存できないため文字列として入力しました。' };
  }
  let v = typed.value;
  // automatic percent entry (Excel UI default): 50 typed into a percent cell is 50%, 0.5 is 50% too
  if (cat === 'percent' && (typed.cls === 'plain' || typed.cls === 'thousands') && Math.abs(v) >= 1) v /= 100;
  const keep = cat === 'number' || (cat !== 'general' && cat !== 'text' && KEEPS[cat].has(typed.cls));
  let finalFmt = ctx.numFmt;
  let numFmt: string | undefined;
  if (!keep && typed.code !== null && !sameCode(typed.code, ctx.numFmt)) {
    numFmt = typed.code;
    finalFmt = typed.code;
  }
  const out: ParsedInput = { kind: 'value', t: isDateFormat(finalFmt) ? 'd' : 'n', v };
  if (numFmt !== undefined) out.numFmt = numFmt;
  return out;
}
