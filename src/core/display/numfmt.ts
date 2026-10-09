/**
 * Display-only number formatting (owner: grid/display G), on top of SheetJS SSF (bundled in `xlsx`).
 * SSF is ~Excel-exact for en-US codes but misses the ja-JP bits, so this wrapper adds:
 *   - ja-JP built-in numFmtId table (14 = yyyy/m/d in Japanese Excel, 27-36/50-58 = 和暦 etc.)
 *     and the reverse of ExcelJS's en-US defaults ("mm-dd-yy" for id 14);
 *   - 和暦 era tokens g/gg/ggg + e/ee and weekday tokens aaa/aaaa (SSF prints them literally);
 *   - bare ¥ / ￥ (SSF throws "unrecognized character ¥");
 *   - [DBNum1..4] / [$-411] / [$-F800] locale tags (SSF prints "$123" for [DBNum1][$-411]General);
 *   - font colour of the chosen section ([Red], [Blue], [Color10], [赤] ...), which SSF drops.
 *
 * STABLE API (imported by the reader and the keys module): formatValue, JA_BUILTIN_NUMFMT, normalizeNumFmt,
 * FormattedValue, dateToSerial. Additions for the grid: formatCached, isDateFormat, displayValue, editText.
 * Pure: no DOM, no React, never throws (on any failure it falls back to String(value)).
 */
import { SSF } from 'xlsx';
import { digitWise, generalText, pad2, positional, type DbStyle } from './dbnum';

export const JA_BUILTIN_NUMFMT: Record<number, string> = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'yyyy/m/d', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
  22: 'yyyy/m/d h:mm',
  27: '[$-411]ge.m.d', 28: '[$-411]ggge"年"m"月"d"日"', 29: '[$-411]ggge"年"m"月"d"日"', 30: 'm/d/yy', 31: 'yyyy"年"m"月"d"日"',
  32: 'h"時"mm"分"', 33: 'h"時"mm"分"ss"秒"', 34: 'yyyy"年"m"月"', 35: 'm"月"d"日"', 36: '[$-411]ge.m.d',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00 ;(#,##0.00)', 40: '#,##0.00 ;[Red](#,##0.00)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
  50: '[$-411]ge.m.d', 51: '[$-411]ggge"年"m"月"d"日"', 52: 'yyyy"年"m"月"', 53: 'm"月"d"日"', 54: '[$-411]ggge"年"m"月"d"日"',
  55: 'yyyy"年"m"月"', 56: 'm"月"d"日"', 57: '[$-411]ge.m.d', 58: '[$-411]ggge"年"m"月"d"日"',
};

/** ExcelJS reports built-in ids with its en-US codes; map those back to what Japanese Excel shows. */
const EXCELJS_EN_DEFAULTS: Record<string, string> = {
  'mm-dd-yy': 'yyyy/m/d',
  'm/d/yy "h":mm': 'yyyy/m/d h:mm',
};

export function normalizeNumFmt(fmt: string | number | undefined | null): string {
  if (fmt === undefined || fmt === null || fmt === '') return 'General';
  if (typeof fmt === 'number') return JA_BUILTIN_NUMFMT[fmt] ?? 'General';
  const mapped = EXCELJS_EN_DEFAULTS[fmt];
  if (mapped) return mapped;
  // System long date / time (Windows regional setting) -> ja-JP defaults
  if (/^\[\$-F800\]/i.test(fmt)) return 'yyyy"年"m"月"d"日"';
  if (/^\[\$-F400\]/i.test(fmt)) return 'h:mm:ss';
  return fmt;
}

// ---------------------------------------------------------------------------------------------
// Section splitting (respects "quoted", [bracketed] and \escaped characters)

function splitSections(fmt: string): string[] {
  const out: string[] = [];
  let cur = '';
  let q = false;
  let b = false;
  for (let i = 0; i < fmt.length; i++) {
    const c = fmt[i];
    if (c === '\\' && !q && i + 1 < fmt.length) {
      cur += c + fmt[++i];
      continue;
    }
    if (c === '"' && !b) q = !q;
    else if (c === '[' && !q) b = true;
    else if (c === ']' && !q) b = false;
    if (c === ';' && !q && !b) {
      out.push(cur);
      cur = '';
      continue;
    }
    cur += c;
  }
  out.push(cur);
  return out;
}

const COLOR_NAMES: Record<string, string> = {
  black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#00FF00', magenta: '#FF00FF', red: '#FF0000', white: '#FFFFFF', yellow: '#FFFF00',
  黒: '#000000', 青: '#0000FF', 水色: '#00FFFF', 緑: '#00FF00', 紫: '#FF00FF', 赤: '#FF0000', 白: '#FFFFFF', 黄: '#FFFF00',
};
/** Legacy 56-colour palette for [Color1]..[Color56]. */
const PALETTE56 = [
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF', '#800000', '#008000', '#000080', '#808000', '#800080', '#008080',
  '#C0C0C0', '#808080', '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF', '#000080', '#FF00FF', '#FFFF00', '#00FFFF',
  '#800080', '#800000', '#008080', '#0000FF', '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99', '#3366FF', '#33CCCC',
  '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696', '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
];

function sectionColor(section: string): string | undefined {
  const m = /\[(black|blue|cyan|green|magenta|red|white|yellow|黒|青|水色|緑|紫|赤|白|黄|color\s*(\d+))\]/i.exec(section);
  if (!m) return undefined;
  if (m[2]) {
    const i = Number(m[2]);
    return i >= 1 && i <= 56 ? PALETTE56[i - 1] : undefined;
  }
  return COLOR_NAMES[m[1].toLowerCase()] ?? COLOR_NAMES[m[1]];
}

function pickSection(sections: string[], v: unknown): string {
  if (sections.some((s) => /\[[<>=]/.test(s))) return sections[0]; // conditional sections: let SSF decide text, no colour guess
  if (typeof v === 'string') return sections[3] ?? (sections.length === 1 && sections[0].includes('@') ? sections[0] : '');
  const n = Number(v);
  if (n < 0 && sections.length >= 2) return sections[1];
  if (n === 0 && sections.length >= 3) return sections[2];
  return sections[0];
}

// ---------------------------------------------------------------------------------------------
// Japanese era / weekday substitution (value dependent, so only done when the tokens occur)

const ERAS = [
  { start: 20190501, name: '令和', short: '令', letter: 'R', base: 2018 },
  { start: 19890108, name: '平成', short: '平', letter: 'H', base: 1988 },
  { start: 19261225, name: '昭和', short: '昭', letter: 'S', base: 1925 },
  { start: 19120730, name: '大正', short: '大', letter: 'T', base: 1911 },
  { start: 18730101, name: '明治', short: '明', letter: 'M', base: 1867 },
];
const WEEK_SHORT = ['日', '月', '火', '水', '木', '金', '土'];

/** The format with "quoted" literals, [brackets] and \escapes removed (lower case). */
function bareFormat(fmt: string): string {
  return fmt.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '').toLowerCase();
}

/** true when the (unquoted, unbracketed) format uses 和暦/曜日 tokens that SSF does not know. */
function needsJaTokens(fmt: string): boolean {
  const bare = bareFormat(fmt);
  if (bare.includes('general')) return false;
  return /g|e(?![+-])|aaa/.test(bare);
}

/** Append a literal as a quoted string, merging with a directly preceding quoted string
 *  (SSF mis-parses adjacent literals such as "R""6"). */
function appendLit(out: string, lit: string): string {
  const t = lit.replace(/"/g, '');
  return out.endsWith('"') ? out.slice(0, -1) + t + '"' : out + '"' + t + '"';
}

/** The next date-token letter after index `from` (skipping "quotes", [brackets], separators); '' when there is none. */
function nextDateLetter(section: string, from: number): string {
  for (let k = from; k < section.length; k++) {
    const c = section[k];
    if (c === '"') k = section.indexOf('"', k + 1) < 0 ? section.length : section.indexOf('"', k + 1);
    else if (c === '[') k = section.indexOf(']', k + 1) < 0 ? section.length : section.indexOf(']', k + 1);
    else if (c === '\\') k++;
    else if (/[a-z]/i.test(c)) return c.toLowerCase();
  }
  return '';
}

/**
 * Replace g/gg/ggg, e/ee, aaa/aaaa outside quotes/brackets by quoted literals for this date.
 * With `db` ([DBNum1] / [DBNum2]) the year / month / day / hour / minute / second tokens become numerals as well:
 * yyyy and the two-letter mm dd hh ss ee digit by digit (二〇二六, 〇一), y yy m d h s e with the units (二十六, 十月, 三十一日).
 */
function substituteJa(section: string, serial: number, db?: DbStyle): string {
  const dc = SSF.parse_date_code(serial);
  if (!dc) return section;
  const ymd = dc.y * 10000 + dc.m * 100 + dc.d;
  const era = ERAS.find((e) => ymd >= e.start) ?? ERAS[ERAS.length - 1];
  const eraYear = dc.y - era.base;
  const wd = dc.q; // SSF: day of week, 0 = Sunday
  const h12 = /am\/pm|a\/p/i.test(bareFormat(section));
  const ja411 = jaNames(section);
  let prevTok = '';
  let out = '';
  let q = false;
  let b = false;
  for (let i = 0; i < section.length; i++) {
    const c = section[i];
    if (c === '\\' && !q) {
      out += c + (section[i + 1] ?? '');
      i++;
      continue;
    }
    if (c === '"' && !b) {
      q = !q;
      out += c;
      continue;
    }
    if (db && !q && c === '[') {
      // elapsed time [h] / [mm] / [ss] as numerals
      const end = section.indexOf(']', i);
      const inner = end > 0 ? section.slice(i + 1, end).toLowerCase() : '';
      if (/^(h+|m+|s+)$/.test(inner)) {
        const total = Math.floor(serial * (inner[0] === 'h' ? 24 : inner[0] === 'm' ? 1440 : 86400) + 1e-9);
        out = appendLit(out, inner.length >= 2 ? pad2(total, db) : positional(total, db));
        prevTok = inner[0];
        i = end;
        continue;
      }
    }
    if (!q && c === '[') b = true;
    if (!q && c === ']') {
      b = false;
      out += c;
      continue;
    }
    if (q || b) {
      out += c;
      continue;
    }
    const lc = c.toLowerCase();
    if (lc === 'g' || lc === 'e' || lc === 'a') {
      let j = i;
      while (j < section.length && section[j].toLowerCase() === lc) j++;
      const run = j - i;
      if (lc === 'e' && (section[j] === '+' || section[j] === '-')) {
        out += section.slice(i, j); // scientific notation, not era year
      } else if (lc === 'g') {
        out = appendLit(out, run >= 3 ? era.name : run === 2 ? era.short : era.letter);
      } else if (lc === 'e') {
        out = appendLit(out, db ? (run >= 2 ? pad2(eraYear, db) : positional(eraYear, db)) : run >= 2 ? String(eraYear).padStart(2, '0') : String(eraYear));
      } else if (run >= 3) {
        out = appendLit(out, run >= 4 ? WEEK_SHORT[wd] + '曜日' : WEEK_SHORT[wd]);
      } else {
        // "a" / "aa" (e.g. AM/PM parts) left to SSF; the whole AM/PM token, so that its M is not taken for a month
        const ampm = /^(am\/pm|a\/p)/i.exec(section.slice(i));
        if (ampm) j = i + ampm[0].length;
        out += section.slice(i, j);
      }
      prevTok = lc;
      i = j - 1;
      continue;
    }
    if ((db || ja411) && (lc === 'm' || lc === 'd' || (db && (lc === 'y' || lc === 'h' || lc === 's')))) {
      let j = i;
      while (j < section.length && section[j].toLowerCase() === lc) j++;
      const run = j - i;
      let lit: string | null = null;
      if (ja411 && lc === 'd' && run >= 3) lit = WEEK_SHORT[wd] + (run >= 4 ? '曜日' : ''); // [$-411]ddd / dddd
      else if (ja411 && lc === 'm' && run >= 3 && !(prevTok === 'h' || nextDateLetter(section, j) === 's')) {
        lit = run >= 5 ? String(dc.m)[0] : String(dc.m) + (run === 4 ? '月' : ''); // [$-411]mmm -> 10, mmmm -> 10月
      } else if (!db) lit = null;
      else if (lc === 'y') lit = run >= 3 ? digitWise(String(dc.y), db) : pad2(dc.y % 100, db);
      else if (lc === 'd' && run <= 2) lit = run === 2 ? pad2(dc.d, db) : positional(dc.d, db);
      else if (lc === 'h' && run <= 2) {
        const hour = h12 ? dc.H % 12 || 12 : dc.H;
        lit = run === 2 ? pad2(hour, db) : positional(hour, db);
      } else if (lc === 's' && run <= 2) lit = run === 2 ? pad2(dc.S, db) : positional(dc.S, db);
      else if (lc === 'm' && run <= 2) {
        // m after h, or before s, is the minute; otherwise the month
        const minute = prevTok === 'h' || nextDateLetter(section, j) === 's';
        const val = minute ? dc.M : dc.m;
        lit = run === 2 ? pad2(val, db) : positional(val, db);
      }
      if (lit !== null) out = appendLit(out, lit);
      else out += section.slice(i, j);
      prevTok = lc;
      i = j - 1;
      continue;
    }
    out += c;
  }
  return out;
}

function cleanForSsf(fmt: string): string {
  const f = fmt
    .replace(/\[DBNum\d\]/gi, '')
    .replace(/\[\$([^\]-]*)-[0-9A-F]+\]/gi, (_m, sym: string) => (sym ? '"' + sym + '"' : '')) // [$-411] -> '', [$¥-411] -> "¥"
    .replace(/\[(黒|青|水色|緑|紫|赤|白|黄)\]/g, ''); // colour already extracted
  // Quote bare ¥ / ￥ outside "quotes", [brackets] and \escapes (SSF throws on them).
  let out = '';
  let q = false;
  let b = false;
  for (let i = 0; i < f.length; i++) {
    const ch = f[i];
    if (ch === '\\' && !q) {
      out += ch + (f[i + 1] ?? '');
      i++;
      continue;
    }
    if (ch === '"' && !b) q = !q;
    else if (ch === '[' && !q) b = true;
    else if (ch === ']' && !q) b = false;
    out += !q && !b && (ch === '¥' || ch === '￥') ? '"' + ch + '"' : ch;
  }
  let isDate = false;
  try {
    isDate = SSF.is_date(out);
  } catch {
    isDate = false;
  }
  return isDate ? quoteDateDots(out) : out;
}

/** SSF throws on '.' separators in date codes (yyyy.m.d, ge.m.d); quote them, keep ss.0 fractions. */
function quoteDateDots(f: string): string {
  let out = '';
  let q = false;
  let b = false;
  for (let i = 0; i < f.length; i++) {
    const ch = f[i];
    if (ch === '\\' && !q) {
      out += ch + (f[i + 1] ?? '');
      i++;
      continue;
    }
    if (ch === '"' && !b) q = !q;
    else if (ch === '[' && !q) b = true;
    else if (ch === ']' && !q) b = false;
    out += !q && !b && ch === '.' && f[i + 1] !== '0' ? '"."' : ch;
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// [DBNum1..3] (数字の表記: 漢数字 / 大字 / 全角, see ./dbnum.ts) and the 午前/午後 of [$-411] AM/PM. Measured with ja-JP Excel (Range.Text):
//   [DBNum1]0 12345 -> 一二三四五   [DBNum2]0 -> 壱弐参四伍   [DBNum3]0 -> １２３４５   [DBNum1]General 123 -> 百二十三
//   [DBNum1]m"月"d"日" -> 十月一日   [DBNum3]yyyy/m/d -> ２０２６/１０/１   [$-411]AM/PM h:mm 18:00 -> 午後 6:00   (plain AM/PM stays PM)

function dbStyleOf(section: string): DbStyle | 0 {
  const m = /\[DBNum([123])\]/i.exec(section);
  return m ? (Number(m[1]) as DbStyle) : 0;
}

function isDateSection(section: string): boolean {
  const bare = bareFormat(section).replace(/general/g, '');
  return /[ymdhs]|g|e(?![+-])|aaa/.test(bare);
}

/** [DBNum1/2] dates are rewritten before SSF (substituteJa with a style) */
function dbDateStyle(section: string): 1 | 2 | undefined {
  const st = dbStyleOf(section);
  return (st === 1 || st === 2) && isDateSection(section) ? st : undefined;
}

/** The SSF text of a number / date with the numerals of the section's [DBNum] tag. */
function applyDbNum(text: string, section: string, v: number): string {
  const style = dbStyleOf(section);
  if (!style) return text;
  if (bareFormat(section).trim() === 'general') {
    // General with numerals writes the value with its 15 significant digits (no 11-character limit, no E notation)
    const full = Number.isFinite(v) ? String(Number(v.toPrecision(15))) : text;
    return generalText(/e/i.test(full) ? text : full, style);
  }
  if (style === 3) return digitWise(text, 3);
  if (isDateSection(section)) return text; // already written by substituteJa
  return digitWise(text, style);
}

/** [$-411] with month / weekday names (mmm, ddd …): the names are written the Japanese way (10, 10月, 木, 木曜日). */
function jaNames(section: string): boolean {
  return /\[\$-411\]/i.test(section) && /m{3,}|d{3,}/.test(bareFormat(section));
}

/** [$-411] (Japanese) turns AM / PM into 午前 / 午後. */
function applyJaAmPm(text: string, section: string): string {
  if (!/\[\$-411\]/i.test(section) || !/am\/pm/i.test(bareFormat(section))) return text;
  return text.replace(/AM|PM/i, (t) => (t.toUpperCase() === 'AM' ? '午前' : '午後'));
}

export interface FormattedValue {
  text: string;
  color?: string;
  /** Excel "General" alignment helper: numbers/dates right, booleans/errors centre, text left */
  kind: 'number' | 'text' | 'bool' | 'error' | 'empty';
}

const EMPTY: FormattedValue = Object.freeze({ text: '', kind: 'empty' }) as FormattedValue;
const ERROR_RE = /^#(REF!|N\/A|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|PYTHON!)$/;

export function formatValue(v: unknown, fmtIn?: string | number | null): FormattedValue {
  if (v === null || v === undefined || v === '') return EMPTY;
  if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE', kind: 'bool' };
  if (v instanceof Date) v = dateToSerial(v);
  const fmt = normalizeNumFmt(fmtIn);
  if (typeof v === 'string') {
    if (v.startsWith('#') && ERROR_RE.test(v)) return { text: v, kind: 'error' };
    if (fmt === 'General' || fmt === '@') return { text: v, kind: 'text' };
  }
  if (typeof v === 'number' && !Number.isFinite(v)) return { text: '#NUM!', kind: 'error' };
  if (typeof v !== 'number' && typeof v !== 'string') return { text: String(v), kind: 'text' };
  try {
    const sections = splitSections(fmt);
    const section = pickSection(sections, v);
    const color = sectionColor(section);
    let f = fmt;
    if (typeof v === 'number' && (needsJaTokens(fmt) || sections.some((x) => dbDateStyle(x) || jaNames(x)))) {
      f = sections.map((x) => substituteJa(x, v as number, dbDateStyle(x))).join(';');
    }
    let text = SSF.format(cleanForSsf(f), v);
    if (typeof v === 'number') text = applyJaAmPm(applyDbNum(text, section, v), section);
    return color ? { text, color, kind: typeof v === 'number' ? 'number' : 'text' } : { text, kind: typeof v === 'number' ? 'number' : 'text' };
  } catch {
    return { text: String(v), kind: typeof v === 'number' ? 'number' : 'text' };
  }
}

/** JS Date (local) -> Excel 1900 serial. */
export function dateToSerial(d: Date): number {
  const utc = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds(), d.getMilliseconds());
  return utc / 86400000 + 25569;
}

// ---------------------------------------------------------------------------------------------
// Cached formatting (SSF costs 2-10 µs per call; a visible screen is ~800 cells per frame)

const FMT_CACHE_MAX = 200_000;
const fmtCache = new Map<string, FormattedValue>();

/** formatValue with a bounded memo (key = format + type + value). Returned objects are shared: do not mutate. */
export function formatCached(v: unknown, fmt?: string | number | null): FormattedValue {
  if (v === null || v === undefined || v === '') return EMPTY;
  if (typeof v === 'string' && (fmt === undefined || fmt === null || fmt === '' || fmt === 'General' || fmt === 0 || fmt === '@' || fmt === 49)) {
    return v.startsWith('#') && ERROR_RE.test(v) ? { text: v, kind: 'error' } : { text: v, kind: 'text' };
  }
  const k = `${fmt ?? ''}\u0001${typeof v}\u0001${String(v)}`;
  let r = fmtCache.get(k);
  if (!r) {
    r = formatValue(v, fmt);
    if (fmtCache.size >= FMT_CACHE_MAX) fmtCache.clear();
    fmtCache.set(k, r);
  }
  return r;
}

const dateFmtCache = new Map<string, boolean>();

/** true when the format code displays a date and/or a time (built-in ids resolved with the ja-JP table). */
export function isDateFormat(fmtIn: string | number | undefined | null): boolean {
  const fmt = normalizeNumFmt(fmtIn);
  if (fmt === 'General' || fmt === '@') return false;
  const cached = dateFmtCache.get(fmt);
  if (cached !== undefined) return cached;
  let r = false;
  try {
    // 和暦-only codes ("ggge") have no y/m/d that SSF recognises -> check the bare tokens as well
    r = !!SSF.is_date(cleanForSsf(fmt)) || /(^|[^a-z])(g{1,3}|e{1,2})([^+-]|$)/.test(bareFormat(fmt).replace(/general/g, ''));
  } catch {
    r = false;
  }
  if (dateFmtCache.size > 5000) dateFmtCache.clear();
  dateFmtCache.set(fmt, r);
  return r;
}

/** Minimal cell shape (structurally compatible with the contract `Cell`). */
export interface ValueCell {
  t?: string;
  v: string | number | boolean | null | undefined;
}

/**
 * Text Excel DISPLAYS for a cell value with the given number format (no formula handling).
 * Honours the contract cell types: 'e' errors, 'b' booleans, 's' text (never treated as an error code),
 * 'd' dates (a date cell without a date format still shows a date), 'z' empty.
 */
export function displayValue(cell: ValueCell | undefined | null, fmt?: string | number | null): FormattedValue {
  if (!cell) return EMPTY;
  const v = cell.v;
  if (v === null || v === undefined || v === '') return EMPTY;
  switch (cell.t) {
    case 'z':
      return EMPTY;
    case 'e':
      return { text: String(v), kind: 'error' };
    case 'b':
      return { text: v === true || v === 1 || v === 'TRUE' ? 'TRUE' : 'FALSE', kind: 'bool' };
    case 's':
      if (typeof v === 'string') {
        const f = normalizeNumFmt(fmt);
        if (f === 'General' || f === '@') return { text: v, kind: 'text' };
      }
      break;
    case 'd':
      if (typeof v === 'number' && !isDateFormat(fmt)) {
        return formatCached(v, Math.abs(v - Math.floor(v)) > 1e-9 ? 'yyyy/m/d h:mm' : 'yyyy/m/d');
      }
      break;
    default:
      break;
  }
  return formatCached(v, fmt);
}

// ---------------------------------------------------------------------------------------------
// Edit text (what Excel shows in the formula bar / in-cell editor for a constant)

/** Number as Excel's formula bar shows it: up to 15 significant digits, no separators, 1E+21 style exponent. */
export function generalEditNumber(n: number): string {
  if (!Number.isFinite(n)) return '#NUM!';
  if (n === 0) return '0';
  let s = String(Number(n.toPrecision(15)));
  const m = /^(-?[\d.]+)e([+-])(\d+)$/.exec(s);
  if (m) s = `${m[1]}E${m[2]}${m[3].padStart(2, '0')}`;
  return s;
}

function hasPercent(fmt: string): boolean {
  return bareFormat(fmt).includes('%');
}

function isTimeOnlyFormat(fmt: string): boolean {
  const bare = bareFormat(fmt);
  return /[hs]/.test(bare) && !/[yd]|g|e(?![+-])|aaa/.test(bare);
}

/**
 * The text Excel puts into the formula bar / in-cell editor for a CONSTANT cell (formulas are handled by
 * the caller): dates as yyyy/m/d [h:mm:ss], percentages as 12.5%, numbers with up to 15 significant digits,
 * booleans TRUE/FALSE, text as is.
 */
export function editText(cell: ValueCell | undefined | null, fmtIn?: string | number | null): string {
  if (!cell) return '';
  const v = cell.v;
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean' || cell.t === 'b') return v === true || v === 1 || v === 'TRUE' ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') return v;
  if (typeof v !== 'number') return String(v);
  const fmt = normalizeNumFmt(fmtIn);
  if (cell.t === 'd' || isDateFormat(fmt)) {
    if (v < 0 || v >= 2958466) return generalEditNumber(v); // outside Excel's date range (shown as #### in Excel)
    try {
      const days = Math.floor(v);
      // round to whole seconds like Excel's formula bar
      const secs = Math.round((v - days) * 86400);
      const hasTime = secs !== 0 && secs !== 86400;
      if (days === 0 && isTimeOnlyFormat(fmt)) return SSF.format('h:mm:ss', v);
      const date = SSF.format('yyyy/m/d', secs === 86400 ? days + 1 : days);
      return hasTime ? `${date} ${SSF.format('h:mm:ss', v)}` : date;
    } catch {
      return generalEditNumber(v);
    }
  }
  if (hasPercent(fmt)) return generalEditNumber(v * 100) + '%';
  return generalEditNumber(v);
}
