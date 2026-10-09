/**
 * Per-type value normalisation for join keys (Excel 統合ツール, ported from research R6).
 *
 * Every normaliser returns '' for "blank / unusable" so callers can count blanks uniformly.
 * All functions are pure except for the module-level NormalizeOptions (see withNormalizeOptions), which the
 * keys facade sets for the duration of one synchronous analyze / join call.
 */
import itaijiJson from './dict/itaiji-name.json';
import type { NormalizeOptions } from '../contracts';

/** A cell VALUE as the keys module sees it (cached formula results; dates as UTC Date objects). */
export type Cell = string | number | boolean | Date | null | undefined;

const ITAIJI: Record<string, string> = (itaijiJson as unknown as { map: Record<string, string> }).map;

// ---------------------------------------------------------------------------
// options (異体字 / 先頭ゼロ / 敬称 / 旧姓 / 大学名)
// ---------------------------------------------------------------------------

export const DEFAULT_NORMALIZE_OPTIONS: NormalizeOptions = Object.freeze({
  itaiji: true,
  leadingZeros: true,
  honorifics: true,
  maidenName: true,
  universityAbbrev: true,
});

let OPTS: NormalizeOptions = DEFAULT_NORMALIZE_OPTIONS;

export function resolveNormalizeOptions(o?: Partial<NormalizeOptions> | null): NormalizeOptions {
  const d = DEFAULT_NORMALIZE_OPTIONS;
  const pick = (k: keyof NormalizeOptions) => (o && typeof o[k] === 'boolean' ? (o[k] as boolean) : d[k]);
  return {
    itaiji: pick('itaiji'),
    leadingZeros: pick('leadingZeros'),
    honorifics: pick('honorifics'),
    maidenName: pick('maidenName'),
    universityAbbrev: pick('universityAbbrev'),
  };
}

export function getNormalizeOptions(): NormalizeOptions {
  return OPTS;
}

/** Runs fn with the given options active (restored afterwards, also on throw). */
export function withNormalizeOptions<T>(o: NormalizeOptions, fn: () => T): T {
  const prev = OPTS;
  OPTS = o;
  try {
    return fn();
  } finally {
    OPTS = prev;
  }
}

// ---------------------------------------------------------------------------
// basic text
// ---------------------------------------------------------------------------

/** Variation selectors (SVS U+FE00-FE0F, IVS U+E0100-E01EF): 葛+IVS → 葛. */
const VS_RE = /[︀-️]|\uDB40[\uDD00-\uDDEF]/g;
/** Zero-width / BOM / bidi / soft hyphen that survive copy-paste from PDFs and web forms. */
const INVISIBLE_RE = /[​-‏‪-‮⁠-⁤﻿­]/g;
/** Every whitespace incl. U+3000 (NFKC already maps U+3000 → U+0020, kept for safety). */
const WS_RE = /[\s　]+/g;
/** Hyphen / dash / minus variants → '-' (incl. the katakana long vowel, which people type inside IDs). */
const HYPHENS_RE = /[‐-―−⁃﹘﹣－⸺⸻ー─━ｰ]/g;
/** Long-vowel look-alikes inside KANA strings → 'ー'. */
const LONG_VOWEL_RE = /[‐-―−－─━ｰ~〜～-]/g;
const PRINTABLE_ASCII_RE = /^[\x20-\x7E]*$/;

/** Excel error literals: never keys. */
const EXCEL_ERROR_RE = /^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|GETTING_DATA)$/i;
/** Placeholder text that means "blank". */
const PLACEHOLDER_RE = /^(-+|―+|ー+|なし|無し|不明|未定|未記入|未入力|該当なし|n\/?a|null|none|nil|undefined|\?+|？+|\.+|・+)$/i;

export function cellText(v: Cell): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number') {
    if (!Number.isFinite(v)) return '';
    // integers: no exponent / no '.0' (8000000001, never 8.000000001e9)
    if (Number.isInteger(v)) return Math.abs(v) < 1e21 ? BigInt(v).toString() : String(v);
    return String(v);
  }
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v instanceof Date) return isoDate(v);
  return String(v);
}

/**
 * Characters NFKC leaves alone and that never need the variation-selector / invisible-character pass: printable ASCII,
 * hiragana, full-width katakana (incl. ー), CJK unified ideographs, 々. A text made of these only (most Japanese names) is
 * returned after a trim.
 */
const NFKC_SAFE_RE = /^[\x20-\x7E\u3041-\u3096\u30A1-\u30F6\u30FC\u3005\u4E00-\u9FFF]*$/;

/** the same plus U+3000 (ideographic space: NFKC turns it into U+0020 and nothing else happens) */
const NFKC_SAFE_SP_RE = /^[\x20-\x7E\u3000\u3041-\u3096\u30A1-\u30F6\u30FC\u3005\u4E00-\u9FFF]*$/;
const IDEOGRAPHIC_SPACE_RE = /\u3000/g;

/** NFKC + drop variation selectors / invisible chars + trim. The base of every normaliser. */
export function nfkc(s: string): string {
  if (NFKC_SAFE_RE.test(s)) return s.trim(); // NFKC is the identity on these (printable ASCII, kana, kanji)
  if (NFKC_SAFE_SP_RE.test(s)) return s.replace(IDEOGRAPHIC_SPACE_RE, ' ').trim();
  return s.replace(VS_RE, '').replace(INVISIBLE_RE, '').normalize('NFKC').trim();
}

/** first characters a blank-like text can start with (placeholders, error literals, whitespace, punctuation) */
// ordinary = A-Z/a-z except n/u, digits, kana except な (U+306A), kanji except 不 (4E0D) 未 (672A) 無 (7121) 該 (8A72)
const BLANK_START_RE = /[^A-MO-TV-Za-mo-tv-z0-9ぁ-どに-ヺ一-丌与-朩末-焠焢-話詳-鿿]/;

export function isBlankLike(v: Cell): boolean {
  if (v === null || v === undefined) return true;
  if (typeof v === 'number') return !Number.isFinite(v);
  if (typeof v !== 'string') return false;
  // fast path: a text starting with an ordinary letter / digit / kana / kanji can only be a placeholder when it starts
  // like one (n/a, null, none, nil, undefined, なし, 無し, 不明, 未定, 未記入, 未入力, 該当なし)
  if (v.length && !BLANK_START_RE.test(v[0])) return false;
  const t = nfkc(v).replace(WS_RE, '');
  return t === '' || EXCEL_ERROR_RE.test(t) || PLACEHOLDER_RE.test(t);
}

/** Generic comparable text used for value-overlap discovery of unknown columns. */
const ASCII_WS_RE = /\s+/g;

/**
 * Longest text that can be a key / a name / an address. Free text (志望動機 pasted into a column, a 100,000-character cell)
 * is never one, and the bracket / e-mail regexes below are quadratic on it.
 */
export const MAX_KEY_TEXT = 300;

export function normGeneric(v: Cell): string {
  if (isBlankLike(v)) return '';
  if (typeof v === 'string' && v.length > MAX_KEY_TEXT) return '';
  if (typeof v === 'number') return cellText(v); // numbers never carry leading zeros
  if (typeof v === 'string' && PRINTABLE_ASCII_RE.test(v)) {
    // fast path: NFKC / 異体字 / exotic hyphens do nothing to printable ASCII
    let a = v.replace(ASCII_WS_RE, '').toUpperCase();
    if (/^\d+\.0+$/.test(a)) a = a.replace(/\.0+$/, '');
    if (OPTS.leadingZeros && /^\d+$/.test(a)) a = a.replace(/^0+(?=\d)/, '');
    return a;
  }
  let s = nfkc(cellText(v)).replace(WS_RE, '').replace(HYPHENS_RE, '-').toUpperCase();
  if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, '');
  if (OPTS.leadingZeros && /^\d+$/.test(s)) s = s.replace(/^0+(?=\d)/, ''); // 00123 ≡ 123 for discovery
  return OPTS.itaiji ? foldItaiji(s) : s;
}

// ---------------------------------------------------------------------------
// IDs
// ---------------------------------------------------------------------------

const THOUSANDS_RE = /^\d{1,3}(,\d{3})+$/;

export interface IdKey {
  /** exact comparable form: NFKC, no spaces, unified hyphens, upper-case */
  exact: string;
  /** all-digit IDs: without leading zeros ('00123' → '123') when the 先頭ゼロ option is on; '' when not all-digit */
  digits: string;
  /** true when the source cell was an Excel number (leading zeros may have been lost) */
  fromNumber: boolean;
  /** 8.00E+09 typed/saved as text: the real digits are lost */
  lossyExponent: boolean;
}

export function normId(v: Cell, displayText?: string): IdKey {
  const empty: IdKey = { exact: '', digits: '', fromNumber: false, lossyExponent: false };
  if (isBlankLike(v) || typeof v === 'boolean' || v instanceof Date) return empty;
  if (typeof v === 'string' && v.length > MAX_KEY_TEXT) return empty;
  let fromNumber = typeof v === 'number';
  // number cell formatted '00000' → Excel SHOWS 00123: the displayed text is the ID
  const shown = displayText ? displayText.trim() : '';
  let s = fromNumber && /^0\d+$/.test(shown) ? shown : cellText(v);
  s = nfkc(s).replace(WS_RE, '').replace(HYPHENS_RE, '-').toUpperCase();
  s = s.replace(/^'+/, ''); // literal apostrophe prefix pasted from CSV
  if (THOUSANDS_RE.test(s)) s = s.replace(/,/g, ''); // 20,230,012 (a number pasted with its thousands separators)
  if (/^\d+\.0+$/.test(s)) {
    s = s.replace(/\.0+$/, '');
    fromNumber = true;
  }
  let lossyExponent = false;
  const m = /^(\d+)(?:\.(\d*))?E\+?(\d+)$/.exec(s);
  if (m) {
    const intDigits = m[1].replace(/^0+/, '');
    const written = (intDigits + (m[2] ?? '')).length;
    lossyExponent = written < intDigits.length + Number(m[3]);
    const n = Number(s);
    if (Number.isSafeInteger(n)) s = String(n);
  }
  if (s === '' || EXCEL_ERROR_RE.test(s)) return empty;
  const digits = /^\d+$/.test(s) ? (OPTS.leadingZeros ? s.replace(/^0+(?=\d)/, '') : s) : '';
  return { exact: s, digits, fromNumber, lossyExponent };
}

/**
 * The matching key of an ID cell ('' = blank).
 * 2.02301E+11 typed as TEXT (a CSV round trip) has lost its digits: every 2023010xxxxx collapses onto it and it would read
 * back as 202301000000 — a real ID of somebody else. Such a cell has NO usable key (the lossyId check item reports it).
 */
export function idKeyOf(v: Cell, displayText?: string): string {
  const k = normId(v, displayText);
  if (k.lossyExponent) return '';
  return k.digits || k.exact.replace(/[-_]/g, '');
}

// ---------------------------------------------------------------------------
// names
// ---------------------------------------------------------------------------

/** every character of the 異体字 table (most names contain none of them: one scan instead of a per-character rebuild) */
const ITAIJI_RE = new RegExp(`[${Object.keys(ITAIJI).map((k) => k.replace(/[\]\\^-]/g, '\\$&')).join('')}]`, 'u');

export function foldItaiji(s: string): string {
  if (!ITAIJI_RE.test(s)) return s;
  let out = '';
  for (const ch of s) out += ITAIJI[ch] ?? ch;
  return out;
}

/** 々 / 〻 (iteration marks) → repeat previous character: 佐々木 ≡ 佐佐木. */
function expandIteration(s: string): string {
  if (!s.includes('々') && !s.includes('〻')) return s;
  const chars = [...s];
  for (let i = 1; i < chars.length; i++) if (chars[i] === '々' || chars[i] === '〻') chars[i] = chars[i - 1];
  return chars.join('');
}

/** Trailing honorifics / notes that are not part of the name. */
const HONORIFIC_RE = /(様|さま|サマ|殿|さん|くん|君|ちゃん|氏)$/;
const PAREN_NOTE_RE = /[(（[［【].*?[)）\]］】]/g;
const NAME_PUNCT_RE = /[・･·•=＝゠.,，、。'’`"]/g;
const KATAKANA_DASH_RE = /(?<=[゠-ヿ])[‐-―−－-]/g;
const DASHISH_RE = /[‐-―−－-]/;
const PAREN_OPEN_RE = /[(（[［【]/;
const LATIN_NAME_START_RE = /^\p{Script=Latin}/u;
const LATIN_WORDS_RE = /^\p{Script=Latin}[\p{Script=Latin}'-]*$/u;

export interface NameKey {
  /** strict: NFKC, no spaces/punctuation, iteration marks expanded, Latin lower-cased (異体字 NOT folded) */
  strict: string;
  /** + 異体字 folded (髙→高, 﨑→崎, 邊/邉→辺, 齋→斎 ...) when the option is on — DEFAULT matching key */
  folded: string;
  /** token-sorted Latin form ('taro yamada' ≡ 'yamada taro'), '' for non-Latin */
  latinSorted: string;
  /** parenthetical note removed from the name, e.g. '(旧姓:佐藤)' */
  note: string;
  /**
   * the name under the OTHER surname a note gives: 山田 花子（旧姓：佐藤） → 佐藤花子 (folded). Only for suggestions — a maiden
   * name is no key. Absent when there is no such note.
   */
  alt?: string;
}

/** 旧姓：佐藤 / 旧 佐藤 / 元 佐藤 / maiden: Sato → the surname inside a note, '' when it does not look like one */
const MAIDEN_NOTE_RE = /^(?:旧姓|旧|元|結婚前|maiden(?:\s*name)?)[:：\s　]*([\p{Script=Han}々ぁ-ヶー]{1,6}|[A-Za-z][A-Za-z'-]{1,20})$/iu;
function maidenSurname(notes: RegExpMatchArray | null): string {
  if (!notes) return '';
  for (const n of notes) {
    const inner = n.slice(1, -1).replace(/[\s　]+/g, ' ').trim();
    const m = MAIDEN_NOTE_RE.exec(inner);
    if (m) return m[1];
  }
  return '';
}

/** a line that is nothing but kana (and spaces / dots): the フリガナ printed under a name in the same cell */
const KANA_LINE_RE = /^[\p{Script=Hiragana}\p{Script=Katakana}ー・･\s　.]+$/u;

/** the lines of a multi-line cell, NFKC-normalised, without blank lines; null for a single-line text */
function textLines(v: string): string[] | null {
  if (!/[\r\n]/.test(v)) return null;
  const lines = v.split(/[\r\n]+/).map((l) => nfkc(l)).filter(Boolean);
  return lines.length >= 2 ? lines : null;
}

/** 山田 太郎⏎ヤマダ タロウ → 山田 太郎 (the kana line is the reading, not part of the name) */
function withoutReadingLines(v: string): string {
  const lines = textLines(v);
  if (!lines) return v;
  const keep = lines.filter((l) => !KANA_LINE_RE.test(l));
  return keep.length >= 1 && keep.length < lines.length ? keep.join(' ') : v;
}

export function normName(v: Cell): NameKey {
  const empty: NameKey = { strict: '', folded: '', latinSorted: '', note: '' };
  if (isBlankLike(v) || typeof v !== 'string' || v.length > MAX_KEY_TEXT) return empty;
  let s = nfkc(withoutReadingLines(v));
  let notes: RegExpMatchArray | null = null;
  if (OPTS.maidenName && PAREN_OPEN_RE.test(s)) {
    notes = s.match(PAREN_NOTE_RE);
    s = s.replace(PAREN_NOTE_RE, ' ').trim();
  }
  const spaced = s; // words still separated (the Latin path below needs them)
  const maiden = maidenSurname(notes);
  s = s.replace(WS_RE, '');
  if (OPTS.honorifics && s.length > 2) s = s.replace(HONORIFIC_RE, '');
  s = s.replace(NAME_PUNCT_RE, '');
  s = expandIteration(s).toLowerCase();
  // katakana in a name: unify long-vowel look-alikes; hiragana/katakana are NOT folded in the name key
  if (DASHISH_RE.test(s)) s = s.replace(KATAKANA_DASH_RE, 'ー');
  if (!s) return empty;
  let latinSorted = '';
  if (LATIN_NAME_START_RE.test(s)) {
    // Latin: the words in any order ('Taro Yamada' ≡ 'YAMADA Taro' ≡ 'Yamada, Taro'); a trailing 様 / さん is not a word
    let words = spaced.split(WS_RE).flatMap((t) => t.split(/[,，]+/)).filter(Boolean);
    if (OPTS.honorifics && words.length > 1 && HONORIFIC_RE.test(words[words.length - 1])) words = words.slice(0, -1);
    if (LATIN_WORDS_RE.test(words.join('').toLowerCase().replace(NAME_PUNCT_RE, ''))) {
      latinSorted = words.map((t) => t.toLowerCase().replace(NAME_PUNCT_RE, '')).sort().join(' ');
    }
  }
  const key: NameKey = {
    strict: s,
    folded: OPTS.itaiji ? foldItaiji(s) : s,
    latinSorted,
    note: notes ? notes.join(' ') : '',
  };
  if (maiden) {
    // 山田 花子（旧姓：佐藤）: the given name is the last word; the other surname replaces the first one
    const words = spaced.split(WS_RE).filter(Boolean);
    if (words.length >= 2) {
      const alt = (maiden + words[words.length - 1]).replace(NAME_PUNCT_RE, '').toLowerCase();
      key.alt = OPTS.itaiji ? foldItaiji(alt) : alt;
    }
  }
  return key;
}

// ---------------------------------------------------------------------------
// kana
// ---------------------------------------------------------------------------

/** hiragana → katakana (ぁ-ゖ, ゝゞ → ァ-ヶ, ヽヾ). Half-width kana are handled by NFKC first. */
export function hiraToKata(s: string): string {
  return s.replace(/[ぁ-ゖゝゞ]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 0x60));
}

const SMALL_TO_LARGE: Record<string, string> = {
  'ァ': 'ア', 'ィ': 'イ', 'ゥ': 'ウ', 'ェ': 'エ', 'ォ': 'オ', 'ッ': 'ツ', 'ャ': 'ヤ', 'ュ': 'ユ', 'ョ': 'ヨ', 'ヮ': 'ワ', 'ヵ': 'カ', 'ヶ': 'ケ',
};
const LOOSE_KANA: Record<string, string> = { 'ヂ': 'ジ', 'ヅ': 'ズ', 'ヲ': 'オ', 'ヰ': 'イ', 'ヱ': 'エ', 'ヴ': 'ブ' };
const KANA_CHAR_RE = /[ァ-ヺー]/;
const KANA_HONORIFIC_RE = /(様|サマ|サン|クン|チャン|殿)$/;

export interface KanaKey {
  /** NFKC (ｶﾞ→ガ), hiragana→katakana, no spaces/・, long-vowel look-alikes → ー */
  strict: string;
  /** + small→large kana, ヂ→ジ ヅ→ズ ヴ→ブ ヲ→オ, ー removed — for fuzzy SUGGESTIONS only */
  loose: string;
  /** share of kana characters in the cell (0..1) */
  kanaRatio: number;
}

export function kanaLoose(strict: string): string {
  let out = '';
  for (const c of strict) {
    const big = SMALL_TO_LARGE[c] ?? c;
    const l = LOOSE_KANA[big] ?? big;
    if (l !== 'ー') out += l;
  }
  return out;
}

/** The strict kana key only (normKana(v).strict without the loose form / ratio). */
export function kanaStrict(v: Cell): string {
  if (isBlankLike(v) || typeof v !== 'string' || v.length > MAX_KEY_TEXT) return '';
  let s = nfkc(v);
  if (OPTS.maidenName) s = s.replace(PAREN_NOTE_RE, ' ');
  s = hiraToKata(s).replace(WS_RE, '').replace(NAME_PUNCT_RE, '').replace(LONG_VOWEL_RE, 'ー');
  if (OPTS.honorifics) s = s.replace(KANA_HONORIFIC_RE, (m) => (s.length > m.length + 1 ? '' : m));
  return s;
}

export function normKana(v: Cell): KanaKey {
  const empty: KanaKey = { strict: '', loose: '', kanaRatio: 0 };
  const s = kanaStrict(v);
  if (!s) return empty;
  const chars = [...s];
  let kana = 0;
  for (const c of chars) if (KANA_CHAR_RE.test(c)) kana++;
  return { strict: s, loose: kanaLoose(s), kanaRatio: chars.length ? kana / chars.length : 0 };
}

// ---------------------------------------------------------------------------
// email / phone
// ---------------------------------------------------------------------------

export const EMAIL_RE = /^[A-Za-z0-9._%+\-']+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$/;

/** an address inside free text: 山田 太郎 <taro@example.com> / (taro@example.com) / 'taro@example.com'、 */
const EMAIL_IN_TEXT_RE = /[a-z0-9._%+\-']+@[a-z0-9-]+(?:\.[a-z0-9-]+)*\.[a-z]{2,}/;

export function normEmail(v: Cell): string {
  if (isBlankLike(v) || typeof v !== 'string' || v.length > MAX_KEY_TEXT) return '';
  const text = nfkc(v).toLowerCase().replace(/＠|\s*[(\[]at[)\]]\s*/g, '@');
  let s = text.replace(WS_RE, '').replace(/^mailto:/i, '');
  s = s.replace(/[。]/g, '.').replace(/\.+$/, '');
  if (EMAIL_RE.test(s)) return s;
  // a display name / brackets / quotes / a trailing 、 around the address: take the (first) address inside
  const m = EMAIL_IN_TEXT_RE.exec(text.replace(/mailto:/g, ' '));
  if (!m) return '';
  const a = m[0].replace(/^'+/, '').replace(/\.+$/, '');
  return EMAIL_RE.test(a) ? a : '';
}

/** Gmail-style canonical form (dots / +tag ignored) — for SUGGESTIONS only, never auto-merge. */
export function emailLoose(s: string): string {
  const [local, domain] = s.split('@');
  if (!domain) return s;
  const l = local.split('+')[0];
  return (/^(gmail|googlemail)\.com$/.test(domain) ? l.replace(/\./g, '') : l) + '@' + domain.replace('googlemail.com', 'gmail.com');
}

export interface PhoneKey {
  /** 10–11 digit domestic form starting with 0 ('' when not a JP phone number) */
  digits: string;
  /** leading zero was restored (Excel number 9012345678 → 09012345678) */
  restoredZero: boolean;
}

export function normPhone(v: Cell): PhoneKey {
  const empty: PhoneKey = { digits: '', restoredZero: false };
  if (isBlankLike(v) || typeof v === 'boolean' || v instanceof Date) return empty;
  if (typeof v === 'string' && v.length > MAX_KEY_TEXT) return empty;
  // two numbers in one cell (090-1234-5678 / 03-1111-2222): the first valid one is the key
  if (typeof v === 'string' && /[\/／、,，;；\r\n]/.test(v)) {
    for (const part of v.split(/[\/／、,，;；\r\n]+/)) {
      if (part.trim() === v.trim()) break;
      const r = normPhone(part);
      if (r.digits) return r;
    }
  }
  let s = nfkc(cellText(v)).replace(/^(tel|phone|電話|携帯)[:：]?/i, '');
  s = s.replace(/(内線|ext\.?|x)\s*\d+$/i, ''); // extension
  const plus = /^\s*\+/.test(s);
  let d = s.replace(/[^\d]/g, '');
  if ((plus || /^\(?\+?81/.test(s)) && d.startsWith('81')) d = '0' + d.slice(2).replace(/^0/, '');
  else if (d.startsWith('0081')) d = '0' + d.slice(4);
  let restoredZero = false;
  if (/^[1-9]\d{8,9}$/.test(d) && (typeof v === 'number' || /^[1-9]/.test(s.trim()))) {
    d = '0' + d;
    restoredZero = true;
  }
  return /^0\d{9,10}$/.test(d) ? { digits: d, restoredZero } : empty;
}

// ---------------------------------------------------------------------------
// organisation (university / faculty)
// ---------------------------------------------------------------------------

const ORG_PREFIX_RE = /^(国立大学法人|公立大学法人|学校法人[^\s　]*?(?=[^\s　]*(大学|学院|学園))|独立行政法人|私立|国立|公立|市立|県立|都立|府立|道立)/;
const UNI_ALIAS: Record<string, string> = {
  '東大': '東京大学', '京大': '京都大学', '阪大': '大阪大学', '東北大': '東北大学', '名大': '名古屋大学', '九大': '九州大学', '北大': '北海道大学',
  '一橋': '一橋大学', '東工大': '東京工業大学', '科学大': '東京科学大学', '早大': '早稲田大学', '早稲田': '早稲田大学', '慶大': '慶応義塾大学', '慶応': '慶応義塾大学',
  '慶応大学': '慶応義塾大学', '上智': '上智大学', '明大': '明治大学', '立教': '立教大学', '中大': '中央大学', '法大': '法政大学', '同志社': '同志社大学', '立命館': '立命館大学', '関学': '関西学院大学', '関大': '関西大学',
  'sfc': '慶応義塾大学', '筑波': '筑波大学', '神大': '神戸大学', '横国': '横浜国立大学', '東京理科': '東京理科大学', '理科大': '東京理科大学', 'icu': '国際基督教大学',
};

export function normOrg(v: Cell): string {
  if (isBlankLike(v) || typeof v !== 'string' || v.length > MAX_KEY_TEXT) return '';
  let s = nfkc(v).replace(PAREN_NOTE_RE, '').replace(WS_RE, '').toLowerCase();
  if (OPTS.itaiji) s = foldItaiji(s); // 慶應 ≡ 慶応
  if (OPTS.universityAbbrev) {
    s = s.replace(ORG_PREFIX_RE, '');
    s = s.replace(/(大学院|大学大学院)$/, '大学').replace(/大学大学院/, '大学');
    s = s.replace(/(キャンパス|校舎)$/, '');
    s = UNI_ALIAS[s] ?? s;
  }
  return s;
}

// ---------------------------------------------------------------------------
// dates (Excel serials, Date objects, 西暦/和暦 strings)
// ---------------------------------------------------------------------------

export function isoDate(d: Date): string {
  if (isNaN(d.getTime())) return '';
  // Date cells are UTC-midnight Dates (see the SourceBook adapter)
  const y = d.getUTCFullYear(), m = d.getUTCMonth() + 1, day = d.getUTCDate();
  return `${y}-${String(m).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

/** Excel 1900 date system serial → UTC Date (handles the fake 1900-02-29; keeps the time fraction). */
export function excelSerialToDate(n: number): Date | null {
  if (!Number.isFinite(n) || n < 0 || n > 2958465.99999) return null;
  const whole = Math.floor(n);
  const days = whole > 59 ? whole - 1 : whole; // serial 60 = fake 1900-02-29
  const ms = Date.UTC(1899, 11, 31) + days * 86400000 + Math.round((n - whole) * 86400000);
  return new Date(ms);
}

/** Excel 1900 date system serial → ISO date (handles the fake 1900-02-29). */
export function excelSerialToIso(n: number): string {
  if (!Number.isFinite(n) || n < 1 || n > 2958465) return '';
  const d = excelSerialToDate(Math.floor(n));
  return d ? isoDate(d) : '';
}

const ERA: Record<string, number> = { '令和': 2018, 'r': 2018, '平成': 1988, 'h': 1988, '昭和': 1925, 's': 1925, '大正': 1911, 't': 1911, '明治': 1867, 'm': 1867 };

function ymd(y: number, m: number, d: number): string {
  if (!(y >= 1868 && y <= 2100 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return '';
  const dt = new Date(Date.UTC(y, m - 1, d));
  if (dt.getUTCMonth() !== m - 1) return ''; // 2/30 etc.
  return isoDate(dt);
}

export interface DateKey {
  iso: string;
  source: 'date' | 'serial' | 'text' | 'wareki' | '';
}

/**
 * @param serialOk  treat plain numbers as Excel serials (only when the column's number format is a date format,
 *                  or the header says 生年月日 and numbers fall in a plausible range)
 */
export function normDate(v: Cell, serialOk = false): DateKey {
  if (isBlankLike(v) || typeof v === 'boolean') return { iso: '', source: '' };
  if (v instanceof Date) {
    const iso = isoDate(v);
    return { iso, source: iso ? 'date' : '' };
  }
  if (typeof v === 'number') {
    if (serialOk && v >= 1 && v < 80000) return { iso: excelSerialToIso(v), source: 'serial' };
    if (Number.isInteger(v) && v >= 18680101 && v <= 21001231) {
      // 20010401
      const s = String(v);
      const iso = ymd(+s.slice(0, 4), +s.slice(4, 6), +s.slice(6, 8));
      return { iso, source: iso ? 'text' : '' };
    }
    return { iso: '', source: '' };
  }
  if (typeof v !== 'string') return { iso: '', source: '' };
  let s = nfkc(v).replace(WS_RE, '').toLowerCase();
  if (s.length > 24) return { iso: '', source: '' };
  // 和暦: 平成13年4月1日 / H13.4.1 / h13/04/01 / 令和元年5月1日 / R1-5-1
  let m = /^(令和|平成|昭和|大正|明治|[rhstm])(\d{1,2}|元)[年./-](\d{1,2})[月./-](\d{1,2})日?(生|生まれ)?$/.exec(s);
  if (m) {
    const y = ERA[m[1]] + (m[2] === '元' ? 1 : +m[2]);
    const iso = ymd(y, +m[3], +m[4]);
    return { iso, source: iso ? 'wareki' : '' };
  }
  s = s.replace(/\(.+?\)$/, ''); // (月) weekday
  m = /^(\d{4})[年./-](\d{1,2})[月./-](\d{1,2})日?(生|生まれ)?$/.exec(s) ?? /^(\d{4})(\d{2})(\d{2})$/.exec(s);
  if (m) {
    const iso = ymd(+m[1], +m[2], +m[3]);
    return { iso, source: iso ? 'text' : '' };
  }
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s); // US m/d/yyyy (CSV exported from English Excel)
  if (m) {
    const iso = ymd(+m[3], +m[1], +m[2]);
    return { iso, source: iso ? 'text' : '' };
  }
  return { iso: '', source: '' };
}

// ---------------------------------------------------------------------------
// value-shape predicates (for column typing)
// ---------------------------------------------------------------------------

const HAN = '\\p{Script=Han}々〆ヶ';
const KANA = '\\p{Script=Hiragana}\\p{Script=Katakana}ー';
/** 山田太郎 / 山田 太郎 / 山田　花子 / 佐藤 さくら / 髙橋 一郎 (2–12 chars, ≥1 Han, optional single space) */
export const KANJI_NAME_RE = new RegExp(`^[${HAN}]{1,6}[ ]?[${HAN}${KANA}]{1,6}$`, 'u');
export const KANA_NAME_RE = /^[ァ-ヺー]{2,20}$/; // after normKana().strict
export const LATIN_NAME_RE = /^\p{Script=Latin}[\p{Script=Latin}'.-]*(,? +\p{Script=Latin}[\p{Script=Latin}'.-]*){1,3}$/u;
export const UNIV_RE = /(大学|大学院|短期大学|短大|高等専門学校|高専|専門学校|university|college|institute|univ\.?$)/i;
export const FACULTY_RE = /(学部|学科|研究科|専攻|学群|学類|学環|学域|学府|課程|コース|faculty|department|school of)/i;
const HAN_RE = /\p{Script=Han}/u;

export function isKanjiNameShape(v: Cell): boolean {
  if (typeof v !== 'string') return false;
  const s = nfkc(withoutReadingLines(v)).replace(/ +/g, ' ');
  return KANJI_NAME_RE.test(s) && HAN_RE.test(s);
}
export function isLatinNameShape(v: Cell): boolean {
  return typeof v === 'string' && LATIN_NAME_RE.test(nfkc(withoutReadingLines(v)).replace(/ +/g, ' '));
}
/** "some human-readable text, not a number/date/email/url" — lenient name evidence */
export function isTextish(v: Cell): boolean {
  if (typeof v !== 'string') return false;
  const s = nfkc(v);
  return s.length >= 2 && s.length <= 30 && !/^[\d\s\-.,:/]+$/.test(s) && !s.includes('@') && !/^https?:/i.test(s);
}
/** alnum ID shape: 3–24 chars of [A-Z0-9-_], at least one digit */
export function isIdShape(v: Cell): boolean {
  if (typeof v === 'number') return Number.isInteger(v) && v >= 0;
  if (typeof v !== 'string') return false;
  const s = nfkc(v).replace(WS_RE, '').replace(HYPHENS_RE, '-');
  return (/^[A-Za-z0-9][A-Za-z0-9\-_/.#]{1,23}$/.test(s) || THOUSANDS_RE.test(s)) && /\d/.test(s);
}
/** 'TEST00001' → 'AAAA99999' (pattern signature used for consistency scoring) */
export function idSignature(v: Cell): string {
  const s = typeof v === 'number' ? cellText(v) : nfkc(cellText(v)).replace(WS_RE, '').toUpperCase();
  return s.replace(/[A-Z]/g, 'A').replace(/\d/g, '9');
}

/** 合計 / 小計 / Total … : summary rows are never keys (kept in place, excluded from matching). */
export const SUMMARY_RE = /^(総?合計|小計|総計|累計|計|合\s*計|合計額|合計金額|total|totals|sub-?total|grand-?total|sum)([(（].*[)）])?[:：]?$/i;
/**
 * Statistic labels of a footer (人数 / 平均 / 件数 / 最大 / 最小 / Average / Count …). Weaker than 合計: a row that
 * starts with one is a summary row only when the rest of the row holds numbers / formulas, not other texts.
 */
export const SUMMARY_WEAK_RE =
  /^(総?人数|人数計|合計人数|総?件数|件数計|総数|平均|平均値|平均点|最大|最大値|最小|最小値|最高|最高点|最高値|最低|最低点|最低値|中央値|最頻値|標準偏差|分散|average|avg|ave|mean|count|max|maximum|min|minimum|median|stdev|std)([(（].*[)）])?[:：]?$/i;
const SUMMARY_HINT_RE = /計|[tTｔＴ][oOｏＯ][tTｔＴ]|[sSｓＳ][uUｕＵ][mMｍＭ]/;
const SUMMARY_WEAK_HINT_RE = /人数|件数|総数|平均|最[大小高低頻]|中央値|標準偏差|分散|[aAａＡ][vVｖＶ]|[mMｍＭ][eEｅＥaAａＡiIｉＩ]|[cCｃＣ][oOｏＯ][uUｕＵ]|[sSｓＳ][tTｔＴ][dDｄＤ]/;
/**
 * A page number line of a paged export: '- 2 -', '― 2 ―', 'Page 2 of 5', '2/5', '(2/5)', '第2ページ', '2ページ', 'p.2'.
 * Never data (and never a key).
 */
const PAGE_DECOR_RE = /[\s\-‐-―−－ー_=~〜・.*＊#■□●○◆◇<>《》【】\[\]()（）"'「」『』:：|｜]+/g;
const PAGE_MARKER_RE = /^(page|p|no|ページ|頁|第)?\d{1,4}(\/\d{1,4}|of\d{1,4}|の\d{1,4})?(ページ|頁|page)?$/;
/**
 * A remark line typed between the rows of a list: ※ 以下は追加分 / ---- 以下余白 ---- / 注）… / ＊ 辞退者 /  ◆ 2次募集.
 * Only ever applied to a row whose single non-blank cell it is.
 */
const NOTE_LINE_RE = /^([※＊*＃#◆◇■□●○▼▽▲△★☆〇]|注[:：)）意]|備考|以下|以上|(?:-{2,}|‐{2,}|―{2,}|={2,}|＝{2,}|~{2,}|〜{2,}|_{2,}))/;
export function isNoteLineText(v: Cell): boolean {
  if (typeof v !== 'string' || v.length > 80) return false;
  return NOTE_LINE_RE.test(nfkc(v).replace(/^[\s　]+/, ''));
}

/** the marker in front of a sample row of a form template: 例） 山田 太郎 / 記入例 / サンプル / 見本 / Sample */
const EXAMPLE_MARK_RE = /^(記入例|入力例|例|サンプル|見本|ダミー|テスト|sample|example|e\.g\.?)[)）:：\s]*$/i;
export function isExampleMark(v: Cell): boolean {
  return typeof v === 'string' && v.length <= 12 && EXAMPLE_MARK_RE.test(nfkc(v).replace(/^[\s　*＊※]+/, ''));
}

export function isPageMarkerText(v: Cell): boolean {
  if (typeof v !== 'string' || v.length > 30) return false;
  const n = nfkc(v).toLowerCase().replace(/ページ|頁/g, 'page'); // (before the decoration strip: ー of ページ is a dash there)
  if (/^\d+$/.test(n.trim())) return false; // a bare number is a value, not a page line
  const t = n.replace(PAGE_DECOR_RE, '');
  return t !== '' && /\d/.test(t) && PAGE_MARKER_RE.test(t);
}

/** decoration around a label: 【合計】, ＜平均＞, ■合計 */
const LABEL_DECOR_RE = /^[【\[〔〈《≪<＜■□◆◇●○◎★☆▼▽▲△・]+|[】\]〕〉》≫>＞]+$/g;

export function summaryLabelText(v: string): string {
  return nfkc(v).replace(WS_RE, '').replace(LABEL_DECOR_RE, '');
}

export function isSummaryText(v: Cell): boolean {
  if (typeof v !== 'string' || v.length > 40 || !SUMMARY_HINT_RE.test(v)) return false;
  const s = summaryLabelText(v);
  return s.length <= 20 && SUMMARY_RE.test(s);
}

/** 東京大学 小計 / ページ計 / 部門別計 / Group Total: a qualified subtotal label (only a summary when nothing but numbers sits beside it) */
const SUMMARY_SUFFIX_RE = /^.{0,20}(小計|合計|総計|累計|ページ計|頁計|別計|subtotal|total)[:：]?$/i;
const SUMMARY_SUFFIX_HINT_RE = /計|total/i;

/**
 * 'strong' = 合計 / 小計 / Total …, 'weak' = 人数 / 平均 / 件数 / 最大 / Count …,
 * 'suffix' = 東京大学 小計 / ページ計 (needs a row without any other text), null = not a summary label.
 */
export function summaryLabelKind(v: Cell): 'strong' | 'weak' | 'suffix' | null {
  if (typeof v !== 'string' || v.length > 40) return null;
  if (isSummaryText(v)) return 'strong';
  if (SUMMARY_SUFFIX_HINT_RE.test(v)) {
    const s = summaryLabelText(v);
    if (s.length <= 24 && SUMMARY_SUFFIX_RE.test(s)) return 'suffix';
  }
  if (!SUMMARY_WEAK_HINT_RE.test(v)) return null;
  const s = summaryLabelText(v);
  return s.length <= 20 && SUMMARY_WEAK_RE.test(s) ? 'weak' : null;
}
