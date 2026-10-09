/**
 * Column profiling + key-type classification (header synonyms × value patterns × uniqueness). Ported from R6.
 *
 * Performance (10,000–20,000 rows × 40 columns must stay usable):
 *  - value-shape ratios are estimated on an even sample of ≤ 1,000 non-blank cells;
 *  - uniqueness for typing uses a cheap "light" key (trimmed text) on every row;
 *  - type-specific keys are computed on every row only for typed columns; generic keys (value-overlap discovery)
 *    are computed lazily, only for columns that take part in a discovery comparison.
 */
import synonyms from './dict/key-synonyms.json';
import {
  type Cell, cellText, nfkc, isBlankLike, normId, idKeyOf, normName, kanaStrict, normEmail, normPhone, normOrg, normDate, normGeneric,
  isKanjiNameShape, isLatinNameShape, isTextish, isIdShape, idSignature, KANA_NAME_RE, UNIV_RE, FACULTY_RE,
} from './normalize';
import type { Table } from './table';

export type KeyType =
  | 'id' | 'studentId' | 'employeeId' | 'orgCode' | 'rowNumber'
  | 'name' | 'lastName' | 'firstName' | 'nameKana' | 'lastNameKana' | 'firstNameKana' | 'nameRoman'
  | 'email' | 'phone' | 'university' | 'faculty' | 'birthDate' | 'gender' | 'graduation';

interface Rule { re: string; score: number; subtype?: string }
interface TypeDef { label: string; isKey: boolean; defaultKey?: boolean; compositePart?: boolean; rules: Rule[]; negative?: string[] }
const DICT = synonyms as unknown as {
  noteWords: string[];
  nonKeyHeaders: { re: string[] };
  entityWords: Record<'subject' | 'staff' | 'other', string[]>;
  types: Record<string, TypeDef>;
};
const NON_KEY_RE = DICT.nonKeyHeaders.re.map((r) => new RegExp(r, 'u'));
export const isNonKeyHeader = (hn: string): boolean => NON_KEY_RE.some((r) => r.test(hn));
const TYPES = Object.keys(DICT.types) as KeyType[];
const COMPILED = Object.fromEntries(
  TYPES.map((t) => [
    t,
    {
      rules: DICT.types[t].rules.map((r) => ({ ...r, rx: new RegExp(r.re, 'u') })),
      neg: (DICT.types[t].negative ?? []).map((n) => new RegExp(n, 'u')),
    },
  ]),
) as Record<KeyType, { rules: (Rule & { rx: RegExp })[]; neg: RegExp[] }>;
const NOTE_RE = new RegExp(DICT.noteWords.map((w) => w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|'), 'gu');

export const TYPE_LABEL: Record<string, string> = Object.fromEntries(TYPES.map((t) => [t, DICT.types[t].label]));
TYPE_LABEL.employeeId = '社員番号';

/** Joinable families: columns of the same family can be compared across files. */
export const FAMILY: Record<string, string> = {
  id: 'idFamily', studentId: 'idFamily', employeeId: 'idFamily', name: 'name', lastName: 'lastName', firstName: 'firstName',
  nameKana: 'kana', lastNameKana: 'lastNameKana', firstNameKana: 'firstNameKana', nameRoman: 'roman', email: 'email', phone: 'phone',
  university: 'university', faculty: 'faculty', birthDate: 'birthDate', gender: 'gender', graduation: 'graduation', orgCode: 'orgCode',
};

export const ID_TYPES: ReadonlySet<string> = new Set(['id', 'studentId', 'employeeId']);

// ---------------------------------------------------------------------------
// header analysis
// ---------------------------------------------------------------------------

export function headerNorm(h: string): string {
  let s = nfkc(h).toLowerCase();
  s = s.replace(/[\s※*＊★☆◆◇■□●○・:：_\-‐()（）［］[\]【】「」『』<>〈〉《》.,、。?？!！'"/／]/gu, '');
  s = s.replace(NOTE_RE, '');
  s = s.replace(/^(q\d+|問\d+|\d+)[.)]?/u, ''); // "Q3. お名前" / "1.氏名"
  return s;
}

const headerCache = new Map<string, Record<KeyType, number>>();

export function headerScores(hn: string): Record<KeyType, number> {
  const cached = headerCache.get(hn);
  if (cached) return cached;
  const out = {} as Record<KeyType, number>;
  for (const t of TYPES) {
    const c = COMPILED[t];
    let best = 0;
    for (const r of c.rules) if (r.rx.test(hn)) best = Math.max(best, r.score);
    if (best > 0 && c.neg.some((n) => n.test(hn))) best = 0;
    // long question-style headers ("お名前をフルネームで教えてください") are weaker evidence
    if (best > 0 && hn.length > 14) best = Math.min(best, 0.7);
    out[t] = best;
  }
  if (headerCache.size > 5000) headerCache.clear();
  headerCache.set(hn, out);
  return out;
}

export function bestHeaderType(hn: string): { type: KeyType | null; score: number } {
  const s = headerScores(hn);
  let type: KeyType | null = null;
  let score = 0;
  for (const t of TYPES) {
    if (s[t] > score) {
      score = s[t];
      type = t;
    }
  }
  return { type, score };
}

function headerNegated(hn: string, t: KeyType): boolean {
  return COMPILED[t].neg.some((n) => n.test(hn));
}

export function entityOf(hn: string): 'subject' | 'staff' | 'other' | null {
  for (const e of ['staff', 'other', 'subject'] as const) if (DICT.entityWords[e].some((w) => hn.includes(w.toLowerCase()))) return e;
  return null;
}

// ---------------------------------------------------------------------------
// column profile
// ---------------------------------------------------------------------------

export interface UniqStats {
  distinct: number;
  uniqueness: number;
  dupRows: number;
  dupKeys: [string, number][];
}

export function uniq(keys: string[]): UniqStats {
  const m = new Map<string, number>();
  let nb = 0;
  for (const k of keys) {
    if (k) {
      nb++;
      m.set(k, (m.get(k) ?? 0) + 1);
    }
  }
  const dupKeys: [string, number][] = [];
  let dupRows = 0;
  for (const e of m) {
    if (e[1] > 1) {
      dupKeys.push(e);
      dupRows += e[1];
    }
  }
  return { distinct: m.size, uniqueness: nb ? m.size / nb : 0, dupRows, dupKeys };
}

const clamp = (x: number, a: number, b: number) => Math.max(a, Math.min(b, x));

/** type-specific key of one cell */
export function keyOf(type: KeyType | 'unknown', v: Cell, display?: string, serialOk = false): string {
  switch (type) {
    case 'id': case 'studentId': case 'employeeId': case 'orgCode':
      return idKeyOf(v, display);
    case 'lastName': case 'firstName':
      return normName(v).folded;
    case 'name': case 'nameRoman': {
      // Latin names are compared as token sets: 'Taro Yamada' ≡ 'YAMADA Taro' ≡ 'Yamada, Taro'
      const k = normName(v);
      return k.latinSorted || k.folded;
    }
    case 'nameKana': case 'lastNameKana': case 'firstNameKana':
      return kanaStrict(v);
    case 'email':
      return normEmail(v);
    case 'phone':
      return normPhone(v).digits;
    case 'university': case 'faculty':
      return normOrg(v);
    case 'birthDate':
      return normDate(v, serialOk).iso;
    default:
      return normGeneric(v);
  }
}

/** Per-row keys of one column (rows from..n-1); string values are memoised (low-cardinality columns are cheap). */
function computeKeys(table: Table, col: number, type: KeyType | 'unknown', serialOk: boolean, from = 0, useMemo = true): string[] {
  const memo = new Map<string, string>();
  const n = table.rows.length;
  const out = new Array<string>(Math.max(0, n - from));
  const isId = ID_TYPES.has(type) || type === 'orgCode';
  for (let i = from; i < n; i++) {
    const v = table.rows[i][col];
    if (v === null || v === undefined) {
      out[i - from] = '';
    } else if (typeof v === 'string') {
      let k = useMemo ? memo.get(v) : undefined;
      if (k === undefined) {
        k = keyOf(type, v, undefined, serialOk);
        if (useMemo) memo.set(v, k);
      }
      out[i - from] = k;
    } else {
      out[i - from] = keyOf(type, v, isId && typeof v === 'number' ? table.display(i, col) : undefined, serialOk);
    }
  }
  return out;
}

export class ColumnProfile {
  table: Table;
  /** index into table.headers / table.rows[i] */
  col: number;
  header: string;
  headerNorm: string;
  entity: 'subject' | 'staff' | 'other' | null;
  rows: number;
  nonBlank: number;
  blank: number;
  fill: number;
  distinct = 0;
  uniqueness = 0;
  /** rows whose (type-specific) key occurs more than once */
  dupRows = 0;
  dupKeys: [string, number][] = [];
  /** uniqueness of the cheap "light" key (trimmed text) — an upper bound of the generic-key uniqueness */
  lightUniq: UniqStats;
  ratios: Record<string, number>;
  headerScores: Record<KeyType, number>;
  valueScores: Record<KeyType, number>;
  typeScores: Record<KeyType, number>;
  type: KeyType | 'unknown' = 'unknown';
  typeScore = 0;
  confidence: 'high' | 'medium' | 'low' = 'low';
  /** joinability of this single column (type evidence × uniqueness × fill) */
  keyScore = 0;
  /** plain numbers are Excel date serials in this column */
  serialOk: boolean;
  notes: string[] = [];
  private keysCache: string[] | null = null;
  private genericCache: string[] | null = null;
  private genericUniqCache: UniqStats | null = null;
  /** clone on an extended table: rows below `rows` are the original's rows (keys reused) */
  private inherit: { from: ColumnProfile; rows: number } | null = null;

  constructor(init: {
    table: Table; col: number; header: string; headerNorm: string; nonBlank: number; lightUniq: UniqStats; ratios: Record<string, number>;
    headerScores: Record<KeyType, number>; valueScores: Record<KeyType, number>; typeScores: Record<KeyType, number>; serialOk: boolean;
  }) {
    this.table = init.table;
    this.col = init.col;
    this.header = init.header;
    this.headerNorm = init.headerNorm;
    this.entity = entityOf(init.headerNorm);
    this.rows = init.table.rows.length;
    this.nonBlank = init.nonBlank;
    this.blank = this.rows - init.nonBlank;
    this.fill = this.rows ? init.nonBlank / this.rows : 0;
    this.lightUniq = init.lightUniq;
    this.ratios = init.ratios;
    this.headerScores = init.headerScores;
    this.valueScores = init.valueScores;
    this.typeScores = init.typeScores;
    this.serialOk = init.serialOk;
  }

  /** per-row normalized key (type-specific), '' = blank */
  get keys(): string[] {
    if (!this.keysCache) {
      if (this.type === 'unknown') this.keysCache = this.generic;
      else if (this.inherit && this.inherit.from.type === this.type) {
        this.keysCache = this.inherit.from.keys.concat(computeKeys(this.table, this.col, this.type, this.serialOk, this.inherit.rows));
      } else this.keysCache = computeKeys(this.table, this.col, this.type, this.serialOk, 0, this.memoWorthIt);
    }
    return this.keysCache;
  }

  /** memoising per distinct text only pays off for repetitive columns */
  private get memoWorthIt(): boolean {
    return this.lightUniq.uniqueness < 0.6;
  }
  set keys(k: string[]) {
    this.keysCache = k;
  }

  /** per-row generic normalized value (for value-overlap discovery) */
  get generic(): string[] {
    if (!this.genericCache) {
      this.genericCache = this.inherit
        ? this.inherit.from.generic.concat(computeKeys(this.table, this.col, 'unknown', this.serialOk, this.inherit.rows))
        : computeKeys(this.table, this.col, 'unknown', this.serialOk, 0, this.memoWorthIt);
    }
    return this.genericCache;
  }

  /** mark this profile as a clone whose first `rows` table rows are `from`'s rows (see cloneProfile) */
  inheritFrom(from: ColumnProfile, rows: number): void {
    this.inherit = { from, rows };
  }

  get genericUniq(): UniqStats {
    if (!this.genericUniqCache) this.genericUniqCache = uniq(this.generic);
    return this.genericUniqCache;
  }

  /** key uniqueness stats: exact for typed columns, the light approximation for unknown ones (display only) */
  setStats(u: UniqStats): void {
    this.distinct = u.distinct;
    this.uniqueness = u.uniqueness;
    this.dupRows = u.dupRows;
    this.dupKeys = u.dupKeys;
  }
}

/**
 * Per-row keys of column p normalised AS `type` (not necessarily p's own type). Manual key candidates pair a column
 * whose type was not recognised ('unknown') with a typed one and normalise both like the typed one (電話番号 ⟷ 連絡先).
 * Reuses the profile's own keys when the type is the same; must run inside withNormalizeOptions like every key access.
 */
export function keysAs(p: ColumnProfile, type: KeyType | 'unknown'): string[] {
  if (type === p.type) return p.keys;
  if (type === 'unknown') return p.generic;
  return computeKeys(p.table, p.col, type, p.serialOk, 0, p.lightUniq.uniqueness < 0.6);
}

const RATIO_SAMPLE = 1000;

function ratio(xs: Cell[], pred: (v: Cell) => boolean): number {
  if (!xs.length) return 0;
  let n = 0;
  for (const v of xs) if (pred(v)) n++;
  return n / xs.length;
}

/** share of consecutive +1 steps, counted only for sequences that START at 0/1 (row numbers), not ID ranges like 000100.. */
function sequenceRatio(xs: Cell[]): number {
  const nums: number[] = [];
  for (const v of xs) {
    if (typeof v === 'number') nums.push(v);
    else if (typeof v === 'string' && /^\d{1,6}$/.test(v.trim())) nums.push(Number(v));
  }
  if (nums.length < 3) return 0;
  let min = Infinity;
  for (const x of nums.slice(0, 50)) min = Math.min(min, x);
  if (min > 1) return 0;
  let ok = 0;
  for (let i = 1; i < nums.length; i++) if (nums[i] - nums[i - 1] === 1) ok++;
  return ok / (nums.length - 1);
}

function lightKey(v: Cell): string {
  if (typeof v === 'string') return v.trim();
  if (typeof v === 'number') return String(v);
  return cellText(v);
}

const GENDER_RE = /^(男|女|男性|女性|m|f|male|female|その他|回答しない|無回答)$/i;
const SHORT_KANJI_RE = /^[\p{Script=Han}々ヶ]{1,4}$/u;
const NAME_LIKE_TYPES: ReadonlySet<string> = new Set(['name', 'nameKana', 'nameRoman', 'lastName', 'firstName', 'email', 'phone']);
const VALUE_ONLY_CAPPED: ReadonlySet<string> = new Set(['lastName', 'firstName', 'lastNameKana', 'firstNameKana', 'studentId', 'employeeId', 'orgCode', 'graduation', 'rowNumber', 'faculty']);
const VETO: Partial<Record<KeyType, number>> = { email: 0.3, phone: 0.3, birthDate: 0.3, nameKana: 0.3, name: 0.2, id: 0.2, studentId: 0.2, university: 0.2, faculty: 0.2 };

export function profileColumn(table: Table, col: number): ColumnProfile {
  const header = table.headers[col];
  const hn = headerNorm(header);
  const n = table.rows.length;
  const values: Cell[] = new Array(n);
  const nonBlankVals: Cell[] = [];
  const light: string[] = new Array(n);
  for (let i = 0; i < n; i++) {
    const v = table.rows[i][col];
    values[i] = v;
    if (isBlankLike(v)) {
      light[i] = '';
    } else {
      nonBlankVals.push(v);
      light[i] = lightKey(v) || ' ';
    }
  }
  const nonBlank = nonBlankVals.length;
  // value-shape ratios are estimated on an even sample; keys/uniqueness use every row
  const sample = nonBlank <= RATIO_SAMPLE ? nonBlankVals : Array.from({ length: RATIO_SAMPLE }, (_, i) => nonBlankVals[Math.floor((i * nonBlank) / RATIO_SAMPLE)]);
  const strSample = sample.filter((v): v is string => typeof v === 'string').map((v) => ({ raw: v, n: nfkc(v) }));
  const strRatio = (pred: (s: { raw: string; n: string }) => boolean) => {
    if (!sample.length) return 0;
    let k = 0;
    for (const s of strSample) if (pred(s)) k++;
    return k / sample.length;
  };

  const hs = headerScores(hn);
  const nonKeyHeader = isNonKeyHeader(hn);
  const isDateFmt = table.dateFmt[col] ?? false;
  // plain numbers are Excel serials / 20010401 only when the header says 生年月日 (an 8-digit ID may look like a date)
  const serialOk = isDateFmt || hs.birthDate > 0;
  const r: Record<string, number> = {
    number: ratio(sample, (v) => typeof v === 'number'),
    date: ratio(sample, (v) => v instanceof Date || (typeof v === 'number' && (isDateFmt || (serialOk && normDate(v, true).iso !== ''))) || (typeof v === 'string' && normDate(v, serialOk).iso !== '')),
    // short plain integers (No. 1..20, scores, ranks): counters, never identifiers
    smallInt: ratio(sample, (v) => (typeof v === 'number' ? Number.isInteger(v) && v >= 0 && v < 1000 : typeof v === 'string' && /^[0-9０-９]{1,3}$/.test(v.trim()))),
    email: strRatio((s) => normEmail(s.raw) !== ''),
    phone: ratio(sample, (v) => normPhone(v).digits !== '' && (typeof v !== 'number' || String(v).length >= 9) && !(typeof v === 'string' && /[a-z]/i.test(v))),
    kana: strRatio((s) => KANA_NAME_RE.test(kanaStrict(s.raw))),
    kanjiName: ratio(sample, isKanjiNameShape),
    latinName: ratio(sample, isLatinNameShape),
    textish: ratio(sample, isTextish),
    univ: strRatio((s) => UNIV_RE.test(s.n)),
    faculty: strRatio((s) => FACULTY_RE.test(s.n) && !UNIV_RE.test(s.n.replace(/大学院|研究科/g, ''))),
    shortKanji: strRatio((s) => SHORT_KANJI_RE.test(s.n.replace(/\s/g, ''))),
    idShape: ratio(sample, isIdShape),
    // strong ID evidence for header-less typing: letters+digits, leading zeros, or >= 4 digits (score 0..100 is not an ID)
    strongId: ratio(sample, (v) => {
      if (typeof v === 'number') return Number.isInteger(v) && v >= 1000;
      const s = typeof v === 'string' ? nfkc(v).replace(/\s/g, '') : '';
      return (/[A-Za-z]/.test(s) && /\d/.test(s)) || /^0\d{2,}$/.test(s) || /^\d{4,}$/.test(s);
    }),
    gender: strRatio((s) => GENDER_RE.test(s.n)),
    sequence: sequenceRatio(values.slice(0, 5000)),
  };
  // birth-year plausibility (1940–2015) for date columns
  const years: number[] = [];
  if (r.date > 0) {
    for (const v of sample) {
      const iso = normDate(v, serialOk).iso;
      if (iso) years.push(Number(iso.slice(0, 4)));
    }
  }
  r.birthYear = years.length ? years.filter((y) => y >= 1940 && y <= 2015).length / years.length : 0;
  // ID pattern consistency: share of the most common signature (AAAA99999)
  const sigs = new Map<string, number>();
  for (const v of sample) {
    const s = idSignature(v);
    sigs.set(s, (sigs.get(s) ?? 0) + 1);
  }
  let maxSig = 0;
  for (const c of sigs.values()) maxSig = Math.max(maxSig, c);
  r.idConsistency = sample.length ? maxSig / sample.length : 0;

  // uniqueness for typing: cheap light key on every row (≥ the generic-key uniqueness)
  const lightUniq = uniq(light);

  const vs = {} as Record<KeyType, number>;
  for (const t of TYPES) vs[t] = 0;
  const uniqF = clamp((lightUniq.uniqueness - 0.5) / 0.45, 0, 1);
  vs.id = r.idShape * (0.5 + 0.5 * r.idConsistency) * uniqF * (1 - r.date);
  vs.studentId = vs.id;
  vs.orgCode = r.idShape * (1 - uniqF * 0.5);
  vs.rowNumber = r.sequence >= 0.8 ? r.sequence : 0;
  const shape = Math.max(r.kanjiName, r.latinName);
  vs.name = (shape + 0.4 * Math.max(0, r.textish - shape)) * (1 - 0.8 * r.kana) * (1 - r.univ) * (1 - r.faculty) * (1 - r.email);
  const nameUniqF = clamp((lightUniq.uniqueness - 0.3) / 0.6, 0, 1); // person names: mostly distinct
  vs.lastName = r.shortKanji * (1 - r.kana);
  vs.firstName = vs.lastName;
  vs.nameKana = r.kana;
  vs.lastNameKana = r.kana;
  vs.firstNameKana = r.kana;
  vs.nameRoman = r.latinName;
  vs.email = r.email;
  vs.phone = r.phone * (1 - r.email);
  vs.university = r.univ;
  vs.faculty = r.faculty;
  vs.birthDate = r.date * (0.4 + 0.6 * r.birthYear);
  vs.gender = r.gender;
  vs.graduation = 0;

  // ---- combine header + value evidence ----
  const ts = {} as Record<KeyType, number>;
  for (const t of TYPES) {
    const h = hs[t];
    let v = vs[t];
    // an explicit 学籍番号 / 社員番号 / 応募者ID header: repeated values (one row per interview …) do not make it less of an ID
    if (h >= 0.85 && ID_TYPES.has(t)) v = Math.max(v, r.idShape * (0.5 + 0.5 * r.idConsistency) * (1 - r.date));
    let s: number;
    if (h >= 0.5) {
      s = 0.6 * h + 0.4 * v;
      const veto = VETO[t];
      if (veto !== undefined && v < veto && nonBlank >= 3) s *= 0.6; // header says X, values disagree
    } else {
      // value-only discovery, capped (header gives no or weak support)
      const vv = NAME_LIKE_TYPES.has(t) ? v * nameUniqF : ID_TYPES.has(t) ? v * r.strongId : v;
      s = 0.75 * vv * (vv >= 0.8 ? 1 : 0.6) + 0.25 * h;
      if (headerNegated(hn, t) || nonKeyHeader) s = Math.min(s, 0.3);
    }
    // value-only types that are too generic to name a column on their own
    if (h < 0.5 && VALUE_ONLY_CAPPED.has(t)) s = Math.min(s, 0.3);
    ts[t] = s;
  }
  // row-number columns: header says No/番号 AND consecutive integers
  if (hs.rowNumber > 0 && r.sequence >= 0.8) ts.rowNumber = 0.95;
  else ts.rowNumber = Math.min(ts.rowNumber, r.sequence >= 0.95 && r.number > 0.9 ? 0.6 : 0.2);
  // subtype from rules (社員番号 → employeeId)
  for (const rule of COMPILED.id.rules) if (rule.subtype === 'employeeId' && rule.rx.test(hn) && !headerNegated(hn, 'id')) ts.employeeId = ts.id + 0.01;

  const p = new ColumnProfile({ table, col, header, headerNorm: hn, nonBlank, lightUniq, ratios: r, headerScores: hs, valueScores: vs, typeScores: ts, serialOk });
  let type: KeyType | 'unknown' = 'unknown';
  let best = 0;
  for (const t of [...TYPES, 'employeeId' as KeyType]) {
    if ((ts[t] ?? 0) > best) {
      best = ts[t];
      type = t;
    }
  }
  if (best < 0.35 || nonBlank === 0) type = 'unknown';
  p.type = type;
  p.typeScore = best;
  p.confidence = best >= 0.75 ? 'high' : best >= 0.55 ? 'medium' : 'low';

  const u = type === 'unknown' ? lightUniq : uniq(p.keys);
  p.setStats(u);
  const isKey = type !== 'unknown' && (DICT.types[type]?.isKey ?? type === 'employeeId');
  p.keyScore = type === 'unknown' || type === 'rowNumber' || type === 'orgCode' ? 0
    : isKey ? best * (0.4 + 0.6 * u.uniqueness) * (0.5 + 0.5 * p.fill)
      : best * 0.5 * (0.5 + 0.5 * p.fill);
  if (type === 'phone' && values.some((v) => typeof v === 'number' && normPhone(v).restoredZero)) p.notes.push('先頭の0が消えた電話番号（数値セル）を補正');
  if (ID_TYPES.has(type)) {
    if (values.some((v) => typeof v === 'string' && /E\+?\d+$/i.test(v) && normId(v).lossyExponent)) p.notes.push('指数表記(8.00E+09)で桁が失われたIDがあります');
    let fromNum = 0;
    let zeroStr = 0;
    for (const v of values) {
      if (typeof v === 'number') fromNum++;
      else if (typeof v === 'string' && /^0\d+$/.test(nfkc(v))) zeroStr++;
    }
    if (fromNum && zeroStr) p.notes.push('数値のIDと先頭ゼロ付き文字列のIDが混在（先頭ゼロを無視して照合）');
  }
  return p;
}

// ---------------------------------------------------------------------------
// sheet-level context rules
// ---------------------------------------------------------------------------

export function profileTable(table: Table): ColumnProfile[] {
  const ps = table.headers.map((_, c) => profileColumn(table, c));
  // (1) 苗字 | 名前  → 名前 is the given name when its values are short and space-free
  for (let i = 0; i < ps.length; i++) {
    const p = ps[i];
    if (p.type !== 'lastName') continue;
    for (const j of [i + 1, i + 2]) {
      const q = ps[j];
      if (!q || !['name', 'unknown'].includes(q.type)) continue;
      const vals = table.rows.map((r) => r[j]).filter((v) => !isBlankLike(v)).map((v) => nfkc(cellText(v)));
      const short = vals.filter((v) => v.length <= 3 && !/\s/.test(v)).length / Math.max(1, vals.length);
      if (/名前|名$|なまえ|first|given/.test(q.headerNorm) && short >= 0.8) retype(q, 'firstName', 0.85, '姓列の隣の「名前」を「名」と判定');
    }
  }
  // (2) several ID columns: prefer subject over staff; mark staff-side keys
  const hasSubjectKey = ps.some((p) => p.entity !== 'staff' && p.keyScore >= 0.5);
  if (hasSubjectKey) for (const p of ps) if (p.entity === 'staff' && p.type !== 'unknown') p.notes.push('担当者/社員側の列（本人のキーではない可能性）');
  return ps;
}

/**
 * The same column classification on another table with the same columns (the base table extended by rows appended
 * from other files). Keys are computed lazily; with `sharedRows` the first rows are known to be p's own rows
 * (same values) and their keys are reused.
 */
export function cloneProfile(p: ColumnProfile, table: Table, sharedRows = 0): ColumnProfile {
  const q = new ColumnProfile({
    table, col: p.col, header: p.header, headerNorm: p.headerNorm, nonBlank: p.nonBlank, lightUniq: p.lightUniq, ratios: p.ratios,
    headerScores: p.headerScores, valueScores: p.valueScores, typeScores: p.typeScores, serialOk: p.serialOk,
  });
  q.entity = p.entity;
  q.type = p.type;
  q.typeScore = p.typeScore;
  q.confidence = p.confidence;
  q.keyScore = p.keyScore;
  q.notes = p.notes;
  q.setStats({ distinct: p.distinct, uniqueness: p.uniqueness, dupRows: p.dupRows, dupKeys: p.dupKeys });
  if (sharedRows > 0 && sharedRows <= table.rows.length && sharedRows === p.table.rows.length) q.inheritFrom(p, sharedRows);
  return q;
}

function retype(p: ColumnProfile, t: KeyType, score: number, note: string) {
  p.type = t;
  p.typeScore = score;
  p.confidence = score >= 0.75 ? 'high' : 'medium';
  p.notes.push(note);
  p.keys = p.table.rows.map((r) => keyOf(t, r[p.col]));
  const u = uniq(p.keys);
  p.setStats(u);
  p.keyScore = score * (0.4 + 0.6 * u.uniqueness) * (0.5 + 0.5 * p.fill);
}
