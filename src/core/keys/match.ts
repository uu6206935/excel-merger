/**
 * Cross-file key candidates, composite keys, tiered join (fallback chain), cross-checks and fuzzy suggestions.
 * Ported from research R6 (match.ts) and extended for the app:
 *   - duplicate policy (first / last / skip) for several other-file rows on one base row;
 *   - manual links (forced pairs, processed before every tier);
 *   - ID tiers: base rows sharing an ID all receive the other row (VLOOKUP semantics, flagged);
 *   - specs can be rebuilt on another profile set (extended base table for the sequential 'full' join);
 *   - suggestions use blocking when unmatched A × B is large.
 */
import { distance } from 'fastest-levenshtein';
import { type ColumnProfile, type KeyType, type UniqStats, FAMILY, ID_TYPES, TYPE_LABEL, keysAs, uniq, isNonKeyHeader } from './detect';
import { type NameKey, hiraToKata, normName, normKana, normGeneric, kanaLoose, emailLoose, normId, foldItaiji } from './normalize';
import { prefixVariant, suffixVariant } from './idvariants';
import type { Table } from './table';

// ---------------------------------------------------------------------------
// key "columns": a single column or a composite of columns (virtual column)
// ---------------------------------------------------------------------------

export interface KeySpec {
  /** display label, e.g. 'ID（応募者ID）', '氏名＋大学', '姓＋名（姓＋名）' */
  label: string;
  /** type-ish id used to align specs across files: 'idFamily', 'name', 'name+university' ... */
  sig: string;
  cols: ColumnProfile[];
  /** per-row key ('' = at least one part blank) */
  keys: string[];
  kind: 'single' | 'composite' | 'discovered';
  /** composite keys: the parts (single or 姓＋名 specs), in order */
  parts?: KeySpec[];
  /** 姓＋名 / セイ＋メイ virtual full-name column */
  virtual?: boolean;
  /**
   * manual candidates (manual.ts): the key type cols[0] is normalised as ('unknown' = generic text), so that
   * rebuildSpec reproduces the same keys on another profile set
   */
  keyType?: KeyType | 'unknown';
  /** ID variant over the same columns: 'prefix' = letter prefix ignored (S2023-0012 ⟷ 20230012), 'suffix' = shorter ID is the tail of the longer */
  variant?: 'prefix' | 'suffix';
  /** pair-specific key recipe for another profile of the same column (rebuildSpec on the extended base) */
  keysFor?: (p: ColumnProfile) => string[];
  /** the default ID＋名前 key (see idNameComposite): `label` is shown as it is ('ID＋部品名') */
  fixedLabel?: boolean;
}

export const SEP = '\u0001';

const LATIN_TOKEN_RE = /^\p{Script=Latin}[\p{Script=Latin}'.-]*$/u;

/**
 * 姓+名 → produces the SAME key as a full-name column (folded, no space) so they can be compared.
 * Latin parts ('Yamada' + 'Taro') become the token-sorted form 'taro yamada' a Latin full-name column has.
 */
function concatNameKeys(parts: ColumnProfile[]): string[] {
  const n = parts[0].keys.length;
  const cols = parts.map((p) => p.keys);
  const out = new Array<string>(n);
  for (let i = 0; i < n; i++) {
    let s = '';
    let ok = true;
    let latin = true;
    for (const ks of cols) {
      const k = ks[i];
      if (!k) {
        ok = false;
        break;
      }
      if (!LATIN_TOKEN_RE.test(k)) latin = false;
      s += k;
    }
    if (ok && latin && cols.length > 1) s = cols.map((ks) => ks[i]).sort().join(' ');
    out[i] = ok ? s : '';
  }
  return out;
}

const latinCache = new WeakMap<ColumnProfile, boolean>();
/** most of the column's keys are Latin-script name tokens (Yamada / Taro / Mary-Ann) */
function isLatinColumn(p: ColumnProfile): boolean {
  let v = latinCache.get(p);
  if (v === undefined) {
    let n = 0;
    let l = 0;
    for (const k of p.keys) {
      if (!k) continue;
      n++;
      if (LATIN_TOKEN_RE.test(k) || /^\p{Script=Latin}[\p{Script=Latin}'.-]*( \p{Script=Latin}[\p{Script=Latin}'.-]*)+$/u.test(k)) l++;
    }
    v = n > 0 && l / n >= 0.8;
    latinCache.set(p, v);
  }
  return v;
}

/** family signature of a column: Latin-script names ('Taro Yamada' under Name / 氏名) form the 'roman' family */
export function sigOf(p: ColumnProfile): string {
  if ((p.type === 'name' || p.type === 'nameRoman') && (p.ratios.latinName ?? 0) >= 0.5 && (p.ratios.kanjiName ?? 0) < 0.3) return 'roman';
  return FAMILY[p.type] ?? p.type;
}

export function singleSpec(p: ColumnProfile): KeySpec {
  return { label: `${TYPE_LABEL[p.type] ?? p.header}（${p.header}）`, sig: sigOf(p), cols: [p], keys: p.keys, kind: 'single' };
}

/** Virtual full-name / full-kana column built from split 姓・名 columns. */
export function splitNameSpecs(ps: ColumnProfile[]): KeySpec[] {
  const out: KeySpec[] = [];
  const pick = (t: KeyType) => ps.filter((p) => p.type === t).sort((a, b) => b.typeScore - a.typeScore)[0];
  const ln = pick('lastName');
  const fn = pick('firstName');
  if (ln && fn) out.push({ label: `姓＋名（${ln.header}＋${fn.header}）`, sig: isLatinColumn(ln) && isLatinColumn(fn) ? 'roman' : 'name', cols: [ln, fn], keys: concatNameKeys([ln, fn]), kind: 'composite', virtual: true });
  const lk = pick('lastNameKana');
  const fk = pick('firstNameKana');
  if (lk && fk) out.push({ label: `セイ＋メイ（${lk.header}＋${fk.header}）`, sig: 'kana', cols: [lk, fk], keys: concatNameKeys([lk, fk]), kind: 'composite', virtual: true });
  return out;
}

export function compositeSpec(parts: KeySpec[]): KeySpec {
  const n = parts[0].keys.length;
  const keys = new Array<string>(n);
  for (let i = 0; i < n; i++) {
    let s = '';
    let ok = true;
    for (let p = 0; p < parts.length; p++) {
      const k = parts[p].keys[i];
      if (!k) {
        ok = false;
        break;
      }
      s += p ? SEP + k : k;
    }
    keys[i] = ok ? s : '';
  }
  return {
    label: parts.map((p) => p.label.replace(/（.*）$/, '')).join('＋'),
    sig: parts.map((p) => p.sig).join('+'),
    cols: parts.flatMap((p) => p.cols),
    keys,
    kind: 'composite',
    parts,
  };
}

/** The same key spec on another profile set of the same columns (e.g. the base table extended by appended rows). */
export function rebuildSpec(s: KeySpec, ps: ColumnProfile[]): KeySpec {
  if (s.parts) {
    const c = compositeSpec(s.parts.map((p) => rebuildSpec(p, ps)));
    return { ...c, label: s.label, sig: s.sig };
  }
  const cols = s.cols.map((c) => ps[c.col]);
  if (s.virtual) return { ...s, cols, keys: concatNameKeys(cols) };
  if (s.keysFor) return { ...s, cols, keys: s.keysFor(cols[0]) };
  if (s.kind === 'discovered') return { ...s, cols, keys: cols[0].generic };
  if (s.keyType !== undefined) return { ...s, cols, keys: keysAs(cols[0], s.keyType) };
  return { ...s, cols, keys: cols[0].keys };
}

// ---------------------------------------------------------------------------
// statistics of one key pair (base A ⟷ other B)
// ---------------------------------------------------------------------------

export interface PairStats {
  a: KeySpec;
  b: KeySpec;
  nonBlankA: number;
  nonBlankB: number;
  blankA: number;
  blankB: number;
  uniqA: number;
  uniqB: number;
  dupRowsA: number;
  dupRowsB: number;
  /** B rows (non-blank key) whose key exists in A — "match rate" shown in the UI */
  matchedB: number;
  matchRate: number;
  /** A rows whose key exists in B (coverage of the base) */
  matchedA: number;
  coverageA: number;
  /** B rows whose key hits >1 row in A (ambiguous) */
  ambiguousB: number;
  jaccard: number;
  cardinality: '1:1' | '1:N' | 'N:1' | 'N:M';
  unmatchedA: number[];
  unmatchedB: number[];
  score: number;
  verdict: '良好' | '一部一致' | '不一致多数' | '使用不可';
  /**
   * ID pairs only: share of the ID-matched rows whose NAMES disagree (same number, other person) — set when ≥ 3 rows can be
   * compared. ≥ 50% means the two columns are different numbering systems that happen to overlap.
   */
  nameConflict?: number;
}

function countIndex(keys: string[]): Map<string, number> {
  const m = new Map<string, number>();
  for (const k of keys) if (k) m.set(k, (m.get(k) ?? 0) + 1);
  return m;
}

function uniqFromCounts(m: Map<string, number>, nonBlank: number): { uniqueness: number; dupRows: number } {
  let dupRows = 0;
  for (const c of m.values()) if (c > 1) dupRows += c;
  return { uniqueness: nonBlank ? m.size / nonBlank : 0, dupRows };
}

export function pairStats(a: KeySpec, b: KeySpec): PairStats {
  const cntA = countIndex(a.keys);
  const cntB = countIndex(b.keys);
  let nonBlankA = 0;
  let nonBlankB = 0;
  let matchedB = 0;
  let ambiguousB = 0;
  const unmatchedB: number[] = [];
  for (let i = 0; i < b.keys.length; i++) {
    const k = b.keys[i];
    if (!k) continue;
    nonBlankB++;
    const hit = cntA.get(k);
    if (hit) {
      matchedB++;
      if (hit > 1) ambiguousB++;
    } else unmatchedB.push(i);
  }
  let matchedA = 0;
  const unmatchedA: number[] = [];
  for (let i = 0; i < a.keys.length; i++) {
    const k = a.keys[i];
    if (!k) continue;
    nonBlankA++;
    if (cntB.has(k)) matchedA++;
    else unmatchedA.push(i);
  }
  let inter = 0;
  const [small, big] = cntA.size <= cntB.size ? [cntA, cntB] : [cntB, cntA];
  for (const k of small.keys()) if (big.has(k)) inter++;
  const uA = uniqFromCounts(cntA, nonBlankA);
  const uB = uniqFromCounts(cntB, nonBlankB);
  const dupA = uA.uniqueness < 1;
  const dupB = uB.uniqueness < 1;
  const matchRate = nonBlankB ? matchedB / nonBlankB : 0;
  const coverageA = nonBlankA ? matchedA / nonBlankA : 0;
  const fillMin = Math.min(nonBlankA / Math.max(1, a.keys.length), nonBlankB / Math.max(1, b.keys.length));
  // ranking score: best of both directions × key-ness × fill (see report §6)
  const score = Math.max(matchRate, coverageA) * Math.min(uA.uniqueness, uB.uniqueness) ** 2 * Math.sqrt(fillMin);
  const best = Math.max(matchRate, coverageA);
  return {
    a, b, nonBlankA, nonBlankB, blankA: a.keys.length - nonBlankA, blankB: b.keys.length - nonBlankB,
    uniqA: uA.uniqueness, uniqB: uB.uniqueness, dupRowsA: uA.dupRows, dupRowsB: uB.dupRows,
    matchedB, matchRate, matchedA, coverageA, ambiguousB,
    jaccard: cntA.size + cntB.size - inter ? inter / (cntA.size + cntB.size - inter) : 0,
    cardinality: dupA && dupB ? 'N:M' : dupA ? 'N:1' : dupB ? '1:N' : '1:1',
    unmatchedA, unmatchedB, score,
    verdict: best >= 0.8 ? '良好' : best >= 0.5 ? '一部一致' : best >= 0.1 ? '不一致多数' : '使用不可',
  };
}

// ---------------------------------------------------------------------------
// candidate discovery between two tables
// ---------------------------------------------------------------------------

const NAME_TYPES_FOR_ENTITY: ReadonlySet<string> = new Set(['name', 'nameKana', 'nameRoman', 'lastName', 'firstName', 'lastNameKana', 'firstNameKana']);
/** cells touched by the value-overlap discovery of one file pair before it stops looking (a few seconds at most) */
const DISCOVERY_WORK_CAP = 4_000_000;
const MAX_DISCOVERY_COLUMNS = 40;
const PART_SIGS = new Set(['name', 'kana', 'roman', 'university', 'faculty', 'birthDate', 'email', 'phone', 'gender', 'graduation']);
const PERSON_PARTS = ['name', 'kana', 'roman'];

export function specsOf(ps: ColumnProfile[]): KeySpec[] {
  const singles = ps.filter((p) => p.type !== 'unknown' && p.type !== 'rowNumber' && p.typeScore >= 0.45).map(singleSpec);
  return [...singles, ...splitNameSpecs(ps)];
}

const NAME_LIKE_SIGS: ReadonlySet<string> = new Set(['name', 'roman']);

/** most specs of one family (ID / 氏名 / メール …) compared between two files: a 100-column export of ID-looking codes stays cheap */
const MAX_SPECS_PER_SIG = 6;

/** the best MAX_SPECS_PER_SIG specs of every family (by type confidence, then uniqueness) */
function limitPerSig(specs: KeySpec[]): KeySpec[] {
  const count = new Map<string, number>();
  for (const s of specs) count.set(s.sig, (count.get(s.sig) ?? 0) + 1);
  if ([...count.values()].every((n) => n <= MAX_SPECS_PER_SIG)) return specs;
  const score = (s: KeySpec) => s.cols[0].typeScore + 0.1 * s.cols[0].uniqueness;
  const keep = new Set<KeySpec>();
  for (const sig of count.keys()) {
    specs.filter((s) => s.sig === sig).sort((a, b) => score(b) - score(a) || a.cols[0].col - b.cols[0].col).slice(0, MAX_SPECS_PER_SIG).forEach((s) => keep.add(s));
  }
  return specs.filter((s) => keep.has(s));
}

export interface Candidates {
  singles: PairStats[];
  discovered: PairStats[];
  composites: PairStats[];
  /** the default key: the ID column ＋ the name column (氏名 / 部品名 …) together (also the first entry of `composites`) */
  idName?: PairStats;
}

const genericSets = new WeakMap<ColumnProfile, { set: Set<string>; list: string[] }>();
/** distinct non-blank generic values of a column (cached per profile) */
function distinctGeneric(p: ColumnProfile): { set: Set<string>; list: string[] } {
  let s = genericSets.get(p);
  if (!s) {
    const set = new Set<string>();
    for (const k of p.generic) if (k) set.add(k);
    s = { set, list: [...set] };
    genericSets.set(p, s);
  }
  return s;
}

/** true when any of 64 row-sampled values of p (generic-normalised on the fly) is in `set` */
function probeGeneric(p: ColumnProfile, set: Set<string>): boolean {
  const rows = p.table.rows;
  const n = rows.length;
  if (!n) return false;
  for (let k = 0; k < 64; k++) {
    const v = rows[Math.floor((k * n) / 64)]?.[p.col];
    if (v === null || v === undefined) continue;
    const g = normGeneric(v);
    if (g && set.has(g)) return true;
  }
  return false;
}

function bestNameSpec(specs: KeySpec[]): KeySpec | undefined {
  return specs.filter((s) => s.sig === 'name' || s.sig === 'roman')
    .sort((x, y) => (x.sig === 'name' ? 0 : 1) - (y.sig === 'name' ? 0 : 1) || y.cols[0].typeScore - x.cols[0].typeScore)[0];
}

/** 'han' (has a kanji) / 'kana' (kana only) / 'other' (Latin …): names are only compared within one kind */
function scriptClass(s: string): string {
  if (/\p{Script=Han}/u.test(s)) return 'han';
  return /^[\p{Script=Hiragana}\p{Script=Katakana}ー]+$/u.test(s) ? 'kana' : 'other';
}

/** share of the ID-matched rows (first row of each ID on the other side) whose names differ; 0 when < 3 rows are comparable */
function nameConflictRate(s: PairStats, nameA?: KeySpec, nameB?: KeySpec): number {
  if (!nameA || !nameB || nameA.sig !== nameB.sig) return 0;
  const first = new Map<string, number>();
  for (let j = 0; j < s.b.keys.length; j++) {
    const k = s.b.keys[j];
    if (k && !first.has(k)) first.set(k, j);
  }
  let agree = 0;
  let disagree = 0;
  for (let i = 0; i < s.a.keys.length; i++) {
    const k = s.a.keys[i];
    if (!k) continue;
    const j = first.get(k);
    if (j === undefined) continue;
    const x = nameA.keys[i];
    const y = nameB.keys[j];
    if (!x || !y) continue;
    // a kanji name against a kana-only one (or a Latin one) cannot be judged: only names of the same kind are compared
    if (scriptClass(x) !== scriptClass(y)) continue;
    const short = Math.min(x.length, y.length);
    // equal, a typo apart, or one is part of the other (surname only ⟷ full name)
    if (x === y || confusableFold(x) === confusableFold(y) || (short >= 3 && distance(x, y) <= 1) || (short >= 2 && (x.includes(y) || y.includes(x)))) agree++;
    else disagree++;
  }
  return agree + disagree >= 3 ? disagree / (agree + disagree) : 0;
}

export function findCandidates(A: ColumnProfile[], B: ColumnProfile[]): Candidates {
  const sa = limitPerSig(specsOf(A));
  const sb = limitPerSig(specsOf(B));
  const singles: PairStats[] = [];
  // (a column of mixed kanji / Latin names may be 'name' in one file and 'roman' in the other: both use the same key)
  const comparable = (x: KeySpec, y: KeySpec) => x.sig === y.sig || (NAME_LIKE_SIGS.has(x.sig) && NAME_LIKE_SIGS.has(y.sig));
  for (const x of sa) for (const y of sb) if (comparable(x, y) && x.sig !== 'orgCode') singles.push(pairStats(x, y));
  // ID columns that do not match literally but are the same numbering (S2023-0012 ⟷ 20230012, 20230012 ⟷ 230012)
  const exactIds = singles.filter((s) => s.a.sig === 'idFamily');
  for (const e of exactIds) {
    const best = Math.max(e.matchRate, e.coverageA);
    if (best >= 0.6 || e.a.cols.length !== 1 || e.b.cols.length !== 1) continue;
    for (const v of [prefixVariant(e.a.keys, e.b.keys), suffixVariant(e.a.keys, e.b.keys)]) {
      if (!v) continue;
      const x: KeySpec = { ...e.a, label: v.label, keys: v.a.keys, variant: v.variant, keysFor: v.a.keysFor };
      const y: KeySpec = { ...e.b, label: v.label, keys: v.b.keys, variant: v.variant, keysFor: v.b.keysFor };
      const st = pairStats(x, y);
      const vb = Math.max(st.matchRate, st.coverageA);
      if (vb >= 0.5 && vb >= best + 0.2) singles.push(st);
    }
  }
  // two ID columns that overlap by accident (No. 1–40 of two different systems): when the people they pair up do not bear
  // the same name, the numbering is not shared — the pair is demoted (never an automatic tier) and says so
  const nameA = bestNameSpec(sa);
  const nameB = bestNameSpec(sb);
  for (const s of singles) {
    if (s.a.sig !== 'idFamily' || Math.max(s.matchRate, s.coverageA) < 0.2) continue;
    const r = nameConflictRate(s, nameA, nameB);
    if (r > 0) s.nameConflict = r;
    if (r >= 0.5 && (s.verdict === '良好' || s.verdict === '一部一致')) s.verdict = '不一致多数';
  }
  singles.sort((p, q) => q.score - p.score);

  // ---- value-overlap discovery (any header): high-uniqueness columns whose normalized values overlap ----
  const discovered: PairStats[] = [];
  const known = new Set(singles.filter((s) => s.matchedB > 0).map((s) => `${s.a.cols.map((c) => c.col)}|${s.b.cols.map((c) => c.col)}`));
  // sparse columns (備考 filled on 3 of 30 rows) and remark/status headers are never keys
  const eligible = (p: ColumnProfile) =>
    p.type !== 'rowNumber' && p.type !== 'birthDate' && p.type !== 'gender' && p.fill >= 0.5 && !isNonKeyHeader(p.headerNorm)
    && p.ratios.date < 0.5 && p.ratios.gender < 0.5 && (p.type !== 'unknown' || (p.ratios.smallInt ?? 0) < 0.8) && p.lightUniq.distinct >= 3;
  // at most MAX_DISCOVERY_COLUMNS per side (the most unique ones): every column of a 100-column export may look like a key
  const top = (ps: ColumnProfile[]) => (ps.length <= MAX_DISCOVERY_COLUMNS ? ps : [...ps].sort((p, q) => q.lightUniq.uniqueness - p.lightUniq.uniqueness || p.col - q.col).slice(0, MAX_DISCOVERY_COLUMNS));
  const EA = top(A.filter(eligible));
  const EB = top(B.filter(eligible));
  const gu = (p: ColumnProfile): UniqStats => p.genericUniq;
  // typed ⟷ typed pairs are the typed candidates' job; discovery looks at unknown columns and ID-like codes
  // (職場ID ⟷ 応募者コード, 職場名 ⟷ 漢字氏名, a header-less column full of IDs …)
  const idish = (p: ColumnProfile) => ID_TYPES.has(p.type) || p.type === 'orgCode';
  // wide files whose columns all carry the same unique-looking values overlap pairwise: bound the work (≈ cells touched)
  let discoveryWork = 0;
  discover: for (const pa of EA) {
    for (const pb of EB) {
      if (discoveryWork > DISCOVERY_WORK_CAP) break discover;
      if (known.has(`${pa.col}|${pb.col}`)) continue;
      if (pa.type !== 'unknown' && pb.type !== 'unknown' && !(idish(pa) && idish(pb))) continue;
      if (pa.ratios.sequence >= 0.8 && pb.ratios.sequence >= 0.8) continue; // two 1..N columns always overlap
      // PK–FK aware: one side must be (nearly) unique; the other may repeat (職場ID in a matching sheet).
      // light uniqueness (trimmed text) is an upper bound of the generic one: cheap pre-filter
      if (Math.max(pa.lightUniq.uniqueness, pb.lightUniq.uniqueness) < 0.88) continue;
      // cheap probe before normalising the whole of the more repetitive column: 64 of its (row-sampled) values
      // against the other column's distinct values — 0 hits means containment ≥ 0.3 is practically impossible
      const [ps, pl] = pa.lightUniq.distinct <= pb.lightUniq.distinct ? [pa, pb] : [pb, pa];
      if (ps.lightUniq.distinct > 128 && !probeGeneric(ps, distinctGeneric(pl).set)) continue;
      const ua = gu(pa);
      const ub = gu(pb);
      if (ua.distinct < 3 || ub.distinct < 3) continue;
      if (Math.max(ua.uniqueness, ub.uniqueness) < 0.9) continue; // neither side is a key
      const da = distinctGeneric(pa);
      const db = distinctGeneric(pb);
      const [small, big] = da.list.length <= db.list.length ? [da.list, db.set] : [db.list, da.set];
      // quick probe: containment ≥ 0.3 of the smaller side is (practically) impossible with 0 of 64 probes hitting
      if (small.length > 128) {
        let probeHits = 0;
        for (let k = 0; k < 64; k++) if (big.has(small[Math.floor((k * small.length) / 64)])) probeHits++;
        if (probeHits === 0) continue;
      }
      let inter = 0;
      discoveryWork += small.length;
      for (const k of small) if (big.has(k)) inter++;
      // containment of the NON-unique (or smaller) side's distinct values in the other side
      const containment = inter / Math.max(1, small.length);
      if (inter >= Math.min(3, small.length) && containment >= 0.3) {
        const x: KeySpec = { label: `値一致（${pa.header}）`, sig: 'discovered', cols: [pa], keys: pa.generic, kind: 'discovered' };
        const y: KeySpec = { label: `値一致（${pb.header}）`, sig: 'discovered', cols: [pb], keys: pb.generic, kind: 'discovered' };
        discoveryWork += x.keys.length + y.keys.length;
        const st = pairStats(x, y);
        // PK-FK relation: rank by containment of distinct values rather than row match rate
        st.score = containment * Math.max(ua.uniqueness, ub.uniqueness) ** 2;
        discovered.push(st);
      }
    }
  }
  discovered.sort((p, q) => q.score - p.score);

  // ---- composite keys ----
  const composites: PairStats[] = [];
  const bySigA = new Map<string, KeySpec>();
  const bySigB = new Map<string, KeySpec>();
  const better = (s: KeySpec, o: KeySpec | undefined) => !o || s.cols[0].typeScore > o.cols[0].typeScore;
  for (const s of sa) if (PART_SIGS.has(s.sig) && better(s, bySigA.get(s.sig))) bySigA.set(s.sig, s);
  for (const s of sb) if (PART_SIGS.has(s.sig) && better(s, bySigB.get(s.sig))) bySigB.set(s.sig, s);
  const common = [...bySigA.keys()].filter((k) => bySigB.has(k));
  const personParts = common.filter((k) => PERSON_PARTS.includes(k));
  const combos: string[][] = [];
  for (const p of personParts) for (const q of common) if (q !== p && !(personParts.includes(q) && q < p)) combos.push([p, q]);
  for (const p of personParts) {
    for (let i = 0; i < common.length; i++) {
      for (let j = i + 1; j < common.length; j++) {
        const q = common[i];
        const r = common[j];
        if (q === p || r === p) continue;
        combos.push([p, q, r]);
      }
    }
  }
  const TEMPLATE_BONUS: Record<string, number> = { 'name+university': 0.05, 'name+birthDate': 0.05, 'kana+university': 0.04, 'name+kana': 0.03, 'kana+birthDate': 0.03, 'name+email': 0.02 };
  for (const c of combos) {
    const x = compositeSpec(c.map((s) => bySigA.get(s)!));
    const y = compositeSpec(c.map((s) => bySigB.get(s)!));
    const st = pairStats(x, y);
    st.score = st.score - 0.03 * (c.length - 2) + (TEMPLATE_BONUS[c.join('+')] ?? 0);
    composites.push(st);
  }
  // drop dominated combos: a strict superset that is not more unique / not better matching than one of its subsets
  const parts = (s: PairStats) => s.a.sig.split('+');
  const dominated = (s: PairStats) =>
    composites.some((o) => o !== s && parts(o).length < parts(s).length && parts(o).every((x) => parts(s).includes(x))
      && o.uniqA >= s.uniqA - 1e-9 && o.uniqB >= s.uniqB - 1e-9 && Math.max(o.matchRate, o.coverageA) >= Math.max(s.matchRate, s.coverageA) - 1e-9);
  // a composite is only interesting when its person part alone is NOT unique in one of the files
  const singleUnique = (sig: string) => singles.some((s) => s.a.sig === sig && s.uniqA >= 0.999 && s.uniqB >= 0.999);
  const kept = composites.filter((s) => !dominated(s) && !singleUnique(parts(s)[0]));
  kept.sort((p, q) => q.score - p.score);
  const idName = idNameComposite(singles, discovered);
  return { singles, discovered, composites: idName ? [idName, ...kept] : kept, ...(idName ? { idName } : {}) };
}

// ---------------------------------------------------------------------------
// the default key: ID ＋ name (a row joins only when BOTH agree)
// ---------------------------------------------------------------------------

/** the ID pair the plan uses (see planTiers): must overlap, be key-like, and not be two numberings that merely overlap */
function pickIdSingle(singles: PairStats[]): PairStats | undefined {
  const usable = (s: PairStats, minRate = 0.1) => Math.max(s.matchRate, s.coverageA) >= minRate;
  // one side (nearly) unique, the other may repeat (one row per interview): an explicit 学籍番号 / 社員番号 / 応募者ID header on both
  // sides makes a repeating ID a foreign key (VLOOKUP / duplicate policy); a vaguer header needs ≥ 50% uniqueness on both
  const explicitId = (s: PairStats) => [...s.a.cols, ...s.b.cols].every((p) => p.type !== 'unknown' && (p.headerScores[p.type] ?? 0) >= 0.85);
  const keyish = (s: PairStats) => Math.max(s.uniqA, s.uniqB) >= 0.9 && (Math.min(s.uniqA, s.uniqB) >= 0.5 || (Math.min(s.uniqA, s.uniqB) >= 0.1 && explicitId(s)));
  const ok = singles.filter((s) => s.a.sig === 'idFamily' && usable(s, 0.2) && keyish(s) && (s.nameConflict ?? 0) < 0.5);
  // a column headed 部品ID / 学生ID … beats a number column that merely looks like an ID (単価, 金額 …) even when the latter overlaps more
  const idHeader = (s: PairStats) => [...s.a.cols, ...s.b.cols].every((p) => /(id|番号|品番|コード|code|no\.?)$/i.test(p.header.trim()));
  return ok.find(explicitId) ?? ok.find(idHeader) ?? ok[0];
}

/**
 * ID ＋ name as ONE key ('ID＋氏名', 'ID＋部品名'): the name is the typed person name (氏名 / ローマ字) when both files have one, else the
 * value-overlap column that pairs up two name-like text columns (部品名 ⟷ 品名). Only built when an ID pair is usable and the name
 * pair overlaps well; a row whose ID agrees but whose name does not (or the other way round) is then NOT joined.
 */
function idNameComposite(singles: PairStats[], discovered: PairStats[]): PairStats | undefined {
  const id = pickIdSingle(singles);
  if (!id || id.a.cols.length !== 1 || id.b.cols.length !== 1) return undefined;
  const rate = (s: PairStats) => Math.max(s.matchRate, s.coverageA);
  const typed = singles.find((s) => (s.a.sig === 'name' || s.a.sig === 'roman') && s.a.cols.length === 1 && s.b.cols.length === 1 && rate(s) >= 0.5);
  const idish = (p: ColumnProfile) => ID_TYPES.has(p.type) || p.type === 'orgCode';
  const found = discovered.find((d) => d.a.cols.length === 1 && d.b.cols.length === 1 && rate(d) >= 0.5
    && d.a.cols[0].col !== id.a.cols[0].col && d.b.cols[0].col !== id.b.cols[0].col
    && !idish(d.a.cols[0]) && !idish(d.b.cols[0]) && (d.a.cols[0].entity ?? 'subject') === (d.b.cols[0].entity ?? 'subject'));
  const nm = typed ?? found;
  if (!nm) return undefined;
  const x = compositeSpec([id.a, nm.a]);
  const y = compositeSpec([id.b, nm.b]);
  x.label = y.label = `ID＋${typed ? '氏名' : nm.a.cols[0].header}`;
  x.fixedLabel = y.fixedLabel = true;
  return pairStats(x, y);
}

// ---------------------------------------------------------------------------
// tiered join (fallback chain)
// ---------------------------------------------------------------------------

export interface Tier {
  name: string;
  a: KeySpec;
  b: KeySpec;
  /** the candidate this tier was made from */
  stats?: PairStats;
}
export interface Flag {
  level: 'info' | 'warn' | 'error';
  msg: string;
}
export interface MatchPair {
  a: number;
  b: number;
  tier: string;
  /** index into the tiers (-1 = manual link) */
  tierIndex: number;
  flags: Flag[];
  /** several B rows had this key; the duplicate policy picked `b`, `dropped` are the others */
  dupSource?: boolean;
  dropped?: number[];
  /** several A rows share this key (ID tier / lookup mode) */
  dupBase?: boolean;
  manual?: boolean;
}
export interface Suggestion {
  a: number;
  b: number;
  score: number;
  reasons: string[];
}

export interface JoinResult {
  pairs: MatchPair[];
  perTier: Record<string, number>;
  unmatchedA: number[];
  /** B rows neither matched nor dropped as a duplicate (incl. candidates of unresolved ambiguous rows) */
  unmatchedB: number[];
  /** rows where a key hit several rows on the other side at every tier */
  ambiguous: { a: number; candidatesB: number[]; tier: string; tierIndex: number; sharedBase?: boolean }[];
  /** rows skipped because a higher-priority key conflicted (both IDs present and different) */
  conflicts: { a: number; b: number; tier: string; msg: string }[];
  dupA: [string, number][];
  dupB: [string, number][];
  suggestions: Suggestion[];
  /** B rows dropped by the duplicate policy (represented by the chosen row) */
  absorbedB: number[];
  /** B rows that are candidates of an unresolved ambiguous A row (not joined, not new either) */
  heldB: number[];
  /** manual links that could not be applied */
  manualRejected: { a: number; b: number; reason: string }[];
  /**
   * groups of rows that share ONE key value on both sides in such numbers that looking at the row pairs is not possible
   * (50,000 rows with 10 distinct names): every base row of the group is reported as ambiguous instead
   */
  limited: { tier: string; rows: number; other: number }[];
}

/** work limits of one join (row pairs looked at inside groups of equal keys): a pathological column must never hang the app */
export const GROUP_WORK_CAP = 4_000_000;
export const TOTAL_WORK_CAP = 20_000_000;
const MAX_CONFLICTS = 20_000;

export interface JoinOptions {
  mode?: 'oneToOne' | 'lookup';
  /** several B rows for one A row: 'first' / 'last' row wins (only when they agree on every tier), 'skip' = ambiguous */
  policy?: 'first' | 'last' | 'skip';
  /** forced pairs (row indices), applied before every tier */
  manual?: Array<{ a: number; b: number }>;
  /** header-independent overlaps used as extra consistency checks */
  discovered?: PairStats[];
  /** fuzzy suggestions only for A rows below this index (default: all) */
  suggestLimitA?: number;
  /** compute fuzzy suggestions (default true) */
  suggestions?: boolean;
}

export interface Side {
  table: Table;
  profiles: ColumnProfile[];
}

/**
 * Plan: tiers in priority order. ID first (when both files have one and it is usable), then composite
 * (name+university / name+birthDate / 姓+名 …), then email, name alone, kana alone.
 */
export function planTiers(c: Candidates): Tier[] {
  const tiers: Tier[] = [];
  const usable = (s: PairStats, minRate = 0.1) => Math.max(s.matchRate, s.coverageA) >= minRate;
  // 1. ID family (応募者ID ⟷ 受験番号 ⟷ 学生ID …): must actually overlap (different ID systems → skipped)
  const id = pickIdSingle(c.singles);
  // 1'. default: ID ＋ name together (氏名 / 部品名) is the FIRST tier. A row whose ID agrees but whose name does not (or the other way
  //     round) has two DIFFERENT first-tier keys, so the veto of executeJoin keeps every later tier (ID alone, name alone …) from joining
  //     it: it stays unmatched (reported as a conflict). Rows where the ID or the name is blank on one side still fall back to the rest.
  const idName = id && c.idName && c.idName.a.parts?.[0] === id.a && c.idName.b.parts?.[0] === id.b ? c.idName : undefined;
  if (idName) tiers.push({ name: idName.a.label, a: idName.a, b: idName.b, stats: idName });
  if (id) tiers.push({ name: 'ID', a: id.a, b: id.b, stats: id });
  // 2. composite keys (氏名＋大学, 氏名＋生年月日, 氏名カナ＋大学 …): best one per person-part
  // kanji name before kana: kana collides more (伊藤/井藤), and a name-tier mismatch then vetoes the kana tier
  const ORDER: Record<string, number> = { name: 0, roman: 1, kana: 2 };
  const comps = [...c.composites].filter((s) => s !== idName).sort((p, q) => (ORDER[p.a.sig.split('+')[0]] ?? 9) - (ORDER[q.a.sig.split('+')[0]] ?? 9) || q.score - p.score);
  const seen = new Set<string>();
  for (const s of comps) {
    const person = s.a.sig.split('+')[0];
    if (seen.has(person) || !usable(s) || s.uniqA < 0.9 || s.uniqB < 0.9) continue;
    seen.add(person);
    tiers.push({ name: s.a.label, a: s.a, b: s.b, stats: s });
  }
  // 3. email (only if unique on both sides — shared family addresses exist)
  // (a shared family address is tolerated up to 10% of the rows: base rows sharing a key are never decided by it, see executeJoin)
  const email = c.singles.find((s) => s.a.sig === 'email' && usable(s) && s.uniqA >= 0.9 && s.uniqB >= 0.9);
  if (email) tiers.push({ name: 'メール', a: email.a, b: email.b, stats: email });
  // 3b. phone number (only if unique on both sides — a shared family number must not join two people)
  const phone = c.singles.find((s) => s.a.sig === 'phone' && usable(s) && s.uniqA >= 0.9 && s.uniqB >= 0.9);
  if (phone) tiers.push({ name: '電話番号', a: phone.a, b: phone.b, stats: phone });
  // 4. name alone, 5. kana alone (per-row ambiguity handling protects 同姓同名); Latin-script names in between
  const name = c.singles.find((s) => s.a.sig === 'name' && usable(s));
  if (name) tiers.push({ name: '氏名', a: name.a, b: name.b, stats: name });
  const roman = c.singles.find((s) => s.a.sig === 'roman' && usable(s));
  if (roman) tiers.push({ name: '氏名ローマ字', a: roman.a, b: roman.b, stats: roman });
  const kana = c.singles.find((s) => s.a.sig === 'kana' && usable(s));
  if (kana) tiers.push({ name: '氏名カナ', a: kana.a, b: kana.b, stats: kana });
  // header-independent discovery only when nothing typed was usable — and never a column of PERSON NAMES across different
  // entities (氏名 ⟷ 紹介者 / 担当者: the referrer or the staff member is not the person of the row; IDs may be named
  // differently on purpose: 職場ID ⟷ 応募者コード)
  const entityOf = (x: KeySpec) => x.cols[0]?.entity ?? 'subject';
  const nameLike = (x: KeySpec) => x.cols.some((p) => NAME_TYPES_FOR_ENTITY.has(p.type) || (p.ratios.kanjiName ?? 0) >= 0.5 || (p.ratios.latinName ?? 0) >= 0.5);
  const disc = c.discovered.find((d) => entityOf(d.a) === entityOf(d.b) || !(nameLike(d.a) || nameLike(d.b)));
  if (!tiers.length && disc && usable(disc, 0.3)) tiers.push({ name: disc.a.label, a: disc.a, b: disc.b, stats: disc });
  return tiers;
}

function pushTo<K, V>(m: Map<K, V[]>, k: K, v: V): void {
  const arr = m.get(k);
  if (arr) arr.push(v);
  else m.set(k, [v]);
}

/**
 * Per-row keys of the columns that tell people apart besides the join keys: 氏名, e-mail, phone, birth date, フリガナ, 大学 (the
 * best column of each type). Two rows of one file with the SAME name (or the same フリガナ) but different values in one of
 * them are two people.
 */
export function attributeKeys(ps: ColumnProfile[]): string[][] {
  const out: string[][] = [];
  for (const type of ['name', 'email', 'phone', 'birthDate', 'nameKana', 'university'] as const) {
    const p = ps.filter((x) => x.type === type && x.typeScore >= 0.45 && (x.entity ?? 'subject') === 'subject').sort((x, y) => y.typeScore - x.typeScore)[0];
    if (p) out.push(p.keys);
  }
  return out;
}

/** true when two rows of a file carry different values in one of the attribute columns (both non-blank) */
export function attributesDisagree(attr: string[][], x: number, y: number): boolean {
  for (const ks of attr) {
    const a = ks[x];
    const b = ks[y];
    if (a && b && a !== b) return true;
  }
  return false;
}

export function executeJoin(A: Side, B: Side, tiers: Tier[], opts: JoinOptions = {}): JoinResult {
  const mode = opts.mode ?? 'oneToOne';
  const lookup = mode === 'lookup';
  const policy = opts.policy ?? 'first';
  const nA = A.table.rows.length;
  const nB = B.table.rows.length;
  const usedA = new Set<number>();
  const usedB = new Set<number>();
  const absorbed = new Set<number>();
  const pairs: MatchPair[] = [];
  const perTier: Record<string, number> = {};
  const conflicts: JoinResult['conflicts'] = [];
  const ambigSeen = new Map<number, JoinResult['ambiguous'][number]>();
  const manualRejected: JoinResult['manualRejected'] = [];
  const limited: JoinResult['limited'] = [];
  let attrA: string[][] | null = null;
  let spent = 0;
  let absorbedVersion = 0;

  // ---- manual links (forced; they override every tier) ----
  for (const m of opts.manual ?? []) {
    const bad = (reason: string) => manualRejected.push({ a: m.a, b: m.b, reason });
    if (!Number.isInteger(m.a) || !Number.isInteger(m.b) || m.a < 0 || m.a >= nA || m.b < 0 || m.b >= nB) {
      bad('指定された行がデータ範囲にありません');
      continue;
    }
    if (usedA.has(m.a)) {
      bad('同じ基準の行に複数の手動結合が指定されています');
      continue;
    }
    if (!lookup && usedB.has(m.b)) {
      bad('この行はすでに別の基準の行と手動で結合されています');
      continue;
    }
    pairs.push({ a: m.a, b: m.b, tier: '手動', tierIndex: -1, flags: [], manual: true });
    usedA.add(m.a);
    if (!lookup) usedB.add(m.b);
    perTier['手動'] = (perTier['手動'] ?? 0) + 1;
  }

  // a higher tier whose keys are present on both sides and DIFFER vetoes this pair (ID違い・同名 → 別人扱い)
  const veto = (i: number, cands: number[], t: number): number[] => {
    if (t === 0) return cands;
    const tierName = tiers[t].name;
    return cands.filter((j) => {
      for (let h = 0; h < t; h++) {
        const ka = tiers[h].a.keys[i];
        const kb = tiers[h].b.keys[j];
        if (ka && kb && ka !== kb) {
          if (conflicts.length < MAX_CONFLICTS) conflicts.push({ a: i, b: j, tier: tierName, msg: `${tierName}は一致するが${tiers[h].name}が異なる（${ka.split(SEP).join('＋')} ≠ ${kb.split(SEP).join('＋')}）` });
          return false;
        }
      }
      return true;
    });
  };
  // B rows sharing a key are "the same record twice" only when no tier key disagrees between them
  const agree = (cands: number[]): boolean => {
    for (const tr of tiers) {
      let seen = '';
      for (const j of cands) {
        const k = tr.b.keys[j];
        if (!k) continue;
        if (!seen) seen = k;
        else if (k !== seen) return false;
      }
    }
    return true;
  };
  // several rows with the same NAME are a resubmission only when nothing else (e-mail, phone, birth date …) tells them apart
  let attrB: string[][] | null = null;
  const distinctPeople = (cands: number[]): boolean => {
    attrB ??= attributeKeys(B.profiles);
    for (let x = 1; x < cands.length; x++) if (attributesDisagree(attrB, cands[0], cands[x])) return true;
    return false;
  };
  const resolve = (cands: number[], person: boolean): { b: number; dropped: number[] } | null => {
    if (cands.length === 1) return { b: cands[0], dropped: [] };
    if (policy === 'skip' || !agree(cands) || (person && distinctPeople(cands))) return null;
    let b = cands[0];
    for (const j of cands) if (policy === 'last' ? j > b : j < b) b = j;
    return { b, dropped: cands.filter((j) => j !== b) };
  };

  for (let t = 0; t < tiers.length; t++) {
    const tier = tiers[t];
    // the default ID＋名前 key carries an ID: it is decided like an ID (several base rows of one person all receive the row)
    const isIdTier = tier.a.sig === 'idFamily' || tier.a.sig.startsWith('idFamily+');
    // a key that names a PERSON (not a workplace / organisation / discovered foreign key): rows sharing it may be different people
    const isPersonKey = /^(name|roman|kana|email|phone)(\+|$)/.test(tier.a.sig);
    perTier[tier.name] = perTier[tier.name] ?? 0;
    const bk = tier.b.keys;
    const ak = tier.a.keys;
    const idx = new Map<string, number[]>();
    for (let j = 0; j < nB; j++) {
      const k = bk[j];
      if (k && !absorbed.has(j) && (lookup || !usedB.has(j))) pushTo(idx, k, j);
    }
    const groups = new Map<string, number[]>();
    for (let i = 0; i < nA; i++) {
      if (usedA.has(i)) continue;
      const k = ak[i];
      if (k && idx.has(k)) pushTo(groups, k, i);
    }
    for (const [k, rowsA] of groups) {
      const all = idx.get(k)!;
      const shared = rowsA.length > 1;
      const work = rowsA.length * all.length;
      if (shared && all.length > 1 && (work > GROUP_WORK_CAP || spent + work > TOTAL_WORK_CAP)) {
        // thousands of rows on BOTH sides carry this one value: nothing can be decided from it, and listing the row pairs
        // would take minutes and gigabytes — the whole group is reported as ambiguous (one shared candidate list)
        const cands = all.filter((j) => !absorbed.has(j) && (lookup || !usedB.has(j)));
        limited.push({ tier: tier.name, rows: rowsA.length, other: cands.length });
        for (const i of rowsA) ambigSeen.set(i, { a: i, candidatesB: cands, tier: tier.name, tierIndex: t, sharedBase: true });
        continue;
      }
      spent += work;
      // a VLOOKUP-style join lets several base rows share one row — unless the key is not an ID and those base rows are
      // different people (イトウ ハナコ = 伊藤 花子 and 井藤 花子; the same address for two brothers)
      let baseIsNotOnePerson = false;
      if (shared && lookup && isPersonKey) {
        attrA ??= attributeKeys(A.profiles);
        for (let x = 1; x < rowsA.length && !baseIsNotOnePerson; x++) baseIsNotOnePerson = attributesDisagree(attrA, rowsA[0], rowsA[x]);
      }
      if (shared && ((!lookup && (!isIdTier || policy === 'skip')) || baseIsNotOnePerson)) {
        // 同姓同名 protection: several base rows share this (non-ID) key → not decided at this tier
        // (an ID shared by several base rows is decided below, unless duplicates must not be joined at all)
        const free = all.filter((j) => !usedB.has(j) && !absorbed.has(j));
        for (const i of rowsA) {
          const c = veto(i, free, t);
          if (c.length) ambigSeen.set(i, { a: i, candidatesB: c, tier: tier.name, tierIndex: t, sharedBase: true });
        }
        continue;
      }
      const pendingUsed: number[] = [];
      let avail: number[] = [];
      let availVersion = -1;
      for (const i of rowsA) {
        if (availVersion !== absorbedVersion) {
          avail = all.filter((j) => !absorbed.has(j) && (lookup || !usedB.has(j)));
          availVersion = absorbedVersion;
        }
        const cands = veto(i, avail, t);
        if (!cands.length) continue;
        const r = resolve(cands, isPersonKey);
        if (!r) {
          ambigSeen.set(i, { a: i, candidatesB: cands, tier: tier.name, tierIndex: t });
          continue;
        }
        pairs.push({ a: i, b: r.b, tier: tier.name, tierIndex: t, flags: [], dupSource: r.dropped.length > 0, dropped: r.dropped.length ? r.dropped : undefined, dupBase: shared });
        usedA.add(i);
        perTier[tier.name]++;
        ambigSeen.delete(i);
        if (r.dropped.length) {
          for (const j of r.dropped) absorbed.add(j);
          absorbedVersion++;
        }
        if (!lookup) {
          pendingUsed.push(r.b);
          for (const j of r.dropped) pendingUsed.push(j);
        }
      }
      // an ID shared by several base rows: all of them received the row; mark it used only now
      for (const j of pendingUsed) usedB.add(j);
    }
  }
  const ambiguous = [...ambigSeen.values()].filter((x) => !usedA.has(x.a));
  const matchedB = new Set(pairs.map((p) => p.b));
  for (const j of absorbed) if (matchedB.has(j)) absorbed.delete(j);
  const held = new Set<number>();
  const seenLists = new Set<number[]>(); // a pathological group shares ONE candidate list between thousands of rows
  for (const x of ambiguous) {
    if (seenLists.has(x.candidatesB)) continue;
    seenLists.add(x.candidatesB);
    for (const j of x.candidatesB) if (!matchedB.has(j) && !absorbed.has(j)) held.add(j);
  }
  const unmatchedA: number[] = [];
  for (let i = 0; i < nA; i++) if (!usedA.has(i)) unmatchedA.push(i);
  const unmatchedB: number[] = [];
  for (let j = 0; j < nB; j++) if (!matchedB.has(j) && !absorbed.has(j)) unmatchedB.push(j);

  // ---- cross-checks on matched pairs ----
  crossCheck(A, B, tiers, pairs, opts.discovered ?? []);

  const limitA = opts.suggestLimitA ?? nA;
  const suggestions = opts.suggestions === false || !A.profiles.length || !B.profiles.length ? [] : suggest(
    A.profiles, B.profiles,
    unmatchedA.filter((i) => i < limitA),
    unmatchedB.filter((j) => lookup || !usedB.has(j)),
    tiers.find((t) => t.a.sig === 'idFamily'),
  );
  const dupA = tiers[0] ? uniq(tiers[0].a.keys.slice(0, limitA)).dupKeys : [];
  const dupB = tiers[0] ? uniq(tiers[0].b.keys).dupKeys : [];
  return { pairs, perTier, unmatchedA, unmatchedB, ambiguous, conflicts, dupA, dupB, suggestions, absorbedB: [...absorbed], heldB: [...held], manualRejected, limited };
}

function crossCheck(A: Side, B: Side, tiers: Tier[], pairs: MatchPair[], discovered: PairStats[]): void {
  if (!A.profiles.length || !B.profiles.length) return;
  // only compare columns that describe the SAME entity as the key (学生名 must not be checked against a 職場ID match)
  const keyEntity = (side: 'a' | 'b') => tiers[0]?.[side].cols[0]?.entity ?? 'subject';
  const sameEntity = (p: ColumnProfile, side: 'a' | 'b') => (p.entity ?? keyEntity(side)) === keyEntity(side);
  const colOf = (ps: ColumnProfile[], side: 'a' | 'b', ...types: KeyType[]) =>
    ps.filter((p) => types.includes(p.type as KeyType) && sameEntity(p, side)).sort((a, b) => b.typeScore - a.typeScore)[0];
  const nameA = colOf(A.profiles, 'a', 'name');
  const nameB = colOf(B.profiles, 'b', 'name');
  const kanaA = colOf(A.profiles, 'a', 'nameKana');
  const kanaB = colOf(B.profiles, 'b', 'nameKana');
  const bdA = colOf(A.profiles, 'a', 'birthDate');
  const bdB = colOf(B.profiles, 'b', 'birthDate');
  const mailA = colOf(A.profiles, 'a', 'email');
  const mailB = colOf(B.profiles, 'b', 'email');
  const phoneA = colOf(A.profiles, 'a', 'phone');
  const phoneB = colOf(B.profiles, 'b', 'phone');
  const univA = colOf(A.profiles, 'a', 'university');
  const univB = colOf(B.profiles, 'b', 'university');
  const tA = A.table;
  const tB = B.table;
  const ida = tiers.find((t) => t.a.sig === 'idFamily');
  const checks = discovered.filter((d) => !tiers.some((t) => t.a.cols[0] === d.a.cols[0] || t.b.cols[0] === d.b.cols[0]) && Math.max(d.matchRate, d.coverageA) >= 0.8);
  const nameMemo = new Map<unknown, NameKey>();
  const normMemo = (v: unknown): NameKey => {
    let k = nameMemo.get(v);
    if (!k) {
      k = normName(v as never);
      if (nameMemo.size < 200_000) nameMemo.set(v, k);
    }
    return k;
  };
  for (const p of pairs) {
    const ra = tA.rows[p.a];
    const rb = tB.rows[p.b];
    if (!ra || !rb) continue;
    if (nameA && nameB) {
      const f = compareNames(ra[nameA.col], rb[nameB.col], normMemo);
      if (f.length) p.flags.push(...f);
    }
    if (kanaA && kanaB) {
      const x = normKana(ra[kanaA.col]);
      const y = normKana(rb[kanaB.col]);
      if (x.strict && y.strict && x.strict !== y.strict) {
        p.flags.push(x.loose === y.loose ? { level: 'info', msg: `カナ表記ゆれ（${x.strict}/${y.strict}）` } : { level: 'warn', msg: `フリガナ不一致（${x.strict}/${y.strict}）` });
      }
    }
    if (bdA && bdB) {
      const x = bdA.keys[p.a];
      const y = bdB.keys[p.b];
      if (x && y && x !== y) p.flags.push({ level: 'warn', msg: `生年月日不一致（${x}/${y}）` });
    }
    // a match made by the NAME alone (氏名 / ローマ字 / フリガナ, no ID, not a composite) that disagrees on e-mail, phone or 大学
    // may be two people of the same name: those differences are worth checking (an ID match just means a second address)
    const tier = p.tierIndex >= 0 ? tiers[p.tierIndex] : undefined;
    const nameOnly = !!tier && (tier.a.sig === 'name' || tier.a.sig === 'roman' || tier.a.sig === 'kana');
    if (mailA && mailB) {
      const x = mailA.keys[p.a];
      const y = mailB.keys[p.b];
      if (x && y && x !== y) p.flags.push({ level: nameOnly ? 'warn' : 'info', msg: `メールが異なる（${x}/${y}）` });
    }
    if (nameOnly && phoneA && phoneB) {
      const x = phoneA.keys[p.a];
      const y = phoneB.keys[p.b];
      if (x && y && x !== y) p.flags.push({ level: 'warn', msg: `電話番号が異なる（${x}/${y}）` });
    }
    if (nameOnly && univA && univB) {
      const x = univA.keys[p.a];
      const y = univB.keys[p.b];
      if (x && y && x !== y) p.flags.push({ level: 'warn', msg: `大学が異なる（${x}/${y}）` });
    }
    // other header-independent overlaps between the same two tables act as extra checks (職場名 ⟷ 漢字氏名)
    for (const d of checks) {
      const x = d.a.keys[p.a];
      const y = d.b.keys[p.b];
      if (x && y && x !== y) p.flags.push({ level: 'warn', msg: `「${d.a.cols[0].header}」と「${d.b.cols[0].header}」が異なる（${x}/${y}）` });
    }
    if (ida && ida.a.cols.length === 1 && ida.b.cols.length === 1) {
      const ca = ida.a.cols[0].col;
      const cb = ida.b.cols[0].col;
      const x = normId(ra[ca], typeof ra[ca] === 'number' ? tA.display(p.a, ca) : undefined);
      const y = normId(rb[cb], typeof rb[cb] === 'number' ? tB.display(p.b, cb) : undefined);
      if (x.exact && y.exact && x.exact !== y.exact) {
        // a number that is only the tail of the other one is weaker evidence than a formatting difference: check it
        if (ida.a.variant === 'suffix') p.flags.push({ level: 'warn', msg: `IDの桁数が異なる（${x.exact} / ${y.exact}）` });
        else p.flags.push({ level: 'info', msg: `ID表記の違いを補正（${x.exact} / ${y.exact}）` });
      }
    }
  }
}

/** 斎/斉 group etc.: DIFFERENT characters that people confuse — used for suggestions only, never auto-merge. */
const CONFUSABLE: Record<string, string> = {
  '斉': '斎', '齊': '斎', '齋': '斎', '邉': '辺', '邊': '辺', '籐': '藤', '澤': '沢', '嶋': '島', '嶌': '島', '之': 'ノ', 'の': 'ノ', '乃': 'ノ',
  'ヶ': 'ケ', 'ヵ': 'カ', '桧': '檜', '舘': '館', '冴': '冱', '淵': '渕', '栁': '柳', '槇': '槙',
  '鴎': '鷗', '峰': '峯', '富': '冨', '条': '條', '藪': '薮', '曽': '曾', '恵': '惠',
};
export function confusableFold(s: string): string {
  let o = '';
  // 山田さくら ≡ 山田サクラ: hiragana and katakana of a name are the same reading — a suggestion, never an automatic match
  for (const c of hiraToKata(s)) o += CONFUSABLE[c] ?? c;
  return o;
}

export function compareNames(a: unknown, b: unknown, norm: (v: unknown) => NameKey = (v) => normName(v as never)): Flag[] {
  if (a === b) return [];
  const x = norm(a);
  const y = norm(b);
  if (!x.strict || !y.strict || x.strict === y.strict) return [];
  if (x.folded === y.folded) return [{ level: 'info', msg: `異体字の違い（${x.strict}/${y.strict}）` }];
  if (x.latinSorted && x.latinSorted === y.latinSorted) return [{ level: 'info', msg: `姓名の順序違い（${a}/${b}）` }];
  if (confusableFold(x.folded) === confusableFold(y.folded)) return [{ level: 'warn', msg: `似た字の違い（${x.strict}/${y.strict}）` }];
  if (distance(x.folded, y.folded) === 1 && Math.min(x.folded.length, y.folded.length) >= 3) return [{ level: 'warn', msg: `氏名1文字違い（${x.strict}/${y.strict}）` }];
  return [{ level: 'error', msg: `ID一致・氏名不一致（${a} / ${b}）` }];
}

/** Per-row person attributes used by suggestions (split 姓/名 and セイ/メイ are concatenated). */
interface PersonRow {
  r: number;
  nameF: string;
  /** the name under the maiden surname given in a note (山田 花子（旧姓：佐藤） → 佐藤花子), '' when there is none */
  nameAlt: string;
  nameConf: string;
  nameLatin: string;
  kana: string;
  kanaLoose: string;
  mail: string;
  /** emailLoose(mail), precomputed */
  mailLoose: string;
  univ: string;
  bd: string;
  phone: string;
  id: string;
}

function personRows(ps: ColumnProfile[], rows: number[], idSpec?: KeySpec): PersonRow[] {
  if (!ps.length) return [];
  const t = ps[0].table;
  const best = (...types: KeyType[]) => ps.filter((p) => types.includes(p.type as KeyType)).sort((a, b) => b.typeScore - a.typeScore)[0];
  const specs = specsOf(ps);
  const nameSpec = specs.filter((s) => s.sig === 'name' || s.sig === 'roman')
    .sort((a, b) => (a.sig === 'name' ? 0 : 1) - (b.sig === 'name' ? 0 : 1) || b.cols[0].typeScore - a.cols[0].typeScore)[0];
  const kanaSpec = specs.filter((s) => s.sig === 'kana').sort((a, b) => b.cols[0].typeScore - a.cols[0].typeScore)[0];
  const e = best('email');
  const u = best('university');
  const d = best('birthDate');
  const ph = best('phone');
  return rows.map((r) => {
    const n = nameSpec?.cols.length === 1 ? normName(t.rows[r][nameSpec.cols[0].col]) : null;
    const kanaStr = kanaSpec ? kanaSpec.keys[r] : '';
    const nameF = nameSpec ? nameSpec.keys[r] : '';
    const mail = e ? e.keys[r] : '';
    return {
      r, nameF, nameAlt: n?.alt ?? '', nameConf: nameF ? confusableFold(nameF) : '', nameLatin: n?.latinSorted ?? '',
      kana: kanaStr, kanaLoose: kanaStr ? kanaLoose(kanaStr) : '',
      mail, mailLoose: mail ? emailLoose(mail) : '',
      univ: u ? u.keys[r] : '', bd: d ? d.keys[r] : '', phone: ph ? ph.keys[r] : '', id: idSpec ? idSpec.keys[r] : '',
    };
  });
}

/** above this many A × B comparisons the suggestion search is blocked (candidate pairs must share a blocking key) */
const FULL_SCAN_LIMIT = 1_000_000;
/** maximum number of row comparisons for suggestions per pair (a full scan always fits; ≈ 1.5 s at most) */
const SUGGEST_BUDGET = 1_500_000;

function blockKeys(p: PersonRow): string[] {
  const out: string[] = [];
  if (p.kanaLoose) out.push('k' + p.kanaLoose);
  if (p.mailLoose) out.push('m' + p.mailLoose);
  if (p.phone) out.push('p' + p.phone);
  if (p.nameLatin) out.push('l' + p.nameLatin);
  if (p.nameF) {
    // equal / confusable names share the folded form; Levenshtein-1 names share the name itself or one of its
    // single-character deletions (substitution: same deletion; insertion/deletion: one is a deletion of the other)
    out.push('c' + p.nameConf);
    if (p.nameAlt) out.push('c' + confusableFold(p.nameAlt));
    const n = [...p.nameF];
    if (n.length >= 3 && n.length <= 30) {
      const seen = new Set<string>([p.nameF]);
      out.push('d' + p.nameF);
      for (let i = 0; i < n.length; i++) {
        const d = n.slice(0, i).join('') + n.slice(i + 1).join('');
        if (!seen.has(d)) {
          seen.add(d);
          out.push('d' + d);
        }
      }
    }
  }
  if (p.id && p.id.length >= 3 && p.id.length <= 24) {
    // single-deletion neighbourhood: one substitution or one adjacent transposition leaves a common deletion
    const seen = new Set<string>();
    for (let i = 0; i < p.id.length; i++) {
      const d = p.id.slice(0, i) + p.id.slice(i + 1);
      if (!seen.has(d)) {
        seen.add(d);
        out.push(`i${p.id.length}:` + d);
      }
    }
  }
  return out;
}

/**
 * Fuzzy suggestions for rows left unmatched (NEVER applied automatically; the user accepts them).
 * score = strongest positive signal + agreements − disagreements (Fellegi–Sunter-like, hand-tuned weights):
 *   positive (max):  0.90 same kana · 0.88 same name (異体字 folded) · 0.85 same email (gmail dots/+tag ignored)
 *                    0.85 same Latin name tokens (order ignored) · 0.80 confusable glyphs (斎/斉, 辺, ノ/之) · 0.80 kana loose
 *                    0.75 same phone · 0.70 name Levenshtein = 1 (len ≥ 3) · 0.55 ID one typo (Damerau-1, len ≥ 3)
 *   agreements:      +0.05 each: university, birthDate, email, phone, ID one typo
 *   disagreements:   −0.35 both IDs present & different (not a one-typo) · −0.15 birthDate differs · −0.10 university differs
 *   shown when ≥ 0.60, top 3 per base row, one B row may appear under several A rows.
 */
export function suggest(A: ColumnProfile[], B: ColumnProfile[], ua: number[], ub: number[], idTier?: Tier): Suggestion[] {
  if (!ua.length || !ub.length) return [];
  // IDs are only comparable when the two ID columns belong to the same numbering system (= an ID tier exists)
  const xs = personRows(A, ua, idTier?.a);
  const ys = personRows(B, ub, idTier?.b);
  let candidatesOf: (x: PersonRow) => Iterable<PersonRow>;
  if (xs.length * ys.length <= FULL_SCAN_LIMIT) {
    candidatesOf = () => ys;
  } else {
    const index = new Map<string, PersonRow[]>();
    for (const y of ys) for (const k of blockKeys(y)) pushTo(index, k, y);
    candidatesOf = (x) => {
      const set = new Set<PersonRow>();
      for (const k of blockKeys(x)) {
        const hit = index.get(k);
        if (hit && hit.length <= 2000) for (const y of hit) set.add(y);
      }
      return set;
    };
  }
  const out: Suggestion[] = [];
  // hard cap on comparisons so a pathological file (thousands of 同姓同名) cannot stall the UI
  let budget = SUGGEST_BUDGET;
  for (const x of xs) {
    if (budget <= 0) break;
    const cands: Suggestion[] = [];
    for (const y of candidatesOf(x)) {
      budget--;
      // strongest positive signal (no closures in this hot loop)
      let s = 0;
      let why: string[] | null = null;
      if (x.kana && y.kana) {
        if (x.kana === y.kana) {
          s = 0.9;
          (why ??= []).push('フリガナ一致');
        } else if (x.kanaLoose && x.kanaLoose === y.kanaLoose) {
          s = 0.8;
          (why ??= []).push('フリガナほぼ一致');
        }
      }
      if (x.mailLoose && x.mailLoose === y.mailLoose) {
        if (s < 0.85) s = 0.85;
        (why ??= []).push('メール一致');
      }
      if (x.phone && x.phone === y.phone) {
        if (s < 0.75) s = 0.75;
        (why ??= []).push('電話一致');
      }
      if (x.nameF && y.nameF) {
        if (x.nameF === y.nameF) {
          if (s < 0.88) s = 0.88;
          (why ??= []).push('氏名一致');
        } else if (x.nameLatin && x.nameLatin === y.nameLatin) {
          if (s < 0.85) s = 0.85;
          (why ??= []).push('姓名順違い');
        } else if (x.nameConf === y.nameConf) {
          if (s < 0.8) s = 0.8;
          (why ??= []).push('似た字（斎/斉・ノ/之など）');
        } else if (Math.min(x.nameF.length, y.nameF.length) >= 3 && Math.abs(x.nameF.length - y.nameF.length) <= 1 && distance(x.nameF, y.nameF) === 1) {
          if (s < 0.7) s = 0.7;
          (why ??= []).push('氏名1文字違い');
        }
      }
      // 山田 花子（旧姓：佐藤）⟷ 佐藤 花子: the same person under the other surname
      if ((x.nameAlt && (x.nameAlt === y.nameF || x.nameAlt === y.nameAlt)) || (y.nameAlt && y.nameAlt === x.nameF)) {
        if (s < 0.8) s = 0.8;
        (why ??= []).push('旧姓で一致');
      }
      const idTypo = !!(x.id && y.id && x.id !== y.id && x.id.length >= 3 && x.id.length === y.id.length && damerau1(x.id, y.id));
      if (idTypo) {
        if (s < 0.55) s = 0.55;
        (why ??= []).push('ID1文字違い');
      }
      if (!s || !why) continue;
      if (x.id && y.id && x.id !== y.id && !idTypo) {
        s -= 0.35;
        why.push('IDが異なる');
      }
      if (idTypo && s > 0.55) s += 0.05;
      if (x.univ && y.univ) {
        if (x.univ === y.univ) {
          s += 0.05;
          why.push('大学一致');
        } else {
          s -= 0.1;
          why.push('大学が異なる');
        }
      }
      if (x.bd && y.bd) {
        if (x.bd === y.bd) {
          s += 0.05;
          why.push('生年月日一致');
        } else {
          s -= 0.15;
          why.push('生年月日が異なる');
        }
      }
      if (x.phone && y.phone && x.phone === y.phone && s > 0.75) s += 0.05;
      cands.push({ a: x.r, b: y.r, score: Math.round(Math.min(0.99, s) * 1000) / 1000, reasons: why });
    }
    cands.sort((p, q) => q.score - p.score || p.b - q.b);
    out.push(...cands.slice(0, 3).filter((c) => c.score >= 0.6));
  }
  return out.sort((p, q) => q.score - p.score || p.a - q.a);
}

function damerau1(a: string, b: string): boolean {
  if (distance(a, b) === 1) return true;
  if (a.length !== b.length) return false;
  const d: number[] = [];
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) d.push(i);
  return d.length === 2 && d[1] === d[0] + 1 && a[d[0]] === b[d[1]] && a[d[1]] === b[d[0]];
}

export { foldItaiji };

/**
 * Join mode from the first tier: base key repeats but the other side's key is unique → 'lookup'
 * (VLOOKUP semantics: many base rows receive the same other row, e.g. 職場ID → 職場回答). Otherwise 1:1.
 * A person list with a couple of duplicates is NOT a lookup: repeated rows must make up ≥ 30% of the base
 * (otherwise one duplicate in a 9-row list would let 同姓同名 rows share a name match).
 */
export function chooseMode(tiers: Tier[]): 'oneToOne' | 'lookup' {
  if (!tiers.length) return 'oneToOne';
  const ua = uniq(tiers[0].a.keys);
  const ub = uniq(tiers[0].b.keys);
  let nonBlankA = 0;
  for (const k of tiers[0].a.keys) if (k) nonBlankA++;
  const dupShare = nonBlankA ? ua.dupRows / nonBlankA : 0;
  return ua.uniqueness < 0.9 && ub.uniqueness >= 0.99 && dupShare >= 0.3 ? 'lookup' : 'oneToOne';
}
