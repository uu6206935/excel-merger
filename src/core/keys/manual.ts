/**
 * Manual key candidates (addManualCandidate in multi.ts): the user picks 1..3 columns of the base sheet and the same
 * number of columns of ONE other file. This file validates the selection and builds the candidate — KeySpec sides,
 * real PairStats, the contract DTO — with the same machinery the auto candidates use (singleSpec-like parts,
 * compositeSpec, pairStats, candidateDTO). Everything here is pure; callers run it inside withNormalizeOptions.
 */
import type { CandidateDTO } from '../contracts';
import { type CandidateEntry, NAME_TYPES, candidateDTO } from './api';
import { type ColumnProfile, type KeyType, FAMILY, ID_TYPES, keysAs } from './detect';
import { type KeySpec, compositeSpec, pairStats } from './match';
import { colLetter } from './region';
import type { Table } from './table';

export const MAX_MANUAL_COLUMNS = 3;

const ARROW = ' ⟷ ';

/** Candidate id of a manual selection: 'manual:<base cols joined by +>|<other cols joined by +>' (sheet columns, user order). */
export function manualId(baseCols: number[], otherCols: number[]): string {
  return `manual:${baseCols.join('+')}|${otherCols.join('+')}`;
}

/**
 * Order-independent identity of a column pairing, to recognise a candidate that already exists: [氏名, 大学] ⟷ [氏名, 大学]
 * and [大学, 氏名] ⟷ [大学, 氏名] are the same key; [氏名, 大学] ⟷ [大学, 氏名] is a different one. Candidates whose sides
 * have different column counts (氏名 ⟷ 姓＋名) have no pairing and compare by their column sets.
 */
export function pairingKey(baseCols: number[], otherCols: number[]): string {
  if (baseCols.length === otherCols.length) return 'p:' + baseCols.map((c, i) => `${c}:${otherCols[i]}`).sort().join(',');
  const s = (xs: number[]) => [...xs].sort((x, y) => x - y).join('+');
  return `s:${s(baseCols)}|${s(otherCols)}`;
}

/** pairingKey of an existing candidate DTO */
export function dtoPairingKey(c: CandidateDTO): string {
  return pairingKey(c.baseColumns.map((x) => x.col), c.otherColumns.map((x) => x.col));
}

// ---------------------------------------------------------------------------
// validation
// ---------------------------------------------------------------------------

function resolveSide(table: Table, profiles: ColumnProfile[], cols: number[], who: string): ColumnProfile[] {
  if (!cols.every((c) => typeof c === 'number' && Number.isInteger(c) && c >= 1)) {
    throw new Error(`${who}の列の指定が正しくありません。列を選び直してください。`);
  }
  if (!table.cols.length) throw new Error(`${who}には結合キーに使える列がありません（シート・見出し行を確認してください）。`);
  const first = table.cols[0];
  const last = table.cols[table.cols.length - 1];
  const out: ColumnProfile[] = [];
  const seen = new Set<number>();
  for (const c of cols) {
    const idx = table.cols.indexOf(c);
    if (idx < 0 || !profiles[idx]) {
      throw new Error(`${who}の列${colLetter(c)}は、読み込んだ表の範囲（${colLetter(first)}〜${colLetter(last)}列）の外にあります。`);
    }
    if (seen.has(c)) throw new Error(`${who}で同じ列（列${colLetter(c)}）が重複して選ばれています。`);
    seen.add(c);
    out.push(profiles[idx]);
  }
  return out;
}

/** The chosen profiles of both sides, or an Error with a Japanese message. */
export function resolveManualColumns(
  base: { table: Table; profiles: ColumnProfile[] },
  other: { table: Table; profiles: ColumnProfile[]; label: string },
  baseCols: number[],
  otherCols: number[],
): { a: ColumnProfile[]; b: ColumnProfile[] } {
  const who = `「${other.label}」`;
  if (!Array.isArray(baseCols) || !Array.isArray(otherCols)) throw new Error('結合に使う列を選んでください。');
  if (!baseCols.length || !otherCols.length) {
    throw new Error(`結合に使う列を、基準ファイルと${who}の両方で1列以上選んでください。`);
  }
  if (baseCols.length > MAX_MANUAL_COLUMNS || otherCols.length > MAX_MANUAL_COLUMNS) {
    throw new Error(`結合キーに使える列は最大${MAX_MANUAL_COLUMNS}列までです（選択: 基準ファイル ${baseCols.length}列 / ${who} ${otherCols.length}列）。`);
  }
  if (baseCols.length !== otherCols.length) {
    throw new Error(`基準ファイルの列数（${baseCols.length}）と${who}の列数（${otherCols.length}）が違います。同じ数だけ選んでください。`);
  }
  return {
    a: resolveSide(base.table, base.profiles, baseCols, '基準ファイル'),
    b: resolveSide(other.table, other.profiles, otherCols, who),
  };
}

// ---------------------------------------------------------------------------
// candidate
// ---------------------------------------------------------------------------

/** joinable group of a detected type ('idFamily' also takes 組織コード: a manual ID-like pair joins like an ID) */
function groupOfType(t: KeyType | 'unknown'): string {
  if (t === 'unknown') return 'unknown';
  if (ID_TYPES.has(t) || t === 'orgCode') return 'idFamily';
  return FAMILY[t] ?? t;
}

const isIdType = (t: string) => ID_TYPES.has(t) || t === 'orgCode';

interface Part {
  a: KeySpec;
  b: KeySpec;
  family: CandidateDTO['family'];
}

/** One column pair normalised as ta (base column) / tb (other column); sig 'manual' = no common type */
function buildPart(p: ColumnProfile, q: ColumnProfile, ta: KeyType | 'unknown', tb: KeyType | 'unknown'): Part {
  const g = groupOfType(ta);
  const sig = g === 'unknown' ? 'manual' : g;
  const spec = (x: ColumnProfile, t: KeyType | 'unknown'): KeySpec => ({ label: x.header, sig, cols: [x], keys: keysAs(x, t), kind: 'single', keyType: t });
  return { a: spec(p, ta), b: spec(q, tb), family: partFamily(p, q, ta, tb) };
}

/**
 * One column pair. Both sides must be normalised the same way for the keys to be comparable:
 *  - the same type family (電話 ⟷ 電話, ID ⟷ 学籍番号, both unrecognised): each column's own keys;
 *  - different types (the header / values of one column were misread: 電話 ⟷ 連絡先 typed as an ID): the detector cannot
 *    know which one is right, so each normaliser the two types suggest — and the generic text key — is tried on both
 *    sides and the best matching one is kept (ties → the base column's type first).
 */
function manualPart(p: ColumnProfile, q: ColumnProfile): Part {
  if (groupOfType(p.type) === groupOfType(q.type)) return buildPart(p, q, p.type, q.type);
  const types: Array<KeyType | 'unknown'> = [];
  for (const t of [p.type, q.type, 'unknown' as const]) if (!types.includes(t)) types.push(t);
  let best: { part: Part; quality: number } | null = null;
  for (const t of types) {
    const part = buildPart(p, q, t, t);
    const st = pairStats(part.a, part.b);
    const quality = Math.max(st.matchRate, st.coverageA);
    if (!best || quality > best.quality + 1e-9) best = { part, quality };
  }
  return best!.part;
}

function partFamily(p: ColumnProfile, q: ColumnProfile, ta: KeyType | 'unknown', tb: KeyType | 'unknown'): CandidateDTO['family'] {
  if (ta !== 'unknown' || tb !== 'unknown') {
    const t = ta !== 'unknown' ? ta : tb;
    if (isIdType(t)) return 'id';
    if (NAME_TYPES.has(t)) return 'name';
    return 'other';
  }
  // generic keys: decided like the value-overlap candidates; ID-shaped values only make an ID when one side is (nearly)
  // unique — scores or codes that repeat (10, 20, 10 …) are not identifiers
  const ps = [p, q];
  if (ps.some((x) => isIdType(x.type))) return 'id';
  if (ps.some((x) => NAME_TYPES.has(x.type))) return 'name';
  if (ps.every((x) => (x.ratios.idShape ?? 0) >= 0.8) && ps.some((x) => x.lightUniq.uniqueness >= 0.9)) return 'id';
  return 'other';
}

/** family of a whole candidate: any ID part → 'id', else any name part → 'name' (composites like 氏名＋大学), else 'other' */
function joinFamilies(fs: CandidateDTO['family'][]): CandidateDTO['family'] {
  if (fs.includes('id')) return 'id';
  if (fs.includes('name')) return 'name';
  return 'other';
}

/**
 * The manual candidate for already validated profiles (user order). `baseCols` / `otherCols` are the sheet columns, for the
 * id and for qualifying a label that another candidate of the pair already uses. Must run inside withNormalizeOptions.
 */
export function buildManualEntry(
  tA: Table,
  tB: Table,
  a: ColumnProfile[],
  b: ColumnProfile[],
  baseCols: number[],
  otherCols: number[],
  usedLabels: ReadonlySet<string>,
): CandidateEntry {
  const parts = a.map((p, i) => manualPart(p, b[i]));
  const headers = (xs: ColumnProfile[]) => xs.map((x) => x.header).join('＋');
  let label = `${headers(a)}${ARROW}${headers(b)}`;
  if (usedLabels.has(label)) label = `${label}（${baseCols.map(colLetter).join('＋')}${ARROW}${otherCols.map(colLetter).join('＋')}）`;
  for (let n = 2; usedLabels.has(label); n++) label = `${label}（${n}）`;
  const single = parts.length === 1;
  const sa = single ? { ...parts[0].a, label } : { ...compositeSpec(parts.map((x) => x.a)), label };
  const sb = single ? { ...parts[0].b, label } : { ...compositeSpec(parts.map((x) => x.b)), label };
  const stats = pairStats(sa, sb);
  const dto = candidateDTO(
    { id: manualId(baseCols, otherCols), kind: single ? 'single' : 'composite', label, family: joinFamilies(parts.map((x) => x.family)), s: stats, tier: -1 },
    tA,
    tB,
  );
  return { dto, stats };
}
