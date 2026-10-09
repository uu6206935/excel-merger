/**
 * keys facade (see ./index.ts): data sheet / region detection, pairwise key analysis base ⟷ every other file, and the
 * multi-file join (tiers in priority order, duplicate policy, manual links, sequential 'full' join).
 *
 * Join semantics (per other file, in the order the files were given to analyzeSources):
 *   1. manual links are applied first and override every tier;
 *   2. tiers are tried in plan order; a higher tier whose keys are present on both sides and differ vetoes the pair;
 *   3. several other-file rows on one base row: the duplicate policy picks the first / last one when those rows do not
 *      disagree on any tier key (else, and for 'skip', the base row is reported as ambiguous);
 *      several base rows sharing an ID all receive the row (VLOOKUP semantics, flag duplicateInBase); several base rows
 *      sharing a NON-ID key (同姓同名) are never decided by that tier;
 *   4. joinType 'full': rows of the other file that matched nothing are appended (duplicates among them collapsed by
 *      the policy); later files are matched against base rows AND rows appended from earlier files, through the base
 *      key columns filled from each appended row's own file (keyFill).
 *   JoinResult.keyPairs (every join, also 'left'): per source the direct (base column, source column) pairs of all enabled
 *   tiers that matched a row — see keyPairsOf.
 *
 * CheckItem codes (JoinResult.items), severity in brackets:
 *   emptySource [error for the base / warning], layout (横持ち) [warning], noCachedValue [warning], lossyId [warning],
 *   noKey [error] (a file with no tier and no manual link — plan.tiers[id] is empty: it contributes no matches, appends no
 *   rows and gets no unmatched / unmatchedOther items; its columns still appear in defaultImports, the UI excludes it),
 *   planInvalid (unknown candidate id in plan.tiers) [warning], manualInvalid [warning],
 *   unmatched (未一致: base rows without a match) [warning],
 *   unmatchedOther (other-file-only rows) [info when appended / warning for 'left'], duplicate (重複 in the other file)
 *   [warning], duplicateBase (重複 in the base) [warning], lookup (VLOOKUP-style pair) [info], ambiguous (候補が複数)
 *   [warning], conflict (ID一致・氏名不一致) [error], checkWarn (似た字・1文字違い・フリガナ/生年月日不一致) [warning],
 *   fallback (2番目以降のキーで一致 = 要確認) [warning], vetoed (IDなどが食い違い結合せず) [info],
 *   variant (表記ゆれ補正) [info], suggestion (もしかして) [info],
 *   dataBelow (Round 5: the table ended at a gap of blank rows, more rows of the same kind follow below it) [warning],
 *   tooManyDuplicates (Round 5: thousands of rows on both sides share one key value — those base rows are not joined but
 *   reported as ambiguous; ID / e-mail / phone should be chosen instead) [warning].
 * Round 5 also renamed / added candidates (ids, labels): 'ID（英字の接頭辞を除く）' (id suffix '~prefix'), 'ID（桁数違い・末尾一致）'
 * ('~suffix'; its matches carry a warn flag 'IDの桁数が異なる'), 'メール', '電話番号', '氏名ローマ字' (Latin-script names in any
 * order, also as 姓＋名 columns). Two-row headers: region.headerRow is the sub-header row, the joined text ('面接評価 評価1') is
 * in KeyAnalysis columns and in defaultImports[].header.
 * `rows` of an item are always rows of the BASE sheet; `otherRows` are rows of the other file named by `sourceId`
 * (unmatchedOther: the other-only rows; duplicate: rows dropped by the policy; ambiguous: the candidates;
 * conflict / checkWarn / vetoed / manualInvalid / suggestion: the other side of the pairs).
 * RowMatch.via: the tier's candidate label ('ID', '氏名＋大学' …), '手動' for manual links, '追加行' for the source
 * an appended row came from.
 */
import type { CheckItem, MatchFlag, MergedRowPlan, RowMatch, SourceBook, TableRegion } from '../types';
import type {
  CandidateDTO,
  ImportColumnSpec,
  JoinPlan,
  JoinResult,
  KeyAnalysis,
  KeyChip,
  KeyFillEntry,
  KeySource,
  NormalizeOptions,
  PairAnalysis,
  PairJoinReport,
  RowRef,
} from '../contracts';
import { type ColumnProfile, cloneProfile, profileTable } from './detect';
import {
  type Candidates,
  type KeySpec,
  type MatchPair,
  type PairStats,
  type Side,
  type Tier,
  SEP,
  attributeKeys,
  attributesDisagree,
  chooseMode,
  executeJoin,
  findCandidates,
  planTiers,
  rebuildSpec,
} from './match';
import { type Cell, cellText, resolveNormalizeOptions, withNormalizeOptions } from './normalize';
import { type RegionOptions, type SourceTable, buildTable, dataBelowTable, detectRegion as detectRegionImpl, findSheet, pickDataSheet as pickDataSheetImpl } from './region';
import { candidateEntries, sheetDTO } from './api';
import { buildManualEntry, dtoPairingKey, pairingKey, resolveManualColumns } from './manual';
import type { Table } from './table';

// ---------------------------------------------------------------------------
// sheet / region
// ---------------------------------------------------------------------------

export function pickDataSheet(book: SourceBook): string {
  try {
    return pickDataSheetImpl(book);
  } catch {
    return book?.sheets?.[0]?.name ?? '';
  }
}

export function detectRegion(book: SourceBook, sheet: string, opts?: RegionOptions): TableRegion {
  try {
    return detectRegionImpl(book, sheet, opts);
  } catch {
    const h = typeof opts?.headerRow === 'number' && Number.isFinite(opts.headerRow) && opts.headerRow >= 1 ? Math.floor(opts.headerRow) : 1;
    return { sheet, headerRow: h, firstDataRow: h + 1, lastDataRow: h, firstCol: 1, lastCol: 1 };
  }
}

// ---------------------------------------------------------------------------
// internal state (KeyAnalysis.internal)
// ---------------------------------------------------------------------------

interface SourceState {
  src: KeySource;
  table: SourceTable;
  profiles: ColumnProfile[];
}

interface CandidateInfo {
  dto: CandidateDTO;
  stats: PairStats;
}

interface PairState {
  sourceId: string;
  candidates: Candidates;
  byId: Map<string, CandidateInfo>;
}

interface Internal {
  kind: 'keys-internal-v1';
  options: NormalizeOptions;
  base: SourceState;
  others: SourceState[];
  pairs: Map<string, PairState>;
}

function internalOf(a: KeyAnalysis): Internal {
  const st = a?.internal as Internal | undefined;
  if (!st || st.kind !== 'keys-internal-v1') throw new Error('キーの分析結果がありません。ファイルを読み込み直してください。');
  return st;
}

function sourceState(src: KeySource): SourceState {
  const table = buildTable(src);
  return { src, table, profiles: profileTable(table) };
}

// ---------------------------------------------------------------------------
// analyze
// ---------------------------------------------------------------------------

export function analyzeSources(base: KeySource, others: KeySource[], options?: Partial<NormalizeOptions>): KeyAnalysis {
  const opts = resolveNormalizeOptions(options);
  return withNormalizeOptions(opts, () => {
    const b = sourceState(base);
    const seen = new Set<string>([base.id]);
    const os: SourceState[] = [];
    for (const o of others ?? []) {
      if (!o || seen.has(o.id)) continue;
      seen.add(o.id);
      os.push(sourceState(o));
    }
    const pairs: Record<string, PairAnalysis> = {};
    const pairStates = new Map<string, PairState>();
    for (const o of os) {
      const c = findCandidates(b.profiles, o.profiles);
      const tiers = planTiers(c);
      const entries = candidateEntries(c, tiers, b.table, o.table);
      const byId = new Map<string, CandidateInfo>();
      for (const e of entries) byId.set(e.dto.id, { dto: e.dto, stats: e.stats });
      pairStates.set(o.src.id, { sourceId: o.src.id, candidates: c, byId });
      pairs[o.src.id] = {
        sourceId: o.src.id,
        other: sheetDTO(o.src.id, o.table, o.profiles),
        mode: chooseMode(tiers),
        candidates: entries.map((e) => e.dto),
      };
    }
    const chips = buildChips(pairs, os.map((o) => o.src.id));
    const internal: Internal = { kind: 'keys-internal-v1', options: opts, base: b, others: os, pairs: pairStates };
    return { baseId: base.id, base: sheetDTO(base.id, b.table, b.profiles), pairs, chips, options: opts, internal };
  });
}

/** Main-screen chips (ID / 氏名): per other source the selected candidate of the family, else the best usable one. */
function buildChips(pairs: Record<string, PairAnalysis>, sourceIds: string[]): KeyChip[] {
  return (['id', 'name'] as const).map((family) => {
    const perSource: KeyChip['perSource'] = {};
    for (const id of sourceIds) {
      const cands = pairs[id].candidates.filter((c) => c.family === family);
      const pick = cands.filter((c) => c.selected).sort((x, y) => (x.tier ?? 99) - (y.tier ?? 99))[0]
        ?? cands.filter((c) => c.verdict !== '使用不可').sort((x, y) => y.score - x.score)[0];
      perSource[id] = pick ? { candidateId: pick.id, label: pick.label, matchRate: pick.matchRate } : null;
    }
    return { family, label: family === 'id' ? 'ID' : '氏名', perSource };
  });
}

export function defaultPlan(analysis: KeyAnalysis): JoinPlan {
  const tiers: Record<string, string[]> = {};
  for (const [id, p] of Object.entries(analysis?.pairs ?? {})) {
    tiers[id] = p.candidates.filter((c) => c.selected).sort((x, y) => (x.tier ?? 99) - (y.tier ?? 99)).map((c) => c.id);
  }
  return { tiers, joinType: 'left', duplicatePolicy: 'first', manualLinks: {} };
}

// ---------------------------------------------------------------------------
// join
// ---------------------------------------------------------------------------

/**
 * Key-column correspondences of a tier, parts aligned one by one: base column ⟷ other column(s).
 *   'copy'  one column each;
 *   'join'  base 氏名 ⟷ other 姓 + 名 (virtual full-name column of the other file);
 *   'split' base 姓 + 名 (virtual) ⟷ other 氏名.
 */
type Alignment =
  | { kind: 'copy'; a: ColumnProfile; b: ColumnProfile }
  | { kind: 'join'; a: ColumnProfile; b: ColumnProfile[] }
  | { kind: 'split'; a: ColumnProfile[]; b: ColumnProfile };

function alignSpecs(a: KeySpec, b: KeySpec): Alignment[] {
  if (a.parts && b.parts && a.parts.length === b.parts.length) return a.parts.flatMap((p, i) => alignSpecs(p, b.parts![i]));
  if (a.cols.length === b.cols.length) return a.cols.map((c, i): Alignment => ({ kind: 'copy', a: c, b: b.cols[i] }));
  if (a.cols.length === 1 && b.virtual && b.cols.length > 1) return [{ kind: 'join', a: a.cols[0], b: b.cols }];
  if (a.virtual && a.cols.length > 1 && b.cols.length === 1) return [{ kind: 'split', a: a.cols, b: b.cols[0] }];
  return [];
}

/** joiner of a base full-name cell filled from split 姓 / 名 columns (the ideographic space Japanese lists use) */
const NAME_JOINER = '　';
const SPLIT_RE = /[\s　]+/;

interface FillEntry {
  /** base column (sheet column) and its index inside the base table */
  baseCol: number;
  baseIdx: number;
  /** source columns (sheet columns) and their indices inside the source table */
  sourceCols: number[];
  srcIdx: number[];
  /** 'join': the source cells joined by NAME_JOINER; 'split': part `part` of the source text split on whitespace */
  mode: 'copy' | 'join' | 'split';
  part?: number;
}

function keyFillOf(tiers: Tier[], tA: Table, tB: Table): FillEntry[] {
  const out: FillEntry[] = [];
  const usedA = new Set<number>();
  const usedB = new Set<number>();
  for (const t of tiers) {
    for (const al of alignSpecs(t.a, t.b)) {
      const as = al.kind === 'split' ? al.a : [al.a];
      const bs = al.kind === 'join' ? al.b : [al.b];
      if (as.some((p) => usedA.has(p.col)) || bs.some((p) => usedB.has(p.col))) continue;
      for (const p of as) usedA.add(p.col);
      for (const p of bs) usedB.add(p.col);
      const srcIdx = bs.map((p) => p.col);
      const sourceCols = srcIdx.map((j) => tB.cols[j]);
      if (al.kind === 'split') {
        as.forEach((p, part) => out.push({ baseCol: tA.cols[p.col], baseIdx: p.col, sourceCols, srcIdx, mode: 'split', part }));
      } else {
        out.push({ baseCol: tA.cols[as[0].col], baseIdx: as[0].col, sourceCols, srcIdx, mode: al.kind });
      }
    }
  }
  return out;
}

/** one aligned (base column, source column) pair of JoinResult.keyPairs, both 1-based sheet columns */
type KeyColumnPair = { baseCol: number; sourceCol: number };

/**
 * JoinResult.keyPairs of ONE source: the direct column-to-column pairs of every enabled tier (single columns; the parts of
 * composite tiers, 氏名＋大学 → 氏名⟷氏名 and 大学⟷大学; 姓⟷姓 / 名⟷名 when both files split the name), de-duplicated, in tier
 * order. 'join' / 'split' alignments (氏名 ⟷ 姓＋名) involve a virtual column and give no pair. Independent of keyFill
 * (which only covers what appended rows need and skips columns an earlier tier already used) and computed for every join.
 *
 * A tier counts only when it matched at least one row: on some matched row (found by ANY tier, manual links included) its
 * two keys are non-blank and equal. That is "the tier identifies the same person in these two columns" — it also holds
 * for a 氏名 tier that matched nothing itself because the ID tier had already taken every row. Without such a row (a
 * wrong manual key, a file whose rows all matched by hand) nothing is known about the columns and no pair is given.
 *
 * `tiers` carry the specs the join ran with (their `a` side possibly rebuilt on the base extended by appended rows, so
 * their keys can be read at row indices of appended rows); the column mapping is the same for both.
 */
function keyPairsOf(tiers: Tier[], tA: Table, tB: Table, matched: MatchPair[]): KeyColumnPair[] {
  const out: KeyColumnPair[] = [];
  const seen = new Set<string>();
  for (const t of tiers) {
    if (!tierMatchedARow(t, matched)) continue;
    for (const al of alignSpecs(t.a, t.b)) {
      if (al.kind !== 'copy') continue;
      const baseCol = tA.cols[al.a.col];
      const sourceCol = tB.cols[al.b.col];
      if (!Number.isInteger(baseCol) || !Number.isInteger(sourceCol)) continue;
      const id = `${baseCol}:${sourceCol}`;
      if (seen.has(id)) continue;
      seen.add(id);
      out.push({ baseCol, sourceCol });
    }
  }
  return out;
}

/** some matched row has equal, non-blank keys on both sides of the tier */
function tierMatchedARow(t: Tier, matched: MatchPair[]): boolean {
  for (const p of matched) {
    const ka = t.a.keys[p.a];
    if (ka && ka === t.b.keys[p.b]) return true;
  }
  return false;
}

/** contract form of the fill entries */
function keyFillEntries(fill: FillEntry[]): KeyFillEntry[] {
  return fill.map((f): KeyFillEntry => {
    if (f.mode === 'join') return { baseCol: f.baseCol, sourceCols: f.sourceCols.slice(), joiner: NAME_JOINER };
    if (f.mode === 'split') return { baseCol: f.baseCol, sourceCol: f.sourceCols[0], splitPart: f.part ?? 0 };
    return { baseCol: f.baseCol, sourceCol: f.sourceCols[0] };
  });
}

function isBlankCell(v: Cell): boolean {
  return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
}

/** value / text of a base key cell of an appended row, from its own file's row */
function fillValue(f: FillEntry, get: (j: number) => Cell): Cell {
  if (f.mode === 'copy') return get(f.srcIdx[0]);
  if (f.mode === 'join') {
    const parts = f.srcIdx.map(get).filter((v) => !isBlankCell(v)).map((v) => cellText(v).trim());
    return parts.length ? parts.join(NAME_JOINER) : null;
  }
  const v = get(f.srcIdx[0]);
  if (isBlankCell(v)) return null;
  return cellText(v).trim().split(SPLIT_RE)[f.part ?? 0] ?? null;
}

interface AppendedRow {
  source: SourceState;
  /** row index in the source table */
  idx: number;
  fill: FillEntry[];
}

/** The base table extended by the appended rows (their base key columns filled from their own file). */
function extendTable(base: SourceTable, appended: AppendedRow[]): Table {
  const n = base.rows.length;
  const width = base.headers.length;
  const rows: Cell[][] = base.rows.slice();
  const excelRow = base.excelRow.slice();
  for (const ap of appended) {
    const row: Cell[] = new Array(width).fill(null);
    const src = ap.source.table.rows[ap.idx];
    for (const f of ap.fill) row[f.baseIdx] = fillValue(f, (j) => src?.[j] ?? null);
    rows.push(row);
    excelRow.push(0);
  }
  const display = (i: number, col: number): string => {
    if (i < n) return base.display(i, col);
    const ap = appended[i - n];
    const f = ap?.fill.find((x) => x.baseIdx === col);
    if (!f) return '';
    const v = fillValue(f, (j) => ap.source.table.display(ap.idx, j));
    return v === null ? '' : String(v);
  };
  return { ...base, rows, excelRow, display };
}

const VARIANT_RE = /^(異体字の違い|カナ表記ゆれ|ID表記の違いを補正|姓名の順序違い)/;

function matchFlags(p: MatchPair): MatchFlag[] {
  const f: MatchFlag[] = [];
  if (p.manual) f.push('manual');
  else if (p.tierIndex > 0) f.push('fallback');
  if (p.flags.some((x) => x.level === 'error' || x.level === 'warn')) f.push('conflict');
  if (p.dupSource) f.push('duplicateInSource');
  if (p.dupBase) f.push('duplicateInBase');
  if (p.flags.some((x) => VARIANT_RE.test(x.msg))) f.push('variant');
  return f;
}

function rowRef(t: Table, i: number): RowRef {
  const preview: string[] = [];
  const width = Math.min(4, t.headers.length);
  for (let c = 0; c < width; c++) preview.push(t.display(i, c));
  return { excelRow: t.excelRow[i], preview };
}

/** unmatched other rows grouped by identity (same tier key, no tier disagreeing) — duplicates among new rows */
function groupNewRows(rows: number[], tiers: Tier[], policy: JoinPlan['duplicatePolicy'], attr: string[][]): Array<{ rep: number; members: number[] }> {
  if (policy === 'skip' || !tiers.length) return rows.map((j) => ({ rep: j, members: [j] }));
  const groups: number[][] = [];
  const index = tiers.map(() => new Map<string, number>());
  const agree = (x: number, y: number) => tiers.every((t) => {
    const a = t.b.keys[x];
    const b = t.b.keys[y];
    return !a || !b || a === b;
  });
  for (const j of rows) {
    let g = -1;
    for (let t = 0; t < tiers.length && g < 0; t++) {
      const k = tiers[t].b.keys[j];
      if (!k) continue;
      const gi = index[t].get(k);
      // rows with the same NAME but a different e-mail / phone / birth date / 大学 are two people, not a resubmission
      if (gi !== undefined && agree(groups[gi][0], j) && (tiers[t].b.sig === 'idFamily' || !attributesDisagree(attr, groups[gi][0], j))) g = gi;
    }
    if (g < 0) {
      g = groups.length;
      groups.push([j]);
    } else {
      groups[g].push(j);
    }
    for (let t = 0; t < tiers.length; t++) {
      const k = tiers[t].b.keys[j];
      if (k && !index[t].has(k)) index[t].set(k, g);
    }
  }
  return groups.map((members) => ({ rep: policy === 'last' ? members[members.length - 1] : members[0], members }));
}

const POLICY_LABEL: Record<JoinPlan['duplicatePolicy'], string> = { first: '先頭の行を採用', last: '最後の行を採用', skip: '結合しない' };

function rowsList(rows: number[], max = 8): string {
  const head = rows.slice(0, max).map((r) => `${r}行目`).join('、');
  return rows.length > max ? `${head} ほか` : head;
}

const EXAMPLE_ROWS = 50;
/** candidates listed per ambiguous base row in the report (a value shared by thousands of rows has thousands of candidates) */
const MAX_CANDIDATES = 20;
const num = (n: number) => String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ',');

export function joinSources(analysis: KeyAnalysis, plan: JoinPlan): JoinResult {
  const st = internalOf(analysis);
  return withNormalizeOptions(st.options, () => joinImpl(st, plan));
}

function joinImpl(st: Internal, planIn: JoinPlan): JoinResult {
  const plan: JoinPlan = {
    tiers: planIn?.tiers ?? {},
    joinType: planIn?.joinType === 'full' ? 'full' : 'left',
    duplicatePolicy: planIn?.duplicatePolicy === 'last' || planIn?.duplicatePolicy === 'skip' ? planIn.duplicatePolicy : 'first',
    manualLinks: planIn?.manualLinks ?? {},
  };
  const base = st.base;
  const nBase = base.table.rows.length;
  const baseIndexOfRow = new Map<number, number>();
  base.table.excelRow.forEach((r, i) => baseIndexOfRow.set(r, i));
  const appended: AppendedRow[] = [];
  /** matches per extended row (base rows, then appended rows) */
  const matches: Array<Record<string, RowMatch>> = [];
  for (let i = 0; i < nBase; i++) matches.push({});
  const reports: Record<string, PairJoinReport> = {};
  const items: CheckItem[] = [];
  const keyFill: Record<string, KeyFillEntry[]> = {};
  const keyPairs: Record<string, KeyColumnPair[]> = {};
  const defaultImports: ImportColumnSpec[] = [];

  sourceNotes(items, base, true);

  for (const o of st.others) {
    const id = o.src.id;
    const label = o.src.label || o.table.file;
    const pair = st.pairs.get(id);
    sourceNotes(items, o, false);

    // ---- tiers of this pair (plan order; unknown ids skipped) ----
    const tierIds: string[] = [];
    const unknown: string[] = [];
    for (const cid of plan.tiers[id] ?? []) {
      if (tierIds.includes(cid)) continue;
      if (pair?.byId.has(cid)) tierIds.push(cid);
      else unknown.push(cid);
    }
    if (unknown.length) {
      items.push({ severity: 'warning', code: 'planInvalid', sourceId: id, message: `「${label}」: 使えない結合キーの指定を無視しました（${unknown.length}件）。キーを選び直してください。` });
    }
    const tiers: Tier[] = tierIds.map((cid) => {
      const c = pair!.byId.get(cid)!;
      return { name: c.dto.label, a: c.stats.a, b: c.stats.b, stats: c.stats };
    });
    const fill = keyFillOf(tiers, base.table, o.table);
    keyFill[id] = keyFillEntries(fill);
    const fillCols = new Set(fill.flatMap((f) => f.sourceCols));
    o.table.cols.forEach((c, j) => {
      if (fillCols.has(c)) return;
      const spec: ImportColumnSpec = { sourceId: id, col: c };
      // a header the premerge would not find in the header cell itself (a merged group title, a two-row header) is passed on
      const h = o.table.headers[j];
      if (h && !o.table.headerFallback[j] && h !== o.table.headerRaw[j]) spec.header = h;
      defaultImports.push(spec);
    });

    // ---- base side: base rows (+ rows appended from earlier files) ----
    let A: Side = { table: base.table, profiles: base.profiles };
    let tiersA = tiers;
    if (appended.length) {
      const ext = extendTable(base.table, appended);
      const ps = base.profiles.map((p) => cloneProfile(p, ext, nBase));
      A = { table: ext, profiles: ps };
      tiersA = tiers.map((t) => ({ ...t, a: rebuildSpec(t.a, ps) }));
    }
    const B: Side = { table: o.table, profiles: o.profiles };
    const mode = chooseMode(tiersA);

    // ---- manual links (Excel rows → table rows) ----
    const otherIndexOfRow = new Map<number, number>();
    o.table.excelRow.forEach((r, i) => otherIndexOfRow.set(r, i));
    const manualIn = plan.manualLinks[id] ?? [];
    const manual = manualIn.map((l) => ({ a: baseIndexOfRow.get(l?.baseRow) ?? -1, b: otherIndexOfRow.get(l?.otherRow) ?? -1 }));
    // no key and no manual link (plan.tiers[id] is empty: the UI excludes the file): nothing can be matched, so the
    // file contributes no matches, appends no rows and is reported once, as noKey — never as "rows only in this file"
    const keyless = !tiers.length && !manualIn.length;

    const r = executeJoin(A, B, tiersA, { mode, policy: plan.duplicatePolicy, manual, discovered: pair?.candidates.discovered ?? [], suggestLimitA: nBase });

    for (const p of r.pairs) {
      matches[p.a][id] = { row: o.table.excelRow[p.b], via: p.manual ? '手動' : p.tier, flags: matchFlags(p) };
    }

    // column pairs of the enabled tiers (read with the keys of tiersA: matched rows can be rows appended from earlier files)
    const kp = keyPairsOf(tiersA, base.table, o.table, r.pairs);
    if (kp.length) keyPairs[id] = kp;

    // ---- 'full': append what matched nothing (duplicates among the new rows collapsed by the policy) ----
    const held = new Set(r.heldB);
    const newRows = r.unmatchedB.filter((j) => !held.has(j));
    let droppedNew = 0;
    const droppedNewRows: number[] = [];
    if (plan.joinType === 'full' && !keyless) {
      const unsplit: number[] = [];
      const splitFill = fill.filter((f) => f.mode === 'split');
      for (const g of groupNewRows(newRows, tiers, plan.duplicatePolicy, attributeKeys(o.profiles))) {
        appended.push({ source: o, idx: g.rep, fill });
        // 姓 / 名 of the base are filled from ONE 氏名 cell: without a space in it the surname column gets the whole name
        for (const f of splitFill) {
          const v = o.table.rows[g.rep]?.[f.srcIdx[0]];
          if (!isBlankCell(v ?? null) && !/[\s　]/.test(cellText(v ?? null).trim())) {
            unsplit.push(o.table.excelRow[g.rep]);
            break;
          }
        }
        const flags: MatchFlag[] = g.members.length > 1 ? ['duplicateInSource'] : [];
        matches.push({ [id]: { row: o.table.excelRow[g.rep], via: '追加行', flags } });
        if (g.members.length > 1) {
          droppedNew += g.members.length - 1;
          for (const j of g.members) if (j !== g.rep) droppedNewRows.push(o.table.excelRow[j]);
        }
      }
      if (unsplit.length) {
        items.push({
          severity: 'warning', code: 'nameNotSplit', sourceId: id, otherRows: unsplit.slice(0, EXAMPLE_ROWS),
          message: `「${label}」から追加した行のうち、氏名にスペースがなく姓と名に分けられなかった行: ${unsplit.length}件（${rowsList(unsplit)}）。姓の欄に氏名全体を入れています。`,
        });
      }
    }

    reports[id] = pairReport(id, base.table, o.table, nBase, tiers, mode, r);
    pairItems(items, { id, label, nBase, base: base.table, other: o.table, tiers, mode, r, plan, newRows, droppedNew, droppedNewRows, manualCount: manualIn.length, keyless });
  }

  // ---- merged row plan: every base region row in order, then the appended rows ----
  const ids = st.others.map((o) => o.src.id);
  const rows: MergedRowPlan[] = [];
  const region = base.table.region;
  for (let r = region.firstDataRow; r <= region.lastDataRow; r++) {
    const i = baseIndexOfRow.get(r);
    const m: Record<string, RowMatch | null> = {};
    for (const id of ids) m[id] = i === undefined ? null : matches[i][id] ?? null;
    rows.push({ baseRow: r, matches: m });
  }
  for (let k = 0; k < appended.length; k++) {
    const ap = appended[k];
    const m: Record<string, RowMatch | null> = {};
    for (const id of ids) m[id] = matches[nBase + k][id] ?? null;
    rows.push({ baseRow: null, appendedFrom: { sourceId: ap.source.src.id, row: ap.source.table.excelRow[ap.idx] }, matches: m });
  }

  let matchedAny = 0;
  let matchedAll = 0;
  for (let i = 0; i < nBase; i++) {
    const n = ids.filter((id) => matches[i][id]).length;
    if (n > 0) matchedAny++;
    if (ids.length && n === ids.length) matchedAll++;
  }
  return { rows, reports, summary: { baseRows: nBase, matchedAny, matchedAll, appended: appended.length }, items, keyFill, defaultImports, keyPairs };
}

function sourceNotes(items: CheckItem[], s: SourceState, isBase: boolean): void {
  const label = s.src.label || s.table.file;
  const who = isBase ? `基準ファイル「${label}」` : `「${label}」`;
  const sourceId = s.src.id;
  if (!s.table.rows.length) {
    items.push({ severity: isBase ? 'error' : 'warning', code: 'emptySource', sourceId, message: `${who}にデータ行がありません（シート・見出し行を確認してください）。` });
  }
  if (s.table.layoutHint) items.push({ severity: 'warning', code: 'layout', sourceId, message: `${who}: ${s.table.layoutHint}` });
  // a blank gap ends a table; rows of the same kind below it are then NOT part of the merge — say so instead of dropping them silently
  const sheet = findSheet(s.src.book, s.table.region.sheet);
  const below = sheet ? dataBelowTable(sheet, s.table.region) : null;
  if (below) {
    items.push({
      severity: 'warning', code: 'dataBelow', sourceId,
      message: `${who}: 表の最終行（${s.table.region.lastDataRow}行目）の下に、空白行をはさんでデータらしい行が ${below.count} 行あります（${below.firstRow}行目〜）。それらは結合に使われていません。続きのデータなら、範囲の最終行を指定し直してください。`,
    });
  }
  if (s.table.formulasWithoutValue > 0) {
    items.push({
      severity: 'warning', code: 'noCachedValue', sourceId,
      message: `${who}: 計算結果が保存されていない数式が ${s.table.formulasWithoutValue} 個あります（空欄として扱います）。Excelで開いて保存し直すと正しく照合できます。`,
    });
  }
  for (const p of s.profiles) {
    if (p.notes.some((n) => n.startsWith('指数表記'))) {
      items.push({ severity: 'warning', code: 'lossyId', sourceId, message: `${who}の「${p.header}」: 指数表記（8.00E+09 など）で桁が失われたIDがあります。元のデータを確認してください。` });
    }
  }
}

/** report entries of the ambiguous base rows; rows sharing one candidate list (a pathological group) share one RowRef list */
function ambiguousReport(list: ReturnType<typeof executeJoin>['ambiguous'], tA: Table, tB: Table, nBase: number): PairJoinReport['ambiguous'] {
  const refs = new Map<number[], RowRef[]>();
  const out: PairJoinReport['ambiguous'] = [];
  for (const x of list) {
    if (x.a >= nBase) continue;
    let c = refs.get(x.candidatesB);
    if (!c) {
      c = x.candidatesB.slice(0, MAX_CANDIDATES).map((j) => rowRef(tB, j));
      refs.set(x.candidatesB, c);
    }
    out.push({ base: rowRef(tA, x.a), candidates: c, tier: x.tierIndex + 1 });
  }
  return out;
}

/** the first candidate rows of the ambiguous base rows (examples only, each shared list once) */
function exampleCandidates(list: ReturnType<typeof executeJoin>['ambiguous']): number[] {
  const out: number[] = [];
  const seen = new Set<number[]>();
  for (const a of list) {
    if (seen.has(a.candidatesB)) continue;
    seen.add(a.candidatesB);
    for (const j of a.candidatesB) {
      out.push(j);
      if (out.length >= EXAMPLE_ROWS * 4) return out;
    }
  }
  return out;
}

function pairReport(id: string, tA: Table, tB: Table, nBase: number, tiers: Tier[], mode: 'oneToOne' | 'lookup', r: ReturnType<typeof executeJoin>): PairJoinReport {
  const flags: PairJoinReport['flags'] = {};
  const addFlag = (kind: string, level: string, example: string) => {
    const g = (flags[kind] ??= { level, count: 0, examples: [] });
    g.count++;
    if (g.examples.length < 5) g.examples.push(example);
  };
  const where = (p: MatchPair) => `${p.a < nBase ? `${tA.excelRow[p.a]}行目` : '追加行'} ⟷ ${tB.excelRow[p.b]}行目`;
  for (const p of r.pairs) {
    for (const f of p.flags) addFlag(f.msg.replace(/（.*$/, ''), f.level, `${where(p)}: ${f.msg}`);
    if (p.dupSource) addFlag('重複（ルールで1行を採用）', 'warn', `${where(p)}（ほかに ${(p.dropped ?? []).map((j) => `${tB.excelRow[j]}行目`).join('、')}）`);
    if (p.manual) addFlag('手動で結合', 'info', where(p));
  }
  const conflicts = [...new Map(r.conflicts.filter((x) => x.a < nBase).map((x) => [`${x.a}|${x.b}`, x])).values()];
  return {
    sourceId: id,
    mode,
    plan: tiers.map((t, i) => ({ order: i + 1, name: t.name })),
    matched: r.pairs.length,
    perTier: { ...r.perTier },
    unmatchedBase: r.unmatchedA.filter((i) => i < nBase).map((i) => rowRef(tA, i)),
    unmatchedOther: r.unmatchedB.map((j) => rowRef(tB, j)),
    duplicatesBase: r.dupA.map(([k, n]) => ({ key: k.split(SEP).join('＋'), count: n })),
    duplicatesOther: r.dupB.map(([k, n]) => ({ key: k.split(SEP).join('＋'), count: n })),
    ambiguous: ambiguousReport(r.ambiguous, tA, tB, nBase),
    conflicts: conflicts.map((x) => ({ base: rowRef(tA, x.a), other: rowRef(tB, x.b), message: x.msg })),
    flags,
    suggestions: r.suggestions.map((s) => ({ base: rowRef(tA, s.a), other: rowRef(tB, s.b), score: s.score, reasons: s.reasons })),
  };
}

interface PairItemCtx {
  id: string;
  label: string;
  nBase: number;
  base: Table;
  other: Table;
  tiers: Tier[];
  mode: 'oneToOne' | 'lookup';
  r: ReturnType<typeof executeJoin>;
  plan: JoinPlan;
  newRows: number[];
  droppedNew: number;
  droppedNewRows: number[];
  manualCount: number;
  /** no tier and no manual link: only the noKey item (and suggestions) are reported */
  keyless: boolean;
}

function pairItems(items: CheckItem[], x: PairItemCtx): void {
  const { id: sourceId, label, nBase, base, other, r, plan } = x;
  const baseRows = (idx: number[]) => [...new Set(idx.filter((i) => i < nBase).map((i) => base.excelRow[i]))].sort((a, b) => a - b);
  /** Excel rows of the other file from its table indices */
  const oRows = (idx: number[]) => idx.filter((j) => j >= 0 && j < other.rows.length).map((j) => other.excelRow[j]);
  /** idx: base table indices (→ rows); otherExcel: Excel rows of the other file (→ otherRows) */
  const push = (severity: CheckItem['severity'], code: string, message: string, idx?: number[], otherExcel?: number[]) => {
    const item: CheckItem = { severity, code, sourceId, message };
    const rows = idx ? baseRows(idx) : [];
    if (rows.length) item.rows = rows.slice(0, EXAMPLE_ROWS);
    const others = otherExcel ? [...new Set(otherExcel.filter((r) => Number.isInteger(r) && r > 0))].sort((a, b) => a - b) : [];
    if (others.length) item.otherRows = others.slice(0, EXAMPLE_ROWS);
    items.push(item);
  };
  if (x.keyless) {
    push('error', 'noKey', `「${label}」と基準ファイルに共通する結合キーが見つかりません。「詳細設定」で結合に使う列を選んでください。`);
  }
  for (const m of r.manualRejected) {
    const br = m.a >= 0 && m.a < nBase ? `基準 ${base.excelRow[m.a]}行目` : '基準の行がデータ範囲外';
    const or = m.b >= 0 && m.b < other.rows.length ? `「${label}」${other.excelRow[m.b]}行目` : `「${label}」の行がデータ範囲外`;
    push('warning', 'manualInvalid', `手動の結合を適用できませんでした（${br} ⟷ ${or}）: ${m.reason}`, m.a >= 0 ? [m.a] : undefined, oRows([m.b]));
  }
  if (r.limited.length) {
    const rows = r.limited.reduce((n, g) => n + g.rows, 0);
    const biggest = r.limited.reduce((m, g) => Math.max(m, g.rows, g.other), 0);
    const tierName = [...new Set(r.limited.map((g) => g.tier))].join('・');
    push('warning', 'tooManyDuplicates',
      `「${label}」: 結合キー「${tierName}」と同じ値の行が大量にあります（最大 ${num(biggest)} 行）。該当する基準の ${num(rows)} 行は組み合わせが多すぎるため自動で結合せず、「候補が複数」として扱いました。ID・メール・電話番号など行ごとに異なる列を結合キーに選んでください。`);
  }
  const unmatchedBase = r.unmatchedA.filter((i) => i < nBase);
  // without any key every row is "unmatched": the noKey error says it once (reports[id] still lists the rows)
  if (unmatchedBase.length && !x.keyless) push('warning', 'unmatched', `「${label}」に見つからない基準の行: ${unmatchedBase.length}件`, unmatchedBase);
  if (x.newRows.length && !x.keyless) {
    const otherRows = x.newRows.map((j) => other.excelRow[j]);
    const tail = plan.joinType === 'full' ? '末尾に追加します' : '結合されません（「他ファイルにしかない行も末尾に追加」で追加できます）';
    push(plan.joinType === 'full' ? 'info' : 'warning', 'unmatchedOther', `「${label}」にだけある行: ${x.newRows.length}件（${rowsList(otherRows)}）— ${tail}`, undefined, otherRows);
  }
  const dupPairs = r.pairs.filter((p) => p.dupSource);
  const dupCount = dupPairs.length + x.droppedNew;
  if (dupCount) {
    push('warning', 'duplicate', `「${label}」でキーが重複している行: ${dupCount}件（${POLICY_LABEL[plan.duplicatePolicy]}）${x.droppedNewRows.length ? `。追加しなかった重複行: ${rowsList(x.droppedNewRows)}` : ''}`, dupPairs.map((p) => p.a), [...oRows(dupPairs.flatMap((p) => p.dropped ?? [])), ...x.droppedNewRows]);
  }
  const dupBaseRows = r.pairs.filter((p) => p.dupBase && !p.manual).map((p) => p.a);
  const dupBaseAmbig = r.ambiguous.filter((a) => a.sharedBase && a.a < nBase).map((a) => a.a);
  if (x.mode === 'lookup') {
    // VLOOKUP-style join (職場ID → 職場回答): many base rows receiving the same row is the point, not a problem
    if (dupBaseRows.length) push('info', 'lookup', `「${label}」は参照（VLOOKUP）形式で結合しました: 同じキーを持つ基準の複数の行に同じ行を結合（${dupBaseRows.length}件）`);
  } else if (dupBaseRows.length || dupBaseAmbig.length) {
    push('warning', 'duplicateBase', `基準ファイルで同じキーの行が複数あります（「${label}」との結合）: ${dupBaseRows.length + dupBaseAmbig.length}件`, [...dupBaseRows, ...dupBaseAmbig]);
  }
  const ambiguous = r.ambiguous.filter((a) => a.a < nBase);
  if (ambiguous.length) {
    push('warning', 'ambiguous', `「${label}」に候補が複数あり結合できなかった行: ${ambiguous.length}件（詳細設定で手動で選べます）`, ambiguous.map((a) => a.a), oRows(exampleCandidates(ambiguous)));
  }
  const errors = r.pairs.filter((p) => p.flags.some((f) => f.level === 'error'));
  if (errors.length) push('error', 'conflict', `ID一致・氏名不一致（「${label}」）: ${errors.length}件`, errors.map((p) => p.a), oRows(errors.map((p) => p.b)));
  const warns = r.pairs.filter((p) => !p.flags.some((f) => f.level === 'error') && p.flags.some((f) => f.level === 'warn'));
  if (warns.length) push('warning', 'checkWarn', `要確認（似た字・1文字違い・フリガナ／生年月日／メール／電話／大学の不一致など、「${label}」）: ${warns.length}件`, warns.map((p) => p.a), oRows(warns.map((p) => p.b)));
  const fallback = r.pairs.filter((p) => !p.manual && p.tierIndex > 0);
  if (fallback.length) {
    push('warning', 'fallback', `2番目以降のキー（${[...new Set(fallback.map((p) => p.tier))].join('・')}）で一致した行（「${label}」）: ${fallback.length}件（要確認）`, fallback.map((p) => p.a));
  }
  const unmatchedSet = new Set(r.unmatchedA);
  const vetoed = [...new Set(r.conflicts.filter((c) => c.a < nBase && unmatchedSet.has(c.a)).map((c) => c.a))];
  if (vetoed.length) {
    push('info', 'vetoed', `別のキー（IDなど）が食い違うため結合しなかった行（「${label}」）: ${vetoed.length}件`, vetoed, oRows(r.conflicts.filter((c) => c.a < nBase && unmatchedSet.has(c.a)).map((c) => c.b)));
  }
  const variant = r.pairs.filter((p) => p.flags.some((f) => VARIANT_RE.test(f.msg)));
  if (variant.length) push('info', 'variant', `表記ゆれ（異体字・全角/半角・先頭ゼロなど）を補正して一致（「${label}」）: ${variant.length}件`, variant.map((p) => p.a));
  if (r.suggestions.length) {
    const rows = [...new Set(r.suggestions.map((s) => s.a))];
    push('info', 'suggestion', `もしかして（「${label}」との候補）: ${rows.length}件 — 詳細設定で確認して結合できます`, rows, oRows(r.suggestions.map((s) => s.b)));
  }
}

/**
 * Manual key candidate — the user picks columns of the base sheet and of ONE other file (1..3 columns each, same count,
 * e.g. [氏名] ⟷ [氏名] or [氏名, 大学] ⟷ [氏名, 大学]; columns are sheet column numbers, 1-based, like
 * CandidateDTO.baseColumns / otherColumns). Returns a NEW KeyAnalysis whose pair for `sourceId` has one more CandidateDTO
 * (appended last; id 'manual:<baseCols>|<otherCols>' in the user's order, kind 'single' | 'composite', label
 * '<base header> ⟷ <other header>', family from the columns' detected types, real statistics with the analysis' normalize
 * options, selected:false). The candidate is registered in the pair's internal state, so joinSources uses it once the UI
 * puts its id into plan.tiers[sourceId]. Nothing of the given analysis is mutated (internal state is copied-on-write).
 *
 *  - invalid columns (not 1..3 / different counts / outside the region / repeated): Error with a Japanese message;
 *  - a candidate over the same column pairing already exists (auto, discovered or manual): `analysis` itself is returned;
 *  - both sides are normalised the same way (see manual.ts manualPart): own key type when the two columns are of the same
 *    type family, the recognised column's type when the other one is unrecognised, else generic text.
 */
export function addManualCandidate(analysis: KeyAnalysis, sourceId: string, baseCols: number[], otherCols: number[]): KeyAnalysis {
  const st = internalOf(analysis);
  const pair = st.pairs.get(sourceId);
  const pa = analysis.pairs?.[sourceId];
  const other = st.others.find((o) => o.src.id === sourceId);
  if (!pair || !pa || !other) {
    throw new Error(sourceId === analysis.baseId
      ? '基準ファイル自身との結合キーは作れません。別のファイルを選んでください。'
      : '指定されたファイルの結合キーの分析結果が見つかりません。ファイルを読み込み直してください。');
  }
  const label = other.src.label || other.table.file;
  const sel = resolveManualColumns(st.base, { table: other.table, profiles: other.profiles, label }, baseCols, otherCols);
  const key = pairingKey(baseCols, otherCols);
  if (pa.candidates.some((c) => dtoPairingKey(c) === key)) return analysis;

  const entry = withNormalizeOptions(st.options, () =>
    buildManualEntry(st.base.table, other.table, sel.a, sel.b, baseCols.slice(), otherCols.slice(), new Set(pa.candidates.map((c) => c.label))),
  );
  const byId = new Map(pair.byId);
  byId.set(entry.dto.id, { dto: entry.dto, stats: entry.stats });
  const pairStates = new Map(st.pairs);
  pairStates.set(sourceId, { ...pair, byId });
  const pairs: Record<string, PairAnalysis> = { ...analysis.pairs, [sourceId]: { ...pa, candidates: [...pa.candidates, entry.dto] } };
  const internal: Internal = { ...st, pairs: pairStates };
  return { ...analysis, pairs, chips: buildChips(pairs, st.others.map((o) => o.src.id)), internal };
}
