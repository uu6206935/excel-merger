/**
 * DTO builders for the UI (contract types in src/core/contracts.ts) + the R6 two-table convenience API
 * (analyzeKeys / runJoin) that the tests use to reproduce the research results.
 */
import type { CandidateDTO, ColumnDTO, SheetDTO } from '../contracts';
import { type ColumnProfile, ID_TYPES, TYPE_LABEL, profileTable } from './detect';
import { type Candidates, type JoinOptions, type PairStats, type Tier, SEP, chooseMode, executeJoin, findCandidates, planTiers } from './match';
import { cellText } from './normalize';
import type { Table } from './table';

const r3 = (x: number) => Math.round(x * 1000) / 1000;

/** short Japanese label of a key signature ('idFamily' → 'ID', 'name+university' → '氏名＋大学') */
const SIG_LABEL: Record<string, string> = {
  idFamily: 'ID', name: '氏名', kana: '氏名カナ', roman: '氏名ローマ字', email: 'メール', phone: '電話番号', university: '大学',
  faculty: '学部・学科', birthDate: '生年月日', gender: '性別', graduation: '卒業年', lastName: '姓', firstName: '名',
  lastNameKana: '姓カナ', firstNameKana: '名カナ',
};
export const NAME_TYPES: ReadonlySet<string> = new Set(['name', 'nameKana', 'nameRoman', 'lastName', 'firstName', 'lastNameKana', 'firstNameKana']);

export function sigLabel(sig: string): string {
  return sig.split('+').map((s) => SIG_LABEL[s] ?? TYPE_LABEL[s] ?? s).join('＋');
}

export function columnDTO(t: Table, p: ColumnProfile): ColumnDTO {
  return {
    col: t.cols[p.col] ?? p.col + 1,
    header: p.header,
    type: p.type,
    typeLabel: TYPE_LABEL[p.type] ?? '—',
    confidence: p.confidence,
    typeScore: r3(p.typeScore),
    keyScore: r3(p.keyScore),
    uniqueness: r3(p.uniqueness),
    fill: r3(p.fill),
    blank: p.blank,
    dupRows: p.dupRows,
    notes: [...p.notes],
  };
}

export function sheetDTO(sourceId: string, t: Table, ps: ColumnProfile[]): SheetDTO {
  const dto: SheetDTO = { sourceId, file: t.file, sheet: t.sheet, headerRow: t.headerRow, rows: t.rows.length, columns: ps.map((p) => columnDTO(t, p)) };
  if (t.layoutHint) dto.layoutHint = t.layoutHint;
  return dto;
}

/** Chip family of a candidate: ID-like / name-like (incl. composites containing a name) / other. */
export function familyOf(s: PairStats): CandidateDTO['family'] {
  const sig = s.a.sig;
  if (sig === 'idFamily' || sig.startsWith('idFamily+')) return 'id';
  if (sig === 'discovered') {
    const ps = [...s.a.cols, ...s.b.cols];
    if (ps.some((p) => ID_TYPES.has(p.type) || p.type === 'orgCode')) return 'id';
    if (ps.some((p) => NAME_TYPES.has(p.type))) return 'name';
    if (ps.every((p) => (p.ratios.idShape ?? 0) >= 0.8)) return 'id';
    return 'other';
  }
  const first = sig.split('+')[0];
  if (first === 'name' || first === 'kana' || first === 'roman') return 'name';
  return 'other';
}

const headersOf = (cols: ColumnProfile[]) => cols.map((c) => c.header).join('＋');

function baseLabel(s: PairStats, kind: CandidateDTO['kind']): string {
  if (kind === 'discovered') {
    const ha = headersOf(s.a.cols);
    const hb = headersOf(s.b.cols);
    return ha === hb ? `値一致（${ha}）` : `値一致（${ha} ⟷ ${hb}）`;
  }
  return s.a.variant || s.a.fixedLabel ? s.a.label : sigLabel(s.a.sig);
}

export function candidateId(kind: CandidateDTO['kind'], s: PairStats, tA: Table, tB: Table): string {
  const ca = s.a.cols.map((c) => tA.cols[c.col] ?? c.col + 1).join('+');
  const cb = s.b.cols.map((c) => tB.cols[c.col] ?? c.col + 1).join('+');
  // an ID variant (letter prefix / tail digits) runs over the same column pair as the literal ID candidate
  return `${kind}:${ca}|${cb}${s.a.variant ? `~${s.a.variant}` : ''}`;
}

export interface CandidateEntry {
  dto: CandidateDTO;
  stats: PairStats;
}

/** The contract DTO of one key pair (shared by the auto candidates and the manual one). tier: 0-based plan tier or -1. */
export function candidateDTO(
  o: { id: string; kind: CandidateDTO['kind']; label: string; family: CandidateDTO['family']; s: PairStats; tier: number },
  tA: Table,
  tB: Table,
): CandidateDTO {
  const { s, tier } = o;
  return {
    id: o.id, kind: o.kind, label: o.label, family: o.family,
    baseColumns: s.a.cols.map((p) => ({ col: tA.cols[p.col] ?? p.col + 1, header: p.header })),
    otherColumns: s.b.cols.map((p) => ({ col: tB.cols[p.col] ?? p.col + 1, header: p.header })),
    matchRate: r3(s.matchRate), coverage: r3(s.coverageA), matchedOther: s.matchedB, nonBlankOther: s.nonBlankB,
    matchedBase: s.matchedA, nonBlankBase: s.nonBlankA, uniqBase: r3(s.uniqA), uniqOther: r3(s.uniqB),
    dupRowsBase: s.dupRowsA, dupRowsOther: s.dupRowsB, blankBase: s.blankA, blankOther: s.blankB,
    ambiguous: s.ambiguousB, cardinality: s.cardinality, score: r3(s.score), verdict: s.verdict,
    selected: tier >= 0, tier: tier >= 0 ? tier + 1 : undefined,
  };
}

/**
 * Candidate list of one pair (base A ⟷ other B) in UI order: auto-selected tiers first, then by score.
 * Labels are short ('ID', '氏名＋大学'); a label used by several candidates is qualified with the column headers.
 */
export function candidateEntries(c: Candidates, tiers: Tier[], tA: Table, tB: Table): CandidateEntry[] {
  const tierOf = (s: PairStats) => tiers.findIndex((t) => t.a === s.a && t.b === s.b);
  const raw: Array<{ s: PairStats; kind: CandidateDTO['kind'] }> = [
    ...c.singles.map((s) => ({ s, kind: 'single' as const })),
    // the default ID＋名前 key always comes first and does not use up one of the five composite slots
    ...(c.idName ? [{ s: c.idName, kind: 'composite' as const }] : []),
    ...c.composites.filter((s) => s !== c.idName).slice(0, 5).map((s) => ({ s, kind: 'composite' as const })),
    ...c.discovered.slice(0, 5).map((s) => ({ s, kind: 'discovered' as const })),
  ];
  const entries: CandidateEntry[] = [];
  const seenIds = new Set<string>();
  for (const { s, kind } of raw) {
    const tier = tierOf(s);
    if (tier < 0 && s.verdict === '使用不可') continue;
    const id = candidateId(kind, s, tA, tB);
    if (seenIds.has(id)) continue;
    seenIds.add(id);
    entries.push({ stats: s, dto: candidateDTO({ id, kind, label: baseLabel(s, kind), family: familyOf(s), s, tier }, tA, tB) });
  }
  entries.sort((p, q) => (p.dto.tier ?? 99) - (q.dto.tier ?? 99) || q.dto.score - p.dto.score);
  // qualify repeated labels (the best-ranked keeps the short one)
  const used = new Set<string>();
  for (const e of entries) {
    if (used.has(e.dto.label)) {
      const ha = headersOf(e.stats.a.cols);
      const hb = headersOf(e.stats.b.cols);
      e.dto.label = `${e.dto.label}（${ha === hb ? ha : `${ha} ⟷ ${hb}`}）`;
    }
    used.add(e.dto.label);
  }
  return entries;
}

// ---------------------------------------------------------------------------
// two-table convenience API (R6 analyzeKeys / runJoin)
// ---------------------------------------------------------------------------

export interface TwoTableAnalysis {
  base: SheetDTO;
  other: SheetDTO;
  mode: 'oneToOne' | 'lookup';
  candidates: CandidateDTO[];
  internal: { A: ColumnProfile[]; B: ColumnProfile[]; tA: Table; tB: Table; c: Candidates; tiers: Tier[]; mode: 'oneToOne' | 'lookup' };
}

export function analyzeKeys(base: Table, other: Table): TwoTableAnalysis {
  const A = profileTable(base);
  const B = profileTable(other);
  const c = findCandidates(A, B);
  const tiers = planTiers(c);
  const mode = chooseMode(tiers);
  const candidates = candidateEntries(c, tiers, base, other).map((e) => e.dto);
  return { base: sheetDTO('base', base, A), other: sheetDTO('other', other, B), mode, candidates, internal: { A, B, tA: base, tB: other, c, tiers, mode } };
}

export function runJoin(an: TwoTableAnalysis, opts: JoinOptions = {}) {
  const { A, B, tA, tB, c, tiers, mode } = an.internal;
  const r = executeJoin({ table: tA, profiles: A }, { table: tB, profiles: B }, tiers, { mode, discovered: c.discovered, ...opts });
  const rowRef = (t: Table, i: number) => ({ excelRow: t.excelRow[i], preview: t.rows[i].slice(0, 4).map((v) => cellText(v)) });
  // flags aggregated by kind so the UI can show "異体字の違い: 5件" instead of 5 lines
  const flagGroups: Record<string, { level: string; count: number; examples: string[] }> = {};
  for (const p of r.pairs) {
    for (const f of p.flags) {
      const kind = f.msg.replace(/（.*$/, '');
      const g = (flagGroups[kind] ??= { level: f.level, count: 0, examples: [] });
      g.count++;
      if (g.examples.length < 5) g.examples.push(`${tA.excelRow[p.a]}行目 ⟷ ${tB.excelRow[p.b]}行目: ${f.msg}`);
    }
  }
  return {
    result: r,
    mode,
    plan: tiers.map((t, i) => ({ order: i + 1, name: t.name })),
    matched: r.pairs.length,
    perTier: r.perTier,
    unmatchedBase: r.unmatchedA.map((i) => rowRef(tA, i)),
    unmatchedOther: r.unmatchedB.map((i) => rowRef(tB, i)),
    duplicatesBase: r.dupA.map(([k, n]) => ({ key: k.split(SEP).join('＋'), count: n })),
    duplicatesOther: r.dupB.map(([k, n]) => ({ key: k.split(SEP).join('＋'), count: n })),
    ambiguous: r.ambiguous.map((x) => ({ base: rowRef(tA, x.a), candidates: x.candidatesB.map((j) => rowRef(tB, j)), tier: x.tierIndex + 1 })),
    conflicts: [...new Map(r.conflicts.map((x) => [`${x.a}|${x.b}`, x])).values()].map((x) => ({ base: rowRef(tA, x.a), other: rowRef(tB, x.b), message: x.msg })),
    flags: flagGroups,
    suggestions: r.suggestions.map((s) => ({ base: rowRef(tA, s.a), other: rowRef(tB, s.b), score: r3(s.score), reasons: s.reasons })),
  };
}
