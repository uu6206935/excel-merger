/**
 * Pure helpers around the keys module's DTOs: key chips, plan edits (詳細画面), import column
 * selection, the summary line and the check lists. No React.
 */
import type {
  CandidateDTO,
  ColumnDTO,
  ImportColumnSpec,
  JoinPlan,
  JoinResult,
  KeyAnalysis,
  KeySource,
  NormalizeOptions,
  PairAnalysis,
  PairJoinReport,
  RowRef,
  SheetDTO,
} from '../core/contracts';
import type { CheckItem, MergedRowPlan } from '../core/types';
import type { FileEntry, ManualKey } from './types';
import { colLetter, stripExt } from './helpers';

// ---------------------------------------------------------------------------------------------
// sources

/** Loaded files whose data region is usable (a header row and at least one data row). */
export function readyFiles(files: readonly FileEntry[]): FileEntry[] {
  return files.filter((f) => f.status === 'ready' && f.book && f.sheet && f.region && f.region.lastDataRow >= f.region.firstDataRow);
}

export function toKeySource(f: FileEntry): KeySource | null {
  if (f.status !== 'ready' || !f.book || !f.region || f.region.lastDataRow < f.region.firstDataRow) return null;
  return { id: f.id, book: f.book, region: f.region, label: stripExt(f.fileName) };
}

/** Base + other sources for analysis, or null when fewer than two usable files. */
export function keySources(files: readonly FileEntry[], baseId: string | null): { base: KeySource; others: KeySource[] } | null {
  const ready = readyFiles(files);
  if (ready.length < 2) return null;
  const baseFile = ready.find((f) => f.id === baseId) ?? ready[0];
  const base = toKeySource(baseFile);
  if (!base) return null;
  const others: KeySource[] = [];
  for (const f of ready) {
    if (f.id === baseFile.id) continue;
    const s = toKeySource(f);
    if (s) others.push(s);
  }
  return { base, others };
}

/**
 * Key of the SET of files the plan belongs to (ids, sheets, regions, base). It deliberately ignores the ORDER of the files:
 * a pure reorder of the file cards keeps the key plan, the manual keys and the region choices (carryOverPlan / the manual keys
 * are kept while this key is unchanged). What depends on the order has its own key, see othersOrderKey.
 */
export function filesKey(files: readonly FileEntry[], baseId: string | null): string {
  const ready = readyFiles(files);
  const base = ready.find((f) => f.id === baseId) ?? ready[0];
  const list = ready.map((f) => [f.id, f.sheet, f.region] as const).sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0));
  return JSON.stringify({ base: base?.id ?? null, files: list });
}

/**
 * The order of the OTHER files (the base is always the first group of the output) = the order of the imported column groups:
 * the analysis pairs, the default import order, the merge preview and the rows appended by 「他ファイルにしかない行も末尾に追加」
 * all follow it. Moving the base card among the others does not change it.
 */
export function othersOrderKey(files: readonly FileEntry[], baseId: string | null): string {
  const ready = readyFiles(files);
  const base = ready.find((f) => f.id === baseId) ?? ready[0];
  return ready
    .filter((f) => f.id !== base?.id)
    .map((f) => f.id)
    .join(',');
}

/** Key of everything the analysis depends on (files, sheets, regions, base, options and the order of the other files). */
export function analysisInputKey(files: readonly FileEntry[], baseId: string | null, options: NormalizeOptions): string {
  return filesKey(files, baseId) + '|' + othersOrderKey(files, baseId) + '|' + JSON.stringify(options);
}

/** True while some file is still being read (the analysis waits for all of them). */
export function anyLoading(files: readonly FileEntry[]): boolean {
  return files.some((f) => f.status === 'loading');
}

/** Effective base id: the chosen one when usable, else the first ready file. */
export function effectiveBaseId(files: readonly FileEntry[], baseId: string | null): string | null {
  const ready = readyFiles(files);
  if (ready.some((f) => f.id === baseId)) return baseId;
  return ready[0]?.id ?? baseId ?? files[0]?.id ?? null;
}

// ---------------------------------------------------------------------------------------------
// candidates / plan

export function findCandidate(analysis: KeyAnalysis, sourceId: string, candidateId: string): CandidateDTO | undefined {
  return analysis.pairs[sourceId]?.candidates.find((c) => c.id === candidateId);
}

export function pairIds(analysis: KeyAnalysis): string[] {
  return Object.keys(analysis.pairs);
}

export interface ChipView {
  family: 'id' | 'name';
  label: string;
  /** some pair has a candidate of this family */
  available: boolean;
  /** every available pair uses this family */
  on: boolean;
  /** only some pairs use it */
  partial: boolean;
  /** per-source lines for the tooltip */
  details: string[];
}

function familyUsed(analysis: KeyAnalysis, plan: JoinPlan, sourceId: string, family: CandidateDTO['family']): boolean {
  const tiers = plan.tiers[sourceId] ?? [];
  return tiers.some((id) => findCandidate(analysis, sourceId, id)?.family === family);
}

const pct = (r: number) => `${Math.round(Math.max(0, Math.min(1, r)) * 100)}%`;

/** The ID / 氏名 chips (always both, in that order). */
export function chipViews(analysis: KeyAnalysis, plan: JoinPlan, labels: Record<string, string>): ChipView[] {
  const fams: Array<'id' | 'name'> = ['id', 'name'];
  return fams.map((family) => {
    const chip = analysis.chips.find((c) => c.family === family);
    const label = chip?.label ?? (family === 'id' ? 'ID' : '氏名');
    const srcs = pairIds(analysis).filter((s) => chip?.perSource[s]);
    const used = srcs.filter((s) => familyUsed(analysis, plan, s, family));
    const details = pairIds(analysis).map((s) => {
      const p = chip?.perSource[s];
      const name = labels[s] ?? s;
      if (!p) return `${name}: 該当する列なし`;
      return `${name}: ${p.label} 一致率 ${pct(p.matchRate)}${familyUsed(analysis, plan, s, family) ? '' : '（使わない）'}`;
    });
    return {
      family,
      label,
      available: srcs.length > 0,
      on: srcs.length > 0 && used.length === srcs.length,
      partial: used.length > 0 && used.length < srcs.length,
      details,
    };
  });
}

/** Rank used to insert a re-enabled candidate at its natural priority. */
function rankOf(c: CandidateDTO | undefined): number {
  if (!c) return Number.MAX_SAFE_INTEGER;
  if (c.tier !== undefined) return c.tier;
  return 1000 - c.score;
}

function insertByRank(analysis: KeyAnalysis, sourceId: string, tiers: string[], candidateId: string): string[] {
  if (tiers.includes(candidateId)) return tiers;
  const r = rankOf(findCandidate(analysis, sourceId, candidateId));
  const out = tiers.slice();
  let at = out.length;
  for (let i = 0; i < out.length; i++) {
    if (rankOf(findCandidate(analysis, sourceId, out[i])) > r) {
      at = i;
      break;
    }
  }
  out.splice(at, 0, candidateId);
  return out;
}

/** Chip toggle: enable / disable a key family in every pair. */
export function toggleFamily(analysis: KeyAnalysis, plan: JoinPlan, family: 'id' | 'name', on: boolean): JoinPlan {
  const chip = analysis.chips.find((c) => c.family === family);
  const tiers: Record<string, string[]> = {};
  for (const s of pairIds(analysis)) {
    let t = plan.tiers[s] ?? [];
    if (!on) {
      t = t.filter((id) => findCandidate(analysis, s, id)?.family !== family);
    } else if (!familyUsed(analysis, plan, s, family)) {
      const best = chip?.perSource[s]?.candidateId;
      if (best && findCandidate(analysis, s, best)) t = insertByRank(analysis, s, t, best);
    }
    tiers[s] = t;
  }
  return { ...plan, tiers: { ...plan.tiers, ...tiers } };
}

/** 使う checkbox in the detail table (new candidates go to the end = lowest priority). */
export function toggleCandidate(plan: JoinPlan, sourceId: string, candidateId: string, on: boolean): JoinPlan {
  const cur = plan.tiers[sourceId] ?? [];
  const next = on ? (cur.includes(candidateId) ? cur : [...cur, candidateId]) : cur.filter((id) => id !== candidateId);
  return { ...plan, tiers: { ...plan.tiers, [sourceId]: next } };
}

/** 優先 ↑↓ */
export function moveCandidate(plan: JoinPlan, sourceId: string, candidateId: string, dir: -1 | 1): JoinPlan {
  const cur = (plan.tiers[sourceId] ?? []).slice();
  const i = cur.indexOf(candidateId);
  const j = i + dir;
  if (i < 0 || j < 0 || j >= cur.length) return plan;
  [cur[i], cur[j]] = [cur[j], cur[i]];
  return { ...plan, tiers: { ...plan.tiers, [sourceId]: cur } };
}

export function hasManualLink(plan: JoinPlan, sourceId: string, baseRow: number, otherRow: number): boolean {
  return (plan.manualLinks[sourceId] ?? []).some((l) => l.baseRow === baseRow && l.otherRow === otherRow);
}

/** 「この組み合わせで結合」 (and its undo). A base row / other row can only be linked once. */
export function setManualLink(plan: JoinPlan, sourceId: string, baseRow: number, otherRow: number, on: boolean): JoinPlan {
  const cur = plan.manualLinks[sourceId] ?? [];
  const rest = cur.filter((l) => l.baseRow !== baseRow && l.otherRow !== otherRow);
  const next = on ? [...rest, { baseRow, otherRow }] : cur.filter((l) => !(l.baseRow === baseRow && l.otherRow === otherRow));
  return { ...plan, manualLinks: { ...plan.manualLinks, [sourceId]: next } };
}

/** Candidates of a pair in display order: used ones in priority order, then the rest by score. */
export function orderedCandidates(pair: PairAnalysis, plan: JoinPlan): Array<{ c: CandidateDTO; priority: number | null }> {
  const tiers = plan.tiers[pair.sourceId] ?? [];
  const used = tiers
    .map((id) => pair.candidates.find((c) => c.id === id))
    .filter((c): c is CandidateDTO => !!c)
    .map((c, i) => ({ c, priority: i + 1 }));
  const rest = pair.candidates
    .filter((c) => !tiers.includes(c.id))
    .sort((a, b) => b.score - a.score)
    .map((c) => ({ c, priority: null }));
  return [...used, ...rest];
}

export interface ExtraChip {
  sourceId: string;
  candidateId: string;
  label: string;
  on: boolean;
}

/** Non-ID/氏名 keys that are (or were auto-) selected, shown as extra chips after ID / 氏名. */
export function extraChips(analysis: KeyAnalysis, plan: JoinPlan): ExtraChip[] {
  const out: ExtraChip[] = [];
  const seen = new Set<string>();
  for (const s of pairIds(analysis)) {
    const tiers = plan.tiers[s] ?? [];
    for (const c of analysis.pairs[s].candidates) {
      if (c.family !== 'other') continue;
      const on = tiers.includes(c.id);
      if (!on && !c.selected) continue;
      const key = `${s}:${c.id}`;
      if (seen.has(key)) continue;
      seen.add(key);
      out.push({ sourceId: s, candidateId: c.id, label: c.label, on });
    }
  }
  return out;
}

/** 「＋ 他の候補 N件」: usable candidates that are neither used nor auto-selected. */
export function otherCandidateCount(analysis: KeyAnalysis, plan: JoinPlan): number {
  let n = 0;
  for (const s of pairIds(analysis)) {
    const tiers = plan.tiers[s] ?? [];
    for (const c of analysis.pairs[s].candidates) {
      if (tiers.includes(c.id) || c.selected || c.verdict === '使用不可') continue;
      n++;
    }
  }
  return n;
}

/** Sources whose plan has no key at all (nothing can match). */
export function sourcesWithoutKeys(analysis: KeyAnalysis, plan: JoinPlan): string[] {
  return pairIds(analysis).filter((s) => (plan.tiers[s] ?? []).length === 0 && (plan.manualLinks[s] ?? []).length === 0);
}

/**
 * Sources that take NO part in the merge: no key is selected for them (and no manual link). Their columns are not
 * imported, they do not block the merge, and they join it again as soon as a key is selected (the set is derived
 * from the plan, there is no separate flag to maintain).
 */
export function excludedSourceIds(analysis: KeyAnalysis | null, plan: JoinPlan | null): string[] {
  if (!analysis || !plan) return [];
  return sourcesWithoutKeys(analysis, plan);
}

/**
 * The join result without the excluded sources. keys.joinSources still runs over every analysed file: a file
 * without tiers gets a 'noKey' item and (join 'full') ALL of its rows would be appended; here its report / items /
 * key fill / imports / appended rows are removed and the summary is recomputed over the remaining files.
 */
export function withoutSources(j: JoinResult, excluded: readonly string[]): JoinResult {
  if (!excluded.length) return j;
  const ex = new Set(excluded);
  const rows: MergedRowPlan[] = [];
  for (const r of j.rows) {
    if (r.baseRow === null && r.appendedFrom && ex.has(r.appendedFrom.sourceId)) continue;
    const matches: MergedRowPlan['matches'] = {};
    for (const [id, m] of Object.entries(r.matches)) if (!ex.has(id)) matches[id] = m;
    rows.push({ ...r, matches });
  }
  const reports: JoinResult['reports'] = {};
  for (const [id, r] of Object.entries(j.reports)) if (!ex.has(id)) reports[id] = r;
  const keyFill: JoinResult['keyFill'] = {};
  for (const [id, f] of Object.entries(j.keyFill)) if (!ex.has(id)) keyFill[id] = f;
  // keyPairs (premerge reads the base's partner of a non-imported key column of an imported formula) only for the files that take part
  let keyPairs: JoinResult['keyPairs'];
  if (j.keyPairs) {
    keyPairs = {};
    for (const [id, p] of Object.entries(j.keyPairs)) if (!ex.has(id)) keyPairs[id] = p;
  }
  const ids = Object.keys(reports);
  let matchedAny = 0;
  let matchedAll = 0;
  let appended = 0;
  for (const r of rows) {
    if (r.baseRow === null) {
      appended++;
      continue;
    }
    const n = ids.filter((id) => r.matches[id]).length;
    if (n > 0) matchedAny++;
    if (ids.length > 0 && n === ids.length) matchedAll++;
  }
  return {
    rows,
    reports,
    summary: { baseRows: j.summary.baseRows, matchedAny, matchedAll, appended },
    items: (j.items ?? []).filter((it) => !(it.sourceId && ex.has(it.sourceId))),
    keyFill,
    defaultImports: j.defaultImports.filter((d) => !ex.has(d.sourceId)),
    ...(keyPairs ? { keyPairs } : {}),
  };
}

/**
 * JoinResult.keyPairs restricted to the sources that take part in the merge (PremergeInput.keyPairs). undefined when the join has none
 * (older / mock join results): premerge then substitutes nothing.
 */
export function keyPairsFor(j: Pick<JoinResult, 'keyPairs'> | null | undefined, sourceIds: readonly string[]): NonNullable<JoinResult['keyPairs']> | undefined {
  if (!j?.keyPairs) return undefined;
  const ids = new Set(sourceIds);
  const out: NonNullable<JoinResult['keyPairs']> = {};
  for (const [id, pairs] of Object.entries(j.keyPairs)) if (ids.has(id) && pairs.length) out[id] = pairs.map((p) => ({ baseCol: p.baseCol, sourceCol: p.sourceCol }));
  return out;
}

/** Number of files that really take part in the join (reports of the stripped join result). */
export function joinedSourceCount(j: JoinResult | null): number {
  return j ? Object.keys(j.reports).length : 0;
}

/** 'ID → 氏名＋大学' */
export function tierLabels(analysis: KeyAnalysis, plan: JoinPlan, sourceId: string): string {
  const labels = (plan.tiers[sourceId] ?? []).map((id) => {
    const c = findCandidate(analysis, sourceId, id);
    return c ? shortLabel(c) : id;
  });
  return labels.length ? labels.join(' → ') : 'なし';
}

/** Short label of a candidate for tables ('ID', '氏名＋大学'); falls back to the full label. */
export function shortLabel(c: CandidateDTO): string {
  if (c.kind === 'composite') return c.label;
  if (c.family === 'id') return 'ID';
  if (c.family === 'name') return c.baseColumns.length > 1 ? c.label : '氏名';
  return c.label.split(/[（(]/)[0] || c.label;
}

/** Verdict -> css tone */
export function verdictTone(verdict: string): 'good' | 'partial' | 'poor' | 'bad' {
  if (verdict === '良好') return 'good';
  if (verdict === '一部一致') return 'partial';
  if (verdict === '不一致多数') return 'poor';
  return 'bad';
}

/**
 * Preserve the user's choices that are still meaningful when a fresh default plan arrives.
 *  - join type / duplicate policy: always kept.
 *  - key tiers and manual links: kept per pair only while the files / sheets / regions / base are the same
 *    (`sameFiles`, i.e. only the normalisation options changed) and every chosen candidate still exists.
 */
export function carryOverPlan(fresh: JoinPlan, previous: JoinPlan | null, opts: { sameFiles: boolean; analysis: KeyAnalysis }): JoinPlan {
  if (!previous) return fresh;
  const tiers: JoinPlan['tiers'] = { ...fresh.tiers };
  const manualLinks: JoinPlan['manualLinks'] = { ...fresh.manualLinks };
  if (opts.sameFiles) {
    for (const s of Object.keys(fresh.tiers)) {
      const prev = previous.tiers[s];
      const pair = opts.analysis.pairs[s];
      if (prev && pair && prev.every((id) => pair.candidates.some((c) => c.id === id))) tiers[s] = prev.slice();
      if (previous.manualLinks[s]) manualLinks[s] = previous.manualLinks[s];
    }
  }
  return { ...fresh, tiers, manualLinks, joinType: previous.joinType, duplicatePolicy: previous.duplicatePolicy };
}

export interface FileKeyColumn {
  /** 'ID' / '氏名' / '氏名＋大学' … */
  label: string;
  header: string;
  col: number;
}

/**
 * Columns a file uses as keys in the current plan (card line 「ID: 学生ID / 氏名: 氏名」). For the base file
 * the base columns of every pair's keys, for another file the columns of its own pair.
 */
export function fileKeyColumns(analysis: KeyAnalysis | null, plan: JoinPlan | null, fileId: string): FileKeyColumn[] {
  if (!analysis || !plan) return [];
  const out: FileKeyColumn[] = [];
  const seen = new Set<number>();
  const add = (c: CandidateDTO, cols: Array<{ col: number; header: string }>) => {
    const label = c.kind === 'composite' ? c.label : c.family === 'id' ? 'ID' : c.family === 'name' ? '氏名' : shortLabel(c);
    for (const x of cols) {
      if (seen.has(x.col)) continue;
      seen.add(x.col);
      out.push({ label, header: x.header, col: x.col });
    }
  };
  if (fileId === analysis.baseId) {
    for (const s of pairIds(analysis)) for (const id of plan.tiers[s] ?? []) {
      const c = findCandidate(analysis, s, id);
      if (c) add(c, c.baseColumns);
    }
  } else if (analysis.pairs[fileId]) {
    for (const id of plan.tiers[fileId] ?? []) {
      const c = findCandidate(analysis, fileId, id);
      if (c) add(c, c.otherColumns);
    }
  }
  return out;
}

/** 「ID: 学生ID / 氏名: 氏名」 */
export function fileKeyLine(cols: readonly FileKeyColumn[]): string {
  return cols.map((k) => (k.label ? `${k.label}: ${k.header}` : k.header)).join(' / ');
}

// ---------------------------------------------------------------------------------------------
// import columns (取り込む列)

export function importKey(sourceId: string, sheet: string, col: number): string {
  return `${sourceId}|${sheet}|${col}`;
}

export interface ImportChoice {
  sourceId: string;
  col: number;
  header: string;
  isDefault: boolean;
  on: boolean;
  /** column is a key of the current plan */
  isKey: boolean;
}

/** All data columns of one other source with their checkbox state. */
export function importChoices(
  analysis: KeyAnalysis,
  plan: JoinPlan | null,
  joinResult: JoinResult | null,
  overrides: Record<string, boolean>,
  sourceId: string,
): ImportChoice[] {
  const pair = analysis.pairs[sourceId];
  if (!pair) return [];
  const sheet = pair.other.sheet;
  const defaults = new Set((joinResult?.defaultImports ?? []).filter((d) => d.sourceId === sourceId).map((d) => d.col));
  const keyCols = new Set<number>();
  for (const id of plan?.tiers[sourceId] ?? []) {
    const c = pair.candidates.find((x) => x.id === id);
    for (const oc of c?.otherColumns ?? []) keyCols.add(oc.col);
  }
  return pair.other.columns.map((col) => {
    const isDefault = defaults.has(col.col);
    const ov = overrides[importKey(sourceId, sheet, col.col)];
    return {
      sourceId,
      col: col.col,
      header: col.header || `（${col.col}列目）`,
      isDefault,
      on: ov ?? isDefault,
      isKey: keyCols.has(col.col),
    };
  });
}

/** The import list handed to buildPremerged (source order, column order). */
export function effectiveImports(
  analysis: KeyAnalysis | null,
  plan: JoinPlan | null,
  joinResult: JoinResult | null,
  overrides: Record<string, boolean>,
): ImportColumnSpec[] {
  if (!analysis || !joinResult) return [];
  // the keys module may give a column its own header text (a two-row header 「面接 1次」, a merged group title): the premerge would
  // not find it in the header cell itself, so it travels with the import (ImportColumnSpec.header)
  const headerOverride = new Map<string, string>();
  for (const d of joinResult.defaultImports) if (d.header) headerOverride.set(`${d.sourceId}|${d.col}`, d.header);
  const out: ImportColumnSpec[] = [];
  const excluded = new Set(excludedSourceIds(analysis, plan));
  for (const s of pairIds(analysis)) {
    if (excluded.has(s)) continue; // no key: the file takes no part in the merge
    for (const ch of importChoices(analysis, plan, joinResult, overrides, s)) {
      if (!ch.on) continue;
      const header = headerOverride.get(`${s}|${ch.col}`);
      out.push(header ? { sourceId: s, col: ch.col, header } : { sourceId: s, col: ch.col });
    }
  }
  return out;
}

/** Columns of the merged table as planned before the merge (base columns, then the imported ones). */
export interface PlannedColumn {
  key: string;
  sourceId: string;
  header: string;
  imported: boolean;
}

export function plannedColumns(
  analysis: KeyAnalysis | null,
  plan: JoinPlan | null,
  joinResult: JoinResult | null,
  overrides: Record<string, boolean>,
): PlannedColumn[] {
  if (!analysis) return [];
  const out: PlannedColumn[] = analysis.base.columns.map((c) => ({
    key: `b${c.col}`,
    sourceId: analysis.baseId,
    header: c.header || `（${c.col}列目）`,
    imported: false,
  }));
  const existing = new Set(out.map((c) => c.header));
  const excluded = new Set(excludedSourceIds(analysis, plan));
  for (const s of pairIds(analysis)) {
    if (excluded.has(s)) continue;
    for (const ch of importChoices(analysis, plan, joinResult, overrides, s)) {
      if (!ch.on) continue;
      // same rule as the premerge's headerConflict 'suffix' (shown only as a preview)
      const header = existing.has(ch.header) ? `${ch.header}（${analysis.pairs[s].other.file.replace(/\.[^.]+$/, '')}）` : ch.header;
      existing.add(header);
      out.push({ key: `i:${s}:${ch.col}`, sourceId: s, header, imported: true });
    }
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// summary + checks

export interface JoinSummary {
  baseRows: number;
  /** base rows found in EVERY other file */
  matched: number;
  /** base rows found in some but not all other files (always 0 with one other file) */
  partial: number;
  /** base rows found in NO other file */
  unmatched: number;
  duplicates: number;
  review: number;
  appended: number;
}

function dupGroups(r: PairJoinReport): number {
  return r.duplicatesOther.length + (r.mode === 'oneToOne' ? r.duplicatesBase.length : 0);
}

function reviewCount(r: PairJoinReport): number {
  let n = r.conflicts.length + r.ambiguous.length;
  for (const f of Object.values(r.flags)) if (f.level === 'error' || f.level === 'warn' || f.level === 'warning') n += f.count;
  return n;
}

export function joinSummary(j: JoinResult): JoinSummary {
  const reports = Object.values(j.reports);
  const baseRows = j.summary.baseRows;
  const matched = reports.length > 1 ? j.summary.matchedAll : j.summary.matchedAny;
  return {
    baseRows,
    matched,
    partial: reports.length > 1 ? Math.max(0, j.summary.matchedAny - j.summary.matchedAll) : 0,
    unmatched: Math.max(0, baseRows - j.summary.matchedAny),
    duplicates: reports.reduce((n, r) => n + dupGroups(r), 0),
    review: reports.reduce((n, r) => n + reviewCount(r), 0),
    appended: j.summary.appended,
  };
}

/**
 * 「基準 30 名 / 一致 28 / 未一致 2 / 重複 1」. With several other files: 一致 = found in every file,
 * 一部一致 = found in some of them, 未一致 = found in none.
 */
export function summaryText(s: JoinSummary, unit: '名' | '行', excluded = 0): string {
  const n = (x: number) => x.toLocaleString('ja-JP');
  let t = `基準 ${n(s.baseRows)} ${unit} / 一致 ${n(s.matched)}`;
  if (s.partial > 0) t += ` / 一部一致 ${n(s.partial)}`;
  t += ` / 未一致 ${n(s.unmatched)} / 重複 ${n(s.duplicates)}`;
  if (s.review > 0) t += ` / 要確認 ${n(s.review)}`;
  if (s.appended > 0) t += ` / 追加 ${n(s.appended)}`;
  if (excluded > 0) t += ` / 結合対象外 ${n(excluded)} ファイル`;
  return t;
}

/** Base rows (Excel rows) that no other file matched — tinted in the grid. */
export function unmatchedEverywhere(j: JoinResult | null): Set<number> {
  const out = new Set<number>();
  if (!j) return out;
  const reports = Object.values(j.reports);
  if (!reports.length) return out;
  const count = new Map<number, number>();
  for (const r of reports) for (const u of r.unmatchedBase) count.set(u.excelRow, (count.get(u.excelRow) ?? 0) + 1);
  for (const [row, n] of count) if (n >= reports.length) out.add(row);
  return out;
}

/** Person lists count in 名; anything without an ID / 氏名 key in 行. */
export function summaryUnit(analysis: KeyAnalysis | null, plan: JoinPlan | null): '名' | '行' {
  if (!analysis || !plan) return '行';
  for (const s of pairIds(analysis)) {
    for (const id of plan.tiers[s] ?? []) {
      const f = findCandidate(analysis, s, id)?.family;
      if (f === 'id' || f === 'name') return '名';
    }
  }
  return '行';
}

export interface CheckRow {
  key: string;
  /** Excel row in the merged sheet to jump to (null = not in the merged sheet) */
  row: number | null;
  text: string;
  sub?: string;
  tone: 'error' | 'warning' | 'info';
}

export interface CheckGroup {
  id: 'unmatched' | 'partial' | 'duplicate' | 'review' | 'toValues' | 'info' | 'otherOnly';
  title: string;
  count: number;
  rows: CheckRow[];
  /**
   * summary lines of the keys module's check items that point into the OTHER file (CheckItem.otherRows), shown
   * above the rows (「「02_面接評価」でキーが重複している行: 2件」 · 「02_面接評価 の 12行目・13行目」); not counted
   */
  summaries?: CheckRow[];
}

/** '12行目・15行目・40行目 ほか 3 件' */
export function rowsList(rows: readonly number[], max = 5): string {
  const uniq = [...new Set(rows.filter((r) => Number.isFinite(r)))];
  const shown = uniq.slice(0, max).map((r) => `${r}行目`).join('・');
  return uniq.length > max ? `${shown} ほか ${uniq.length - max} 件` : shown;
}

/**
 * Where a check item points to: 「基準の 5行目・9行目」 and / or 「02_面接評価 の 12行目・13行目」 (otherRows are rows of
 * the OTHER file named by sourceId). undefined when the item has no rows.
 */
export function checkItemWhere(item: CheckItem, labels: Record<string, string>): string | undefined {
  const parts: string[] = [];
  if (item.otherRows?.length) parts.push(`${(item.sourceId && labels[item.sourceId]) || 'このファイル'} の ${rowsList(item.otherRows)}`);
  if (item.rows?.length) parts.push(`基準の ${rowsList(item.rows)}`);
  return parts.length ? parts.join(' / ') : undefined;
}

/** Group of the side panel that a keys check item belongs to (by its code). */
function groupOfItem(code: string): CheckGroup['id'] {
  switch (code) {
    case 'duplicate':
      return 'duplicate';
    case 'unmatchedOther':
      return 'otherOnly';
    case 'unmatched':
      return 'unmatched';
    case 'variant':
      return 'info';
    default:
      return 'review';
  }
}

const previewText = (r: RowRef, max = 3) =>
  r.preview
    .filter((t) => t !== '')
    .slice(0, max)
    .join(' ');

/**
 * Check lists for the side panel. `labels` maps source id -> display name.
 * Base rows keep their Excel row in the merged sheet (only columns are inserted by the merge).
 */
export function checkGroups(j: JoinResult, labels: Record<string, string>, toValues?: { count: number; rows: CheckRow[] }): CheckGroup[] {
  const reports = Object.values(j.reports);
  const multi = reports.length > 1;

  // 未一致: base rows that some file did not match
  const missing = new Map<number, { ref: RowRef; files: string[] }>();
  for (const r of reports) {
    for (const u of r.unmatchedBase) {
      const e = missing.get(u.excelRow) ?? { ref: u, files: [] };
      e.files.push(labels[r.sourceId] ?? r.sourceId);
      missing.set(u.excelRow, e);
    }
  }
  const sorted = [...missing.values()].sort((a, b) => a.ref.excelRow - b.ref.excelRow);
  // 未一致 = found in no other file; 一部のファイルに該当なし = missing in some of them (several files only)
  const unmatchedRows: CheckRow[] = sorted
    .filter((e) => e.files.length >= reports.length)
    .map((e) => ({
      key: `u${e.ref.excelRow}`,
      row: e.ref.excelRow,
      text: `${e.ref.excelRow}行目 ${previewText(e.ref)}`,
      sub: multi ? 'どのファイルにも該当なし' : '該当する行なし',
      tone: 'warning',
    }));
  const partialRows: CheckRow[] = sorted
    .filter((e) => e.files.length < reports.length)
    .map((e) => ({
      key: `p${e.ref.excelRow}`,
      row: e.ref.excelRow,
      text: `${e.ref.excelRow}行目 ${previewText(e.ref)}`,
      sub: `${e.files.join('・')} に該当なし`,
      tone: 'info',
    }));

  const dupRows: CheckRow[] = [];
  for (const r of reports) {
    const name = labels[r.sourceId] ?? r.sourceId;
    for (const d of r.duplicatesOther) {
      dupRows.push({ key: `do${r.sourceId}${d.key}`, row: null, text: `${d.key}（${d.count} 行）`, sub: `${name} で重複`, tone: 'warning' });
    }
    if (r.mode === 'oneToOne') {
      for (const d of r.duplicatesBase) {
        dupRows.push({ key: `db${r.sourceId}${d.key}`, row: null, text: `${d.key}（${d.count} 行）`, sub: '基準ファイルで重複', tone: 'warning' });
      }
    }
  }

  const review: CheckRow[] = [];
  const info: CheckRow[] = [];
  // 要確認 is counted like the summary line (joinSummary.review): every conflict / ambiguous row once, a flag by its count
  let reviewCount = 0;
  for (const r of reports) {
    const name = labels[r.sourceId] ?? r.sourceId;
    r.conflicts.forEach((c, i) => {
      reviewCount++;
      review.push({ key: `c${r.sourceId}${i}`, row: c.base.excelRow, text: `${c.base.excelRow}行目 ${previewText(c.base, 2)}`, sub: `${name}: ${c.message}`, tone: 'error' });
    });
    r.ambiguous.forEach((a, i) => {
      reviewCount++;
      review.push({
        key: `a${r.sourceId}${i}`,
        row: a.base.excelRow,
        text: `${a.base.excelRow}行目 ${previewText(a.base, 2)}`,
        sub: `${name}: 候補が ${a.candidates.length} 行（${a.candidates.map((c) => c.excelRow + '行目').join('、')}）`,
        tone: 'warning',
      });
    });
    for (const [kind, f] of Object.entries(r.flags)) {
      const serious = f.level === 'error' || f.level === 'warn' || f.level === 'warning';
      const target = serious ? review : info;
      if (serious) reviewCount += f.count;
      target.push({
        key: `f${r.sourceId}${kind}`,
        row: firstRowOf(f.examples[0]),
        text: `${kind}: ${f.count} 件`,
        sub: `${name}${f.examples.length ? ' — ' + f.examples.slice(0, 2).join(' / ') : ''}`,
        tone: f.level === 'error' ? 'error' : serious ? 'warning' : 'info',
      });
    }
  }

  const otherOnly: CheckRow[] = [];
  for (const r of reports) {
    const name = labels[r.sourceId] ?? r.sourceId;
    for (const u of r.unmatchedOther) {
      otherOnly.push({ key: `o${r.sourceId}${u.excelRow}`, row: null, text: `${name} の ${u.excelRow}行目`, sub: previewText(u) || undefined, tone: 'info' });
    }
  }

  const groups: CheckGroup[] = [
    { id: 'unmatched', title: '未一致', count: unmatchedRows.length, rows: unmatchedRows },
    ...(multi ? [{ id: 'partial' as const, title: '一部のファイルに該当なし', count: partialRows.length, rows: partialRows }] : []),
    { id: 'duplicate', title: '重複', count: dupRows.length, rows: dupRows },
    { id: 'review', title: '要確認', count: reviewCount, rows: review },
  ];
  if (toValues) groups.push({ id: 'toValues', title: '数式→値', count: toValues.count, rows: toValues.rows });
  groups.push({ id: 'otherOnly', title: '他ファイルにしかない行', count: otherOnly.length, rows: otherOnly });
  if (info.length) groups.push({ id: 'info', title: '補正した表記の違い', count: info.length, rows: info });

  // check items of the keys module that point into the other file (CheckItem.otherRows)
  (j.items ?? []).forEach((it, i) => {
    if (!it.otherRows?.length) return;
    const g = groups.find((x) => x.id === groupOfItem(it.code)) ?? groups.find((x) => x.id === 'review')!;
    (g.summaries ??= []).push({
      key: `it${i}`,
      row: it.rows?.length ? it.rows[0] : null,
      text: it.message,
      sub: checkItemWhere(it, labels),
      tone: it.severity,
    });
  });
  return groups;
}

/** '17行目 ⟷ 16行目: …' -> 17 */
function firstRowOf(example: string | undefined): number | null {
  if (!example) return null;
  const m = /^(\d+)行目/.exec(example);
  return m ? Number(m[1]) : null;
}

/** Per-file row of the 結合プレビュー table. */
export interface PreviewRow {
  sourceId: string;
  isBase: boolean;
  /** no key selected: the file takes no part in the merge */
  excluded: boolean;
  rows: number;
  keys: string;
  matched: number | null;
  unmatched: number | null;
  duplicates: number | null;
  blanks: number | null;
}

export function previewRows(analysis: KeyAnalysis, plan: JoinPlan, j: JoinResult | null): PreviewRow[] {
  const excluded = new Set(excludedSourceIds(analysis, plan));
  const out: PreviewRow[] = [
    {
      sourceId: analysis.baseId,
      isBase: true,
      excluded: false,
      rows: analysis.base.rows,
      keys: '基準',
      matched: null,
      unmatched: j && joinedSourceCount(j) > 0 ? joinSummary(j).unmatched : null,
      duplicates: null,
      blanks: null,
    },
  ];
  for (const s of pairIds(analysis)) {
    const pair = analysis.pairs[s];
    if (excluded.has(s)) {
      out.push({ sourceId: s, isBase: false, excluded: true, rows: pair.other.rows, keys: '結合対象外', matched: null, unmatched: null, duplicates: null, blanks: null });
      continue;
    }
    const rep = j?.reports[s];
    const first = findCandidate(analysis, s, (plan.tiers[s] ?? [])[0] ?? '');
    out.push({
      sourceId: s,
      isBase: false,
      excluded: false,
      rows: pair.other.rows,
      keys: tierLabels(analysis, plan, s),
      matched: rep ? rep.matched : null,
      unmatched: rep ? rep.unmatchedOther.length : null,
      duplicates: rep ? rep.duplicatesOther.reduce((n, d) => n + d.count, 0) : null,
      blanks: first ? first.blankOther : null,
    });
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// match rate sentence (Power Query style) under the key table of the detail modal

export interface MatchRate {
  /** rows of the other file / of the base table */
  otherRows: number;
  baseRows: number;
  /** other-file rows that found a partner in the base (rows absorbed as duplicates count as found) */
  matchedOther: number;
  /** base rows that found a row in the other file */
  matchedBase: number;
}

/** Numbers behind 「このキーで 58 / 63 行が一致します」 for one file, or null while there is no report. */
export function matchRate(analysis: KeyAnalysis, joinResult: JoinResult | null, sourceId: string): MatchRate | null {
  const pair = analysis.pairs[sourceId];
  const rep = joinResult?.reports[sourceId];
  if (!pair || !rep) return null;
  const otherRows = Math.max(0, pair.other.rows);
  const baseRows = Math.max(0, analysis.base.rows);
  const clamp = (x: number, max: number) => Math.max(0, Math.min(max, x));
  return {
    otherRows,
    baseRows,
    matchedOther: clamp(otherRows - rep.unmatchedOther.length, otherRows),
    matchedBase: clamp(baseRows - rep.unmatchedBase.length, baseRows),
  };
}

/** 「このキーで 58 / 63 行が一致します（基準 60 行中 58 行が見つかりました）」 */
export function matchRateText(m: MatchRate): string {
  const n = (x: number) => x.toLocaleString('ja-JP');
  return `このキーで ${n(m.matchedOther)} / ${n(m.otherRows)} 行が一致します（基準 ${n(m.baseRows)} 行中 ${n(m.matchedBase)} 行が見つかりました）`;
}

/** Tone of the sentence: how much of the other file / of the base was found. */
export function matchRateTone(m: MatchRate): 'good' | 'partial' | 'poor' {
  const r = m.otherRows ? m.matchedOther / m.otherRows : 0;
  return r >= 0.9 ? 'good' : r >= 0.5 ? 'partial' : 'poor';
}

// ---------------------------------------------------------------------------------------------
// manual key candidates (詳細画面 「列を選んで結合キーに追加」)

/** Candidate that the user added by hand (can be removed; automatic ones can only be unchecked). */
export function isManualCandidate(c: Pick<CandidateDTO, 'id'>, manual?: readonly Pick<ManualKey, 'id'>[]): boolean {
  return c.id.startsWith('manual:') || !!manual?.some((m) => m.id === c.id);
}

export interface ColumnOption {
  col: number;
  header: string;
  letter: string;
  /** 「学生ID（B列）」 */
  label: string;
  typeLabel: string;
}

/** Every column of a sheet's data table for the picker dropdowns: header text + column letter. */
export function columnOptions(sheet: SheetDTO): ColumnOption[] {
  return sheet.columns.map((c) => {
    const letter = colLetter(c.col);
    const header = (c.header ?? '').trim();
    return { col: c.col, header, letter, label: `${header || '（見出しなし）'}（${letter}列）`, typeLabel: c.typeLabel };
  });
}

export interface ManualKeyHint {
  /** ok: same detected type on both sides · warn: types differ · none: nothing to say yet */
  tone: 'ok' | 'warn' | 'none';
  text: string;
  /** the selection can be added */
  ready: boolean;
}

const sideText = (cols: readonly ColumnDTO[]) => cols.map((c) => `${c.header || colLetter(c.col) + '列'}（${c.typeLabel}）`).join(' ＋ ');

/**
 * Live hint under the picker: the detected type of the chosen columns, and why 「追加」 is not possible yet.
 * `base` / `other` hold the chosen columns in order (undefined = a dropdown still on 「選んでください」).
 */
export function manualKeyHint(base: ReadonlyArray<ColumnDTO | undefined>, other: ReadonlyArray<ColumnDTO | undefined>): ManualKeyHint {
  const b = base.filter((c): c is ColumnDTO => !!c);
  const o = other.filter((c): c is ColumnDTO => !!c);
  if (!b.length && !o.length) return { tone: 'none', text: '両方のファイルの列を選ぶと、検出した列の種類をここに表示します。', ready: false };
  if (b.length < base.length || o.length < other.length) return { tone: 'none', text: '「選んでください」のままの列があります。列を選ぶか、不要な列を外してください。', ready: false };
  if (!b.length || !o.length) return { tone: 'none', text: `${b.length ? 'このファイル' : '基準ファイル'}の列も選んでください。`, ready: false };
  if (b.length !== o.length) return { tone: 'none', text: `基準ファイルとこのファイルで同じ数の列を選んでください（基準 ${b.length} 列・このファイル ${o.length} 列）。`, ready: false };
  if (new Set(b.map((c) => c.col)).size < b.length || new Set(o.map((c) => c.col)).size < o.length) {
    return { tone: 'none', text: '同じ列が重複して選ばれています。', ready: false };
  }
  const same = b.every((c, i) => c.typeLabel === o[i].typeLabel);
  const text = `検出した種類: 基準 ${sideText(b)} ⟷ このファイル ${sideText(o)}`;
  return same ? { tone: 'ok', text: `${text}。同じ種類の列です。`, ready: true } : { tone: 'warn', text: `${text}。種類が違います。同じ値が入っている列か確かめてください。`, ready: true };
}

const sameCols = (a: readonly number[], b: readonly number[]) => a.length === b.length && a.every((x, i) => x === b[i]);

/** The candidate of a pair that already uses exactly these columns (same order), if any. */
export function candidateByColumns(pair: PairAnalysis | undefined, baseCols: readonly number[], otherCols: readonly number[]): CandidateDTO | undefined {
  return pair?.candidates.find((c) => sameCols(c.baseColumns.map((x) => x.col), baseCols) && sameCols(c.otherColumns.map((x) => x.col), otherCols));
}

/** The candidate that addManualCandidate added: the one id that is new, else the one with those columns. */
export function addedCandidate(before: PairAnalysis | undefined, after: PairAnalysis | undefined, baseCols: readonly number[], otherCols: readonly number[]): CandidateDTO | undefined {
  if (!after) return undefined;
  const old = new Set((before?.candidates ?? []).map((c) => c.id));
  const fresh = after.candidates.filter((c) => !old.has(c.id));
  if (fresh.length) return fresh.find((c) => sameCols(c.baseColumns.map((x) => x.col), baseCols) && sameCols(c.otherColumns.map((x) => x.col), otherCols)) ?? fresh[fresh.length - 1];
  return candidateByColumns(after, baseCols, otherCols);
}
