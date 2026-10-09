/**
 * The row-rearrangement angle of the consistency oracle (task V1): rows of the 2nd file are matched to the base by KEY, so merged row i is
 * not source row i. For every imported formula that stayed a formula (FormulaTrace outcome 'kept') the value Excel computes in the merged
 * sheet must equal the value of the same formula in the SOURCE file (opened and recalculated by Excel) for the SAME PERSON; and the person
 * of every placed row must agree on at least one key pair (JoinResult.keyPairs) — i.e. the row mapping itself is right.
 * Ranges that were widened to the merged table are compared with the source after Excel deleted the rows that were not placed (Excel
 * shrinks the ranges the same way); results that are POSITIONS (MATCH on its own, ROWS …) or that read a changed header cell are skipped
 * and counted (tests/xlsx/formula-import-findings.md §5 explains the same rules).
 */
import fs from 'node:fs';
import path from 'node:path';
import type { Merged } from './lib';
import { norm, type BookDump, type XVal } from './excel';
import { sameValue, type Finding } from './check';

export interface SrcCell {
  sourceId: string;
  srcRow: number;
  srcCol: number;
  row: number;
  col: number;
  widened: boolean;
  exact: boolean;
  source: string;
  formula: string;
}

export interface SrcCheck {
  combo: string;
  sourceId: string;
  sourceSheet: string;
  mergedSheet: string;
  /** dump job ids */
  twinJob: string;
  reducedJob?: string;
  deleteRows: number[];
  cells: SrcCell[];
  /** every placed source row sits on the same row number in the merged sheet */
  identical: boolean;
  placed: Array<[number, number]>;
  /** [baseCol, sourceCol] pairs of the enabled key tiers */
  keyPairs: Array<[number, number]>;
  stats: { formulas: number; kept: number };
}

const safe = (s: string) => s.replace(/[^A-Za-z0-9_.-]/g, '_');

/** the source rows that were placed into the merged sheet: [source row, merged row] */
export function placedRows(merged: Merged, sid: string): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  let appendAt = merged.base.region.lastDataRow + 1;
  for (const r of merged.join.rows) {
    const target = r.baseRow == null ? appendAt++ : r.baseRow;
    const m = r.matches?.[sid];
    const sr = m ? m.row : r.appendedFrom && r.appendedFrom.sourceId === sid ? r.appendedFrom.row : null;
    if (sr != null) out.push([sr, target]);
  }
  return out;
}

/** Writes the source files (the original bytes) into pkgDir and describes what to compare. */
export function buildSrcChecks(combo: string, merged: Merged, pkgDir: string, strictKeys: boolean): SrcCheck[] {
  const checks: SrcCheck[] = [];
  for (const src of merged.others) {
    const placed = placedRows(merged, src.id);
    const placedSet = new Set(placed.map((p) => p[0]));
    const r = src.region;
    const del: number[] = [];
    const grid = src.book.sheets.find((x) => x.name === r.sheet);
    for (let row = r.firstDataRow; row <= r.lastDataRow; row++) if (!placedSet.has(row) && grid?.rows[row - 1]?.some((c) => c && (c.v != null || c.f))) del.push(row);
    const kept = merged.trace.filter((t) => t.sourceId === src.id && t.outcome === 'kept');
    const cells: SrcCell[] = kept.map((t) => ({ sourceId: src.id, srcRow: t.srcRow, srcCol: t.srcCol, row: t.row, col: t.col, widened: !!t.widened, exact: t.exact !== false, source: t.source, formula: t.formula ?? '' }));
    const file = path.join(pkgDir, `${safe(combo)}.src-${src.id}.xlsx`);
    fs.writeFileSync(file, src.book.bytes);
    // the person of a placed row is checked by key text only where the keys are written the same way in both files (IDs); the sample files
    // write them differently on purpose (000142 / 142, full-width, 異体字: matching them is the keys module's job)
    const keyPairs: Array<[number, number]> = strictKeys ? (merged.join.keyPairs?.[src.id] ?? []).map((p) => [p.baseCol, p.sourceCol]) : [];
    checks.push({
      combo, sourceId: src.id, sourceSheet: r.sheet, mergedSheet: merged.result.layout.sheet,
      twinJob: `${safe(combo)}.src-${src.id}`,
      ...(cells.some((c) => c.widened) && del.length ? { reducedJob: `${safe(combo)}.src-${src.id}.red` } : {}),
      deleteRows: del.sort((a, b) => b - a),
      cells, identical: placed.every(([s, t]) => s === t), placed, keyPairs,
      stats: { formulas: merged.trace.filter((t) => t.sourceId === src.id).length, kept: kept.length },
    });
  }
  return checks;
}

const POSITION_ONLY = /\bMATCH\(/;

export interface SrcStats {
  keptCompared: number;
  widenedCompared: number;
  skipped: number;
  personRows: number;
}

export function judgeSrc(chk: SrcCheck, pre: BookDump, twin: BookDump, reduced: BookDump | null): { findings: Finding[]; stats: SrcStats } {
  const findings: Finding[] = [];
  const stats: SrcStats = { keptCompared: 0, widenedCompared: 0, skipped: 0, personRows: 0 };
  const add = (cls: string, detail: string) => findings.push({ scenario: chk.combo, cls, detail });
  if (pre.open !== 'normal' || twin.open !== 'normal') {
    add('open-src', `pre ${pre.open} twin ${twin.open}`);
    return { findings, stats };
  }
  const sheetOf = (d: BookDump, name: string) => d.sheets.find((s) => s.name.toLowerCase() === name.toLowerCase());
  const m = sheetOf(pre, chk.mergedSheet);
  const t = sheetOf(twin, chk.sourceSheet);
  const rd = reduced ? sheetOf(reduced, chk.sourceSheet) : undefined;
  if (!m || !t) {
    add('open-src', `sheet missing: merged ${!!m}, source ${!!t}`);
    return { findings, stats };
  }
  const at = (rows: XVal[][], r: number, c: number) => norm(rows[r - 1]?.[c - 1]);
  // ---- the person of every placed row
  if (chk.keyPairs.length) {
    let bad = 0;
    for (const [sr, mr] of chk.placed) {
      stats.personRows++;
      const agree = chk.keyPairs.some(([bc, sc]) => {
        const a = at(m.v, mr, bc);
        const b = at(t.v, sr, sc);
        return a !== null && b !== null && String(a).trim() === String(b).trim();
      });
      if (!agree) {
        bad++;
        if (bad <= 3) add('row-person', `${chk.sourceId}: source row ${sr} was placed in merged row ${mr}, but no key pair agrees (${chk.keyPairs.map(([bc, sc]) => `${JSON.stringify(at(m.v, mr, bc))} vs ${JSON.stringify(at(t.v, sr, sc))}`).join(', ')})`);
      }
    }
    if (bad > 3) add('row-person', `... ${bad - 3} more rows`);
  }
  // ---- the imported formulas
  const delBelow = (row: number) => row - chk.deleteRows.filter((d) => d < row).length;
  let diffs = 0;
  for (const c of chk.cells) {
    if (!c.exact && !c.widened) { stats.skipped++; continue; }
    if (c.widened && /\b(ROWS|COUNTBLANK)\(/.test(c.source)) { stats.skipped++; continue; }
    if (c.widened && /\b(UNIQUE|FILTER|SORT|SORTBY|TAKE)\(/.test(c.source)) { stats.skipped++; continue; }
    if (c.widened && /\bCOUNTA\(/.test(c.source)) { stats.skipped++; continue; }
    if (POSITION_ONLY.test(c.source) && !/\bINDEX\(/.test(c.source) && !chk.identical) { stats.skipped++; continue; }
    if (/\b(TODAY|NOW|RAND|RANDBETWEEN)\(/i.test(c.source)) { stats.skipped++; continue; }
    const act = at(m.v, c.row, c.col);
    let exp: unknown;
    if (c.widened) {
      if (reduced && rd) {
        exp = at(rd.v, delBelow(c.srcRow), c.srcCol);
        stats.widenedCompared++;
      } else exp = at(t.v, c.srcRow, c.srcCol); // nothing was deleted: the table is the same set of rows
    } else exp = at(t.v, c.srcRow, c.srcCol);
    stats.keptCompared++;
    if (!sameValue(exp, act)) {
      diffs++;
      if (diffs <= 6) add('src-value', `${chk.sourceId} ${chk.sourceSheet}!R${c.srcRow}C${c.srcCol} (=${c.source}) -> ${chk.mergedSheet}!R${c.row}C${c.col} (=${c.formula}${c.widened ? ', widened' : ''}): source ${JSON.stringify(exp)} merged ${JSON.stringify(act)}`);
    }
  }
  if (diffs > 6) add('src-value', `... ${diffs - 6} more`);
  return { findings, stats };
}
