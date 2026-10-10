/**
 * Shared helpers of the formula-import tooling (task F1): corpus manifest, the app's own key analysis + join (src/core/keys),
 * PremergeInput assembly and the classification of imported formula cells (live / value / EMPTY).
 * Used by diagnose.ts (findings) and run.mjs (real-Excel value-invariance oracle).
 */
import fs from 'node:fs';
import path from 'node:path';
import type { ImportColumnSpec, JoinResult, KeyAnalysis, KeySource, PremergeInput, PremergeResult } from '../../../src/core/contracts';
import { analyzeSources, defaultPlan, detectRegion, joinSources, pickDataSheet } from '../../../src/core/keys';
import type { Cell, SheetGrid, SourceBook } from '../../../src/core/types';
import { readBook } from '../../../src/core/xlsx';
import { buildPremerged, type FormulaTrace } from '../../../src/core/xlsx/premerge';

export interface CorpusEntry {
  name: string;
  order: 'same' | 'rev' | 'shuf';
  file: string;
  desc: string;
  sheet: string;
  headerRow: number;
  firstDataRow: number;
  lastDataRow: number;
  comSaveXlsm?: boolean;
  /** the base workbook of the pair (default base.xlsx); base_samples.xlsx = the app's samples/01_学生名簿.xlsx */
  base?: string;
}

export function loadManifest(dir: string): CorpusEntry[] {
  return JSON.parse(fs.readFileSync(path.join(dir, 'manifest.json'), 'utf8')) as CorpusEntry[];
}

export async function loadKeySource(file: string, id: string): Promise<KeySource> {
  const bytes = new Uint8Array(fs.readFileSync(file));
  const name = path.basename(file);
  const book = await readBook(bytes, name, { id });
  const sheet = pickDataSheet(book);
  return { id, book, region: detectRegion(book, sheet), label: name.replace(/\.[^.]+$/, '') };
}

export interface Pipeline {
  base: KeySource;
  src: KeySource;
  analysis: KeyAnalysis;
  join: JoinResult;
  imports: ImportColumnSpec[];
  /** JoinResult.keyPairs when the keys module provides it, else derived from the enabled tiers (same meaning) */
  keyPairs: Record<string, Array<{ baseCol: number; sourceCol: number }>>;
}

/** The app's pipeline for ONE other file: analysis -> default plan -> join. `imports`: 'default' = the app's default (key columns excluded), 'all' = every column. */
export function buildPipeline(base: KeySource, src: KeySource, opts: { imports: 'default' | 'all' | number[]; joinType: 'left' | 'full' }): Pipeline {
  const analysis = analyzeSources(base, [src]);
  const plan = { ...defaultPlan(analysis), joinType: opts.joinType };
  const join = joinSources(analysis, plan);
  let imports: ImportColumnSpec[] = join.defaultImports.filter((d) => d.sourceId === src.id).map((d) => (d.header ? { sourceId: d.sourceId, col: d.col, header: d.header } : { sourceId: d.sourceId, col: d.col }));
  // an explicit list of source columns (the property test: some columns of the source are NOT imported)
  if (Array.isArray(opts.imports)) imports = opts.imports.map((col) => ({ sourceId: src.id, col }));
  else if (opts.imports === 'all') {
    const r = src.region;
    imports = [];
    for (let c = r.firstCol; c <= r.lastCol; c++) imports.push({ sourceId: src.id, col: c });
  }
  let keyPairs = join.keyPairs ?? {};
  if (!join.keyPairs) {
    keyPairs = {};
    const pair = analysis.pairs[src.id];
    const list: Array<{ baseCol: number; sourceCol: number }> = [];
    for (const cid of plan.tiers[src.id] ?? []) {
      const c = pair?.candidates.find((x) => x.id === cid);
      if (!c) continue;
      c.baseColumns.forEach((b, i) => { const o = c.otherColumns[i]; if (o && c.baseColumns.length === c.otherColumns.length) list.push({ baseCol: b.col, sourceCol: o.col }); });
    }
    keyPairs[src.id] = list;
  }
  return { base, src, analysis, join, imports, keyPairs };
}

export function toPremergeInput(p: Pipeline, options: Partial<PremergeInput['options']> = {}, withKeyPairs = true): PremergeInput {
  return {
    base: p.base,
    sources: [p.src],
    rows: p.join.rows,
    imports: p.imports,
    keyFill: p.join.keyFill,
    ...(withKeyPairs ? { keyPairs: p.keyPairs } : {}),
    options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix', dropOtherSheets: false, ...options },
  };
}

export function sheetOf(book: SourceBook, name: string): SheetGrid | undefined {
  return book.sheets.find((s) => s.name === name);
}

export function cellAt(sheet: SheetGrid | undefined, row: number, col: number): Cell | undefined {
  return sheet?.rows[row - 1]?.[col - 1];
}

/** one imported formula cell of a source (matched or appended row) and where it went */
export interface FormulaCellRec {
  srcRow: number;
  srcCol: number;
  row: number;
  col: number;
  f: string;
  arr: boolean;
  cached: boolean;
}

/** merged row of every placed source row (matched or appended) of source `sid`: [srcRow, mergedRow][] in plan order */
export function placedRows(join: JoinResult, sid: string, lastDataRow: number): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  let appendAt = lastDataRow + 1;
  for (const r of join.rows) {
    let target: number;
    if (r.baseRow == null) target = appendAt++;
    else target = r.baseRow;
    const m = r.matches?.[sid];
    const sr = m ? m.row : r.appendedFrom && r.appendedFrom.sourceId === sid ? r.appendedFrom.row : null;
    if (sr != null) out.push([sr, target]);
  }
  return out;
}

export function listFormulaCells(p: Pipeline, insertAt: number): FormulaCellRec[] {
  const sheet = sheetOf(p.src.book, p.src.region.sheet);
  const out: FormulaCellRec[] = [];
  const pr = placedRows(p.join, p.src.id, p.base.region.lastDataRow);
  p.imports.forEach((imp, j) => {
    for (const [sr, target] of pr) {
      const c = cellAt(sheet, sr, imp.col);
      if (c?.f) out.push({ srcRow: sr, srcCol: imp.col, row: target, col: insertAt + j, f: c.f, arr: !!c.arr || !!c.dyn, cached: c.t !== 'z' && c.v != null });
    }
  });
  return out;
}

export type Cat = 'live' | 'value' | 'EMPTY';

export function classify(merged: SheetGrid | undefined, rec: FormulaCellRec): { cat: Cat; f?: string } {
  const c = cellAt(merged, rec.row, rec.col);
  if (c?.f) return { cat: 'live', f: c.f };
  if (!c || c.t === 'z' || c.v == null) return { cat: 'EMPTY' };
  return { cat: 'value' };
}

// ---------------------------------------------------------------------------------------------------------------------
// counters of one premerge run (the expected.json snapshot of tests/fixtures/formulaimport, the findings tables)
// ---------------------------------------------------------------------------------------------------------------------
export interface Counts {
  /** formula cells of the imported columns (matched / appended rows) */
  total: number;
  kept: number;
  value: number;
  empty: number;
  /** reason -> number (values and empty cells together) */
  reasons: Record<string, number>;
  /** kept: a whole-table range now reads the merged table */
  widened: number;
  /** kept without a stored result (Excel recalculates it on load) */
  uncached: number;
  /** kept: a key column was replaced by the base's partner column */
  substituted: number;
  /** kept: reads the copies of the source's sheets (includeSourceSheets; absent when none) */
  viaCopy?: number;
}

export function summarize(trace: FormulaTrace[]): Counts {
  const c: Counts = { total: trace.length, kept: 0, value: 0, empty: 0, reasons: {}, widened: 0, uncached: 0, substituted: 0 };
  for (const t of trace) {
    if (t.outcome === 'kept') {
      c.kept++;
      if (t.widened) c.widened++;
      if (t.cached === false) c.uncached++;
      if (t.substituted) c.substituted++;
      if (t.viaCopy) c.viaCopy = (c.viaCopy ?? 0) + 1;
    } else {
      c[t.outcome]++;
      const why = t.reason ?? 'unknown';
      c.reasons[why] = (c.reasons[why] ?? 0) + 1;
    }
  }
  c.reasons = Object.fromEntries(Object.entries(c.reasons).sort(([a], [b]) => (a < b ? -1 : 1)));
  return c;
}

export interface RunCfg { results: 'cached' | 'nocache'; imports: 'default' | 'all'; join: 'left' | 'full' }

/** the app's pipeline (keys module -> buildPremerged) for one corpus file; `fixtures` = a directory with base.xlsx, cached/, nocache/ */
export async function runCorpusFile(fixtures: string, file: string, cfg: RunCfg, baseFile = 'base.xlsx', options: Partial<PremergeInput['options']> = {}): Promise<{ counts: Counts; trace: FormulaTrace[]; result: PremergeResult; pipeline: Pipeline }> {
  const base = await loadKeySource(path.join(fixtures, baseFile), 'b');
  const src = await loadKeySource(path.join(fixtures, cfg.results, file), 's1');
  const pipeline = buildPipeline(base, src, { imports: cfg.imports, joinType: cfg.join });
  const trace: FormulaTrace[] = [];
  const result = await buildPremerged(toPremergeInput(pipeline, options), undefined, (t) => trace.push(t));
  return { counts: summarize(trace), trace, result, pipeline };
}
