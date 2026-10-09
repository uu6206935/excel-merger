/**
 * Findings tool (task F1, step 1): for every corpus pair (base.xlsx + one source file; cached = Excel-saved results, nocache =
 * openpyxl / XlsxWriter without results) run the app's pipeline (keys module -> buildPremerged) and classify EVERY formula cell of
 * the imported columns as live (still a formula) / value (cached value written) / EMPTY (nothing), with the reason.
 *
 *   npx tsx tests/oracle/import-values/diagnose.ts --dir <corpusDir> [--engine legacy|new|both] [--md out.md] [--json out.json] [--only regex]
 *
 * 'legacy' = tests/oracle/import-values/legacy/premerge-legacy.ts (frozen copy of the pre-round-8 premerge; its reasons are derived
 * here with legacyReason(), which replays the old rule token by token); 'new' = src/core/xlsx/premerge.ts (reasons from its trace hook).
 */
import fs from 'node:fs';
import path from 'node:path';
import type { PremergeInput, PremergeResult } from '../../../src/core/contracts';
import { isUnsafeFormula } from '../../../src/core/export/comJob';
import { POSITIONAL_FUNCTIONS } from '../../../src/core/formula/forms';
import { funcName } from '../../../src/core/formula/calls';
import { tokenize } from '../../../src/core/formula/excelRef';
import { readBook } from '../../../src/core/xlsx';
import { buildPremerged as buildNew } from '../../../src/core/xlsx/premerge';
import { buildPremerged as buildLegacy } from './legacy/premerge-legacy';
import { buildPipeline, classify, listFormulaCells, loadKeySource, loadManifest, sheetOf, toPremergeInput, type CorpusEntry, type FormulaCellRec, type Pipeline } from './lib';

const args = process.argv.slice(2);
const opt = (k: string, d = ''): string => { const i = args.indexOf(`--${k}`); return i >= 0 && args[i + 1] ? args[i + 1] : d; };
const dir = opt('dir');
const engines = (opt('engine', 'both') === 'both' ? ['legacy', 'new'] : [opt('engine')]) as Array<'legacy' | 'new'>;
const only = opt('only') ? new RegExp(opt('only')) : null;
/** --foreign: the repository's foreign-producer corpus (tests/fixtures/foreign/corpus: files of openpyxl, XlsxWriter, SheetJS, POI, LibreOffice … as the OTHER file of partner.xlsx) */
const foreign = args.includes('--foreign');

export interface Row {
  engine: string;
  item: string;
  order: string;
  results: 'cached' | 'nocache';
  imports: 'default' | 'all';
  join: 'left' | 'full';
  total: number;
  live: number;
  value: number;
  empty: number;
  /** reason -> {value, empty} */
  reasons: Record<string, { value: number; empty: number }>;
  /** examples of formulas that did not stay live: [source formula, reason, category] */
  examples: Array<[string, string, string]>;
  keptExamples: Array<[string, string]>;
  error?: string;
}

/** the reason the PRE-round-8 premerge had for not keeping a formula (first failing token of its rule) */
function legacyReason(rec: FormulaCellRec, p: Pipeline): string {
  if (rec.arr) return 'array';
  if (isUnsafeFormula(rec.f)) return 'unsafe';
  const r = p.src.region;
  const imported = new Set(p.imports.map((i) => i.col));
  for (const t of tokenize(rec.f)) {
    switch (t.type) {
      case 'unknown': return 'unknownToken';
      case 'name': if (!/^_xl(pm|op|eta)\./i.test(t.text)) return 'definedName'; break;
      case 'sref': return 'tableRef';
      case 'error': if (t.prefix) return 'unknownToken'; break;
      case 'func':
        if (funcName(t) === 'INDIRECT' || POSITIONAL_FUNCTIONS.has(funcName(t))) return 'positional';
        if (t.prefix) return 'unknownToken';
        break;
      case 'ref': {
        const pf = t.prefix;
        if (pf && (pf.book !== undefined || pf.sheet2 !== undefined)) return 'external/3D';
        if (pf && pf.sheet && pf.sheet.toLowerCase() !== r.sheet.toLowerCase()) return 'otherSheet';
        const a = t.area!;
        if (a.kind !== 'cell' && a.kind !== 'area') return 'wholeColumnRow';
        if (a.r1 !== rec.srcRow || a.r2 !== rec.srcRow) return a.r1 >= r.firstDataRow && a.r2 <= r.lastDataRow ? 'otherRow(range/other row)' : 'outsideTable';
        for (let c = Math.min(a.c1, a.c2); c <= Math.max(a.c1, a.c2); c++) if (!imported.has(c)) return 'notImportedColumn';
        break;
      }
    }
  }
  return 'consecutiveColumns';
}

async function runOne(engine: 'legacy' | 'new', e: CorpusEntry, results: 'cached' | 'nocache', imports: 'default' | 'all', join: 'left' | 'full'): Promise<Row> {
  const row: Row = { engine, item: e.name, order: e.order, results, imports, join, total: 0, live: 0, value: 0, empty: 0, reasons: {}, examples: [], keptExamples: [] };
  try {
    const base = await loadKeySource(path.join(dir, foreign ? 'partner.xlsx' : e.base ?? 'base.xlsx'), 'b');
    const src = await loadKeySource(path.join(dir, foreign ? '' : results === 'cached' ? 'cached' : 'nocache', e.file), 's1');
    const p = buildPipeline(base, src, { imports, joinType: join });
    const input: PremergeInput = toPremergeInput(p);
    const traceReason = new Map<string, string>();
    let res: PremergeResult;
    if (engine === 'legacy') res = await buildLegacy(input);
    else res = await (buildNew as unknown as (i: PremergeInput, pr?: undefined, tr?: (t: { row: number; col: number; reason?: string }) => void) => Promise<PremergeResult>)(input, undefined, (t) => { if (t.reason) traceReason.set(`${t.row}:${t.col}`, t.reason); });
    const out = await readBook(res.bytes, 'out.xlsx', { id: 'out' });
    const merged = sheetOf(out, res.layout.sheet);
    const recs = listFormulaCells(p, res.layout.insertAt);
    for (const rec of recs) {
      const { cat, f } = classify(merged, rec);
      row.total++;
      if (cat === 'live') { row.live++; if (row.keptExamples.length < 4) row.keptExamples.push([rec.f, f ?? '']); continue; }
      row[cat === 'value' ? 'value' : 'empty']++;
      const reason = engine === 'legacy' ? legacyReason(rec, p) : traceReason.get(`${rec.row}:${rec.col}`) ?? 'unknown';
      const slot = (row.reasons[reason] ??= { value: 0, empty: 0 });
      slot[cat === 'value' ? 'value' : 'empty']++;
      if (row.examples.length < 6) row.examples.push([rec.f, reason, cat]);
    }
  } catch (err) {
    row.error = String((err as Error)?.stack ?? err).split('\n').slice(0, 3).join(' | ');
  }
  return row;
}

async function main(): Promise<void> {
  const all: CorpusEntry[] = foreign
    ? fs.readdirSync(dir).filter((f) => /.(xlsx|xlsm)$/i.test(f) && f !== 'partner.xlsx' && (!only || only.test(f))).sort().map((f) => ({ name: f.replace(/.[^.]+$/, ''), order: 'same' as const, file: f, desc: '', sheet: '', headerRow: 1, firstDataRow: 2, lastDataRow: 2 }))
    : loadManifest(dir).filter((e) => !only || only.test(e.file));
  const rows: Row[] = [];
  for (const engine of engines) {
    for (const e of all) {
      for (const results of foreign ? (['cached'] as const) : (['cached', 'nocache'] as const)) {
        for (const imports of ['default', 'all'] as const) {
          rows.push(await runOne(engine, e, results, imports, 'left'));
        }
      }
    }
    // join 'full' (appended rows) on the shuffled order only
    if (!foreign) for (const e of all.filter((x) => x.order === 'shuf')) rows.push(await runOne(engine, e, 'cached', 'default', 'full'));
  }
  const jsonOut = opt('json');
  if (jsonOut) fs.writeFileSync(jsonOut, JSON.stringify(rows, null, 1));
  const md = report(rows);
  const mdOut = opt('md');
  if (mdOut) fs.writeFileSync(mdOut, md);
  console.log(md);
}

function agg(rows: Row[]): { total: number; live: number; value: number; empty: number } {
  return rows.reduce((a, r) => ({ total: a.total + r.total, live: a.live + r.live, value: a.value + r.value, empty: a.empty + r.empty }), { total: 0, live: 0, value: 0, empty: 0 });
}

function report(rows: Row[]): string {
  const out: string[] = [];
  const cell = (a: { total: number; live: number; value: number; empty: number }) => `${a.live} / ${a.value} / ${a.empty}`;
  for (const engine of engines) {
    out.push(`\n### ${engine}: live / value / EMPTY per item (all three orders, join left)\n`);
    out.push('| item | cached, default imports | cached, all columns | no results, default imports | no results, all columns |');
    out.push('|---|---|---|---|---|');
    const items = [...new Set(rows.filter((r) => r.engine === engine).map((r) => r.item))];
    for (const it of items) {
      const sel = (res: string, imp: string) => agg(rows.filter((r) => r.engine === engine && r.item === it && r.results === res && r.imports === imp && r.join === 'left'));
      out.push(`| ${it} | ${cell(sel('cached', 'default'))} | ${cell(sel('cached', 'all'))} | ${cell(sel('nocache', 'default'))} | ${cell(sel('nocache', 'all'))} |`);
    }
    const tot = (res: string, imp: string) => agg(rows.filter((r) => r.engine === engine && r.results === res && r.imports === imp && r.join === 'left'));
    out.push(`| **total** | ${cell(tot('cached', 'default'))} | ${cell(tot('cached', 'all'))} | ${cell(tot('nocache', 'default'))} | ${cell(tot('nocache', 'all'))} |`);
    const errs = rows.filter((r) => r.engine === engine && r.error);
    if (errs.length) out.push(`\nERRORS (${errs.length}): ${errs.slice(0, 5).map((r) => `${r.item}/${r.order}/${r.results}/${r.imports}: ${r.error}`).join('\n')}`);
    // reasons
    const reason = new Map<string, { value: number; empty: number }>();
    for (const r of rows.filter((x) => x.engine === engine && x.join === 'left')) for (const [k, v] of Object.entries(r.reasons)) {
      const s = reason.get(k) ?? { value: 0, empty: 0 };
      s.value += v.value; s.empty += v.empty;
      reason.set(k, s);
    }
    out.push(`\n### ${engine}: why formulas did not stay live (value = cached value written, EMPTY = nothing written)\n`);
    out.push('| reason | value | EMPTY |');
    out.push('|---|---|---|');
    for (const [k, v] of [...reason].sort((a, b) => b[1].value + b[1].empty - a[1].value - a[1].empty)) out.push(`| ${k} | ${v.value} | ${v.empty} |`);
  }
  return out.join('\n') + '\n';
}

main().catch((e) => { console.error(e); process.exit(1); });
