#!/usr/bin/env node
// Diff of two workbooks (owner: T) — made for "same op log, two engines": the JS-engine output vs the Excel-engine
// output (or a golden Excel result vs replayLog). Two views, each reporting only REAL differences:
//
//   Excel view   : both files opened (and recalculated) in our own hidden Excel (tools/inspect.ps1): formulas
//                  (Formula2), values, spills, merges, CF, DV, tables, charts, pivots, widths, style probes …
//   package view : both packages parsed and normalized (tools/lib/package-diff.mjs): formulas (shared formulas
//                  expanded), constants, RESOLVED cell styles, columns, rows, merges, pane, filter, CF, DV,
//                  hyperlinks, tables, drawings/charts, comments, pivots, names … — no Excel needed
//
// Known-benign differences are filtered and listed at the end with counts (see tools/README.md): calcChain,
// docProps, style xf indexes, sharedStrings order, shared-formula representation, cached formula values,
// spill extents, empty <c/> cells, <dimension>, selection/scroll, CF priority numbers, _xlnm._FilterDatabase,
// pivot location, drawing offsets, XML attribute order / namespaces / revision ids, values of volatile functions.
//
//   node tools/compare.mjs <a.xlsx|a.json> <b.xlsx|b.json> [--excel-only | --package-only] [--no-values]
//        [--no-recalc] [--pivot-layout] [--all] [--json]
//
// .json arguments are inspect dumps (tools/inspect.mjs --json) -> Excel view only. Without Excel the Excel view is
// skipped (said so in the output). --no-values: Excel view without cell values; --no-recalc: do not recalculate
// before reading values; --all: also print the rows that are equal.
// Exit code 0 = no real differences, 1 = differences, 2 = could not run.
import fs from 'node:fs';
import path from 'node:path';
import { excelAvailable, inspectFiles } from './lib/inspect.mjs';
import { compareInspections, diffSummary, formatComparison } from './lib/compare-core.mjs';
import { comparePackages } from './lib/package-diff.mjs';

const args = process.argv.slice(2);
const files = args.filter((a) => !a.startsWith('--'));
const flag = (f) => args.includes(f);
const known = ['--all', '--no-values', '--recalc', '--no-recalc', '--pivot-layout', '--excel-only', '--package-only', '--json'];
const unknown = args.filter((a) => a.startsWith('--') && !known.includes(a));
if (files.length !== 2 || unknown.length || (flag('--excel-only') && flag('--package-only'))) {
  process.stderr.write(
    `${unknown.length ? `unknown option ${unknown.join(' ')}\n` : ''}usage: node tools/compare.mjs <a.xlsx|a.json> <b.xlsx|b.json> [--excel-only|--package-only] [--no-values] [--no-recalc] [--pivot-layout] [--all] [--json]\n`,
  );
  process.exit(2);
}
for (const f of files) {
  if (!fs.existsSync(f)) {
    process.stderr.write(`ファイルがありません: ${f}\n`);
    process.exit(2);
  }
}

const isJson = (f) => f.toLowerCase().endsWith('.json');
const bothPackages = files.every((f) => /\.(xlsx|xlsm)$/i.test(f));
const opts = { ignoreValues: flag('--no-values'), pivotLayout: flag('--pivot-layout') };
const report = { a: path.resolve(files[0]), b: path.resolve(files[1]), excel: null, package: null, notes: [] };

try {
  // ---- package view
  if (!flag('--excel-only')) {
    if (bothPackages) {
      const r = await comparePackages(path.resolve(files[0]), path.resolve(files[1]), opts);
      report.package = { rows: r.rows, filtered: r.filtered, diffs: r.rows.filter((x) => !x.ok).length };
    } else report.notes.push('package view skipped: needs two .xlsx / .xlsm files');
  }
  // ---- Excel view
  if (!flag('--package-only')) {
    const toInspect = files.filter((f) => !isJson(f));
    if (toInspect.length && !excelAvailable()) report.notes.push('Excel view skipped: Excel is not installed');
    else {
      const recalc = !flag('--no-recalc');
      const inspected = toInspect.length ? await inspectFiles(toInspect.map((f) => path.resolve(f)), { recalc }) : [];
      const dumps = files.map((f) => (isJson(f) ? JSON.parse(fs.readFileSync(f, 'utf8')) : inspected[toInspect.indexOf(f)]));
      const failed = dumps.filter((d) => d.error);
      if (failed.length) report.notes.push(`Excel view: inspection failed: ${failed.map((d) => `${d.file}: ${d.error}`).join(' / ')}`);
      const rows = compareInspections(dumps[0], dumps[1], opts);
      report.excel = { rows, diffs: rows.filter((x) => !x.ok).length, recalculated: recalc };
    }
  }
} catch (e) {
  process.stderr.write(`比較を実行できませんでした: ${e instanceof Error ? e.stack ?? e.message : String(e)}\n`);
  process.exit(2);
}

const total = (report.excel?.diffs ?? 0) + (report.package?.diffs ?? 0);
report.identical = total === 0;
report.formulaDiffs = [...(report.excel?.rows ?? []), ...(report.package?.rows ?? [])].filter((r) => !r.ok && r.group === 'formulas').length;

if (flag('--json')) {
  process.stdout.write(JSON.stringify(report, null, 1) + '\n');
} else {
  const out = [];
  const head = (n, rows) => (n ? `${n} differences (${diffSummary(rows)})` : 'identical');
  out.push(
    `${path.basename(files[0])} vs ${path.basename(files[1])}: ${total ? `${total} differences` : 'identical'} — formula differences: ${report.formulaDiffs}`,
  );
  for (const n of report.notes) out.push(`(${n})`);
  if (report.excel) {
    out.push('', `## Excel view (${report.excel.recalculated ? 'recalculated' : 'as saved'}${opts.ignoreValues ? ', values not compared' : ''}): ${head(report.excel.diffs, report.excel.rows)}`);
    const t = formatComparison(report.excel.rows, { onlyDiffs: !flag('--all') });
    if (t) out.push(t);
  }
  if (report.package) {
    out.push('', `## Package view (normalized XML): ${head(report.package.diffs, report.package.rows)}`);
    const t = formatComparison(report.package.rows, { onlyDiffs: !flag('--all') });
    if (t) out.push(t);
  }
  process.stdout.write(out.join('\n') + '\n');
}
process.exit(total ? 1 : 0);
