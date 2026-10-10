// check-output + compare for the outputs of the UI flows (tests/ui/real.mjs, real-excluded.mjs) with ONE Excel start
// (every Excel start costs time and ~80 MB of Office diagnostics in %TEMP%):
//   node tests/ui/verify-outputs.mjs --case "<base.xlsx>|<sheet or ->|<header row or ->|<excel-output.xlsx>|<js-output.xlsx>[|<sheets>[|<source;source;...>]]" [--case ...]
// <sheets> = what the output holds, as real.mjs printed on its `SHEETS` line: single (the default, the app's default: only the merged sheet →
// check-output --single-sheet) | base (the base's other sheets are kept: no flag) | all (+ the sheets of the other files: --with-copies, compared with the
// <source> files, given in card order and ';'-separated). For every case: tools/lib/verdict.mjs (the same verdict as tools/check-output.mjs) for BOTH outputs, then the same
// comparison as tools/compare.mjs (Excel view + package view) of the Excel本体 output against the built-in engine's.
// Exit 0 = every verdict PASS and no formula differences between the engines, 1 = otherwise, 2 = could not run.
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const lib = (f) => import(pathToFileURL(path.join(root, 'tools', 'lib', f)).href);
const { inspectFiles, excelAvailable } = await lib('inspect.mjs');
const { checkOutput, formatVerdict } = await lib('verdict.mjs');
const { compareInspections, diffSummary, formatComparison } = await lib('compare-core.mjs');
const { comparePackages } = await lib('package-diff.mjs');

const args = process.argv.slice(2);
const cases = [];
for (let i = 0; i < args.length; i++) {
  if (args[i] !== '--case') continue;
  const [base, sheet, headerRow, excel, js, mode = 'single', sources = ''] = String(args[++i]).split('|');
  cases.push({
    base: path.resolve(base),
    sheet: sheet && sheet !== '-' ? sheet : undefined,
    headerRow: headerRow && headerRow !== '-' ? Number(headerRow) : undefined,
    excel: path.resolve(excel),
    js: path.resolve(js),
    mode,
    sources: sources ? sources.split(';').filter(Boolean).map((s) => path.resolve(s)) : [],
  });
}
if (!cases.length) {
  process.stderr.write('usage: node tests/ui/verify-outputs.mjs --case "<base>|<sheet or ->|<headerRow or ->|<excel out>|<js out>" ...\n');
  process.exit(2);
}
if (!excelAvailable()) {
  process.stderr.write('Excel is not available\n');
  process.exit(2);
}

const files = [...new Set(cases.flatMap((c) => [c.base, c.excel, c.js, ...c.sources]))];
const t0 = Date.now();
const dumps = await inspectFiles(files, { recalc: true });
console.log(`inspected ${files.length} files in ONE Excel instance in ${Date.now() - t0} ms`);
const dumpOf = (f) => dumps[files.indexOf(f)];
let bad = 0;
for (const f of files) if (dumpOf(f).error && dumpOf(f).openMode === 'failed') {
  console.log(`FAIL: cannot inspect ${path.basename(f)}: ${dumpOf(f).error}`);
  bad++;
}

for (const c of cases) {
  console.log(`\n==== ${path.basename(c.base)} -> ${path.basename(c.excel)} / ${path.basename(c.js)}  (sheets: ${c.mode})`);
  for (const [label, file] of [['Excel本体', c.excel], ['内蔵エンジン', c.js]]) {
    const v = checkOutput(dumpOf(c.base), dumpOf(file), {
      sheet: c.sheet,
      headerRow: c.headerRow,
      singleSheet: c.mode === 'single',
      withCopies: c.mode === 'all',
      sources: c.sources.map((s) => ({ name: path.basename(s), dump: dumpOf(s) })),
    });
    console.log(`-- check-output (${label}): ${v.verdict} (open mode: ${dumpOf(file).openMode})`);
    console.log(formatVerdict(v, path.basename(file)));
    if (!v.ok || v.verdict !== 'PASS' || dumpOf(file).openMode !== 'normal') bad++;
  }
  const pkg = await comparePackages(c.excel, c.js, {});
  const pkgRows = pkg.rows;
  const xlRows = compareInspections(dumpOf(c.excel), dumpOf(c.js), {});
  const all = [...xlRows, ...pkgRows];
  const formulaDiffs = all.filter((r) => !r.ok && r.group === 'formulas').length;
  const total = all.filter((r) => !r.ok).length;
  console.log(`-- compare Excel本体 vs 内蔵エンジン: ${total ? `${total} differences` : 'identical'} — formula differences: ${formulaDiffs}`);
  console.log(`   Excel view: ${xlRows.filter((r) => !r.ok).length ? `${xlRows.filter((r) => !r.ok).length} differences (${diffSummary(xlRows)})` : 'identical'}`);
  const tx = formatComparison(xlRows, { onlyDiffs: true });
  if (tx) console.log(tx);
  console.log(`   Package view: ${pkgRows.filter((r) => !r.ok).length ? `${pkgRows.filter((r) => !r.ok).length} differences (${diffSummary(pkgRows)})` : 'identical'}`);
  const tp = formatComparison(pkgRows, { onlyDiffs: true });
  if (tp) console.log(tp);
  // like the E2E specs (expectSameFormulas): the engines must agree on every formula; other differences are listed
  if (formulaDiffs > 0) bad++;
}
console.log(bad ? `\nRESULT: ${bad} problem(s)` : '\nRESULT: all verdicts PASS, engines identical');
process.exit(bad ? 1 : 0);
