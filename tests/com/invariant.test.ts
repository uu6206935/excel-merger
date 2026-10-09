/**
 * The JS-vs-Excel invariant (docs/ARCHITECTURE.md §6) against the REAL Excel (skipped unless EXCEL_TESTS=1):
 * the SAME op log goes through the Excel runner (excel-runner.ps1, the user's Excel) and through replayLog (the JS
 * package patcher), and the two workbooks must be the same workbook:
 *   - what Excel shows for every cell of every sheet: formula (Formula / Formula2), value, number format, alignment,
 *     array / spill, merged areas, stored widths, tables, defined names                        (lib.diffDumps)
 *   - what the packages contain: CF / DV (sqref + formulas), merges, tables, autofilter, names, widths (lib.packageDiff)
 *   - tools/compare.mjs: the Excel view (recalculated; formulas, values, CF, DV, hidden columns, freeze, filter,
 *     style probes ...) and the normalized package view (resolved styles, columns, rows ...)
 * Every scenario (tests/com/scenarios.ts, fixtures authored by Excel itself) is a log with unmerge, numFmt / dyn
 * writes, toValues over arrays, table moves, inserts (columns and rows), deletes, moves, widths.
 *
 * Differences that are understood and accepted are listed in KNOWN with the engine that is right and the evidence;
 * everything else fails the test and is printed.
 */
import { afterAll, describe, expect, it } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { logToRunnerJob } from '../../src/core/export/comJob';
import { runRunner, type RunnerEvent } from '../../server/excel/run-job.mjs';
import { compareInspections, diffSummary, formatComparison } from '../../tools/lib/compare-core.mjs';
import { excelAvailable, inspectFiles } from '../../tools/lib/inspect.mjs';
import { comparePackages } from '../../tools/lib/package-diff.mjs';
import { EXCEL_TESTS, diffDumps, dumpBook, makeFixture, packageDiff, tempDir } from './lib';
import { SCENARIOS } from './scenarios';

const RUNNER = path.resolve(__dirname, '..', '..', 'server', 'excel', 'excel-runner.ps1');

/**
 * Accepted differences: [scenario or '*', regex over the difference text, who is right + evidence].
 *
 * History of differences this test found (fixed ones are listed so the reasoning stays visible):
 *  - r2 numFmt writes: the runner set Range.NumberFormat BEFORE the value, so Excel did not widen the standard-width
 *    column for a date (stored width 8.38 -> "########" on screen) while replayLog widened it to 9.63 like Excel does
 *    when the format is applied to a cell that already holds the number (tests/golden/RESULTS.md, probe-numfmt-width).
 *    replayLog (= typing in Excel) was right; FIXED: the runner applies every format except "@" AFTER the value.
 *  - r2 '001' with numFmt '@' inside a column of mixed formats: the runner wrote "'001" (quotePrefix) although the cell
 *    itself has the text format (Excel does not set quotePrefix for typing into a text cell). FIXED in Cells.Build.
 *  - r2 "=text" (N7): the runner guessed quotePrefix from the text (Risky), replayLog had no flag: "quotePrefix: あり -> なし".
 *    FIXED in round 3 by the explicit CellWrite.quote flag (both engines: quote === true -> quotePrefix, anything else -> none; the
 *    runner no longer guesses). The accepted-difference entry is gone: the two engines must agree.
 *  - cfdv: Excel-authored fixture with an AutoFilter whose hidden name _xlnm._FilterDatabase (A1:J21) disagrees with
 *    <autoFilter ref="A1:J22"> (Excel adopted the whole current region): fixture artifact, removed by asking for A1:J22.
 */
const KNOWN: Array<[string, RegExp, string]> = [];

/** tools/compare rows of CF / DV are order / fragmentation differences only when the coverage-normalized comparison is clean. */
const RULE_GROUPS = new Set(['cf', 'dv']);

describe.skipIf(!EXCEL_TESTS)('JS-vs-Excel invariant: the same log through the runner and through replayLog', () => {
  const dirs: string[] = [];
  afterAll(() => {
    for (const d of dirs) {
      if (process.env.KEEP_COM_OUTPUT !== '1') fs.rmSync(d, { recursive: true, force: true });
      else console.log(`[tests/com] outputs kept in ${d}`);
    }
  });

  for (const sc of SCENARIOS) {
    it(
      `${sc.name}: ${sc.covers}`,
      async () => {
        const dir = tempDir(`inv-${sc.name}`);
        dirs.push(dir);
        const fixture = path.join(dir, 'fixture.xlsx');
        await makeFixture(sc.spec(), fixture);
        const before = await dumpBook(fixture);
        const { log } = sc.build(before);

        // --- Excel runner
        const job = logToRunnerJob(log, { format: 'xlsx', sheet: sc.main });
        const excelOut = path.join(dir, 'excel.xlsx');
        const jobFile = path.join(dir, 'job.json');
        fs.writeFileSync(jobFile, JSON.stringify({ ...job, input: fixture, output: excelOut, workDir: path.join(dir, 'w') }));
        const events: RunnerEvent[] = [];
        const r = await runRunner({ runner: RUNNER, jobFile, env: { XLRUNNER_CACHE: path.join(dir, 'cache') }, onEvent: (e) => events.push(e) });
        expect(r.ok, JSON.stringify(r)).toBe(true);
        const warnings = events.filter((e) => e.type === 'warn').map((e) => `${e.code}: ${e.message}`);

        // --- JS engine
        const { replayLog } = await import('../../src/core/xlsx/replay');
        const js = await replayLog(new Uint8Array(fs.readFileSync(fixture)), log);
        const jsOut = path.join(dir, 'js.xlsx');
        fs.writeFileSync(jsOut, js);
        const excelBytes = new Uint8Array(fs.readFileSync(excelOut));

        const diffs: string[] = [];
        // --- Excel's own per-cell view
        const [xd, jd] = [await dumpBook(excelOut), await dumpBook(jsOut)];
        diffs.push(...diffDumps(xd, jd).map((d) => `[cells] ${d}`));
        diffs.push(...(await packageDiff(excelBytes, js)).map((d) => `[package] ${d}`));
        // --- tools/compare.mjs
        expect(excelAvailable()).toBe(true);
        const [ia, ib] = await inspectFiles([excelOut, jsOut], { recalc: true });
        expect(ia.openMode, 'Excel opens the runner output normally').toBe('normal');
        expect(ib.openMode, 'Excel opens the replayLog output without repair').toBe('normal');
        const excelView = compareInspections(ia, ib, {});
        const pkgView = await comparePackages(excelBytes, js, {});
        const bad = [...excelView.filter((x) => !x.ok).map((x) => ({ ...x, view: 'excel' })), ...pkgView.rows.filter((x) => !x.ok).map((x) => ({ ...x, view: 'package' }))];
        // CF / DV: Excel re-creates the rules of moved cells (other priority order, ranges split in pieces: "C14:D22 C2:D13");
        // the rules themselves (type, formulas, covered cells) are compared by packageDiff above. The tool's text comparison
        // of the same rules only adds order / fragmentation noise, so its cf / dv rows count only when packageDiff found a rule difference.
        const ruleDiffs = diffs.some((d) => d.startsWith('[package]') && / rule /.test(d));
        const ruleNoise = bad.filter((x) => RULE_GROUPS.has(x.group) && !ruleDiffs);
        diffs.push(...bad.filter((x) => !ruleNoise.includes(x)).map((x) => `[compare ${x.view}/${x.group}] ${x.sheet}: ${x.detail}`));

        const accepted = diffs.filter((d) => KNOWN.some(([s, re]) => (s === '*' || s === sc.name) && re.test(d)));
        const real = diffs.filter((d) => !accepted.includes(d));
        console.log(
          `[tests/com] invariant ${sc.name}: ${log.length} log entries -> ${job.ops.length} runner ops; runner warnings ${JSON.stringify(warnings)}; ` +
            `${real.length} differences (${accepted.length} accepted, ${ruleNoise.length} CF/DV order or fragmentation rows ignored)` +
            `; compare: excel view ${diffSummary(excelView) || 'identical'}, package view ${diffSummary(pkgView.rows) || 'identical'}`,
        );
        if (real.length) {
          console.warn(`[tests/com] invariant ${sc.name}: Excel runner (A) vs replayLog (B)\n${real.join('\n')}\n${formatComparison(excelView, { onlyDiffs: true })}`);
        }
        expect(real).toEqual([]);
      },
      900_000,
    );
  }
});
