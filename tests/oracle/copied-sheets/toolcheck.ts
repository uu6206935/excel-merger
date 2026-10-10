/**
 * The QA tools on the oracle's own outputs (task V9): tools/lib/verdict.mjs --single-sheet / --with-copies (the logic of tools/check-output.mjs) and
 * tools/lib/package-diff.mjs / compare-core.mjs (tools/compare.mjs) are run over a few of the scenarios' outputs — both engines — in ONE Excel instance:
 *   - check-output must not FAIL for a reason that has to do with the output shapes (names of the copies, a copy that differs from its source, a sheet
 *     that should not be there, a formula that reads a removed sheet, no repair). Edits of the merged sheet can legitimately change what the
 *     other, header-based checks of the verdict see (a deleted column ...): those are reported, not judged here;
 *   - compare (Excel view + package view) of the two engines' outputs: formula differences must be 0.
 */
import path from 'node:path';
import { compareInspections, diffSummary } from '../../../tools/lib/compare-core.mjs';
import { inspectFiles, type InspectResult } from '../../../tools/lib/inspect.mjs';
import { comparePackages } from '../../../tools/lib/package-diff.mjs';
import { isStrictOoxml } from '../../../tools/lib/strict.mjs';
import { checkOutput } from '../../../tools/lib/verdict.mjs';
import type { Finding } from '../consistency/check';
import type { CopyRec } from './judge';

/** the verdict codes that belong to the output shapes (any of them FAILing is a finding) */
const SHAPE_FAILS = /^(open-failed|repair|inspect-mismatch|not-single-sheet|dangling-sheet-ref|copy-missing|copy-unexpected|copy-order|copy-differs|copy-new-error|new-ref-text|new-error-#REF!|new-error-#NAME\?|sheets-lost|charts-lost|pivots-lost|table-lost)$/;

export interface ToolCheckStats {
  scenarios: number;
  verdicts: number;
  compared: number;
  summary: string[];
}

export async function toolChecks(picks: Array<{ rec: CopyRec; js: string; xl: string | null }>, log: (s: string) => void): Promise<{ findings: Finding[]; stats: ToolCheckStats }> {
  const findings: Finding[] = [];
  // one Excel instance per scenario (base + both outputs + its sources): a hidden Excel that dies takes only that scenario's files with it (inspectFiles retries
  // the failed ones in a fresh instance; a dump that still failed is read once more on its own)
  const cache = new Map<string, InspectResult>();
  const load = async (list: string[]): Promise<void> => {
    const missing = [...new Set(list)].filter((f) => !cache.has(f));
    if (!missing.length) return;
    const got = await inspectFiles(missing, { recalc: true });
    missing.forEach((f, i) => cache.set(f, got[i]));
    for (const f of missing.filter((f) => cache.get(f)!.error)) cache.set(f, (await inspectFiles([f], { recalc: true }))[0]);
  };
  const dumpOf = (f: string): InspectResult => cache.get(f)!;
  const stats: ToolCheckStats = { scenarios: picks.length, verdicts: 0, compared: 0, summary: [] };
  for (const { rec, js, xl } of picks) {
    log(`tools: inspecting ${rec.id}`);
    await load([rec.baseFile, js, ...(xl ? [xl] : []), ...rec.sources.map((s) => s.path)]);
    const base = dumpOf(rec.baseFile);
    const sources = await Promise.all(rec.sources.map(async (s) => ({ name: s.fileName, dump: dumpOf(s.path), strict: await isStrictOoxml(s.path) })));
    for (const [label, file] of [['js', js], ...(xl ? ([['xl', xl]] as const) : [])] as Array<[string, string]>) {
      const out = dumpOf(file);
      const v = checkOutput(base, out, { sheet: rec.mergedSheet, headerRow: rec.headerRow, singleSheet: rec.drop, withCopies: true, sources, outStrict: await isStrictOoxml(file) });
      stats.verdicts++;
      const bad = v.items.filter((i) => i.level === 'fail' && SHAPE_FAILS.test(i.code));
      const other = v.items.filter((i) => i.level === 'fail' && !SHAPE_FAILS.test(i.code)).map((i) => i.code);
      stats.summary.push(`${rec.id} ${label}: ${v.verdict}${other.length ? ` (other: ${[...new Set(other)].join(',')})` : ''}`);
      for (const i of bad) findings.push({ scenario: rec.id, cls: `tool-check-output-${label}`, detail: `${i.code}: ${i.message}${i.examples?.length ? ` — ${i.examples.slice(0, 3).join(' | ')}` : ''}`.slice(0, 600) });
      const copies = v.items.find((i) => i.code === 'copy-equal' || i.code === 'copy-differs');
      if (!copies) findings.push({ scenario: rec.id, cls: `tool-check-output-${label}`, detail: 'the verdict says nothing about the copied sheets (--with-copies did not run)' });
    }
    if (xl) {
      const rows = [...compareInspections(dumpOf(js), dumpOf(xl), {}), ...(await comparePackages(js, xl, {})).rows];
      const formulaDiffs = rows.filter((r) => !r.ok && !r.filtered && r.group === 'formulas');
      stats.compared++;
      stats.summary.push(`${rec.id} compare js vs xl: ${rows.filter((r) => !r.ok && !r.filtered).length} difference(s) (${diffSummary(rows) || 'none'})`);
      // (the other groups are reported with an example each: they are how the two engines write things, not what this oracle judges)
      for (const r of rows.filter((r) => !r.ok && !r.filtered && r.group !== 'formulas').slice(0, 4)) stats.summary.push(`    ${r.group} ${r.sheet}: ${String(r.detail).replace(/\s+/g, ' ').slice(0, 220)}`);
      for (const r of formulaDiffs.slice(0, 4)) findings.push({ scenario: rec.id, cls: 'tool-compare-formulas', detail: `${r.sheet}: ${String(r.detail).slice(0, 300)}` });
    }
    log(`tools: ${path.basename(js)} checked`);
  }
  return { findings, stats };
}
