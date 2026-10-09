/**
 * The consistency oracle (oracle.ts) under vitest.
 *  - always: the Node half runs without Excel — scenarios are built through the app's pipeline (readBook -> keys -> joinSources ->
 *    buildPremerged -> createEditor), random op logs are made through the doc API, replayLog exports every log, and the doc's own view of
 *    the result must equal the package read back (tests/doc/compare.ts) for every scenario, and every reference to one cell / one column
 *    must still point at the same column id;
 *  - EXCEL_TESTS=1: the whole oracle with real Excel (run.ps1: one hidden Excel per batch of workbooks for the recalculation, one Excel
 *    start per scenario for the app's Excel engine): no value changed that nothing explains, no new error, both engines equal, every
 *    package opens without repair, imported formulas equal their source for the same person.
 *    EXCEL_ORACLE_COUNT (default 40) sets the number of scenarios; the big run (>= 300) is run.ps1 -Count 340.
 */
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { INFO_CLASSES, KNOWN_OPEN } from './check';
import { runOracle } from './oracle';
import { HERE } from './lib';

const EXCEL = process.env.EXCEL_TESTS === '1';

describe('consistency oracle: Node half (no Excel)', () => {
  it('builds scenarios over the sample / rich / row-rearrangement sets and the doc equals replayLog for every random log', async () => {
    const out = fs.mkdtempSync(path.join(os.tmpdir(), 'cons-oracle-'));
    try {
      // one set of each family: the samples, the rich base with a formula source of F1's corpus, a source whose rows are shuffled
      const only = '^(smp-demo|smp-02base|rich-f04_lookups-c|rich-table|row-shuf-c|row-mix-n)$';
      const r = await runOracle(['--out', out, '--count', '36', '--seed-base', '900', '--no-excel', '--only', only]);
      expect(r.recs.length).toBeGreaterThanOrEqual(20);
      expect(r.recs.filter((x) => x.docVsReplay.length).map((x) => `${x.id}: ${x.docVsReplay.join(' | ')}`)).toEqual([]);
      expect(r.recs.some((x) => x.opsKinds.move)).toBe(true);
      expect(r.recs.some((x) => x.opsKinds.insert)).toBe(true);
      expect(r.recs.some((x) => x.opsKinds.edit)).toBe(true);
      // every kind of set took part (samples, rich base with F1's sources, rows rearranged)
      for (const prefix of ['smp-', 'rich-', 'row-']) expect(r.recs.some((x) => x.set.startsWith(prefix)), prefix).toBe(true);
    } finally {
      fs.rmSync(out, { recursive: true, force: true });
    }
  }, 300_000);
});

describe.skipIf(!EXCEL)('consistency oracle: real Excel (EXCEL_TESTS=1)', () => {
  it('formulas stay consistent through moves / inserts / deletes in both engines, rows rearranged by key keep their values', async () => {
    const out = fs.mkdtempSync(path.join(os.tmpdir(), 'cons-oracle-xl-'));
    try {
      const ps1 = path.join(HERE, 'run.ps1');
      const count = process.env.EXCEL_ORACLE_COUNT ?? '40';
      // asynchronous: a blocking call would starve vitest's worker heartbeat for the minutes Excel needs
      const code = await new Promise<number | null>((resolve, reject) => {
        const child = spawn('powershell.exe', ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', ps1, '-Count', count, '-Out', out], { stdio: ['ignore', 'inherit', 'inherit'] });
        const timer = setTimeout(() => child.kill(), 90 * 60_000);
        child.on('error', reject);
        child.on('close', (c) => { clearTimeout(timer); resolve(c); });
      });
      const findings = JSON.parse(fs.readFileSync(path.join(out, 'findings.json'), 'utf8')) as Array<{ scenario: string; cls: string; detail: string }>;
      const failing = findings.filter((f) => !INFO_CLASSES.test(f.cls) && !KNOWN_OPEN.test(f.cls));
      expect(failing.map((f) => `${f.scenario} ${f.cls}: ${f.detail}`.slice(0, 300))).toEqual([]);
      expect(code).toBe(0);
    } finally {
      fs.rmSync(out, { recursive: true, force: true });
    }
  }, 100 * 60_000);
});
