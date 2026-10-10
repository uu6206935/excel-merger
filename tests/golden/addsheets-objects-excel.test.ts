/**
 * Real Excel on the outputs whose copied sheets own objects (EXCEL_TESTS=1; needs Excel, Windows PowerShell and a few minutes) — task A2:
 *   1. tests/golden/tools/make-addsheets-objects-outputs.ts writes the premerged workbooks (every Excel-authored fixture of
 *      tests/fixtures/addsheets-objects alone, all together, twice (name / id collisions), as the base, a base with another theme, Strict
 *      packages, after edits replayed by the built-in engine and by the app's Excel engine, the golden fixtures and the foreign corpus files that
 *      own objects) and a manifest;
 *   2. tests/golden/com/check-addsheets-objects.ps1 opens every output in ONE hidden Excel of our own: no repair, the sheet names as planned,
 *      every copied sheet owns what its source sheet owns (shapes, charts and their series formulas / values / colours, notes, tables, pivot
 *      tables incl. PivotCache.Refresh on the copy), no new #REF! / #NAME?.
 * The results of the last run are in tests/golden/RESULTS.md ("Round 9, objects").
 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '../..');

describe.skipIf(process.env.EXCEL_TESTS !== '1')('real Excel opens every output with copied objects and finds the copies equal to their sources', () => {
  it('>= 35 combinations', () => {
    const dir = fs.mkdtempSync(path.join(process.env.ADDSHEETS_OUT ?? os.tmpdir(), 'addsheets-objects-'));
    try {
      const gen = spawnSync('npx', ['tsx', 'tests/golden/tools/make-addsheets-objects-outputs.ts', dir], { cwd: root, encoding: 'utf8', shell: true, maxBuffer: 1 << 26, env: { ...process.env, EXCEL_ENGINE: '1' } });
      expect(gen.status, gen.stdout + gen.stderr).toBe(0);
      expect(gen.stdout).not.toContain('FAILED');
      const manifest = JSON.parse(fs.readFileSync(path.join(dir, 'manifest.json'), 'utf8')) as unknown[];
      expect(manifest.length).toBeGreaterThanOrEqual(35);
      const chk = spawnSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(here, 'com/check-addsheets-objects.ps1'), '-Dir', dir], { cwd: root, encoding: 'utf8', maxBuffer: 1 << 26, timeout: 25 * 60_000 });
      const lines = chk.stdout.split(/\r?\n/).filter((l) => l.startsWith('{')).map((l) => JSON.parse(l) as Record<string, unknown>);
      const rows = lines.filter((l) => l.type === undefined);
      expect(rows.length).toBe(manifest.length);
      const notOk = rows.filter((r) => r.ok !== true).map((r) => `${r.name}: open=${r.open} names=${r.names} lint=${r.lint} shapes=${r.shapeDiffs} charts=${r.chartDiffs} notes=${r.noteDiffs} tables=${r.tableDiffs} pivots=${r.pivotDiffs} newErr=${r.newErr} ${JSON.stringify(r.info)}`);
      expect(notOk).toEqual([]);
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  }, 30 * 60_000);
});
