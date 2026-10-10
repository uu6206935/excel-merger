/**
 * Real Excel on the outputs with copied sheets (EXCEL_TESTS=1; needs Excel, Windows PowerShell and a few minutes):
 *   1. tests/golden/tools/make-addsheets-outputs.ts writes >= 40 premerged workbooks WITH copied sheets (samples in several card orders, the
 *      Excel-authored fixtures, the foreign corpus, same-name files, edits replayed by the built-in engine and, with EXCEL_ENGINE=1, by the
 *      app's Excel engine) and a manifest;
 *   2. tests/golden/com/check-addsheets.ps1 opens every output in ONE hidden Excel of our own: no repair, the sheet names exactly as planned,
 *      every copied sheet equal to its source (values after a full recalculation, formulas modulo the renamed sheet prefixes, sampled
 *      formats, widths, tab colour, conditional-format / validation / hyperlink counts), no new #REF! / #NAME?.
 * The results of the last run are in tests/golden/RESULTS.md.
 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '../..');

describe.skipIf(process.env.EXCEL_TESTS !== '1')('real Excel opens every output with copied sheets and finds the copies equal to their sources', () => {
  it('>= 40 combinations', () => {
    const dir = fs.mkdtempSync(path.join(process.env.ADDSHEETS_OUT ?? os.tmpdir(), 'addsheets-'));
    try {
      const gen = spawnSync('npx', ['tsx', 'tests/golden/tools/make-addsheets-outputs.ts', dir], { cwd: root, encoding: 'utf8', shell: true, maxBuffer: 1 << 26, env: { ...process.env, EXCEL_ENGINE: '1' } });
      expect(gen.status, gen.stdout + gen.stderr).toBe(0);
      const manifest = JSON.parse(fs.readFileSync(path.join(dir, 'manifest.json'), 'utf8')) as unknown[];
      expect(manifest.length).toBeGreaterThanOrEqual(40);
      const chk = spawnSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(here, 'com/check-addsheets.ps1'), '-Dir', dir], { cwd: root, encoding: 'utf8', maxBuffer: 1 << 26, timeout: 25 * 60_000 });
      const lines = chk.stdout.split(/\r?\n/).filter((l) => l.startsWith('{')).map((l) => JSON.parse(l) as Record<string, unknown>);
      const rows = lines.filter((l) => l.type === undefined);
      expect(rows.length).toBe(manifest.length);
      const notOk = rows.filter((r) => r.ok !== true).map((r) => `${r.name}: open=${r.open} names=${r.names} newErr=${r.newErr} values=${r.valueDiffs} formulas=${r.formulaDiffs} formats=${r.formatDiffs} ${JSON.stringify(r.info)}`);
      expect(notOk).toEqual([]);
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  }, 30 * 60_000);
});
