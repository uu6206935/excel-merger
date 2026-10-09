/**
 * Excel golden test: replayLog(fixture, ops) must equal the workbook Excel itself saved after the same ops
 * (tests/fixtures/golden/out/<case>.xlsx, produced by tests/golden/com/run-ops.ps1). Compared through the XML
 * snapshot (formulas, values, styles, CF/DV per cell, names, tables, charts, pivots, merges, widths, panes, filters,
 * hyperlinks, notes, anchors). No Excel is needed at test time.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { replayLog } from '../../src/core/xlsx/replay';
import { CASES } from './cases';
import { diffSnap, snapshot } from './snapshot';

const here = path.dirname(fileURLToPath(import.meta.url));
const FIX = path.resolve(here, '../fixtures/golden');
const read = (p: string) => new Uint8Array(fs.readFileSync(p));

describe('replayLog vs Excel goldens', () => {
  for (const c of CASES) {
    it(c.name, async () => {
      const input = read(path.join(FIX, c.fixture));
      if (c.expectError) {
        await expect(replayLog(input, c.ops)).rejects.toThrow(/[ぁ-んァ-ン一-龥]/);
        return;
      }
      const ours = await replayLog(input, c.ops);
      const golden = read(path.join(FIX, c.golden ?? `out/${c.name}.xlsx`));
      const [a, b] = await Promise.all([snapshot(ours), snapshot(golden)]);
      const known = Object.keys(c.known ?? {});
      const diffs = diffSnap(a, b).filter((d) => !known.some((k) => d.path.startsWith(k)));
      expect(diffs.map((d) => `${d.path}: ours=${JSON.stringify(d.a)} excel=${JSON.stringify(d.b)}`)).toEqual([]);
      // every output: no calcChain, full recalculation on load
      expect(a.calcChain).toBe(false);
      expect(a.fullCalcOnLoad).toBe(true);
    });
  }

  it('documents every known deviation with a reason and still differs there (not stale)', async () => {
    for (const c of CASES.filter((x) => x.known && Object.keys(x.known).length)) {
      const ours = await replayLog(read(path.join(FIX, c.fixture)), c.ops);
      const a = await snapshot(ours);
      const b = await snapshot(read(path.join(FIX, c.golden ?? `out/${c.name}.xlsx`)));
      const diffs = diffSnap(a, b);
      for (const [k, why] of Object.entries(c.known!)) {
        expect(why.length).toBeGreaterThan(20);
        expect(diffs.some((d) => d.path.startsWith(k)), `${c.name}: known deviation ${k} no longer differs`).toBe(true);
      }
    }
  });
});
