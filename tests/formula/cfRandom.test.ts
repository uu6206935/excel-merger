// Randomized differential run against Excel 365: 899 scenarios = one or two conditional-format rules (one
// <conditionalFormatting> block) on one or two areas, `=5` / `=ref>0` / `=ref>ref` with relative and absolute-column
// references, then one column insert / delete / move (p6 single rules, p7 blocks of two rules) or 1 to 3 successive
// ops (p8: the output of an op is the input of the next, like the op log of the editor)
// (fixtures/formula/cf-random-excel-results.json; generator com/make-cf-random-spec.mjs + excel-cf-probe.ps1 +
// extract-cf-probe.mjs). Every scenario is compared per cell (effective rule of each cell); the few that differ are
// known quirks (DEVIATIONS.md section 11). CFR_REPORT=1 prints every mismatch.
import { describe, expect, it } from 'vitest';
import { probeDiffs, type ProbeScenario } from './cfHelpers';
import { readFixture } from './helpers';

const FX = readFixture<{ scenarios: ProbeScenario[] }>('cf-random-excel-results.json');

/** scenario id -> why Excel's result is not reproduced (each must still differ) */
const KNOWN: Record<string, string> = {
  // two-area block with a move landing right after the second area and formulas that mix references with an absolute
  // column: Excel takes the cells that joined the block at the destination out again, we keep them (cfProbe p4)
  p7_042: 'dest-piece', p7_071: 'dest-piece', p7_197: 'dest-piece', p8_078: 'dest-piece', p8_276: 'dest-piece',
  // a moved piece that is a whole area of a two-area rule, one of its cells referring to another cell of the piece:
  // Excel writes #REF! for that cell (after 3 ops), we keep the reference following the piece
  p8_168: 'moved-piece-ref',
};

describe('CF vs Excel 365: 899 random rules x column ops', () => {
  it('uses the whole fixture', () => {
    expect(FX.scenarios.length).toBe(899);
    expect(FX.scenarios.filter((s) => s.ops.length > 1).length).toBeGreaterThan(150);
    const ops = new Set(FX.scenarios.map((s) => s.ops[0].kind));
    expect(ops).toEqual(new Set(['ins', 'del', 'mov']));
    expect(FX.scenarios.filter((s) => s.rules.length === 2).length).toBeGreaterThan(50);
    expect(FX.scenarios.filter((s) => s.rules[0].addr.includes(',')).length).toBeGreaterThan(50);
  });

  it('every scenario equals Excel except the listed quirks', () => {
    const bad: string[] = [];
    let pass = 0;
    for (const sc of FX.scenarios) {
      const d = probeDiffs(sc);
      if (!d.length) { pass++; continue; }
      if (process.env.CFR_REPORT) console.log(`--- ${sc.id} ${sc.rules.map((r) => r.addr + ' ' + r.f).join(' ; ')} ${JSON.stringify(sc.ops[0])}\n  ${d.slice(0, 4).join('\n  ')}`);
      if (!KNOWN[sc.id]) bad.push(`${sc.id} ${sc.rules.map((r) => r.addr + ' ' + r.f).join(' ; ')} ${JSON.stringify(sc.ops[0])}: ${d.slice(0, 2).join(' | ')}`);
    }
    expect(bad).toEqual([]);
    expect(pass).toBe(899 - Object.keys(KNOWN).length);
  });

  it('every listed quirk still differs', () => {
    const byId = new Map(FX.scenarios.map((s) => [s.id, s]));
    expect(Object.keys(KNOWN).filter((id) => !byId.has(id) || probeDiffs(byId.get(id)!).length === 0)).toEqual([]);
  });
});
