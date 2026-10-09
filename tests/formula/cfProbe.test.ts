// Conditional formats after ONE column move, small rule families against Excel 365
// (fixtures/formula/cf-probe-excel-results.json; generator: fixtures/formula/com/{make-cf-probe-spec.mjs,
// excel-cf-probe.ps1, extract-cf-probe.mjs}; the move is the Excel runner's clipboard-free sequence: insert blank
// columns, Range.Cut(Destination), delete the emptied source).
//   p1 / p2  which cells of a rule share a formula after the move, and what a reference to the moved column becomes
//            (the part of a rule right of the moved column is ONE range object: the formula of its first cell; a
//            relative reference to the moved column dies in the delete step, an absolute one follows it)
//   p3 / p4  absolute and mixed references around the moved column, and whether the cells that joined the rule by
//            format inheritance at the destination stay in it
//   p5       several rules on one range (one <conditionalFormatting> block): the block is fragmented as a whole and
//            a rule referring to the moved column makes the destination leave the block for every rule in it
// Every mismatch must be listed in KNOWN with its reason (and must still differ). CFP_REPORT=1 prints them.
import { describe, expect, it } from 'vitest';
import { cover, probeDiffs, probeOurs, type ProbeScenario } from './cfHelpers';
import { readFixture } from './helpers';

const FX = readFixture<{ scenarios: ProbeScenario[] }>('cf-probe-excel-results.json');

/** scenario id -> why Excel's result is not reproduced (each must still differ) */
const KNOWN: Record<string, string> = {
  // Excel drops the cells left of the moved column from a rule whose ABSOLUTE column reference points at the moved
  // column (the whole rule shrinks to the columns from the moved cell on): the CF_INSIDE_MOVE family of Excel
  // quirks (tests/golden/RESULTS.md). We keep "every cell keeps its rule" (cell-set semantics).
  p3_absD_sD_xH: 'abs-ref-to-moved-column',
  p3_absE_sE_xH: 'abs-ref-to-moved-column',
  // a formula mixing a relative reference to the moved column with an absolute one before / between source and
  // destination: Excel takes the destination cell out of the rule, we keep it (one extra cell covered)
  p4_07: 'dest-piece', p4_09: 'dest-piece', p4_16: 'dest-piece', p4_26: 'dest-piece',
};

describe('CF after one column move vs Excel 365 (small rule families)', () => {
  it('uses the whole fixture', () => {
    expect(FX.scenarios.length).toBe(129);
    expect(new Set(FX.scenarios.map((s) => s.id.slice(0, 2)))).toEqual(new Set(['p1', 'p2', 'p3', 'p4', 'p5']));
    for (const s of FX.scenarios) expect(s.excel.length).toBeGreaterThan(0);
  });

  it('every scenario equals Excel except the listed quirks', () => {
    const bad: string[] = [];
    let pass = 0;
    for (const sc of FX.scenarios) {
      const d = probeDiffs(sc);
      if (!d.length) { pass++; continue; }
      if (process.env.CFP_REPORT) console.log(`--- ${sc.id} ${sc.rules.map((r) => r.addr + ' ' + r.f).join(' ; ')} ${JSON.stringify(sc.ops[0])}\n  ${d.slice(0, 4).join('\n  ')}`);
      if (!KNOWN[sc.id]) bad.push(`${sc.id}: ${d.slice(0, 3).join(' | ')}`);
    }
    expect(bad).toEqual([]);
    expect(pass).toBe(FX.scenarios.length - Object.keys(KNOWN).length);
  });

  it('every listed quirk still differs', () => {
    const byId = new Map(FX.scenarios.map((s) => [s.id, s]));
    const stale = Object.keys(KNOWN).filter((id) => !byId.has(id) || probeDiffs(byId.get(id)!).length === 0);
    expect(stale).toEqual([]);
  });

  it('rule blocks: a sibling that refers to the moved column takes the destination out of the block', () => {
    const get = (id: string) => FX.scenarios.find((s) => s.id === id)!;
    const at = (m: Map<string, string>) => (m.get('5,2') ?? '').split(' || ');
    // alone, `=C2>0` keeps the cell it gained at the destination: E2 is covered by the moved cell (B2>0 -> C[-3]) and by
    // the spread one (C[2])
    expect(at(cover(probeOurs(get('p5_a_alone')))).filter((x) => x.includes('C[2]') || x.includes('C[-3]')).length).toBe(2);
    // next to `=A2>0` (whose cell C refers to the moved B) the block loses the destination: only the moved cells remain
    expect(at(cover(probeOurs(get('p5_ab'))))).toEqual(['cf:expression|{cell:R[0]C[-3]:R[0]C[-3]}>0', 'cf:expression|{cell:R[0]C[-4]:R[0]C[-4]}>0']);
    for (const id of ['p5_ab', 'p5_ba', 'p5_a_const']) expect(probeDiffs(get(id))).toEqual([]);
  });
});
