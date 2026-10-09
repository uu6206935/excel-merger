// Performance: 100,000 realistic formulas adjusted by one structural op in < 3 s (target sheet of 10,000 rows ×
// 40 columns with formula columns + formulas on other sheets). Also the helpers used over whole sheets.
// Timings are the BEST of up to 3 runs (a run under budget ends the measurement early): the full suite runs many
// workers in parallel and one run can be slowed down by the machine, not by the code. The budgets are unchanged.
import { describe, expect, it } from 'vitest';
import { adjustFormula, adjustFormulaSeq, planMove, lookupWarnings, refsTouchingColumns, type StructOp } from '../../src/core/formula/excelRef';
import { toFileForm, toUiForm } from '../../src/core/formula/forms';

const N = 100_000;
const templates = [
  (r: number) => `IF($B${r}="済",VLOOKUP($A${r},'集計 シート'!$A:$F,3,FALSE),SUM(C${r}:E${r})*TAX_RATE)`,
  (r: number) => `_xlfn.XLOOKUP($A${r},Data!$A$2:$A$10001,Data!$K$2:$K$10001,"")`,
  (r: number) => `ROUND(F${r}*G${r}/100,0)+IFERROR(INDEX(Data!$C:$C,MATCH(A${r},Data!$A:$A,0)),0)`,
  (r: number) => `TEXT(H${r},"yyyy/mm/dd")&" "&Data!AB${r}&"様"`,
  (r: number) => `SUMIFS(Data!$M:$M,Data!$A:$A,$A${r},Data!$D:$D,">="&$D$1)`,
  (r: number) => `_xlfn.LET(_xlpm.x,C${r}+D${r},IF(_xlpm.x>0,_xlpm.x,""))`,
  (r: number) => `COUNTIF($B$2:$B$10001,B${r})>1`,
  (r: number) => `Data!AN${r}*2`,
];
const formulas = Array.from({ length: N }, (_, i) => templates[i % templates.length]((i % 10000) + 2));
const hostOf = (i: number) => (i % 3 === 0 ? 'Other' : 'Data');

const time = (fn: () => void) => { const t0 = performance.now(); fn(); return performance.now() - t0; };
/** Best (minimum) time of up to `runs` runs; stops as soon as one run is within `budget` ms. */
function best(fn: () => void, budget: number, runs = 3): number {
  let min = Infinity;
  for (let i = 0; i < runs; i++) {
    min = Math.min(min, time(fn));
    if (min < budget) break;
  }
  return min;
}

describe('performance', () => {
  it('adjusts 100,000 formulas by one move in < 3 s', () => {
    const op: StructOp = { type: 'move', axis: 'col', from: 3, count: 2, to: 30 };
    let changed = 0;
    time(() => { for (let i = 0; i < 2000; i++) adjustFormula(formulas[i], op, { hostSheet: hostOf(i), targetSheet: 'Data' }); }); // warm-up
    const ms = best(() => {
      changed = 0;
      for (let i = 0; i < N; i++) if (adjustFormula(formulas[i], op, { hostSheet: hostOf(i), targetSheet: 'Data' }).changed) changed++;
    }, 3000);
    console.log(`100,000 formulas adjusted in ${ms.toFixed(0)} ms (${(ms / 100).toFixed(1)} us/formula), changed=${changed}`);
    expect(changed).toBeGreaterThan(N / 2);
    expect(ms).toBeLessThan(3000);
  });
  it('insert and delete ops are as fast', () => {
    for (const op of [{ type: 'insert', axis: 'col', at: 5, count: 3 }, { type: 'delete', axis: 'col', at: 7, count: 1 }] as StructOp[]) {
      const ms = best(() => { for (let i = 0; i < N; i++) adjustFormula(formulas[i], op, { hostSheet: hostOf(i), targetSheet: 'Data' }); }, 3000);
      expect(ms).toBeLessThan(3000);
    }
  });
  it('a 3-block multi-selection move (adjustFormulaSeq) over 100,000 formulas stays usable', () => {
    const ops = planMove([2, 5, 6, 9], 20);
    expect(ops.length).toBe(3);
    const ms = best(() => { for (let i = 0; i < N; i++) adjustFormulaSeq(formulas[i], ops, { hostSheet: hostOf(i), targetSheet: 'Data' }); }, 4500);
    console.log(`adjustFormulaSeq (3 ops) x 100,000: ${ms.toFixed(0)} ms`);
    expect(ms).toBeLessThan(4500); // tokenized once (about 2x one op)
  });
  it('dependents / lookup checks and form conversion over 100,000 formulas', () => {
    const op: StructOp = { type: 'delete', axis: 'col', at: 3, count: 1 };
    const ms1 = best(() => { for (let i = 0; i < N; i++) refsTouchingColumns(formulas[i], hostOf(i), 'Data', [3]); }, 3000);
    const ms2 = best(() => { for (let i = 0; i < N; i++) lookupWarnings(formulas[i], op, { hostSheet: hostOf(i), targetSheet: 'Data' }); }, 3000);
    const ms3 = best(() => { for (let i = 0; i < N; i++) toFileForm(toUiForm(formulas[i])); }, 6000);
    console.log(`refsTouchingColumns ${ms1.toFixed(0)} ms, lookupWarnings ${ms2.toFixed(0)} ms, toUiForm+toFileForm ${ms3.toFixed(0)} ms (x100,000)`);
    expect(ms1).toBeLessThan(3000);
    expect(ms2).toBeLessThan(3000);
    expect(ms3).toBeLessThan(12000); // ~2-3 s alone; wide margin so the parallel full suite does not flake
  });
});
