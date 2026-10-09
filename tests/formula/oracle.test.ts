// adjustFormula() vs what Excel 365 itself produced via COM (R2 oracle: fixtures/formula/excel-results.json).
// 16 ops x (76 formulas on sheet F referencing Data + formulas on Data + defined names + CF formulas/sqref + chart series).
import { describe, expect, it } from 'vitest';
import { adjustFormula, adjustSqref, type StructOp } from '../../src/core/formula/excelRef';
import { readFixture } from './helpers';

interface OracleState {
  formulas: string[];
  dataSheet: Record<string, string>;
  names: Record<string, string>;
  cf: string[];
  dv: string | null;
  series?: string;
}
const R = readFixture<Record<string, OracleState>>('excel-results.json');

const OPS: Record<string, StructOp> = {
  ins_C: { type: 'insert', axis: 'col', at: 3, count: 1 },
  ins_CD: { type: 'insert', axis: 'col', at: 3, count: 2 },
  del_C: { type: 'delete', axis: 'col', at: 3, count: 1 },
  del_CD: { type: 'delete', axis: 'col', at: 3, count: 2 },
  del_B: { type: 'delete', axis: 'col', at: 2, count: 1 },
  mv_E_to_B: { type: 'move', axis: 'col', from: 5, count: 1, to: 2 },
  mv_B_to_F: { type: 'move', axis: 'col', from: 2, count: 1, to: 6 },
  mv_CD_to_B: { type: 'move', axis: 'col', from: 3, count: 2, to: 2 },
  mv_BC_to_G: { type: 'move', axis: 'col', from: 2, count: 2, to: 7 },
  mv_A_to_D: { type: 'move', axis: 'col', from: 1, count: 1, to: 4 },
  mv_C_to_E: { type: 'move', axis: 'col', from: 3, count: 1, to: 5 },
  mv_E_to_B_3step: { type: 'move', axis: 'col', from: 5, count: 1, to: 2 },
  mv_B_to_F_3step: { type: 'move', axis: 'col', from: 2, count: 1, to: 6 },
  ins_row3: { type: 'insert', axis: 'row', at: 3, count: 1 },
  del_row3: { type: 'delete', axis: 'row', at: 3, count: 1 },
  mv_row2_to_5: { type: 'move', axis: 'row', from: 2, count: 1, to: 6 },
};
// Known, deliberate deviations (see tests/formula/DEVIATIONS.md): Excel implements a move as insert+cut+delete, so a
// reference to the last column (XFD) / a range ending at XFC / row 1048576 is pushed off the sheet (#REF!) or
// normalised to 1:1 during a move. We keep such refs intact (a move never creates #REF!). Chart SERIES: Excel drops a
// deleted category argument (chart policy; the app warns before deleting instead).
const KNOWN = (op: string, base: string) => (op.startsWith('mv_') && /XFD1|XFC1|C1048576/.test(base)) || base.startsWith('=SERIES');
const strip = (f: string) => (f.startsWith('=') ? f.slice(1) : f);

describe('R2 oracle (Excel 365 via COM)', () => {
  it('reproduces every formula / name / CF / series result except the documented deviations', () => {
    let pass = 0, known = 0;
    const fails: string[] = [];
    const check = (label: string, op: string, base: string, got: string, want: string) => {
      if (got === want) { pass++; return; }
      if (KNOWN(op, base)) { known++; return; }
      fails.push(`${label} [${op}] ${base}  -> got ${got}  want ${want}`);
    };
    for (const [opName, op] of Object.entries(OPS)) {
      const res = R[opName];
      R.none.formulas.forEach((base, i) => {
        check('F!A' + (i + 1), opName, base, '=' + adjustFormula(strip(base), op, { hostSheet: 'F', targetSheet: 'Data' }).formula, res.formulas[i]);
      });
      const baseD = Object.values(R.none.dataSheet);
      const wantD = Object.values(res.dataSheet);
      baseD.forEach((b, i) => check('Data!*', opName, b, '=' + adjustFormula(strip(b), op, { hostSheet: 'Data', targetSheet: 'Data' }).formula, wantD[i]));
      for (const [nm, refersTo] of Object.entries(R.none.names)) {
        if (nm.startsWith('_xlfn')) continue;
        check('Name ' + nm, opName, refersTo, '=' + adjustFormula(strip(refersTo), op, { hostSheet: null, targetSheet: 'Data' }).formula, res.names[nm]);
      }
      R.none.cf.forEach((line, i) => {
        const [, f] = line.split(' | ');
        const [, want] = res.cf[i].split(' | ');
        check('CF' + i, opName, f, '=' + adjustFormula(strip(f), op, { hostSheet: 'Data', targetSheet: 'Data' }).formula, want);
      });
      if (op.type !== 'move') {
        const [sq] = R.none.cf[0].split(' | ');
        const [wantSq] = res.cf[0].split(' | ');
        check('CFsqref', opName, sq, adjustSqref(sq, op, ','), wantSq);
      }
      if (res.series) check('SERIES', opName, R.none.series!, '=' + adjustFormula(strip(R.none.series!), op, { hostSheet: 'F', targetSheet: 'Data' }).formula, res.series);
    }
    expect(fails).toEqual([]);
    expect(pass).toBe(1393);
    expect(known).toBe(20);
  });
});
