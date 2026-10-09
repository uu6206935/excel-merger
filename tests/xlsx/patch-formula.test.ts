/**
 * The patcher rewrites every formula ONCE with the whole op list (adjustFormulaOps). It must give exactly what
 * excelRef.adjustFormula gives when applied op by op (the doc module and the Excel engine rely on that).
 */
import { describe, expect, it } from 'vitest';
import { adjustFormula, adjustFormulaSeq, type StructOp, type TableRef } from '../../src/core/formula/excelRef';
import { adjustFormulaOps, extendArea, parseSqref, renderSqref, type OpEntry } from '../../src/core/xlsx/patch/formula';

function rng(seed: number) {
  let s = seed >>> 0;
  return () => { s = (s * 1664525 + 1013904223) >>> 0; return s / 2 ** 32; };
}

const SHEETS = ['Data', 'Other', "'集計 シート'", '名簿'];
const sheetName = (q: string) => q.replace(/^'|'$/g, '');

function randomRef(r: () => number): string {
  const col = () => String.fromCharCode(65 + Math.floor(r() * 12));
  const row = () => String(1 + Math.floor(r() * 30));
  const d = () => (r() < 0.3 ? '$' : '');
  const pre = r() < 0.5 ? SHEETS[Math.floor(r() * SHEETS.length)] + '!' : '';
  const k = r();
  if (k < 0.45) return `${pre}${d()}${col()}${d()}${row()}`;
  if (k < 0.8) return `${pre}${d()}${col()}${d()}${row()}:${d()}${col()}${d()}${row()}`;
  if (k < 0.9) return `${pre}${d()}${col()}:${d()}${col()}`;
  return `${pre}${d()}${row()}:${d()}${row()}`;
}

function randomFormula(r: () => number): string {
  const parts: string[] = [];
  const n = 1 + Math.floor(r() * 4);
  for (let i = 0; i < n; i++) {
    const k = r();
    if (k < 0.6) parts.push(randomRef(r));
    else if (k < 0.7) parts.push('tbl[点数]');
    else if (k < 0.75) parts.push('tbl[[#This Row],[氏名]:[点数]]');
    else if (k < 0.8) parts.push('"Data!C5"');
    else if (k < 0.85) parts.push('[1]Data!C5');
    else if (k < 0.9) parts.push('Data:Other!C5');
    else parts.push(String(Math.floor(r() * 100)));
  }
  return `SUM(${parts.join(',')})`;
}

function randomOp(r: () => number): StructOp {
  const axis = r() < 0.7 ? 'col' : 'row';
  const k = r();
  const count = 1 + Math.floor(r() * 3);
  if (k < 0.35) return { type: 'insert', axis, at: 1 + Math.floor(r() * 14), count };
  if (k < 0.7) return { type: 'delete', axis, at: 1 + Math.floor(r() * 14), count };
  const from = 1 + Math.floor(r() * 12);
  let to = 1 + Math.floor(r() * 15);
  if (to >= from && to <= from + count) to = from + count + 1;
  return { type: 'move', axis, from, count, to };
}

describe('adjustFormulaOps == excelRef.adjustFormula applied op by op', () => {
  it('cell formulas (hosted on each sheet), 3000 random formulas x random op sequences', () => {
    const r = rng(12345);
    const fails: string[] = [];
    for (let i = 0; i < 3000; i++) {
      const f = randomFormula(r);
      const nOps = 1 + Math.floor(r() * 4);
      const seq: Array<{ sheet: string; op: StructOp }> = [];
      for (let k = 0; k < nOps; k++) seq.push({ sheet: sheetName(SHEETS[Math.floor(r() * SHEETS.length)]), op: randomOp(r) });
      const host = sheetName(SHEETS[i % SHEETS.length]);
      let want = f;
      for (const e of seq) want = adjustFormula(want, e.op, { hostSheet: host, targetSheet: e.sheet }).formula;
      const got = adjustFormulaOps(f, seq.map((e) => ({ sheet: e.sheet, op: e.op })), host).text;
      if (got !== want) fails.push(`${f} @${host} ${JSON.stringify(seq)}: got ${got} want ${want}`);
    }
    expect(fails.slice(0, 10)).toEqual([]);
  });

  it('defined names (relative parts are never adjusted)', () => {
    const r = rng(777);
    const fails: string[] = [];
    for (let i = 0; i < 2000; i++) {
      const f = randomRef(r).replace(/^(?![A-Za-z'集名]+!)/, 'Data!');
      const seq = [{ sheet: 'Data', op: randomOp(r) }, { sheet: 'Data', op: randomOp(r) }];
      let want = f;
      for (const e of seq) want = adjustFormula(want, e.op, { hostSheet: null, targetSheet: e.sheet, isName: true }).formula;
      const got = adjustFormulaOps(f, seq, null, 'name').text;
      if (got !== want) fails.push(`${f}: got ${got} want ${want}`);
    }
    expect(fails.slice(0, 10)).toEqual([]);
  });

  it('structured references to deleted table columns become #REF! (like excelRef with ctx.tables)', () => {
    const tables: TableRef[] = [{ name: 'tbl', sheet: 'Data', ref: 'C3:F10', columns: ['ID', '氏名', '点数', '合計'] }];
    const cases: Array<[string, StructOp]> = [
      ['SUM(tbl[点数])', { type: 'delete', axis: 'col', at: 5, count: 1 }],
      ['SUM(tbl[[#This Row],[氏名]:[点数]])', { type: 'delete', axis: 'col', at: 4, count: 1 }],
      ['SUM(tbl[点数])', { type: 'delete', axis: 'col', at: 3, count: 1 }],
      ['ROWS(tbl)', { type: 'delete', axis: 'col', at: 3, count: 4 }],
      ['SUM(tbl[合計])+Data!C4', { type: 'delete', axis: 'col', at: 1, count: 1 }],
    ];
    for (const [f, op] of cases) {
      const want = adjustFormulaSeq(f, [op], { hostSheet: 'Other', targetSheet: 'Data', tables }).formula;
      const hits = new Map<string, { whole: boolean; cols: Set<string> }>();
      const lo = 3, hi = 6, s1 = (op as { at: number }).at, s2 = s1 + op.count - 1;
      if (!(s2 < lo || s1 > hi)) {
        const cols = new Set<string>();
        for (let c = Math.max(lo, s1); c <= Math.min(hi, s2); c++) cols.add(tables[0].columns[c - lo].toLowerCase());
        hits.set('tbl', { whole: s1 <= lo && s2 >= hi, cols });
      }
      const e: OpEntry = { sheet: 'Data', op, tableHits: hits.size ? hits : undefined };
      expect(adjustFormulaOps(f, [e], 'Other').text).toBe(want);
    }
  });

  it('only applies entries from the given index (cells written after earlier ops)', () => {
    const ops: OpEntry[] = [
      { sheet: 'Data', op: { type: 'insert', axis: 'col', at: 1, count: 1 } },
      { sheet: 'Data', op: { type: 'insert', axis: 'col', at: 1, count: 2 } },
    ];
    expect(adjustFormulaOps('Data!B2', ops, null, 'cell', 0).text).toBe('Data!E2');
    expect(adjustFormulaOps('Data!B2', ops, null, 'cell', 1).text).toBe('Data!D2');
    expect(adjustFormulaOps('Data!B2', ops, null, 'cell', 2).text).toBe('Data!B2');
  });

  it('skipSheet leaves references to that sheet alone (already handled by the anchored CF/DV pass)', () => {
    const ops: OpEntry[] = [{ sheet: 'Data', op: { type: 'insert', axis: 'col', at: 1, count: 1 } }, { sheet: 'Lists', op: { type: 'insert', axis: 'row', at: 1, count: 1 } }];
    expect(adjustFormulaOps('$B2>Lists!$A$1', ops, 'Data', 'cell', 0, 'Data').text).toBe('$B2>Lists!$A$2');
  });
});

describe('premerge range extension', () => {
  const e: OpEntry = { sheet: 'S', op: { type: 'insert', axis: 'row', at: 7, count: 2 }, ext: { firstDataRow: 2, lastDataRow: 6 } };
  it('grows ranges ending at the last data row and starting at/above the first data row', () => {
    expect(adjustFormulaOps('SUM(D2:D6)', [e], 'S').text).toBe('SUM(D2:D8)');
    expect(adjustFormulaOps('SUM($D$1:$D$6)', [e], 'S').text).toBe('SUM($D$1:$D$8)');
    expect(adjustFormulaOps('SUM(D3:D6)', [e], 'S').text).toBe('SUM(D3:D6)');
    expect(adjustFormulaOps('SUM(D2:D5)', [e], 'S').text).toBe('SUM(D2:D5)');
    expect(adjustFormulaOps('D6', [e], 'S').text).toBe('D6');
    expect(adjustFormulaOps('SUM(D2:D9)', [e], 'S').text).toBe('SUM(D2:D11)');
    expect(adjustFormulaOps('Scores', [e], 'S').text).toBe('Scores');
    expect(adjustFormulaOps('S!$D$2:$D$6', [e], null, 'name').text).toBe('S!$D$2:$D$8');
    expect(adjustFormulaOps('S!D2:D6', [e], null, 'name').text).toBe('S!D2:D6'); // relative name: never adjusted
  });
  it('extendArea is a no-op for other ops', () => {
    const a = parseSqref('D2:D6')[0];
    expect(extendArea(a, a, { sheet: 'S', op: { type: 'insert', axis: 'col', at: 7, count: 2 }, ext: e.ext })).toBe(a);
    expect(renderSqref(parseSqref('$A$1:B2 C3'))).toBe('A1:B2 C3');
  });
});
