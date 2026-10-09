// Helpers other modules use: refsTouchingColumns (dependents checks), lookupWarnings (VLOOKUP & co. numeric
// positions that Excel does not adjust), adjustTableRef, adjustNameArea, hostCell resolution of [@col].
import { describe, expect, it } from 'vitest';
import {
  refsTouchingColumns, lookupWarnings, adjustTableRef, adjustFormula, adjustFormulaSeq, advanceContext, adjustAnchoredRule,
  parseAreaText, renderArea, sqrefTopLeft, planMove, type StructOp, type TableRef, type AdjustContext, type AdjustResult,
} from '../../src/core/formula/excelRef';

const T: TableRef[] = [{ name: 'tblMain', sheet: 'Data', ref: 'C3:H10', columns: ['氏名', '大学', '点数', 'メール', '備考 メモ', '合計'] }];
const colsOf = (r: ReturnType<typeof refsTouchingColumns>) => r.map((x) => `${x.text}=${x.cols.join(',')}`);

describe('refsTouchingColumns', () => {
  it('A1 refs on the target sheet only', () => {
    expect(colsOf(refsTouchingColumns('SUM(Data!B2:D9)+Data!F1+E1+"Data!C1"', 'Other', 'Data', [3]))).toEqual(['Data!B2:D9=3']);
    expect(colsOf(refsTouchingColumns('E1+$C:$C+5:5+A1', 'Data', 'Data', [3, 5]))).toEqual(['E1=5', '$C:$C=3', '5:5=3,5']);
    expect(refsTouchingColumns('Data:Other!C1+[1]Data!C1+Other!C1', 'Other', 'Data', [3])).toEqual([]);
    expect(refsTouchingColumns('C1', 'Data', 'data', [])).toEqual([]);
    expect(colsOf(refsTouchingColumns("'データ 表'!$C$2", null, 'データ 表', new Set([3])))).toEqual(["'データ 表'!$C$2=3"]);
  });
  it('structured refs through the table layout', () => {
    expect(colsOf(refsTouchingColumns('SUM(tblMain[点数])', 'Other', 'Data', [5], { tables: T }))).toEqual(['tblMain[点数]=5']);
    expect(refsTouchingColumns('SUM(tblMain[点数])', 'Other', 'Data', [4], { tables: T })).toEqual([]);
    expect(colsOf(refsTouchingColumns('SUM(tblMain[[#This Row],[大学]:[点数]])', 'Other', 'Data', [4], { tables: T }))).toEqual(['tblMain[[#This Row],[大学]:[点数]]=4']);
    expect(colsOf(refsTouchingColumns('ROWS(tblMain[])', 'Other', 'Data', [8, 9], { tables: T }))).toEqual(['tblMain[]=8']);
    expect(colsOf(refsTouchingColumns('[@点数]*2', 'Data', 'Data', [5], { tables: T, hostCell: { row: 4, col: 8 } }))).toEqual(['[@点数]=5']);
    expect(refsTouchingColumns('[@点数]*2', 'Data', 'Data', [5], { tables: T, hostCell: { row: 20, col: 8 } })).toEqual([]);
  });
});

describe('lookupWarnings', () => {
  const ctx = { hostSheet: 'Other', targetSheet: 'Data' };
  const mv = (from: number, to: number, count = 1): StructOp => ({ type: 'move', axis: 'col', from, count, to });
  const ins = (at: number, count = 1): StructOp => ({ type: 'insert', axis: 'col', at, count });
  const del = (at: number, count = 1): StructOp => ({ type: 'delete', axis: 'col', at, count });
  const w = (f: string, op: StructOp, c = ctx) => lookupWarnings(f, op, c).map((x) => `${x.fn} ${x.ref}`);
  it('VLOOKUP with a literal column index: exact', () => {
    const f = 'VLOOKUP(A1,Data!A:E,3,FALSE)';
    expect(w(f, mv(4, 2))).toEqual(['VLOOKUP Data!A:E']);  // A D B C E: position 3 is now B
    expect(w(f, mv(5, 4))).toEqual([]);                    // D and E swap: C still at 3
    expect(w(f, del(5))).toEqual([]);                       // E deleted: C still 3rd
    expect(w(f, del(2))).toEqual(['VLOOKUP Data!A:E']);
    expect(w(f, ins(3))).toEqual(['VLOOKUP Data!A:E']);     // inserted before C
    expect(w(f, ins(4))).toEqual([]);                       // after C
    expect(w(f, ins(6))).toEqual([]);                       // outside
    expect(w(f, mv(7, 1))).toEqual([]);                     // G moved before A: the range shifts as a whole
    expect(w(f, mv(1, 6, 5))).toEqual([]);                  // the whole table moves
    expect(w(f, mv(7, 3))).toEqual(['VLOOKUP Data!A:E']);   // G lands inside the range before C
    expect(w(f, { type: 'insert', axis: 'row', at: 1, count: 1 })).toEqual([]);
    expect(w(f, del(2), { hostSheet: 'Other', targetSheet: 'Other' })).toEqual([]);
    expect(lookupWarnings(f, del(2), ctx)[0].message).toMatch(/VLOOKUP の列番号は数値のため自動では変わりません/);
  });
  it('VLOOKUP with a computed index warns on any change inside the range', () => {
    // (the MATCH over the whole header row also reports: the column number it returns changes)
    expect(w('VLOOKUP(A1,Data!$A$1:$E$99,MATCH("x",Data!1:1,0),0)', del(5))).toEqual(['VLOOKUP Data!$A$1:$E$99', 'MATCH Data!1:1']);
    expect(w('VLOOKUP(A1,Data!1:5,3,0)', mv(4, 2))).toEqual(['VLOOKUP Data!1:5']);     // whole rows: columns reordered inside
    expect(w('VLOOKUP(A1,Data!1:5,3,0)', mv(5, 4))).toEqual([]);
    expect(w('INDEX(Data!2:2,5)', del(1))).toEqual(['INDEX Data!2:2']);
    expect(w('INDEX(Data!2:2,5)', del(9))).toEqual([]);
    expect(w('VLOOKUP(A1,Data!$A$1:$E$99,B1,0)', ins(9))).toEqual([]);
  });
  it('HLOOKUP / INDEX / MATCH / XLOOKUP / OFFSET / CHOOSE', () => {
    expect(w('HLOOKUP(1,Data!A1:E5,3,0)', { type: 'insert', axis: 'row', at: 2, count: 1 })).toEqual(['HLOOKUP Data!A1:E5']);
    expect(w('HLOOKUP(1,Data!A1:E5,3,0)', ins(2))).toEqual([]);
    expect(w('INDEX(Data!A1:E10,2,4)', del(2))).toEqual(['INDEX Data!A1:E10']);
    expect(w('INDEX(Data!A1:E10,2,4)', del(5))).toEqual([]);
    expect(w('INDEX(Data!A1:E1,3)', ins(2))).toEqual(['INDEX Data!A1:E1']);
    expect(w('INDEX(Data!A1:A10,3)', ins(1))).toEqual([]);
    expect(w('INDEX(Data!A1:E10,2,0)', del(2))).toEqual([]);
    expect(w('MATCH(1,Data!A1:E1,0)', mv(3, 1))).toEqual(['MATCH Data!A1:E1']);
    expect(w('MATCH(1,Data!A:A,0)', mv(3, 1))).toEqual([]);
    expect(w('_xlfn.XLOOKUP(1,Data!A:A,Data!B:D)', ins(3))).toEqual(['XLOOKUP Data!B:D']);
    expect(w('_xlfn.XLOOKUP(1,Data!A:A,Data!B:D)', ins(6))).toEqual([]);
    expect(w('OFFSET(Data!A1,0,2)', del(2))).toEqual(['OFFSET Data!A1']);
    expect(w('OFFSET(Data!A1,0,2)', ins(5))).toEqual([]);
    expect(w('OFFSET(Data!A1,1,0)', del(1 + 1))).toEqual([]);
    expect(w('OFFSET(Data!A1,0,B1)', del(3))).toEqual(['OFFSET Data!A1']);
    expect(w('CHOOSE(2,Data!A1:C1,Data!E1)', del(2))).toEqual(['CHOOSE Data!A1:C1']);
    expect(w('SUM(Data!A1:C1)', del(2))).toEqual([]);
  });
  it('tables and structured refs (columns reordered INSIDE a table keep the table range)', () => {
    // tblMain on Data!C3:H10: C 氏名 | D 大学 | E 点数 | F メール | G 備考 メモ | H 合計
    const tc = { hostSheet: 'Other', targetSheet: 'Data', tables: T };
    const wt = (f: string, op: StructOp) => lookupWarnings(f, op, tc).map((x) => `${x.fn} ${x.ref}`);
    expect(wt('VLOOKUP(A1,tblMain[],3,FALSE)', mv(7, 4))).toEqual(['VLOOKUP tblMain[]']);    // 3rd column is now 大学
    expect(wt('VLOOKUP(A1,tblMain,3,FALSE)', mv(7, 4))).toEqual(['VLOOKUP tblMain']);        // UI form (bare table name)
    expect(wt('VLOOKUP(A1,tblMain[],3,FALSE)', mv(8, 7))).toEqual([]);                       // swap of the last two
    expect(wt('VLOOKUP(A1,tblMain[#All],3,FALSE)', ins(4))).toEqual(['VLOOKUP tblMain[#All]']);
    expect(wt('VLOOKUP(A1,tblMain[],3,FALSE)', ins(9))).toEqual([]);                         // right after the table
    expect(wt('VLOOKUP(A1,tblMain[],3,FALSE)', mv(3, 12, 6))).toEqual([]);                   // the whole table moves
    expect(wt('INDEX(tblMain[[大学]:[メール]],2,2)', mv(6, 4))).toEqual(['INDEX tblMain[[大学]:[メール]]']);
    expect(wt('INDEX(tblMain[[大学]:[メール]],2,1)', mv(8, 3))).toEqual([]);                 // 大学 stays first of the span
    expect(wt('MATCH("x",tblMain[#Headers],0)', mv(5, 4))).toEqual(['MATCH tblMain[#Headers]']);
    expect(wt('SUM(tblMain[点数])+VLOOKUP(1,tblMain[点数],1,0)', mv(5, 4))).toEqual([]);     // single column
    expect(wt('VLOOKUP(A1,tblMain[],3,FALSE)', { type: 'insert', axis: 'row', at: 4, count: 1 })).toEqual([]);
    expect(lookupWarnings('VLOOKUP(A1,tblMain[],3,FALSE)', mv(7, 4), ctx)).toEqual([]);      // tables not given
    // calculated column inside the table: [@...] resolves through hostCell
    expect(lookupWarnings('INDEX(tblMain[@],1,3)', mv(7, 4), { ...tc, hostSheet: 'Data', hostCell: { row: 5, col: 8 } }).map((x) => x.fn)).toEqual(['INDEX']);
  });
  it('defined names through opts.resolveName', () => {
    const names: Record<string, string> = { 社員マスタ: 'Data!$A:$E', 他シート: 'Other!$A:$E', 行: 'Data!$2:$2' };
    const resolveName = (n: string) => names[n] ?? null;
    const f = 'VLOOKUP(A1,社員マスタ,3,0)';
    expect(lookupWarnings(f, mv(4, 2), ctx, { resolveName }).map((x) => `${x.fn} ${x.ref}`)).toEqual(['VLOOKUP 社員マスタ']);
    expect(lookupWarnings(f, mv(5, 4), ctx, { resolveName })).toEqual([]);
    expect(lookupWarnings(f, mv(4, 2), ctx)).toEqual([]);                                    // names not resolvable
    expect(lookupWarnings('VLOOKUP(A1,他シート,3,0)', mv(4, 2), ctx, { resolveName })).toEqual([]);
    expect(lookupWarnings('MATCH("x",行,0)', mv(4, 2), ctx, { resolveName }).map((x) => x.fn)).toEqual(['MATCH']);
  });
});

describe('adjustTableRef', () => {
  it('moves inside a table keep its range; other ops follow the range rules', () => {
    expect(adjustTableRef('C3:H10', { type: 'move', axis: 'col', from: 7, count: 1, to: 3 })).toBe('C3:H10');
    expect(adjustTableRef('C3:H10', { type: 'move', axis: 'col', from: 3, count: 6, to: 13 })).toBe('G3:L10');
    expect(adjustTableRef('C3:H10', { type: 'move', axis: 'col', from: 10, count: 1, to: 1 })).toBe('D3:I10');
    expect(adjustTableRef('C3:H10', { type: 'insert', axis: 'col', at: 9, count: 1 })).toBe('C3:H10');
    expect(adjustTableRef('C3:H10', { type: 'insert', axis: 'col', at: 5, count: 1 })).toBe('C3:I10');
    expect(adjustTableRef('C3:H10', { type: 'delete', axis: 'col', at: 3, count: 6 })).toBeNull();
    expect(adjustTableRef('Data!C3:H10', { type: 'delete', axis: 'col', at: 1, count: 1 })).toBe('Data!B3:G10');
  });
});

describe('adjustFormula extras', () => {
  it('defined names: relative parts are offsets (not adjusted), absolute parts follow', () => {
    const del: StructOp = { type: 'delete', axis: 'col', at: 2, count: 1 };
    const n = (f: string) => adjustFormula(f, del, { hostSheet: null, targetSheet: 'Data', isName: true }).formula;
    expect(n('Data!C5')).toBe('Data!C5');
    expect(n('Data!$C$5')).toBe('Data!$B$5');
    expect(n('Data!$C5:E5')).toBe('Data!$B5:E5');
    expect(n('Data!$B$5')).toBe('Data!#REF!');
  });
  it('structured refs to deleted table columns become #REF! (qualified, or through hostCell)', () => {
    const del: StructOp = { type: 'delete', axis: 'col', at: 5, count: 1 };
    const ctx = { hostSheet: 'Data', targetSheet: 'Data', tables: T };
    expect(adjustFormula('tblMain[[#This Row],[点数]]*2', del, ctx).formula).toBe('#REF!*2');
    expect(adjustFormula('[@点数]*2', del, { ...ctx, hostCell: { row: 4, col: 8 } }).formula).toBe('#REF!*2');
    expect(adjustFormula('[@点数]*2', del, ctx).formula).toBe('[@点数]*2');
    expect(adjustFormula('SUM(tblMain[大学])', del, ctx).createdRefError).toBe(false);
    expect(adjustFormula('ROWS(tblMain[])', { type: 'delete', axis: 'col', at: 3, count: 6 }, ctx).formula).toBe('ROWS(#REF!)');
  });
});

describe('adjustAnchoredRule options', () => {
  const rule = (sq: string, f: string) => ({ sqref: sq.split(' ').map((s) => parseAreaText(s)!), formulas: [f] });
  const show = (rs: ReturnType<typeof adjustAnchoredRule>) => rs.map((r) => `${r.sqref.map(renderArea).join(' ')}|${r.formulas.join(';')}`);
  it('every CF type and every validation grow over an insertion right after them; inheritFormat: false does not', () => {
    const op: StructOp = { type: 'insert', axis: 'col', at: 3, count: 1 };
    expect(show(adjustAnchoredRule(rule('B2:B20', 'B2<0'), op, 'S'))).toEqual(['B2:C20|B2<0']);
    expect(show(adjustAnchoredRule(rule('B2:B20', 'B2<0'), op, 'S', { kind: 'cf', type: 'cellIs' }))).toEqual(['B2:C20|B2<0']);
    expect(show(adjustAnchoredRule(rule('B2:B20', 'B2<0'), op, 'S', { kind: 'cf', type: 'colorScale' }))).toEqual(['B2:C20|B2<0']);
    expect(show(adjustAnchoredRule(rule('B2:B20', 'B2<0'), op, 'S', { kind: 'dv', type: 'list' }))).toEqual(['B2:C20|B2<0']);
    expect(show(adjustAnchoredRule(rule('B2:B20', 'B2<0'), op, 'S', { inheritFormat: false }))).toEqual(['B2:B20|B2<0']);
    expect(show(adjustAnchoredRule(rule('B2:B20', 'B2<0'), op, 'S', { kind: 'dv', type: 'list', inheritFormat: false }))).toEqual(['B2:B20|B2<0']);
    // strictly inside an area the inserted column is absorbed whatever the options
    expect(show(adjustAnchoredRule(rule('B2:D20', 'B2<0'), op, 'S', { kind: 'dv', inheritFormat: false }))).toEqual(['B2:E20|B2<0']);
  });
  it('spread onto a moved block: cellIs / expression / validations yes, the statistical CF types no', () => {
    const op: StructOp = { type: 'move', axis: 'col', from: 3, count: 1, to: 8 }; // C before H
    const g = (opts: Parameters<typeof adjustAnchoredRule>[3]) => show(adjustAnchoredRule(rule('G2:G20', 'G2<0'), op, 'S', opts));
    expect(g(undefined)).toEqual(['F2:G20|F2<0']);
    for (const type of ['cellIs', 'expression', 'containsText', 'containsBlanks']) expect(g({ kind: 'cf', type })).toEqual(['F2:G20|F2<0']);
    for (const type of ['whole', 'list', 'custom', 'date']) expect(g({ kind: 'dv', type })).toEqual(['F2:G20|F2<0']);
    for (const type of ['colorScale', 'dataBar', 'iconSet', 'top10', 'uniqueValues', 'duplicateValues', 'aboveAverage']) {
      expect(g({ kind: 'cf', type })).toEqual(['F2:F20|F2<0']);
    }
    expect(g({ kind: 'cf', type: 'cellIs', inheritFormat: false })).toEqual(['F2:F20|F2<0']);
    // round-1 option names keep working
    expect(g({ cfType: 'cellValue' })).toEqual(['F2:F20|F2<0']);
    expect(g({ cfType: 'colorScale' })).toEqual(['F2:F20|F2<0']);
    // a reference to the moved cell follows it (per-cell split when the rule moves with relative references)
    expect(show(adjustAnchoredRule(rule('L2:M20', 'L$2>C$2'), op, 'S'))).toEqual(['L2:M20|L$2>G$2']);
    // a rule whose cells do not move keeps ONE formula, anchored at its first cell, for all of them although the
    // second cell's relative reference points into the moved column (Excel, R2b c_m13: D before F)
    const d: StructOp = { type: 'move', axis: 'col', from: 4, count: 1, to: 6 };
    expect(show(adjustAnchoredRule(rule('L2:M20', 'L$2>C$2'), d, 'S'))).toEqual(['L2:M20|L$2>C$2']);
  });
  it('whole-column rules survive row ops unchanged; bounding-box anchor', () => {
    expect(show(adjustAnchoredRule(rule('E1:E1048576', 'E1>0'), { type: 'insert', axis: 'row', at: 1, count: 3 }, 'S'))).toEqual(['E1:E1048576|E1>0']);
    expect(show(adjustAnchoredRule(rule('E1:E1048576', 'E1>0'), { type: 'delete', axis: 'row', at: 5, count: 2 }, 'S'))).toEqual(['E1:E1048576|E1>0']);
    expect(sqrefTopLeft([parseAreaText('G5:G9')!, parseAreaText('D7:D20')!])).toEqual({ row: 5, col: 4 });
  });
});

describe('helpers never throw on junk', () => {
  const parts = ['VLOOKUP(', 'INDEX(', 'OFFSET(', 'MATCH(', 'CHOOSE(', 'XLOOKUP(', 'Data!', 'A:E', 'A1', '3', ',', ')', '(',
    '{', '}', '"', 'tblMain[点数]', '[@点数]', '#REF!', ' ', '0', 'B1', '1:1', "'x'!", '[1]Data!A1'];
  let seed = 7;
  const rnd = (n: number) => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed % n; };
  const ops: StructOp[] = [
    { type: 'move', axis: 'col', from: 2, count: 2, to: 6 }, { type: 'insert', axis: 'col', at: 3, count: 1 },
    { type: 'delete', axis: 'col', at: 1, count: 2 }, { type: 'delete', axis: 'row', at: 4, count: 1 },
  ];
  it('10,000 random formulas through lookupWarnings / refsTouchingColumns / adjustFormula', () => {
    for (let k = 0; k < 10000; k++) {
      let f = '';
      for (let j = 0, n = 1 + rnd(10); j < n; j++) f += parts[rnd(parts.length)];
      const op = ops[rnd(ops.length)];
      expect(() => lookupWarnings(f, op, { hostSheet: 'Data', targetSheet: 'Data' })).not.toThrow();
      expect(() => refsTouchingColumns(f, 'Data', 'Data', [1, 3, 5], { tables: T, hostCell: { row: 5, col: 5 } })).not.toThrow();
      const r = adjustFormula(f, op, { hostSheet: 'Data', targetSheet: 'Data', tables: T, hostCell: { row: 5, col: 5 } });
      expect(typeof r.formula).toBe('string');
    }
  });
});

describe('adjustFormulaSeq / advanceContext', () => {
  /** the reference implementation: one adjustFormula per op, context unchanged */
  const naive = (f: string, ops: StructOp[], ctx: AdjustContext): AdjustResult => {
    let cur: AdjustResult = { formula: f, changed: false, createdRefError: false, unsafe: false, warnings: [] };
    for (const op of ops) {
      const r = adjustFormula(cur.formula, op, ctx);
      cur = { formula: r.formula, changed: cur.changed || r.changed, createdRefError: cur.createdRefError || r.createdRefError,
        unsafe: cur.unsafe || r.unsafe, warnings: [...new Set([...cur.warnings, ...r.warnings])] };
    }
    return cur;
  };
  const parts = ['A1', '$B$2', 'C3:E9', 'D:D', '$F:$H', '5:7', 'Data!', "'Other'!", 'XFD1', 'XFA2:XFD2', 'SUM(', 'VLOOKUP(', ')', ',',
    '+', '"C3"', '#REF!', 'A1048576', 'nm', ' ', '[1]Data!B2', 'Data:Other!C1'];
  let seed = 11;
  const rnd = (n: number) => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed % n; };
  const randOp = (): StructOp => {
    const axis = rnd(4) === 0 ? 'row' : 'col';
    const k = rnd(3);
    if (k === 0) return { type: 'insert', axis, at: 1 + rnd(12), count: 1 + rnd(3) };
    if (k === 1) return { type: 'delete', axis, at: 1 + rnd(12), count: 1 + rnd(3) };
    const from = 1 + rnd(10), count = 1 + rnd(3);
    let to = 1 + rnd(14);
    if (to >= from && to <= from + count) to = from + count + 1 + rnd(3);
    return { type: 'move', axis, from, count, to };
  };
  it('tokenizes once but equals one adjustFormula per op (5,000 random formulas x random op sequences)', () => {
    for (let k = 0; k < 5000; k++) {
      let f = '';
      for (let j = 0, n = 1 + rnd(8); j < n; j++) f += parts[rnd(parts.length)];
      const ops = Array.from({ length: 1 + rnd(4) }, randOp);
      for (const ctx of [{ hostSheet: 'Data', targetSheet: 'Data' }, { hostSheet: 'Other', targetSheet: 'Data' }, { hostSheet: null, targetSheet: 'Data', isName: true }]) {
        expect(adjustFormulaSeq(f, ops, ctx)).toEqual(naive(f, ops, ctx));
      }
    }
  });
  it('planMove sequences give the intended permutation of every column', () => {
    const ops = planMove([2, 5, 6, 9], 12);
    const f = Array.from({ length: 14 }, (_, i) => `${String.fromCharCode(65 + i)}1`).join('+');
    const got = adjustFormulaSeq(f, ops, { hostSheet: 'Other', targetSheet: 'Other' }).formula;
    // B, E, F, I land before L (in that order): A C D G H J K B E F I L M N
    expect(got).toBe('A1+H1+B1+C1+I1+J1+D1+E1+K1+F1+G1+L1+M1+N1');
  });
  it('the table layout follows each op: a later delete breaks the right structured refs', () => {
    const ctx: AdjustContext = { hostSheet: 'Other', targetSheet: 'Data', tables: T };
    // delete D (大学) then D again (now 点数): 大学 and 点数 break, メール survives
    const ops: StructOp[] = [{ type: 'delete', axis: 'col', at: 4, count: 1 }, { type: 'delete', axis: 'col', at: 4, count: 1 }];
    expect(adjustFormulaSeq('tblMain[大学]&tblMain[点数]&tblMain[メール]', ops, ctx).formula).toBe('#REF!&#REF!&tblMain[メール]');
    // move 点数 (E) to the front of the table (before C), then delete C: 点数 breaks, 氏名 survives
    const ops2: StructOp[] = [{ type: 'move', axis: 'col', from: 5, count: 1, to: 3 }, { type: 'delete', axis: 'col', at: 3, count: 1 }];
    expect(adjustFormulaSeq('tblMain[氏名]&tblMain[点数]', ops2, ctx).formula).toBe('tblMain[氏名]&#REF!');
    // insert inside the table, then delete the inserted column: nothing breaks
    const ops3: StructOp[] = [{ type: 'insert', axis: 'col', at: 5, count: 1 }, { type: 'delete', axis: 'col', at: 5, count: 1 }];
    expect(adjustFormulaSeq('tblMain[大学]&tblMain[点数]', ops3, ctx).formula).toBe('tblMain[大学]&tblMain[点数]');
    // [@col] of a calculated column follows its host cell: H4 (合計) moves to G4 after deleting C
    const host: AdjustContext = { hostSheet: 'Data', targetSheet: 'Data', tables: T, hostCell: { row: 4, col: 8 } };
    const ops4: StructOp[] = [{ type: 'delete', axis: 'col', at: 3, count: 1 }, { type: 'delete', axis: 'col', at: 4, count: 1 }];
    expect(adjustFormulaSeq('[@点数]*2+[@大学]', ops4, host).formula).toBe('#REF!*2+[@大学]');
  });
  it('advanceContext', () => {
    const ctx: AdjustContext = { hostSheet: 'Data', targetSheet: 'Data', tables: T, hostCell: { row: 5, col: 8 } };
    const a = advanceContext(ctx, { type: 'move', axis: 'col', from: 5, count: 1, to: 3 });
    expect(a.tables![0].ref).toBe('C3:H10');
    expect(a.tables![0].columns).toEqual(['点数', '氏名', '大学', 'メール', '備考 メモ', '合計']);
    expect(a.hostCell).toEqual({ row: 5, col: 8 });
    const b = advanceContext(ctx, { type: 'insert', axis: 'col', at: 4, count: 2 });
    expect(b.tables![0].ref).toBe('C3:J10');
    expect(b.tables![0].columns.filter((c) => !c.startsWith(' '))).toEqual(T[0].columns);
    expect(b.tables![0].columns[1].startsWith(' ')).toBe(true);
    expect(b.hostCell).toEqual({ row: 5, col: 10 });
    const c = advanceContext(ctx, { type: 'delete', axis: 'col', at: 3, count: 6 });
    expect(c.tables).toEqual([]);
    expect(c.hostCell).toBeUndefined();
    const d = advanceContext(ctx, { type: 'delete', axis: 'row', at: 1, count: 1 });
    expect(d.tables![0]).toEqual({ ...T[0], ref: 'C2:H9' });
    expect(d.hostCell).toEqual({ row: 4, col: 8 });
    const other: AdjustContext = { hostSheet: 'Other', targetSheet: 'Other', tables: T };
    expect(advanceContext(other, { type: 'delete', axis: 'col', at: 1, count: 1 })).toBe(other);
  });
});
