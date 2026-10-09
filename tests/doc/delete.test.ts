import { describe, expect, it } from 'vitest';
import { analyzeDelete, deleteColumns } from '../../src/core/doc';
import type { EditorSnapshot, LoggedOp } from '../../src/core/types';
import { F, XF, allFormulas, at, ext, fAt, ids, ok, refused, roster } from './helpers';

/** roster + a formula that uses the name 大学セル */
const base = () =>
  roster((sp) => {
    sp.sheets[1].cells!.B7 = F('大学セル&"様"', '東大様');
  });

const noRefOutsideNames = (s: EditorSnapshot) => {
  const d = s.doc;
  const cellsAndRules = allFormulas({ ...d, names: [] });
  expect(cellsAndRules.filter((f) => f.includes('#REF!'))).toEqual([]);
};

describe('analyzeDelete', () => {
  it('lists every formula / name that would become #REF! with its location', () => {
    const s = base();
    const imp = analyzeDelete(s, ['b3']); // 大学 (C)
    expect(imp.breakCount).toBe(3);
    expect(imp.breaks).toEqual(
      expect.arrayContaining([
        { where: '集計!B3', text: '=名簿!C2' },
        { where: '集計!B7', text: '=大学セル&"様"' },
        { where: '名前「大学セル」', text: '=名簿!$C$2（値に変換できないため #REF! になります）' },
      ]),
    );
    expect(imp.lookupWarnings.join('\n')).toMatch(/VLOOKUP|列番号/);
    expect(imp.lookupWarnings.join('\n')).toContain('集計!B4');
  });

  it('reports merged-sheet formulas and ignores formulas inside the deleted columns', () => {
    const imp = analyzeDelete(base(), ['b4']); // 点数 (D): E2:E6 and 集計 break; D8 disappears with D
    const where = imp.breaks.map((b) => b.where);
    expect(where).toEqual(expect.arrayContaining(['名簿!E2', '名簿!E6', '集計!B2', '集計!B6', '名前「点数範囲」']));
    expect(where).not.toContain('名簿!D8');
    expect(imp.breakCount).toBe(8); // E2..E6, 集計!B2, 集計!B6 (uses the name), name 点数範囲
    expect(imp.breaks.find((b) => b.where === '名簿!E2')!.text).toBe('=D2*2');
  });

  it('is empty for columns nothing depends on', () => {
    expect(analyzeDelete(base(), ['b6'])).toEqual({ breaks: [], breakCount: 0, lookupWarnings: [] });
    expect(analyzeDelete(base(), ['nope']).breakCount).toBe(0);
  });
});

describe('deleteColumns', () => {
  it('without conversion behaves like Excel (#REF!) and says so', () => {
    const s0 = base();
    const r = deleteColumns(s0, ['b3'], { convertDependents: false });
    const s = ok(r);
    expect(ids(s.doc)).toEqual(['b1', 'b2', 'b4', 'b5', 'i:s2:3', 'i:s2:4', 'b6']);
    expect(ext(s.doc, '集計', 'B3')).toBe('名簿!#REF!');
    expect(ext(s.doc, '集計', 'B4')).toBe('VLOOKUP("00123",名簿!A2:D6,4,FALSE)');
    expect(s.doc.names.find((n) => n.name === '大学セル')!.refersTo).toBe('名簿!#REF!');
    expect(fAt(s.doc, 'D2')).toBe('C2*2');
    expect(s.log).toEqual([{ kind: 'struct', sheet: '名簿', op: { type: 'delete', axis: 'col', at: 3, count: 1 } }]);
    expect(r.ok && r.warnings.join('\n')).toMatch(/#REF!/);
    expect(s.label).toBe('列を削除');
  });

  it('converts dependents to their cached values BEFORE deleting: no #REF! anywhere', () => {
    const s0 = base();
    const r = deleteColumns(s0, ['b3'], { convertDependents: true });
    const s = ok(r);
    const cellsLog: LoggedOp = {
      kind: 'cells',
      sheet: '集計',
      reason: 'toValues',
      writes: [
        { row: 3, col: 2, cell: { t: 's', v: '東大' } },
        { row: 7, col: 2, cell: { t: 's', v: '東大様' } },
      ],
    };
    expect(s.log).toEqual([cellsLog, { kind: 'struct', sheet: '名簿', op: { type: 'delete', axis: 'col', at: 3, count: 1 } }]);
    expect(ext(s.doc, '集計', 'B3')).toBeUndefined();
    expect(ext(s.doc, '集計', 'B7')).toBeUndefined();
    noRefOutsideNames(s);
    // names cannot be converted: listed in the warnings
    expect(r.ok && r.warnings.join('\n')).toMatch(/大学セル/);
    expect(r.ok && r.warnings[0]).toMatch(/2 個を値に変換/);
  });

  it('converts merged-sheet formulas in place (values keep their style)', () => {
    const s = ok(deleteColumns(base(), ['b4'], { convertDependents: true }));
    const d = s.doc;
    expect(ids(d)).toEqual(['b1', 'b2', 'b3', 'b5', 'i:s2:3', 'i:s2:4', 'b6']);
    expect(at(d, 'D2')).toEqual({ t: 'n', v: 160, s: XF.int });
    expect(at(d, 'D6')).toEqual({ t: 'n', v: 100, s: XF.int });
    noRefOutsideNames(s);
    const first = s.log[0];
    expect(first.kind === 'cells' && first.sheet).toBe('名簿');
    expect(first.kind === 'cells' && first.writes.map((w) => `${w.col}:${w.row}=${w.cell?.v}`)).toEqual(['5:2=160', '5:3=140', '5:4=180', '5:5=120', '5:6=100']);
    const kinds = s.log.map((l) => (l.kind === 'cells' ? `cells:${l.sheet}` : l.kind));
    expect(kinds).toEqual(['cells:名簿', 'cells:集計', 'struct']);
    // 集計!B6 used the broken name 点数範囲: converted too
    expect(ext(d, '集計', 'B6')).toBeUndefined();
    expect(ext(d, '集計', 'B2')).toBeUndefined();
    expect(ext(d, '集計', 'B5')).toBe('COUNTA(名簿!$B:$B)');
  });

  it('deletes non-contiguous selections as contiguous runs, right to left', () => {
    const s = ok(deleteColumns(base(), ['b6', 'b2', 'i:s2:3', 'b3'], { convertDependents: true }));
    const structs = s.log.filter((l) => l.kind === 'struct').map((l) => (l.kind === 'struct' ? l.op : null));
    expect(structs).toEqual([
      { type: 'delete', axis: 'col', at: 8, count: 1 },
      { type: 'delete', axis: 'col', at: 6, count: 1 },
      { type: 'delete', axis: 'col', at: 2, count: 2 },
    ]);
    expect(ids(s.doc)).toEqual(['b1', 'b4', 'b5', 'i:s2:4']);
    expect(fAt(s.doc, 'C2')).toBe('B2*2');
    expect(s.label).toBe('4 列を削除');
    noRefOutsideNames(s);
  });

  it('deletes table columns: structured references to them break (and can be converted)', () => {
    const tbl = () =>
      roster((sp) => {
        sp.sheets[0].features = {
          merges: [], rules: [], arrays: [], hyperlinks: [],
          tables: [{ name: '名簿T', ref: 'A1:E6', headerRow: true, totalsRow: false, columns: ['学生ID', '氏名', '大学', '点数', '倍'] }],
        };
        for (let r = 2; r <= 6; r++) sp.sheets[0].cells![`E${r}`] = F('名簿T[[#This Row],[点数]]*2', (r - 1) * 10, { s: XF.int });
        sp.sheets[1].cells!.B8 = F('SUM(名簿T[点数])', 350);
        sp.sheets[1].cells!.B9 = F('COUNTA(名簿T[氏名])', 5);
      });
    const imp = analyzeDelete(tbl(), ['b4']);
    expect(imp.breaks.map((b) => b.where)).toEqual(expect.arrayContaining(['名簿!E2', '集計!B8']));
    expect(imp.breaks.find((b) => b.where === '集計!B8')!.text).toBe('=SUM(名簿T[点数])');
    const plain = ok(deleteColumns(tbl(), ['b4'], { convertDependents: false })).doc;
    expect(ext(plain, '集計', 'B8')).toBe('SUM(#REF!)');
    expect(fAt(plain, 'D2')).toBe('#REF!*2');
    expect(ext(plain, '集計', 'B9')).toBe('COUNTA(名簿T[氏名])');
    expect(plain.features.tables[0]).toMatchObject({ ref: 'A1:D6', columns: ['学生ID', '氏名', '大学', '倍'] });
    const conv = ok(deleteColumns(tbl(), ['b4'], { convertDependents: true }));
    noRefOutsideNames(conv);
    expect(at(conv.doc, 'D2')).toEqual({ t: 'n', v: 10, s: XF.int });
    expect(ext(conv.doc, '集計', 'B9')).toBe('COUNTA(名簿T[氏名])');
  });

  it('clears formulas that have no cached value and says so', () => {
    const s0 = roster((sp) => {
      sp.sheets[1].cells!.C3 = { t: 'z', v: null, f: '名簿!C3&"!"' };
    });
    const r = deleteColumns(s0, ['b3'], { convertDependents: true });
    const s = ok(r);
    const w = s.log[0];
    expect(w.kind === 'cells' && w.writes).toEqual([
      { row: 3, col: 2, cell: { t: 's', v: '東大' } },
      { row: 3, col: 3, cell: null },
    ]);
    expect(r.ok && r.warnings.join('\n')).toMatch(/空欄/);
  });

  it('refuses what cannot be done safely', () => {
    const s0 = base();
    expect(refused(deleteColumns(s0, ids(s0.doc), { convertDependents: true }))).toMatch(/すべての列/);
    expect(refused(deleteColumns(s0, ['nope'], { convertDependents: true }))).toMatch(/列/);
    // an array formula (other sheet) that depends on the column is converted as a whole (the master write replaces
    // the whole array), or becomes #REF! without conversion
    const arr = roster((sp) => {
      sp.sheets[1].cells!.D2 = F('名簿!C2:C3', '東大', { arr: 'D2:D3' });
    });
    const conv = deleteColumns(arr, ['b3'], { convertDependents: true });
    const cs = ok(conv);
    expect(cs.doc.externalFormulas.find((e) => e.col === 4)).toBeUndefined();
    expect(cs.log[0]).toEqual({
      kind: 'cells', sheet: '集計', reason: 'toValues',
      writes: [{ row: 3, col: 2, cell: { t: 's', v: '東大' } }, { row: 2, col: 4, cell: { t: 's', v: '東大' } }],
    });
    expect(conv.ok && conv.warnings.join()).toMatch(/配列数式 1 個/);
    expect(ok(deleteColumns(arr, ['b3'], { convertDependents: false })).doc.externalFormulas.find((e) => e.col === 4)!.f).toBe('名簿!#REF!');
    // part of a legacy array on the merged sheet (Excel: 配列の一部を変更することはできません)
    const arr2 = roster((sp) => {
      sp.sheets[0].cells!.F2 = { t: 's', v: '法', f: 'UPPER(B2:C2)', arr: 'F2:G2' };
      sp.sheets[0].features = { merges: [], rules: [], tables: [], hyperlinks: [], arrays: [{ r1: 2, c1: 6, r2: 2, c2: 7 }] };
    });
    expect(refused(deleteColumns(arr2, ['i:s2:4'], { convertDependents: false }))).toMatch(/配列/);
    // ... still refused when converting, as long as the array formula itself does not depend on the deleted column
    expect(refused(deleteColumns(arr2, ['i:s2:4'], { convertDependents: true }))).toMatch(/配列/);
    // an empty selection is a no-op
    const same = deleteColumns(s0, [], { convertDependents: true });
    expect(same.ok && same.snapshot).toBe(s0);
  });

  it('shrinks merges, removes fully deleted rules and keeps freeze consistent', () => {
    const s0 = roster((sp) => {
      sp.sheets[0].features = {
        merges: [{ r1: 8, c1: 1, r2: 8, c2: 3 }],
        rules: [{ sqref: 'C2:C6', formulas: ['C2="東大"'], kind: 'cf' }, { sqref: 'B2:D6', formulas: ['$D2>60'], kind: 'cf' }],
        tables: [], arrays: [], hyperlinks: [],
        freeze: { cols: 3, rows: 1 },
        autoFilter: 'A1:H6',
      };
    });
    const d = ok(deleteColumns(s0, ['b3'], { convertDependents: false })).doc;
    expect(d.features.merges).toEqual([{ r1: 8, c1: 1, r2: 8, c2: 2 }]);
    expect(d.features.rules).toEqual([{ sqref: 'B2:C6', formulas: ['$C2>60'], kind: 'cf' }]);
    expect(d.features.freeze).toEqual({ cols: 2, rows: 1 });
    expect(d.features.autoFilter).toBe('A1:G6');
  });
});
