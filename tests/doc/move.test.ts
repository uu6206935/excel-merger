import { describe, expect, it } from 'vitest';
import { planMove } from '../../src/core/formula/excelRef';
import { moveColumns } from '../../src/core/doc';
import { F, XF, at, ext, fAt, idAt, ids, ok, refused, roster, rowValues } from './helpers';

describe('moveColumns', () => {
  it('moves one column with Excel reference semantics everywhere', () => {
    const s0 = roster();
    const r = moveColumns(s0, ['b4'], 1); // 点数 (D) before B
    const s = ok(r);
    const d = s.doc;
    expect(ids(d)).toEqual(['b1', 'b4', 'b2', 'b3', 'b5', 'i:s2:3', 'i:s2:4', 'b6']);
    expect(rowValues(d, 1)).toEqual(['学生ID', '点数', '氏名', '大学', '倍', '学部', 'メール', '備考']);
    // merged sheet: E stays, its reference follows D -> B; the total travels with its column
    expect(fAt(d, 'E2')).toBe('B2*2');
    expect(fAt(d, 'B8')).toBe('SUM(B2:B6)');
    // other sheet
    expect(ext(d, '集計', 'B2')).toBe('SUM(名簿!B2:B6)');
    expect(ext(d, '集計', 'B3')).toBe('名簿!D2');
    expect(ext(d, '集計', 'B4')).toBe('VLOOKUP("00123",名簿!A2:E6,4,FALSE)'); // contains source and destination
    expect(ext(d, '集計', 'B5')).toBe('COUNTA(名簿!$C:$C)');
    // names
    expect(d.names.map((n) => n.refersTo)).toEqual(['名簿!$B$2:$B$6', '名簿!$D$2']);
    // widths and cell styles travel with the column
    expect(d.columns[1].props.width).toBe(8);
    expect(at(d, 'B2')).toEqual({ t: 'n', v: 80, s: XF.int });
    // log
    expect(s.log).toEqual([{ kind: 'struct', sheet: '名簿', op: { type: 'move', axis: 'col', from: 4, count: 1, to: 2 } }]);
    expect(s.label).toBe('列を移動');
    // VLOOKUP over A:E now sees another column order
    expect(r.ok && r.warnings.join('\n')).toMatch(/VLOOKUP/);
    // the input snapshot is untouched
    expect(ids(s0.doc)).toEqual(['b1', 'b2', 'b3', 'b4', 'b5', 'i:s2:3', 'i:s2:4', 'b6']);
    expect(fAt(s0.doc, 'E2')).toBe('D2*2');
    expect(d.version).toBeGreaterThan(s0.doc.version);
  });

  it('decomposes a non-contiguous selection with planMove and logs every block move', () => {
    const s0 = roster();
    const s = ok(moveColumns(s0, ['b6', 'b2', 'b4'], 0));
    expect(ids(s.doc)).toEqual(['b2', 'b4', 'b6', 'b1', 'b3', 'b5', 'i:s2:3', 'i:s2:4']);
    const want = planMove([2, 4, 8], 1, 'col');
    expect(want.length).toBe(3);
    expect(s.log).toEqual(want.map((op) => ({ kind: 'struct', sheet: '名簿', op })));
    // 倍 (b5) is now column F, its reference follows 点数 (now B)
    expect(fAt(s.doc, 'F2')).toBe('B2*2');
    expect(s.label).toBe('3 列を移動');
  });

  it('moves a contiguous block as one op', () => {
    const s = ok(moveColumns(roster(), ['i:s2:3', 'i:s2:4'], 1));
    expect(ids(s.doc)).toEqual(['b1', 'i:s2:3', 'i:s2:4', 'b2', 'b3', 'b4', 'b5', 'b6']);
    expect(s.log).toEqual([{ kind: 'struct', sheet: '名簿', op: { type: 'move', axis: 'col', from: 6, count: 2, to: 2 } }]);
    expect(fAt(s.doc, 'G2')).toBe('F2*2');
    expect(fAt(s.doc, 'F8')).toBe('SUM(F2:F6)');
  });

  it('returns the same snapshot for moves that change nothing', () => {
    const s0 = roster();
    for (const gap of [3, 4]) {
      const r = moveColumns(s0, ['b4'], gap);
      expect(r.ok && r.snapshot).toBe(s0);
    }
    // target inside the selection
    const r2 = moveColumns(s0, ['b2', 'b3', 'b4'], 2);
    expect(r2.ok && r2.snapshot).toBe(s0);
    const r3 = moveColumns(s0, [], 0);
    expect(r3.ok && r3.snapshot).toBe(s0);
  });

  it('moves to the end (gap = n)', () => {
    const s = ok(moveColumns(roster(), ['b1'], 8));
    expect(ids(s.doc)).toEqual(['b2', 'b3', 'b4', 'b5', 'i:s2:3', 'i:s2:4', 'b6', 'b1']);
    expect(s.log[0]).toEqual({ kind: 'struct', sheet: '名簿', op: { type: 'move', axis: 'col', from: 1, count: 1, to: 9 } });
    expect(rowValues(s.doc, 2)[7]).toBe('00123');
  });

  it('shares every unchanged column and cell with the previous snapshot', () => {
    const s0 = roster();
    const s = ok(moveColumns(s0, ['b4'], 1));
    const col = (doc: typeof s.doc, id: string) => doc.columns.find((c) => c.id === id)!;
    for (const id of ['b1', 'b2', 'b3', 'i:s2:3', 'i:s2:4', 'b6']) expect(col(s.doc, id)).toBe(col(s0.doc, id));
    // E changed (its formulas) but its unchanged header cell is shared
    expect(col(s.doc, 'b5')).not.toBe(col(s0.doc, 'b5'));
    expect(col(s.doc, 'b5').cells[0]).toBe(col(s0.doc, 'b5').cells[0]);
    // untouched features / styles are shared
    expect(s.doc.styles).toBe(s0.doc.styles);
    expect(s.doc.sourceInfo).toBe(s0.doc.sourceInfo);
  });

  it('adjusts freeze panes, autofilter, CF, DV and hyperlinks', () => {
    const s0 = roster((sp) => {
      sp.sheets[0].features = {
        freeze: { cols: 2, rows: 1 },
        autoFilter: 'A1:H6',
        rules: [
          { sqref: 'D2:D6', formulas: ['D2>=80'], kind: 'cf' },
          { sqref: 'H2:H6', formulas: ['$C$2:$C$6'], kind: 'dv' },
        ],
        hyperlinks: [{ ref: 'G2', target: 'mailto:u0@x.jp' }],
        merges: [],
        tables: [],
        arrays: [],
      };
    });
    const d = ok(moveColumns(s0, ['b4'], 0)).doc; // D before A
    expect(d.features.freeze).toEqual({ cols: 3, rows: 1 }); // R2b R10: only the insert step counts
    expect(d.features.autoFilter).toBe('B1:H6'); // R2b m12 (T before A: A1:T20 -> B1:T20)
    expect(d.features.rules).toEqual([
      { sqref: 'A2:A6', formulas: ['A2>=80'], kind: 'cf' },
      { sqref: 'H2:H6', formulas: ['$D$2:$D$6'], kind: 'dv' },
    ]);
    expect(d.features.hyperlinks).toEqual([{ ref: 'G2', target: 'mailto:u0@x.jp' }]);
  });

  it('refuses like Excel: part of a merged area, into / out of a table, part of an array', () => {
    const merged = roster((sp) => {
      sp.sheets[0].features = { merges: [{ r1: 8, c1: 1, r2: 8, c2: 3 }], rules: [], tables: [], arrays: [], hyperlinks: [] };
    });
    expect(refused(moveColumns(merged, ['b2'], 6))).toMatch(/結合/);
    expect(refused(moveColumns(merged, ['b6'], 2))).toMatch(/結合/); // into the middle of A8:C8
    const whole = ok(moveColumns(merged, ['b1', 'b2', 'b3'], 8));
    expect(whole.doc.features.merges).toEqual([{ r1: 8, c1: 6, r2: 8, c2: 8 }]);

    const table = roster((sp) => {
      sp.sheets[0].features = {
        merges: [], rules: [], arrays: [], hyperlinks: [],
        tables: [{ name: '名簿T', ref: 'A1:E6', headerRow: true, totalsRow: false, columns: ['学生ID', '氏名', '大学', '点数', '倍'] }],
      };
    });
    expect(refused(moveColumns(table, ['b6'], 2))).toMatch(/テーブル/); // into
    expect(refused(moveColumns(table, ['b2'], 8))).toMatch(/テーブル/); // out of
    expect(refused(moveColumns(table, ['b5', 'i:s2:3'], 0))).toMatch(/テーブル/); // partly over
    const inside = ok(moveColumns(table, ['b4'], 1));
    expect(inside.doc.features.tables[0]).toMatchObject({ ref: 'A1:E6', columns: ['学生ID', '点数', '氏名', '大学', '倍'] });
    const wholeTable = ok(moveColumns(table, ['b1', 'b2', 'b3', 'b4', 'b5'], 8));
    expect(wholeTable.doc.features.tables[0].ref).toBe('D1:H6');

    const arr = roster((sp) => {
      sp.sheets[0].cells!.F2 = { t: 's', v: '法', f: 'UPPER(B2:C2)', arr: 'F2:G2' };
      sp.sheets[0].features = { merges: [], rules: [], tables: [], hyperlinks: [], arrays: [{ r1: 2, c1: 6, r2: 2, c2: 7 }] };
    });
    const why = refused(moveColumns(arr, ['i:s2:3'], 0));
    expect(why).toMatch(/配列/);
    const both = ok(moveColumns(arr, ['i:s2:3', 'i:s2:4'], 0));
    expect(both.doc.features.arrays).toEqual([{ r1: 2, c1: 1, r2: 2, c2: 2 }]);
    expect(at(both.doc, 'A2')).toMatchObject({ f: 'UPPER(D2:E2)', arr: 'A2:B2' });
  });

  it('refuses bad input with a Japanese reason', () => {
    const s0 = roster();
    expect(refused(moveColumns(s0, ['nope'], 0))).toMatch(/列/);
    expect(refused(moveColumns(s0, ['b1'], -1))).toMatch(/位置/);
    expect(refused(moveColumns(s0, ['b1'], 99))).toMatch(/位置/);
    expect(refused(moveColumns(s0, ['b1'], 1.5))).toMatch(/位置/);
  });

  it('warns only when a column-number lookup would return another column', () => {
    const r = moveColumns(roster(), ['b6'], 6); // H between E and F: VLOOKUP A:E untouched
    expect(r.ok && r.warnings).toEqual([]);
    const r2 = moveColumns(roster(), ['b3'], 0);
    expect(r2.ok && r2.warnings[0]).toMatch(/列番号/);
    expect(r2.ok && r2.warnings.join('\n')).toContain('集計!B4');
    expect(r2.ok && r2.warnings.join('\n')).toMatch(/VLOOKUP の列番号/);
    // B moves inside A:E but column 4 of the range is still 点数: exact check, no warning
    const r3 = moveColumns(roster(), ['b2'], 3);
    expect(ok(r3).doc.columns.map((c) => c.id).slice(0, 5)).toEqual(['b1', 'b3', 'b2', 'b4', 'b5']);
    expect(r3.ok && r3.warnings).toEqual([]);
    // many filled-down lookups: counted once each, a few examples listed
    const many = roster((sp) => {
      for (let r = 2; r <= 6; r++) sp.sheets[0].cells![`H${r}`] = F(`INDEX($A$2:$E$6,${r - 1},4)`, 0);
    });
    const r4 = moveColumns(many, ['b4'], 0);
    expect(r4.ok && r4.warnings[0]).toMatch(/6 個/); // 5 INDEX + 1 VLOOKUP
  });

  it('keeps ids stable over many moves (round trip)', () => {
    let s = roster();
    const start = s.doc;
    s = ok(moveColumns(s, ['b1', 'b3'], 8));
    s = ok(moveColumns(s, ['b6'], 0));
    s = ok(moveColumns(s, ['b1'], 1));
    // put everything back in the original order
    const order = start.columns.map((c) => c.id);
    for (let i = 0; i < order.length; i++) s = ok(moveColumns(s, [order[i]], i));
    expect(ids(s.doc)).toEqual(order);
    expect(fAt(s.doc, 'E2')).toBe('D2*2');
    expect(ext(s.doc, '集計', 'B2')).toBe('SUM(名簿!D2:D6)');
    expect(idAt(s.doc, 'D')).toBe('b4');
  });
});
