import { describe, expect, it } from 'vitest';
import { insertColumns } from '../../src/core/doc';
import { XF, at, ext, fAt, ids, ok, refused, roster, rowValues } from './helpers';

describe('insertColumns', () => {
  it('inserts blank columns with Excel "format from left" and adjusts references', () => {
    const s0 = roster();
    const r = insertColumns(s0, 2, 2); // before C
    const s = ok(r);
    const d = s.doc;
    expect(ids(d)).toEqual(['b1', 'b2', 'n1', 'n2', 'b3', 'b4', 'b5', 'i:s2:3', 'i:s2:4', 'b6']);
    expect(d.columns[2].origin).toEqual({ kind: 'new' });
    // left neighbour B: header style per row, width
    expect(at(d, 'C1')).toEqual({ t: 'z', v: null, s: XF.header });
    expect(at(d, 'D1')).toEqual({ t: 'z', v: null, s: XF.header });
    expect(at(d, 'C2')).toBeUndefined();
    expect(d.columns[2].props.width).toBe(12);
    expect(d.columns[3].cells.length).toBe(d.rowCount);
    // references
    expect(fAt(d, 'G2')).toBe('F2*2');
    expect(fAt(d, 'F8')).toBe('SUM(F2:F6)');
    expect(ext(d, '集計', 'B3')).toBe('名簿!E2');
    expect(ext(d, '集計', 'B4')).toBe('VLOOKUP("00123",名簿!A2:G6,4,FALSE)');
    expect(d.names.map((n) => n.refersTo)).toEqual(['名簿!$F$2:$F$6', '名簿!$E$2']);
    expect(s.log).toEqual([{ kind: 'struct', sheet: '名簿', op: { type: 'insert', axis: 'col', at: 3, count: 2 } }]);
    expect(s.label).toBe('2 列を挿入');
    // the lookup range A:E grew inside: VLOOKUP's column index now points elsewhere
    expect(r.ok && r.warnings.join('\n')).toMatch(/VLOOKUP/);
  });

  it('copies per-row cell formats of the left column', () => {
    const d = ok(insertColumns(roster(), 4, 1)).doc; // right of 点数 (D)
    expect(at(d, 'E1')?.s).toBe(XF.header);
    for (let r = 2; r <= 6; r++) expect(at(d, `E${r}`)).toEqual({ t: 'z', v: null, s: XF.int });
    expect(at(d, 'E7')).toBeUndefined();
    expect(at(d, 'E8')).toEqual({ t: 'z', v: null, s: XF.int });
    expect(d.columns[4].props.width).toBe(8);
  });

  it('inserts at column A without inherited formats', () => {
    const d = ok(insertColumns(roster(), 0, 1)).doc;
    expect(ids(d)[0]).toBe('n1');
    expect(d.columns[0].cells.every((c) => c === undefined)).toBe(true);
    expect(d.columns[0].props.width).toBeUndefined();
    expect(fAt(d, 'F2')).toBe('E2*2');
  });

  it('numbers new ids after the existing ones', () => {
    let s = ok(insertColumns(roster(), 1, 2));
    s = ok(insertColumns(s, 0, 1));
    expect(ids(s.doc).filter((i) => i.startsWith('n'))).toEqual(['n3', 'n1', 'n2']);
  });

  it('grows a table and writes the names Excel gives to new table columns', () => {
    const tbl = roster((sp) => {
      sp.sheets[0].features = {
        merges: [], rules: [], arrays: [], hyperlinks: [],
        tables: [{ name: '名簿T', ref: 'A1:E6', headerRow: true, totalsRow: false, columns: ['学生ID', '氏名', '大学', '点数', '倍'] }],
      };
    });
    const d = ok(insertColumns(tbl, 2, 1)).doc;
    expect(d.features.tables[0]).toMatchObject({ ref: 'A1:F6', columns: ['学生ID', '氏名', '列1', '大学', '点数', '倍'] });
    expect(at(d, 'C1')).toEqual({ t: 's', v: '列1', s: XF.header });
    // right after the table: not part of it
    expect(ok(insertColumns(tbl, 5, 1)).doc.features.tables[0].ref).toBe('A1:E6');
    // at the table's first column: the table shifts
    expect(ok(insertColumns(tbl, 0, 1)).doc.features.tables[0].ref).toBe('B1:F6');
    // numbered neighbour
    const t2 = roster((sp) => {
      sp.sheets[0].features = {
        merges: [], rules: [], arrays: [], hyperlinks: [],
        tables: [{ name: 'T', ref: 'A1:B6', headerRow: true, totalsRow: false, columns: ['項目1', '項目3'] }],
      };
    });
    expect(ok(insertColumns(t2, 1, 2)).doc.features.tables[0].columns).toEqual(['項目1', '項目2', '項目12', '項目3']);
  });

  it('grows merges and CF like Excel', () => {
    const s0 = roster((sp) => {
      sp.sheets[0].features = {
        merges: [{ r1: 8, c1: 1, r2: 8, c2: 3 }],
        rules: [{ sqref: 'D2:D6', formulas: ['D2>=80'], kind: 'cf' }],
        tables: [], arrays: [], hyperlinks: [],
        freeze: { cols: 2, rows: 1 },
      };
    });
    const d = ok(insertColumns(s0, 1, 1)).doc; // before B
    expect(d.features.merges).toEqual([{ r1: 8, c1: 1, r2: 8, c2: 4 }]);
    expect(d.features.rules).toEqual([{ sqref: 'E2:E6', formulas: ['E2>=80'], kind: 'cf' }]);
    expect(d.features.freeze).toEqual({ cols: 3, rows: 1 });
    // CF area ending right before the insertion point grows (format inheritance)
    const d2 = ok(insertColumns(s0, 4, 1)).doc; // right of D
    expect(d2.features.rules[0].sqref).toBe('D2:E6');
  });

  it('refuses inside a legacy array formula and bad input', () => {
    const arr = roster((sp) => {
      sp.sheets[0].cells!.F2 = { t: 's', v: '法', f: 'UPPER(B2:C2)', arr: 'F2:G2' };
      sp.sheets[0].features = { merges: [], rules: [], tables: [], hyperlinks: [], arrays: [{ r1: 2, c1: 6, r2: 2, c2: 7 }] };
    });
    expect(refused(insertColumns(arr, 6, 1))).toMatch(/配列/);
    const moved = ok(insertColumns(arr, 5, 1)).doc;
    expect(at(moved, 'G2')).toMatchObject({ f: 'UPPER(B2:C2)', arr: 'G2:H2' });
    expect(moved.features.arrays).toEqual([{ r1: 2, c1: 7, r2: 2, c2: 8 }]);
    const s0 = roster();
    expect(refused(insertColumns(s0, -1, 1))).toMatch(/位置/);
    expect(refused(insertColumns(s0, 9, 1))).toMatch(/位置/);
    expect(refused(insertColumns(s0, 0, 0))).toMatch(/列数/);
    expect(refused(insertColumns(s0, 0, 16384))).toMatch(/16,384/);
    expect(rowValues(s0.doc, 1)).toHaveLength(8);
  });
});
