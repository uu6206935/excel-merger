import { describe, expect, it } from 'vitest';
import { createEditor } from '../../src/core/doc';
import { F, XF, ext, fAt, ids, makeBook, makeEditor, roster, rosterSpec } from './helpers';

describe('createEditor', () => {
  it('builds a column-major doc with stable ids from the premerge layout', () => {
    const s = roster();
    const d = s.doc;
    expect(ids(d)).toEqual(['b1', 'b2', 'b3', 'b4', 'b5', 'i:s2:3', 'i:s2:4', 'b6']);
    expect(d.columns[5].origin).toEqual({ kind: 'import', sourceId: 's2', col: 3 });
    expect(d.columns[7].origin).toEqual({ kind: 'base', col: 6 });
    expect(d.columns[0].origin).toEqual({ kind: 'base', col: 1 });
    expect(s.log).toEqual([]);
    expect(d.sheetName).toBe('名簿');
    expect(d.sheetNames).toEqual(['名簿', '集計']);
    expect(d.headerRow).toBe(1);
    expect(d.firstDataRow).toBe(2);
    expect(d.lastDataRow).toBe(6);
    expect(d.rowCount).toBe(8);
    expect(d.columns.every((c) => c.cells.length === d.rowCount)).toBe(true);
    expect(fAt(d, 'E2')).toBe('D2*2');
    expect(d.columns[0].cells[1]).toEqual({ t: 's', v: '00123', s: XF.text });
    expect(d.columns[2].props.width).toBe(14);
    expect(d.sourceInfo.s2.label).toBe('成績');
    expect(d.version).toBeGreaterThan(0);
  });

  it('collects formulas of the other sheets, names, rules and styles', () => {
    const spec = rosterSpec();
    spec.sheets[1].features = { rules: [{ sqref: 'B2:B6', formulas: ['名簿!$D2>50'], kind: 'cf' }] };
    const d = makeEditor(spec).doc;
    expect(d.externalFormulas.map((e) => `${e.sheet}!${e.col},${e.row}=${e.f}`)).toEqual([
      '集計!2,2=SUM(名簿!D2:D6)',
      '集計!2,3=名簿!C2',
      '集計!2,4=VLOOKUP("00123",名簿!A2:E6,4,FALSE)',
      '集計!2,5=COUNTA(名簿!$B:$B)',
      '集計!2,6=SUM(点数範囲)',
    ]);
    // contract objects only (no private fields leaking into the doc)
    expect(Object.keys(d.externalFormulas[0]).sort()).toEqual(['col', 'f', 'row', 'sheet']);
    expect(d.externalRules).toEqual([{ sheet: '集計', rule: { sqref: 'B2:B6', formulas: ['名簿!$D2>50'], kind: 'cf' } }]);
    expect(d.names).toHaveLength(2);
    expect(d.styles[XF.int].numFmt).toBe('0');
    expect(ext(d, '集計', 'B3')).toBe('名簿!C2');
  });

  it('extends lastDataRow over appended rows and keeps sheet features', () => {
    const spec = rosterSpec();
    spec.sheets[0].cells!.A7 = '00999';
    spec.appended = [{ row: 7, sourceId: 's2', sourceRow: 12 }];
    spec.sheets[0].features = {
      merges: [{ r1: 8, c1: 1, r2: 8, c2: 3 }],
      freeze: { cols: 2, rows: 1 },
      autoFilter: 'A1:H6',
      tables: [],
      arrays: [],
      hyperlinks: [{ ref: 'G2', target: 'mailto:u0@x.jp' }],
      rules: [{ sqref: 'D2:D6', formulas: ['D2>80'], kind: 'cf' }],
    };
    const d = makeEditor(spec).doc;
    expect(d.lastDataRow).toBe(7);
    expect(d.features.freeze).toEqual({ cols: 2, rows: 1 });
    expect(d.features.autoFilter).toBe('A1:H6');
    expect(d.features.merges).toHaveLength(1);
    expect(d.features.hyperlinks[0].ref).toBe('G2');
  });

  it('maps base columns right of the imported block back to their original numbers', () => {
    // imported block in the middle of nowhere (after base column 2): C, D imported; E -> b3
    const s = makeEditor({
      sheets: [{ name: 'S', cells: { A1: 'a', B1: 'b', C1: 'x', D1: 'y', E1: 'c' } }],
      imported: [
        { col: 3, sourceId: 'src', sourceCol: 5, header: 'x' },
        { col: 4, sourceId: 'src', sourceCol: 9, header: 'y' },
      ],
    });
    expect(ids(s.doc)).toEqual(['b1', 'b2', 'i:src:5', 'i:src:9', 'b3']);
  });

  it('gives duplicate import ids a unique suffix', () => {
    const s = makeEditor({
      sheets: [{ name: 'S', cells: { A1: 'a', B1: 'x', C1: 'x2' } }],
      imported: [
        { col: 2, sourceId: 'src', sourceCol: 5, header: 'x' },
        { col: 3, sourceId: 'src', sourceCol: 5, header: 'x2' },
      ],
    });
    expect(new Set(ids(s.doc)).size).toBe(3);
  });

  it('throws a Japanese error when the merged sheet is missing', () => {
    const book = makeBook({ sheets: [{ name: 'A', cells: { A1: F('1', 1) } }] });
    expect(() =>
      createEditor(book, { sheet: 'ない', insertAt: 0, imported: [], appendAt: null, appended: [] }, {}, { headerRow: 1, firstDataRow: 2, lastDataRow: 2 }),
    ).toThrow(/シート/);
  });
});
