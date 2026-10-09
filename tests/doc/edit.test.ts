import { describe, expect, it } from 'vitest';
import { formulaForDisplay, moveColumns, setCellInput, setColumnWidth } from '../../src/core/doc';
import { F, XF, at, ext, fAt, ok, refused, roster } from './helpers';

const tableRoster = () =>
  roster((sp) => {
    sp.sheets[0].features = {
      merges: [], rules: [], arrays: [], hyperlinks: [],
      tables: [{ name: '名簿T', ref: 'A1:E6', headerRow: true, totalsRow: false, columns: ['学生ID', '氏名', '大学', '点数', '倍'] }],
    };
    for (let r = 2; r <= 6; r++) sp.sheets[0].cells![`E${r}`] = F('名簿T[[#This Row],[点数]]*2', 0, { s: XF.int });
    sp.sheets[1].cells!.B8 = F('SUM(名簿T[点数])', 350);
    sp.sheets[1].cells!.B9 = F('名簿T[[#Headers],[点数]]', '点数');
    sp.names!.push({ name: '点数列', scope: null, refersTo: '名簿T[点数]' });
  });

describe('setCellInput', () => {
  it('writes a parsed value keeping the style and logs the write', () => {
    const s0 = roster();
    const s = ok(setCellInput(s0, 'b4', 3, '８５'));
    expect(at(s.doc, 'D3')).toEqual({ t: 'n', v: 85, s: XF.int });
    expect(s.log).toEqual([{ kind: 'cells', sheet: '名簿', writes: [{ row: 3, col: 4, cell: { t: 'n', v: 85 } }], reason: 'edit' }]);
    expect(s.label).toBe('セルを編集');
    expect(at(s0.doc, 'D3')).toEqual({ t: 'n', v: 70, s: XF.int });
    // only the edited column is copied
    expect(s.doc.columns[2]).toBe(s0.doc.columns[2]);
    expect(s.doc.version).toBeGreaterThan(s0.doc.version);
  });

  it('uses the CURRENT position of the column in the log', () => {
    const moved = ok(moveColumns(roster(), ['b4'], 0));
    const s = ok(setCellInput(moved, 'b4', 2, '99'));
    expect(s.log[s.log.length - 1]).toEqual({ kind: 'cells', sheet: '名簿', writes: [{ row: 2, col: 1, cell: { t: 'n', v: 99 } }], reason: 'edit' });
  });

  it('writes formulas in FILE form without a cached value; displays them in UI form', () => {
    const s = ok(setCellInput(roster(), 'b6', 2, '=xlookup(a2,a2:a6,b2:b6)&"さん"'));
    expect(at(s.doc, 'H2')).toEqual({ t: 'z', v: null, f: '_xlfn.XLOOKUP(A2,A2:A6,B2:B6)&"さん"' });
    expect(s.log[0]).toEqual({ kind: 'cells', sheet: '名簿', writes: [{ row: 2, col: 8, cell: { t: 'z', v: null, f: '_xlfn.XLOOKUP(A2,A2:A6,B2:B6)&"さん"' } }], reason: 'edit' });
    expect(formulaForDisplay(s.doc, 7, 2)).toBe('=XLOOKUP(A2,A2:A6,B2:B6)&"さん"');
    expect(formulaForDisplay(s.doc, 7, 3)).toBeNull();
    // a typed formula follows later moves like any other
    const m = ok(moveColumns(s, ['b1'], 8));
    expect(fAt(m.doc, 'G2')).toBe('_xlfn.XLOOKUP(H2,H2:H6,A2:A6)&"さん"');
  });

  it('clears values but keeps the cell format; no-ops return the same snapshot', () => {
    const s0 = roster();
    const s = ok(setCellInput(s0, 'b4', 3, ''));
    expect(at(s.doc, 'D3')).toEqual({ t: 'z', v: null, s: XF.int });
    expect(s.log[0]).toEqual({ kind: 'cells', sheet: '名簿', writes: [{ row: 3, col: 4, cell: null }], reason: 'edit' });
    for (const [id, row, text] of [['b6', 3, ''], ['b4', 3, '70'], ['b2', 2, '佐藤']] as const) {
      const r = setCellInput(s0, id, row, text);
      expect(r.ok && r.snapshot, `${id}${row}`).toBe(s0);
    }
    // replacing a formula by its own value is a change
    const e = ok(setCellInput(s0, 'b5', 2, '160'));
    expect(at(e.doc, 'E2')).toEqual({ t: 'n', v: 160, s: XF.int });
  });

  it('keeps text IDs as text in an ID column and in text-formatted cells', () => {
    const s = ok(setCellInput(roster(), 'b1', 4, '00200'));
    expect(at(s.doc, 'A4')).toEqual({ t: 's', v: '00200', s: XF.text });
    const s2 = ok(setCellInput(roster((sp) => {
      for (let r = 2; r <= 6; r++) (sp.sheets[0].cells![`A${r}`] as { s?: number }).s = undefined;
    }), 'b1', 4, '200'));
    expect(at(s2.doc, 'A4')).toMatchObject({ t: 's', v: '200' });
    // a plain number column
    const s3 = ok(setCellInput(roster(), 'b2', 3, '200'));
    expect(at(s3.doc, 'B3')).toEqual({ t: 'n', v: 200 });
  });

  it('dates become serials in date-formatted cells; in General cells they also set the number format like Excel', () => {
    const s0 = roster((sp) => {
      sp.sheets[0].cells!.H2 = { t: 'z', v: null, s: XF.date };
    });
    const d = ok(setCellInput(s0, 'b6', 2, '令和8年10月20日'));
    expect(at(d.doc, 'H2')).toEqual({ t: 'd', v: 46315, s: XF.date });
    expect(d.log[0]).toEqual({ kind: 'cells', sheet: '名簿', writes: [{ row: 2, col: 8, cell: { t: 'd', v: 46315 } }], reason: 'edit' });
    // General cell: serial + the format Excel-ja gives it (a derived display style)
    const g = ok(setCellInput(s0, 'b6', 3, '2026/10/20'));
    const h3 = at(g.doc, 'H3')!;
    expect(h3).toMatchObject({ t: 'd', v: 46315 });
    expect(g.doc.styles[h3.s!]).toEqual({ numFmt: 'yyyy/m/d' });
    expect(h3.s).toBe(XF.date); // an identical existing style is reused
    expect(g.log[0]).toEqual({ kind: 'cells', sheet: '名簿', writes: [{ row: 3, col: 8, cell: { t: 'd', v: 46315, numFmt: 'yyyy/m/d' } }], reason: 'edit' });
    // a percentage into a styled General cell: the style is kept, only the number format changes
    const p = ok(setCellInput(s0, 'b2', 1, '50%'));
    const b1 = at(p.doc, 'B1')!;
    expect(b1).toMatchObject({ t: 'n', v: 0.5 });
    expect(p.doc.styles[b1.s!]).toEqual({ bold: true, fill: '#DDEBF7', numFmt: '0%' });
    expect(p.doc.styles.length).toBe(s0.doc.styles.length + 1);
    expect(s0.doc.styles.length).toBe(6); // the input snapshot is untouched
    expect(p.log[0]).toMatchObject({ writes: [{ row: 1, col: 2, cell: { t: 'n', v: 0.5, numFmt: '0%' } }] });
    // the same value typed with a format change is not a no-op
    const q = ok(setCellInput(ok(setCellInput(s0, 'b6', 4, '0.5')), 'b6', 4, '50%'));
    expect(q.doc.styles[at(q.doc, 'H4')!.s!]).toEqual({ numFmt: '0%' });
    // yen / thousands / m/d
    const y = ok(setCellInput(s0, 'b6', 5, '¥1,000'));
    expect(y.doc.styles[at(y.doc, 'H5')!.s!]).toEqual({ numFmt: '"¥"#,##0;[Red]"¥"\\-#,##0' });
    const k = ok(setCellInput(s0, 'b6', 5, '1,000'));
    expect(k.log[0]).toMatchObject({ writes: [{ cell: { t: 'n', v: 1000, numFmt: '#,##0' } }] });
  });

  it('a new cell takes the row format (row with customFormat) like Excel and the patcher', () => {
    const s0 = roster();
    const doc = { ...s0.doc, rowProps: [undefined, undefined, undefined, { s: XF.header }] };
    const s = ok(setCellInput({ ...s0, doc }, 'b6', 4, 'メモ'));
    expect(at(s.doc, 'H4')).toEqual({ t: 's', v: 'メモ', s: XF.header });
  });

  it('refuses like Excel: inside array formulas, hidden cells of merged areas, broken formulas', () => {
    const s0 = roster((sp) => {
      sp.sheets[0].cells!.F2 = { t: 's', v: '法', f: 'UPPER(B2:C2)', arr: 'F2:G2' };
      sp.sheets[0].features = { merges: [{ r1: 8, c1: 1, r2: 8, c2: 3 }], rules: [], tables: [], hyperlinks: [], arrays: [{ r1: 2, c1: 6, r2: 2, c2: 7 }] };
    });
    expect(refused(setCellInput(s0, 'i:s2:4', 2, 'x'))).toMatch(/配列/);
    expect(refused(setCellInput(s0, 'i:s2:3', 2, 'x'))).toMatch(/配列/);
    expect(refused(setCellInput(s0, 'b2', 8, 'x'))).toMatch(/結合/);
    ok(setCellInput(s0, 'b1', 8, '合計'));
    expect(refused(setCellInput(s0, 'b6', 2, '=SUM(A1'))).toMatch(/かっこ/);
    expect(refused(setCellInput(s0, 'b6', 2, '=存在しない!A1'))).toMatch(/存在しない/);
    expect(refused(setCellInput(s0, 'nope', 2, 'x'))).toMatch(/列/);
    expect(refused(setCellInput(s0, 'b6', 0, 'x'))).toMatch(/位置/);
    expect(refused(setCellInput(s0, 'b6', 99, 'x'))).toMatch(/位置/);
    expect(formulaForDisplay(s0.doc, 5, 2)).toBe('{=UPPER(B2:C2)}');
  });

  it('allows re-entering a dynamic-array formula in its own cell but not typing into its spill range', () => {
    const s0 = roster((sp) => {
      sp.sheets[0].cells!.H2 = { t: 's', v: '九大', f: '_xlfn._xlws.SORT(C2:C6)', arr: 'H2:H6', dyn: true };
      for (let r = 3; r <= 6; r++) sp.sheets[0].cells![`H${r}`] = 'x';
      sp.sheets[0].features = { merges: [], rules: [], tables: [], hyperlinks: [], arrays: [{ r1: 2, c1: 8, r2: 6, c2: 8 }] };
    });
    expect(refused(setCellInput(s0, 'b6', 4, 'y'))).toMatch(/スピル/);
    expect(formulaForDisplay(s0.doc, 7, 2)).toBe('=SORT(C2:C6)');
    const s = ok(setCellInput(s0, 'b6', 2, '=unique(b2:b6)'));
    // UNIQUE returns an array: stored as a dynamic-array master (Excel computes the spill range on load)
    expect(at(s.doc, 'H2')).toEqual({ t: 'z', v: null, f: '_xlfn.UNIQUE(B2:B6)', dyn: true, arr: 'H2' });
    expect(s.log[0]).toEqual({ kind: 'cells', sheet: '名簿', writes: [{ row: 2, col: 8, cell: { t: 'z', v: null, f: '_xlfn.UNIQUE(B2:B6)', dyn: true } }], reason: 'edit' });
    expect(s.doc.features.arrays).toEqual([{ r1: 2, c1: 8, r2: 2, c2: 8 }]);
    // the old spilled values disappear (like the patcher / Excel)
    for (let r = 3; r <= 6; r++) expect(at(s.doc, `H${r}`), `H${r}`).toBeUndefined();
    // a scalar formula over the master is a plain formula again
    const t = ok(setCellInput(s, 'b6', 2, '=B2&"さん"'));
    expect(at(t.doc, 'H2')).toEqual({ t: 'z', v: null, f: 'B2&"さん"' });
    expect(t.doc.features.arrays).toEqual([]);
  });

  it('flags typed formulas that Excel 365 stores as dynamic arrays', () => {
    const s0 = roster((sp) => {
      sp.names!.push({ name: '単価', scope: null, refersTo: '名簿!$D$2' }, { name: '列', scope: '名簿', refersTo: '名簿!$D$2:$D$6' });
    });
    const dyn = (text: string) => {
      const s = ok(setCellInput(s0, 'b6', 3, text));
      const w = s.log[0];
      return w.kind === 'cells' ? !!w.writes[0].cell?.dyn : null;
    };
    expect(dyn('=SUM(D2:D6)')).toBe(false);
    expect(dyn('=D2:D6')).toBe(true);
    expect(dyn('=D2:D6*2')).toBe(true);
    expect(dyn('=SEQUENCE(3)')).toBe(true);
    expect(dyn('=単価*2')).toBe(false); // a single-cell name
    expect(dyn('=列*2')).toBe(true); // a sheet-scoped range name
    expect(dyn('=VLOOKUP("00123",A2:E6,4,FALSE)')).toBe(false);
  });

  it('editing a table header renames the table column and its structured references', () => {
    const s0 = tableRoster();
    const r = setCellInput(s0, 'b4', 1, '得点');
    const s = ok(r);
    expect(at(s.doc, 'D1')).toEqual({ t: 's', v: '得点', s: XF.header });
    expect(s.doc.features.tables[0].columns).toEqual(['学生ID', '氏名', '大学', '得点', '倍']);
    expect(fAt(s.doc, 'E2')).toBe('名簿T[[#This Row],[得点]]*2');
    expect(ext(s.doc, '集計', 'B8')).toBe('SUM(名簿T[得点])');
    expect(ext(s.doc, '集計', 'B9')).toBe('名簿T[[#Headers],[得点]]');
    expect(s.doc.names.find((n) => n.name === '点数列')!.refersTo).toBe('名簿T[得点]');
    expect(s.log).toEqual([{ kind: 'cells', sheet: '名簿', writes: [{ row: 1, col: 4, cell: { t: 's', v: '得点' } }], reason: 'edit' }]);
    expect(s.label).toBe('見出しを変更');
    // duplicates get a number like Excel, formulas are refused, numbers stay text
    const dup = setCellInput(s0, 'b4', 1, '氏名');
    expect(at(ok(dup).doc, 'D1')?.v).toBe('氏名2');
    expect(dup.ok && dup.warnings[0]).toMatch(/氏名2/);
    expect(refused(setCellInput(s0, 'b4', 1, '=1+1'))).toMatch(/見出し/);
    expect(at(ok(setCellInput(s0, 'b4', 1, '2026')).doc, 'D1')).toEqual({ t: 's', v: '2026', s: XF.header });
    const cleared = ok(setCellInput(s0, 'b4', 1, ''));
    expect(at(cleared.doc, 'D1')?.v).toBe('列1');
  });

  it('typed [@列] inside a table is stored with the table name and shown like Excel', () => {
    const s = ok(setCellInput(tableRoster(), 'b5', 3, '=[@点数]*3'));
    expect(fAt(s.doc, 'E3')).toBe('名簿T[[#This Row],[点数]]*3');
    expect(formulaForDisplay(s.doc, 4, 3)).toBe('=[@点数]*3');
    // outside the table the table name is shown
    const s2 = ok(setCellInput(tableRoster(), 'b6', 3, '=名簿T[@点数]'));
    expect(fAt(s2.doc, 'H3')).toBe('名簿T[[#This Row],[点数]]');
    expect(formulaForDisplay(s2.doc, 7, 3)).toBe('=名簿T[@点数]');
  });
});

describe('setColumnWidth', () => {
  it('sets the width and logs it at the current position', () => {
    const s0 = roster();
    const s = ok(setColumnWidth(s0, 'b2', 20.5));
    expect(s.doc.columns[1].props.width).toBe(20.5);
    expect(s.doc.columns[1].props.customWidth).toBe(true);
    expect(s.log).toEqual([{ kind: 'colWidth', sheet: '名簿', col: 2, width: 20.5 }]);
    expect(s.label).toBe('列幅を変更');
    expect(s.doc.columns[1].cells).toBe(s0.doc.columns[1].cells);
    const same = setColumnWidth(s0, 'b2', 12);
    expect(same.ok && same.snapshot).toBe(s0);
  });

  it('a width the user sets is a custom width even when it equals the stored auto width (Excel: customWidth="1")', () => {
    const s0 = roster((sp) => {
      sp.sheets[0].colProps!.B = { width: 12 }; // an auto width as Excel wrote it
    });
    expect(s0.doc.columns[1].props.customWidth).toBeUndefined();
    const r = setColumnWidth(s0, 'b2', 12);
    const s = ok(r);
    expect(s).not.toBe(s0);
    expect(s.doc.columns[1].props).toEqual({ width: 12, customWidth: true });
    expect(s.log).toEqual([{ kind: 'colWidth', sheet: '名簿', col: 2, width: 12 }]);
    // and now typed dates do not widen it any more
    const narrow = ok(setColumnWidth(s0, 'b2', 6));
    expect(ok(setCellInput(narrow, 'b2', 2, '2026/10/20')).doc.columns[1].props.width).toBe(6);
  });

  it('unhides, clamps and refuses bad widths', () => {
    const s0 = roster((sp) => {
      sp.sheets[0].colProps!.C = { width: 14, hidden: true };
    });
    const s = ok(setColumnWidth(s0, 'b3', 14));
    expect(s.doc.columns[2].props).toEqual({ width: 14, customWidth: true });
    const r = setColumnWidth(s0, 'b3', 300);
    expect(ok(r).doc.columns[2].props.width).toBe(255);
    expect(r.ok && r.warnings[0]).toMatch(/255/);
    expect(refused(setColumnWidth(s0, 'b3', -1))).toMatch(/列幅/);
    expect(refused(setColumnWidth(s0, 'b3', Number.NaN))).toMatch(/列幅/);
    expect(refused(setColumnWidth(s0, 'zz', 5))).toMatch(/列/);
  });
});
