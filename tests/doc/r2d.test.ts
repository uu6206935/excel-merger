// Round-2 doc features: title merges that block a move (fix 'unmergeTitles' + the 'unmerge' op), CF / DV spreading
// by rule type, converting array / spill masters before a delete.
import { describe, expect, it } from 'vitest';
import { analyzeDelete, deleteColumns, insertColumns, moveColumns, setCellInput } from '../../src/core/doc';
import type { FeatureRule, XStyle } from '../../src/core/types';
import { F, a1, at, ext, ids, makeEditor, ok, refused, roster } from './helpers';

const TITLE: XStyle = { bold: true, size: 14, hAlign: 'center', indent: 1 };

/** 報告書: title merged over A1:E1 (row 1), header row 3, data rows 4..6. */
function titled(extra: { merges?: Array<{ r1: number; c1: number; r2: number; c2: number }>; tables?: boolean } = {}) {
  const cells: Record<string, unknown> = {
    A1: { t: 's', v: '2026年度 名簿', s: 1 },
    B1: { t: 'z', v: null, s: 1 },
    A3: '番号', B3: '氏名', C3: '点数', D3: '倍', E3: '備考', F3: '外',
  };
  for (let r = 4; r <= 6; r++) {
    cells[`A${r}`] = r - 3;
    cells[`B${r}`] = `名${r}`;
    cells[`C${r}`] = r * 10;
    cells[`D${r}`] = F(`C${r}*2`, r * 20);
  }
  return makeEditor({
    styles: [{}, TITLE],
    sheets: [
      {
        name: '報告',
        cells: cells as never,
        maxCol: 6,
        features: {
          merges: extra.merges ?? [{ r1: 1, c1: 1, r2: 1, c2: 5 }],
          tables: extra.tables ? [{ name: 'T1', ref: 'A3:D6', headerRow: true, totalsRow: false, columns: ['番号', '氏名', '点数', '倍'] }] : [],
        },
      },
    ],
    region: { headerRow: 3, firstDataRow: 4, lastDataRow: 6 },
  });
}

describe('moveColumns: title merges above the header row', () => {
  it('refuses with a fix, then unmerges (same look) and moves in ONE snapshot', () => {
    const s0 = titled();
    const r = moveColumns(s0, ['b2'], 4);
    expect(r.ok).toBe(false);
    if (r.ok) return;
    expect(r.reason).toMatch(/結合されたセル（A1:E1）/);
    expect(r.fix).toEqual({
      kind: 'unmergeTitles',
      refs: ['A1:E1'],
      message: 'タイトル行の結合セル（A1:E1）が列の移動を妨げています。見た目はそのままで結合だけを解除（選択範囲内で中央揃え）して移動しますか？',
    });
    const done = moveColumns(s0, ['b2'], 4, { unmergeTitles: true });
    const s = ok(done);
    expect(s.log[0]).toEqual({ kind: 'unmerge', sheet: '報告', refs: ['A1:E1'], centerAcross: true });
    expect(s.log.slice(1)).toEqual([{ kind: 'struct', sheet: '報告', op: { type: 'move', axis: 'col', from: 2, count: 1, to: 5 } }]);
    expect(s.label).toBe('結合を解除して列を移動');
    expect(done.ok && done.warnings.join()).toMatch(/選択範囲内で中央/);
    expect(s.doc.features.merges).toEqual([]);
    expect(ids(s.doc)).toEqual(['b1', 'b3', 'b4', 'b2', 'b5', 'b6']);
    // every cell of the former area: its style + centerContinuous (indent dropped, like Excel); missing cells are created
    const centered: XStyle = { bold: true, size: 14, hAlign: 'centerContinuous' };
    expect(at(s.doc, 'A1')).toMatchObject({ t: 's', v: '2026年度 名簿' });
    for (const a of ['A1', 'D1']) expect(s.doc.styles[at(s.doc, a)!.s!], a).toEqual(centered); // B1 moved to D1
    for (const a of ['B1', 'C1', 'E1']) {
      const c = at(s.doc, a)!;
      expect(c, a).toMatchObject({ t: 'z', v: null });
      expect(s.doc.styles[c.s!], a).toEqual({ hAlign: 'centerContinuous' });
    }
    expect(at(s.doc, 'F1')).toBeUndefined();
    // the formula followed its column; the input snapshot is untouched
    expect(at(s.doc, 'C4')!.f).toBe('B4*2');
    expect(s0.doc.features.merges.length).toBe(1);
    expect(s0.doc.styles.length).toBe(2);
    expect(at(s0.doc, 'C1')).toBeUndefined();
  });

  it('does not unmerge when the move is not blocked', () => {
    const s = ok(moveColumns(titled(), ['b6'], 5, { unmergeTitles: true })); // F moves nowhere near the title
    expect(s.log.every((e) => e.kind === 'struct')).toBe(true);
    const moved = ok(moveColumns(titled(), ['b1', 'b2', 'b3', 'b4', 'b5'], 6, { unmergeTitles: true })); // the whole title moves
    expect(moved.log.every((e) => e.kind === 'struct')).toBe(true);
    expect(moved.doc.features.merges).toEqual([{ r1: 1, c1: 2, r2: 1, c2: 6 }]);
  });

  it('keeps refusing merges that touch the header / data rows, and other obstacles', () => {
    const header = titled({ merges: [{ r1: 1, c1: 1, r2: 1, c2: 5 }, { r1: 3, c1: 2, r2: 3, c2: 3 }] });
    const r = moveColumns(header, ['b2'], 4, { unmergeTitles: true });
    expect(r.ok).toBe(false);
    expect(!r.ok && r.reason).toMatch(/B3:C3/);
    expect(!r.ok && r.fix).toBeUndefined();
    const tall = titled({ merges: [{ r1: 1, c1: 1, r2: 3, c2: 2 }] }); // reaches the header row
    const t = moveColumns(tall, ['b2'], 4);
    expect(!t.ok && t.fix).toBeUndefined();
    // a table in the way: no title fix offered
    const tab = titled({ tables: true });
    const u = moveColumns(tab, ['b2'], 5);
    expect(u.ok).toBe(false);
    expect(!u.ok && u.reason).toMatch(/テーブル/);
    expect(!u.ok && u.fix).toBeUndefined();
  });

  it('offers only the title merges that actually block', () => {
    const s0 = titled({ merges: [{ r1: 1, c1: 1, r2: 1, c2: 3 }, { r1: 2, c1: 4, r2: 2, c2: 5 }] });
    const r = moveColumns(s0, ['b2'], 0);
    expect(!r.ok && r.fix?.refs).toEqual(['A1:C1']);
    const s = ok(moveColumns(s0, ['b2'], 0, { unmergeTitles: true }));
    expect(s.doc.features.merges).toEqual([{ r1: 2, c1: 4, r2: 2, c2: 5 }]);
  });
});

describe('setCellInput: number formats of typed values', () => {
  it('widens a standard-width column like Excel (and the patcher) when the formatted value no longer fits', () => {
    const s0 = roster((sp) => {
      delete sp.sheets[0].colProps!.H;
    });
    const s = ok(setCellInput(s0, 'b6', 2, '2026/10/20'));
    expect(s.doc.columns[7].props.width).toBe(11.375); // 10.75 characters, as Excel shows it (golden r-numfmt-w)
    expect(s.log.length).toBe(1); // no width op: the replay widens by itself, like Excel
    expect(ok(setCellInput(s0, 'b6', 2, '9/30')).doc.columns[7].props.width).toBeUndefined(); // 9月30日 fits
    expect(ok(setCellInput(s0, 'b6', 2, '50%')).doc.columns[7].props.width).toBeUndefined();
    // a column with a width keeps it
    expect(ok(setCellInput(roster(), 'b6', 2, '2026/10/20')).doc.columns[7].props.width).toBe(15);
  });
});

describe('CF / DV spreading by rule type (FeatureRule.type)', () => {
  const withRules = (rules: FeatureRule[]) =>
    roster((sp) => {
      sp.sheets[0].features = { merges: [], tables: [], arrays: [], hyperlinks: [], rules };
    });
  const sq = (s: { doc: { features: { rules: FeatureRule[] } } }) => s.doc.features.rules.map((r) => `${r.type}:${r.sqref}`);

  it('a block moved right after an area picks up cellIs / expression / DV rules, not statistical ones', () => {
    const s0 = withRules([
      { sqref: 'B2:B6', formulas: ['B2="佐藤"'], kind: 'cf', type: 'expression' },
      { sqref: 'B2:B6', formulas: ['"x"'], kind: 'cf', type: 'cellIs' },
      { sqref: 'B2:B6', formulas: [], kind: 'cf', type: 'colorScale' },
      { sqref: 'B2:B6', formulas: [], kind: 'cf', type: 'duplicateValues' },
      { sqref: 'B2:B6', formulas: ['"a,b"'], kind: 'dv', type: 'list' },
    ]);
    const s = ok(moveColumns(s0, ['b6'], 2)); // 備考 (H) lands right after B
    expect(sq(s)).toEqual(['expression:B2:C6', 'cellIs:B2:C6', 'colorScale:B2:B6', 'duplicateValues:B2:B6', 'list:B2:C6']);
  });

  it('every type grows over a column inserted right after its area (DV included)', () => {
    const s0 = withRules([
      { sqref: 'B2:B6', formulas: [], kind: 'cf', type: 'dataBar' },
      { sqref: 'B2:B6', formulas: ['"a,b"'], kind: 'dv', type: 'list' },
    ]);
    const s = ok(insertColumns(s0, 2, 1));
    expect(sq(s)).toEqual(['dataBar:B2:C6', 'list:B2:C6']);
  });
});

describe('structural ops cutting through a dynamic-array (spill) range', () => {
  /** 名簿 H2:I4 = SEQUENCE(3,2) (dynamic array, cached values in every cell of the range) */
  const spilled = () =>
    roster((sp) => {
      const g = sp.sheets[0].cells!;
      g.H2 = { t: 'n', v: 1, f: '_xlfn.SEQUENCE(3,2)', arr: 'H2:I4', dyn: true };
      g.I2 = 2;
      g.H3 = 3;
      g.I3 = 4;
      g.H4 = 5;
      g.I4 = 6;
      sp.sheets[0].maxCol = 9;
      sp.sheets[0].features = { merges: [], rules: [], tables: [], hyperlinks: [], arrays: [{ r1: 2, c1: 8, r2: 4, c2: 9 }] };
    });
  const spillColumnValues = (s: ReturnType<typeof spilled>, col: string) => [2, 3, 4].map((r) => at(s.doc, `${col}${r}`)?.v ?? null);

  it('an insert inside the range collapses it to the master and drops the stale spilled values (as the patcher does)', () => {
    const s0 = spilled();
    const r = insertColumns(s0, 8, 1); // between H and I
    const s = ok(r);
    expect(s.doc.features.arrays).toEqual([{ r1: 2, c1: 8, r2: 2, c2: 8 }]);
    expect(at(s.doc, 'H2')).toMatchObject({ f: '_xlfn.SEQUENCE(3,2)', arr: 'H2', dyn: true });
    expect(spillColumnValues(s, 'H')).toEqual([1, null, null]);
    expect(spillColumnValues(s, 'J')).toEqual([null, null, null]); // the old I
    expect(r.ok && r.warnings.join()).toMatch(/スピル範囲（H2:I4）/);
    // the input snapshot is untouched
    expect(s0.doc.features.arrays).toEqual([{ r1: 2, c1: 8, r2: 4, c2: 9 }]);
    expect(at(s0.doc, 'I3')!.v).toBe(4);
  });

  it('inserting beside the range, or moving the whole range, keeps it (values travel with it)', () => {
    const before = ok(insertColumns(spilled(), 7, 1)); // before H
    expect(before.doc.features.arrays).toEqual([{ r1: 2, c1: 9, r2: 4, c2: 10 }]);
    expect(at(before.doc, 'J3')!.v).toBe(4);
    const after = ok(insertColumns(spilled(), 9, 1)); // after I (the end)
    expect(after.doc.features.arrays).toEqual([{ r1: 2, c1: 8, r2: 4, c2: 9 }]);
    const moved = ok(moveColumns(spilled(), [s0ColumnId(spilled(), 'H'), s0ColumnId(spilled(), 'I')], 0));
    expect(moved.doc.features.arrays).toEqual([{ r1: 2, c1: 1, r2: 4, c2: 2 }]);
    expect(at(moved.doc, 'B3')!.v).toBe(4);
    expect(at(moved.doc, 'A2')).toMatchObject({ arr: 'A2:B4' });
  });

  it('moving one column of the range (or landing inside it) collapses it', () => {
    const s0 = spilled();
    const out = ok(moveColumns(s0, [s0ColumnId(s0, 'I')], 0)); // I to the front: the master keeps its column
    expect(out.doc.features.arrays).toEqual([{ r1: 2, c1: 9, r2: 2, c2: 9 }]);
    expect(at(out.doc, 'I2')).toMatchObject({ arr: 'I2', dyn: true });
    expect(spillColumnValues(out, 'A')).toEqual([null, null, null]);
    const landing = ok(moveColumns(s0, ['b1'], 8)); // A lands between H and I
    expect(landing.doc.features.arrays).toEqual([{ r1: 2, c1: 7, r2: 2, c2: 7 }]);
    expect(at(landing.doc, 'G2')).toMatchObject({ arr: 'G2' });
  });

  it('deleting the master column leaves no orphan values; deleting a child column collapses the range', () => {
    const s0 = spilled();
    const noMaster = ok(deleteColumns(s0, [s0ColumnId(s0, 'H')], { convertDependents: false }));
    expect(noMaster.doc.features.arrays).toEqual([]);
    expect(spillColumnValues(noMaster, 'H')).toEqual([null, null, null]); // the old I
    const child = ok(deleteColumns(s0, [s0ColumnId(s0, 'I')], { convertDependents: false }));
    expect(child.doc.features.arrays).toEqual([{ r1: 2, c1: 8, r2: 2, c2: 8 }]);
    expect(spillColumnValues(child, 'H')).toEqual([1, null, null]);
    expect(at(child.doc, 'H2')).toMatchObject({ arr: 'H2' });
  });

  it('legacy (CSE) arrays are still refused, never collapsed', () => {
    const s0 = roster((sp) => {
      sp.sheets[0].cells!.H2 = { t: 'n', v: 1, f: 'C2:C3', arr: 'H2:I2' };
      sp.sheets[0].cells!.I2 = 2;
      sp.sheets[0].maxCol = 9;
      sp.sheets[0].features = { merges: [], rules: [], tables: [], hyperlinks: [], arrays: [{ r1: 2, c1: 8, r2: 2, c2: 9 }] };
    });
    expect(refused(insertColumns(s0, 8, 1))).toMatch(/配列数式/);
  });
});

/** id of the column at a letter of a fresh doc */
const s0ColumnId = (s: { doc: { columns: Array<{ id: string }> } }, letter: string): string => s.doc.columns[letter.charCodeAt(0) - 65].id;

describe('deleteColumns: converting array / spill masters', () => {
  it('converts a legacy array master on the merged sheet as a whole', () => {
    const s0 = roster((sp) => {
      sp.sheets[0].cells!.H2 = { t: 's', v: '東大', f: 'C2:C3', arr: 'H2:H3' };
      sp.sheets[0].cells!.H3 = '京大';
      sp.sheets[0].features = { merges: [], rules: [], tables: [], hyperlinks: [], arrays: [{ r1: 2, c1: 8, r2: 3, c2: 8 }] };
    });
    const r = deleteColumns(s0, ['b3'], { convertDependents: true });
    const s = ok(r);
    // the master write comes first, at the position before the delete
    expect(s.log[0]).toMatchObject({ kind: 'cells', sheet: '名簿', reason: 'toValues' });
    const w = s.log[0].kind === 'cells' ? s.log[0].writes : [];
    expect(w).toContainEqual({ row: 2, col: 8, cell: { t: 's', v: '東大' } });
    expect(s.log[s.log.length - 1]).toEqual({ kind: 'struct', sheet: '名簿', op: { type: 'delete', axis: 'col', at: 3, count: 1 } });
    expect(at(s.doc, 'G2')).toEqual({ t: 's', v: '東大' });
    expect(at(s.doc, 'G3')).toEqual({ t: 's', v: '京大' });
    expect(s.doc.features.arrays).toEqual([]);
    expect(r.ok && r.warnings.join()).toMatch(/配列数式 1 個/);
  });

  it('converting the array lifts the "part of an array" refusal', () => {
    // H2:I2 = {=C2&{1,2}} (array over two columns): deleting I (part of the array) is refused, unless the array itself
    // is converted because it depends on a deleted column (C)
    const s0 = roster((sp) => {
      sp.sheets[0].cells!.H2 = { t: 's', v: '東大1', f: 'C2&{1,2}', arr: 'H2:I2' };
      sp.sheets[0].cells!.I2 = '東大2';
      sp.sheets[0].maxCol = 9;
      sp.sheets[0].features = { merges: [], rules: [], tables: [], hyperlinks: [], arrays: [{ r1: 2, c1: 8, r2: 2, c2: 9 }] };
    });
    expect(refused(deleteColumns(s0, [s0.doc.columns[8].id], { convertDependents: true }))).toMatch(/配列/);
    const s = ok(deleteColumns(s0, ['b3', s0.doc.columns[8].id], { convertDependents: true }));
    expect(s.doc.features.arrays).toEqual([]);
    expect(at(s.doc, 'G2')).toEqual({ t: 's', v: '東大1' });
  });

  it('formulas reading a converted spill (A1#) break with it and are converted too (cascade)', () => {
    const withSpill = (sp: Parameters<typeof roster>[0] extends ((s: infer S) => void) | undefined ? S : never) => {
      const g = sp.sheets[0].cells!;
      g.H2 = { t: 's', v: '東大x', f: 'C2:C6&"x"', arr: 'H2:H6', dyn: true };
      ['京大', '阪大', '東北大', '九大'].forEach((u, i) => (g[`H${i + 3}`] = `${u}x`));
      g.H8 = F('COUNTA(_xlfn.ANCHORARRAY(H2))', 5); // a reader on the merged sheet
      sp.sheets[0].features = { merges: [], rules: [], tables: [], hyperlinks: [], arrays: [{ r1: 2, c1: 8, r2: 6, c2: 8 }] };
      const t = sp.sheets[1].cells!;
      t.B7 = F('_xlfn.ANCHORARRAY(名簿!H2)', '東大x', { arr: 'B7:B11', dyn: true }); // a reader that is a spill itself
      t.C2 = F('COUNTA(_xlfn.ANCHORARRAY(名簿!H2))', 5);
      t.C3 = F('SUM(B7:B11)', 0); // reads the second spill by a plain range: untouched
      sp.names = [...(sp.names ?? []), { name: 'スピル', scope: null, refersTo: '_xlfn.ANCHORARRAY(名簿!$H$2)' }];
    };
    const s0 = roster(withSpill);
    const impact = analyzeDelete(s0, ['b3']);
    const where = impact.breaks.map((b) => b.where).sort();
    expect(where).toEqual(['名前「スピル」', '名前「大学セル」', '名簿!H2', '名簿!H8', '集計!B3', '集計!B7', '集計!C2'].sort());
    expect(impact.breakCount).toBe(7);
    const r = deleteColumns(s0, ['b3'], { convertDependents: true });
    const s = ok(r);
    const writes = s.log.filter((e) => e.kind === 'cells').map((e) => (e.kind === 'cells' ? `${e.sheet}:${e.reason}:${e.writes.map((w) => `${w.row},${w.col}`).join(' ')}` : ''));
    // readers of the spill first (once H2 is a plain value, H2# is #REF! and their own spilled values are gone in Excel), the master last
    expect(writes).toEqual(['名簿:toValues:8,8', '集計:toValues:3,2 7,2 2,3', '名簿:toValues:2,8']);
    expect(s.doc.features.arrays).toEqual([]);
    expect(at(s.doc, 'G2')).toEqual({ t: 's', v: '東大x' }); // the master is a plain value now
    expect(at(s.doc, 'G8')).toEqual({ t: 'n', v: 5 });
    expect(s.doc.externalFormulas.map((e) => `${e.sheet}!${a1(e.col, e.row)}`)).toEqual(['集計!B2', '集計!B4', '集計!B5', '集計!B6', '集計!C3']);
    // the names cannot be converted: they are reported
    expect(r.ok && r.warnings.join('\n')).toMatch(/スピル/);
  });

  it('a spill reader of a master that is NOT converted stays a formula', () => {
    const s0 = roster((sp) => {
      const g = sp.sheets[0].cells!;
      g.H2 = { t: 's', v: '1', f: 'D2:D6&""', arr: 'H2:H6', dyn: true };
      sp.sheets[0].features = { merges: [], rules: [], tables: [], hyperlinks: [], arrays: [{ r1: 2, c1: 8, r2: 6, c2: 8 }] };
      sp.sheets[1].cells!.C2 = F('COUNTA(_xlfn.ANCHORARRAY(名簿!H2))', 5);
    });
    const s = ok(deleteColumns(s0, ['b3'], { convertDependents: true })); // C does not feed H
    expect(s.doc.features.arrays).toEqual([{ r1: 2, c1: 7, r2: 6, c2: 7 }]);
    expect(ext(s.doc, '集計', 'C2')).toBe('COUNTA(_xlfn.ANCHORARRAY(名簿!G2))');
  });

  it('converts a dynamic-array (spill) master typed in this session', () => {
    const typed = ok(setCellInput(roster(), 'b6', 2, '=SORT(C2:C6)'));
    expect(at(typed.doc, 'H2')).toMatchObject({ dyn: true, arr: 'H2' });
    const s = ok(deleteColumns(typed, ['b3'], { convertDependents: true }));
    // never calculated: the master is cleared (no cached value), the spill range is gone
    expect(at(s.doc, 'G2')).toBeUndefined();
    expect(s.doc.features.arrays).toEqual([]);
    expect(s.log[1]).toMatchObject({ kind: 'cells', reason: 'toValues', writes: [{ row: 2, col: 8, cell: null }] });
  });
});
