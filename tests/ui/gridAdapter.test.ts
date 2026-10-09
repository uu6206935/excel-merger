import { beforeAll, describe, expect, it } from 'vitest';
import type { EditorSnapshot, MergedDoc, SourceBook } from '../../src/core/types';
import { excelWidthToPx } from '../../src/core/display/styles';
import {
  addedColumnIds,
  buildGridModel,
  cellAddress,
  columnListItems,
  docMdw,
  isDataRow,
  keyRolesFromPlan,
  makeGetCell,
} from '../../src/state/gridAdapter';
import { editor, svc, type Joined } from './fixtures';

let j: Joined;
let snap: EditorSnapshot;
let mergedBook: SourceBook;
let doc: MergedDoc;
const headerText = (b: () => SourceBook) => (c: Parameters<typeof svc.cellText>[1]) => svc.cellText(b(), c);

beforeAll(async () => {
  const e = await editor();
  j = e.j;
  snap = e.snapshot;
  mergedBook = e.mergedBook;
  doc = snap.doc;
});

const model = (d: MergedDoc = doc, live?: Record<string, number>) =>
  buildGridModel(d, { baseId: 's1', keyRoles: keyRolesFromPlan(j.analysis, j.plan), headerText: headerText(() => mergedBook), liveWidths: live });

describe('grid model', () => {
  it('columns: titles from the header row, sources base-first, key badges on the base key columns', () => {
    const m = model();
    expect(m.columns).toHaveLength(18);
    expect(m.columns.slice(0, 3).map((c) => c.title)).toEqual(['No', '学生ID', '氏名']);
    expect(m.sources.map((s) => s.name)).toEqual(['学生回答（基準）', '職場回答', '人事回答']);
    expect(m.sourceIds).toEqual(['s1', 's2', 's3']);
    expect(m.columns[0].source).toBe(0);
    const imported = m.columns.find((c) => c.id.startsWith('i:s3'))!;
    expect(imported.source).toBe(2);
    expect(m.columns.filter((c) => c.keyRole).map((c) => [c.title, c.keyRole])).toEqual([
      ['学生ID', 'ID'],
      ['氏名', '氏名'],
    ]);
  });

  it('rows = the data rows (Excel numbers), frozen columns from the sheet', () => {
    const m = model();
    expect(m.rows[0]).toBe(3);
    expect(m.rows).toHaveLength(30);
    expect(m.frozenCols).toBe(3);
    expect(isDataRow(doc, 2)).toBe(false);
    expect(isDataRow(doc, 3)).toBe(true);
    expect(isDataRow(doc, 32)).toBe(true);
    expect(isDataRow(doc, 33)).toBe(false);
  });

  it('widths: Excel widths of the file via the default font, live widths win', () => {
    const mdw = docMdw(doc);
    expect(mdw).toBe(8); // 游ゴシック 11
    const m = model();
    expect(m.columns[1].width).toBe(excelWidthToPx(doc.columns[1].props.width!, mdw));
    expect(model(doc, { [doc.columns[1].id]: 222 }).columns[1].width).toBe(222);
  });

  it('empty header -> a placeholder title with the column letter; hidden columns are thin', () => {
    const cols = doc.columns.slice();
    cols[4] = { ...cols[4], cells: cols[4].cells.map((c, i) => (i === doc.headerRow - 1 ? undefined : c)), props: { ...cols[4].props, hidden: true } };
    const m = model({ ...doc, columns: cols });
    expect(m.columns[4].title).toBe('（E列）');
    expect(m.columns[4].hidden).toBe(true);
    expect(m.columns[4].width).toBe(8);
  });

  it('a header text function that throws never breaks the model', () => {
    const m = buildGridModel(doc, {
      baseId: 's1',
      keyRoles: new Map(),
      headerText: () => {
        throw new Error('x');
      },
    });
    expect(m.columns[0].title).toBe('（A列）');
  });
});

describe('getCell', () => {
  it('returns cached views with the formula in UI form', () => {
    const m = model();
    const get = makeGetCell(doc, m, (ci, row) => svc.formulaForDisplay(doc, ci, row));
    const total = m.columns[10];
    expect(total.title).toBe('合計');
    const v = get(3, total)!;
    expect(v.formula).toBe('=I3+J3');
    expect(typeof v.value).toBe('number');
    expect(get(3, total)).toBe(v); // cached
    expect(get(3, m.columns[1])!.value).toBe('10001');
    expect(get(3, { id: 'nope', title: '', width: 10 })).toBeUndefined();
  });

  it('a throwing formula display falls back to "=" + file form', () => {
    const m = model();
    const get = makeGetCell(doc, m, () => {
      throw new Error('x');
    });
    expect(get(4, m.columns[10])!.formula).toBe('=I4+J4');
  });

  it('empty cells of a styled column get the column style', () => {
    const cols = doc.columns.slice();
    cols[0] = { ...cols[0], props: { ...cols[0].props, s: 3 }, cells: cols[0].cells.map((c, i) => (i === 9 ? undefined : c)) };
    const d = { ...doc, columns: cols };
    const m = model(d);
    const get = makeGetCell(d, m, () => null);
    expect(get(10, m.columns[0])).toMatchObject({ value: null, styleId: 3 });
  });
});

describe('column list', () => {
  it('lists every column with letter, colour and source', () => {
    const items = columnListItems(model(), '#6b7280');
    expect(items[0]).toMatchObject({ letter: 'A', title: 'No', sourceLabel: '学生回答（基準）', color: '#2563eb' });
    expect(items[1].keyRole).toBe('ID');
    expect(items.at(-1)!.letter).toBe('R');
  });

  it('added column ids / addresses', async () => {
    const r = svc.insertColumns(snap, 2, 2);
    if (!r.ok) throw new Error(r.reason);
    const added = addedColumnIds(snap.doc, r.snapshot.doc);
    expect(added).toHaveLength(2);
    const items = columnListItems(model(r.snapshot.doc), '#6b7280');
    expect(items[2]).toMatchObject({ sourceLabel: 'アプリで追加', color: '#6b7280', title: '（C列）' });
    expect(cellAddress(0, 3)).toBe('A3');
    expect(cellAddress(27, 100)).toBe('AB100');
  });
});
