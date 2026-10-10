/**
 * Round 9 (plan A), pure logic of the UI: the coupling of the two sheet checkboxes, the planned sheet list of the merge preview
 * (conflicts, long names, card order, files without a key), the counts / lines after the merge, the reducer actions, and the
 * messages in the 警告 notes. No DOM.
 */
import { describe, expect, it } from 'vitest';
import type { PremergeLayout } from '../../src/core/contracts';
import type { MergeReport, SourceBook } from '../../src/core/types';
import { planSheetNames } from '../../src/core/xlsx/sheetnames';
import { droppedBaseSheetsLine, fidelityNotes, SHEET_COPY_LIMITS } from '../../src/state/notes';
import { initialState, reducer } from '../../src/state/reducer';
import {
  addedSheetCount,
  addedSheetsLine,
  isNameShortened,
  normalizeSheetOptions,
  outputSheetCounts,
  outputSheetRows,
  outputSheetsLine,
  plannedOutputSheets,
  sameSheetOptions,
  suffixNumber,
  withOtherSheets,
  withSourceSheets,
  type SheetOptions,
} from '../../src/state/sheetPlan';
import type { AppState, FileEntry } from '../../src/state/types';
import { book } from './fixtures';

type SheetSpec = string | [name: string, state: 'visible' | 'hidden' | 'veryHidden'];

/** a loaded file with the given sheets (the first sheet is the data sheet unless `data` says otherwise) */
function fe(id: string, fileName: string, sheets: SheetSpec[], data?: string): FileEntry {
  const grids = sheets.map((s) => ({ name: typeof s === 'string' ? s : s[0], state: typeof s === 'string' ? 'visible' : s[1] }));
  const sheet = data ?? grids[0].name;
  return {
    id,
    fileName,
    size: 1,
    color: `#${id.padStart(6, '0')}`,
    status: 'ready',
    book: { id, fileName, sheets: grids } as unknown as SourceBook,
    sheet,
    region: { sheet, headerRow: 1, firstDataRow: 2, lastDataRow: 9, firstCol: 1, lastCol: 3 },
  };
}

const both: SheetOptions = { includeOtherSheets: true, includeSourceSheets: true };
const baseOnly: SheetOptions = { includeOtherSheets: true, includeSourceSheets: false };
const off: SheetOptions = { includeOtherSheets: false, includeSourceSheets: false };
const names = (p: { sheets: Array<{ name: string }> }) => p.sheets.map((s) => s.name);

describe('the two checkboxes are coupled', () => {
  it('turning the first on turns the second on (the user may switch it off again)', () => {
    expect(withOtherSheets(off, true)).toEqual(both);
    expect(withSourceSheets(both, false)).toEqual(baseOnly);
  });

  it('turning the second on while the first is off turns the first on too (the user wants all sheets)', () => {
    expect(withSourceSheets(off, true)).toEqual(both);
    expect(withSourceSheets(baseOnly, true)).toEqual(both);
  });

  it('turning the first off turns the second off', () => {
    expect(withOtherSheets(both, false)).toEqual(off);
    expect(withOtherSheets(baseOnly, false)).toEqual(off);
  });

  it('switching the second off keeps the first on; switching the first on again brings the second back (default on)', () => {
    const a = withSourceSheets(both, false);
    expect(a).toEqual(baseOnly);
    expect(withOtherSheets(withOtherSheets(a, false), true)).toEqual(both);
  });

  it('a request that changes nothing returns the very same object', () => {
    expect(withOtherSheets(off, false)).toBe(off);
    expect(withOtherSheets(both, true)).toBe(both);
    expect(withSourceSheets(both, true)).toBe(both);
    expect(withSourceSheets(off, false)).toBe(off);
    expect(withSourceSheets(baseOnly, false)).toBe(baseOnly);
    expect(withOtherSheets(baseOnly, true)).toBe(baseOnly);
  });

  it('every sequence of clicks ends in one of the three allowed states (never "copies without the base\'s sheets")', () => {
    const clicks: Array<(o: SheetOptions) => SheetOptions> = [
      (o) => withOtherSheets(o, true),
      (o) => withOtherSheets(o, false),
      (o) => withSourceSheets(o, true),
      (o) => withSourceSheets(o, false),
    ];
    let seen = 0;
    const walk = (o: SheetOptions, depth: number) => {
      expect(!(o.includeSourceSheets && !o.includeOtherSheets)).toBe(true);
      seen++;
      if (depth === 0) return;
      for (const c of clicks) walk(c(o), depth - 1);
    };
    walk(off, 5);
    expect(seen).toBeGreaterThan(1000);
  });

  it('normalizeSheetOptions repairs the unreachable state; sameSheetOptions compares by value', () => {
    expect(normalizeSheetOptions({ includeOtherSheets: false, includeSourceSheets: true })).toEqual(off);
    expect(normalizeSheetOptions(both)).toBe(both);
    expect(withOtherSheets({ includeOtherSheets: false, includeSourceSheets: true }, false)).toEqual(off);
    expect(sameSheetOptions(both, { ...both })).toBe(true);
    expect(sameSheetOptions(both, baseOnly)).toBe(false);
  });
});

describe('reducer: output/otherSheets and output/sourceSheets', () => {
  const step = (s: AppState, a: Parameters<typeof reducer>[1]) => reducer(s, a);
  const opts = (s: AppState) => ({ includeOtherSheets: s.includeOtherSheets, includeSourceSheets: s.includeSourceSheets });

  it('both are off by default (the single-sheet output)', () => {
    expect(opts(initialState)).toEqual(off);
  });

  it('follows the coupling rules', () => {
    let s = step(initialState, { type: 'output/otherSheets', include: true });
    expect(opts(s)).toEqual(both);
    s = step(s, { type: 'output/sourceSheets', include: false });
    expect(opts(s)).toEqual(baseOnly);
    s = step(s, { type: 'output/otherSheets', include: false });
    expect(opts(s)).toEqual(off);
    s = step(s, { type: 'output/sourceSheets', include: true });
    expect(opts(s)).toEqual(both);
    s = step(s, { type: 'output/otherSheets', include: false });
    expect(opts(s)).toEqual(off);
  });

  it('an action that changes nothing keeps the state object (no re-render)', () => {
    expect(step(initialState, { type: 'output/sourceSheets', include: false })).toBe(initialState);
    expect(step(initialState, { type: 'output/otherSheets', include: false })).toBe(initialState);
    const s = step(initialState, { type: 'output/sourceSheets', include: true });
    expect(step(s, { type: 'output/sourceSheets', include: true })).toBe(s);
    expect(step(s, { type: 'output/otherSheets', include: true })).toBe(s);
  });

  it('is a merge setting: ignored in the merged phase, kept by 結合をやり直す', () => {
    const s = step(initialState, { type: 'output/sourceSheets', include: true });
    const m: AppState = { ...s, phase: 'merged' };
    expect(step(m, { type: 'output/sourceSheets', include: false })).toBe(m);
    expect(step(m, { type: 'output/otherSheets', include: false })).toBe(m);
    const back = step(m, { type: 'merge/reset' });
    expect(opts(back)).toEqual(both);
    expect(opts(step(back, { type: 'output/otherSheets', include: false }))).toEqual(off);
  });
});

describe('plannedOutputSheets: the sheet list of the merge preview', () => {
  const base = () => fe('s1', '学生名簿.xlsx', ['名簿', '集計']);
  const interview = () => fe('s2', '02_面接評価.xlsx', ['面接評価', '備考']);
  const attendance = () => fe('s3', '研修出欠.xlsx', ['出欠']);

  it('1. the merged sheet, then the base\'s sheets, then <file>_<sheet> in card order, sheet by sheet', () => {
    const p = plannedOutputSheets([base(), interview(), attendance()], 's1', [], both);
    expect(names(p)).toEqual(['名簿', '集計', '02_面接評価_面接評価', '02_面接評価_備考', '研修出欠_出欠']);
    expect(p.sheets.map((s) => s.kind)).toEqual(['merged', 'base', 'copy', 'copy', 'copy']);
    expect(p.sheets[2]).toMatchObject({ sourceId: 's2', file: '02_面接評価.xlsx', original: '面接評価', suffixed: false, shortened: false });
    expect(p.sheets[4]).toMatchObject({ sourceId: 's3', file: '研修出欠.xlsx', original: '出欠' });
    expect(p.skipped).toEqual([]);
  });

  it('reordering the cards reorders the copies live (the base stays first)', () => {
    const files = [base(), attendance(), interview()];
    expect(names(plannedOutputSheets(files, 's1', [], both))).toEqual(['名簿', '集計', '研修出欠_出欠', '02_面接評価_面接評価', '02_面接評価_備考']);
    // the base card moved to the end: it is still the merged sheet and comes first
    expect(names(plannedOutputSheets([attendance(), interview(), base()], 's1', [], both))).toEqual(['名簿', '集計', '研修出欠_出欠', '02_面接評価_面接評価', '02_面接評価_備考']);
  });

  it('another file as the base: the first file becomes a source and is copied', () => {
    const p = plannedOutputSheets([base(), interview()], 's2', [], both);
    expect(names(p)).toEqual(['面接評価', '備考', '学生名簿_名簿', '学生名簿_集計']);
    expect(p.sheets.map((s) => s.kind)).toEqual(['merged', 'base', 'copy', 'copy']);
  });

  it('the data sheet of a source is copied too (a full original copy), not only its other sheets', () => {
    const f = fe('s2', '面接.xlsx', ['メモ', '評価'], '評価');
    expect(names(plannedOutputSheets([base(), f], 's1', [], both))).toEqual(['名簿', '集計', '面接_メモ', '面接_評価']);
  });

  it('the merged sheet is the base DATA sheet, wherever it stands in the base file', () => {
    const b = fe('s1', '学生名簿.xlsx', ['集計', 'マスタ', '名簿'], '名簿');
    const p = plannedOutputSheets([b, interview()], 's1', [], both);
    expect(p.sheets[0]).toMatchObject({ kind: 'merged', name: '名簿' });
    expect(names(p)).toEqual(['名簿', '集計', 'マスタ', '02_面接評価_面接評価', '02_面接評価_備考']);
  });

  it('only the merged sheet when both options are off; the base\'s sheets when only the first is on', () => {
    expect(names(plannedOutputSheets([base(), interview()], 's1', [], off))).toEqual(['名簿']);
    const first = plannedOutputSheets([base(), interview()], 's1', [], baseOnly);
    expect(names(first)).toEqual(['名簿', '集計']);
    expect(first.skipped).toEqual([]);
  });

  it('names that are taken get (1), (2) … on the sheet that is added LATER; the merged sheet and the base\'s sheets keep theirs', () => {
    const a = fe('s2', '名簿.xlsx', ['名簿']);
    const b = fe('s3', '名簿.xlsx', ['名簿']);
    const c = fe('s4', '名簿.xlsx', ['名簿']);
    const p = plannedOutputSheets([fe('s1', 'x.xlsx', ['名簿']), a, b, c], 's1', [], both);
    expect(names(p)).toEqual(['名簿', '名簿_名簿', '名簿_名簿(1)', '名簿_名簿(2)']);
    expect(p.sheets.map((s) => s.suffixed ?? false)).toEqual([false, false, true, true]);
    expect(p.sheets.map((s) => s.suffixNo)).toEqual([undefined, undefined, 1, 2]);
  });

  it('a name that is taken by a sheet of the base file or by the merged sheet is avoided too', () => {
    const b = fe('s1', '基準.xlsx', ['データ', '応募_回答']);
    const o = fe('s2', '応募.xlsx', ['回答']);
    const p = plannedOutputSheets([b, o], 's1', [], both);
    expect(names(p)).toEqual(['データ', '応募_回答', '応募_回答(1)']);
    expect(p.sheets[2]).toMatchObject({ suffixed: true, suffixNo: 1 });
    // the base's sheet is not kept (first option off, API level only): nothing to collide with
    expect(names(plannedOutputSheets([b, o], 's1', [], { includeOtherSheets: false, includeSourceSheets: true }))).toEqual(['データ', '応募_回答']);
  });

  it('names are compared like Excel does: case-insensitive, full / half width folded', () => {
    const a = fe('s2', 'ａ.xlsx', ['Ｓheet']);
    const b = fe('s3', 'A.xlsx', ['sheet']);
    expect(names(plannedOutputSheets([fe('s1', 'x.xlsx', ['x']), a, b], 's1', [], both))).toEqual(['x', 'ａ_Ｓheet', 'A_sheet(1)']);
  });

  it('too long names are shortened to 31 characters and flagged; short ones are not', () => {
    const long = fe('s2', 'とても長いファイル名のインターンシップ応募者一覧2026年度版.xlsx', ['回答']);
    const short = fe('s3', '短い.xlsx', ['回答']);
    const p = plannedOutputSheets([fe('s1', 'x.xlsx', ['x']), long, short], 's1', [], both);
    expect(p.sheets[1].name.length).toBeLessThanOrEqual(31);
    expect(p.sheets[1].shortened).toBe(true);
    expect(p.sheets[1].name.endsWith('_回答')).toBe(true);
    expect(p.sheets[2]).toMatchObject({ name: '短い_回答', shortened: false });
    // exactly 31 characters is not shortened; 32 is
    expect(isNameShortened('a'.repeat(10) + '.xlsx', 'b'.repeat(20))).toBe(false);
    expect(isNameShortened('a'.repeat(10) + '.xlsx', 'b'.repeat(21))).toBe(true);
  });

  it('a shortened name that collides still gets its (n) within 31 characters', () => {
    const f1 = fe('s2', 'あ'.repeat(40) + '.xlsx', ['回答']);
    const f2 = fe('s3', 'あ'.repeat(40) + '.xlsx', ['回答']);
    const p = plannedOutputSheets([fe('s1', 'x.xlsx', ['x']), f1, f2], 's1', [], both);
    expect(p.sheets[2].name.endsWith('(1)')).toBe(true);
    expect(p.sheets[2].name.length).toBeLessThanOrEqual(31);
    expect(p.sheets[2]).toMatchObject({ suffixed: true, shortened: true });
  });

  it('characters Excel forbids in a sheet name become "_" in the planned name', () => {
    const p = plannedOutputSheets([fe('s1', 'x.xlsx', ['x']), fe('s2', '報告/集計.xlsx', ['Q1:Q2?'])], 's1', [], both);
    expect(p.sheets[1].name).not.toMatch(/[\\/?*[\]:]/);
    expect(p.sheets[1].original).toBe('Q1:Q2?');
  });

  it('hidden sheets are flagged (a copy stays hidden)', () => {
    const o = fe('s2', 'a.xlsx', ['公開', ['内部', 'hidden'], ['奥', 'veryHidden']]);
    const p = plannedOutputSheets([fe('s1', 'x.xlsx', ['x', ['隠し', 'hidden']]), o], 's1', [], both);
    expect(p.sheets.map((s) => s.hidden)).toEqual([false, true, false, true, true]);
  });

  it('a file without a join key (結合対象外) takes no part: no copies, and the list says so', () => {
    const p = plannedOutputSheets([base(), interview(), attendance()], 's1', ['s2'], both);
    expect(names(p)).toEqual(['名簿', '集計', '研修出欠_出欠']);
    expect(p.skipped).toEqual(['02_面接評価.xlsx']);
  });

  it('files that are loading or failed are not part of the plan; without a base nothing is planned', () => {
    const loading: FileEntry = { id: 's3', fileName: 'c.xlsx', size: 1, color: '#000', status: 'loading' };
    const failed: FileEntry = { id: 's4', fileName: 'd.xlsx', size: 1, color: '#000', status: 'error', error: '壊れています' };
    expect(names(plannedOutputSheets([base(), interview(), loading, failed], 's1', [], both))).toEqual(['名簿', '集計', '02_面接評価_面接評価', '02_面接評価_備考']);
    expect(plannedOutputSheets([loading], null, [], both)).toEqual({ sheets: [], skipped: [] });
    expect(plannedOutputSheets([], null, [], both)).toEqual({ sheets: [], skipped: [] });
  });

  it('is exactly planSheetNames of the engine (same input -> same names)', () => {
    const files = [base(), interview(), fe('s3', '02_面接評価.xlsx', ['面接評価']), attendance()];
    const p = plannedOutputSheets(files, 's1', [], both);
    const engine = planSheetNames(
      ['名簿', '集計'],
      files.slice(1).map((f) => ({ file: f.book!.fileName, sheets: f.book!.sheets.map((s) => s.name) })),
    );
    expect(p.sheets.slice(2).map((s) => s.name)).toEqual(engine.map((e) => e.name));
    expect(p.sheets.slice(2).map((s) => s.suffixed)).toEqual(engine.map((e) => e.suffixed));
  });

  it('works with the real mock books too (応募者一覧 + 集計 as the base)', async () => {
    const b1 = await book('学生回答.xlsx', 's1');
    const b2 = await book('職場回答.xlsx', 's2');
    const mk = (id: string, bk: SourceBook, data: string): FileEntry => ({
      id,
      fileName: bk.fileName,
      size: 1,
      color: '#000',
      status: 'ready',
      book: bk,
      sheet: data,
      region: { sheet: data, headerRow: 1, firstDataRow: 2, lastDataRow: 9, firstCol: 1, lastCol: 3 },
    });
    expect(names(plannedOutputSheets([mk('s1', b1, '応募者一覧'), mk('s2', b2, '回答')], 's1', [], both))).toEqual(['応募者一覧', '集計', '職場回答_回答']);
  });
});

describe('suffix number', () => {
  it('reads the (n) the engine appended', () => {
    expect(suffixNumber('名簿_名簿(1)')).toBe(1);
    expect(suffixNumber('名簿_名簿(12)')).toBe(12);
    expect(suffixNumber('名簿_名簿')).toBeUndefined();
    expect(suffixNumber('a(1)b')).toBeUndefined();
  });
});

describe('after the merge: counts, line, rows', () => {
  it('「出力は N シート（結合 1 + 基準のほか M + 他ファイル K）」', () => {
    expect(outputSheetsLine(outputSheetCounts(7, 5))).toBe('出力は 7 シート（結合 1 + 基準のほか 1 + 他ファイル 5）');
    expect(outputSheetsLine(outputSheetCounts(3, 2))).toBe('出力は 3 シート（結合 1 + 基準のほか 0 + 他ファイル 2）');
    expect(outputSheetsLine(outputSheetCounts(1, 0))).toBe('出力は 1 シート（結合 1 + 基準のほか 0 + 他ファイル 0）');
  });

  it('counts never go negative (a report that is larger than the workbook)', () => {
    expect(outputSheetCounts(2, 5)).toEqual({ total: 2, base: 0, added: 5 });
    expect(outputSheetCounts(0, 0)).toEqual({ total: 1, base: 0, added: 0 });
  });

  it('addedSheetCount: the layout first, else the report', () => {
    const layout = { addedSheets: [{ sourceId: 's2', file: 'a.xlsx', original: 'x', name: 'a_x', suffixed: false, complete: true }] };
    expect(addedSheetCount(layout, { addedSheets: ['a', 'b', 'c'] })).toBe(1);
    expect(addedSheetCount({}, { addedSheets: ['a', 'b'] })).toBe(2);
    expect(addedSheetCount(undefined, undefined)).toBe(0);
    expect(addedSheetCount({ addedSheets: [] }, { addedSheets: ['a'] })).toBe(0);
  });

  const layout: Pick<PremergeLayout, 'sheet' | 'addedSheets'> = {
    sheet: '名簿',
    addedSheets: [
      { sourceId: 's2', file: '02_面接評価.xlsx', original: '面接評価', name: '02_面接評価_面接評価', suffixed: false, complete: true },
      { sourceId: 's3', file: '02_面接評価.xlsx', original: '面接評価', name: '02_面接評価_面接評価(1)', suffixed: true, complete: false },
    ],
  };
  const sheets = [
    { name: '名簿', state: 'visible' },
    { name: '集計', state: 'visible' },
    { name: '02_面接評価_面接評価', state: 'visible' },
    { name: '02_面接評価_面接評価(1)', state: 'hidden' },
  ];

  it('outputSheetRows tells the merged sheet, the base\'s sheets and the copies apart (names only)', () => {
    const rows = outputSheetRows(sheets, layout, { addedSheets: layout.addedSheets!.map((a) => a.name) });
    expect(rows.map((r) => [r.kind, r.name])).toEqual([
      ['merged', '名簿'],
      ['base', '集計'],
      ['copy', '02_面接評価_面接評価'],
      ['copy', '02_面接評価_面接評価(1)'],
    ]);
    expect(rows[3]).toMatchObject({ hidden: true, suffixed: true, complete: false, sourceId: 's3', original: '面接評価' });
    expect(rows[1].hidden).toBe(false);
  });

  it('falls back to the report\'s name list when the layout has no details', () => {
    const rows = outputSheetRows(sheets, { sheet: '名簿' }, { addedSheets: ['02_面接評価_面接評価', '02_面接評価_面接評価(1)'] });
    expect(rows.map((r) => r.kind)).toEqual(['merged', 'base', 'copy', 'copy']);
    expect(rows[2].file).toBeUndefined();
  });

  it('without any added sheets every other sheet is a base sheet', () => {
    const rows = outputSheetRows(sheets.slice(0, 2), { sheet: '名簿' }, {});
    expect(rows.map((r) => r.kind)).toEqual(['merged', 'base']);
  });

  it('addedSheetsLine for the toast detail', () => {
    expect(addedSheetsLine([])).toBe('');
    expect(addedSheetsLine(['a_b'])).toBe('他のファイルのシート 1 枚を追加しました（a_b）');
    expect(addedSheetsLine(['02_面接評価_面接評価', 'x'])).toBe('他のファイルのシート 2 枚を追加しました（02_面接評価_面接評価 ほか）');
    expect(addedSheetsLine(['あ'.repeat(30), 'x'])).toContain(`${'あ'.repeat(19)}…`);
  });
});

describe('警告 notes with the sheets of the other files', () => {
  const report = (patch: Partial<MergeReport> = {}): MergeReport => ({ mergedRows: 3, matchedRows: 3, baseOnlyRows: 0, appendedRows: 0, importedColumns: 2, formulasKept: 0, formulasToValues: 0, items: [], ...patch });
  const keys = (notes: Array<{ key: string }>) => notes.map((n) => n.key);

  async function files(): Promise<FileEntry[]> {
    const b1 = await book('学生回答.xlsx', 's1');
    const b2 = await book('職場回答.xlsx', 's2');
    const mk = (id: string, bk: SourceBook, data: string): FileEntry => ({
      id,
      fileName: bk.fileName,
      size: 1,
      color: '#000',
      status: 'ready',
      book: bk,
      sheet: data,
      region: { sheet: data, headerRow: 1, firstDataRow: 2, lastDataRow: 5, firstCol: 1, lastCol: 3 },
    });
    return [mk('s1', b1, '応募者一覧'), mk('s2', b2, '回答')];
  }

  it('the sheetCopyLimits items of the report are listed as warnings with their own key (the message names the sheet)', async () => {
    const item = { severity: 'warning' as const, code: SHEET_COPY_LIMITS, message: '「職場回答_回答」: グラフ 1 個はコピーできませんでした', sourceId: 's2' };
    const notes = fidelityNotes(await files(), 's1', true, report({ addedSheets: ['職場回答_回答'], items: [item, { severity: 'info', code: 'x', message: 'ほか' }] }), {
      merged: true,
      includeSourceSheets: true,
    });
    const n = notes.find((x) => x.key.startsWith('sheetcopy'));
    expect(n).toMatchObject({ severity: 'warning', text: '「職場回答_回答」: グラフ 1 個はコピーできませんでした' });
    expect(keys(notes)).toContain('r1');
    expect(SHEET_COPY_LIMITS).toBe('sheetCopyLimits');
  });

  it('a file with charts / pictures / pivots: the note says they go with the copied sheets, not "not imported"', async () => {
    const f = await files();
    f[1] = { ...f[1], book: { ...f[1].book!, features: { ...f[1].book!.features, charts: 2, images: 1 } } };
    const on = fidelityNotes(f, 's1', true, null, { includeSourceSheets: true });
    const objs = on.find((n) => n.key === 's2-objs')!;
    expect(objs.text).toContain('コピーしたシート側に入ります');
    expect(objs.text).not.toContain('取り込まれません（取り込むのはセルの値と書式です）');
    const off = fidelityNotes(f, 's1', true, null, { includeSourceSheets: false });
    expect(off.find((n) => n.key === 's2-objs')!.text).toContain('グラフ・図・ピボットは取り込まれません');
  });

  it('asked for but nothing was copied: a warning instead of a promise', async () => {
    const f = await files();
    const none = fidelityNotes(f, 's1', true, report(), { merged: true, includeSourceSheets: true });
    expect(none.find((n) => n.key === 'sheets-not-added')).toMatchObject({ severity: 'warning' });
    const ok = fidelityNotes(f, 's1', true, report({ addedSheets: ['職場回答_回答'] }), { merged: true, includeSourceSheets: true });
    expect(keys(ok)).not.toContain('sheets-not-added');
    // option off, or not merged yet: never shown
    expect(keys(fidelityNotes(f, 's1', true, report(), { merged: true }))).not.toContain('sheets-not-added');
    expect(keys(fidelityNotes(f, 's1', true, null, { includeSourceSheets: true }))).not.toContain('sheets-not-added');
  });

  it('dropped base sheets: with copies the line does not claim a one-sheet output', async () => {
    const f = await files();
    const only = fidelityNotes(f, 's1', true, report({ droppedSheets: ['集計'] }), { merged: true, dropOtherSheets: true });
    expect(only.find((n) => n.key === 'dropped-sheets')!.text).toContain('出力は結合した 1 シートのみです');
    const copies = fidelityNotes(f, 's1', true, report({ droppedSheets: ['集計'], addedSheets: ['職場回答_回答'] }), { merged: true, dropOtherSheets: true, includeSourceSheets: true });
    const t = copies.find((n) => n.key === 'dropped-sheets')!.text;
    expect(t).not.toContain('1 シートのみ');
    expect(t).toBe(droppedBaseSheetsLine(['集計']));
    expect(droppedBaseSheetsLine(['集計', 'マスタ'])).toBe('基準ファイルの他のシート 2 枚「集計」「マスタ」 は含めません');
  });
});
