/**
 * Round 7: single-sheet output in the messages (info lines, 警告 notes). The old notes about the objects of the base book (charts,
 * pivots, …) must not promise anything about objects that may sit on a dropped sheet.
 */
import { beforeAll, describe, expect, it } from 'vitest';
import type { MergeReport, SourceBook } from '../../src/core/types';
import { droppedSheetsLine, fidelityNotes, formulasDetachedLine, otherSheetNames } from '../../src/state/notes';
import type { FileEntry } from '../../src/state/types';
import { book } from './fixtures';

let base: SourceBook;
let other: SourceBook;
const entry = (id: string, b: SourceBook, region = true): FileEntry => ({
  id,
  fileName: b.fileName,
  size: 1,
  color: '#000',
  status: 'ready',
  book: b,
  sheet: b.sheets[0].name,
  region: region ? { sheet: b.sheets[0].name, headerRow: 1, firstDataRow: 2, lastDataRow: 5, firstCol: 1, lastCol: 3 } : undefined,
});
const report = (patch: Partial<MergeReport> = {}): MergeReport => ({ mergedRows: 3, matchedRows: 3, baseOnlyRows: 0, appendedRows: 0, importedColumns: 2, formulasKept: 0, formulasToValues: 0, items: [], ...patch });
const keys = (notes: Array<{ key: string }>) => notes.map((n) => n.key);

beforeAll(async () => {
  base = await book('学生回答.xlsx', 's1'); // two sheets (応募者一覧, 集計) and one chart
  other = await book('職場回答.xlsx', 's2');
});

describe('info lines', () => {
  it('「出力は結合した 1 シートのみです（他のシート N 枚「集計」「マスタ」… は含めません）」', () => {
    expect(droppedSheetsLine(['集計'])).toBe('出力は結合した 1 シートのみです（他のシート 1 枚「集計」 は含めません）');
    expect(droppedSheetsLine(['集計', 'マスタ'])).toBe('出力は結合した 1 シートのみです（他のシート 2 枚「集計」「マスタ」 は含めません）');
    expect(droppedSheetsLine(['a', 'b', 'c', 'd', 'e'])).toBe('出力は結合した 1 シートのみです（他のシート 5 枚「a」「b」「c」… は含めません）');
    // a very long sheet name is shortened
    expect(droppedSheetsLine(['あ'.repeat(30)])).toContain(`「${'あ'.repeat(15)}…」`);
  });

  it('「他のシートを参照していた数式 N 個は値にしました」', () => {
    expect(formulasDetachedLine(7)).toBe('他のシートを参照していた数式 7 個は値にしました');
  });

  it('otherSheetNames: the sheets of the base book except the data sheet (a prediction before the merge)', () => {
    expect(otherSheetNames(entry('s1', base))).toEqual(['集計']);
    expect(otherSheetNames(entry('s2', other))).toEqual([]);
    expect(otherSheetNames(undefined)).toEqual([]);
    expect(otherSheetNames({ book: undefined })).toEqual([]);
  });
});

describe('警告 notes with the single-sheet output', () => {
  const files = () => [entry('s1', base), entry('s2', other)];

  it('the merge report adds the two info lines (and only when there is something to say)', () => {
    const notes = fidelityNotes(files(), 's1', true, report({ droppedSheets: ['集計', 'マスタ'], formulasDetached: 3 }), { merged: true, dropOtherSheets: true });
    expect(notes.find((n) => n.key === 'dropped-sheets')).toMatchObject({ severity: 'info', text: droppedSheetsLine(['集計', 'マスタ']) });
    expect(notes.find((n) => n.key === 'formulas-detached')).toMatchObject({ severity: 'info', text: '他のシートを参照していた数式 3 個は値にしました' });
    const none = fidelityNotes(files(), 's1', true, report({ droppedSheets: [], formulasDetached: 0 }), { merged: true, dropOtherSheets: true });
    expect(keys(none)).not.toContain('dropped-sheets');
    expect(keys(none)).not.toContain('formulas-detached');
    expect(keys(fidelityNotes(files(), 's1', true, report(), { merged: true }))).not.toContain('dropped-sheets');
  });

  it('with the other sheets dropped, the notes that promise to keep the base book\'s objects are not shown (they may sit on a dropped sheet)', () => {
    // before the merge: the base book has another sheet
    expect(keys(fidelityNotes(files(), 's1', true, null, { dropOtherSheets: true }))).not.toContain('base-charts');
    // after the merge: the report says a sheet was dropped
    expect(keys(fidelityNotes(files(), 's1', true, report({ droppedSheets: ['集計'] }), { merged: true, dropOtherSheets: true }))).not.toContain('base-charts');
  });

  it('with the other sheets kept, the old notes are unchanged — and the Excel engine hint stays correct', () => {
    const withExcel = fidelityNotes(files(), 's1', true, null, { dropOtherSheets: false });
    expect(withExcel.find((n) => n.key === 'base-charts')?.text).toBe('グラフ 1 個: Excel本体モードで確実に保持されます');
    const without = fidelityNotes(files(), 's1', false, null, { dropOtherSheets: false });
    expect(without.find((n) => n.key === 'base-charts')?.text).toContain('Excel のある PC');
    // callers that know nothing about the option (older call sites) get the old behaviour
    expect(keys(fidelityNotes(files(), 's1', true, null))).toContain('base-charts');
  });

  it('a base book with a single sheet loses nothing: its notes stay even in the single-sheet mode', () => {
    const oneSheet: SourceBook = { ...base, sheets: [base.sheets[0]] };
    const f = [entry('s1', oneSheet), entry('s2', other)];
    expect(keys(fidelityNotes(f, 's1', true, null, { dropOtherSheets: true }))).toContain('base-charts');
    // after the merge: nothing was dropped -> still true
    expect(keys(fidelityNotes(f, 's1', true, report({ droppedSheets: [] }), { merged: true, dropOtherSheets: true }))).toContain('base-charts');
  });

  it('notes that are not about sheet objects stay (values-only base, macros, external links)', () => {
    const vba: SourceBook = { ...base, features: { ...base.features, vba: true, externalLinks: 2 } };
    const notes = fidelityNotes([entry('s1', vba), entry('s2', other)], 's1', true, null, { dropOtherSheets: true });
    expect(keys(notes)).toEqual(expect.arrayContaining(['base-vba', 'base-ext']));
  });
});
