import { beforeAll, describe, expect, it } from 'vitest';
import type { SourceBook } from '../../src/core/types';
import { fidelityNotes, noteCounts } from '../../src/state/notes';
import type { FileEntry } from '../../src/state/types';
import { book } from './fixtures';

let base: SourceBook;
let other: SourceBook;
const entry = (id: string, b: SourceBook): FileEntry => ({ id, fileName: b.fileName, size: 1, color: '#000', status: 'ready', book: b });

beforeAll(async () => {
  base = await book('学生回答.xlsx', 's1');
  other = await book('職場回答.csv', 's2');
});

describe('警告 notes', () => {
  it('charts of the base: engine hint depends on Excel availability', () => {
    const withExcel = fidelityNotes([entry('s1', base)], 's1', true);
    expect(withExcel.find((n) => n.key === 'base-charts')?.text).toBe('グラフ 1 個: Excel本体モードで確実に保持されます');
    const without = fidelityNotes([entry('s1', base)], 's1', false);
    expect(without.find((n) => n.key === 'base-charts')?.text).toContain('Excel のある PC');
  });

  it('values-only inputs and merge report items', () => {
    const notes = fidelityNotes([entry('s1', base), entry('s2', other)], 's1', true, {
      mergedRows: 1,
      matchedRows: 1,
      baseOnlyRows: 0,
      appendedRows: 0,
      importedColumns: 1,
      formulasKept: 0,
      formulasToValues: 3,
      items: [{ severity: 'warning', code: 'x', message: '見出しが重複したため名前を変えました' }],
    });
    expect(notes.find((n) => n.key === 's2-values')?.text).toContain('値のみ取り込みます');
    expect(notes.find((n) => n.key === 'toValues')?.text).toContain('3 件');
    expect(notes.some((n) => n.text === '見出しが重複したため名前を変えました')).toBe(true);
    const c = noteCounts(notes);
    expect(c.warning).toBeGreaterThanOrEqual(2);
  });

  it('report items name the rows of the other file; export warnings are listed', () => {
    const notes = fidelityNotes(
      [entry('s1', base), entry('s2', other)],
      's1',
      true,
      {
        mergedRows: 1,
        matchedRows: 1,
        baseOnlyRows: 0,
        appendedRows: 0,
        importedColumns: 1,
        formulasKept: 0,
        formulasToValues: 0,
        items: [{ severity: 'info', code: 'x', message: '数式を値にしました', sourceId: 's2', otherRows: [12, 13] }],
      },
      { labels: { s2: '02_面接評価' }, exportWarnings: ['マクロは保持されませんでした'] },
    );
    expect(notes.some((n) => n.text === '数式を値にしました（02_面接評価 の 12行目・13行目）')).toBe(true);
    expect(notes.find((n) => n.key === 'export-w0')).toEqual({ key: 'export-w0', severity: 'warning', text: '出力: マクロは保持されませんでした' });
  });

  it('a values-only input the reader already warned about gets one note, not two', () => {
    const xls = { ...other, fileName: '04_適性検査.xls', format: 'xls' as const, warnings: ['xls形式のファイルは値（と表示形式）だけを読み込みました。'] };
    const notes = fidelityNotes([entry('s1', base), entry('s2', xls)], 's1', true);
    expect(notes.find((n) => n.key === 's2-values')).toBeUndefined();
    expect(notes.filter((n) => n.key.startsWith('s2-')).map((n) => n.text)).toEqual(['04_適性検査.xls: xls形式のファイルは値（と表示形式）だけを読み込みました。']);
  });

  it('a file without a key is listed first as a warning (wording before / after the merge)', () => {
    const files = [entry('s1', base), entry('s2', other)];
    const labels = { s2: '04_適性検査' };
    const before = fidelityNotes(files, 's1', true, null, { labels, excluded: ['s2'] });
    expect(before[0]).toMatchObject({ key: 'excluded-s2', severity: 'warning' });
    expect(before[0].text).toBe('「04_適性検査」は結合キーが見つからないため結合の対象外です（列は取り込まれません）。ファイルカードの「キーを選ぶ…」でキーを選ぶと結合に含まれます。');
    const after = fidelityNotes(files, 's1', true, null, { labels, excluded: ['s2'], merged: true });
    expect(after[0].text).toContain('列は取り込んでいません');
    expect(after[0].text).toContain('もう一度結合');
    // the file name is the fallback when no label is known
    expect(fidelityNotes(files, 's1', true, null, { excluded: ['s2'] })[0].text).toContain(`「${other.fileName}」`);
    expect(fidelityNotes(files, 's1', true, null, { labels, excluded: [] }).some((n) => n.key.startsWith('excluded-'))).toBe(false);
  });

  it('a values-only base is a warning; no files -> no notes', () => {
    const csvBase = { ...other, fileName: 'x.csv' };
    expect(fidelityNotes([entry('s2', csvBase)], 's2', true)[0]).toMatchObject({ severity: 'warning', key: 'base-values' });
    expect(fidelityNotes([], null, true)).toEqual([]);
  });
});
