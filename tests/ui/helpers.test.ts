import { describe, expect, it } from 'vitest';
import type { SourceBook, TableRegion } from '../../src/core/types';
import {
  SOURCE_COLORS,
  colLetter,
  dateStamp,
  encodingLabel,
  exportFileName,
  formatSize,
  headerRowChoices,
  isAcceptedFileName,
  isExcelLockFile,
  lastRowBounds,
  normalizeExportName,
  pickColor,
  readErrorMessage,
  opWarningsMessage,
  refusalMessage,
  regionRowCount,
  regionWithHeaderRow,
  regionWithLastRow,
  sanitizeFileName,
  undoRedoAction,
} from '../../src/state/helpers';

const key = (k: string, mods: Partial<{ ctrl: boolean; shift: boolean; alt: boolean; meta: boolean }> = {}) => ({
  key: k,
  ctrlKey: !!mods.ctrl,
  shiftKey: !!mods.shift,
  altKey: !!mods.alt,
  metaKey: !!mods.meta,
});

describe('file names', () => {
  it('accepts the supported extensions only (case-insensitive)', () => {
    for (const n of ['a.xlsx', 'b.XLSM', '名簿.xls', 'c.xlsb', 'x.csv', 'y.ods']) expect(isAcceptedFileName(n)).toBe(true);
    for (const n of ['a.txt', 'b.pdf', 'noext', 'a.xlsx.pdf']) expect(isAcceptedFileName(n)).toBe(false);
  });
  it('recognises Excel lock files', () => {
    expect(isExcelLockFile('~$名簿.xlsx')).toBe(true);
    expect(isExcelLockFile('名簿.xlsx')).toBe(false);
  });
  it('builds the default export name <base>_統合_<yyyyMMdd>.xlsx', () => {
    const d = new Date(2026, 8, 30);
    expect(dateStamp(d)).toBe('20260930');
    expect(exportFileName('学生回答.xlsx', 'xlsx', d)).toBe('学生回答_統合_20260930.xlsx');
    expect(exportFileName('マクロ付き.xlsm', 'xlsm', d)).toBe('マクロ付き_統合_20260930.xlsm');
    expect(exportFileName('a:b?.csv', 'xlsx', d)).toBe('a_b__統合_20260930.xlsx');
  });
  it('normalises typed names: forbidden characters, trailing dots, extension forced', () => {
    expect(normalizeExportName('結果', 'xlsx')).toBe('結果.xlsx');
    expect(normalizeExportName('結果.xlsx', 'xlsx')).toBe('結果.xlsx');
    expect(normalizeExportName('結果.xls', 'xlsx')).toBe('結果.xlsx');
    expect(normalizeExportName('結果.xlsx', 'xlsm')).toBe('結果.xlsm');
    expect(normalizeExportName('  a/b*c. ', 'xlsx')).toBe('a_b_c.xlsx');
    expect(normalizeExportName('', 'xlsx')).toBe('統合結果.xlsx');
    expect(sanitizeFileName('<>')).toBe('__');
  });
  it('formats sizes', () => {
    expect(formatSize(512)).toBe('512 B');
    expect(formatSize(32600)).toBe('31.8 KB');
    expect(formatSize(5 * 1024 * 1024)).toBe('5.0 MB');
    expect(formatSize(-1)).toBe('');
  });
  it('turns library errors into Japanese card messages', () => {
    expect(readErrorMessage('a.xlsx', new Error('File is password-protected'))).toContain('パスワード');
    expect(readErrorMessage('a.xlsx', new Error('シートがありません'))).toBe('a.xlsx: シートがありません');
    const m = readErrorMessage('a.xlsx', new Error('bad zip'));
    expect(m).toMatch(/^a\.xlsx: 読み込めませんでした/);
    expect(m).toContain('bad zip');
    expect(readErrorMessage('a.xlsx', 'oops')).toContain('oops');
  });
});

describe('colours', () => {
  it('picks the first unused palette colour and cycles when exhausted', () => {
    expect(pickColor([])).toBe(SOURCE_COLORS[0]);
    expect(pickColor([SOURCE_COLORS[0]])).toBe(SOURCE_COLORS[1]);
    expect(pickColor([...SOURCE_COLORS])).toBe(SOURCE_COLORS[SOURCE_COLORS.length % SOURCE_COLORS.length]);
  });
});

describe('columns', () => {
  it('colLetter', () => {
    expect([1, 26, 27, 52, 53, 702, 703, 16384].map(colLetter)).toEqual(['A', 'Z', 'AA', 'AZ', 'BA', 'ZZ', 'AAA', 'XFD']);
  });
});

describe('keyboard', () => {
  it('maps Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z', () => {
    expect(undoRedoAction(key('z', { ctrl: true }))).toBe('undo');
    expect(undoRedoAction(key('Z', { ctrl: true, shift: true }))).toBe('redo');
    expect(undoRedoAction(key('y', { ctrl: true }))).toBe('redo');
    expect(undoRedoAction(key('z', { meta: true }))).toBe('undo');
    expect(undoRedoAction(key('z'))).toBeNull();
    expect(undoRedoAction(key('z', { ctrl: true, alt: true }))).toBeNull();
    expect(undoRedoAction(key('y', { ctrl: true, shift: true }))).toBeNull();
  });
});

function sheetBook(rows: Array<Array<string | number | null>>): SourceBook {
  return {
    id: 'b',
    fileName: 'b.xlsx',
    size: 1,
    format: 'xlsx',
    bytes: new Uint8Array(),
    fidelity: 'full',
    styles: [{}],
    sheets: [
      {
        name: 'S',
        state: 'visible',
        maxRow: rows.length,
        maxCol: Math.max(...rows.map((r) => r.length)),
        rows: rows.map((r) => r.map((v) => (v === null ? undefined : typeof v === 'number' ? { t: 'n' as const, v } : { t: 's' as const, v }))),
        colProps: [],
        rowProps: [],
        features: { merges: [], rules: [], tables: [], arrays: [], hyperlinks: [] },
      },
    ],
    names: [],
    features: { charts: 0, pivotTables: 0, vba: false, drawings: 0, slicers: 0, sparklineGroups: 0, externalLinks: 0, threadedComments: 0, formControls: 0, tables: 0, images: 0 },
    date1904: false,
    warnings: [],
  };
}

describe('regions', () => {
  const book = sheetBook([
    ['名簿', null, null],
    [null, null, null],
    ['ID', '氏名', '点'],
    [1, '青木', 70],
    [2, '石川', 80],
    [null, null, null],
    [3, '上田', 90],
    [null, null, null],
    [null, null, null],
    ['メモ', null, null],
  ]);
  const detected: TableRegion = { sheet: 'S', headerRow: 3, firstDataRow: 4, lastDataRow: 7, firstCol: 1, lastCol: 3 };

  it('keeps the detected region for the detected header row', () => {
    expect(regionWithHeaderRow(book, detected, 3)).toEqual(detected);
  });
  it('recomputes columns from the chosen header row and keeps the detected last row below it', () => {
    const r = regionWithHeaderRow(book, detected, 1);
    expect(r).toEqual({ sheet: 'S', headerRow: 1, firstDataRow: 2, lastDataRow: 7, firstCol: 1, lastCol: 1 });
  });
  it('scans for data (stopping at two blank rows) when nothing was detected below', () => {
    const none: TableRegion = { sheet: 'S', headerRow: 0, firstDataRow: 1, lastDataRow: 0, firstCol: 1, lastCol: 3 };
    const r = regionWithHeaderRow(book, none, 3);
    expect(r.lastDataRow).toBe(7);
    expect(regionRowCount(r)).toBe(4);
  });
  it('last data row: bounds and the local fallback clamp to the sheet', () => {
    expect(lastRowBounds(book.sheets[0], detected)).toEqual({ min: 4, max: 10 });
    expect(lastRowBounds(undefined, detected)).toBeNull();
    expect(regionWithLastRow(detected, 5, 10)).toEqual({ ...detected, lastDataRow: 5 });
    expect(regionWithLastRow(detected, 99, 10).lastDataRow).toBe(10);
    expect(regionWithLastRow(detected, 1, 10).lastDataRow).toBe(4);
  });
  it('header row choices cover the first rows and the current one', () => {
    expect(headerRowChoices(book.sheets[0])).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
    expect(headerRowChoices(book.sheets[0], 40)).toContain(40);
    expect(headerRowChoices(undefined)).toEqual([1]);
  });
});

describe('messages', () => {
  it('refusals become one short sentence; the full reason stays as the detail', () => {
    const r = refusalMessage('結合されたセル（A1:K1）の一部の列だけを移動することはできません。結合を解除するか、結合範囲の列をまとめて移動してください。', '列を移動できませんでした');
    expect(r.text).toBe('結合されたセル（A1:K1）の一部の列だけを移動できません');
    expect(r.detail).toContain('結合を解除するか');
    const t = refusalMessage('テーブル「表1」の列をテーブルの外へ移動することはできません。（Excel と同じ制限です）', 'x');
    expect(t.text).toBe('テーブル「表1」の列をテーブルの外へ移動できません');
    expect(refusalMessage('テーブルの一部は移動できません', 'x')).toEqual({ text: 'テーブルの一部は移動できません' });
    expect(refusalMessage('', '列を移動できませんでした')).toEqual({ text: '列を移動できませんでした' });
    expect(refusalMessage(undefined, 'x')).toEqual({ text: 'x' });
  });
  it('warnings of a successful op: a short toast, every warning in the detail', () => {
    const w = opWarningsMessage([
      'タイトル行の結合（A1:K1）を解除し、「選択範囲内で中央」にしました。',
      'VLOOKUP の列番号は数値のため自動では変わりません。範囲 名簿!$A$4:$K$63 の列が移動するため、別の列の値を返す可能性があります。',
      '集計!B28: =VLOOKUP("000108",名簿!$A$4:$K$63,2,FALSE)',
    ]);
    expect(w?.text).toBe('タイトル行の結合（A1:K1）を解除し、「選択範囲内で中央」にしました。（ほか 2 件の注意）');
    expect(w?.detail?.split('\n')).toHaveLength(3);
    expect(opWarningsMessage(['列幅を変更しました。'])).toEqual({ text: '列幅を変更しました。' });
    expect(opWarningsMessage(['一文目。二文目。'])).toEqual({ text: '一文目。', detail: '一文目。二文目。' });
    expect(opWarningsMessage([])).toBeNull();
    expect(opWarningsMessage(undefined)).toBeNull();
    expect(opWarningsMessage(['  ', ''])).toBeNull();
  });
  it('encoding names for the CSV card tag', () => {
    expect(encodingLabel('Shift_JIS')).toBe('Shift_JIS');
    expect(encodingLabel('SJIS')).toBe('Shift_JIS');
    expect(encodingLabel('UTF8')).toBe('UTF-8');
    expect(encodingLabel('UTF-16LE')).toBe('UTF-16');
    expect(encodingLabel('ISO-2022-JP')).toBe('JIS');
    expect(encodingLabel('windows-1252')).toBe('windows-1252');
    expect(encodingLabel(undefined)).toBeNull();
    expect(encodingLabel(' ')).toBeNull();
  });
});
