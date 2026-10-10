/**
 * includeSourceSheets on the app's own sample files (samples/*): every other file's sheets are copied after the merged sheet, named
 * "<file>_<sheet>", the package passes the linter and every copy equals its source cell by cell.
 */
import { describe, expect, it } from 'vitest';
import { readBook } from '../../src/core/xlsx';
import { lintPackage, lintText, plannedNames, premergeCopies, readFix, sheetNamesOf, compareSheets, type FileIn } from './addsheets-helpers';

const NAMES = ['01_学生名簿.xlsx', '02_面接評価.xlsx', '03_連絡先.csv', '04_適性検査.xls', '05_研修出欠.xlsx'];
const file = (n: string): FileIn => ({ bytes: readFix('samples', n), fileName: n });

describe('samples: includeSourceSheets', () => {
  it('copies every sheet of the other four files (order: merged, base sheets, then file by file)', async () => {
    const res = await premergeCopies(file(NAMES[0]), NAMES.slice(1).map(file), { dropOtherSheets: false });
    const names = await sheetNamesOf(res.bytes);
    const srcSheets = await Promise.all(NAMES.map(async (n) => (await readBook(readFix('samples', n), n)).sheets.map((s) => s.name)));
    const expected = plannedNames(['名簿', '集計', 'ピボット', 'マスタ'], NAMES.slice(1).map((n, i) => ({ fileName: n, sheets: srcSheets[i + 1] })));
    expect(names).toEqual(['名簿', '集計', 'ピボット', 'マスタ', ...expected]);
    expect(res.layout.addedSheets?.map((s) => s.name)).toEqual(expected);
    expect(res.report.addedSheets).toEqual(expected);
    expect(lintText(await lintPackage(res.bytes))).toBe('');
  });

  it('dropOtherSheets + includeSourceSheets: the merged sheet and the copies only', async () => {
    const res = await premergeCopies(file(NAMES[0]), NAMES.slice(1).map(file), { dropOtherSheets: true });
    const names = await sheetNamesOf(res.bytes);
    expect(names[0]).toBe('名簿');
    expect(names.length).toBe(1 + (res.layout.addedSheets?.length ?? 0));
    expect(names).not.toContain('集計');
    expect(lintText(await lintPackage(res.bytes))).toBe('');
  });

  it('every copy equals its source (values, formulas with renamed references, resolved styles)', async () => {
    const res = await premergeCopies(file(NAMES[0]), NAMES.slice(1).map(file), { dropOtherSheets: true });
    const out = await readBook(res.bytes, 'out.xlsx');
    for (const a of res.layout.addedSheets ?? []) {
      const src = await readBook(readFix('samples', a.file), a.file);
      const rename = new Map((res.layout.addedSheets ?? []).filter((x) => x.file === a.file).map((x) => [x.original.toLowerCase(), x.name]));
      const diffs = compareSheets(src, a.original, out, a.name, { rename, styles: true, allowValueFormulas: a.complete === false });
      expect(diffs.map((d) => `${a.name} ${d.where}: ${d.detail}`), a.name).toEqual([]);
    }
  });
});
