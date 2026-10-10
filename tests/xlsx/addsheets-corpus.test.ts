/**
 * includeSourceSheets on the foreign corpus (tests/fixtures/foreign/corpus: Google Sheets, LibreOffice, POI, Open XML SDK, Mac Excel 2011 /
 * 1904, Excel 2007, Strict OOXML, openpyxl, XlsxWriter, pandas, ExcelJS, SheetJS ...): every file is the OTHER file of a merge into a sample
 * base. The output passes the package linter and every copied sheet equals its source (values, formulas with renamed references, resolved
 * styles, columns, rows, merges, rules, links); the base file is not changed.
 */
import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { readBook } from '../../src/core/xlsx';
import { REPO, compareSheets, lintPackage, lintText, premergeCopies, readFix, type FileIn } from './addsheets-helpers';

const DIR = path.join(REPO, 'tests/fixtures/foreign/corpus');
const files = fs.readdirSync(DIR).filter((f) => /\.(xlsx|xlsm)$/i.test(f)).sort();
const base = (): FileIn => ({ bytes: readFix('samples', '01_学生名簿.xlsx'), fileName: '01_学生名簿.xlsx' });

/** files whose package cannot be copied at all: (none expected) */
const REFUSED = new Set<string>();

describe('corpus: every file as the other file of a merge (copied, linted, compared)', () => {
  it('has the corpus', () => {
    expect(files.length).toBeGreaterThanOrEqual(60);
  });

  for (const f of files) {
    it(`${f}`, async () => {
      const bytes = new Uint8Array(fs.readFileSync(path.join(DIR, f)));
      const other: FileIn = { bytes, fileName: f };
      let res;
      try {
        res = await premergeCopies(base(), [other], { dropOtherSheets: true });
      } catch (e) {
        // a file the app cannot read / merge is refused with a Japanese message (the same as without copies)
        expect(REFUSED.has(f) || /[぀-ヿ一-鿿]/.test((e as Error).message), (e as Error).message).toBe(true);
        return;
      }
      expect(lintText(await lintPackage(res.bytes)), 'linter').toBe('');
      const out = await readBook(res.bytes, 'out.xlsx');
      const src = await readBook(bytes, f);
      expect(res.layout.addedSheets?.length, 'sheets copied').toBe(src.sheets.length);
      const rename = new Map(res.layout.addedSheets!.map((s) => [s.original.toLowerCase(), s.name]));
      const problems: string[] = [];
      for (const a of res.layout.addedSheets!) {
        // formulas that cannot stay formulas (other workbooks, tables ...) keep their cached value: only those sheets may differ in formula text
        const diffs = compareSheets(src, a.original, out, a.name, { rename, styles: true, allowValueFormulas: !a.complete, dateShift: 0, noRich: false });
        for (const d of diffs) problems.push(`${a.name} ${d.where}: ${d.detail}`);
      }
      expect(problems).toEqual([]);
    });
  }
});
