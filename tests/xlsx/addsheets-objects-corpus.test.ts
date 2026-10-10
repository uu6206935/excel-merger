/**
 * The object phase of the sheet copier on every package the repository has (task A2): the foreign corpus (Google Sheets, LibreOffice, POI,
 * Open XML SDK, Mac Excel 2011, Excel 2007, Strict OOXML, openpyxl, XlsxWriter, pandas, ExcelJS, SheetJS ...), the golden fixtures, the
 * Excel-authored fixtures of round 9, the samples: each one is the OTHER file of a merge. The output passes the linter of A1 plus the rules
 * of the object parts, every copied sheet owns the objects its source sheet owns (counts of charts / pictures / shapes / groups / connectors,
 * notes, tables, pivot tables; the things that are left out on purpose are counted separately), and the formulas of the copies read
 * the copies.
 */
import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { REPO, premergeCopies, type FileIn } from './addsheets-helpers';
import { lintObjects, lintObjectsText, objectsOf } from './addsheets-objects-helpers';

const dirs = ['tests/fixtures/foreign/corpus', 'tests/fixtures/golden', 'tests/fixtures/addsheets', 'tests/fixtures/addsheets-objects', 'tests/fixtures/dropsheets', 'samples'];
const files = dirs.flatMap((d) => fs.readdirSync(path.join(REPO, d)).filter((f) => /\.(xlsx|xlsm)$/i.test(f)).sort().map((f) => ({ dir: d, f })));
const base = (): FileIn => ({ bytes: new Uint8Array(fs.readFileSync(path.join(REPO, 'samples/02_面接評価.xlsx'))), fileName: '02_面接評価.xlsx' });

describe('every package of the repository as the other file of a merge, objects included', () => {
  it('has the packages', () => {
    expect(files.length).toBeGreaterThanOrEqual(90);
  });

  for (const { dir, f } of files) {
    it(`${dir}/${f}`, async () => {
      const bytes = new Uint8Array(fs.readFileSync(path.join(REPO, dir, f)));
      let res;
      try {
        res = await premergeCopies(base(), [{ bytes, fileName: f }], { dropOtherSheets: true });
      } catch (e) {
        // a file the app cannot read / merge is refused with a Japanese message (the same as without copies)
        expect(/[぀-ヿ一-鿿]/.test((e as Error).message), (e as Error).message).toBe(true);
        return;
      }
      expect(lintObjectsText(await lintObjects(res.bytes)), 'linter').toBe('');
      const from = await objectsOf(bytes);
      const to = await objectsOf(res.bytes);
      const problems: string[] = [];
      const strict = (await import('jszip')).default.loadAsync(bytes).then(async (z) => ((await z.file('xl/workbook.xml')?.async('string')) ?? '').includes('purl.oclc.org'));
      const isStrictSrc = await strict;
      // (the test helper reads plain packages only: for a source whose workbook uses prefixes / odd relationship targets it finds no sheet, then the source must not own objects)
      if (!from.length) {
        const zip = await (await import('jszip')).default.loadAsync(bytes);
        expect(Object.keys(zip.files).filter((n) => /^xl[\\/](drawings|charts|tables|comments|pivot)/i.test(n)), 'objects of a package the helper cannot read').toEqual([]);
        return;
      }
      for (const a of res.layout.addedSheets ?? []) {
        const s = from.find((x) => x.name === a.original);
        const d = to.find((x) => x.name === a.name);
        if (!s || !d) { problems.push(`${a.name}: missing`); continue; }
        // a sheet that cannot be copied at all stays as an empty sheet (reported); everything else keeps what it has
        if (a.complete === false && res.report.items.some((i) => i.code === 'sheetCopyLimits' && i.message.includes('空のシートになります') && i.message.includes(`「${a.original}」`))) continue;
        const left = { slicers: s.slicers, controls: s.vmlControls };
        // drawings are not copied between a Strict and a transitional package (their values differ): named in the warning, tested in addsheets-objects-fixtures
        const drawings = isStrictSrc ? 'skip' : 'copy';
        const same: Array<[string, number, number]> = [
          ['charts', s.charts, d.charts], ['pictures', s.pictures, d.pictures], ['shapes', s.shapes, d.shapes], ['connectors', s.connectors, d.connectors], ['groups', s.groups, d.groups],
          ['notes', s.notes, d.notes], ['tables', s.tables.length, d.tables.length], ['pivots', s.pivots.length, d.pivots.length],
        ];
        for (const [what, x, y] of same) if (x !== y && !(drawings === 'skip' && ['charts', 'pictures', 'shapes', 'connectors', 'groups'].includes(what))) problems.push(`${a.name}: ${what} ${x} -> ${y} (left out on purpose: ${JSON.stringify(left)})`);
        if (s.background !== d.background) problems.push(`${a.name}: background ${s.background} -> ${d.background}`);
        // formulas of the charts read the copies (their sheets exist: lintObjects), and they are the source's formulas with the sheet names renamed
        if (drawings === 'copy' && s.chartFormulas.length !== d.chartFormulas.length) problems.push(`${a.name}: chart formulas ${s.chartFormulas.length} -> ${d.chartFormulas.length}`);
      }
      expect(problems).toEqual([]);
    });
  }
});
