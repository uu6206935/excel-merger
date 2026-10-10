/**
 * The Excel-authored fixtures of the sheet copier (tests/fixtures/addsheets/*.xlsx, made with tests/golden/com/make-fixtures-addsheets.ps1 in a
 * real Excel): copied into a sample base, linted, compared with their source cell by cell through the reader, and checked for the features that
 * only Excel itself writes (x14 rules with inline dxf, sheet-scoped names, spills, sparklines, theme colours that differ from the base).
 * Real Excel opens the same outputs without repair: tests/golden/RESULTS.md (tests/golden/com/check-addsheets.ps1).
 */
import { describe, expect, it } from 'vitest';
import { quoteSheetName } from '../../src/core/formula/excelRef';
import { readBook } from '../../src/core/xlsx';
import { compareSheets, definedNamesOf, lintPackage, lintText, premergeCopies, readFiles, readFix, sheetsOf, type FileIn } from './addsheets-helpers';

/** a sheet reference prefix as Excel writes it (the names of these copies contain a hyphen: quoted) */
const q = (name: string) => `${quoteSheetName(name)}!`;
const fx = (n: string): FileIn => ({ bytes: readFix('tests/fixtures/addsheets', n), fileName: n });
const base = (): FileIn => ({ bytes: readFix('samples', '01_学生名簿.xlsx'), fileName: '01_学生名簿.xlsx' });

async function merge(n: string, extra = {}) {
  const res = await premergeCopies(base(), [fx(n)], { dropOtherSheets: true, ...extra });
  const out = await readBook(res.bytes, 'out.xlsx');
  const src = await readBook(readFix('tests/fixtures/addsheets', n), n);
  const files = await readFiles(res.bytes);
  const rename = new Map(res.layout.addedSheets!.map((s) => [s.original.toLowerCase(), s.name]));
  const part = (name: string) => files.get(sheetsOf(files).find((s) => s.name === name)!.part)!;
  return { res, out, src, files, rename, part };
}

describe('Excel-authored fixtures: every copy equals its source', () => {
  for (const n of ['ex-formats.xlsx', 'ex-theme-other.xlsx', 'ex-two-sheets.xlsx', 'ex-1904.xlsx']) {
    it(`${n}`, async () => {
      const { res, out, src, rename } = await merge(n);
      expect(lintText(await lintPackage(res.bytes))).toBe('');
      expect(res.layout.addedSheets!.length).toBe(src.sheets.length);
      for (const a of res.layout.addedSheets!) {
        const diffs = compareSheets(src, a.original, out, a.name, { rename, styles: true });
        expect(diffs.map((d) => `${a.name} ${d.where}: ${d.detail}`), a.name).toEqual([]);
      }
    });
  }
});

describe('what only Excel writes', () => {
  it('ex-formats: names, rules, validations, links, sparklines, spill, freeze, widths', async () => {
    const { res, part, files } = await merge('ex-formats.xlsx');
    const data = part('ex-formats_名簿');
    const names = await definedNamesOf(res.bytes);
    // the workbook names the sheet uses are sheet-scoped on its copy and point at the copies; the sheet's own names follow it
    expect(names.get('税率|ex-formats_名簿')).toBe(`${q('ex-formats_マスタ')}$B$1`);
    expect(names.get('候補|ex-formats_名簿')).toBe(`${q('ex-formats_マスタ')}$A$2:$A$5`);
    expect(names.get('基準日|ex-formats_名簿')).toBe(`${q('ex-formats_マスタ')}$D$1`);
    expect(names.get('局所|ex-formats_名簿')).toBe(`${q('ex-formats_名簿')}$C$3:$C$12`);
    expect(names.get('_xlnm.Print_Area|ex-formats_名簿')).toBe(`${q('ex-formats_名簿')}$A$1:$H$12`);
    expect([...names.keys()].filter((k) => k.endsWith('|ex-formats_マスタ'))).toEqual([]); // マスタ uses no name
    // x14 rules read the copy of マスタ; the inline dxf of the expression rule survives
    expect(data).toContain(`<xm:f>${q('ex-formats_マスタ')}$B$1&gt;0.05</xm:f><x14:dxf><font><b/><i val="0"/></font></x14:dxf>`);
    expect(data).toContain(`<x14:formula1><xm:f>${q('ex-formats_マスタ')}$A$2:$A$5</xm:f></x14:formula1>`);
    expect(data).toMatch(/<formula1>候補<\/formula1>/);
    expect(data).toContain(`<xm:f>${q('ex-formats_名簿')}C3:E3</xm:f>`);
    expect(data).toContain(`<hyperlink ref="B4" location="${q('ex-formats_マスタ')}A1" display="マスタへ"/>`);
    expect(data).toMatch(/<pane xSplit="2" ySplit="2" topLeftCell="C3" activePane="bottomRight" state="frozen"\/>/);
    expect(data).toMatch(/<col min="5" max="5" width="0" hidden="1" customWidth="1"\/>/);
    expect(data).toMatch(/<row r="9"[^>]*outlineLevel="1"/);
    expect(data).toMatch(/<tabColor theme="6"\/>/); // the same Office theme on both sides: theme colours stay theme colours
    // the dynamic-array spill (SORT) is one array formula with cm into the base's metadata
    expect(data).toMatch(/<c r="M3"[^>]* cm="1"><f t="array" ref="M3:M12">_xlfn\._xlws\.SORT\(C3:C12\)<\/f>/);
    expect(files.get('xl/metadata.xml')).toMatch(/XLDAPR/);
  });

  it('ex-theme-other: theme colours (cell fills / fonts / borders, rich text runs, a conditional format) are rgb colours of the source theme', async () => {
    const { res, out, part, files } = await merge('ex-theme-other.xlsx');
    const styles = files.get('xl/styles.xml')!;
    // accent 1 of the fixture's theme is C0504D; the base (Office 2023 theme) has another accent 1, so a theme="5" reference would be wrong
    expect(styles).toMatch(/rgb="FF[0-9A-F]{6}"/);
    const g = out.sheets.find((s) => s.name === 'ex-theme-other_配色')!;
    const fill1 = out.styles[g.rows[0]![0]!.s ?? 0].fill;
    expect(fill1).toBeTruthy();
    const x = part('ex-theme-other_配色');
    expect(x).not.toMatch(/<(?:fgColor|bgColor|color) theme=/);
    expect(res.report.items.some((i) => i.code === 'sheetCopyLimits')).toBe(false);
  });

  it('ex-two-sheets: sheet names with a space are quoted in formulas, a hidden sheet stays hidden, names that point at it follow it', async () => {
    const { res, out, files } = await merge('ex-two-sheets.xlsx');
    const sheets = sheetsOf(files);
    expect(sheets.map((s) => [s.name, s.state])).toEqual([['名簿', 'visible'], ['ex-two-sheets_Summary', 'visible'], ['ex-two-sheets_売上 2025', 'visible'], ['ex-two-sheets_作業', 'hidden']]);
    const s = out.sheets.find((x) => x.name === 'ex-two-sheets_Summary')!;
    expect(s.rows[1]![0]!.f).toBe("'ex-two-sheets_売上 2025'!A2");
    expect(s.rows[1]![2]!.f).toBe("VLOOKUP(A2,'ex-two-sheets_売上 2025'!$A$2:$C$6,3,FALSE)");
    expect(s.rows[1]![4]!.f).toBe('B2*(1+消費税)');
    const names = await definedNamesOf(res.bytes);
    expect(names.get('消費税|ex-two-sheets_Summary')).toBe(`${q('ex-two-sheets_売上 2025')}$E$1`);
    expect(names.get('隠し値|ex-two-sheets_Summary')).toBe(`${q('ex-two-sheets_作業')}$A$1`);
    expect(names.get('合計行|ex-two-sheets_売上 2025')).toBe(`${q('ex-two-sheets_売上 2025')}$B$2:$B$6`);
    // the dynamic-array formulas (XLOOKUP / LET / SORT spill) stay formulas
    expect(s.rows[2]![3]!.f).toMatch(/XLOOKUP/);
    expect(s.rows[3]![3]!.f).toMatch(/LET\(/);
    expect(s.rows[1]![7]!.dyn).toBe(true);
  });

  it('ex-1904: a 1904 workbook copied into the 1900 base keeps the same calendar dates (+1462 days in the file), with a warning', async () => {
    const { res, out, src, files, part } = await merge('ex-1904.xlsx');
    expect(src.date1904).toBe(true);
    expect(out.date1904).toBe(false);
    const stored = /<c r="A2"[^>]*><v>([\d.]+)<\/v>/.exec(part('ex-1904_日付'))![1];
    expect(Number(stored)).toBe(41000 + 1462);
    expect(res.report.items.some((i) => i.code === 'sheetCopyLimits' && /1904/.test(i.message))).toBe(true);
    // the time of day (0.25) and the duration (1.5 days: [h]:mm) are not dates: untouched
    expect(/<c r="C2"[^>]*><v>([\d.]+)<\/v>/.exec(part('ex-1904_日付'))![1]).toBe('0.25');
    expect(/<c r="D2"[^>]*><v>([\d.]+)<\/v>/.exec(part('ex-1904_日付'))![1]).toBe('1.5');
    void files;
  });
});
