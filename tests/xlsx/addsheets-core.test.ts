/**
 * The core of the sheet copier on the hand-made "kitchen sink" source (addsheets-kitchen.ts): names, order, states, the package linter, a cell by
 * cell comparison with the source, sheet-level parts (views, columns, rules, links, page setup), shared strings / styles / dxfs, defined names.
 */
import { describe, expect, it } from 'vitest';
import { excelWidthToPx, pxToExcelWidth } from '../../src/core/display/styles';
import { readBook } from '../../src/core/xlsx';
import { buildPackage, cn, fml, grid, inl, num, themeXml, OFFICE_COLORS } from './addsheets-build';
import { KITCHEN_NAME, kitchenSource } from './addsheets-kitchen';
import { compareSheets, definedNamesOf, lintPackage, lintText, premergeCopies, readFiles, sheetsOf, type FileIn } from './addsheets-helpers';

const BASE_REGION = { sheet: '名簿', headerRow: 1, firstDataRow: 2, lastDataRow: 4, firstCol: 1, lastCol: 2 };

/** a small base workbook: the merged sheet 名簿 at position `at` of [集計, 名簿] */
async function baseBook(opts: { mergedFirst?: boolean; date1904?: boolean; names?: string } = {}): Promise<FileIn> {
  const merged = { name: '名簿', data: grid([['ID', '氏名'], ['A001', '山田'], ['A002', '鈴木'], ['A003', '佐藤']]) };
  const other = { name: '集計', data: grid([['n', { f: 'COUNTA(名簿!A2:A4)', v: 3 }]]) };
  const bytes = await buildPackage({
    sheets: opts.mergedFirst ? [merged, other] : [other, merged],
    theme: themeXml(OFFICE_COLORS),
    date1904: opts.date1904,
    names: opts.names ?? '<definedName name="Mine">名簿!$A$2</definedName>',
    app: true,
  });
  return { bytes, fileName: '名簿.xlsx', region: BASE_REGION };
}

async function kitchen(): Promise<FileIn> {
  return { bytes: await kitchenSource(), fileName: KITCHEN_NAME, sheet: 'Data' };
}

describe('kitchen sink source copied into a small base', () => {
  it('names, order and tab states (merged sheet first, base sheets, then the source sheets in workbook order)', async () => {
    const res = await premergeCopies(await baseBook(), [await kitchen()], { dropOtherSheets: false });
    const f = await readFiles(res.bytes);
    const sheets = sheetsOf(f);
    expect(sheets.map((s) => s.name)).toEqual(['名簿', '集計', 'キッチン_Data', 'キッチン_Calc', 'キッチン_Bad', 'キッチン_Lookup', "キッチン_Q'uote Sheet", 'キッチン_Hidden', 'キッチン_VHidden']);
    expect(sheets.map((s) => s.state)).toEqual(['visible', 'visible', 'visible', 'visible', 'visible', 'visible', 'visible', 'hidden', 'veryHidden']);
    expect(res.layout.addedSheets?.map((s) => [s.sourceId, s.original, s.name, s.suffixed])).toEqual([
      ['s2', 'Data', 'キッチン_Data', false], ['s2', 'Calc', 'キッチン_Calc', false], ['s2', 'Bad', 'キッチン_Bad', false], ['s2', 'Lookup', 'キッチン_Lookup', false],
      ['s2', "Q'uote Sheet", "キッチン_Q'uote Sheet", false], ['s2', 'Hidden', 'キッチン_Hidden', false], ['s2', 'VHidden', 'キッチン_VHidden', false],
    ]);
    expect(res.report.addedSheets).toEqual(res.layout.addedSheets?.map((s) => s.name));
    expect(res.report.items.find((i) => i.code === 'sheetsCopied')?.message).toBe('他のファイルのシート 7 枚をコピーしました（キッチン_Data ほか）');
    expect(lintText(await lintPackage(res.bytes))).toBe('');
  });

  it('the merged sheet moves to the first tab when the base has it elsewhere (names, active tab, app.xml follow)', async () => {
    const res = await premergeCopies(await baseBook({ mergedFirst: false }), [await kitchen()], { dropOtherSheets: false });
    const f = await readFiles(res.bytes);
    expect(sheetsOf(f).slice(0, 2).map((s) => s.name)).toEqual(['名簿', '集計']);
    const names = await definedNamesOf(res.bytes);
    expect(names.get('Mine|')).toBe('名簿!$A$2');
    expect(f.get('xl/workbook.xml')).not.toMatch(/activeTab=/);
    expect(lintText(await lintPackage(res.bytes))).toBe('');
    // only the merged sheet is the selected tab
    const part = (n: string) => f.get(sheetsOf(f).find((s) => s.name === n)!.part)!;
    expect(part('名簿')).toMatch(/tabSelected="1"/);
    expect(part('集計')).not.toMatch(/tabSelected/);
    expect(part('キッチン_Data')).not.toMatch(/tabSelected/);
  });

  it('dropOtherSheets: the base keeps only the merged sheet, the copies are named after the files', async () => {
    const res = await premergeCopies(await baseBook(), [await kitchen()], { dropOtherSheets: true });
    const names = sheetsOf(await readFiles(res.bytes)).map((s) => s.name);
    expect(names.slice(0, 2)).toEqual(['名簿', 'キッチン_Data']);
    expect(names).not.toContain('集計');
    expect(lintText(await lintPackage(res.bytes))).toBe('');
  });

  it('every copied sheet equals its source cell by cell (renamed references, resolved styles)', async () => {
    const src = await readBook(await kitchenSource(), KITCHEN_NAME);
    const res = await premergeCopies(await baseBook(), [await kitchen()], { dropOtherSheets: true });
    const out = await readBook(res.bytes, 'out.xlsx');
    const rename = new Map(res.layout.addedSheets!.map((s) => [s.original.toLowerCase(), s.name]));
    for (const a of res.layout.addedSheets!) {
      const diffs = compareSheets(src, a.original, out, a.name, { rename, styles: true, allowValueFormulas: a.original === 'Bad' || a.original === 'Calc' || a.original === 'Data' });
      expect(diffs.map((d) => `${a.name} ${d.where}: ${d.detail}`), a.name).toEqual([]);
    }
  });
});

describe('formulas of the copies', () => {
  const read = async () => {
    const res = await premergeCopies(await baseBook(), [await kitchen()], { dropOtherSheets: true });
    const out = await readBook(res.bytes, 'out.xlsx');
    const sheet = (n: string) => out.sheets.find((s) => s.name === n)!;
    const cell = (n: string, ref: string) => {
      const m = /^([A-Z]+)(\d+)$/.exec(ref)!;
      let c = 0;
      for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64;
      return sheet(n).rows[Number(m[2]) - 1]?.[c - 1];
    };
    return { res, out, cell };
  };

  it('references to the other sheets of the same file follow the renamed copies; the file never reads the merged sheet or the base', async () => {
    const { cell } = await read();
    expect(cell('キッチン_Data', 'I2')!.f).toBe('キッチン_Calc!B2+キッチン_Lookup!A2');
    expect(cell('キッチン_Calc', 'B2')!.f).toBe('キッチン_Data!C2+1');
    expect(cell('キッチン_Calc', 'C3')!.f).toBe("'キッチン_Q''uote Sheet'!A1+1");
    expect(cell('キッチン_Q\'uote Sheet', 'B1')!.f).toBe('キッチン_Calc!A1+A1');
    expect(cell('キッチン_Data', 'H2')!.f).toBe('SUM($C$2:$C$11)');
  });

  it('shared, array and dynamic-array formulas stay what they are', async () => {
    const { cell } = await read();
    expect(cell('キッチン_Data', 'J2')!.f).toBe('C2*2');
    expect(cell('キッチン_Data', 'J5')!.f).toBe('C5*2'); // a shared dependent became a plain formula
    expect(cell('キッチン_Data', 'K2')!.arr).toBe('K2:K3');
    expect(cell('キッチン_Data', 'L2')!.dyn).toBe(true);
    expect(cell('キッチン_Data', 'L3')!.v).toBe(2);
  });

  it('what cannot stay a formula keeps its cached value (external, 3D, call-out, uncopied table, unusable names, missing sheet)', async () => {
    const { cell, res } = await read();
    for (const [ref, v] of [['A1', 5], ['B1', 111], ['A2', 'web'], ['B2', 7], ['D2', 6], ['A3', 9], ['A4', 4]] as const) {
      const c = cell('キッチン_Bad', ref)!;
      expect(c.f, ref).toBeUndefined();
      expect(c.v, ref).toBe(v);
    }
    // what can: names (copied as sheet-scoped names), quoted sheet names, ranges on other copies, INDIRECT / HYPERLINK strings
    expect(cell('キッチン_Bad', 'C1')!.f).toBe('HYPERLINK("#キッチン_Data!A1","go")');
    expect(cell('キッチン_Bad', 'D1')!.f).toBe('INDIRECT("キッチン_Calc!A1")');
    expect(cell('キッチン_Bad', 'C2')!.f).toBe('Rate*3');
    expect(cell('キッチン_Bad', 'B3')!.f).toBe("'キッチン_Q''uote Sheet'!A1");
    expect(cell('キッチン_Bad', 'C3')!.f).toBe('SUM(キッチン_Calc!A1:A3)');
    expect(cell('キッチン_Bad', 'B4')!.f).toBe('Chain');
    const warn = res.report.items.find((i) => i.code === 'sheetCopyLimits' && i.message.includes('キッチン_Bad'))!;
    expect(warn.severity).toBe('warning');
    expect(warn.message).toMatch(/他のブック・外部を参照または呼び出す数式/);
    expect(warn.message).toMatch(/複数シートにまたがる参照/);
        expect(warn.message).toMatch(/テーブル参照の数式/);
  });

  it('defined names: sheet-scoped names of the source, the workbook names a sheet uses (also through other names) become sheet-scoped on each copy', async () => {
    const res = await premergeCopies(await baseBook(), [await kitchen()], { dropOtherSheets: true });
    const names = await definedNamesOf(res.bytes);
    expect(names.get('Mine|')).toBe('名簿!$A$2'); // the base's own names are untouched
    expect(names.get('Local|キッチン_Data')).toBe('キッチン_Data!$C$2:$C$11');
    expect(names.get('_xlnm.Print_Area|キッチン_Data')).toBe('キッチン_Data!$A$1:$H$12');
    expect(names.get('_xlnm.Print_Titles|キッチン_Data')).toBe('キッチン_Data!$1:$1');
    expect(names.get('_xlnm._FilterDatabase|キッチン_Data')).toBe('キッチン_Data!$A$1:$M$11');
    // Data uses Choices (validation) and, through the formulas of Calc, Rate and Chain
    expect(names.get('Choices|キッチン_Data')).toBe('キッチン_Calc!$A$1:$A$5');
    expect(names.get('Rate|キッチン_Calc')).toBe('キッチン_Calc!$A$1');
    expect(names.get('Chain|キッチン_Calc')).toBe('Rate*2');
    expect(names.get('Shadow|キッチン_Calc')).toBe('キッチン_Calc!$A$2'); // the sheet's own name hides the workbook name
    expect(names.get('Rate|キッチン_Bad')).toBe('キッチン_Calc!$A$1');
    expect(names.get('Chain|キッチン_Bad')).toBe('Rate*2');
    expect(names.get('Rate|キッチン_Lookup')).toBe('キッチン_Calc!$A$1');
    // never copied: unused workbook names, macro names, names into other workbooks, and no workbook-scoped name is added to the base
    for (const k of names.keys()) expect(k, k).not.toMatch(/^(Orphan|MacroFn|ExtName)\|/);
    expect(names.get('Unused2|キッチン_Bad')).toBe('Unused'); // used by a formula: copied with the name it needs
    expect(names.get('Unused|キッチン_Bad')).toBe('キッチン_Calc!$C$3');
    for (const k of names.keys()) if (!k.startsWith('Mine|')) expect(k.endsWith('|'), k).toBe(false);
  });

  it('a named formula that cannot be copied makes the formulas that use it values', async () => {
    const { cell } = await read();
    expect(cell('キッチン_Bad', 'D2')!.f).toBeUndefined(); // ExtName
    expect(cell('キッチン_Bad', 'A4')!.f).toBeUndefined(); // MacroFn
    expect(cell('キッチン_Bad', 'C4')!.f).toBe('Unused2'); // a fine name that needs another fine name: stays a formula
  });
});

describe('sheet-level parts of a copy', () => {
  it('views / columns / rows / merges / filters / page setup', async () => {
    const res = await premergeCopies(await baseBook(), [await kitchen()], { dropOtherSheets: true });
    const f = await readFiles(res.bytes);
    const x = f.get(sheetsOf(f).find((s) => s.name === 'キッチン_Data')!.part)!;
    expect(x).not.toMatch(/xr:uid=/);
    expect(x).not.toMatch(/codeName=/);
    expect(x).not.toMatch(/tabSelected/);
    expect(x).not.toMatch(/pivotSelection/);
    expect(x).toMatch(/<tabColor rgb="FF[0-9A-F]{6}"\/>/); // theme colour 5 of the OTHER theme became an rgb colour
    expect(x).toMatch(/<outlinePr summaryBelow="0"\/>/);
    expect(x).toMatch(/<pageSetUpPr fitToPage="1"\/>/);
    expect(x).toMatch(/zoomScale="85"/);
    expect(x).toMatch(/<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"\/>/);
    // the source's default font is Calibri (7 px per digit), the base's 游ゴシック (8 px): the columns keep their pixel widths
    const w = (chars: number) => Math.round(pxToExcelWidth(excelWidthToPx(chars, 7), 8) * 1e6) / 1e6;
    expect(x).toContain(`<col min="2" max="3" width="${w(14)}" customWidth="1"/>`);
    expect(x).toContain(`<col min="4" max="4" width="${w(9)}" hidden="1" customWidth="1"/>`);
    expect(x).toContain(`<col min="5" max="5" width="${w(12)}" outlineLevel="1" customWidth="1"/>`);
    // columns without a width follow: 8 characters of Calibri = 64 px = 8 characters of 游ゴシック
    expect(x).toContain(`<sheetFormatPr defaultRowHeight="18.75" defaultColWidth="${Math.round(pxToExcelWidth(64, 8) * 1e6) / 1e6}"/>`);
    expect(x).toMatch(/<row r="1" ht="24" customHeight="1">/);
    expect(x).toMatch(/<row r="5" ht="30" customHeight="1" hidden="1">/);
    expect(x).toMatch(/<mergeCells count="1"><mergeCell ref="A12:C12"\/><\/mergeCells>/);
    expect(x).toMatch(/<autoFilter ref="A1:M11">/);
    expect(x).toMatch(/<sortState ref="A2:M11">/);
    expect(x).toMatch(/<sheetProtection sheet="1" objects="1" scenarios="1"\/>/);
    expect(x).toMatch(/<pageSetup paperSize="9" orientation="landscape"\/>/); // the printer settings part is not copied
    expect(x).not.toMatch(/printerSettings/);
    expect(x).toMatch(/<headerFooter><oddHeader>&amp;CKitchen<\/oddHeader><\/headerFooter>/);
    expect(x).toMatch(/<rowBreaks count="1" manualBreakCount="1"><brk id="8" max="16383" man="1"\/><\/rowBreaks>/);
    expect(x).toMatch(/<ignoredErrors><ignoredError sqref="I2:I11" formula="1"\/><\/ignoredErrors>/);
    expect(x).toMatch(/<phoneticPr fontId="\d+" type="Hiragana"\/>/);
    expect([...f.keys()].some((k) => /printerSettings/i.test(k))).toBe(false);
  });

  it('hyperlinks: the external target gets a new relationship, the location follows the renamed sheet', async () => {
    const res = await premergeCopies(await baseBook(), [await kitchen()], { dropOtherSheets: true });
    const f = await readFiles(res.bytes);
    const sh = sheetsOf(f).find((s) => s.name === 'キッチン_Data')!;
    const x = f.get(sh.part)!;
    const rels = f.get(sh.part.replace('worksheets/', 'worksheets/_rels/') + '.rels')!;
    expect(rels).toMatch(/Target="https:\/\/example\.com\/a\?b=1&amp;c=2" TargetMode="External"/);
    expect(x).toMatch(/<hyperlink ref="A2" r:id="rId1" tooltip="site"\/>/);
    expect(x).toMatch(/<hyperlink ref="A3" location="キッチン_Calc!B2" display="to calc"\/>/);
    expect(x).toMatch(/<hyperlink ref="A4" location="キッチン_Lookup!A1"\/>/);
  });

  it('conditional formats and validations: dxfs remapped into the base, formulas renamed, x14 links with fresh GUIDs, no xr:uid', async () => {
    const res = await premergeCopies(await baseBook(), [await kitchen()], { dropOtherSheets: true });
    const f = await readFiles(res.bytes);
    const x = f.get(sheetsOf(f).find((s) => s.name === 'キッチン_Data')!.part)!;
    const styles = f.get('xl/styles.xml')!;
    expect(x).toMatch(/<formula>キッチン_Calc!\$A\$1&gt;C2<\/formula>/);
    expect(x).toMatch(/<formula1>キッチン_Calc!\$A\$1:\$A\$5<\/formula1>/);
    expect(x).toMatch(/<formula1>"a,b,c"<\/formula1>/);
    expect(x).toMatch(/<formula1>Choices<\/formula1>/);
    expect(x).toMatch(/<xm:f>キッチン_Calc!\$B\$1:\$B\$3<\/xm:f>/);
    expect(x).toMatch(/<xm:f>キッチン_Calc!A2:C2<\/xm:f>/);
    expect(x).toMatch(/<xm:f>キッチン_Data!C2:C11<\/xm:f>/);
    expect(x).not.toMatch(/xr:uid/);
    expect(x).not.toMatch(/unknownThing/);
    // the base had no dxfs: the three of the source were appended (numFmt of the third re-registered)
    expect((styles.match(/<dxf>/g) ?? []).length).toBe(3);
    expect(styles).toMatch(/<numFmt numFmtId="\d+" formatCode="0\.00%"\/>/);
    const dxfIds = [...x.matchAll(/\sdxfId="(\d+)"/g)].map((m) => Number(m[1]));
    expect(dxfIds.sort()).toEqual([0, 1, 2]);
    // the link between the classic data bar and its x14 twin is a fresh GUID (the same on both sides)
    const classic = /<x14:id>(\{[^}]+\})<\/x14:id>/.exec(x)?.[1];
    const twin = /<x14:cfRule type="dataBar" id="(\{[^}]+\})"/.exec(x)?.[1];
    expect(classic).toBeTruthy();
    expect(classic).toBe(twin);
    expect(classic).not.toBe('{11111111-2222-4333-8444-555555555555}');
    // the data-bar colours of the other theme are rgb now
    expect(x).toMatch(/<x14:borderColor rgb="FF[0-9A-F]{6}"\/>/);
    expect(x).toMatch(/<x14:colorSeries rgb="FF[0-9A-F]{6}"\/>/);
    expect(x).toMatch(/<color rgb="FF[0-9A-F]{6}"\/><\/dataBar>/);
  });

  it('shared strings: plain and rich (runs, phonetic data, theme colours of the other theme) land in the base sst', async () => {
    const res = await premergeCopies(await baseBook(), [await kitchen()], { dropOtherSheets: true });
    const f = await readFiles(res.bytes);
    const sst = f.get('xl/sharedStrings.xml')!;
    expect(sst).toMatch(/<si><r><rPr><b\/><sz val="11"\/><color rgb="FF[0-9A-F]{6}"\/><rFont val="Calibri"\/><family val="2"\/><\/rPr><t>太<\/t><\/r>/);
    expect(sst).toMatch(/<rPh sb="0" eb="1"><t>タ<\/t><\/rPh>/);
    const out = await readBook(res.bytes, 'o.xlsx');
    const b2 = out.sheets.find((s) => s.name === 'キッチン_Data')!.rows[1]![1]!;
    expect(b2.v).toBe('太 郎');
    expect(b2.rich?.map((r) => r.text)).toEqual(['太', ' 郎']);
    const a13 = out.sheets.find((s) => s.name === 'キッチン_Data')!.rows[0]![0]!;
    expect(a13.v).toBe('ID');
  });
});

describe('the base is not changed', () => {
  it('merged sheet cells, names and the base sheets stay as they were', async () => {
    const base = await baseBook({ mergedFirst: true });
    const plain = await premergeCopies(base, [], { includeSourceSheets: false });
    const withCopies = await premergeCopies(base, [await kitchen()], { dropOtherSheets: false });
    const a = await readBook(plain.bytes, 'a.xlsx');
    const b = await readBook(withCopies.bytes, 'b.xlsx');
    for (const n of ['名簿', '集計']) {
      expect(compareSheets(a, n, b, n, { rename: new Map(), styles: false })).toEqual([]);
    }
    expect(await definedNamesOf(withCopies.bytes)).toEqual(expect.objectContaining({}));
    const nm = await definedNamesOf(withCopies.bytes);
    expect(nm.get('Mine|')).toBe('名簿!$A$2');
  });

  it('includeSourceSheets false / undefined = the old behaviour (no extra sheet, no report field)', async () => {
    const base = await baseBook();
    const off = await premergeCopies(base, [await kitchen()], { includeSourceSheets: false });
    const undef = await premergeCopies(base, [await kitchen()], { includeSourceSheets: undefined });
    for (const r of [off, undef]) {
      expect((await sheetsOf(await readFiles(r.bytes))).map((s) => s.name)).toEqual(['集計', '名簿']);
      expect(r.layout.addedSheets).toBeUndefined();
      expect(r.report.addedSheets).toBeUndefined();
    }
  });
});

void cn; void fml; void inl; void num;
