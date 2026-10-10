/**
 * Cell values, number formats, date systems, names and counts of the sheet copier on small hand-made workbooks:
 *  - 1904 <-> 1900 date systems (the reader's rule: dates shift by 1462 days, times of day < 1 and durations never) + the warning,
 *  - number formats whose custom ids collide between the base and the source, theme / palette conversion, the 64,000-format limit,
 *  - values-only sources (csv / xls), an xlsm base, strict OOXML on either side,
 *  - naming: 3 files with the same name, 31-character shortening, forbidden characters, names that collide with the base's sheets,
 *  - empty sheets, formatted-but-empty cells, odd strings.
 */
import { describe, expect, it } from 'vitest';
import { readBook } from '../../src/core/xlsx';
import { MAIN_NS, PLAIN_STYLES, OFFICE_COLORS, OTHER_COLORS, buildPackage, boolc, fml, grid, inl, num, row, sst, themeXml, type SheetSpec } from './addsheets-build';
import { REPO, compareSheets, definedNamesOf, lintPackage, lintText, plannedNames, premergeCopies, readFiles, readFix, sheetsOf, type FileIn } from './addsheets-helpers';

const REGION = { sheet: '名簿', headerRow: 1, firstDataRow: 2, lastDataRow: 3, firstCol: 1, lastCol: 2 };
async function baseFile(over: Parameters<typeof buildPackage>[0] extends infer T ? Partial<T> : never = {}): Promise<FileIn> {
  const bytes = await buildPackage({ sheets: [{ name: '名簿', data: grid([['ID', '氏名'], ['A001', '山田'], ['A002', '鈴木']]) }], ...over });
  return { bytes, fileName: '名簿.xlsx', region: REGION };
}
const styles = (xfs: string, fmts = '', fonts = '<font><sz val="11"/><name val="Calibri"/></font>') =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="${MAIN_NS}">${fmts}<fonts count="${(fonts.match(/<font>/g) ?? []).length}">${fonts}</fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${(xfs.match(/<xf /g) ?? []).length}">${xfs}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/></styleSheet>`;

describe('date systems', () => {
  const xfs = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="20" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="46" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>';
  const sheet = (): SheetSpec => ({
    name: '日付',
    data: row(1, num('A1', 43831, 1) + num('B1', 0.5, 2) + num('C1', 1.25, 3) + num('D1', 43831.75, 4) + num('E1', 12345))
      + row(2, fml('A2', 'A1+1', 43832, 'n', 1) + num('B2', 2000, 1)),
  });
  // in a 1904 workbook the same calendar date is 1462 days smaller
  const src1904 = () => buildPackage({ sheets: [{ ...sheet(), data: sheet().data!.replace('43831', '42369').replace('43832', '42370').replace('43831.75', '42369.75').replace('2000', '538') }], styles: styles(xfs), date1904: true });

  it('a 1904 source into a 1900 base: calendar dates, date-times and times from 1 day on move, times of day / durations / plain numbers do not', async () => {
    const res = await premergeCopies(await baseFile(), [{ bytes: await src1904(), fileName: '旧.xlsx', sheet: '日付' }], { dropOtherSheets: true });
    const src = await readBook(await src1904(), '旧.xlsx');
    const out = await readBook(res.bytes, 'o.xlsx');
    expect(compareSheets(src, '日付', out, '旧_日付', { rename: new Map([['日付', '旧_日付']]), styles: true })).toEqual([]);
    const g = out.sheets.find((s) => s.name === '旧_日付')!;
    expect(g.rows[0]![0]!.v).toBe(43831); // 2020-01-01 in the 1900 system
    expect(g.rows[0]![1]!.v).toBe(0.5);
    expect(g.rows[0]![4]!.v).toBe(12345);
    const files = await readFiles(res.bytes);
    const xml = files.get(sheetsOf(files).find((s) => s.name === '旧_日付')!.part)!;
    expect(xml).toMatch(/<c r="A1" s="\d+"><v>43831<\/v><\/c>/); // stored in the base's system
    expect(xml).toMatch(/<c r="A2" s="\d+"><f>A1\+1<\/f><v>43832<\/v><\/c>/); // the cached result of a formula too
    const w = res.report.items.find((i) => i.code === 'sheetCopyLimits' && i.message.includes('日付の基準'))!;
    expect(w.message).toMatch(/1904 年方式.*1900 年方式/);
    expect(res.layout.addedSheets![0].complete).toBe(false);
    expect(lintText(await lintPackage(res.bytes))).toBe('');
  });

  it('a 1900 source into a 1904 base: the other way round (dates are 1462 days smaller, a time of day stays)', async () => {
    const src1900 = await buildPackage({ sheets: [sheet()], styles: styles(xfs) });
    const base = await baseFile({ date1904: true });
    const res = await premergeCopies(base, [{ bytes: src1900, fileName: '新.xlsx', sheet: '日付' }], { dropOtherSheets: true });
    const files = await readFiles(res.bytes);
    const xml = files.get(sheetsOf(files).find((s) => s.name === '新_日付')!.part)!;
    expect(xml).toMatch(/<c r="A1" s="\d+"><v>42369<\/v><\/c>/);
    expect(xml).toMatch(/<c r="B1" s="\d+"><v>0\.5<\/v><\/c>/);
    expect(xml).toMatch(/<c r="E1" s="\d+"><v>12345<\/v><\/c>/);
    const src = await readBook(src1900, '新.xlsx');
    const out = await readBook(res.bytes, 'o.xlsx');
    expect(compareSheets(src, '日付', out, '新_日付', { rename: new Map([['日付', '新_日付']]), styles: true })).toEqual([]);
  });

  it('the same date system on both sides: nothing is shifted, no warning', async () => {
    const src = await buildPackage({ sheets: [sheet()], styles: styles(xfs), date1904: true });
    const res = await premergeCopies(await baseFile({ date1904: true }), [{ bytes: src, fileName: '同.xlsx', sheet: '日付' }], { dropOtherSheets: true });
    expect(res.report.items.some((i) => i.message.includes('日付の基準'))).toBe(false);
    const files = await readFiles(res.bytes);
    expect(files.get(sheetsOf(files).find((s) => s.name === '同_日付')!.part)!).toMatch(/<c r="A1" s="\d+"><v>43831<\/v><\/c>/);
  });
});

describe('styles', () => {
  it('custom number formats: ids that collide between the two workbooks are renumbered, the code is what counts', async () => {
    const baseStyles = styles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>', '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0&quot;base&quot;"/></numFmts>');
    const srcStyles = styles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>', '<numFmts count="2"><numFmt numFmtId="164" formatCode="0.00&quot;src&quot;"/><numFmt numFmtId="165" formatCode="0.0&quot;base&quot;"/></numFmts>');
    const base = await baseFile({ styles: baseStyles });
    const src = await buildPackage({ sheets: [{ name: 'F', data: row(1, num('A1', 1.234, 1) + num('B1', 5.5, 2)) }], styles: srcStyles });
    const res = await premergeCopies(base, [{ bytes: src, fileName: 's.xlsx', sheet: 'F' }], { dropOtherSheets: true });
    const out = await readBook(res.bytes, 'o.xlsx');
    const g = out.sheets.find((s) => s.name === 's_F')!;
    expect(out.styles[g.rows[0]![0]!.s ?? 0].numFmt).toBe('0.00"src"');
    expect(out.styles[g.rows[0]![1]!.s ?? 0].numFmt).toBe('0.0"base"');
    expect(lintText(await lintPackage(res.bytes))).toBe('');
  });

  it('theme colours of the source are converted to rgb when the themes differ and kept as they are when the themes agree', async () => {
    const fonts = '<font><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><color theme="4"/><name val="Calibri"/></font>';
    const xfs = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>';
    const src = (theme: string[]) => buildPackage({ sheets: [{ name: 'T', data: row(1, num('A1', 1, 1)) }], styles: styles(xfs, '', fonts), theme: themeXml(theme) });
    const differ = await premergeCopies(await baseFile(), [{ bytes: await src(OTHER_COLORS), fileName: 'd.xlsx', sheet: 'T' }], { dropOtherSheets: true });
    const out1 = await readBook(differ.bytes, 'a.xlsx');
    const color1 = out1.styles[out1.sheets.find((s) => s.name === 'd_T')!.rows[0]![0]!.s ?? 0].color;
    expect(color1?.toLowerCase()).toBe('#c0504d'); // accent1 of the OTHER theme
    const same = await premergeCopies(await baseFile(), [{ bytes: await src(OFFICE_COLORS), fileName: 's.xlsx', sheet: 'T' }], { dropOtherSheets: true });
    const files = await readFiles(same.bytes);
    expect(files.get('xl/styles.xml')).toMatch(/<color theme="4"\/>/);
    const out2 = await readBook(same.bytes, 'b.xlsx');
    expect(out2.styles[out2.sheets.find((s) => s.name === 's_T')!.rows[0]![0]!.s ?? 0].color?.toLowerCase()).toBe('#4472c4');
  });

  it('the 64,000-format limit of a workbook is respected: what does not fit is shown with the default format and reported, the package stays valid', async () => {
    const many = Array.from({ length: 63998 }, (_x, i) => `<xf numFmtId="0" fontId="${i % 2}" fillId="${i % 2}" borderId="0" xfId="0" applyFont="1"><alignment indent="${i % 15}" horizontal="left"/></xf>`).join('');
    // only 1 in 8 xfs are distinct records; the exact count of distinct ones does not matter: the base is full to the brim
    const baseStyles = styles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' + many, '', '<font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font>');
    const base = await baseFile({ styles: baseStyles });
    const sx = Array.from({ length: 6 }, (_x, i) => `<xf numFmtId="${200 + i}" fontId="0" fillId="0" borderId="0" xfId="0"/>`).join('');
    const srcStyles = styles('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' + sx, `<numFmts count="6">${Array.from({ length: 6 }, (_x, i) => `<numFmt numFmtId="${200 + i}" formatCode="0.${'0'.repeat(i + 1)}&quot;z&quot;"/>`).join('')}</numFmts>`);
    const src = await buildPackage({ sheets: [{ name: 'X', data: row(1, [1, 2, 3, 4, 5, 6].map((k) => num(`${'ABCDEF'[k - 1]}1`, k, k)).join('')) }], styles: srcStyles });
    const res = await premergeCopies(base, [{ bytes: src, fileName: 'x.xlsx', sheet: 'X' }], { dropOtherSheets: true });
    const files = await readFiles(res.bytes);
    const xfCount = (files.get('xl/styles.xml')!.match(/<cellXfs\b[^>]*count="(\d+)"/) ?? [])[1];
    expect(Number(xfCount)).toBeLessThanOrEqual(64000);
    const w = res.report.items.find((i) => i.code === 'sheetCopyLimits' && i.message.includes('上限'))!;
    expect(w.message).toMatch(/書式の種類が Excel の上限（64,000）に達した/);
    expect(res.layout.addedSheets![0].complete).toBe(false);
    expect(lintText(await lintPackage(res.bytes))).toBe('');
  }, 60_000);
});

describe('kinds of sources and bases', () => {
  it('a csv and an xls (values only) are copied with values and number formats; the warning says what is lost', async () => {
    const res = await premergeCopies({ bytes: readFix('samples', '01_学生名簿.xlsx'), fileName: '01_学生名簿.xlsx' }, [
      { bytes: readFix('samples', '03_連絡先.csv'), fileName: '03_連絡先.csv' },
      { bytes: readFix('samples', '04_適性検査.xls'), fileName: '04_適性検査.xls' },
    ], { dropOtherSheets: true });
    const out = await readBook(res.bytes, 'o.xlsx');
    const csv = await readBook(readFix('samples', '03_連絡先.csv'), '03_連絡先.csv');
    const xls = await readBook(readFix('samples', '04_適性検査.xls'), '04_適性検査.xls');
    for (const [src, name, file] of [[csv, csv.sheets[0].name, '03_連絡先'], [xls, xls.sheets[0].name, '04_適性検査']] as const) {
      const a = res.layout.addedSheets!.find((x) => x.original === name && x.name.startsWith(file))!;
      expect(compareSheets(src, name, out, a.name, { rename: new Map([[name.toLowerCase(), a.name]]), styles: true })).toEqual([]);
      expect(a.complete).toBe(false);
    }
    expect(res.report.items.filter((i) => i.code === 'sheetCopyLimits' && i.message.includes('値だけを読み込める形式')).length).toBe(2);
    expect(lintText(await lintPackage(res.bytes))).toBe('');
  });

  it('an xlsm base stays macro-enabled (vbaProject, workbook content type) and gets the copies', async () => {
    const xlsm = new Uint8Array((await import('node:fs')).readFileSync(`${REPO}/tests/fixtures/foreign/corpus/hc-31-macro-enabled.xlsm`));
    const res = await premergeCopies({ bytes: xlsm, fileName: 'hc-31.xlsm' }, [{ bytes: readFix('samples', '02_面接評価.xlsx'), fileName: '02_面接評価.xlsx' }], { dropOtherSheets: false });
    const files = await readFiles(res.bytes);
    expect(files.get('[Content_Types].xml')).toMatch(/application\/vnd\.ms-excel\.sheet\.macroEnabled\.main\+xml/);
    expect([...files.keys()].some((k) => /vbaProject\.bin$/.test(k))).toBe(true);
    expect(sheetsOf(files).map((s) => s.name).slice(-2)).toEqual(['02_面接評価_面接評価', '02_面接評価_評価テーブル']);
    expect(lintText(await lintPackage(res.bytes))).toBe('');
  });

  it('strict OOXML as the source (hc-08) and as the base: the copy is written in the target\'s namespaces', async () => {
    const strict = new Uint8Array((await import('node:fs')).readFileSync(`${REPO}/tests/fixtures/foreign/corpus/hc-08-strict-ooxml.xlsx`));
    const kitchenLike = await buildPackage({ sheets: [{ name: 'T', data: grid([['a', 1], ['b', 2]]) }] });
    const asSource = await premergeCopies({ bytes: readFix('samples', '01_学生名簿.xlsx'), fileName: '01.xlsx' }, [{ bytes: strict, fileName: 'strict.xlsx' }], { dropOtherSheets: true });
    const f1 = await readFiles(asSource.bytes);
    const copy = f1.get(sheetsOf(f1).find((s) => s.name === 'strict_名簿')!.part)!;
    expect(copy).not.toMatch(/purl\.oclc\.org/);
    expect(copy).toMatch(/<c r="F2" s="\d+" t="d"><v>2024-04-12<\/v><\/c>/); // an ISO date cell keeps its text
    expect(lintText(await lintPackage(asSource.bytes))).toBe('');
    const out = await readBook(asSource.bytes, 'o.xlsx');
    const src = await readBook(strict, 'strict.xlsx');
    expect(compareSheets(src, '名簿', out, 'strict_名簿', { rename: new Map([['名簿', 'strict_名簿']]), styles: true })).toEqual([]);

    const asBase = await premergeCopies({ bytes: strict, fileName: 'strict.xlsx' }, [{ bytes: kitchenLike, fileName: 'k.xlsx' }], { dropOtherSheets: true });
    const f2 = await readFiles(asBase.bytes);
    const sh = sheetsOf(f2).find((s) => s.name === 'k_T')!;
    expect(f2.get(sh.part)).toMatch(/xmlns="http:\/\/purl\.oclc\.org\/ooxml\/spreadsheetml\/main"/);
    expect(f2.get('xl/_rels/workbook.xml.rels')).toMatch(/purl\.oclc\.org\/ooxml\/officeDocument\/relationships\/worksheet" Target="worksheets\/sheet2\.xml"/);
  });
});

describe('names of the copies', () => {
  const one = (n: string, rows = 1) => buildPackage({ sheets: [{ name: n, data: grid(Array.from({ length: rows }, (_x, i) => [i])) }] });

  it('three files with the same name: the one added later gets (1), (2)', async () => {
    const bytes = await one('名簿');
    const res = await premergeCopies(await baseFile(), [
      { bytes, fileName: '名簿.xlsx' }, { bytes, fileName: '名簿.xlsx' }, { bytes, fileName: 'C:\\x\\名簿.xlsx' },
    ], { dropOtherSheets: true });
    const names = sheetsOf(await readFiles(res.bytes)).map((s) => s.name);
    // the base's 名簿 is the merged sheet and is never renamed; the first copy is 名簿_名簿
    expect(names).toEqual(['名簿', '名簿_名簿', '名簿_名簿(1)', '名簿_名簿(2)']);
    expect(res.layout.addedSheets!.map((s) => s.suffixed)).toEqual([false, true, true]);
    const info = res.report.items.filter((i) => i.code === 'sheetRenamed').map((i) => i.message);
    expect(info).toEqual(['シート名が重なったため「名簿_名簿(1)」にしました。', 'シート名が重なったため「名簿_名簿(2)」にしました。']);
  });

  it('31 characters at most (the file part is shortened first), forbidden characters become "_", the planner\'s names are applied as they are', async () => {
    const longFile = '営業部門の月次実績報告書（2026年度第3四半期確定版）.xlsx';
    const res = await premergeCopies(await baseFile(), [
      { bytes: await one('実績データ（全社）'), fileName: longFile },
      { bytes: await one('a:b/c'), fileName: 'we?ird*.xlsx' },
      { bytes: await one('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'), fileName: 'f.xlsx' },
    ], { dropOtherSheets: true });
    const names = sheetsOf(await readFiles(res.bytes)).map((s) => s.name);
    const expected = plannedNames(['名簿'], [{ fileName: longFile, sheets: ['実績データ（全社）'] }, { fileName: 'we?ird*.xlsx', sheets: ['a:b/c'] }, { fileName: 'f.xlsx', sheets: ['ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'] }]);
    expect(names).toEqual(['名簿', ...expected]);
    for (const n of names) expect(n.length).toBeLessThanOrEqual(31);
    expect(names[2]).toBe('we_ird__a_b_c');
    expect(lintText(await lintPackage(res.bytes))).toBe('');
  });

  it('a copy whose name is taken by a sheet of the base gets the suffix; the base keeps its name; case differences count', async () => {
    const base = await baseFile({ sheets: [{ name: '名簿', data: grid([['ID', '氏名'], ['A001', '山田'], ['A002', '鈴木']]) }, { name: 'X_Y', data: grid([[1]]) }] });
    const res = await premergeCopies(base, [{ bytes: await one('y'), fileName: 'x.xlsx' }], { dropOtherSheets: false });
    const names = sheetsOf(await readFiles(res.bytes)).map((s) => s.name);
    expect(names).toEqual(['名簿', 'X_Y', 'x_y(1)']);
    expect(res.layout.addedSheets![0]).toMatchObject({ name: 'x_y(1)', suffixed: true, original: 'y' });
  });

  it('many files and sheets (50 copies) keep their order and pass the linter', async () => {
    const book = await buildPackage({ sheets: Array.from({ length: 5 }, (_x, i) => ({ name: `シート${i + 1}`, data: grid([[i, { f: `シート${((i + 1) % 5) + 1}!A1`, v: i }]]) })) });
    const others = Array.from({ length: 10 }, (_x, i) => ({ bytes: book, fileName: `file${String(i).padStart(2, '0')}.xlsx` }));
    const res = await premergeCopies(await baseFile(), others, { dropOtherSheets: true });
    const names = sheetsOf(await readFiles(res.bytes)).map((s) => s.name);
    expect(names.length).toBe(51);
    expect(names.slice(1, 7)).toEqual(['file00_シート1', 'file00_シート2', 'file00_シート3', 'file00_シート4', 'file00_シート5', 'file01_シート1']);
    expect(lintText(await lintPackage(res.bytes))).toBe('');
    const out = await readBook(res.bytes, 'o.xlsx');
    expect(out.sheets.find((s) => s.name === 'file07_シート5')!.rows[0]![1]!.f).toBe('file07_シート1!A1'); // each file's formulas read its own copies
    expect((await definedNamesOf(res.bytes)).size).toBe(0);
  });
});

describe('odd content', () => {
  it('empty sheets, formatted empty cells, strings with leading / trailing spaces, CR, tabs and escapes, errors, booleans', async () => {
    const xfs = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>';
    const src = await buildPackage({
      sheets: [
        { name: '空', data: '' },
        { name: '書式だけ', data: row(3, '<c r="C3" s="1"/><c r="D3" s="1"/>') },
        { name: '文字', data: row(1, sst('A1', 0) + sst('B1', 1) + sst('C1', 2) + sst('D1', 3) + boolc('E1', true) + '<c r="F1" t="e"><v>#REF!</v></c>' + inl('G1', 'a\tb\n_x0041_')) },
      ],
      sst: [' lead', 'trail ', 'cr_x000D_lf', { xml: '<r><t>ru</t></r><r><rPr><b/><sz val="11"/><rFont val="Calibri"/></rPr><t>n</t></r>' }],
      styles: styles(xfs, '', '<font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font>'),
    });
    const res = await premergeCopies(await baseFile(), [{ bytes: src, fileName: 'odd.xlsx', sheet: '文字' }], { dropOtherSheets: true });
    const out = await readBook(res.bytes, 'o.xlsx');
    const orig = await readBook(src, 'odd.xlsx');
    for (const n of ['空', '書式だけ', '文字']) expect(compareSheets(orig, n, out, `odd_${n}`, { rename: new Map(), styles: true }), n).toEqual([]);
    expect(out.sheets.find((s) => s.name === 'odd_文字')!.rows[0]![0]!.v).toBe(' lead');
    expect(out.sheets.find((s) => s.name === 'odd_文字')!.rows[0]![2]!.v).toBe('cr\rlf');
    expect(lintText(await lintPackage(res.bytes))).toBe('');
  });
});

void MAIN_NS; void PLAIN_STYLES;
