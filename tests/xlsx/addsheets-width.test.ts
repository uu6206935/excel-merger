/**
 * Column widths of a copied sheet. A stored width is a number of characters of the workbook's default font (Calibri 11 / Arial 10: 7 px per
 * digit, the Japanese fonts at 11 pt: 8 px), so a sheet copied into a workbook with another default font keeps its PIXEL widths — the copy
 * is not 14 % wider or narrower than the original. A default font whose digit width is not known leaves the widths as they are.
 */
import { describe, expect, it } from 'vitest';
import { excelWidthToPx, pxToExcelWidth } from '../../src/core/display/styles';
import { MAIN_NS, PLAIN_STYLES, OFFICE_COLORS, WS_NS, buildPackage, grid, themeXml } from './addsheets-build';
import { lintPackage, premergeCopies, readFiles, sheetsOf, type FileIn } from './addsheets-helpers';

const REGION = { sheet: '名簿', headerRow: 1, firstDataRow: 2, lastDataRow: 3, firstCol: 1, lastCol: 2 };
const baseFile = async (over: Partial<Parameters<typeof buildPackage>[0]> = {}): Promise<FileIn> => ({
  bytes: await buildPackage({ sheets: [{ name: '名簿', data: grid([['ID', '氏名'], ['A001', '山田'], ['A002', '鈴木']]) }], ...over }),
  fileName: '名簿.xlsx',
  region: REGION,
});

const stylesOf = (font: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="${MAIN_NS}"><fonts count="1"><font>${font}</font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/></styleSheet>`;
const CALIBRI = stylesOf('<sz val="11"/><name val="Calibri"/><family val="2"/>');
const ARIAL10 = stylesOf('<sz val="10"/><name val="Arial"/>');
const VERDANA = stylesOf('<sz val="9"/><name val="Verdana"/>');
const SCHEME_MINOR = stylesOf('<sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/>');

const COLS = '<cols><col min="1" max="1" width="10" customWidth="1"/><col min="2" max="3" width="14.5" customWidth="1"/><col min="5" max="5" width="30" hidden="1" customWidth="1"/></cols>';
const sheetXml = (formatPr: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet ${WS_NS}><dimension ref="A1:B2"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>${formatPr}${COLS}<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row></sheetData></worksheet>`;

async function copyOf(src: Uint8Array, base?: FileIn): Promise<string> {
  const res = await premergeCopies(base ?? (await baseFile()), [{ bytes: src, fileName: '他.xlsx', sheet: 'S' }], { dropOtherSheets: true });
  expect(await lintPackage(res.bytes)).toEqual([]);
  const f = await readFiles(res.bytes);
  return f.get(sheetsOf(f).find((s) => s.name === '他_S')!.part)!;
}
const widths = (x: string): number[] => [...x.matchAll(/<col\b[^>]*\swidth="([^"]+)"/g)].map((m) => Number(m[1]));
const pxWidth = (chars: number, mdw: number, to: number): number => Math.round(pxToExcelWidth(excelWidthToPx(chars, mdw), to) * 1e6) / 1e6;

describe('column widths of a copy', () => {
  it('Calibri 11 source into a 游ゴシック 11 base: pixel widths are kept (7 px -> 8 px per digit), the default width too', async () => {
    const x = await copyOf(await buildPackage({ sheets: [{ name: 'S', xml: sheetXml('<sheetFormatPr defaultRowHeight="15"/>') }], styles: CALIBRI }));
    expect(widths(x)).toEqual([pxWidth(10, 7, 8), pxWidth(14.5, 7, 8), pxWidth(30, 7, 8)]);
    expect(widths(x)[0]).not.toBe(10);
    // 8 characters of Calibri = 64 px = 8 characters of 游ゴシック
    expect(x).toContain(`<sheetFormatPr defaultRowHeight="15" defaultColWidth="${Math.round(pxToExcelWidth(64, 8) * 1e6) / 1e6}"/>`);
    expect(x.indexOf('<sheetFormatPr')).toBeLessThan(x.indexOf('<cols>'));
  });

  it('a sheet WITHOUT <sheetFormatPr> gets one (schema order) so that its default column width stays 64 px', async () => {
    const x = await copyOf(await buildPackage({ sheets: [{ name: 'S', xml: sheetXml('') }], styles: CALIBRI }));
    const dcw = Math.round(pxToExcelWidth(64, 8) * 1e6) / 1e6;
    expect(x).toContain(`</sheetViews><sheetFormatPr defaultRowHeight="15" defaultColWidth="${dcw}"/><cols>`);
    expect((x.match(/<sheetFormatPr\b/g) ?? []).length).toBe(1);
  });

  it('Arial 10 is 7 px per digit as well', async () => {
    const x = await copyOf(await buildPackage({ sheets: [{ name: 'S', xml: sheetXml('') }], styles: ARIAL10 }));
    expect(widths(x)).toEqual([pxWidth(10, 7, 8), pxWidth(14.5, 7, 8), pxWidth(30, 7, 8)]);
  });

  it('the same default font on both sides: widths and <sheetFormatPr> are copied as they are', async () => {
    const x = await copyOf(await buildPackage({ sheets: [{ name: 'S', xml: sheetXml('') }], styles: PLAIN_STYLES }));
    expect(widths(x)).toEqual([10, 14.5, 30]);
    expect(x).not.toContain('<sheetFormatPr');
  });

  it('a default font whose digit width is not known: widths are left alone and nothing is invented', async () => {
    const x = await copyOf(await buildPackage({ sheets: [{ name: 'S', xml: sheetXml('') }], styles: VERDANA }));
    expect(widths(x)).toEqual([10, 14.5, 30]);
    expect(x).not.toContain('<sheetFormatPr');
  });

  it('a scheme (theme) font is resolved through the theme: Calibri named, but the theme maps the Japanese script to a 8 px font -> no conversion', async () => {
    const theme = themeXml(OFFICE_COLORS, 'Calibri Light', 'Calibri').replace('<a:cs typeface=""/></a:minorFont>', '<a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/></a:minorFont>');
    const x = await copyOf(await buildPackage({ sheets: [{ name: 'S', xml: sheetXml('') }], styles: SCHEME_MINOR, theme }));
    expect(widths(x)).toEqual([10, 14.5, 30]);
  });

  it('the same scheme font without a Japanese script entry stays Calibri (7 px) and is converted', async () => {
    const x = await copyOf(await buildPackage({ sheets: [{ name: 'S', xml: sheetXml('') }], styles: SCHEME_MINOR, theme: themeXml(OFFICE_COLORS, 'Calibri Light', 'Calibri') }));
    expect(widths(x)).toEqual([pxWidth(10, 7, 8), pxWidth(14.5, 7, 8), pxWidth(30, 7, 8)]);
  });

  it('the base has the narrow font: a 游ゴシック source is converted the other way round', async () => {
    const base = await baseFile({ styles: CALIBRI });
    const x = await copyOf(await buildPackage({ sheets: [{ name: 'S', xml: sheetXml('') }], styles: PLAIN_STYLES }), base);
    expect(widths(x)).toEqual([pxWidth(10, 8, 7), pxWidth(14.5, 8, 7), pxWidth(30, 8, 7)]);
    // 8 characters of 游ゴシック (72 px incl. padding) in Calibri units
    expect(x).toContain(`defaultColWidth="${Math.round(pxToExcelWidth(72, 7) * 1e6) / 1e6}"`);
  });
});
