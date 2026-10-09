/**
 * SheetGrid.defaultColWidth from <sheetFormatPr>: defaultColWidth when positive, else derived from baseColWidth (Mac Excel
 * writes <sheetFormatPr baseColWidth="10" defaultRowHeight="16"/> and no defaultColWidth in every file), else undefined
 * (= the standard 8 characters). The doc's auto-widening (doc/autowiden.ts) and the grid take this value; the JS patcher
 * derives the same width itself from baseColWidth (patch/engine.ts autoWiden), so all three must agree — and Excel.
 *
 * ORACLE: pixel widths measured in a hidden Excel instance (2026-10, Windows, 96 dpi) as Range.Width / 0.75 of a column
 * without a <col> element, for workbooks that differ only in the default font and baseColWidth (the default font = Normal
 * style = cellXfs[0]'s font). tests/doc/basecolwidth.test.ts is the regression for the doc / patcher side.
 */
import { describe, expect, it } from 'vitest';
import { defaultColumnPx, excelWidthToPx, mdwFor } from '../../src/core/display/styles';
import { readBook, sheetDefaultColWidth } from '../../src/core/xlsx/read';
import { makePackage } from '../fixtures/xlsx/helpers';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

const sheetXml = (fmt: string, prefix = '') =>
  prefix
    ? `<?xml version="1.0"?><${prefix}:worksheet xmlns:${prefix}="${NS}">${fmt}<${prefix}:sheetData><${prefix}:row r="1"><${prefix}:c r="A1"><${prefix}:v>1</${prefix}:v></${prefix}:c></${prefix}:row></${prefix}:sheetData></${prefix}:worksheet>`
    : `<?xml version="1.0"?><worksheet xmlns="${NS}">${fmt}<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`;

/** styles.xml whose cellXfs[0] uses fonts[xf0Font] */
function stylesXml(fonts: Array<[string, number]>, xf0Font = 0): string {
  return (
    `<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="${NS}"><fonts count="${fonts.length}">` +
    fonts.map(([n, s]) => `<font><sz val="${s}"/><name val="${n}"/><family val="2"/></font>`).join('') +
    '</fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="1"><xf numFmtId="0" fontId="${xf0Font}" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`
  );
}

const fmtOf = (attrs: string) => `<sheetFormatPr${attrs} defaultRowHeight="16"/>`;

/** the reader's defaultColWidth of a one-sheet workbook with the given sheetFormatPr attributes and default font */
async function widthOf(attrs: string, font: [string, number] = ['游ゴシック', 11]): Promise<number | undefined> {
  const pkg = await makePackage({ sheets: [{ name: 'S', xml: sheetXml(fmtOf(attrs)) }], styles: stylesXml([font]) });
  return (await readBook(pkg, 'f.xlsx')).sheets[0].defaultColWidth;
}

/** the pixel width the grid / doc / patcher give the default column of a sheet whose reader value is `w` (undefined = 8 characters) */
const pxOf = (w: number | undefined, font: [string, number]) => {
  const mdw = mdwFor(font[0], font[1]);
  return w === undefined ? defaultColumnPx(mdw, 8) : excelWidthToPx(w, mdw);
};

// Excel's own numbers (px of Range.Width / 0.75) per default font and baseColWidth
const YUGO11: Array<[number, number]> = [[1, 16], [5, 48], [9, 80], [10, 88], [11, 96], [12, 104], [15, 128], [20, 168], [100, 808]];
const CALIBRI11: Array<[number, number]> = [[1, 16], [5, 40], [9, 72], [10, 80], [11, 88], [12, 96], [15, 112], [20, 152], [100, 712]];

describe('SheetGrid.defaultColWidth: baseColWidth without defaultColWidth (Mac Excel writes baseColWidth="10")', () => {
  it('absent baseColWidth, baseColWidth="8" (the schema default) and no <sheetFormatPr> at all: undefined = the standard 8 characters', async () => {
    expect(await widthOf('')).toBeUndefined();
    expect(await widthOf(' baseColWidth="8"')).toBeUndefined();
    expect(await widthOf(' baseColWidth="8.0"')).toBeUndefined();
    const pkg = await makePackage({ sheets: [{ name: 'S', xml: sheetXml('') }], styles: stylesXml([['游ゴシック', 11]]) });
    expect((await readBook(pkg, 'f.xlsx')).sheets[0].defaultColWidth).toBeUndefined();
    // Excel: 72 px for 游ゴシック 11 and 64 px for Calibri 11 — the 8-character default the consumers assume
    expect(pxOf(undefined, ['游ゴシック', 11])).toBe(72);
    expect(pxOf(undefined, ['Calibri', 11])).toBe(64);
  });

  it('baseColWidth="10": 88 px for 游ゴシック 11 (stored width 11), 80 px for Calibri 11 — the pixel widths Excel shows', async () => {
    const yugo = await widthOf(' baseColWidth="10"', ['游ゴシック', 11]);
    expect(yugo).toBeCloseTo(11, 2);
    expect(excelWidthToPx(yugo!, 8)).toBe(88);
    const calibri = await widthOf(' baseColWidth="10"', ['Calibri', 11]);
    expect(calibri).toBeCloseTo(11.42, 1);
    expect(excelWidthToPx(calibri!, 7)).toBe(80);
  });

  it('the Mac default (Calibri 12 / Aptos Narrow 12, <sheetFormatPr baseColWidth="10" defaultRowHeight="16"/>): 88 px', async () => {
    for (const font of ['Calibri', 'Aptos Narrow']) {
      const w = await widthOf(' baseColWidth="10"', [font, 12]);
      expect(w, font).toBeGreaterThan(10.9);
      expect(excelWidthToPx(w!, mdwFor(font, 12)), font).toBe(88); // Excel: Range.Width 66 pt, 10.38 characters
    }
    // Excel (Aptos Narrow 11, 游ゴシック 10, ＭＳ Ｐゴシック 10, 游明朝 11): 80 / 80 / 80 / 88 px
    expect(pxOf(await widthOf(' baseColWidth="10"', ['Aptos Narrow', 11]), ['Aptos Narrow', 11])).toBe(80);
    expect(pxOf(await widthOf(' baseColWidth="10"', ['游ゴシック', 10]), ['游ゴシック', 10])).toBe(80);
    expect(pxOf(await widthOf(' baseColWidth="10"', ['ＭＳ Ｐゴシック', 10]), ['ＭＳ Ｐゴシック', 10])).toBe(80);
    expect(pxOf(await widthOf(' baseColWidth="10"', ['游明朝', 11]), ['游明朝', 11])).toBe(88);
  });

  it('every baseColWidth Excel was measured with gives Excel\'s pixel width (游ゴシック 11, ＭＳ Ｐゴシック 11 / Calibri 11, メイリオ 9)', async () => {
    for (const [font, table] of [
      [['游ゴシック', 11], YUGO11],
      [['ＭＳ Ｐゴシック', 11], YUGO11],
      [['Calibri', 11], CALIBRI11],
      [['メイリオ', 9], CALIBRI11],
    ] as Array<[[string, number], Array<[number, number]>]>) {
      for (const [base, px] of table) {
        const w = await widthOf(` baseColWidth="${base}"`, font);
        expect(w, `${font[0]} ${base}`).toBeDefined();
        expect(pxOf(w, font), `${font[0]} ${font[1]} base ${base}`).toBe(px);
      }
    }
  });

  it('the stored unit is the one Excel writes (width = px / mdw) within one 1/256 step: it round-trips to the same pixels', async () => {
    // Excel saved defaultColWidth="11" / "11.42578125" / "12.5703125" for 游ゴシック 11 / Calibri 11 base 10 / Calibri 11 base 11
    expect(Math.abs((await widthOf(' baseColWidth="10"', ['游ゴシック', 11]))! - 11)).toBeLessThanOrEqual(1 / 256);
    expect(Math.abs((await widthOf(' baseColWidth="10"', ['Calibri', 11]))! - 11.42578125)).toBeLessThanOrEqual(1 / 256);
    expect(Math.abs((await widthOf(' baseColWidth="11"', ['Calibri', 11]))! - 12.5703125)).toBeLessThanOrEqual(1 / 256);
  });

  it('an explicit defaultColWidth wins over baseColWidth (Excel: defaultColWidth 9.5 with baseColWidth 10 = 76 px)', async () => {
    expect(await widthOf(' baseColWidth="10" defaultColWidth="9.5"')).toBe(9.5);
    expect(pxOf(9.5, ['游ゴシック', 11])).toBe(76);
    expect(await widthOf(' defaultColWidth="9.5"')).toBe(9.5);
    expect(await widthOf(' baseColWidth="8" defaultColWidth="12"')).toBe(12);
  });

  it('a non-positive defaultColWidth is ignored like the patcher does: baseColWidth decides, else undefined', async () => {
    expect(await widthOf(' defaultColWidth="0"')).toBeUndefined();
    expect(await widthOf(' defaultColWidth="-2"')).toBeUndefined();
    expect(pxOf(await widthOf(' baseColWidth="10" defaultColWidth="0"'), ['游ゴシック', 11])).toBe(88);
    expect(pxOf(await widthOf(' baseColWidth="10" defaultColWidth="-2"'), ['游ゴシック', 11])).toBe(88);
    expect(await widthOf(' defaultColWidth="abc"')).toBeUndefined();
  });

  it('baseColWidth that is not positive, not a number or beyond what a column can be: undefined (Excel falls back to 8 characters)', async () => {
    expect(await widthOf(' baseColWidth="0"')).toBeUndefined(); // Excel: 72 px, same as 8
    expect(await widthOf(' baseColWidth="-4"')).toBeUndefined();
    expect(await widthOf(' baseColWidth=""')).toBeUndefined();
    expect(await widthOf(' baseColWidth="ten"')).toBeUndefined();
    expect(await widthOf(' baseColWidth="255"')).toBeUndefined(); // Excel ignores it (72 px)
    expect(await widthOf(' baseColWidth="65536"')).toBeUndefined(); // Excel ignores it (72 px)
    expect(await widthOf(' baseColWidth="1000000000"')).toBeUndefined();
    expect(await widthOf(' baseColWidth="1e308"')).toBeUndefined();
    // the last width a column can have: 255 characters
    const w = await widthOf(' baseColWidth="254"');
    expect(w).toBeGreaterThan(254);
    expect(excelWidthToPx(w!, 8)).toBe(2040);
  });

  it('is per sheet; baseColWidth of one sheet does not leak into the next', async () => {
    const pkg = await makePackage({
      sheets: [
        { name: 'A', xml: sheetXml(fmtOf(' baseColWidth="10"')) },
        { name: 'B', xml: sheetXml(fmtOf('')) },
        { name: 'C', xml: sheetXml(fmtOf(' baseColWidth="12"')) },
        { name: 'D', xml: sheetXml(fmtOf(' baseColWidth="10" defaultColWidth="9.5"')) },
      ],
      styles: stylesXml([['游ゴシック', 11]]),
    });
    const sheets = (await readBook(pkg, 'f.xlsx')).sheets;
    const px = sheets.map((s) => pxOf(s.defaultColWidth, ['游ゴシック', 11]));
    expect(px).toEqual([88, 72, 104, 76]);
    expect(sheets[1].defaultColWidth).toBeUndefined();
  });

  it('the default font is cellXfs[0]\'s font (what the patcher measures with), not fonts[0]', async () => {
    // fonts: Calibri 11 (7 px digits) and 游ゴシック 11 (8 px); cellXfs[0] uses the second one
    const pkg = await makePackage({
      sheets: [{ name: 'S', xml: sheetXml(fmtOf(' baseColWidth="10"')) }],
      styles: stylesXml([['Calibri', 11], ['游ゴシック', 11]], 1),
    });
    const w = (await readBook(pkg, 'f.xlsx')).sheets[0].defaultColWidth;
    expect(excelWidthToPx(w!, 8)).toBe(88);
    // without a styles part the default font is unknown: 8 px digits (display/styles' fallback), also what the patcher assumes
    const bare = await makePackage({ sheets: [{ name: 'S', xml: sheetXml(fmtOf(' baseColWidth="10"')) }], styles: null });
    expect(excelWidthToPx((await readBook(bare, 'f.xlsx')).sheets[0].defaultColWidth!, 8)).toBe(88);
  });

  it('a prefixed worksheet (x:worksheet / x:sheetFormatPr) and attribute order read the same', async () => {
    const pkg = await makePackage({
      sheets: [{ name: 'S', xml: sheetXml('<x:sheetFormatPr defaultRowHeight="16" baseColWidth="10"/>', 'x') }],
      styles: stylesXml([['游ゴシック', 11]]),
    });
    const w = (await readBook(pkg, 'f.xlsx')).sheets[0].defaultColWidth;
    expect(excelWidthToPx(w!, 8)).toBe(88);
  });

  it('the default row height is read next to it (Mac Excel: Calibri 12, defaultRowHeight 16)', async () => {
    const pkg = await makePackage({ sheets: [{ name: 'S', xml: sheetXml(fmtOf(' baseColWidth="10"')) }], styles: stylesXml([['Calibri', 12]]) });
    const s = (await readBook(pkg, 'f.xlsx')).sheets[0];
    expect(s.defaultRowHeight).toBe(16);
    expect(s.defaultColWidth).toBeGreaterThan(10.9);
  });
});

describe('sheetDefaultColWidth (the rule itself)', () => {
  const yugo = { font: '游ゴシック', size: 11 };

  it('precedence: positive defaultColWidth > baseColWidth != 8 > undefined', () => {
    expect(sheetDefaultColWidth({ defaultColWidth: 9.5, baseColWidth: 10 }, yugo)).toBe(9.5);
    expect(sheetDefaultColWidth({ defaultColWidth: 0, baseColWidth: 10 }, yugo)).toBeCloseTo(11, 2);
    expect(sheetDefaultColWidth({ baseColWidth: 10 }, yugo)).toBeCloseTo(11, 2);
    expect(sheetDefaultColWidth({ baseColWidth: 8 }, yugo)).toBeUndefined();
    expect(sheetDefaultColWidth({}, yugo)).toBeUndefined();
  });

  it('a missing default font or size measures with 11 pt / 8 px digits like the patcher (fontMdw)', () => {
    expect(sheetDefaultColWidth({ baseColWidth: 10 }, undefined)).toBeCloseTo(11, 2);
    expect(sheetDefaultColWidth({ baseColWidth: 10 }, {})).toBeCloseTo(11, 2);
    expect(sheetDefaultColWidth({ baseColWidth: 10 }, { font: '游ゴシック', size: 0 })).toBeCloseTo(11, 2);
    expect(sheetDefaultColWidth({ baseColWidth: 10 }, { font: 'Calibri' })).toBeCloseTo(11.42, 1);
  });

  it('agrees with the patcher\'s formula: excelWidthToPx(result) === defaultColumnPx(mdw, baseColWidth) for every base, font and size', () => {
    for (const font of ['游ゴシック', 'Calibri', 'ＭＳ Ｐゴシック', 'Arial', 'Aptos Narrow', 'Unknown Font']) {
      for (const size of [8, 9, 10, 11, 12, 14, 18, 24]) {
        const mdw = mdwFor(font, size);
        for (let base = 1; base <= 254; base++) {
          if (base === 8) continue;
          const w = sheetDefaultColWidth({ baseColWidth: base }, { font, size });
          const want = defaultColumnPx(mdw, base);
          if (want > 255 * mdw) expect(w, `${font} ${size} base ${base}`).toBeUndefined();
          else expect(excelWidthToPx(w!, mdw), `${font} ${size} base ${base}`).toBe(want);
        }
      }
    }
  });
});
