/**
 * Style resolution vs Excel: theme colours x tints and the indexed palette (excel-colors.json, recorded via COM on
 * theme-probe.xlsx whose theme has a custom accent1), applyX semantics (apply-probe.xlsx), raw record parsing.
 */
import { beforeAll, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { readBook } from '../../src/core/xlsx/read';
import { buildDisplayStyles, dateModeOf, isDateCode, numFmtCode, parseStyleSheet } from '../../src/core/xlsx/styles';
import { applyTint, parseThemeColors, resolveColor, DEFAULT_THEME } from '../../src/core/xlsx/theme';
import type { SourceBook } from '../../src/core/types';
import { COM_HALIGN, cellAt, fx, fxJson, probes } from '../fixtures/xlsx/helpers';

interface ColorProbe {
  tints: number[];
  theme: Array<{ addr: string; comTheme: number; tint: number; fill: string; font: string }>;
  indexed: Array<{ addr: string; colorIndex: number; fill: string; font: string }>;
}

describe('colours vs Excel (theme-probe.xlsx)', () => {
  let book: SourceBook;
  const colors = fxJson<ColorProbe>('excel-colors.json');
  beforeAll(async () => {
    book = await readBook(fx('theme-probe.xlsx'), 'theme-probe.xlsx');
  });

  it('theme colours 0..11 x 14 tints (fill and font) match Excel exactly', () => {
    const s = book.sheets.find((x) => x.name === 'Probe')!;
    const bad: string[] = [];
    for (const e of colors.theme) {
      const st = book.styles[cellAt(s, e.addr)?.s ?? 0];
      if (st.fill !== '#' + e.fill) bad.push(`${e.addr} fill ${st.fill} vs ${e.fill}`);
      if (st.color !== '#' + e.font) bad.push(`${e.addr} font ${st.color} vs ${e.font}`);
    }
    expect(colors.theme.length).toBe(168);
    expect(bad).toEqual([]);
  });

  it('indexed palette colours 1..56 (fill and font) match Excel', () => {
    const s = book.sheets.find((x) => x.name === 'Probe')!;
    const bad: string[] = [];
    for (const e of colors.indexed) {
      const st = book.styles[cellAt(s, e.addr)?.s ?? 0];
      if (st.fill !== '#' + e.fill) bad.push(`${e.addr} fill ${st.fill} vs ${e.fill}`);
      if (st.color !== '#' + e.font) bad.push(`${e.addr} font ${st.color} vs ${e.font}`);
    }
    expect(colors.indexed.length).toBe(56);
    expect(bad).toEqual([]);
  });

  it('custom theme is read from theme1.xml (accent1 = CC9933) in theme-index order lt1, dk1, lt2, dk2, …', async () => {
    const zip = await JSZip.loadAsync(fx('theme-probe.xlsx'));
    const theme = parseThemeColors(await zip.file('xl/theme/theme1.xml')!.async('string'));
    expect(theme[0]).toBe('FFFFFF'); // lt1 (bg1) — listed SECOND in clrScheme
    expect(theme[1]).toBe('000000'); // dk1 (tx1)
    expect(theme[4]).toBe('CC9933');
    expect(parseThemeColors(null)).toEqual([...DEFAULT_THEME]);
  });

  it('applyTint: zero tint is identity; integer HLS math (known Excel values)', () => {
    expect(applyTint('4472c4', 0)).toBe('4472C4');
    expect(applyTint('CC9933', -0.249977111117893)).toBe('997426');
    expect(applyTint('CC9933', 0.39997558519241921)).toBe('E0C285');
    expect(applyTint('000000', 0.499984740745262)).toBe('808080');
    expect(applyTint('FFFFFF', -0.0499893185216834)).toBe('F2F2F2');
    expect(resolveColor({ auto: true })).toBeUndefined();
    expect(resolveColor({ rgb: 'FF00B050' })).toBe('#00B050');
    expect(resolveColor({ indexed: 64 })).toBe('#000000');
  });
});

describe('applyX semantics (apply-probe.xlsx vs Excel)', () => {
  it('Excel renders the cellXf own records whatever applyX says', async () => {
    const book = await readBook(fx('apply-probe.xlsx'), 'apply-probe.xlsx');
    const s = book.sheets[0];
    const pf = probes('apply-probe.xlsx').sheets[0];
    for (const pc of pf.cells) {
      const st = book.styles[cellAt(s, pc.addr)?.s ?? 0];
      expect(!!st.bold, pc.addr).toBe(pc.font!.bold);
      expect(st.font, pc.addr).toBe(pc.font!.name);
      expect(st.size, pc.addr).toBe(pc.font!.size);
      expect(st.fill, pc.addr).toBe(pc.fill!.pattern === -4142 ? undefined : '#' + pc.fill!.color);
      expect(st.hAlign, pc.addr).toBe(COM_HALIGN[pc.hAlign!]);
      expect(st.numFmt ?? 'G/標準', pc.addr).toBe(pc.nf);
    }
    // A3: parent style is bold/filled/date, applyX=0 with plain own records → plain (as Excel shows)
    expect(book.styles[2]).toEqual({ size: 11, font: '游ゴシック', color: '#000000' });
  });
});

describe('parseStyleSheet (raw records for the patcher)', () => {
  it('keeps every record with its verbatim XML, in file order', async () => {
    const zip = await JSZip.loadAsync(fx('torture.xlsx'));
    const xml = await zip.file('xl/styles.xml')!.async('string');
    const ss = parseStyleSheet(xml);
    expect(ss.prefix).toBe('');
    expect(ss.numFmts.map((n) => [n.id, n.code])).toEqual([
      [179, 'yyyy/mm/dd'],
      [180, 'm"月"d"日"\\(aaa\\)'],
      [181, '[$¥-411]#,##0'],
      [182, '0.0%'],
      [183, '#,##0\\ "個"'],
    ]);
    expect(ss.numFmts[1].xml).toBe('<numFmt numFmtId="180" formatCode="m&quot;月&quot;d&quot;日&quot;\\(aaa\\)"/>');
    expect([ss.fonts.length, ss.fills.length, ss.borders.length, ss.cellStyleXfs.length, ss.cellXfs.length, ss.cellStyles.length, ss.dxfs.length]).toEqual([8, 4, 4, 2, 27, 2, 3]);
    for (const f of ss.fonts) expect(xml).toContain(f.xml);
    for (const x of ss.cellXfs) expect(x.xml.startsWith('<xf ')).toBe(true);
    expect(ss.fonts[2]).toMatchObject({ bold: true, size: 16, name: '游ゴシック', color: { theme: 1, tint: -0.249977111117893 } });
    expect(ss.fills[2]).toMatchObject({ patternType: 'solid', fg: { theme: 4, tint: 0.39997558519241921 }, bg: { indexed: 64 } });
    expect(ss.borders[3].bottom).toEqual({ style: 'medium', color: { rgb: 'FFFF0000' } });
    expect(ss.cellXfs[23]).toMatchObject({ fontId: 5, borderId: 2, xfId: 1, applyBorder: true, applyAlignment: true, alignment: { vertical: 'top', wrapText: true } });
    expect(ss.cellXfs[23].applyFont).toBeUndefined();
    expect(ss.cellStyles.map((c) => [c.name, c.xfId, c.builtinId])).toEqual([
      ['ハイパーリンク', 1, 8],
      ['標準', 0, 0],
    ]);
    expect(ss.dxfs[0].xml).toBe('<dxf><font><b/><i val="0"/><color rgb="FFC00000"/></font></dxf>');
  });

  it('<i val="0"/> is NOT italic (ExcelJS bug), prefixed styles part, missing part → defaults', () => {
    const ss = parseStyleSheet(
      '<x:styleSheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:fonts count="1"><x:font><x:b val="0"/><x:i val="0"/><x:u/><x:sz val="9"/><x:name val="Meiryo"/></x:font></x:fonts>' +
        '<x:cellXfs count="1"><x:xf numFmtId="14" fontId="0"/></x:cellXfs></x:styleSheet>',
    );
    expect(ss.prefix).toBe('x:');
    expect(ss.fonts[0]).toMatchObject({ bold: false, italic: false, underline: 'single', size: 9, name: 'Meiryo' });
    const ds = buildDisplayStyles(ss);
    expect(ds.styles[0]).toEqual({ underline: true, size: 9, font: 'Meiryo', numFmt: 'yyyy/m/d' });
    expect(ds.isDate).toEqual([true]);
    const empty = parseStyleSheet(null);
    expect(empty.cellXfs.length).toBe(1);
    expect(buildDisplayStyles(empty).styles[0]).toMatchObject({ size: 11, font: '游ゴシック' });
  });

  it('built-in number formats resolve to ja-JP codes; custom ids win', () => {
    const ss = parseStyleSheet('<styleSheet xmlns="x"><numFmts><numFmt numFmtId="14" formatCode="d/m/yyyy"/></numFmts></styleSheet>');
    expect(numFmtCode(ss, 14)).toBe('d/m/yyyy');
    expect(numFmtCode(parseStyleSheet(null), 14)).toBe('yyyy/m/d');
    expect(numFmtCode(parseStyleSheet(null), 22)).toBe('yyyy/m/d h:mm');
    expect(numFmtCode(parseStyleSheet(null), 58)).toBe('[$-411]ggge"年"m"月"d"日"');
    expect(numFmtCode(parseStyleSheet(null), 999)).toBe('General');
  });
});

describe('date format detection', () => {
  it.each([
    ['General', false],
    ['@', false],
    ['0.00', false],
    ['#,##0\\ "個"', false],
    ['[$¥-411]#,##0', false],
    ['0.00E+00', false],
    ['"yyyy"0', false],
    ['[Red]0.0', false],
    ['yyyy/m/d', true],
    ['m"月"d"日"\\(aaa\\)', true],
    ['[$-411]ggge"年"m"月"d"日"', true],
    ['[$-411]ge.m.d', true],
    ['h:mm', true],
    ['[h]:mm:ss', true],
    ['mm:ss', true],
    ['yyyy/m/d h:mm AM/PM', true],
  ])('isDateCode(%s) = %s', (code, expected) => {
    expect(isDateCode(code)).toBe(expected);
  });

  it.each([
    ['yyyy/m/d', 'date'],
    ['mmm-yy', 'date'],
    ['h:mm', 'time'],
    ['h:mm AM/PM', 'time'],
    ['mm:ss', 'time'],
    ['[h]:mm', 'elapsed'],
    ['yyyy/m/d h:mm', 'date'],
    ['[$-411]ge.m.d', 'date'],
    ['0.0', 'none'],
  ])('dateModeOf(%s) = %s', (code, mode) => {
    expect(dateModeOf(code)).toBe(mode);
  });
});
