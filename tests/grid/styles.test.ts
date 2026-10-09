import { describe, expect, it } from 'vitest';
import {
  DEFAULT_FONT_STACK,
  OFFICE_2013_THEME,
  applyTint,
  charsToPx,
  defaultColumnPx,
  excelWidthToPx,
  fontFamilyFor,
  makeStyleCache,
  mdwFor,
  ptToPx,
  pxToExcelWidth,
  pxToPt,
  resolveExcelColor,
  styleToCss,
} from '../../src/core/display/styles';
import type { XStyle } from '../../src/core/types';

describe('styleToCss', () => {
  it('default style: Excel defaults (bottom aligned, General alignment, default font)', () => {
    const c = styleToCss(undefined);
    expect(c.box).toEqual({});
    expect(c.vAlign).toBe('bottom');
    expect(c.textAlign).toBeUndefined();
    expect(c.wrap).toBe(false);
    expect(c.font).toContain('11pt');
    expect(c.font).toContain('Yu Gothic UI');
  });
  it('fill, font colour, bold/italic/underline/strike, size, family', () => {
    const c = styleToCss({ fill: '#FFFF00', color: '#FF0000', bold: true, italic: true, underline: true, strike: true, size: 14, font: 'Arial' });
    expect(c.box.backgroundColor).toBe('#FFFF00');
    expect(c.box.color).toBe('#FF0000');
    expect(c.box.fontWeight).toBe(700);
    expect(c.box.fontStyle).toBe('italic');
    expect(c.box.textDecoration).toBe('underline line-through');
    expect(c.box.fontSize).toBe('14pt');
    expect(c.box.fontFamily).toBe(`'Arial', ${DEFAULT_FONT_STACK}`);
    expect(c.font).toBe(`italic 700 14pt 'Arial', ${DEFAULT_FONT_STACK}`);
  });
  it('游ゴシック renders with the default UI stack; 11pt is not repeated', () => {
    const c = styleToCss({ font: '游ゴシック', size: 11 });
    expect(c.box.fontFamily).toBeUndefined();
    expect(c.box.fontSize).toBeUndefined();
    expect(fontFamilyFor('Yu Gothic Medium')).toBe(DEFAULT_FONT_STACK);
    expect(fontFamilyFor("O'Brien")).toBe(`'OBrien', ${DEFAULT_FONT_STACK}`);
  });
  it('alignment, wrap, indent', () => {
    expect(styleToCss({ hAlign: 'center' }).textAlign).toBe('center');
    expect(styleToCss({ hAlign: 'centerContinuous' }).textAlign).toBe('center');
    expect(styleToCss({ hAlign: 'right', indent: 2 }).box.paddingRight).toBe(21);
    expect(styleToCss({ hAlign: 'left', indent: 1 }).box.paddingLeft).toBe(12);
    expect(styleToCss({ hAlign: 'general' }).textAlign).toBeUndefined();
    expect(styleToCss({ vAlign: 'top' }).box.alignItems).toBe('flex-start');
    expect(styleToCss({ vAlign: 'middle' }).box.alignItems).toBe('center');
    const w = styleToCss({ wrap: true });
    expect(w.wrap).toBe(true);
    expect(w.box.whiteSpace).toBe('pre-wrap');
    expect(w.box.wordBreak).toBe('break-all');
  });
  it('borders map to CSS on a separate overlay', () => {
    const c = styleToCss({
      border: { top: { style: 'thin' }, right: { style: 'medium', color: '#00FF00' }, bottom: { style: 'double', color: '#DC2626' }, left: { style: 'dashed' } },
    });
    expect(c.border).toEqual({ borderTop: '1px solid #000000', borderRight: '2px solid #00FF00', borderBottom: '3px double #DC2626', borderLeft: '1px dashed #000000' });
    expect(styleToCss({ border: {} }).border).toBeUndefined();
    expect(styleToCss({ border: { top: { style: 'thick' } } }).border).toEqual({ borderTop: '3px solid #000000' });
  });
  it('number format is passed through', () => {
    expect(styleToCss({ numFmt: 'yyyy/m/d' }).numFmt).toBe('yyyy/m/d');
    expect(styleToCss({ numFmt: '' }).numFmt).toBeUndefined();
  });
  it('makeStyleCache caches per id and tolerates bad ids', () => {
    const styles: XStyle[] = [{}, { bold: true }];
    const cssOf = makeStyleCache(styles);
    expect(cssOf(1)).toBe(cssOf(1));
    expect(cssOf(1).box.fontWeight).toBe(700);
    expect(cssOf(undefined)).toBe(cssOf(0));
    expect(cssOf(99).box).toEqual({});
    expect(cssOf(-1).box).toEqual({});
  });
});

describe('Excel colours', () => {
  it('argb / rgb / theme / indexed', () => {
    expect(resolveExcelColor({ argb: 'FFFF0000' })).toBe('#FF0000');
    expect(resolveExcelColor({ rgb: '00ff00' })).toBe('#00FF00');
    expect(resolveExcelColor({ theme: 4 })).toBe('#' + OFFICE_2013_THEME[4]);
    expect(resolveExcelColor({ indexed: 2 })).toBe('#FF0000');
    expect(resolveExcelColor({ indexed: 64 })).toBe('#000000');
    expect(resolveExcelColor({ indexed: 65 })).toBe('#FFFFFF');
    expect(resolveExcelColor({ argb: 'zz' })).toBeUndefined();
    expect(resolveExcelColor(undefined)).toBeUndefined();
  });
  it('tint (Excel HLS algorithm)', () => {
    // accent1 4472C4 lighter 80% / darker 25% as shown in Excel's palette
    expect(applyTint('4472C4', 0.7999816888943144)).toBe('DAE3F3');
    expect(applyTint('4472C4', -0.249977111117893)).toBe('2F5597');
    // white darker 5% -> F2F2F2
    expect(applyTint('FFFFFF', -0.0499893185216834)).toBe('F2F2F2');
    expect(resolveExcelColor({ theme: 0, tint: -0.0499893185216834 })).toBe('#F2F2F2');
  });
});

describe('units', () => {
  it('column widths (ECMA-376): Calibri 8.43 = 64 px, 游ゴシック 8.38 = 72 px', () => {
    expect(excelWidthToPx(9.140625, 7)).toBe(64);
    expect(excelWidthToPx(9.00390625, 8)).toBe(72);
    expect(pxToExcelWidth(72, 8)).toBe(9.00390625);
    expect(pxToExcelWidth(64, 7)).toBe(9.140625);
    expect(charsToPx(8.38, 8)).toBe(72);
    expect(excelWidthToPx(0)).toBe(0);
    expect(pxToExcelWidth(0)).toBe(0);
    for (let px = 10; px < 400; px++) expect(excelWidthToPx(pxToExcelWidth(px, 8), 8)).toBe(px);
    for (let px = 10; px < 400; px++) expect(excelWidthToPx(pxToExcelWidth(px, 7), 7)).toBe(px);
  });
  it('default column / mdw / row heights', () => {
    expect(defaultColumnPx(8)).toBe(72);
    expect(defaultColumnPx(7)).toBe(64);
    expect(mdwFor('Calibri')).toBe(7);
    expect(mdwFor('游ゴシック')).toBe(8);
    expect(mdwFor('Unknown Font', 22)).toBe(16);
    expect(ptToPx(15)).toBe(20);
    expect(ptToPx(18.75)).toBe(25);
    expect(pxToPt(20)).toBe(15);
  });
});
