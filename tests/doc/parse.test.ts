import { describe, expect, it } from 'vitest';
import { dateSerial, foldWidth, isDateFormat, isPercentFormat, isTextFormat, normalizeFormulaCase, parseInput } from '../../src/core/doc/parse';
import type { ParseContext } from '../../src/core/doc/parse';

const NOW = new Date(2026, 4, 1); // 2026-05-01
const P = (text: string, ctx: ParseContext = {}) => parseInput(text, { now: NOW, ...ctx });
const DATE: ParseContext = { numFmt: 'yyyy/m/d' };

describe('parseInput', () => {
  it('clears on empty input', () => {
    expect(P('')).toEqual({ kind: 'clear' });
  });

  it('parses numbers like Excel (half / full width, separators, currency, sign, exponent)', () => {
    const v = (text: string, ctx?: ParseContext) => {
      const r = P(text, ctx);
      return r.kind === 'value' ? [r.t, r.v] : r;
    };
    expect(v('123')).toEqual(['n', 123]);
    expect(v('１２３')).toEqual(['n', 123]);
    expect(v('1,234,567')).toEqual(['n', 1234567]);
    expect(v('１，２３４')).toEqual(['n', 1234]);
    expect(v('¥1,000')).toEqual(['n', 1000]);
    expect(v('￥１０００')).toEqual(['n', 1000]);
    expect(v('$12.5')).toEqual(['n', 12.5]);
    expect(v('-5')).toEqual(['n', -5]);
    expect(v('－５')).toEqual(['n', -5]);
    expect(v('(5)')).toEqual(['n', -5]);
    expect(v('+7')).toEqual(['n', 7]);
    expect(v('.5')).toEqual(['n', 0.5]);
    expect(v('1.5e3')).toEqual(['n', 1500]);
    expect(v(' 42 ')).toEqual(['n', 42]);
    expect(v('00123')).toEqual(['n', 123]);
    expect(v('1,23')).toEqual(['s', '1,23']);
    expect(v('12-34')).toEqual(['s', '12-34']);
    expect(v('1.2.3')).toEqual(['s', '1.2.3']);
  });

  it('handles percent input and percent-formatted cells', () => {
    // like Excel, a percentage typed into a General cell gives the cell a percent format
    expect(P('50%')).toEqual({ kind: 'value', t: 'n', v: 0.5, numFmt: '0%' });
    expect(P('1.5%')).toEqual({ kind: 'value', t: 'n', v: 0.015, numFmt: '0.00%' });
    expect(P('50%', { numFmt: '0%' })).toEqual({ kind: 'value', t: 'n', v: 0.5 });
    expect(P('50%', { numFmt: '0.00' })).toEqual({ kind: 'value', t: 'n', v: 0.5 }); // number formats keep everything
    expect(P('50%', { numFmt: 'yyyy/m/d' })).toEqual({ kind: 'value', t: 'n', v: 0.5, numFmt: '0%' });
    expect(P('５０％', { numFmt: '0.0%' })).toEqual({ kind: 'value', t: 'n', v: 0.5 });
    expect(P('50', { numFmt: '0%' })).toEqual({ kind: 'value', t: 'n', v: 0.5 });
  });

  it('parses dates and times into serials in date / time formatted cells', () => {
    expect(P('2026/10/20', DATE)).toEqual({ kind: 'value', t: 'd', v: 46315 });
    expect(P('2026-10-20', DATE)).toEqual({ kind: 'value', t: 'd', v: 46315 });
    expect(P('２０２６／１０／２０', DATE)).toEqual({ kind: 'value', t: 'd', v: 46315 });
    expect(P('2026年10月20日', DATE)).toEqual({ kind: 'value', t: 'd', v: 46315 });
    expect(P('10/20', DATE)).toEqual({ kind: 'value', t: 'd', v: 46315 }); // current year
    expect(P('10月20日', DATE)).toEqual({ kind: 'value', t: 'd', v: 46315 });
    expect(P('R8.10.20', DATE)).toEqual({ kind: 'value', t: 'd', v: 46315 });
    expect(P('r8/10/20', DATE)).toEqual({ kind: 'value', t: 'd', v: 46315 });
    expect(P('令和8年10月20日', DATE)).toEqual({ kind: 'value', t: 'd', v: 46315 });
    expect(P('令和元年5月1日', DATE)).toEqual({ kind: 'value', t: 'd', v: 43586 });
    expect(P('H30.1.5', DATE)).toEqual({ kind: 'value', t: 'd', v: 43105 });
    expect(P('2026/10/20 9:30', { numFmt: 'yyyy/m/d h:mm' })).toEqual({ kind: 'value', t: 'd', v: 46315 + 9.5 / 24 });
    const t = P('9:30', { numFmt: 'h:mm' });
    expect(t.kind === 'value' && t.t).toBe('d');
    expect(t.kind === 'value' && (t.v as number)).toBeCloseTo(9.5 / 24, 12);
    expect(P('46315', DATE)).toEqual({ kind: 'value', t: 'd', v: 46315 });
  });

  it('rejects impossible dates (kept as text)', () => {
    expect(P('2026/2/30', DATE)).toMatchObject({ t: 's', v: '2026/2/30' });
    expect(P('R0.5.1', DATE)).toMatchObject({ t: 's' });
    expect(P('1899/12/31', DATE)).toMatchObject({ t: 's' });
    expect(P('13/40', DATE)).toMatchObject({ t: 's' });
    // measured: Excel lets an era year run past the era's end / start before its first day
    expect(P('R1.4.30', DATE)).toEqual({ kind: 'value', t: 'd', v: 43585 });
    expect(P('H31.5.1', DATE)).toEqual({ kind: 'value', t: 'd', v: 43586 });
  });

  it('gives a General cell the number format Excel-ja gives it for dates, percentages and currency', () => {
    expect(P('2026/10/20')).toEqual({ kind: 'value', t: 'd', v: 46315, numFmt: 'yyyy/m/d' });
    expect(P('10/20')).toEqual({ kind: 'value', t: 'd', v: 46315, numFmt: 'm"月"d"日"' }); // current year (NOW = 2026-05-01)
    expect(P('R8.10.20')).toEqual({ kind: 'value', t: 'd', v: 46315, numFmt: '[$-411]ge.m.d' });
    expect(P('令和8年10月20日')).toEqual({ kind: 'value', t: 'd', v: 46315, numFmt: '[$-411]ggge"年"m"月"d"日"' });
    expect(P('2026年10月20日')).toEqual({ kind: 'value', t: 'd', v: 46315, numFmt: 'yyyy"年"m"月"d"日"' });
    expect(P('¥1,000')).toEqual({ kind: 'value', t: 'n', v: 1000, numFmt: '"¥"#,##0;[Red]"¥"\\-#,##0' });
    expect(P('1,000')).toEqual({ kind: 'value', t: 'n', v: 1000, numFmt: '#,##0' });
    expect(P('9:30')).toMatchObject({ kind: 'value', t: 'd', numFmt: 'h:mm' });
    // a cell that already shows dates keeps its format
    expect(P('10/20', DATE)).toEqual({ kind: 'value', t: 'd', v: 46315 });
    // plain numbers never change the format
    expect(P('123')).toEqual({ kind: 'value', t: 'n', v: 123 });
  });

  it('parses booleans, text and the quote prefix', () => {
    expect(P('TRUE')).toEqual({ kind: 'value', t: 'b', v: true });
    expect(P('false')).toEqual({ kind: 'value', t: 'b', v: false });
    expect(P('ＴＲＵＥ')).toEqual({ kind: 'value', t: 'b', v: true });
    expect(P('東京大学')).toEqual({ kind: 'value', t: 's', v: '東京大学' });
    expect(P("'123")).toEqual({ kind: 'value', t: 's', v: '123', quote: true });
    expect(P('   ')).toEqual({ kind: 'value', t: 's', v: '   ' });
  });

  it('keeps text in text-formatted cells and ID columns', () => {
    expect(P('00123', { numFmt: '@' })).toEqual({ kind: 'value', t: 's', v: '00123' });
    expect(P('=A1', { numFmt: '@' })).toEqual({ kind: 'value', t: 's', v: '=A1' });
    expect(P('00123', { textIds: true })).toEqual({ kind: 'value', t: 's', v: '00123', quote: true });
    expect(P('１２３', { textIds: true })).toEqual({ kind: 'value', t: 's', v: '１２３', quote: true });
    expect(P('12.5', { textIds: true })).toEqual({ kind: 'value', t: 'n', v: 12.5 });
  });

  it('keeps integers Excel cannot store exactly as text', () => {
    const r = P('1234567890123456789');
    expect(r).toMatchObject({ t: 's', v: '1234567890123456789', quote: true });
    expect(r.kind === 'value' && r.hint).toMatch(/15 桁/);
  });

  it('turns formulas into FILE form (future functions, table refs, Excel case)', () => {
    expect(P('=sum(a1:a3)')).toEqual({ kind: 'formula', f: 'SUM(A1:A3)' });
    expect(P('=xlookup(a2,a2:a6,b2:b6)')).toEqual({ kind: 'formula', f: '_xlfn.XLOOKUP(A2,A2:A6,B2:B6)' });
    expect(P('=LET(x,1,x+1)')).toEqual({ kind: 'formula', f: '_xlfn.LET(_xlpm.x,1,_xlpm.x+1)' });
    expect(P('=A1#')).toEqual({ kind: 'formula', f: '_xlfn.ANCHORARRAY(A1)' });
    expect(P('=[@点数]*3', { formOptions: { hostTable: '名簿T', tables: ['名簿T'] } })).toEqual({ kind: 'formula', f: '名簿T[[#This Row],[点数]]*3' });
    expect(P('="a"&b2&"ｂ"')).toEqual({ kind: 'formula', f: '"a"&B2&"ｂ"' });
    expect(P("='集計 シート'!$c$5")).toEqual({ kind: 'formula', f: "'集計 シート'!$C$5" });
    expect(P('=if(true,#n/a,1)')).toEqual({ kind: 'formula', f: 'IF(TRUE,#N/A,1)' });
    expect(P('=')).toEqual({ kind: 'value', t: 's', v: '=' });
  });

  it('refuses broken formulas with a Japanese reason', () => {
    for (const bad of ['=SUM(A1', '=SUM(A1))', '=A1+', '="abc', '={1,2']) {
      const r = P(bad);
      expect(r.kind, bad).toBe('error');
      expect(r.kind === 'error' && r.reason).toMatch(/数式/);
    }
    const long = P('x'.repeat(40000));
    expect(long.kind === 'error' && long.reason).toMatch(/32,767/);
  });
});

describe('helpers', () => {
  it('classifies number formats', () => {
    expect(isDateFormat('yyyy/m/d')).toBe(true);
    expect(isDateFormat('[$-ja-JP]ggge"年"m"月"d"日"')).toBe(true);
    expect(isDateFormat('h:mm:ss')).toBe(true);
    expect(isDateFormat('0.00E+00')).toBe(false);
    expect(isDateFormat('#,##0"円"')).toBe(false);
    expect(isDateFormat('General')).toBe(false);
    expect(isDateFormat('[Red]0.0')).toBe(false);
    expect(isTextFormat('@')).toBe(true);
    expect(isTextFormat('0')).toBe(false);
    expect(isPercentFormat('0.0%')).toBe(true);
    expect(isPercentFormat('0"%"')).toBe(false);
  });

  it('folds widths and computes serials', () => {
    expect(foldWidth('ＡＢＣ　１２３－')).toBe('ABC 123-');
    expect(dateSerial(2026, 10, 20)).toBe(46315);
    expect(dateSerial(1900, 1, 1)).toBe(1);
    expect(dateSerial(1900, 2, 29)).toBe(60); // Excel's 1900 leap-year bug
    expect(dateSerial(1900, 3, 1)).toBe(61);
    expect(dateSerial(1899, 12, 31)).toBeNull();
    expect(dateSerial(2026, 2, 29)).toBeNull();
  });

  it('normalizes formula case like Excel without touching strings, names and sheets', () => {
    expect(normalizeFormulaCase('sum(a1:$b$2)+myName+"abc"')).toBe('SUM(A1:$B$2)+myName+"abc"');
    expect(normalizeFormulaCase("'Data sheet'!c5")).toBe("'Data sheet'!C5");
    expect(normalizeFormulaCase('_xlfn.xlookup(a1,b:b,c:c)')).toBe('_xlfn.XLOOKUP(A1,B:B,C:C)');
    expect(normalizeFormulaCase('SUM(A1)')).toBe('SUM(A1)');
  });
});
