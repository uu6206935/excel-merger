import { describe, expect, it } from 'vitest';
import {
  JA_BUILTIN_NUMFMT,
  dateToSerial,
  displayValue,
  editText,
  formatCached,
  formatValue,
  generalEditNumber,
  isDateFormat,
  normalizeNumFmt,
} from '../../src/core/display/numfmt';

const d = 45566.5; // 2024-10-01 (Tue) 12:00

// R3 numfmt cases (value, format, expected text, expected colour)
const CASES: Array<[unknown, string | number | undefined, string, string?]> = [
  [1234.5678, 'General', '1234.5678'],
  [1234.5678, undefined, '1234.5678'],
  [0.1 + 0.2, 'General', '0.3'],
  [1234567, '#,##0', '1,234,567'],
  [12345, '"¥"#,##0', '¥12,345'],
  [12345, '¥#,##0', '¥12,345'],
  [12345, '￥#,##0', '￥12,345'],
  [-12345, '¥#,##0;[Red]-¥#,##0', '-¥12,345', '#FF0000'],
  [-12345, String.raw`\¥#,##0;[Red]\-\¥#,##0`, '-¥12,345', '#FF0000'],
  [12345, '[$¥-411]#,##0', '¥12,345'],
  [-5, '#,##0_);[Red](#,##0)', '(5)', '#FF0000'],
  [0.1234, '0.0%', '12.3%'],
  [d, 14, '2024/10/1'],
  [d, 'mm-dd-yy', '2024/10/1'],
  [d, 22, '2024/10/1 12:00'],
  [d, 'yyyy"年"m"月"d"日"', '2024年10月1日'],
  [d, '[$-411]ggge"年"m"月"d"日"', '令和6年10月1日'],
  [d, 28, '令和6年10月1日'],
  [d, 27, 'R6.10.1'],
  [d, '[$-411]gge"年"m"月"d"日"', '令6年10月1日'],
  [d, '[$-411]ggg ee"年"', '令和 06年'],
  [32874, '[$-411]ggge"年"m"月"d"日"', '平成2年1月1日'], // 1990-01-01
  [43585, '[$-411]ggge"年"m"月"d"日"', '平成31年4月30日'], // last day of 平成
  [43586, '[$-411]ggge"年"m"月"d"日"', '令和1年5月1日'], // first day of 令和 (Excel prints 1, not 元)
  [d, 'm/d(aaa)', '10/1(火)'],
  [d, 'yyyy/m/d aaaa', '2024/10/1 火曜日'],
  [d, '[$-F800]dddd, mmmm dd, yyyy', '2024年10月1日'],
  [d, 'yyyy.m.d', '2024.10.1'],
  // Excel (ja-JP) writes General with [DBNum1] in kanji numerals (measured, see numfmt-excel.test.ts): 百二十三, not 123
  [123, '[DBNum1][$-411]General', '百二十三'],
  [123456, '0.00E+00', '1.23E+05'],
  [3000, '#,##0"円"', '3,000円'],
  [-3, '[Blue]#,##0;[Red]-#,##0;0', '-3', '#FF0000'],
  [3, '[Blue]#,##0;[Red]-#,##0;0', '3', '#0000FF'],
  [3, '[Color10]0', '3', '#008000'],
  [3, '[赤]0', '3', '#FF0000'],
  ['テキスト', '@', 'テキスト'],
  ['abc', '#,##0', 'abc'],
  [true, undefined, 'TRUE'],
  [false, '0', 'FALSE'],
  ['#REF!', undefined, '#REF!'],
  [1234567, '"〒"000-0000', '〒123-4567'],
  [0.5, 'h:mm', '12:00'],
  [1.5, '[h]:mm', '36:00'],
  [null, 'General', ''],
  ['', '0', ''],
];

describe('formatValue (SSF + ja-JP shim)', () => {
  for (const [v, f, want, color] of CASES) {
    it(`${JSON.stringify(v)} with ${JSON.stringify(f)} -> ${want}`, () => {
      const r = formatValue(v, f);
      expect(r.text).toBe(want);
      if (color !== undefined) expect(r.color).toBe(color);
    });
  }
  it('kinds drive Excel General alignment', () => {
    expect(formatValue(1, 'General').kind).toBe('number');
    expect(formatValue('a', 'General').kind).toBe('text');
    expect(formatValue(true).kind).toBe('bool');
    expect(formatValue('#N/A').kind).toBe('error');
    expect(formatValue(null).kind).toBe('empty');
    expect(formatValue(Number.NaN).kind).toBe('error');
  });
  it('never throws on garbage formats', () => {
    for (const f of ['[', '"', '0;0;0;0;0', '\\', 'yyyy"', '[$-]', '[Color99]0', '#,##0.0.0']) {
      expect(() => formatValue(1234.5, f)).not.toThrow();
      expect(typeof formatValue(1234.5, f).text).toBe('string');
    }
  });
  it('Date objects are converted to serials', () => {
    expect(formatValue(new Date(2024, 9, 1), 'yyyy/m/d').text).toBe('2024/10/1');
    expect(dateToSerial(new Date(2024, 9, 1, 12))).toBe(d);
  });
});

describe('normalizeNumFmt / JA_BUILTIN_NUMFMT', () => {
  it('built-in ids resolve to the ja-JP codes', () => {
    expect(JA_BUILTIN_NUMFMT[14]).toBe('yyyy/m/d');
    expect(normalizeNumFmt(14)).toBe('yyyy/m/d');
    expect(normalizeNumFmt(22)).toBe('yyyy/m/d h:mm');
    expect(normalizeNumFmt(49)).toBe('@');
    expect(normalizeNumFmt(999)).toBe('General');
    expect(normalizeNumFmt(undefined)).toBe('General');
    expect(normalizeNumFmt('')).toBe('General');
    expect(normalizeNumFmt('mm-dd-yy')).toBe('yyyy/m/d');
    expect(normalizeNumFmt('[$-F400]h:mm:ss AM/PM')).toBe('h:mm:ss');
  });
});

describe('isDateFormat', () => {
  it('detects dates and times, including 和暦-only codes', () => {
    for (const f of [14, 22, 27, 'yyyy/m/d', 'h:mm', '[h]:mm:ss', '[$-411]ggge"年"m"月"d"日"', 'm"月"d"日"', 'yyyy.m.d']) {
      expect(isDateFormat(f), String(f)).toBe(true);
    }
    for (const f of [0, 49, 'General', '@', '#,##0', '0.00E+00', '"円"0', '0.0%', '[Red]0', undefined]) {
      expect(isDateFormat(f), String(f)).toBe(false);
    }
  });
});

describe('displayValue (contract cell types)', () => {
  it('text cells are never errors; error cells always are', () => {
    expect(displayValue({ t: 's', v: '#N/A' })).toEqual({ text: '#N/A', kind: 'text' });
    expect(displayValue({ t: 'e', v: '#DIV/0!' })).toEqual({ text: '#DIV/0!', kind: 'error' });
  });
  it('date cells without a date format still show a date', () => {
    expect(displayValue({ t: 'd', v: 45566 }).text).toBe('2024/10/1');
    expect(displayValue({ t: 'd', v: d }).text).toBe('2024/10/1 12:00');
    expect(displayValue({ t: 'd', v: 45566 }, '[$-411]ggge"年"m"月"d"日"').text).toBe('令和6年10月1日');
  });
  it('booleans, empties, numbers', () => {
    expect(displayValue({ t: 'b', v: true }).text).toBe('TRUE');
    expect(displayValue({ t: 'b', v: false }).text).toBe('FALSE');
    expect(displayValue({ t: 'z', v: null }).kind).toBe('empty');
    expect(displayValue(undefined).kind).toBe('empty');
    expect(displayValue({ t: 'n', v: 1234.5 }, '#,##0.0').text).toBe('1,234.5');
    expect(displayValue({ t: 's', v: '00123' }, '@').text).toBe('00123');
  });
});

describe('editText (formula bar / F2 text of a constant)', () => {
  it('dates, times, percentages, numbers, text, booleans', () => {
    expect(editText({ t: 'd', v: 45566 }, 'yyyy"年"m"月"d"日"')).toBe('2024/10/1');
    expect(editText({ t: 'n', v: 45566 }, 14)).toBe('2024/10/1');
    expect(editText({ t: 'd', v: d }, 22)).toBe('2024/10/1 12:00:00');
    expect(editText({ t: 'n', v: 0.5 }, 'h:mm')).toBe('12:00:00');
    expect(editText({ t: 'n', v: 0.125 }, '0.0%')).toBe('12.5%');
    expect(editText({ t: 'n', v: 1234567.891 }, '#,##0')).toBe('1234567.891');
    expect(editText({ t: 'n', v: 0.1 + 0.2 })).toBe('0.3');
    expect(editText({ t: 's', v: '00123' })).toBe('00123');
    expect(editText({ t: 'b', v: true })).toBe('TRUE');
    expect(editText({ t: 'z', v: null })).toBe('');
    expect(editText(undefined)).toBe('');
  });
  it('generalEditNumber: 15 significant digits, Excel exponent style', () => {
    expect(generalEditNumber(123456789012345678)).toBe('123456789012346000');
    expect(generalEditNumber(1e21)).toBe('1E+21');
    expect(generalEditNumber(1.5e-7)).toBe('1.5E-07');
    expect(generalEditNumber(0)).toBe('0');
    expect(generalEditNumber(-2.5)).toBe('-2.5');
  });
});

describe('formatCached', () => {
  it('returns the same results as formatValue and is fast', () => {
    for (const [v, f, want] of CASES) expect(formatCached(v, f).text).toBe(want);
    const t0 = performance.now();
    for (let i = 0; i < 200_000; i++) formatCached(40000 + (i % 500), 'yyyy/m/d');
    expect(performance.now() - t0).toBeLessThan(1500);
  });
});
