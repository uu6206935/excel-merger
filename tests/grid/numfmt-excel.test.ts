/**
 * The grid shows what Excel shows: number formats compared with the text of the installed ja-JP Excel (Range.Text) for about
 * 100 (value, format) pairs measured once (numfmt-excel-cases.json: era dates, weekdays, AM/PM, [DBNum1..3], fractions, scientific,
 * thousands, conditions, out-of-range dates …). Not measured / not reproduced on purpose: a `*` fill character (depends on the cell
 * width), [DBNum1/2] dates (二〇二六年十月一日 uses 十) and 99999999999.5 in General (SSF rounds it differently).
 */
import { describe, expect, it } from 'vitest';
import { formatValue } from '../../src/core/display/numfmt';
import fixture from './numfmt-excel-cases.json';

interface Case {
  id: string;
  v: number;
  f: string;
  text: string;
}
const cases = (fixture as { cases: Case[] }).cases;

describe('number formats vs Excel (ja-JP, measured)', () => {
  it('has a meaningful table', () => {
    expect(cases.length).toBeGreaterThan(80);
  });

  for (const c of cases) {
    it(`${c.id}: ${c.f}`, () => {
      const got = formatValue(c.v, c.f).text;
      if (c.text.startsWith('#')) {
        // Excel shows ##### when a date / time does not fit its range (negative, > 9999-12-31): the grid shows nothing for it
        expect(got === '' || got.startsWith('#')).toBe(true);
      } else {
        expect(got.trim()).toBe(c.text.trim());
      }
    });
  }

  it('[DBNum1/2/3] and 午前/午後 spot checks', () => {
    expect(formatValue(12345, '[DBNum1]0').text).toBe('一二三四五');
    expect(formatValue(12345, '[DBNum2]0').text).toBe('壱弐参四伍');
    expect(formatValue(12345, '[DBNum3]#,##0').text).toBe('１２,３４５');
    expect(formatValue(0, '[DBNum1]0').text).toBe('〇');
    expect(formatValue(0.75, '[$-411]AM/PM h:mm').text).toBe('午後 6:00');
    expect(formatValue(0.25, '[$-411]h:mm AM/PM').text).toBe('6:00 午前');
    expect(formatValue(0.75, 'h:mm AM/PM').text).toBe('6:00 PM');
    // a text value never goes through the digit substitution
    expect(formatValue('abc123', '[DBNum3]@').text).toBe('abc123');
  });
});
