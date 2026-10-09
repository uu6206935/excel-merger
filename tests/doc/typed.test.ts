/**
 * parseInput vs Excel: every case of fixtures/typed-excel.json was typed into a cell of the given number format by
 * Excel 365 (ja-JP) via COM (Range.FormulaLocal) and read back with readBook. parseInput must produce the same
 * value type, value and resulting number format — except for the documented deliberate deviations (parse.ts):
 *  - automatic percent entry (Excel UI option, not applied by COM): a plain number >= 1 typed into a percent cell
 *    is divided by 100;
 *  - integers with more than 15 digits stay text (Excel rounds them).
 */
import { describe, expect, it } from 'vitest';
import fixture from './fixtures/typed-excel.json';
import { formatCategory, parseInput } from '../../src/core/doc/parse';

type Case = [string, number, string, string | number | boolean | null, string | 0];

const NOW = new Date(2026, 8, 30);
const formats = fixture.formats as string[];
const cases = fixture.cases as Case[];

interface Outcome {
  t: string;
  v: string | number | boolean | null;
  code: string;
}

function ours(input: string, fmt: string): Outcome | { skip: string } {
  const numFmt = fmt === 'General' ? undefined : fmt;
  const r = parseInput(input, { numFmt, now: NOW });
  if (r.kind === 'clear') return { t: 'z', v: null, code: fmt };
  if (r.kind === 'formula') return { skip: 'formula' };
  if (r.kind === 'error') return { skip: 'error: ' + r.reason };
  return { t: r.t, v: r.v, code: r.numFmt ?? fmt };
}

/**
 * Excel stores "minutes:seconds.tenths" input as built-in numFmtId 47, which readBook's ja table spells 'mmss.0'
 * (the ECMA-376 spelling); Excel itself shows it as mm:ss.0, which is the code the doc writes (a custom format
 * that displays identically and that the grid's formatter understands).
 */
const MMSS: Record<string, string> = { 'mmss.0': 'mm:ss.0' };

const sameValue =(a: Outcome['v'], b: Outcome['v']): boolean => {
  if (typeof a === 'number' && typeof b === 'number') return a === b || Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b));
  return a === b;
};

describe('parseInput replays what Excel-ja does with typed input', () => {
  it(`matches ${cases.length} measured cases`, () => {
    const bad: string[] = [];
    let deviations = 0;
    for (const [input, fi, t, v, code0] of cases) {
      const fmt = formats[fi];
      const code = code0 === 0 ? fmt : code0;
      const got = ours(input, fmt);
      if ('skip' in got) {
        bad.push(`${JSON.stringify(input)} @ ${fmt}: ${got.skip}`);
        continue;
      }
      let want: Outcome = { t, v, code: MMSS[code] ?? code };
      // deliberate deviation 1: automatic percent entry
      if (formatCategory(fmt) === 'percent' && t === 'n' && typeof v === 'number' && Math.abs(v) >= 1 && code === fmt && /^[\s\d,.+\-()（）０-９，．＋－€]+$/.test(input)) {
        want = { ...want, v: v / 100 };
        deviations++;
      }
      // deliberate deviation 2: > 15 integer digits stay text
      if ((t === 'n' || t === 'd') && got.t === 's' && got.v === input && /\d{16,}/.test(input.replace(/[,，]/g, ''))) {
        deviations++;
        continue;
      }
      if (got.t !== want.t || !sameValue(got.v, want.v) || got.code !== want.code) {
        bad.push(`${JSON.stringify(input)} @ ${fmt}: excel ${want.t}:${String(want.v)} [${want.code}]  ours ${got.t}:${String(got.v)} [${got.code}]`);
      }
    }
    expect(bad.slice(0, 40)).toEqual([]);
    expect(deviations).toBeGreaterThan(0);
  });

  it('covers the inputs the editor spec names (General cell)', () => {
    const g = (s: string) => cases.find((c) => c[0] === s && formats[c[1]] === 'General');
    const expected: Array<[string, string, number, string]> = [
      ['10/20', 'd', 46315, 'm"月"d"日"'],
      ['2026/10/20', 'd', 46315, 'yyyy/m/d'],
      ['R8.10.20', 'd', 46315, '[$-411]ge.m.d'],
      ['令和8年10月20日', 'd', 46315, '[$-411]ggge"年"m"月"d"日"'],
      ['50%', 'n', 0.5, '0%'],
      ['¥1,000', 'n', 1000, '"¥"#,##0;[Red]"¥"\\-#,##0'],
      ['1,000', 'n', 1000, '#,##0'],
    ];
    for (const [input, t, v, code] of expected) {
      const c = g(input);
      expect(c, input).toBeDefined();
      expect([c![2], c![3], c![4]], input).toEqual([t, v, code]);
      const r = parseInput(input, { now: NOW });
      expect(r, input).toEqual({ kind: 'value', t, v, numFmt: code });
    }
  });
});
