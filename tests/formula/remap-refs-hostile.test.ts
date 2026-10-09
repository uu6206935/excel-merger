/**
 * FormulaRemapper: sheet prefixes (other sheets, 3D, external, Japanese / quoted names), hostile and odd formulas, a seeded fuzz and the
 * speed requirement (100,000 formulas in well under 5 s).
 */
import { describe, expect, it } from 'vitest';
import { FormulaRemapper, type RemapReason } from '../../src/core/formula/remap';
import { tokenize } from '../../src/core/formula/excelRef';
import { identity, ok, reversed, run, spec, why } from './remapHelpers';

describe('sheet prefixes', () => {
  it.each<[string, RemapReason]>([
    ['Master!A1', 'otherSheet'],
    ["'マスタ'!B3", 'otherSheet'],
    ['Sheet2!$B$3*C5', 'otherSheet'],
    ["'It''s'!A1", 'otherSheet'],
    ["SUM('Q1 2026'!C2:C11)", 'otherSheet'],
    ['SUM(Sheet2!A:A)', 'otherSheet'],
    ['評価:Sheet2!B3', 'otherSheet'],
    ["SUM('評価:Sheet2'!B3)", 'otherSheet'],
    ['VLOOKUP(C5,Master!$A$2:$B$100,2,FALSE)', 'otherSheet'],
    ['IFERROR(VLOOKUP(C5,Master!$A:$B,2,0),"未登録")', 'otherSheet'],
    ['[1]Sheet1!A1', 'unsafe'],
    ["'[book.xlsx]S'!A1", 'unsafe'],
    ["'C:\\data\\[other.xlsx]Sheet1'!A1", 'unsafe'],
    ['SUM([2]Data!$A$1:$A$9)', 'unsafe'],
    ['Sheet2!#REF!', 'unsupported'],
    ['#REF!', 'unsupported'],
    ['SUM(C5,#REF!)', 'unsupported'],
    ['C5+Sheet2!C5', 'otherSheet'],
  ])('%s -> %s', (f, want) => {
    expect(why(f)).toBe(want);
  });

  it('the data sheet itself, quoted or not, any case', () => {
    expect(ok("'評価'!C5")).toBe('E5');
    expect(ok('評価!C5')).toBe('E5');
    expect(ok('SUM(評価!C2:C11)')).toBe('SUM(E2:E11)');
    expect(ok('DATA!C5', 5, 5, { sheet: 'data' })).toBe('E5');
    expect(ok('data!C5', 5, 5, { sheet: 'DATA' })).toBe('E5');
    expect(ok("'It''s'!C5", 5, 5, { sheet: "It's" })).toBe('E5');
    expect(ok("'データ 2026'!C5+'データ 2026'!D5", 5, 5, { sheet: 'データ 2026' })).toBe('E5+F5');
    expect(ok("'Q1-2026'!C5", 5, 5, { sheet: 'Q1-2026' })).toBe('E5');
  });

  it('an unqualified reference is a reference to the data sheet', () => {
    expect(ok('C5')).toBe('E5');
  });
});

describe('odd and hostile formulas', () => {
  it('empty, constant-only and reference-free formulas stay as they are', () => {
    expect(ok('')).toBe('');
    expect(ok('1+1')).toBe('1+1');
    expect(ok('"text"')).toBe('"text"');
    expect(ok('3.14*2')).toBe('3.14*2');
    expect(ok('1E3')).toBe('1E3');
    expect(ok('.5')).toBe('.5');
    expect(ok('TRUE')).toBe('TRUE');
    expect(ok('{1,2;3,4}')).toBe('{1,2;3,4}');
  });

  it('strings keep their content, including doubled quotes and reference-looking text', () => {
    expect(ok('"a""b"&C5')).toBe('"a""b"&E5');
    expect(ok('"C5:D5"&C5')).toBe('"C5:D5"&E5');
    expect(ok('"評価!C5"&C5')).toBe('"評価!C5"&E5');
    expect(ok('SUBSTITUTE(C5,"C5","D5")')).toBe('SUBSTITUTE(E5,"C5","D5")');
    expect(ok('"日本語の文字列：Ｃ５"&C5')).toBe('"日本語の文字列：Ｃ５"&E5');
  });

  it('whitespace, line feeds and the intersection operator are preserved', () => {
    expect(ok('C5 +\n D5')).toBe('E5 +\n F5');
    expect(ok('C5\t+\tD5')).toBe('E5\t+\tF5');
    expect(ok('SUM(C5:D5 D5:E5)')).toBe('SUM(E5:F5 F5:G5)');
    expect(ok('  C5  ')).toBe('  E5  ');
  });

  it('unterminated strings, unknown characters and full-width operators / digits cannot be remapped', () => {
    expect(why('"unterminated')).toBe('unsupported');
    expect(why('C5＋1')).toBe('unsupported');
    expect(why('C5 ? D5')).toBe('unsupported');
    expect(why('C5`')).toBe('unsupported');
    expect(why('Ｃ５')).toBe('unresolvedName');
    expect(why('SUM(Ａ１:Ａ９)')).toBe('unresolvedName');
  });

  it('unbalanced brackets are copied as they are (the tokens decide, not the nesting)', () => {
    expect(ok('SUM(C5')).toBe('SUM(E5');
    expect(ok('C5)')).toBe('E5)');
    expect(ok('((C5')).toBe('((E5');
  });

  it('a leading = is just an operator character here (FILE form has none)', () => {
    expect(ok('=C5')).toBe('=E5');
  });

  it('the spill operator and the implicit-intersection operator', () => {
    expect(why('C2#')).toBe('array');
    expect(ok('@C5')).toBe('@E5');
    expect(ok('_xlfn.SINGLE(C2:C11)')).toBe('_xlfn.SINGLE(E2:E11)');
  });

  it('very long formulas and deep nesting', () => {
    const long = Array.from({ length: 1500 }, () => 'C5').join('+');
    expect(ok(long)).toBe(Array.from({ length: 1500 }, () => 'E5').join('+'));
    const deep = 'SUM('.repeat(150) + 'C5' + ')'.repeat(150);
    expect(ok(deep)).toBe('SUM('.repeat(150) + 'E5' + ')'.repeat(150));
    const many = 'SUM(' + Array.from({ length: 200 }, (_x, i) => `C${2 + (i % 10)}`).join(',') + ')';
    expect(why(many)).toBe('ok');
  });

  it('a result longer than Excel\'s 8,192 characters is not kept (a name inlined many times)', () => {
    const names = [{ name: 'Rng', scope: null, refersTo: '評価!$C$2:$C$11' }];
    const f = Array.from({ length: 900 }, () => 'SUM(Rng)').join('+'); // 900 x ("Rng" -> "$E$2:$E$11") grows the formula past the limit
    expect(f.length).toBeLessThan(8192);
    expect(why(f, 5, 5, { names })).toBe('unsupported');
    expect(why(f.slice(0, f.length / 2).replace(/\+$/, ''), 5, 5, { names })).toBe('ok');
  });

  it('rows beyond Excel\'s limits and impossible mappings never throw', () => {
    expect(why('C1048576')).toBe('externalCell');
    expect(why('SUM(C1:C1048576)')).toBe('ok');
    expect(why('XFD5')).toBe('externalCell');
    expect(() => run('A1:XFD1048576')).not.toThrow();
    expect(why('A1:XFD1048576')).toBe('notImported'); // the whole sheet: columns that are not imported
    expect(why('1:1048576')).toBe('externalCell');
  });

  it('a spec with nothing imported only keeps constant formulas', () => {
    expect(why('C5', 5, 5, { colMap: new Map() })).toBe('notImported');
    expect(why('1+1', 5, 5, { colMap: new Map() })).toBe('ok');
  });

  it('a spec with no rows placed: only the own row exists', () => {
    expect(ok('C5+D5', 5, 5, { rowPairs: [] })).toBe('E5+F5');
    expect(why('C4', 5, 5, { rowPairs: [] })).toBe('rowRelation');
    expect(why('SUM(C2:C11)', 5, 5, { rowPairs: [] })).toBe('ok');
  });

  it('the result never contains a sheet prefix of the source and always tokenizes without unknown tokens', () => {
    for (const f of ["SUM('評価'!C2:C11)", '評価!$C$2', "'評価'!C5:E5", 'IF(評価!C5>1,評価!D5,0)']) {
      const out = ok(f);
      expect(out).not.toContain('評価');
      expect(tokenize(out).every((t) => t.type !== 'unknown')).toBe(true);
      expect(tokenize(out).map((t) => t.text).join('')).toBe(out);
    }
  });
});

describe('seeded fuzz: remapped formulas stay well-formed and only read merged columns', () => {
  function rng(seed: number) {
    let a = seed >>> 0;
    return () => {
      a = (a + 0x6d2b79f5) >>> 0;
      let t = Math.imul(a ^ (a >>> 15), 1 | a);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }
  it('1,500 random formulas over a reversed order, partners and appended rows', () => {
    const r = rng(20261009);
    const pick = <T,>(xs: T[]): T => xs[Math.floor(r() * xs.length)];
    const cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'Z'];
    const ref = () => {
      const c1 = pick(cols), c2 = pick(cols);
      const r1 = 1 + Math.floor(r() * 14), r2 = r1 + Math.floor(r() * 6);
      const d = () => (r() < 0.3 ? '$' : '');
      const form = r();
      if (form < 0.45) return `${d()}${c1}${d()}${r1}`;
      if (form < 0.8) return `${d()}${c1}${d()}${r1}:${d()}${c2}${d()}${r2}`;
      if (form < 0.9) return `${c1}:${c2}`;
      return `${pick(['', '評価!', "'評価'!", 'Other!'])}${c1}${r1}`;
    };
    const fn = () => pick(['SUM', 'AVERAGE', 'MAX', 'COUNT', 'RANK', 'IF', 'ROW', 'OFFSET', 'VLOOKUP', 'INDEX', 'ABS']);
    let nOk = 0, nFail = 0;
    const rem = new FormulaRemapper(spec({ rowPairs: reversed(), partnerCols: new Map([[1, 1], [2, 2]]), target: { headerRow: 1, firstDataRow: 2, lastDataRow: 13 } }));
    for (let i = 0; i < 1500; i++) {
      const n = 1 + Math.floor(r() * 3);
      const parts: string[] = [];
      for (let k = 0; k < n; k++) parts.push(r() < 0.5 ? `${fn()}(${ref()},${ref()})` : ref());
      const f = parts.join(pick(['+', '&', '*', '-']));
      const res = rem.remap(f, 2 + Math.floor(r() * 10), 5);
      if (!res.ok) { nFail++; continue; }
      nOk++;
      const toks = tokenize(res.formula);
      expect(toks.some((t) => t.type === 'unknown'), f).toBe(false);
      for (const t of toks) {
        if (t.type !== 'ref') continue;
        expect(t.prefix, `${f} -> ${res.formula}`).toBeUndefined();
        const a = t.area!;
        if (a.kind !== 'rows') for (const c of [a.c1, a.c2]) if (c) expect(c <= 8, `${f} -> ${res.formula}`).toBe(true);
        if (a.kind !== 'cols') for (const rr of [a.r1, a.r2]) expect(rr >= 1 && rr <= 13, `${f} -> ${res.formula}`).toBe(true);
      }
    }
    expect(nOk).toBeGreaterThan(100);
    expect(nFail).toBeGreaterThan(100);
  });
});

describe('speed', () => {
  it('remaps 100,000 imported formulas in well under 5 seconds', () => {
    const N = 100000;
    const pairs: Array<[number, number]> = [];
    for (let r = 2; r <= N + 1; r++) pairs.push([r, N + 3 - r]);
    const t0 = Date.now();
    const rem = new FormulaRemapper({
      sheet: '評価', region: { headerRow: 1, firstDataRow: 2, lastDataRow: N + 1, firstCol: 1, lastCol: 8 }, target: { headerRow: 1, firstDataRow: 2, lastDataRow: N + 1 },
      colMap: new Map([[3, 5], [4, 6], [5, 7], [6, 8]]), rowPairs: pairs,
    });
    let okN = 0;
    for (let i = 0; i < N; i++) {
      const r = 2 + i;
      const forms = [`C${r}+D${r}*2`, `IF(E${r}>=80,"A","B")`, `C${r}/SUM(C$2:C$${N + 1})`, `RANK(C${r},$C$2:$C$${N + 1})`, `VLOOKUP(C${r},$C$2:$F$${N + 1},2,FALSE)`];
      if (rem.remap(forms[i % forms.length], r, N + 3 - r).ok) okN++;
    }
    const ms = Date.now() - t0;
    expect(okN).toBe(N);
    expect(ms).toBeLessThan(5000);
  });

  it('identity helpers stay consistent', () => {
    expect(identity()).toHaveLength(10);
  });
});
