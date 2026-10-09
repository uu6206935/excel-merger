// Tokenizer tests (ported from the R2 prototype): torture list, UI-form extras, grammar fuzzing with known
// ground truth (lossless round trip + exact ref recognition), translate, performance.
// The live differential test against @borgar/fx is dropped (fx is not a dependency of this project); the
// prototype recorded 0 disagreements on every formula fx could tokenize.
import { describe, expect, it } from 'vitest';
import { tokenize, adjustFormula, translateFormula, type Token } from '../../src/core/formula/excelRef';
import { readFixture, Tally } from './helpers';

interface TortureCase { id: string; f: string; refs?: string[]; names?: string[]; sref?: string[]; refsCount?: number }
const cases = readFixture<TortureCase[]>('torture.json');

const eqSet = (a: string[], b: string[]) => JSON.stringify([...a].sort()) === JSON.stringify([...b].sort());
const refsOf = (t: Token[]) => t.filter((x) => x.type === 'ref').map((x) => x.text);
const namesOf = (t: Token[]) => t.filter((x) => x.type === 'name').map((x) => x.text);
const srefsOf = (t: Token[]) => t.filter((x) => x.type === 'sref').map((x) => x.text);
const split2 = (arr: string[]) => arr.flatMap((r) => { const m = /^([^']+![^:!]+):(.+!.+)$/.exec(r); return m ? [m[1], m[2]] : [r]; });

describe('tokenizer: torture list (84 cases)', () => {
  it('round-trips, has no unknown tokens and finds exactly the expected refs / names / srefs', () => {
    const t = new Tally();
    expect(cases.length).toBe(84);
    for (const c of cases) {
      const toks = tokenize(c.f);
      t.ok(toks.map((x) => x.text).join('') === c.f, `${c.id}: round trip`);
      t.ok(!toks.some((x) => x.type === 'unknown'), () => `${c.id}: unknown token ${JSON.stringify(toks.filter((x) => x.type === 'unknown').map((x) => x.text))}`);
      if (c.id === 'long') { t.ok(refsOf(toks).length === c.refsCount, `long: ${refsOf(toks).length} refs`); continue; }
      const want = split2((c.refs ?? []).map((r) => r.replace(/#$/, '')));
      const got = split2(refsOf(toks));
      if (c.id === 'range-chain') t.ok(got.length === 2, `${c.id}: refs ${JSON.stringify(got)}`);
      else t.ok(eqSet(got, want), () => `${c.id}: refs got ${JSON.stringify(got)} want ${JSON.stringify(want)}`);
      if (c.names) t.ok(eqSet(namesOf(toks), c.names), () => `${c.id}: names got ${JSON.stringify(namesOf(toks))} want ${JSON.stringify(c.names)}`);
      if (c.sref) t.ok(eqSet(srefsOf(toks), c.sref), () => `${c.id}: srefs got ${JSON.stringify(srefsOf(toks))}`);
    }
    expect(t.fails).toEqual([]);
  });

  it('handles UI-form extras (leading =, spill #, quoted external path, 3D, chains)', () => {
    const extra: [string, string[]][] = [
      ["='C:\\Users\\x\\[Book1.xlsx]Sheet1'!$A$1+1", ["'C:\\Users\\x\\[Book1.xlsx]Sheet1'!$A$1"]],
      ['=[Book1.xlsx]Sheet1!A1', ['[Book1.xlsx]Sheet1!A1']],
      ["='[1]Sheet 1'!A1:B2", ["'[1]Sheet 1'!A1:B2"]],
      ['=SUM(A1#)', ['A1']],
      ['=LOG10!B2+SUM(LOG10!C:C)', ['LOG10!B2', 'LOG10!C:C']],
      ['=売上・集計!B2+項目①!C3+Ｓｈｅｅｔ!D4', ['売上・集計!B2', '項目①!C3', 'Ｓｈｅｅｔ!D4']],
      ["=SUM('Sheet1:TAX2025'!A1)", ["'Sheet1:TAX2025'!A1"]],
      ['=A1:B2:C3', ['A1:B2', 'C3']],
      ['=$A$1:$B$2 $B$1:$C$3', ['$A$1:$B$2', '$B$1:$C$3']],
      ['=x:y', ['x:y']], // Excel parses this as columns X:Y
    ];
    const t = new Tally();
    for (const [f, want] of extra) {
      const toks = tokenize(f);
      t.ok(toks.map((x) => x.text).join('') === f, `extra ${f}: round trip`);
      t.ok(eqSet(refsOf(toks), want), () => `extra ${f}: refs got ${JSON.stringify(refsOf(toks))} want ${JSON.stringify(want)}`);
    }
    expect(t.fails).toEqual([]);
  });

  it('is lossless for arbitrary junk (property)', () => {
    let seed = 99;
    const rnd = () => { seed = (seed * 16807) % 2147483647; return seed / 2147483647; };
    const alphabet = ['A', '1', '$', ':', '!', "'", '"', '[', ']', '#', '@', '(', ')', ',', ' ', '集', '計', 'Ａ', '、', '.', '_', 'x', 'REF!', 'TRUE', '{', '}', ';', '-', '%', '\\', '\n', '\u3000'];
    for (let i = 0; i < 3000; i++) {
      const len = 1 + Math.floor(rnd() * 25);
      let f = '';
      for (let k = 0; k < len; k++) f += alphabet[Math.floor(rnd() * alphabet.length)];
      const toks = tokenize(f);
      expect(toks.map((x) => x.text).join('')).toBe(f);
      // tokens are contiguous and non-empty
      let pos = 0;
      for (const tk of toks) { expect(tk.start).toBe(pos); expect(tk.end).toBeGreaterThan(tk.start); pos = tk.end; }
    }
  });
});

describe('tokenizer: grammar fuzzing with ground truth', () => {
  it('recognises exactly the generated refs and adjust never touches non-ref text', () => {
    let seed = 12345;
    const rnd = () => { seed = (seed + 0x6D2B79F5) | 0; let t = Math.imul(seed ^ (seed >>> 15), 1 | seed); t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t; return ((t ^ (t >>> 14)) >>> 0) / 4294967296; }; // mulberry32
    const pick = <T,>(a: T[]) => a[Math.floor(rnd() * a.length)];
    const col = (n: number) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
    const sheets = ['Sheet1', '集計', "'集計 シート'", "'O''Brien'", '売上・集計', "'2024'", "'A1'", 'LOG10', "'データ（1）'", 'Data_2', 'a.b'];
    const d = () => (rnd() < 0.3 ? '$' : '');
    const cell = () => `${d()}${col(1 + Math.floor(rnd() * (rnd() < 0.1 ? 16384 : 60)))}${d()}${1 + Math.floor(rnd() * (rnd() < 0.1 ? 1048576 : 500))}`;
    function genRef(): string {
      const pre = rnd() < 0.5 ? pick(sheets) + '!' : '';
      const k = rnd();
      if (k < 0.4) return pre + cell();
      if (k < 0.7) return pre + cell() + ':' + cell();
      if (k < 0.85) { const a = 1 + Math.floor(rnd() * 30); return pre + `${d()}${col(a)}:${d()}${col(a + Math.floor(rnd() * 3))}`; }
      const r = 1 + Math.floor(rnd() * 99); return pre + `${d()}${r}:${d()}${r + Math.floor(rnd() * 5)}`;
    }
    const names = ['TAX_RATE', '単価', 'my.name', 'A1B', 'XFE1', 'R1C1', '_xlpm.x', 'Total2'];
    const funcs = ['SUM', 'IF', '_xlfn.XLOOKUP', 'VLOOKUP', 'LOG10', 'ATAN2', '_xlfn._xlws.FILTER', 'MYUDF', 'INDEX'];
    function genExpr(depth: number, refs: string[]): string {
      const k = rnd();
      if (depth > 3 || k < 0.35) { const r = genRef(); refs.push(r); return r; }
      if (k < 0.45) return pick(['"A1"', '"C:C"', '"Sheet1!A1"', '"x""A1"""', '1.5', '1E+3', 'TRUE', '#N/A', '{1,2;3,4}', '#REF!']);
      if (k < 0.55) return pick(names);
      if (k < 0.65) return 'Table1[[#This Row],[氏名]]';
      if (k < 0.85) { const n = 1 + Math.floor(rnd() * 3); const args = Array.from({ length: n }, () => genExpr(depth + 1, refs)); return `${pick(funcs)}(${args.join(',')})`; }
      return `${genExpr(depth + 1, refs)}${pick(['+', '-', '*', '&', '=', '<>', ' ', ','])}${genExpr(depth + 1, refs)}`;
    }
    const t = new Tally();
    const seen = new Set<string>();
    const skeleton = (s: string) => tokenize(s).filter((x) => x.type !== 'ref').map((x) => x.text).join('');
    for (let i = 0; i < 20000; i++) {
      const refs: string[] = [];
      const f = genExpr(0, refs);
      if (f.startsWith('{') || f.startsWith('#') || seen.has(f)) continue;
      seen.add(f);
      const toks = tokenize(f);
      t.ok(toks.map((x) => x.text).join('') === f, `fuzz round trip: ${f}`);
      const got = refsOf(toks);
      t.ok(JSON.stringify(got) === JSON.stringify(refs), () => `fuzz refs: ${f}\n      got  ${JSON.stringify(got)}\n      want ${JSON.stringify(refs)}`);
      const adj = adjustFormula(f, { type: 'insert', axis: 'col', at: 2, count: 1 }, { hostSheet: 'Sheet1', targetSheet: 'Sheet1' }).formula;
      t.ok(skeleton(adj) === skeleton(f), `fuzz adjust changed non-ref text: ${f} -> ${adj}`);
    }
    expect(seen.size).toBeGreaterThan(13000);
    expect(t.fails.slice(0, 20)).toEqual([]);
  });
});

describe('translateFormula (relative copy)', () => {
  it('shifts relative parts only; off-sheet becomes #REF!', () => {
    expect(translateFormula('A1+$B1+C$1+$D$1+SUM(E:E)+SUM(2:3)+Sheet2!F5', 1, 2)).toBe('C2+$B2+E$1+$D$1+SUM(G:G)+SUM(3:4)+Sheet2!H6');
    expect(translateFormula('A1', -1, 0)).toBe('#REF!');
    expect(translateFormula('"A1"&A1', 1, 1)).toBe('"A1"&B2');
  });
});

describe('tokenizer performance', () => {
  it('handles an 11,675-char formula with 1,200 refs quickly', () => {
    const long = cases.find((c) => c.id === 'long')!.f;
    let t0 = performance.now();
    for (let i = 0; i < 50; i++) tokenize(long);
    const perTok = (performance.now() - t0) / 50;
    t0 = performance.now();
    for (let i = 0; i < 50; i++) adjustFormula(long, { type: 'move', axis: 'col', from: 3, count: 1, to: 1 }, { hostSheet: 'Sheet1', targetSheet: 'Sheet1' });
    const perAdj = (performance.now() - t0) / 50;
    expect(perTok).toBeLessThan(50);
    expect(perAdj).toBeLessThan(50);
  });
});
