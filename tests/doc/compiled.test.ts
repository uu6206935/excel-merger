// The doc's compiled-formula fast path must be EXACTLY excelRef.adjustFormulaSeq (op by op), and fill-down
// neighbours reused through compileWithHint must equal a fresh tokenization.
import { describe, expect, it } from 'vitest';
import { adjustFormulaSeq, translateFormula, type StructOp } from '../../src/core/formula/excelRef';
import { adjustCompiled, compile, compileWithHint, render } from '../../src/core/doc/compiled';

function rng(seed: number) {
  let s = seed >>> 0;
  return () => {
    s = (s * 1664525 + 1013904223) >>> 0;
    return s / 4294967296;
  };
}

const SHEETS = ['Data', 'data', "'データ 表'", 'Other', "'O''Brien'", '[1]Data', 'Data:Other'];

function randomFormula(r: () => number): string {
  const pick = <T,>(a: T[]) => a[Math.floor(r() * a.length)];
  const col = () => {
    const n = 1 + Math.floor(r() * 30);
    return (r() < 0.3 ? '$' : '') + String.fromCharCode(64 + Math.min(n, 26)) + (n > 26 ? '' : '');
  };
  const row = () => (r() < 0.3 ? '$' : '') + String(1 + Math.floor(r() * 40));
  const cell = () => col() + row();
  const ref = (): string => {
    const k = r();
    let body: string;
    if (k < 0.35) body = cell();
    else if (k < 0.7) body = `${cell()}:${cell()}`;
    else if (k < 0.85) body = `${col()}:${col()}`;
    else if (k < 0.93) body = `${row()}:${row()}`;
    else body = r() < 0.5 ? 'XFD5' : 'A1:XFD1';
    const pre = r() < 0.5 ? '' : `${pick(SHEETS)}!`;
    return pre + body;
  };
  const term = (): string => {
    const k = r();
    if (k < 0.5) return ref();
    if (k < 0.6) return String(Math.floor(r() * 100));
    if (k < 0.7) return '"A1&B2"';
    if (k < 0.8) return `SUM(${ref()},${ref()})`;
    if (k < 0.9) return `VLOOKUP(${ref()},${ref()},3,FALSE)`;
    return `_xlfn.XLOOKUP(${ref()},${ref()},${ref()})`;
  };
  const n = 1 + Math.floor(r() * 4);
  const parts: string[] = [];
  for (let i = 0; i < n; i++) parts.push(term());
  return parts.join(pick(['+', '&', ',', ' ', '*']));
}

function randomOps(r: () => number): StructOp[] {
  const n = 1 + Math.floor(r() * 3);
  const ops: StructOp[] = [];
  for (let i = 0; i < n; i++) {
    const k = r();
    const a = 1 + Math.floor(r() * 25);
    const c = 1 + Math.floor(r() * 4);
    if (k < 0.4) ops.push({ type: 'move', axis: 'col', from: a, count: c, to: 1 + Math.floor(r() * 32) });
    else if (k < 0.7) ops.push({ type: 'insert', axis: 'col', at: a, count: c });
    else ops.push({ type: 'delete', axis: 'col', at: a, count: c });
  }
  return ops;
}

describe('compiled formulas', () => {
  it('adjustCompiled == excelRef.adjustFormulaSeq on 20,000 random formulas / op sequences', () => {
    const r = rng(20260930);
    const fails: string[] = [];
    for (let i = 0; i < 20000; i++) {
      const f = randomFormula(r);
      const ops = randomOps(r);
      const host = r() < 0.5 ? 'Data' : 'Other';
      const c = compile(f);
      expect(render(c)).toBe(f);
      const out = adjustCompiled(c, ops, host.toLowerCase(), 'data');
      const want = adjustFormulaSeq(f, ops, { hostSheet: host, targetSheet: 'Data' });
      const got = out ? out.text : f;
      const gotErr = out ? out.createdRefError : false;
      if (got !== want.formula || gotErr !== want.createdRefError) fails.push(`${host}: ${f} ${JSON.stringify(ops)} -> ${want.formula} (${want.createdRefError}) got ${got} (${gotErr})`);
      // the output's compiled form renders its own text (it is reused by the next op)
      if (out) expect(render(out.compiled)).toBe(out.text);
      if (fails.length > 10) break;
    }
    expect(fails).toEqual([]);
  });

  it('compileWithHint reuses a filled-down neighbour exactly like a fresh tokenization', () => {
    const r = rng(7);
    for (let i = 0; i < 5000; i++) {
      const f = randomFormula(r);
      const dr = 1 + Math.floor(r() * 3);
      const g = translateFormula(f, dr, 0);
      const hinted = compileWithHint(g, compile(f), dr, 0);
      const fresh = compile(g);
      expect(render(hinted)).toBe(g);
      expect(hinted.refs.map((x) => x.text)).toEqual(fresh.refs.map((x) => x.text));
      expect(hinted.refs.map((x) => x.area)).toEqual(fresh.refs.map((x) => ({ ...x.area })));
      expect(hinted.parts).toEqual(fresh.parts);
    }
  });

  it('flags structured refs, lookups and unknown tokens', () => {
    expect(compile('SUM(T[a])').sref).toBe(true);
    expect(compile('VLOOKUP(A1,B:D,2,0)').lookup).toBe(true);
    expect(compile('_xlfn.CHOOSECOLS(A1:C3,2)').lookup).toBe(true);
    expect(compile('SUM(A1)').lookup).toBe(false);
    expect(compile('"abc').unsafe).toBe(true);
  });
});
