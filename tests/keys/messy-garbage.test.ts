/**
 * Round 5 — garbage in, no exception out: random tables full of odd cell values (booleans, NaN / Infinity, huge numbers, dates
 * far away, control characters, lone surrogates, emoji, RTL text, 10,000-character strings, formulas without results),
 * odd headers (blank, duplicate, numeric, very long), ragged rows, merges outside the data, hidden sheets … must analyse and join
 * without throwing, and the result must be internally consistent.
 */
import { describe, expect, it } from 'vitest';
import type { KeySource } from '../../src/core/contracts';
import { addManualCandidate, analyzeSources, defaultPlan, detectRegion, joinSources, pickDataSheet } from '../../src/core/keys';
import { type CellSpec, bookFromRows, sourceOf } from './helpers';

class Rng {
  private s: number;
  constructor(seed: number) {
    this.s = (seed * 2246822519) >>> 0 || 1;
  }
  next(): number {
    this.s = (Math.imul(this.s, 1664525) + 1013904223) >>> 0;
    return this.s / 4294967296;
  }
  int(n: number): number {
    return Math.floor(this.next() * n);
  }
  chance(p: number): boolean {
    return this.next() < p;
  }
  pick<T>(xs: readonly T[]): T {
    return xs[this.int(xs.length)];
  }
}

const HEADERS = ['氏名', '名前', 'ID', '学籍番号', 'メール', '電話', 'フリガナ', '生年月日', '大学', '姓', '名', 'No', '備考', '', '  ', '1', '氏名', 'Name', '合計', '（必須）氏名', 'x'.repeat(500), '\u0000', '😀', 'שלום'];
const ODD: CellSpec[] = [
  true, false, NaN, Infinity, -Infinity, 0, -1, 1e21, 1e-9, 2 ** 53, -(2 ** 53), 0.1 + 0.2, '', ' ', '\t', '\n', '\u0000', '\ud800', '\udfff', '😀', '山田😀太郎', 'שלום עולם', '١٢٣٤', '№5', '#N/A', '#REF!', '=SUM(A1)', "'123", '2.02E+07', '1E+400', '0x1F', 'NaN', 'null', 'undefined', '-', '―', '(旧姓：佐藤)', '（', ')', '[', '山田 太郎 様', 'yamada@', '@', 'a@b.c', '090-1234-567', '+81', '03-', '令和元年', 'R1.5.1', 'H99.13.40', '2001/2/30', '9999/99/99', '山'.repeat(10000),
  new Date(Date.UTC(1900, 0, 1)), new Date(Date.UTC(9999, 11, 31)), new Date(NaN), { v: null, f: 'A1+1' }, { v: 5, f: 'SUM(B2:B9)' }, { v: 'x', f: 'A1&"x"', t: 'z' },
];

function table(rng: Rng): { rows: CellSpec[][]; merges: Array<{ r1: number; c1: number; r2: number; c2: number }> } {
  const nCols = 1 + rng.int(9);
  const nRows = rng.int(40);
  const head: CellSpec[] = Array.from({ length: nCols }, () => (rng.chance(0.85) ? rng.pick(HEADERS) : null));
  const rows: CellSpec[][] = [];
  if (rng.chance(0.3)) rows.push([rng.pick(ODD), null, rng.pick(ODD)]);
  rows.push(head);
  for (let r = 0; r < nRows; r++) {
    const width = rng.chance(0.15) ? rng.int(nCols + 3) : nCols;
    const row: CellSpec[] = [];
    for (let c = 0; c < width; c++) {
      const x = rng.next();
      row.push(x < 0.35 ? rng.pick(ODD) : x < 0.5 ? null : x < 0.7 ? `山田 太郎${rng.int(6)}` : x < 0.85 ? rng.int(1000) : `ID${rng.int(20)}`);
    }
    rows.push(row);
  }
  const merges: Array<{ r1: number; c1: number; r2: number; c2: number }> = [];
  for (let k = rng.int(4); k > 0; k--) {
    const r1 = 1 + rng.int(nRows + 6);
    const c1 = 1 + rng.int(nCols + 2);
    merges.push({ r1, c1, r2: r1 + rng.int(4), c2: c1 + rng.int(4) });
  }
  return { rows, merges };
}

function check(res: ReturnType<typeof joinSources>, ids: string[]): void {
  expect(res.summary.baseRows).toBeGreaterThanOrEqual(0);
  expect(res.rows.length).toBe(res.rows.filter((r) => r.baseRow !== null).length + res.summary.appended);
  for (const r of res.rows) for (const id of ids) expect(id in r.matches).toBe(true);
  for (const it of res.items) {
    expect(typeof it.message).toBe('string');
    expect(['error', 'warning', 'info']).toContain(it.severity);
    for (const n of [...(it.rows ?? []), ...(it.otherRows ?? [])]) expect(Number.isInteger(n) && n > 0).toBe(true);
  }
  for (const id of ids) {
    const rep = res.reports[id];
    expect(rep.matched).toBeGreaterThanOrEqual(0);
    for (const s of rep.suggestions) expect(s.score).toBeGreaterThanOrEqual(0.6);
  }
}

describe('random garbage tables', () => {
  it('200 random pairs / triples: analyse + join (left / full, every duplicate policy) never throw and stay consistent', () => {
    for (let seed = 1; seed <= 200; seed++) {
      const rng = new Rng(seed);
      const n = 2 + rng.int(2);
      const srcs: KeySource[] = [];
      for (let k = 0; k < n; k++) {
        const t = table(rng);
        const book = bookFromRows(`g${k}.xlsx`, [{ name: rng.chance(0.2) ? '' : 'S', rows: t.rows, merges: t.merges, state: rng.chance(0.1) ? 'hidden' : 'visible' }, { name: 'メモ', rows: [['x']] }], `g${k}`);
        srcs.push(sourceOf(book, `g${k}`));
      }
      try {
        const an = analyzeSources(srcs[0], srcs.slice(1));
        const base = defaultPlan(an);
        for (const joinType of ['left', 'full'] as const) {
          for (const duplicatePolicy of ['first', 'last', 'skip'] as const) {
            check(joinSources(an, { ...base, joinType, duplicatePolicy }), srcs.slice(1).map((s) => s.id));
          }
        }
        // every candidate of every pair can be used as the only tier
        for (const s of srcs.slice(1)) {
          for (const c of an.pairs[s.id].candidates) check(joinSources(an, { ...base, tiers: { ...base.tiers, [s.id]: [c.id] } }), srcs.slice(1).map((x) => x.id));
        }
        // manual candidates over arbitrary columns
        const other = an.pairs[srcs[1].id];
        const nb = an.base.columns.length;
        const no = other.other.columns.length;
        if (nb && no) {
          const a1 = an.base.columns[rng.int(nb)].col;
          const b1 = other.other.columns[rng.int(no)].col;
          const an2 = addManualCandidate(an, srcs[1].id, [a1], [b1]);
          check(joinSources(an2, { ...base, tiers: { ...base.tiers, [srcs[1].id]: [`manual:${a1}|${b1}`] } }), srcs.slice(1).map((x) => x.id));
        }
      } catch (e) {
        throw new Error(`seed ${seed}: ${(e as Error).stack ?? e}`);
      }
    }
  });

  it('detectRegion / pickDataSheet on garbage sheets and options never throw', () => {
    for (let seed = 1; seed <= 200; seed++) {
      const rng = new Rng(seed * 7);
      const t = table(rng);
      const book = bookFromRows('g.xlsx', [{ name: 'S', rows: t.rows, merges: t.merges }], 'g');
      const sheet = pickDataSheet(book);
      const r = detectRegion(book, sheet);
      expect(r.firstCol).toBeGreaterThanOrEqual(1);
      expect(r.lastCol).toBeGreaterThanOrEqual(r.firstCol);
      expect(r.lastDataRow).toBeGreaterThanOrEqual(r.headerRow);
      for (const opts of [{ headerRow: 1 + rng.int(50) }, { lastDataRow: rng.int(80) - 5 }, { headerRow: 0 }, { headerRow: NaN }, { lastDataRow: Infinity }, { headerRow: 1e9 }]) {
        expect(() => detectRegion(book, sheet, opts as never)).not.toThrow();
      }
    }
  });
});
