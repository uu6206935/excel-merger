/**
 * Random corpus of the includeSourceSheets property test (round 9, task F2).
 *
 * For a seed it writes <dir>/base.xlsx and, per random source workbook, <dir>/nocache/<name>__<order>.xlsx plus <dir>/manifest.json
 * (the layout of the formula-import corpus, see tests/fixtures/formulaimport/README.md) — three row orders of every workbook (the
 * base's order, reversed, shuffled). The source has the data sheet 評価 (ID, 氏名, 点1 点2 点3, 備考 and 8 random formula columns), the
 * other sheets レート (a score -> grade table and a rate) and Master (ID -> 区分, 係数), a rate in a side cell (P1) and a footer total.
 * 3 base people are missing in the source, 3 people exist only in the source (join 'full' appends them). Columns C-E hold DISTINCT
 * numbers, so MATCH / INDEX / RANK do not depend on which of two equal rows comes first.
 *
 * The formulas are drawn from TEMPLATES: same-row arithmetic, side cells, footer, other sheets (lookups into Master / レート), previous
 * rows and running totals, ranges over the whole table, ranges that are only half imported (SUMIF / SUMPRODUCT / INDEX-MATCH over an
 * imported and a non-imported column), ROW(ref) / OFFSET, and a few formulas that must stay values (ROW()). The source stores NO results:
 * the real-Excel oracle (oracle.ts --random) lets Excel compute the source and the merged workbook and compares the people.
 */
import fs from 'node:fs';
import path from 'node:path';
import { buildBook, col, type F, type G } from '../../xlsx/formulaimport-helpers';
import type { CorpusEntry } from './lib';

export function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

export const BASE_N = 14;
export const SRC_COMMON = 11;
export const SRC_EXTRA = 3;
/** the last data row of the source (header row 1, data rows 2..LAST) */
export const LAST = 1 + SRC_COMMON + SRC_EXTRA;
/** the footer total row of the source */
export const FOOT = LAST + 2;
export const FIRST_FORMULA_COL = 7;
export const N_FORMULA_COLS = 8;

interface Ctx {
  r: number;
  /** a numeric column letter (C, D, E or an earlier numeric formula column) */
  c: string;
  /** a second one, to the right of `c` when possible */
  d: string;
  /** a threshold */
  k: number;
  /** an earlier numeric formula column or C */
  f: string;
}
interface Tpl {
  id: string;
  out: 'num' | 'text';
  /**
   * false: reads a range of the whole table, a position, a running range or the footer — in the merged table (round 8: "the table the formulas
   * now belong to") its value legitimately differs from the source's when rows are missing / added, and the reduced twin the oracle compares a
   * widened formula with recalculates it after Excel deleted rows. A formula that READS such a column would differ for that reason, not
   * because of the copies, so such a column is never chained (it is compared itself). 'row-bare' (ROW()) never stays a formula: the source
   * stores no result, so it becomes an empty cell and whatever reads it would see an empty cell.
   */
  stable?: boolean;
  /** null = no formula for this row (a row-2 variant is used instead) */
  make: (x: Ctx) => string | null;
}

const L = LAST;
export const TEMPLATES: Tpl[] = [
  { id: 'add', out: 'num', make: (x) => `${x.c}${x.r}+${x.d}${x.r}` },
  { id: 'mul', out: 'num', make: (x) => `${x.c}${x.r}*${x.d}${x.r}-1` },
  { id: 'if', out: 'text', make: (x) => `IF(${x.c}${x.r}>${x.k},"A","B")` },
  { id: 'round', out: 'num', make: (x) => `ROUND(${x.c}${x.r}/${x.d}${x.r},1)` },
  { id: 'max', out: 'num', make: (x) => `MAX(C${x.r}:E${x.r})` },
  { id: 'avg3', out: 'num', make: (x) => `AVERAGE(C${x.r}:E${x.r})` },
  { id: 'side', out: 'num', make: (x) => `SUM(C${x.r}:D${x.r})*$P$1` },
  { id: 'rate', out: 'num', make: (x) => `${x.c}${x.r}*レート!$D$1` },
  { id: 'vlookup', out: 'text', make: (x) => `VLOOKUP($A${x.r},Master!$A$2:$C$${L},2,FALSE)` },
  { id: 'index-match', out: 'num', make: (x) => `INDEX(Master!$C:$C,MATCH($A${x.r},Master!$A:$A,0))` },
  { id: 'sumif-master', out: 'num', make: (x) => `SUMIF(Master!$A:$A,$A${x.r},Master!$C:$C)` },
  { id: 'countif-master', out: 'num', make: (x) => `COUNTIF(Master!$A:$A,$A${x.r})` },
  { id: 'grade', out: 'num', make: (x) => `IFERROR(VLOOKUP(${x.c}${x.r},レート!$A$2:$B$6,2,TRUE),0)` },
  { id: 'sum-rate', out: 'num', make: (x) => `SUM(レート!$B$2:$B$6)*${x.c}${x.r}` },
  { id: 'lookup-vec', out: 'num', make: (x) => `LOOKUP(${x.c}${x.r},レート!$A$2:$A$6,レート!$B$2:$B$6)` },
  { id: 'index-grade', out: 'num', make: (x) => `INDEX(レート!$B$2:$B$6,MATCH(${x.c}${x.r},レート!$A$2:$A$6,1))` },
  { id: 'prev', out: 'num', make: (x) => (x.r > 2 ? `${x.c}${x.r}-${x.c}${x.r - 1}` : `${x.c}${x.r}`) },
  { id: 'running', out: 'num', stable: false, make: (x) => `SUM(${x.c}$2:${x.c}${x.r})` },
  { id: 'share', out: 'num', stable: false, make: (x) => `${x.c}${x.r}/SUM(${x.c}$2:${x.c}$${L})` },
  { id: 'rank', out: 'num', stable: false, make: (x) => `RANK(${x.c}${x.r},${x.c}$2:${x.c}$${L})` },
  { id: 'countif-gt', out: 'num', stable: false, make: (x) => `COUNTIF(${x.c}$2:${x.c}$${L},">"&${x.c}${x.r})` },
  { id: 'counta', out: 'num', stable: false, make: (x) => `COUNTA($A$2:$A${x.r})` },
  { id: 'footer', out: 'num', stable: false, make: (x) => `${x.c}${x.r}/$C$${FOOT}` },
  { id: 'sumif-half', out: 'num', stable: false, make: (x) => `SUMIF($E$2:$E$${L},">"&${(x.k % 9) + 1},$C$2:$C$${L})` },
  { id: 'sumproduct-half', out: 'num', stable: false, make: () => `SUMPRODUCT($C$2:$C$${L},$E$2:$E$${L})` },
  { id: 'index-match-half', out: 'num', stable: false, make: (x) => `INDEX($C$2:$C$${L},MATCH(E${x.r},$E$2:$E$${L},0))` },
  { id: 'sumproduct-master', out: 'num', stable: false, make: () => `SUMPRODUCT($C$2:$C$${L},Master!$C$2:$C$${L})` },
  { id: 'sumproduct-cond', out: 'num', stable: false, make: (x) => `SUMPRODUCT(--(C$2:C$${L}>${x.k}),D$2:D$${L})` },
  { id: 'sum-master-over', out: 'num', stable: false, make: (x) => `SUM(Master!$C$2:$C$${L})/COUNT(${x.c}$2:${x.c}$${L})` },
  { id: 'match-pos', out: 'num', stable: false, make: (x) => `MATCH(${x.c}${x.r},${x.c}$2:${x.c}$${L},0)` },
  { id: 'exists', out: 'text', make: (x) => `IF(ISNUMBER(MATCH($A${x.r},Master!$A:$A,0)),"有","無")` },
  { id: 'eq-master', out: 'text', make: (x) => `IF($A${x.r}=Master!$A$2,"先頭","他")` },
  { id: 'avg-master', out: 'num', make: (x) => `AVERAGE(Master!$C$2:$C$${L})*${x.c}${x.r}` },
  { id: 'row-ref', out: 'num', make: (x) => `ROW(${x.c}${x.r})` },
  { id: 'offset', out: 'num', make: (x) => `OFFSET(${x.c}${x.r},0,1)` },
  { id: 'sum-offset', out: 'num', make: (x) => `SUM(OFFSET(C${x.r},0,0,1,2))` },
  { id: 'row-bare', out: 'num', stable: false, make: () => 'ROW()-1' },
  { id: 'name', out: 'text', make: (x) => `$A${x.r}&"様"` },
  { id: 'text', out: 'text', make: (x) => `TEXT(${x.c}${x.r},"0.0")&"pt"` },
  { id: 'chain', out: 'num', make: (x) => `${x.f}${x.r}*2` },
];

function shuffle<T>(a: T[], rnd: () => number): T[] {
  const out = a.slice();
  for (let i = out.length - 1; i > 0; i--) {
    const j = Math.floor(rnd() * (i + 1));
    [out[i], out[j]] = [out[j], out[i]];
  }
  return out;
}

const pad = (n: number, w = 3) => String(n).padStart(w, '0');

/**
 * The names of the files and of the other sheets (what planSheetNames makes of them is what the formulas must read, in real Excel):
 * plain, with spaces / parentheses / an apostrophe, a leading digit, and so long that the copies' names are cut to 31 characters and the
 * two long sheet names collide ("(1)" on the one that is added later).
 */
const LONG = 'とても長いレートシート名その一二三四五六七八九十';
export const NAME_VARIANTS: Array<{ stem: string; rate: string; master: string }> = [
  { stem: '', rate: 'レート', master: 'Master' },
  { stem: '面接 評価(最終) ', rate: 'レート 表', master: 'マスタ(旧)' },
  { stem: '令和七年度第三回新卒採用面接評価シート最終版_', rate: "Rate's", master: '2025年度 Master' },
  { stem: "O'Brien ", rate: LONG, master: LONG + 'マスタ' },
  { stem: '', rate: 'レート 表', master: 'Master' },
  { stem: '面接 評価(最終) ', rate: 'レート', master: '2025年度 Master' },
];
/** a sheet name as it is written in a formula */
export const quoteName = (n: string): string => (/^[A-Za-z_\u3040-\u30ff\u4e00-\u9fff][A-Za-z0-9_\u3040-\u30ff\u4e00-\u9fff]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);
const f = (text: string): F => ({ f: text });

/** the formula columns of workbook `w`: the chosen templates with their constants */
export function pickColumns(rnd: () => number): Array<{ tpl: Tpl; c: string; d: string; k: number; f: string }> {
  const numeric: string[] = ['C', 'D', 'E'];
  const out: Array<{ tpl: Tpl; c: string; d: string; k: number; f: string }> = [];
  const ids = shuffle(TEMPLATES, rnd).slice(0, N_FORMULA_COLS);
  ids.forEach((tpl, i) => {
    const c = numeric[Math.floor(rnd() * numeric.length)];
    const rest = numeric.filter((x) => x > c);
    const d = rest.length ? rest[Math.floor(rnd() * rest.length)] : numeric.find((x) => x !== c) ?? 'D';
    const earlier = numeric.filter((x) => x > 'E');
    out.push({ tpl, c, d, k: 30 + Math.floor(rnd() * 50), f: earlier.length ? earlier[Math.floor(rnd() * earlier.length)] : 'C' });
    if (tpl.out === 'num' && tpl.stable !== false) numeric.push(col(FIRST_FORMULA_COL + i));
  });
  return out;
}

export async function writeRandomCorpus(dir: string, count: number, seed: number): Promise<CorpusEntry[]> {
  fs.mkdirSync(path.join(dir, 'nocache'), { recursive: true });
  const baseIds = Array.from({ length: BASE_N }, (_x, i) => `S${pad(i + 1)}`);
  const baseRows: G[][] = [['ID', '氏名', '大学'], ...baseIds.map((id, i) => [id, `氏名${id}`, `大学${(i % 4) + 1}`])];
  fs.writeFileSync(path.join(dir, 'base.xlsx'), await buildBook({ sheet: '名簿', rows: baseRows }));
  const manifest: CorpusEntry[] = [];
  for (let n = 0; n < count; n++) {
    const rnd = rng(seed * 1000 + n);
    // who is in the source: SRC_COMMON of the base's people + SRC_EXTRA people the base does not have
    const common = shuffle(baseIds, rnd).slice(0, SRC_COMMON).sort();
    const extras = Array.from({ length: SRC_EXTRA }, (_x, i) => `X${pad(n * 10 + i + 1, 2)}`);
    const people = [...common, ...extras];
    const cols = pickColumns(rnd);
    const distinct = (lo: number, hi: number) => shuffle(Array.from({ length: hi - lo + 1 }, (_x, i) => lo + i), rnd).slice(0, people.length);
    const c = distinct(40, 99), d = distinct(20, 79), e = distinct(1, 14);
    const orders: Array<'same' | 'rev' | 'shuf'> = ['same', 'rev', 'shuf'];
    const nv = NAME_VARIANTS[n % NAME_VARIANTS.length];
    for (const order of orders) {
      const idx = people.map((_p, i) => i);
      const seq = order === 'same' ? idx : order === 'rev' ? idx.slice().reverse() : shuffle(idx, rng(seed * 7919 + n * 31 + 5));
      const header: G[] = ['ID', '氏名', '点1', '点2', '点3', '備考', ...cols.map((_x, i) => `式${i + 1}`)];
      while (header.length < 15) header.push(null);
      header.push(1.5); // P1: a rate in a side cell
      const rows: G[][] = [header];
      seq.forEach((pi, i) => {
        const r = i + 2;
        const id = people[pi];
        const row: G[] = [id, `氏名${id}`, c[pi], d[pi], e[pi], `備考${i}`];
        for (const k of cols) {
          const text = k.tpl.make({ r, c: k.c, d: k.d, k: k.k, f: k.f });
          // the templates say レート! / Master!: the sheets of this workbook are called nv.rate / nv.master
          row.push(text == null ? null : f(text.replace(/(レート|Master)!/g, (_m, nm: string) => `${quoteName(nm === 'レート' ? nv.rate : nv.master)}!`)));
        }
        rows.push(row);
      });
      rows.push([]);
      rows.push(['合計', null, f(`SUM(C2:C${LAST})`)]);
      const master: G[][] = [['ID', '区分', '係数'], ...people.slice().reverse().map((id, i) => [id, i % 2 ? 'A' : 'B', 1 + (i % 3)])];
      const rates: G[][] = [['点', '評価', null, 1.1], [0, 1], [20, 2], [40, 3], [60, 4], [80, 5]];
      const file = `${nv.stem}rnd${pad(n, 2)}__${order}.xlsx`;
      fs.writeFileSync(path.join(dir, 'nocache', file), await buildBook({ sheet: '評価', rows, extra: [{ name: nv.rate, rows: rates }, { name: nv.master, rows: master }] }));
      manifest.push({
        name: `rnd${pad(n, 2)}`, order, file, desc: cols.map((x) => x.tpl.id).join(' '), sheet: '評価', headerRow: 1, firstDataRow: 2, lastDataRow: LAST,
      });
    }
  }
  fs.writeFileSync(path.join(dir, 'manifest.json'), JSON.stringify(manifest, null, 1));
  return manifest;
}
