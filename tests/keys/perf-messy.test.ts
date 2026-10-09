/**
 * Round 5 — performance on big and pathological data (generated in memory; reading the files is the reader's budget):
 *   - a 50,000-row base + 3 other files (analysis + join < 6 s, memory reasonable);
 *   - columns that repeat: 50,000 rows with 10 distinct names (no ID / unique IDs with 20% blanks / every key-like column at
 *     once) must finish with a clear message (check item tooManyDuplicates) instead of hanging or running out of memory;
 *   - wide files whose columns all overlap, one-to-many files, a typo in every row, 100,000-character cells.
 */
import { describe, expect, it } from 'vitest';
import type { JoinResult, KeyAnalysis, KeySource } from '../../src/core/contracts';
import { analyzeSources, defaultPlan, joinSources } from '../../src/core/keys';
import { kanaStrict, normEmail, normGeneric, normId, normName, normOrg, normPhone } from '../../src/core/keys/normalize';
import { type CellSpec, budgetMs, sourceFromRows, timeCpu } from './helpers';

const K1 = [...'山田中村佐藤鈴木高橋渡辺伊藤小林加藤吉田松本井上木村林斎清水森池田橋本阿部石川前田藤原後藤岡田長谷川村上近藤石井坂本遠藤青木'];
const K2 = [...'太郎一花健美咲翔結衣大輝陽菜蓮優子拓也直樹彩誠瞳隆悠真葵駿亮涼智恵雅和紀幸明秀順豊勇博'];
const pick = (xs: string[], n: number) => xs[n % xs.length];
/** 50,000 distinct-ish names (4 kanji drawn from two pools) */
function nameOf(i: number): { sei: string; mei: string } {
  return { sei: pick(K1, i) + pick(K1, Math.floor(i / 40) * 7 + 3), mei: pick(K2, Math.floor(i / 1600) + i * 13) + pick(K2, Math.floor(i / 50) * 11 + 5) };
}
const full = (i: number, sep = '　') => {
  const n = nameOf(i);
  return `${n.sei}${sep}${n.mei}`;
};

/** (result, ms): min of wall clock and CPU time — see timeCpu in helpers.ts */
function time<T>(fn: () => T): [T, number] {
  return timeCpu(fn);
}

/**
 * best of up to `tries` fresh runs (this machine is shared with other sessions; the budget itself is unchanged). It yields to the
 * event loop after every run: the test worker must answer the runner's messages while a minute of synchronous work goes by.
 */
async function bestOf<T>(tries: number, budget: number, fn: () => T): Promise<{ result: T; ms: number; samples: number[] }> {
  const samples: number[] = [];
  let result!: T;
  for (let k = 0; k < tries; k++) {
    const [r, ms] = time(fn);
    result = r;
    samples.push(ms);
    await new Promise<void>((done) => setTimeout(done, 0));
    if (ms < budget) break;
  }
  return { result, ms: Math.min(...samples), samples };
}

const run = (srcs: KeySource[], join: (an: KeyAnalysis) => JoinResult = (an) => joinSources(an, defaultPlan(an))): { an: KeyAnalysis; res: JoinResult } => {
  const an = analyzeSources(srcs[0], srcs.slice(1));
  return { an, res: join(an) };
};
const rssMb = () => process.memoryUsage().rss / 1048576;

describe('performance: 50,000 rows', () => {
  it('50,000-row base + 3 other files (50,000 / 45,000 / 55,000 rows): analysis + join < 6 s, memory stays reasonable', async () => {
    const N = 50000;
    const base: CellSpec[][] = [['学籍番号', '氏名', 'フリガナ', 'メールアドレス', '電話番号', '大学', '学部', ...Array.from({ length: 8 }, (_, k) => `項目${k}`)]];
    for (let i = 0; i < N; i++) {
      base.push([`S${String(i).padStart(7, '0')}`, full(i), `カナ${i % 977}　タロウ${i % 991}`, `user${i}@example.ac.jp`, `090-${1000 + (i % 9000)}-${String(i % 10000).padStart(4, '0')}`, `大学${i % 60}`, `学部${i % 9}`, ...Array.from({ length: 8 }, (_, k) => (i * (k + 3)) % 100)]);
    }
    const o1: CellSpec[][] = [['学籍番号', '氏名', '評価', '備考']];
    for (let k = 0; k < N; k++) {
      const i = (k * 7919) % N;
      o1.push([k % 50 === 0 ? null : `S${String(i).padStart(7, '0')}`, full(i, ' '), 'ABC'[k % 3], k % 10 === 0 ? '要面談' : null]);
    }
    const o2: CellSpec[][] = [['メール', 'お名前', '研修']];
    for (let k = 0; k < N * 0.9; k++) {
      const i = (k * 31) % (N + 500);
      o2.push([`User${i}@Example.ac.jp `, `${full(i, '')}様`, k % 2 ? '済' : '未']);
    }
    const o3: CellSpec[][] = [['氏名', '大学', '電話', '点数']];
    for (let k = 0; k < N * 1.1; k++) {
      const i = (k * 13) % (N + 5000);
      o3.push([full(i), `大学${i % 60}`, `090${1000 + (i % 9000)}${String(i % 10000).padStart(4, '0')}`, k % 100]);
    }
    const srcs = [sourceFromRows('base', base), sourceFromRows('o1', o1), sourceFromRows('o2', o2), sourceFromRows('o3', o3)];
    // what the UI does: analysis, then the join of the default plan (< 6 s together); the 'full' join is a second run
    const a = await bestOf(3, budgetMs(6000), () => {
      const [an, tAnalyze] = time(() => analyzeSources(srcs[0], srcs.slice(1)));
      const [res, tJoin] = time(() => joinSources(an, defaultPlan(an)));
      console.log(`50k + 3: analyze ${tAnalyze.toFixed(0)} ms, join ${tJoin.toFixed(0)} ms, rss ${rssMb().toFixed(0)} MB`);
      return { an, res };
    });
    expect(a.ms).toBeLessThan(budgetMs(6000));
    const { an, res: left } = a.result;
    expect(left.summary.baseRows).toBe(N);
    const f = await bestOf(2, budgetMs(6000), () => joinSources(an, { ...defaultPlan(an), joinType: 'full' }));
    console.log(`50k + 3: join (full) ${f.ms.toFixed(0)} ms`);
    expect(f.ms).toBeLessThan(budgetMs(6000));
    const res = f.result;
    expect(an.pairs.o1.candidates.find((c) => c.selected)?.label).toBe('ID');
    expect(res.reports.o1.matched).toBe(N);
    expect(res.reports.o3.matched).toBeGreaterThan(N * 0.99);
    expect(res.summary.appended).toBeGreaterThan(4000);
    expect(rssMb()).toBeLessThan(2500);
  }, 180000);

  it('50,000 rows with only 10 distinct names (no ID): finishes with a clear message, nothing joined wrongly, no out-of-memory', async () => {
    const N = 50000;
    const names = Array.from({ length: 10 }, (_, i) => full(i * 97));
    const a: CellSpec[][] = [['氏名', '点数']];
    const b: CellSpec[][] = [['氏名', '評価']];
    for (let i = 0; i < N; i++) {
      a.push([names[i % 10], i % 100]);
      b.push([names[(i * 3) % 10], 'ABC'[i % 3]]);
    }
    const srcs = [sourceFromRows('a', a), sourceFromRows('b', b)];
    const r = await bestOf(2, budgetMs(8000), () => run(srcs));
    expect(r.ms).toBeLessThan(budgetMs(8000));
    const { res } = r.result;
    expect(res.reports.b.matched).toBe(0);
    expect(res.reports.b.ambiguous).toHaveLength(N);
    // a value shared by 5,000 rows has 5,000 candidates: only a few are listed, and the list is shared between the rows
    expect(res.reports.b.ambiguous[0].candidates.length).toBeLessThanOrEqual(20);
    expect(res.reports.b.ambiguous[0].candidates).toBe(res.reports.b.ambiguous[1].candidates);
    const msg = res.items.find((i) => i.code === 'tooManyDuplicates');
    expect(msg?.severity).toBe('warning');
    expect(msg?.message).toMatch(/大量にあります/);
    expect(msg?.message).toMatch(/50,000 行/);
    expect(res.items.find((i) => i.code === 'ambiguous')?.message).toMatch(/50000件/);
    expect(rssMb()).toBeLessThan(2500);
  }, 120000);

  it('unique IDs but 10 distinct names and 20% of the IDs blank in the other file: IDs join 80%, the name tier gives up on the rest with the same message', async () => {
    const N = 50000;
    const names = Array.from({ length: 10 }, (_, i) => full(i * 97));
    const a: CellSpec[][] = [['学籍番号', '氏名', '点数']];
    const b: CellSpec[][] = [['学籍番号', '氏名', '評価']];
    for (let i = 0; i < N; i++) {
      a.push([`S${i}`, names[i % 10], i % 100]);
      b.push([i % 5 === 0 ? null : `S${i}`, names[i % 10], 'ABC'[i % 3]]);
    }
    const r = await bestOf(2, budgetMs(8000), () => run([sourceFromRows('a', a), sourceFromRows('b', b)]));
    expect(r.ms).toBeLessThan(budgetMs(8000));
    expect(r.result.res.reports.b.perTier.ID).toBe(N * 0.8);
    expect(r.result.res.reports.b.matched).toBe(N * 0.8);
    expect(r.result.res.items.some((i) => i.code === 'tooManyDuplicates')).toBe(true);
  }, 120000);

  it('10 distinct values in EVERY key-like column (ID / 氏名 / フリガナ / メール / 電話 / 大学), join "full": no hang, nothing joined', async () => {
    const N = 50000;
    const mk = (k: number): CellSpec[][] => {
      const rows: CellSpec[][] = [['学籍番号', '氏名', 'フリガナ', 'メール', '電話', '大学', '評価']];
      for (let i = 0; i < N; i++) {
        const v = (i * k) % 10;
        rows.push([`S00${v}`, `山田　太郎${v}`, `ヤマダ　タロウ${'ABCDEFGHIJ'[v]}`, `user${v}@example.com`, `090-0000-000${v}`, `大学${v}`, i % 100]);
      }
      return rows;
    };
    const srcs = [sourceFromRows('a', mk(1)), sourceFromRows('b', mk(3))];
    const r = await bestOf(2, budgetMs(8000), () => run(srcs, (an) => joinSources(an, { ...defaultPlan(an), joinType: 'full' })));
    expect(r.ms).toBeLessThan(budgetMs(8000));
    expect(r.result.res.reports.b.matched).toBe(0);
    expect(r.result.res.summary.appended).toBeLessThan(N);
  }, 120000);

  it('a one-to-many other file (every person 4 times, no ID), 30,000 people: policy first and skip < 5 s', async () => {
    const N = 30000;
    const a: CellSpec[][] = [['氏名', '大学']];
    const b: CellSpec[][] = [['氏名', '評価']];
    for (let i = 0; i < N; i++) a.push([full(i), `大学${i % 60}`]);
    for (let i = 0; i < N * 4; i++) b.push([full(i % N), 'ABC'[i % 3]]);
    const srcs = [sourceFromRows('a', a), sourceFromRows('b', b)];
    const an = analyzeSources(srcs[0], [srcs[1]]);
    const first = await bestOf(2, budgetMs(5000), () => joinSources(an, { ...defaultPlan(an), duplicatePolicy: 'first' }));
    const skip = await bestOf(2, budgetMs(5000), () => joinSources(an, { ...defaultPlan(an), duplicatePolicy: 'skip' }));
    expect(first.ms).toBeLessThan(budgetMs(5000));
    expect(skip.ms).toBeLessThan(budgetMs(5000));
    expect(first.result.reports.b.matched).toBeGreaterThan(N * 0.98); // (the generated names repeat a little: those rows are ambiguous)
    expect(skip.result.reports.b.matched).toBe(0);
    expect(skip.result.reports.b.ambiguous.length).toBe(N);
  }, 120000);

  it('every name of the other file has a one-character typo (nothing joins, 50,000 suggestions to find) < 6 s', async () => {
    const N = 50000;
    const a: CellSpec[][] = [['氏名', 'フリガナ', '大学']];
    const b: CellSpec[][] = [['氏名', '評価']];
    for (let i = 0; i < N; i++) {
      a.push([full(i), `カナ${i}`, `大学${i % 60}`]);
      b.push([full(i).replace(/.$/, 'x'), 'A']);
    }
    const r = await bestOf(2, budgetMs(6000), () => run([sourceFromRows('a', a), sourceFromRows('b', b)]));
    expect(r.ms).toBeLessThan(budgetMs(6000));
    expect(r.result.res.reports.b.matched).toBe(0);
    expect(r.result.res.reports.b.suggestions.length).toBeGreaterThan(1000);
  }, 120000);
});

describe('performance: wide files and long cells', () => {
  it('10,000 rows × 100 columns that all carry the same unique-looking values: discovery is bounded (< 10 s)', async () => {
    const N = 10000;
    const C = 100;
    const mk = (seed: number): CellSpec[][] => {
      const rows: CellSpec[][] = [Array.from({ length: C }, (_, j) => `列${j}`)];
      for (let i = 0; i < N; i++) rows.push(Array.from({ length: C }, () => `V${(i * (seed + 1)) % N}`));
      return rows;
    };
    const srcs = [sourceFromRows('a', mk(0)), sourceFromRows('b', mk(6))];
    const r = await bestOf(2, budgetMs(10000), () => run(srcs));
    expect(r.ms).toBeLessThan(budgetMs(10000));
    expect(r.result.res.reports.b.matched).toBe(N);
  }, 180000);

  it('100,000-character cells (an essay in a name column, 100,000 open brackets) cost nothing: no key, no quadratic regex', async () => {
    const texts = ['('.repeat(100000), '【'.repeat(100000), '山'.repeat(100000), '1'.repeat(100000), 'a@'.repeat(50000), '0/'.repeat(50000), ' '.repeat(99999) + 'a'];
    const [, ms] = time(() => {
      for (const t of texts) {
        expect(normName(t).folded).toBe('');
        expect(kanaStrict(t)).toBe('');
        expect(normOrg(t)).toBe('');
        expect(normEmail(t)).toBe('');
        expect(normPhone(t).digits).toBe('');
        expect(normGeneric(t)).toBe('');
        expect(normId(t).exact).toBe('');
      }
    });
    expect(ms).toBeLessThan(budgetMs(1000));
  });

  it('a column of long texts among the names does not slow the analysis down', async () => {
    const N = 3000;
    const a: CellSpec[][] = [['氏名', '志望動機']];
    const b: CellSpec[][] = [['氏名', '評価']];
    const essay = '私は貴社（御社）の「理念」に共感し【強く】志望しております。'.repeat(200);
    for (let i = 0; i < N; i++) {
      a.push([full(i), essay + i]);
      b.push([full(i), 'A']);
    }
    const r = await bestOf(2, budgetMs(4000), () => run([sourceFromRows('a', a), sourceFromRows('b', b)]));
    expect(r.ms).toBeLessThan(budgetMs(4000));
    expect(r.result.res.reports.b.matched).toBeGreaterThan(N * 0.9);
  }, 60000);
});
