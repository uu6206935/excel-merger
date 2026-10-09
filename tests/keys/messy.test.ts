/**
 * Round 5 — key detection and joins on MESSY real-world data.
 *
 * tests/fixtures/keys/messy/make-*.ts build 40+ realistic pairs / triples (names, kana, romaji, 異体字, IDs, phones, e-mails,
 * dates in every notation, paged exports, two-row headers, 同姓同名, one-to-many, more rows than the base …) together with the
 * GROUND TRUTH (which person every row is). The committed xlsx files (./out) are read with the app's reader and joined
 * with the default plan; the score is always checked against the truth:
 *   - ZERO wrong automatic matches (a match between two different people / a non-person row);
 *   - the expected number of correct matches;
 *   - what is not joined on purpose is surfaced (suggestion / ambiguous), not dropped silently.
 */
import fs from 'node:fs';
import { beforeAll, describe, expect, it } from 'vitest';
import { allCases } from '../fixtures/keys/messy/all';
import { toXlsxBytes } from '../fixtures/keys/messy/lib';
import type { MessyCase } from '../fixtures/keys/messy/lib';
import type { JoinPlan, JoinResult } from '../../src/core/contracts';
import { analyzeSources, defaultPlan, joinSources } from '../../src/core/keys';
import { buildTable } from '../../src/core/keys/region';
import { type CellSpec, bookFromRows, sourceFromRows, sourceOf } from './helpers';
import { type CaseRun, committedPath, runCase, sourceOfFile } from './messy-helpers';

const byName = (n: string): MessyCase => {
  const c = allCases.find((x) => x.name === n);
  if (!c) throw new Error(`no messy case ${n}`);
  return c;
};

describe('messy fixtures: committed files are in sync with the generators', () => {
  it('there are at least 25 pairs / triples', () => {
    expect(allCases.length).toBeGreaterThanOrEqual(25);
  });
  it('every committed xlsx equals what make-*.ts generates (run `npx tsx tests/fixtures/keys/messy/write-all.ts` after editing them)', () => {
    const stale: string[] = [];
    for (const c of allCases) {
      for (const f of c.files) {
        const p = committedPath(c.name, f);
        if (!fs.existsSync(p)) stale.push(`missing: ${c.name}/${f.fileName}`);
        else if (!Buffer.from(toXlsxBytes(f)).equals(fs.readFileSync(p))) stale.push(`differs: ${c.name}/${f.fileName}`);
      }
    }
    expect(stale).toEqual([]);
  });
});

describe('messy fixtures: the default plan never joins the wrong person', () => {
  const runs = new Map<string, CaseRun>();
  beforeAll(async () => {
    for (const c of allCases) runs.set(c.name, await runCase(c, { fromDisk: true }));
  }, 120000);

  for (const c of allCases) {
    it(`${c.name}: ${c.note.slice(0, 90)}`, () => {
      const run = runs.get(c.name)!;
      for (const [id, e] of Object.entries(run.evals)) {
        // the quality bar: no silent wrong automatic matches
        expect(e.wrong, `${c.name}/${id} wrong matches`).toEqual([]);
        expect(e.junkAppended, `${c.name}/${id} junk rows appended`).toEqual([]);
        expect(e.duplicatedPeople, `${c.name}/${id} people appended twice`).toEqual([]);
        const want = c.expect.matched?.[id];
        if (want !== undefined) expect(e.correct, `${c.name}/${id} correct matches`).toBe(want);
        const surfaced = e.missed.filter((m) => m.surfaced !== null).length;
        const unsurfaced = e.missed.length - surfaced;
        expect(surfaced, `${c.name}/${id} surfaced (suggestion / ambiguous)`).toBe(c.expect.surfaced?.[id] ?? 0);
        expect(unsurfaced, `${c.name}/${id} missed and not surfaced`).toBe(c.expect.unsurfaced?.[id] ?? 0);
        const plan = c.expect.plan?.[id];
        if (plan) expect(e.plan, `${c.name}/${id} plan`).toEqual(plan);
      }
      for (const code of c.expect.items ?? []) expect(run.codes.has(code), `${c.name}: check item ${code}`).toBe(true);
      for (const code of c.expect.noItems ?? []) expect(run.codes.has(code), `${c.name}: no check item ${code}`).toBe(false);
      if (c.expect.appended !== undefined) expect(run.res.summary.appended, `${c.name} appended rows`).toBe(c.expect.appended);
    });
  }

  it('the joins are deterministic (same input, same rows)', async () => {
    for (const n of ['typos-never-silent', 'namesakes-university', 'two-files-without-id-full']) {
      const a = runs.get(n)!;
      const b = await runCase(byName(n), { fromDisk: true });
      expect(b.res.rows).toEqual(a.res.rows);
      expect(b.res.items).toEqual(a.res.items);
    }
  });
});

// ---------------------------------------------------------------------------------------------------------------------
// details that the score alone does not show
// ---------------------------------------------------------------------------------------------------------------------

describe('messy fixtures: details', () => {
  const get = async (n: string, opts?: Parameters<typeof runCase>[1]) => runCase(byName(n), { fromDisk: true, ...opts });
  const rowOf = (run: CaseRun, id: string, baseRow: number) => run.res.rows.find((r) => r.baseRow === baseRow)?.matches[id] ?? null;

  it('English names: token order, commas and capitals do not matter; the family is the Latin-name one', async () => {
    const run = await get('english-name-orders');
    for (const id of ['s1', 's2', 's3']) {
      const cand = run.an.pairs[id].candidates.find((c) => c.selected)!;
      expect(cand.family).toBe('name');
      expect(cand.label).toBe('氏名ローマ字');
    }
    // 姓名の順序違い is reported as a variant, not as an error
    expect(Object.keys(run.res.reports.s1.flags)).toContain('姓名の順序違い');
    expect(run.res.items.some((i) => i.code === 'conflict')).toBe(false);
  });

  it('split English columns (Last Name / First Name) join a one-column roster', async () => {
    const run = await get('english-split-columns');
    expect(run.evals.s1.correct).toBe(14);
    expect(run.evals.s2.correct).toBe(8);
    expect(run.an.pairs.s1.candidates.find((c) => c.selected)?.otherColumns.map((x) => x.header)).toEqual(['Last Name', 'First Name']);
  });

  it('web-form headers お名前（姓） / お名前（名） / セイ（カナ） / メイ（カナ） are split-name columns', async () => {
    const run = await get('split-name-header-variants');
    const types = Object.fromEntries(run.an.pairs.s1.other.columns.map((c) => [c.header, c.type]));
    // (header texts are NFKC-normalised: the full-width parentheses become half-width ones)
    expect(types).toMatchObject({ 'お名前(姓)': 'lastName', 'お名前(名)': 'firstName', 'セイ(カナ)': 'lastNameKana', 'メイ(カナ)': 'firstNameKana' });
  });

  it('異体字 / 旧姓 / 読み仮名 are joined and flagged as variants; 斎 → 斉 is only suggested (with a reason)', async () => {
    const run = await get('name-variants-maiden');
    const flags = run.res.reports.s1.flags;
    expect(flags['異体字の違い']?.count).toBeGreaterThanOrEqual(3);
    const miss = run.evals.s1.missed;
    expect(miss).toHaveLength(1);
    expect(miss[0].person).toBe('p15');
    const sug = run.res.reports.s1.suggestions.find((s) => s.base.preview[0].includes('斎藤'));
    expect(sug?.reasons.join()).toMatch(/似た字/);
    // with 異体字 switched off the variants are no automatic match any more
    const off = await get('name-variants-maiden', { plan: undefined });
    expect(off.evals.s1.wrong).toEqual([]);
  });

  it('options off: without 異体字 / 敬称 / 旧姓 the variants stay unmatched, never wrongly matched', async () => {
    const c = byName('name-variants-maiden');
    const { sources } = await runCase(c, { fromDisk: true });
    const an = analyzeSources(sources[0], sources.slice(1), { itaiji: false, maidenName: false });
    const res = joinSources(an, defaultPlan(an));
    const matched = res.rows.filter((r) => r.baseRow !== null && r.matches.s1).length;
    expect(matched).toBeLessThan(17);
    const h = byName('honorifics');
    const hs = (await runCase(h, { fromDisk: true })).sources;
    const an2 = analyzeSources(hs[0], hs.slice(1), { honorifics: false });
    const res2 = joinSources(an2, defaultPlan(an2));
    const m2 = res2.rows.filter((r) => r.matches.s1).length;
    expect(m2).toBeLessThan(10);
  });

  it('ID with a letter prefix: S2023-0012 ⟷ 20230012 is a named variant candidate (info flag), B/M/D prefixes that carry identity are not stripped', async () => {
    const run = await get('id-prefix');
    const cand = run.an.pairs.s1.candidates.find((c) => c.selected)!;
    expect(cand.label).toBe('ID（英字の接頭辞を除く）');
    expect(cand.id).toMatch(/~prefix$/);
    expect(Object.keys(run.res.reports.s1.flags)).toContain('ID表記の違いを補正');
    const id2 = await get('id-prefix-reverse-and-identity');
    expect(id2.an.pairs.s2.candidates.some((c) => c.selected)).toBe(false);
    expect(id2.codes.has('s2:noKey')).toBe(true);
    expect(id2.evals.s2.wrong).toEqual([]);
  });

  it('IDs of different lengths: the tail variant is joined but every row is flagged for checking (要確認)', async () => {
    const run = await get('id-different-lengths');
    const rep = run.res.reports.s2;
    expect(rep.matched).toBe(14);
    expect(Object.keys(rep.flags)).toEqual(['IDの桁数が異なる']);
    expect(rep.flags['IDの桁数が異なる'].level).toBe('warn');
    expect(run.res.items.find((i) => i.code === 'checkWarn' && i.sourceId === 's2')?.severity).toBe('warning');
  });

  it('IDs that went through text as 2.02301E+11 have NO key: they never read back as somebody else’s number', async () => {
    const run = await get('id-scientific-notation');
    const t = run.an.pairs.s1;
    expect(t.other.columns[0].notes.join()).toMatch(/指数表記/);
    expect(run.res.items.find((i) => i.code === 'lossyId')?.severity).toBe('warning');
    // the lookalike 202301234567 (→ 2.02301E+11) must not join the base person whose number is 202301000000
    expect(run.evals.s1.wrong).toEqual([]);
    const special = run.res.rows.find((r) => r.baseRow === 22); // p19b
    expect(special?.matches.s1 ?? null).toBeNull();
  });

  it('leading zeros, 全角 digits, trailing spaces, apostrophes: one ID', async () => {
    const run = await get('id-leading-zeros');
    expect(run.res.reports.s1.matched).toBe(20);
    expect(run.res.reports.s1.perTier.ID).toBe(20);
  });

  it('two ID systems that overlap by accident are not an automatic key (names disagree): the ID pair is demoted and says so', async () => {
    const run = await get('id-systems-overlap-by-accident');
    const cand = run.an.pairs.s1.candidates.find((c) => c.label === 'ID');
    expect(cand).toBeDefined();
    expect(cand!.selected).toBe(false);
    expect(cand!.verdict).toBe('不一致多数');
    expect(run.evals.s1.plan).toEqual(['氏名']);
  });

  it('phone numbers in every notation; two numbers in one cell; e-mail in every notation', async () => {
    const ph = await get('phone-variants');
    expect(ph.res.reports.s1.perTier['電話番号']).toBe(18);
    const two = await get('phone-two-numbers');
    expect(two.res.reports.s1.matched).toBe(16);
    const em = await get('email-variants');
    expect(em.res.reports.s1.perTier['メール']).toBe(16);
  });

  it('a shared family e-mail address decides nothing: both base rows are ambiguous, the others join', async () => {
    const run = await get('email-shared-family-address');
    expect(run.res.reports.s1.ambiguous.map((a) => a.base.excelRow).sort((a, b) => a - b)).toEqual([3, 13]);
    expect(run.res.items.find((i) => i.code === 'ambiguous')?.message).toMatch(/2件/);
  });

  it('dates of birth: 和暦 / serial / 20010401 / 生まれ; the two 田中 太郎 are told apart by the birth date', async () => {
    const run = await get('birthdate-notations');
    for (const id of ['s1', 's2', 's3']) {
      expect(run.an.pairs[id].candidates.find((c) => c.selected)?.label).toBe('氏名＋生年月日');
      const tanaka = run.res.rows.filter((r) => r.baseRow !== null && r.baseRow >= 22);
      // s3 lists only the second 田中 太郎 (base row 23)
      expect(tanaka.map((r) => !!r.matches[id])).toEqual(id === 's3' ? [false, true] : [true, true]);
    }
  });

  it('paged export: the table runs across the pages, repeated headers / page lines / subtotals are skipped (as base and as other)', async () => {
    const run = await get('paged-export-as-other');
    const src = run.sources[1];
    expect(src.region).toMatchObject({ headerRow: 3, firstDataRow: 4, lastDataRow: 44 });
    const t = buildTable(src);
    expect(t.rows).toHaveLength(21);
    // per page break: the subtotal (in four notations), 2 blank rows, the page line, the repeated header
    expect(t.skippedRows).toEqual([9, 10, 11, 12, 13, 19, 20, 21, 22, 23, 29, 30, 31, 32, 33, 39, 40, 41, 42, 43]);
    expect(run.res.summary.appended).toBe(3);
    const asBase = await get('paged-export-as-base');
    expect(asBase.res.summary.baseRows).toBe(18);
    expect(asBase.evals.s1.correct).toBe(18);
  });

  it('two-row merged header: the sub-header row is the header, the group title is joined to it, imports carry that header', async () => {
    const run = await get('two-row-merged-header');
    const a = run.sources[1];
    expect(a.region).toMatchObject({ headerRow: 2, firstDataRow: 3, lastDataRow: 16 });
    const t = buildTable(a);
    expect(t.headers).toEqual(['学籍番号', '氏名', '面接評価 評価1', '面接評価 評価2', '面接評価 評価3', '備考']);
    expect(t.rows).toHaveLength(14);
    const imps = run.res.defaultImports.filter((i) => i.sourceId === 's1');
    // 評価1 sits in the header cell itself; the joined text '面接評価 評価1' is what the merged sheet gets
    // (備考 is merged over both header rows: its own header cell in row 2 is blank, so the text is passed on as well)
    expect(imps.map((i) => [i.col, i.header])).toEqual([[3, '面接評価 評価1'], [4, '面接評価 評価2'], [5, '面接評価 評価3'], [6, '備考']]);
    const b = run.sources[2];
    expect(b.region).toMatchObject({ headerRow: 2, firstDataRow: 3 });
    expect(buildTable(b).headers[0]).toBe('学籍番号');
  });

  it('a hidden first sheet is never the data sheet; a hidden column still carries the key', async () => {
    const run = await get('hidden-column-and-sheet');
    expect(run.sources[0].region.sheet).toBe('名簿');
    expect(run.evals.s1.plan).toEqual(['ID']);
  });

  it('remark lines between rows, vertical merged cells, a list in two pieces', async () => {
    const remark = await get('remark-lines-inside-list');
    expect(remark.res.summary.appended).toBe(0);
    const t = buildTable(remark.sources[1]);
    expect(t.rows).toHaveLength(18);
    expect(t.skippedRows).toEqual([8, 15]);
    const merged = await get('vertical-merged-university');
    const tb = buildTable(merged.sources[0]);
    expect(tb.rows.map((r) => r[1])).toEqual(['東京大学', '東京大学', '東京大学', '東京大学', '早稲田大学', '早稲田大学', '早稲田大学', '早稲田大学', '京都大学', '京都大学', '京都大学', '京都大学', '東京大学', '早稲田大学']);
    const two = await get('list-in-two-pieces');
    const it = two.res.items.find((i) => i.code === 'dataBelow');
    expect(it?.sourceId).toBe('s1');
    expect(it?.message).toMatch(/8 行/);
  });

  it('bare No columns and category-only files: no key, nothing joined', async () => {
    const no = await get('bare-no-columns');
    expect(no.res.reports.s2.matched).toBe(0);
    expect(no.codes.has('s2:noKey')).toBe(true);
    const cat = await get('no-person-key-at-all');
    for (const id of ['s1', 's2']) expect(cat.res.reports[id].matched).toBe(0);
  });

  it('a referrer column is not the person of the row (紹介者 ⟷ 氏名 is never a tier)', async () => {
    const run = await get('referral-columns');
    expect(run.an.pairs.s1.candidates.filter((c) => c.selected)).toEqual([]);
    expect(run.codes.has('s1:noKey')).toBe(true);
  });

  it('typos: every one-character typo is suggested with the right person, never joined; kana-correct rows are vetoed by the name tier', async () => {
    const run = await get('typos-never-silent');
    const sug = run.res.reports.s1.suggestions;
    expect(sug.length).toBeGreaterThanOrEqual(8);
    for (const s of run.evals.s1.missed) expect(s.surfaced).toBe('suggestion');
    expect(sug.every((s) => s.score >= 0.6)).toBe(true);
    // s2: the name tier vetoes a kana-tier match whose name differs
    expect(run.res.reports.s2.conflicts.length + run.res.reports.s2.suggestions.length).toBeGreaterThan(0);
    expect(run.evals.s2.wrong).toEqual([]);
  });

  it('a maiden name in one file and the married name (旧姓：…) in the other: suggested with the reason 旧姓で一致, never joined', async () => {
    const run = await get('maiden-name-in-one-file');
    const sug = run.res.reports.s1.suggestions;
    expect(sug).toHaveLength(2);
    for (const s of sug) expect(s.reasons).toContain('旧姓で一致');
    // (both base rows stay unmatched: the suggestion is the user's to accept)
    const unmatched = run.res.reports.s1.unmatchedBase.map((r) => r.preview[0]);
    expect(unmatched).toEqual(expect.arrayContaining(['小林　花子（旧姓：加藤）', '山口　優子（旧姓：井上）']));
  });

  it('同姓同名: resolved by 氏名＋大学 / ＋生年月日, else ambiguous (both candidates listed) — never guessed', async () => {
    const run = await get('namesakes-university');
    const amb = run.res.reports.s2.ambiguous;
    expect(amb).toHaveLength(2);
    for (const a of amb) expect(a.candidates).toHaveLength(2);
    const fac = await get('namesakes-same-university-faculty');
    expect(fac.evals.s1.plan[0]).toBe('氏名＋生年月日');
    expect(fac.res.reports.s2.ambiguous).toHaveLength(2);
  });

  it('one-to-many: first / last / skip pick what they say', async () => {
    const f = await get('one-to-many-first');
    const l = await get('one-to-many-last');
    const s = await get('one-to-many-skip');
    // interview rows of p0..p14 (see make-relations): the joined row is the first / last of the person
    const hist = byName('one-to-many-first').files[1].truth;
    const rowsOf = (p: string) => Object.entries(hist).filter(([, id]) => id === p).map(([r]) => Number(r));
    for (const [run, pick] of [[f, (x: number[]) => x[0]], [l, (x: number[]) => x[x.length - 1]]] as const) {
      for (const [baseRowStr, p] of Object.entries(byName('one-to-many-first').files[0].truth)) {
        const rows = rowsOf(p);
        if (!rows.length) continue;
        expect(rowOf(run, 's1', Number(baseRowStr))?.row, `${p}`).toBe(pick(rows));
      }
    }
    expect(f.res.items.find((i) => i.code === 'duplicate')?.severity).toBe('warning');
    // skip: people with several rows are ambiguous, the rest joins
    expect(s.res.reports.s1.ambiguous).toHaveLength(10);
  });

  it('the base with several rows per person is a VLOOKUP-style join (lookup info, no duplicateBase warning)', async () => {
    const run = await get('one-to-many-lookup-by-id');
    expect(run.an.pairs.s1.mode).toBe('lookup');
    expect(run.res.items.find((i) => i.code === 'lookup')?.severity).toBe('info');
    const few = await get('one-to-many-name-only-few');
    expect(few.an.pairs.s1.mode).toBe('oneToOne');
    expect(few.res.reports.s1.ambiguous).toHaveLength(2);
  });

  it('more rows than the base with "full": 姓/名 of the base are filled from 氏名 (split at the space), repeated new people once', async () => {
    const run = await get('more-rows-than-base-full');
    expect(run.res.summary).toEqual({ baseRows: 10, matchedAny: 10, matchedAll: 10, appended: 12 });
    expect(run.res.keyFill.s1).toEqual([{ baseCol: 1, sourceCol: 1, splitPart: 0 }, { baseCol: 2, sourceCol: 1, splitPart: 1 }]);
    const dup = run.res.items.find((i) => i.code === 'duplicate');
    expect(dup?.message).toMatch(/3件/);
    // p25 has no space in its 氏名: it cannot be split into 姓 / 名 — said so, the whole name goes into the 姓 column
    const ns = run.res.items.find((i) => i.code === 'nameNotSplit');
    expect(ns?.severity).toBe('warning');
    expect(ns?.message).toMatch(/1件/);
    expect(ns?.otherRows).toHaveLength(1);
  });

  it('two files without ID share 氏名＋大学: the second finds the rows the first appended', async () => {
    const run = await get('two-files-without-id-full');
    expect(run.res.summary.appended).toBe(5);
    const both = run.res.rows.filter((r) => r.baseRow === null && r.matches.s1 && r.matches.s2);
    expect(both).toHaveLength(2);
    expect(run.res.reports.s2.plan[0].name).toBe('氏名＋大学');
  });

  it('every file keyed differently: each pair has its own tier list', async () => {
    const run = await get('every-file-keyed-differently');
    expect(run.evals.s1.plan).toEqual(['ID（英字の接頭辞を除く）']);
    expect(run.evals.s2.plan).toEqual(['氏名']);
    expect(run.evals.s3.plan).toEqual(['メール']);
    expect(run.evals.s4.plan).toEqual(['電話番号']);
    expect(run.evals.s5.plan).toEqual(['氏名カナ']);
  });

  it('all cases: every check item points at rows of the files (no row beyond the sheet)', async () => {
    for (const c of allCases) {
      const run = await get(c.name);
      const maxBase = run.sources[0].book.sheets.find((s) => s.name === run.sources[0].region.sheet)!.maxRow;
      for (const it of run.res.items) {
        for (const r of it.rows ?? []) expect(r).toBeLessThanOrEqual(maxBase);
        if (it.sourceId) {
          const src = run.sources.find((s) => s.id === it.sourceId)!;
          const max = src.book.sheets.find((s) => s.name === src.region.sheet)!.maxRow;
          for (const r of it.otherRows ?? []) expect(r).toBeLessThanOrEqual(max);
        }
      }
    }
  });

  it('sourceOfFile reads the committed bytes exactly like the generated ones', async () => {
    const c = byName('paged-export-as-other');
    const a = await sourceOfFile(c.files[1]);
    const b = await sourceOfFile(c.files[1], new Uint8Array(fs.readFileSync(committedPath(c.name, c.files[1]))));
    expect(b.region).toEqual(a.region);
  });
});

describe('messy fixtures: small hand-made lists', () => {
  const join = (base: CellSpec[][], other: CellSpec[][], plan: Partial<JoinPlan> = {}) => {
    const b = sourceFromRows('base', base);
    const o = sourceFromRows('o', other);
    const an = analyzeSources(b, [o]);
    return { an, res: joinSources(an, { ...defaultPlan(an), ...plan }) };
  };
  const matchOf = (res: JoinResult, baseRow: number) => res.rows.find((r) => r.baseRow === baseRow)?.matches.o ?? null;

  it('homophones in a tiny list (lookup mode by the 30% rule): 伊藤 花子 and 井藤 花子 still do not share one answer', () => {
    const { an, res } = join(
      [['氏名', 'フリガナ'], ['伊藤 花子', 'イトウ ハナコ'], ['井藤 花子', 'イトウ ハナコ'], ['山田 太郎', 'ヤマダ タロウ']],
      [['フリガナ', '評価'], ['イトウ ハナコ', 'A'], ['ヤマダ タロウ', 'B']],
    );
    expect(an.pairs.o.mode).toBe('lookup');
    expect(matchOf(res, 2)).toBeNull();
    expect(matchOf(res, 3)).toBeNull();
    expect(matchOf(res, 4)?.row).toBe(3);
    expect(res.reports.o.ambiguous).toHaveLength(2);
  });

  it('the same name sharing base rows in lookup mode still joins when the rows are one person (same e-mail / no disagreement)', () => {
    const { res } = join(
      [['氏名', 'メール', '企業'], ['山田 太郎', 'y@example.com', 'A社'], ['山田 太郎', 'y@example.com', 'B社'], ['山田 太郎', 'y@example.com', 'C社'], ['鈴木 花子', 's@example.com', 'A社']],
      [['氏名', '大学'], ['山田 太郎', '東京大学'], ['鈴木 花子', '京都大学']],
    );
    expect(matchOf(res, 2)?.row).toBe(2);
    expect(matchOf(res, 3)?.row).toBe(2);
    expect(matchOf(res, 4)?.row).toBe(2);
    expect(matchOf(res, 5)?.row).toBe(3);
  });

  it('two rows with the same name but a different e-mail are two people (not a resubmission): ambiguous under first / last', () => {
    for (const duplicatePolicy of ['first', 'last'] as const) {
      const { res } = join(
        [['氏名', '大学'], ['田中 太郎', '東京大学'], ['山田 花子', '京都大学']],
        [['氏名', 'メール'], ['田中 太郎', 'a@example.com'], ['田中 太郎', 'b@example.com'], ['山田 花子', 'h@example.com']],
        { duplicatePolicy },
      );
      expect(matchOf(res, 2)).toBeNull();
      expect(matchOf(res, 3)?.row).toBe(4);
      expect(res.reports.o.ambiguous).toHaveLength(1);
    }
    // the very same person twice (same e-mail): the policy picks the first / last row
    const same = join(
      [['氏名', '大学'], ['田中 太郎', '東京大学'], ['山田 花子', '京都大学']],
      [['氏名', 'メール'], ['田中 太郎', 'a@example.com'], ['田中 太郎', 'a@example.com'], ['山田 花子', 'h@example.com']],
      { duplicatePolicy: 'last' },
    );
    expect(matchOf(same.res, 2)?.row).toBe(3);
  });

  it('with "full", two new people of one name but different e-mails are appended as two rows, a resubmission as one', () => {
    const { res } = join(
      [['氏名', '大学'], ['山田 花子', '京都大学']],
      [['氏名', 'メール'], ['山田 花子', 'h@example.com'], ['田中 太郎', 'a@example.com'], ['田中 太郎', 'b@example.com'], ['佐藤 一郎', 's@example.com'], ['佐藤 一郎', 's@example.com']],
      { joinType: 'full' },
    );
    expect(res.summary.appended).toBe(3);
  });

  it('a name pasted as hiragana in one file and katakana in the other is suggested (似た字), not joined', () => {
    const run = join(
      [['氏名', '大学'], ['渡辺 さくら', '京都大学'], ['山田 太郎', '東京大学']],
      [['氏名', '評価'], ['渡辺 サクラ', 'A'], ['山田 太郎', 'B']],
    );
    expect(matchOf(run.res, 2)).toBeNull();
    expect(matchOf(run.res, 3)?.row).toBe(3);
    const sug = run.res.reports.o.suggestions;
    expect(sug).toHaveLength(1);
    expect(sug[0].reasons.join()).toMatch(/似た字/);
  });

  it('the same ID system with abbreviated names (surname only) or names in kana is still an ID key — only a clash of comparable names demotes an ID pair', () => {
    const base: CellSpec[][] = [['学籍番号', '氏名'], [1001, '山田 太郎'], [1002, '鈴木 花子'], [1003, '佐藤 一郎'], [1004, '高橋 健太']];
    const surname = join(base, [['学籍番号', '氏名', '評価'], [1001, '山田', 'A'], [1002, '鈴木', 'B'], [1003, '佐藤', 'C'], [1004, '高橋', 'D']]);
    expect(surname.an.pairs.o.candidates.find((c) => c.label === 'ID')?.selected).toBe(true);
    const kana = join(base, [['学籍番号', '氏名', '評価'], [1001, 'ヤマダ タロウ', 'A'], [1002, 'スズキ ハナコ', 'B'], [1003, 'サトウ イチロウ', 'C'], [1004, 'タカハシ ケンタ', 'D']]);
    expect(kana.an.pairs.o.candidates.find((c) => c.label === 'ID')?.selected).toBe(true);
    expect([2, 3, 4, 5].map((r) => matchOf(kana.res, r)?.row)).toEqual([2, 3, 4, 5]);
    // a real clash (same numbers, other kanji names) does demote it
    const clash = join(base, [['学籍番号', '氏名', '評価'], [1001, '中村 結衣', 'A'], [1002, '小林 花子', 'B'], [1003, '加藤 翔', 'C'], [1004, '吉田 蓮', 'D']]);
    expect(clash.an.pairs.o.candidates.find((c) => c.label === 'ID')?.selected).toBe(false);
  });

  it('international students: kanji and Latin names (accents, hyphens, capitals, any order) in one column', () => {
    const { an, res } = join(
      [['氏名', '大学'], ['山田 太郎', '東大'], ['Michael Smith', '慶應'], ['李 明', '早稲田'], ['Nguyen Van An', '京大'], ['Park Ji-min', '阪大'], ['Anna Müller', '上智'], ['José García', '九大']],
      [['氏名', '評価'], ['Smith Michael', 'A'], ['山田　太郎', 'B'], ['Nguyen  Van  An', 'C'], ['PARK JI-MIN', 'D'], ['李　明', 'E'], ['Müller, Anna', 'F'], ['GARCÍA, José', 'G']],
    );
    expect(an.pairs.o.candidates.find((c) => c.selected)?.family).toBe('name');
    expect([2, 3, 4, 5, 6, 7, 8].map((r) => matchOf(res, r)?.row)).toEqual([3, 2, 6, 4, 5, 7, 8]);
  });

  it('IDs with thousands separators and with a No. prefix', () => {
    const { res, an } = join(
      [['学籍番号', '氏名'], [20230012, '山田 太郎'], [20230013, '鈴木 花子'], [20230014, '佐藤 一郎']],
      [['学籍番号', '評価'], ['20,230,012', 'A'], ['20,230,013', 'B'], ['20,230,014', 'C']],
    );
    expect(an.pairs.o.candidates.find((c) => c.selected)?.label).toBe('ID');
    expect([2, 3, 4].map((r) => matchOf(res, r)?.row)).toEqual([2, 3, 4]);
    const p = join(
      [['学籍番号', '氏名'], ['NO.123', '山田 太郎'], ['NO.124', '鈴木 花子'], ['NO.125', '佐藤 一郎']],
      [['学籍番号', '評価'], [123, 'A'], [124, 'B'], [125, 'C']],
    );
    expect(p.an.pairs.o.candidates.find((c) => c.selected)?.label).toBe('ID（英字の接頭辞を除く）');
    expect([2, 3, 4].map((r) => matchOf(p.res, r)?.row)).toEqual([2, 3, 4]);
  });

  it('1904-date-system workbooks: dates of birth are read with the right offset (xlsx, raw serials)', () => {
    const a = bookFromRows('a.xlsx', [{ name: 'S', rows: [['氏名', '生年月日'], ['山田 太郎', new Date(Date.UTC(2001, 3, 1))], ['山田 太郎', new Date(Date.UTC(2003, 5, 5))]] }], 'a');
    const b = bookFromRows('b.xlsx', [{ name: 'S', rows: [['氏名', '生年月日', '評価'], ['山田 太郎', new Date(Date.UTC(2001, 3, 1)), 'A'], ['山田 太郎', new Date(Date.UTC(2003, 5, 5)), 'B']] }], 'b');
    // the same serials read as a 1904-system book are 1462 days later in real time: the two dates are then 4 years apart
    b.date1904 = true;
    const an = analyzeSources(sourceOf(a, 'a'), [sourceOf(b, 'b')]);
    const res = joinSources(an, defaultPlan(an));
    expect(res.reports.b.matched).toBe(0);
    const c = bookFromRows('c.xlsx', [{ name: 'S', rows: [['氏名', '生年月日', '評価'], ['山田 太郎', new Date(Date.UTC(2001, 3, 1) - 1462 * 86400000), 'A'], ['山田 太郎', new Date(Date.UTC(2003, 5, 5) - 1462 * 86400000), 'B']] }], 'c');
    c.date1904 = true;
    const an2 = analyzeSources(sourceOf(a, 'a'), [sourceOf(c, 'c')]);
    const res2 = joinSources(an2, defaultPlan(an2));
    expect(res2.reports.c.matched).toBe(2);
  });
});
