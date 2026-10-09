/**
 * Round 4: files without a usable key are EXCLUDED from the merge (they do not block it), and the manual key picker
 * (addManualKey / removeManualKey) against the ?mock=1 services, which implement keys.addManualCandidate like the real
 * module (the real one is covered by tests/keys/manual.test.ts).
 */
import { describe, expect, it } from 'vitest';
import type { JoinPlan } from '../../src/core/contracts';
import { createController } from '../../src/state/controller';
import { candidateByColumns, excludedSourceIds, joinedSourceCount, previewRows, summaryText, toggleCandidate, withoutSources, joinSummary } from '../../src/state/keysPlan';
import { createMockServices } from '../../src/state/mock';
import type { Services } from '../../src/state/services';
import { file, joined, store } from './fixtures';

const NOKEY = '参加状況（キーなし）.xlsx';

function setup(patch: Partial<Services> = {}) {
  const st = store();
  const services = { ...createMockServices('?mock=1&delay=0.05'), ...patch };
  const ctrl = createController({ services, dispatch: st.dispatch, getState: st.getState });
  return { st, ctrl, services };
}

/** 学生回答 (base, s1) + 職場回答 (s2, has keys) + the key-less file (s3) */
async function loaded(patch: Partial<Services> = {}) {
  const env = setup(patch);
  env.ctrl.init();
  await env.ctrl.addFiles([file('学生回答.xlsx'), file('職場回答.xlsx'), file(NOKEY)]);
  await env.ctrl.idle();
  await env.ctrl.ensureAnalysis();
  return env;
}

const present = (env: { st: ReturnType<typeof store> }) => env.st.state.merged!.history.present;

describe('files without a key are excluded', () => {
  it('the analysis finds no key for the file; it is excluded and not part of the join result', async () => {
    const { st } = await loaded();
    const s = st.state;
    expect(s.analysis?.pairs.s3.candidates).toEqual([]);
    expect(s.plan?.tiers).toEqual({ s2: ['id', 'name'], s3: [] });
    expect(excludedSourceIds(s.analysis, s.plan)).toEqual(['s3']);
    expect(Object.keys(s.joinResult!.reports)).toEqual(['s2']);
    expect(s.joinResult!.defaultImports.every((d) => d.sourceId === 's2')).toBe(true);
    expect(Object.keys(s.joinResult!.keyFill)).toEqual(['s2']);
    // the summary counts rows of the files that join (with one joined file matchedAll = matchedAny)
    const sum = joinSummary(s.joinResult!);
    expect(sum.baseRows).toBe(30);
    expect(sum.matched).toBe(s.joinResult!.summary.matchedAny);
    expect(sum.matched).toBeGreaterThan(20);
    expect(summaryText(sum, '名', 1)).toMatch(/ \/ 結合対象外 1 ファイル$/);
    expect(summaryText(sum, '名')).not.toContain('結合対象外');
  });

  it('the preview lists the file as 結合対象外 and imports nothing from it', async () => {
    const { st } = await loaded();
    const rows = previewRows(st.state.analysis!, st.state.plan!, st.state.joinResult);
    expect(rows.map((r) => [r.sourceId, r.excluded, r.keys, r.matched])).toEqual([
      ['s1', false, '基準', null],
      ['s2', false, 'ID → 氏名', 28],
      ['s3', true, '結合対象外', null],
    ]);
  });

  it('join 「末尾に追加」 does not append the rows of an excluded file', async () => {
    const { st, ctrl, services } = await loaded();
    ctrl.setJoinType('full');
    const j = st.state.joinResult!;
    const appended = j.rows.filter((r) => r.baseRow === null);
    expect(appended.length).toBe(j.summary.appended);
    expect(appended.every((r) => r.appendedFrom?.sourceId === 's2')).toBe(true);
    expect(appended.length).toBeLessThan(5); // the raw join appends all 20 rows of the key-less file
    // (the keys module itself does append them: the UI strips them)
    const raw = services.joinSources(st.state.analysis!, st.state.plan!);
    expect(raw.rows.filter((r) => r.baseRow === null && r.appendedFrom?.sourceId === 's3').length).toBe(20);
    expect(j.items.some((it) => it.sourceId === 's3')).toBe(false);
  });

  it('the merge goes ahead without it: no columns of the file, a note in the toast', async () => {
    const env = await loaded();
    await env.ctrl.merge();
    expect(env.st.state.phase).toBe('merged');
    const d = present(env).doc;
    expect(d.columns).toHaveLength(15); // 11 base columns + 配属先 評価 面接日 コメント
    expect(d.columns.some((c) => c.origin.kind === 'import' && c.origin.sourceId === 's3')).toBe(false);
    expect(env.st.state.toast?.text).toContain('結合対象外 1 ファイル（参加状況（キーなし））');
    expect(env.st.state.merged!.layout.imported.every((i) => i.sourceId === 's2')).toBe(true);
  });

  it('the merge input handed to buildPremerged only has the joined files', async () => {
    let seen: { sources: string[]; imports: string[]; rows: number } | null = null;
    const base = createMockServices('?mock=1&delay=0.05');
    const env = await loaded({
      buildPremerged: async (input, onProgress) => {
        seen = { sources: input.sources.map((s) => s.id), imports: [...new Set(input.imports.map((i) => i.sourceId))], rows: input.rows.length };
        return base.buildPremerged(input, onProgress);
      },
    });
    await env.ctrl.merge();
    expect(seen).toEqual({ sources: ['s2'], imports: ['s2'], rows: 30 });
  });

  it('selecting a key later (plan edit) brings the file back into the merge without any other action', async () => {
    const env = await loaded();
    const { st, ctrl } = env;
    expect(excludedSourceIds(st.state.analysis, st.state.plan)).toEqual(['s3']);
    const res = ctrl.addManualKey('s3', [3], [2]); // 氏名 ⟷ 参加者
    expect(res).toEqual({ ok: true, id: 'manual:3|2', existing: false });
    expect(excludedSourceIds(st.state.analysis, st.state.plan)).toEqual([]);
    expect(Object.keys(st.state.joinResult!.reports).sort()).toEqual(['s2', 's3']);
    expect(st.state.joinResult!.defaultImports.some((d) => d.sourceId === 's3')).toBe(true);
    await ctrl.merge();
    expect(st.state.phase).toBe('merged');
    const d = present(env).doc;
    // 番号 班 点数 of the key-less file (参加者 is the key column, not imported by default)
    expect(d.columns.filter((c) => c.origin.kind === 'import' && c.origin.sourceId === 's3')).toHaveLength(3);
    expect(st.state.toast?.text).not.toContain('結合対象外');
  });

  it('with NO key at all the merge is refused (and the summary of the join is empty)', async () => {
    const env = await loaded();
    const { st, ctrl } = env;
    ctrl.updatePlan({ ...st.state.plan!, tiers: { s2: [], s3: [] } });
    expect(excludedSourceIds(st.state.analysis, st.state.plan)).toEqual(['s2', 's3']);
    expect(joinedSourceCount(st.state.joinResult)).toBe(0);
    await ctrl.merge();
    expect(st.state.phase).toBe('files');
    expect(st.state.toast?.text).toContain('どのファイルにも結合キーが選ばれていません');
    // a manual link counts as a way to join: the file is not excluded
    ctrl.updatePlan({ ...st.state.plan!, manualLinks: { s2: [{ baseRow: 3, otherRow: 2 }] } });
    expect(excludedSourceIds(st.state.analysis, st.state.plan)).toEqual(['s3']);
  });
});

describe('withoutSources (pure)', () => {
  it('drops the reports / items / appended rows / imports of the excluded files and recomputes the summary', async () => {
    const j = await joined();
    const plan: JoinPlan = { ...j.plan, joinType: 'full', tiers: { ...j.plan.tiers, s3: [] } };
    const raw = svcJoin(j, plan);
    const out = withoutSources(raw, ['s3']);
    expect(Object.keys(out.reports)).toEqual(['s2']);
    expect(out.rows.every((r) => r.baseRow !== null ? Object.keys(r.matches).join() === 's2' : r.appendedFrom?.sourceId === 's2')).toBe(true);
    expect(out.summary.appended).toBe(out.rows.filter((r) => r.baseRow === null).length);
    expect(out.summary.matchedAll).toBe(out.summary.matchedAny);
    expect(out.defaultImports.some((d) => d.sourceId === 's3')).toBe(false);
    expect(out.items.some((it) => it.sourceId === 's3')).toBe(false);
    expect(withoutSources(raw, [])).toBe(raw);
  });
});

describe('manual key picker (controller)', () => {
  it('adds the candidate checked at the lowest priority and recomputes the join', async () => {
    const { st, ctrl } = await loaded();
    const before = st.state.analysis!;
    const res = ctrl.addManualKey('s2', [3], [2]); // 氏名 ⟷ 氏名 is an auto candidate already
    expect(res).toMatchObject({ ok: true, existing: true, id: 'name' });
    // composite of two columns (氏名 + 大学) — the mock has no such column pair in 職場回答 -> error
    const bad = ctrl.addManualKey('s2', [3, 5], [2]);
    expect(bad).toEqual({ ok: false, error: '基準ファイルとこのファイルで同じ数の列を選んでください' });
    expect(st.state.analysis).toBe(before); // nothing changed
    const add = ctrl.addManualKey('s3', [3], [2]);
    expect(add.ok).toBe(true);
    const pair = st.state.analysis!.pairs.s3;
    const c = pair.candidates.find((x) => x.id === 'manual:3|2')!;
    expect(c).toMatchObject({ kind: 'single', label: '氏名 ⟷ 参加者', selected: false, matchedOther: 20, nonBlankOther: 20 });
    expect(st.state.plan!.tiers.s3).toEqual(['manual:3|2']);
    expect(st.state.manualKeys).toEqual({ s3: [{ id: 'manual:3|2', baseCols: [3], otherCols: [2] }] });
    expect(st.state.joinResult!.reports.s3.matched).toBe(20);
    expect(candidateByColumns(pair, [3], [2])?.id).toBe('manual:3|2');
    // a second manual key goes after the first one
    const second = ctrl.addManualKey('s3', [1], [1]); // No ⟷ 番号
    expect(second.ok).toBe(true);
    expect(st.state.plan!.tiers.s3).toEqual(['manual:3|2', 'manual:1|1']);
  });

  it('reports the Japanese error message of the keys module (never throws)', async () => {
    const { st, ctrl } = await loaded({
      addManualCandidate: () => {
        throw new Error('列が範囲外です');
      },
    });
    const before = st.state.analysis;
    expect(ctrl.addManualKey('s3', [3], [2])).toEqual({ ok: false, error: '列が範囲外です' });
    expect(st.state.analysis).toBe(before);
    expect(ctrl.addManualKey('zz', [3], [2])).toMatchObject({ ok: false });
  });

  it('is refused after the merge (the keys are frozen until 「結合をやり直す」)', async () => {
    const env = await loaded();
    await env.ctrl.merge();
    const res = env.ctrl.addManualKey('s3', [3], [2]);
    expect(res).toMatchObject({ ok: false });
    expect((res as { error: string }).error).toContain('結合後');
  });

  it('× removes a manual candidate by analysing again without it; the other choices stay', async () => {
    const { st, ctrl } = await loaded();
    ctrl.updatePlan({ ...st.state.plan!, tiers: { ...st.state.plan!.tiers, s2: ['id'] }, duplicatePolicy: 'last' });
    expect(ctrl.addManualKey('s3', [3], [2]).ok).toBe(true);
    expect(ctrl.addManualKey('s3', [1], [1]).ok).toBe(true);
    await ctrl.removeManualKey('s3', 'manual:3|2');
    const s = st.state;
    expect(s.analysis!.pairs.s3.candidates.map((c) => c.id)).toEqual(['manual:1|1']);
    expect(s.plan!.tiers).toEqual({ s2: ['id'], s3: ['manual:1|1'] });
    expect(s.plan!.duplicatePolicy).toBe('last');
    expect(s.manualKeys.s3.map((m) => m.id)).toEqual(['manual:1|1']);
    expect(s.busy).toBeNull();
    // removing the last one leaves the file without a key again: excluded
    await ctrl.removeManualKey('s3', 'manual:1|1');
    expect(st.state.analysis!.pairs.s3.candidates).toEqual([]);
    expect(st.state.manualKeys.s3 ?? []).toEqual([]);
    expect(excludedSourceIds(st.state.analysis, st.state.plan)).toEqual(['s3']);
    expect(Object.keys(st.state.joinResult!.reports)).toEqual(['s2']);
    // the same columns can be added again
    expect(ctrl.addManualKey('s3', [3], [2])).toMatchObject({ ok: true, existing: false, id: 'manual:3|2' });
  });

  it('manual keys survive a re-analysis of the same files (options) but not a change of the base', async () => {
    const { st, ctrl } = await loaded();
    expect(ctrl.addManualKey('s3', [3], [2]).ok).toBe(true);
    st.dispatch({ type: 'options/set', options: { itaiji: false } });
    expect(st.state.analysis).toBeNull();
    await ctrl.ensureAnalysis();
    expect(st.state.analysis!.pairs.s3.candidates.map((c) => c.id)).toEqual(['manual:3|2']);
    expect(st.state.plan!.tiers.s3).toEqual(['manual:3|2']);
    expect(st.state.manualKeys.s3).toHaveLength(1);
    expect(excludedSourceIds(st.state.analysis, st.state.plan)).toEqual([]);
    // a different base: fresh plan, the manual keys are gone
    st.dispatch({ type: 'base/set', id: 's2' });
    expect(st.state.manualKeys).toEqual({});
    await ctrl.ensureAnalysis();
    expect(st.state.manualKeys).toEqual({});
    expect(Object.values(st.state.analysis!.pairs).every((p) => p.candidates.every((c) => !c.id.startsWith('manual:')))).toBe(true);
  });

  it('a manual key whose columns do not exist any more is dropped quietly by the next analysis', async () => {
    const { st, ctrl, services } = await loaded();
    expect(ctrl.addManualKey('s3', [3], [2]).ok).toBe(true);
    services.addManualCandidate = () => {
      throw new Error('列がありません');
    };
    st.dispatch({ type: 'options/set', options: { itaiji: false } });
    await ctrl.ensureAnalysis();
    expect(st.state.analysis!.pairs.s3.candidates).toEqual([]);
    expect(st.state.manualKeys.s3 ?? []).toEqual([]);
    expect(st.state.plan!.tiers.s3).toEqual([]); // the chosen id is gone: the automatic plan (no key)
    expect(excludedSourceIds(st.state.analysis, st.state.plan)).toEqual(['s3']);
    expect(st.state.keyError).toBeNull();
  });
});

function svcJoin(j: Awaited<ReturnType<typeof joined>>, plan: JoinPlan) {
  return createMockServices('?mock=1&delay=0').joinSources(j.analysis, plan);
}

describe('toggle helpers used by the picker', () => {
  it('the new candidate is appended last when toggled on', async () => {
    const j = await joined();
    expect(toggleCandidate(j.plan, 's2', 'x', true).tiers.s2).toEqual(['id', 'name', 'x']);
  });
});
