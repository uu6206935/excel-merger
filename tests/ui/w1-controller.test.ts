/**
 * Round 7 flows of the controller against the ?mock=1 services and the real reducer: file card reorder (files phase / merged phase)
 * and the single-sheet output (「他のシートも出力に含める」).
 */
import { describe, expect, it } from 'vitest';
import { createController } from '../../src/state/controller';
import { createMockServices } from '../../src/state/mock';
import type { Services } from '../../src/state/services';
import { file, store } from './fixtures';

function setup(search = '?mock=1', patch: Partial<Services> = {}) {
  const st = store();
  const services = { ...createMockServices(`${search}&delay=0.05`), ...patch };
  const ctrl = createController({ services, dispatch: st.dispatch, getState: st.getState });
  return { st, ctrl, services };
}

async function loaded(search = '?mock=1', patch: Partial<Services> = {}) {
  const env = setup(search, patch);
  env.ctrl.init();
  await env.ctrl.addFiles([file('学生回答.xlsx'), file('職場回答.xlsx'), file('人事回答.xlsx')]);
  await env.ctrl.idle();
  await env.ctrl.ensureAnalysis();
  return env;
}

async function merged(search = '?mock=1', patch: Partial<Services> = {}) {
  const env = await loaded(search, patch);
  await env.ctrl.merge();
  expect(env.st.state.phase).toBe('merged');
  return env;
}

const present = (env: { st: ReturnType<typeof store> }) => env.st.state.merged!.history.present;
const originsOf = (env: { st: ReturnType<typeof store> }) =>
  present(env).doc.columns.map((c) => (c.origin.kind === 'import' ? c.origin.sourceId : c.origin.kind === 'base' ? 's1' : 'new'));
const runsOf = (list: string[]) => list.filter((x, i) => i === 0 || x !== list[i - 1]);

describe('file cards: reorder', () => {
  it('before the merge: a pure reorder keeps the key plan and join type; the re-analysis follows the new order; the output imports in that order', async () => {
    const { st, ctrl } = await loaded();
    ctrl.updatePlan({ ...st.state.plan!, tiers: { ...st.state.plan!.tiers, s2: ['id'] }, joinType: 'full' });
    const planBefore = st.state.plan!;
    expect(Object.keys(st.state.analysis!.pairs)).toEqual(['s2', 's3']);

    expect(ctrl.reorderFile('s3', 's2')).toBe(true);
    expect(st.state.files.map((f) => f.id)).toEqual(['s1', 's3', 's2']);
    expect(st.state.analysis).toBeNull(); // depends on the order: recomputed
    await ctrl.ensureAnalysis();
    expect(Object.keys(st.state.analysis!.pairs)).toEqual(['s3', 's2']);
    expect(st.state.plan!.tiers.s2).toEqual(['id']);
    expect(st.state.plan!.tiers.s3).toEqual(planBefore.tiers.s3);
    expect(st.state.plan!.joinType).toBe('full');
    // the default imports (and so the planned columns / the merge preview rows) follow the new order
    expect(st.state.joinResult!.defaultImports[0].sourceId).toBe('s3');

    await ctrl.merge();
    expect(runsOf(originsOf({ st }))).toEqual(['s1', 's3', 's2']);
  });

  it('a reorder that changes nothing does not touch the state; a busy app refuses it', async () => {
    const { st, ctrl } = await loaded();
    const files = st.state.files;
    expect(ctrl.reorderFile('s2', 's2')).toBe(false);
    expect(ctrl.reorderFile('s2', 's3')).toBe(false);
    expect(st.state.files).toBe(files);
    st.dispatch({ type: 'busy/set', busy: { message: 'x', ratio: null } });
    expect(ctrl.reorderFile('s3', 's1')).toBe(false);
    expect(st.state.files).toBe(files);
  });

  it("after the merge: moves the file's whole column group in the grid, one undo step with the file name; the group is selected", async () => {
    const env = await merged();
    const before = originsOf(env);
    expect(runsOf(before)).toEqual(['s1', 's2', 's3']);
    expect(env.ctrl.reorderFile('s3', 's2')).toBe(true);
    expect(runsOf(originsOf(env))).toEqual(['s1', 's3', 's2']);
    const h = env.st.state.merged!.history;
    expect(h.past).toHaveLength(1);
    expect(h.present.label).toBe('人事回答の列を移動');
    const sel = env.st.state.selection;
    expect(sel.cols.size).toBe(before.filter((x) => x === 's3').length);
    expect([...sel.cols].every((id) => id.startsWith('i:s3:'))).toBe(true);
    // the files themselves are untouched after the merge (the cards are drawn in the grid's order by the UI)
    expect(env.st.state.files.map((f) => f.id)).toEqual(['s1', 's2', 's3']);
    env.ctrl.undo();
    expect(originsOf(env)).toEqual(before);
  });

  it('after the merge: to the very end / the very front, and a drop where the group already is = no-op (no undo entry)', async () => {
    const env = await merged();
    expect(env.ctrl.reorderFile('s1', null)).toBe(true); // the base group behind all the others
    expect(runsOf(originsOf(env))).toEqual(['s2', 's3', 's1']);
    expect(env.ctrl.reorderFile('s1', null)).toBe(false);
    expect(env.ctrl.reorderFile('s2', 's2')).toBe(false);
    expect(env.ctrl.reorderFile('s2', 's3')).toBe(false);
    expect(env.st.state.merged!.history.past).toHaveLength(1);
    expect(env.ctrl.reorderFile('s1', 's2')).toBe(true);
    expect(runsOf(originsOf(env))).toEqual(['s1', 's2', 's3']);
    expect(env.st.state.merged!.history.past).toHaveLength(2);
  });

  it('after the merge: a file without columns in the table (結合対象外) says so and moves nothing', async () => {
    const { st, ctrl } = await loaded();
    ctrl.updatePlan({ ...st.state.plan!, tiers: { ...st.state.plan!.tiers, s3: [] } });
    await ctrl.merge();
    const past = st.state.merged!.history.past.length;
    expect(ctrl.reorderFile('s3', 's1')).toBe(false);
    expect(st.state.toast?.text).toContain('「人事回答.xlsx」の列は表にありません');
    expect(st.state.merged!.history.past).toHaveLength(past);
  });

  it('after the merge: a title-merge refusal opens the same 「結合を解除して移動」 confirmation; the confirmed move keeps the card label', async () => {
    const real = createMockServices('?mock=1&delay=0.05');
    const env = await merged('?mock=1', {
      moveColumns: (s, ids, gap, opts) =>
        opts?.unmergeTitles
          ? real.moveColumns(s, ids, gap)
          : { ok: false, reason: 'タイトルの結合セルがあります', fix: { kind: 'unmergeTitles', refs: ['A1:K1'], message: '見出しの上のタイトルがセル結合されています' } },
    });
    expect(env.ctrl.reorderFile('s3', 's2')).toBe(false);
    const m = env.st.state.modal;
    expect(m).toMatchObject({ kind: 'confirm', purpose: 'unmergeMove', okLabel: '結合を解除して移動' });
    expect(m && m.kind === 'confirm' ? m.move?.label : null).toBe('人事回答の列を移動');
    expect(env.st.state.merged!.history.past).toHaveLength(0);
    env.ctrl.confirmModal();
    expect(env.st.state.merged!.history.past).toHaveLength(1);
    expect(env.st.state.merged!.history.present.label).toBe('人事回答の列を移動');
    expect(runsOf(originsOf(env))).toEqual(['s1', 's3', 's2']);
  });

  it('after the merge: other refusals are shown like a column drag (toast), nothing changes', async () => {
    const env = await merged('?mock=1', { moveColumns: () => ({ ok: false, reason: '表の途中へは移動できません（Excel と同じ制限です）。' }) });
    expect(env.ctrl.reorderFile('s3', 's2')).toBe(false);
    expect(env.st.state.toast).toMatchObject({ kind: 'error', text: '表の途中へは移動できません' });
    expect(env.st.state.merged!.history.past).toHaveLength(0);
  });
});

describe('「他のシートも出力に含める」', () => {
  it('is off by default: the premerge is asked to drop the other sheets, the merged book has ONE sheet, the toast says what is left out', async () => {
    const real = createMockServices('?mock=1&delay=0.05');
    const seen: Array<boolean | undefined> = [];
    const env = await loaded('?mock=1', {
      buildPremerged: (input, p) => {
        seen.push(input.options.dropOtherSheets);
        return real.buildPremerged(input, p);
      },
    });
    await env.ctrl.merge();
    expect(seen).toEqual([true]);
    const m = env.st.state.merged!;
    expect(m.book.sheets.map((s) => s.name)).toEqual(['応募者一覧']);
    expect(m.report.droppedSheets).toEqual(['集計']);
    expect(env.st.state.toast?.text).toContain('他のシート 1 枚は含めません');
    expect(env.st.state.toast?.detail).toContain('出力は結合した 1 シートのみです（他のシート 1 枚「集計」 は含めません）');
  });

  it('ticked: dropOtherSheets is false, every sheet stays, nothing is reported as dropped', async () => {
    const real = createMockServices('?mock=1&delay=0.05');
    const seen: Array<boolean | undefined> = [];
    const env = await loaded('?mock=1', {
      buildPremerged: (input, p) => {
        seen.push(input.options.dropOtherSheets);
        return real.buildPremerged(input, p);
      },
    });
    env.ctrl.setIncludeOtherSheets(true);
    expect(env.st.state.includeOtherSheets).toBe(true);
    // (round 9) ticking it also ticks 「他のファイルのシートも含める」; this test is about the BASE file's sheets only, so that one is cleared again
    expect(env.st.state.includeSourceSheets).toBe(true);
    env.ctrl.setIncludeSourceSheets(false);
    expect(env.st.state.includeOtherSheets).toBe(true);
    await env.ctrl.merge();
    expect(seen).toEqual([false]);
    expect(env.st.state.merged!.book.sheets.map((s) => s.name)).toEqual(['応募者一覧', '集計']);
    expect(env.st.state.merged!.report.droppedSheets).toBeUndefined();
    expect(env.st.state.toast?.text).not.toContain('他のシート');
  });

  it('the setting cannot change after the merge (the controls are disabled; the reducer ignores it too)', async () => {
    const env = await merged();
    env.ctrl.setIncludeOtherSheets(true);
    expect(env.st.state.includeOtherSheets).toBe(false);
  });

  it('single-sheet output: no formula of another sheet depends on the table any more (the delete impact counts the table only)', async () => {
    const env = await merged();
    const d = present(env).doc;
    expect(d.sheetNames).toEqual(['応募者一覧']);
    expect(d.externalFormulas).toEqual([]);
    const written = d.columns.find((c) => c.cells[d.headerRow - 1]?.v === '筆記')!.id;
    env.ctrl.requestDelete([written]);
    const modal = env.st.state.modal;
    expect(modal?.kind).toBe('deleteConfirm');
    // 30 rows of 合計 = 筆記 + 面接 (the 31st dependent of the old behaviour was 集計!B3 on the dropped sheet)
    expect(modal?.kind === 'deleteConfirm' ? modal.impact.breakCount : -1).toBe(30);
  });

  it('formulas that read a dropped sheet and were turned into values are reported', async () => {
    const real = createMockServices('?mock=1&delay=0.05');
    const env = await loaded('?mock=1', {
      buildPremerged: async (input, p) => {
        const r = await real.buildPremerged(input, p);
        return { ...r, report: { ...r.report, formulasDetached: 3 } };
      },
    });
    await env.ctrl.merge();
    expect(env.st.state.toast?.text).toContain('他のシートを参照していた数式 3 個は値にしました');
  });
});
