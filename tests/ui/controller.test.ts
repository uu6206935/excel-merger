/**
 * The async flows of the app (controller) against the ?mock=1 services and the real reducer:
 * files -> analysis -> merge -> editor ops -> export, with the failure paths.
 */
import { describe, expect, it } from 'vitest';
import { createController, mergedRegion, outputFormatOf, regionProblem } from '../../src/state/controller';
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
const colIdByTitle = (env: { st: ReturnType<typeof store> }, title: string) => {
  const d = present(env).doc;
  return d.columns.find((c) => c.cells[d.headerRow - 1]?.v === title)!.id;
};

describe('files and analysis', () => {
  it('reads files one by one, detects sheets / regions, then analyses and joins', async () => {
    const { st } = await loaded();
    expect(st.state.server.status?.excel.installed).toBe(true);
    expect(st.state.files.map((f) => [f.status, f.sheet, f.region?.headerRow])).toEqual([
      ['ready', '応募者一覧', 2],
      ['ready', '回答', 1],
      ['ready', 'Sheet1', 1],
    ]);
    expect(st.state.analysis).not.toBeNull();
    expect(st.state.plan?.tiers.s2).toEqual(['id', 'name']);
    expect(st.state.joinResult?.summary.baseRows).toBe(30);
    expect(st.state.busy).toBeNull();
    // the analysis does not run again for the same inputs
    const n = st.actions.filter((a) => a.type === 'analysis/done').length;
    expect(n).toBe(1);
  });

  it('refuses unsupported files and skips duplicates / Excel lock files with a Japanese message', async () => {
    const { st, ctrl } = setup();
    await ctrl.addFiles([file('memo.txt'), file('~$名簿.xlsx')]);
    expect(st.state.files).toHaveLength(0);
    expect(st.state.toast?.kind).toBe('error');
    expect(st.state.toast?.text).toContain('対応していない形式です');
    await ctrl.addFiles([file('学生回答.xlsx')]);
    await ctrl.idle();
    await ctrl.addFiles([file('学生回答.xlsx')]);
    expect(st.state.files).toHaveLength(1);
    expect(st.state.toast?.text).toContain('すでに読み込まれています');
  });

  it('a broken / protected / empty file only shows an error on its card', async () => {
    const { st, ctrl } = setup();
    await ctrl.addFiles([file('壊れた.xlsx'), file('パスワード付き.xlsx'), file('空.xlsx', 0)]);
    await ctrl.idle();
    expect(st.state.files.map((f) => f.status)).toEqual(['error', 'error', 'error']);
    expect(st.state.files[0].error).toContain('読み込めませんでした');
    expect(st.state.files[1].error).toContain('パスワード');
    expect(st.state.files[2].error).toContain('空のファイル');
  });

  it('a region that cannot be detected leaves the file usable after choosing the header row', async () => {
    const { st, ctrl } = setup('?mock=1', {
      detectRegion: () => {
        throw new Error('見出しなし');
      },
    });
    await ctrl.addFiles([file('学生回答.xlsx')]);
    await ctrl.idle();
    const f = st.state.files[0];
    expect(f.status).toBe('ready');
    expect(f.region).toBeUndefined();
    expect(f.error).toContain('見出し行を自動で見つけられませんでした');
    ctrl.changeHeaderRow(f.id, 2);
    expect(st.state.files[0].region).toMatchObject({ headerRow: 2, firstDataRow: 3, lastDataRow: 32, firstCol: 1, lastCol: 11 });
    expect(st.state.files[0].error).toBeUndefined();
    ctrl.changeHeaderRow(f.id, 32);
    expect(st.state.files[0].error).toContain('データがありません');
  });

  it('analysis errors become a Japanese key error, not a crash', async () => {
    const { st } = await loaded('?mock=1', {
      analyzeSources: () => {
        throw new Error('boom');
      },
    });
    expect(st.state.analysis).toBeNull();
    // an unexpected (English) error is not shown raw: a Japanese sentence, the original text stays as a short technical note
    expect(st.state.keyError).toBe('結合キーを分析できませんでした: 予期しないエラーが発生しました［技術情報: boom］');
    expect(st.state.busy).toBeNull();
  });

  it('options change -> re-analysis keeps the user tiers; base change -> fresh plan', async () => {
    const { st, ctrl } = await loaded();
    ctrl.updatePlan({ ...st.state.plan!, tiers: { ...st.state.plan!.tiers, s2: ['id'] }, joinType: 'full' });
    expect(st.state.joinResult?.summary.appended).toBe(2);
    st.dispatch({ type: 'options/set', options: { itaiji: false } });
    expect(st.state.analysis).toBeNull();
    await ctrl.ensureAnalysis();
    expect(st.state.plan?.tiers.s2).toEqual(['id']);
    expect(st.state.plan?.joinType).toBe('full');
    st.dispatch({ type: 'base/set', id: 's2' });
    await ctrl.ensureAnalysis();
    expect(st.state.analysis?.baseId).toBe('s2');
    expect(st.state.plan?.joinType).toBe('left');
  });

  it('a file without a key does not block the merge: it is excluded (round 4) — details in excluded.test.ts', async () => {
    const { st, ctrl } = await loaded();
    ctrl.updatePlan({ ...st.state.plan!, tiers: { ...st.state.plan!.tiers, s3: [] } });
    await ctrl.merge();
    expect(st.state.phase).toBe('merged');
    expect(st.state.toast?.text).toContain('結合対象外 1 ファイル（人事回答）');
    expect(present({ st }).doc.columns).toHaveLength(15); // 11 base columns + 4 of 職場回答, none of 人事回答
  });

  it('merge is refused with a message when NO other file has a key', async () => {
    const { st, ctrl } = await loaded();
    ctrl.updatePlan({ ...st.state.plan!, tiers: { s2: [], s3: [] } });
    await ctrl.merge();
    expect(st.state.phase).toBe('files');
    expect(st.state.toast).toMatchObject({ kind: 'error' });
    expect(st.state.toast?.text).toContain('どのファイルにも結合キーが選ばれていません');
  });

  it('errors that are already Japanese (modules, server) are shown as they are', async () => {
    const { st, ctrl } = await loaded('?mock=1', {
      buildPremerged: async () => {
        throw new Error('出力フォルダがいっぱいです');
      },
    });
    await ctrl.merge();
    expect(st.state.toast).toMatchObject({ kind: 'error', text: '結合できませんでした: 出力フォルダがいっぱいです' });
    expect(st.state.toast?.detail).toBeUndefined();
  });

  it('merge failures are reported and leave the key screen usable', async () => {
    const { st, ctrl } = await loaded('?mock=1', {
      buildPremerged: async () => {
        throw new Error('disk full');
      },
    });
    await ctrl.merge();
    expect(st.state.phase).toBe('files');
    expect(st.state.toast).toMatchObject({ kind: 'error', text: '結合できませんでした: 予期しないエラーが発生しました', detail: 'disk full' });
    expect(st.state.busy).toBeNull();
  });
});

describe('editor', () => {
  it('merge builds the editor from the read-back premerged book', async () => {
    const env = await merged();
    const d = present(env).doc;
    expect(d.columns).toHaveLength(18);
    expect(d.firstDataRow).toBe(3);
    expect(d.lastDataRow).toBe(32);
    expect(env.st.state.toast?.text).toMatch(/^結合しました: 7 列を追加/);
    expect(env.st.state.merged?.format).toBe('xlsx');
  });

  it('move / no-op move / insert / undo / redo', async () => {
    const env = await merged();
    const { ctrl, st } = env;
    const ids = present(env).doc.columns.map((c) => c.id);
    ctrl.moveColumns([ids[4], ids[6]], 1);
    expect(present(env).doc.columns.slice(0, 4).map((c) => c.id)).toEqual([ids[0], ids[4], ids[6], ids[1]]);
    expect([...st.state.selection.cols]).toEqual([ids[4], ids[6]]);
    const past = st.state.merged!.history.past.length;
    ctrl.moveColumns([ids[0]], 0); // already there
    expect(st.state.merged!.history.past.length).toBe(past);
    ctrl.insertColumns(2, 2);
    const d = present(env).doc;
    expect(d.columns[2].origin.kind).toBe('new');
    expect([...st.state.selection.cols]).toEqual([d.columns[2].id, d.columns[3].id]);
    expect(st.state.active?.col).toBe(d.columns[2].id);
    ctrl.undo();
    ctrl.undo();
    expect(present(env).doc.columns.map((c) => c.id)).toEqual(ids);
    ctrl.redo();
    expect(present(env).doc.columns[1].id).toBe(ids[4]);
  });

  it('delete with dependents asks, then converts the formulas to values (never #REF!)', async () => {
    // (round 7: the default output holds ONLY the merged sheet, so the formula of the 集計 sheet that reads 筆記 is not part of it;
    //  the other sheets are kept here to keep covering the formulas of another sheet: 30 rows of 合計 + 集計!B3 = 31)
    const env = await loaded();
    env.ctrl.setIncludeOtherSheets(true);
    await env.ctrl.merge();
    const { ctrl, st } = env;
    const written = colIdByTitle(env, '筆記');
    ctrl.requestDelete([written]);
    expect(st.state.modal?.kind).toBe('deleteConfirm');
    if (st.state.modal?.kind !== 'deleteConfirm') throw new Error();
    expect(st.state.modal.impact.breakCount).toBe(31);
    ctrl.confirmDelete();
    expect(st.state.modal).toBeNull();
    const d = present(env).doc;
    expect(d.columns.find((c) => c.id === written)).toBeUndefined();
    const total = d.columns.find((c) => c.cells[d.headerRow - 1]?.v === '合計')!;
    expect(total.cells[2]?.f).toBeUndefined();
    expect(typeof total.cells[2]?.v).toBe('number');
    for (const c of d.columns) for (const cell of c.cells) expect(cell?.f ?? '').not.toContain('#REF!');
  });

  it('delete without dependents happens at once; deleting every column is refused', async () => {
    const env = await merged();
    const { ctrl, st } = env;
    const mail = colIdByTitle(env, 'メール');
    ctrl.requestDelete([mail]);
    expect(st.state.modal).toBeNull();
    expect(present(env).doc.columns.some((c) => c.id === mail)).toBe(false);
    ctrl.requestDelete(present(env).doc.columns.map((c) => c.id));
    expect(st.state.toast?.text).toContain('すべての列を削除することはできません');
  });

  it('cell edits: unchanged text is not an edit; refusals are shown', async () => {
    const env = await merged();
    const { ctrl, st } = env;
    const name = colIdByTitle(env, '氏名');
    const past = st.state.merged!.history.past.length;
    expect(ctrl.editCell({ row: 3, col: name }, '青木 翔太')).toBe(false);
    expect(st.state.merged!.history.past.length).toBe(past);
    expect(ctrl.editCell({ row: 3, col: name }, '青木 翔')).toBe(true);
    expect(present(env).doc.columns.find((c) => c.id === name)!.cells[2]?.v).toBe('青木 翔');
    const total = colIdByTitle(env, '合計');
    expect(ctrl.editCell({ row: 3, col: total }, '=I3+J3')).toBe(false); // the formula as shown
  });

  it('refused ops show their Japanese reason', async () => {
    const env = await merged('?mock=1', { moveColumns: () => ({ ok: false, reason: 'テーブルの一部は移動できません' }) });
    env.ctrl.moveColumns([present(env).doc.columns[0].id], 5);
    expect(env.st.state.toast).toMatchObject({ kind: 'error', text: 'テーブルの一部は移動できません' });
  });

  it('a long refusal becomes a short toast; the full reason is its detail', async () => {
    const reason = 'テーブル「表1」の中へ、テーブル外の列を移動することはできません。（Excel と同じ制限です）';
    const env = await merged('?mock=1', { moveColumns: () => ({ ok: false, reason }) });
    expect(env.ctrl.moveColumns([present(env).doc.columns[0].id], 5)).toBe(false);
    expect(env.st.state.toast).toMatchObject({ kind: 'error', text: 'テーブル「表1」の中へ、テーブル外の列を移動できません', detail: reason });
    expect(env.st.state.modal).toBeNull();
  });

  it('merged title cells: asks 「結合を解除して移動」, then unmerges + moves as ONE undo step', async () => {
    const env = await merged('?mock=1&titlemerge=1');
    const { ctrl, st } = env;
    const ids = present(env).doc.columns.map((c) => c.id);
    const past = st.state.merged!.history.past.length;
    expect(ctrl.moveColumns([ids[4]], 1)).toBe(false); // E inside the title merge A1:K1
    expect(st.state.modal).toMatchObject({ kind: 'confirm', purpose: 'unmergeMove', okLabel: '結合を解除して移動', move: { ids: [ids[4]], gap: 1, refs: ['A1:K1'] } });
    if (st.state.modal?.kind !== 'confirm') throw new Error();
    expect(st.state.modal.message).toContain('A1:K1');
    expect(st.state.modal.danger).toBeFalsy();
    expect(st.state.toast?.kind).not.toBe('error');
    ctrl.confirmModal();
    expect(st.state.modal).toBeNull();
    const d = present(env).doc;
    expect(d.columns.slice(0, 2).map((c) => c.id)).toEqual([ids[0], ids[4]]);
    expect(d.features.merges).toEqual([]);
    const log = present(env).log;
    expect(log[0]).toEqual({ kind: 'unmerge', sheet: d.sheetName, refs: ['A1:K1'], centerAcross: true });
    expect(log[1].kind).toBe('struct');
    expect(st.state.merged!.history.past.length).toBe(past + 1);
    expect([...st.state.selection.cols]).toEqual([ids[4]]);
    expect(st.state.toast).toMatchObject({ kind: 'info', text: 'タイトル行の結合（A1:K1）を解除し、「選択範囲内で中央」にしました。' });
    ctrl.undo();
    expect(present(env).doc.columns.map((c) => c.id)).toEqual(ids);
    expect(present(env).doc.features.merges).toHaveLength(1);
  });

  it('cancelling the unmerge question changes nothing; imported columns outside the title move freely', async () => {
    const env = await merged('?mock=1&titlemerge=1');
    const { ctrl, st } = env;
    const ids = present(env).doc.columns.map((c) => c.id);
    ctrl.moveColumns([ids[2]], 0);
    expect(st.state.modal?.kind).toBe('confirm');
    st.dispatch({ type: 'modal/close' });
    expect(present(env).doc.columns.map((c) => c.id)).toEqual(ids);
    expect(st.state.merged!.history.past.length).toBe(0);
    // imported columns (after K) swap without touching the title
    expect(ctrl.moveColumns([ids[13]], 11)).toBe(true);
    expect(st.state.modal).toBeNull();
    expect(present(env).log.every((op) => op.kind === 'struct')).toBe(true);
  });

  it('the pending unmerge-move is dropped when the table changed meanwhile; a second refusal is only a toast', async () => {
    const env = await merged('?mock=1&titlemerge=1');
    const { ctrl, st } = env;
    const ids = present(env).doc.columns.map((c) => c.id);
    ctrl.moveColumns([ids[4]], 1);
    ctrl.insertColumns(15, 1); // (not possible from the UI while the dialog is open)
    ctrl.confirmModal();
    expect(st.state.toast).toMatchObject({ kind: 'error' });
    expect(st.state.toast?.text).toContain('移動を取りやめました');
    expect(present(env).doc.columns.slice(0, 2).map((c) => c.id)).toEqual([ids[0], ids[1]]);

    const env2 = await merged('?mock=1&titlemerge=1', {
      moveColumns: () => ({
        ok: false,
        reason: '結合されたセル（A1:K1）の一部の列だけを移動することはできません。',
        fix: { kind: 'unmergeTitles', refs: ['A1:K1'], message: 'タイトルの結合を解除しますか' },
      }),
    });
    env2.ctrl.moveColumns([present(env2).doc.columns[4].id], 1);
    env2.ctrl.confirmModal();
    expect(env2.st.state.modal).toBeNull();
    expect(env2.st.state.toast).toMatchObject({ kind: 'error', text: '結合されたセル（A1:K1）の一部の列だけを移動できません' });
  });

  it('several warnings of a successful op: one short toast, the full list as its detail', async () => {
    const mock = createMockServices('?mock=1&delay=0.05');
    const env = await merged('?mock=1', {
      insertColumns: (s, i, n) => {
        const r = mock.insertColumns(s, i, n);
        return r.ok ? { ...r, warnings: ['列番号で列を指定する関数を使った数式が 1 個あります。別の列の値を返す可能性があります。', '集計!B28: =VLOOKUP("1",名簿!$A$4:$K$63,2,FALSE)'] } : r;
      },
    });
    env.ctrl.insertColumns(2, 1);
    expect(env.st.state.toast).toMatchObject({ kind: 'info', text: '列番号で列を指定する関数を使った数式が 1 個あります。（ほか 1 件の注意）' });
    expect(env.st.state.toast?.detail).toContain('集計!B28');
  });

  it('rename header, resize (live + commit + cancel), reset widths', async () => {
    const env = await merged();
    const { ctrl, st } = env;
    const id = colIdByTitle(env, '学部');
    ctrl.renameColumn(id, ' 所属学部 ');
    const d = present(env).doc;
    expect(d.columns.find((c) => c.id === id)!.cells[d.headerRow - 1]?.v).toBe('所属学部');
    ctrl.renameColumn(id, '  ');
    expect(st.state.toast?.text).toBe('見出しを入力してください');

    const before = st.state.merged!.history.past.length;
    ctrl.resizeColumn(id, 150, null);
    expect(st.state.liveWidths[id]).toBe(150);
    ctrl.resizeColumn(id, 150, { from: 150 }); // cancelled gesture
    expect(st.state.liveWidths[id]).toBeUndefined();
    expect(st.state.merged!.history.past.length).toBe(before);
    ctrl.resizeColumn(id, 200, { from: 109 });
    const w = present(env).doc.columns.find((c) => c.id === id)!.props.width!;
    expect(w).toBeGreaterThan(20);
    expect(st.state.liveWidths[id]).toBeUndefined();
    ctrl.resetColumnWidths([id]);
    expect(present(env).doc.columns.find((c) => c.id === id)!.props.width).toBe(st.state.merged!.initialWidths[id]);
  });
});

describe('header row / last data row (keys.detectRegion with the user choices)', () => {
  const spy = (env: ReturnType<typeof setup>) => {
    const calls: unknown[] = [];
    const real = env.services.detectRegion;
    env.services.detectRegion = (b, sh, o) => {
      calls.push(o);
      return real(b, sh, o);
    };
    return calls;
  };

  it('the header row goes through detectRegion(book, sheet, { headerRow })', async () => {
    const env = setup('?mock=1');
    const calls = spy(env);
    await env.ctrl.addFiles([file('学生回答.xlsx')]);
    await env.ctrl.idle();
    const id = env.st.state.files[0].id;
    env.ctrl.changeHeaderRow(id, 1);
    expect(calls.at(-1)).toEqual({ headerRow: 1 });
    expect(env.st.state.files[0].region).toMatchObject({ headerRow: 1, firstDataRow: 2, lastDataRow: 32 });
    env.ctrl.changeHeaderRow(id, 2);
    expect(env.st.state.files[0].region).toMatchObject({ headerRow: 2, firstDataRow: 3, lastDataRow: 32, lastCol: 11 });
  });

  it('falls back to the local helper when detectRegion ignores the chosen header row', async () => {
    const env = setup('?mock=1');
    const real = env.services.detectRegion;
    env.services.detectRegion = (b, sh) => real(b, sh); // an implementation without opts
    await env.ctrl.addFiles([file('学生回答.xlsx')]);
    await env.ctrl.idle();
    env.ctrl.changeHeaderRow(env.st.state.files[0].id, 1);
    expect(env.st.state.files[0].region).toMatchObject({ headerRow: 1, firstDataRow: 2, lastDataRow: 32 });
  });

  it('最終行: the chosen row (clamped), then 「自動」 back to the detected one', async () => {
    const env = setup('?mock=1');
    const calls = spy(env);
    await env.ctrl.addFiles([file('学生回答.xlsx'), file('職場回答.xlsx')]);
    await env.ctrl.idle();
    await env.ctrl.ensureAnalysis();
    const id = env.st.state.files[0].id;
    expect(env.st.state.joinResult!.summary.baseRows).toBe(30);
    env.ctrl.changeLastDataRow(id, 12);
    expect(calls.at(-1)).toEqual({ headerRow: 2, lastDataRow: 12 });
    expect(env.st.state.files[0]).toMatchObject({ region: { headerRow: 2, firstDataRow: 3, lastDataRow: 12 }, lastRowManual: true });
    expect(env.st.state.analysis).toBeNull(); // re-analysed with the shorter table
    await env.ctrl.ensureAnalysis();
    expect(env.st.state.joinResult!.summary.baseRows).toBe(10);
    env.ctrl.changeLastDataRow(id, 1); // above the data -> the first data row
    expect(env.st.state.files[0].region?.lastDataRow).toBe(3);
    env.ctrl.changeLastDataRow(id, 9999); // below the sheet -> the last row of the sheet
    expect(env.st.state.files[0].region?.lastDataRow).toBe(32);
    env.ctrl.changeLastDataRow(id, 20);
    env.ctrl.changeLastDataRow(id, null);
    expect(env.st.state.files[0].region?.lastDataRow).toBe(32);
    expect(env.st.state.files[0].lastRowManual).toBeUndefined();
    // a new header row resets the manual last row
    env.ctrl.changeLastDataRow(id, 20);
    env.ctrl.changeHeaderRow(id, 2);
    expect(env.st.state.files[0].lastRowManual).toBeUndefined();
  });

  it('the last data row falls back locally when detectRegion throws', async () => {
    const env = setup('?mock=1');
    await env.ctrl.addFiles([file('学生回答.xlsx')]);
    await env.ctrl.idle();
    env.services.detectRegion = () => {
      throw new Error('boom');
    };
    const id = env.st.state.files[0].id;
    env.ctrl.changeLastDataRow(id, 10);
    expect(env.st.state.files[0].region).toMatchObject({ headerRow: 2, firstDataRow: 3, lastDataRow: 10 });
    env.ctrl.changeHeaderRow(id, 2);
    expect(env.st.state.files[0].region).toMatchObject({ headerRow: 2, firstDataRow: 3 });
  });
});

describe('re-merge and adding files after the merge', () => {
  it('without edits returns directly; with edits asks first', async () => {
    const env = await merged();
    const { ctrl, st } = env;
    ctrl.requestRemerge();
    expect(st.state.phase).toBe('files');
    await ctrl.merge();
    ctrl.insertColumns(0, 1);
    ctrl.requestRemerge();
    expect(st.state.modal).toMatchObject({ kind: 'confirm', purpose: 'remerge' });
    ctrl.confirmModal();
    expect(st.state.phase).toBe('files');
    expect(st.state.analysis).not.toBeNull();
  });

  it('dropping files after the merge asks, then adds them on the key screen', async () => {
    const env = await merged();
    const { ctrl, st } = env;
    await ctrl.addFiles([file('追加データ.xlsx')]);
    expect(st.state.modal).toMatchObject({ kind: 'confirm', purpose: 'addFiles' });
    expect(st.state.files).toHaveLength(3);
    ctrl.confirmModal();
    await ctrl.idle();
    expect(st.state.phase).toBe('files');
    expect(st.state.files).toHaveLength(4);
    expect(st.state.files[3].status).toBe('ready');
  });
});

describe('export', () => {
  it('Excel engine: progress, then saved into the output folder', async () => {
    const env = await merged();
    const { ctrl, st } = env;
    ctrl.moveColumns([present(env).doc.columns[5].id], 0);
    ctrl.openExport();
    expect(st.state.modal?.kind).toBe('export');
    const p = ctrl.runExport({ fileName: '結果', engine: 'excel' });
    expect(st.state.exportState.phase).toBe('running');
    expect(st.state.exportState.engine).toBe('excel');
    await p;
    expect(st.state.exportState).toMatchObject({ phase: 'done', engine: 'excel', fellBack: false, savedName: '結果.xlsx' });
    expect(st.state.exportState.savedPath).toMatch(/\\結果\.xlsx$/);
    await ctrl.openSavedFile();
    expect(st.state.toast?.text).toContain('Excel で開いています');
  });

  it('Excel engine warnings are kept with the result', async () => {
    const env = await merged('?mock=excelwarn');
    await env.ctrl.runExport({ fileName: 'a', engine: 'excel' });
    expect(env.st.state.exportState.phase).toBe('done');
    expect(env.st.state.exportState.warnings).toHaveLength(2);
    expect(env.st.state.exportState.warnings![0]).toContain('マクロ');
    // the built-in engine has none; a new export replaces the old ones
    await env.ctrl.runExport({ fileName: 'b', engine: 'js' });
    expect(env.st.state.exportState.warnings).toEqual([]);
  });

  it('an older api client that returns bare bytes still works', async () => {
    const env = await merged('?mock=1', { runExcelJob: (async () => new Uint8Array([1, 2, 3])) as unknown as Services['runExcelJob'] });
    await env.ctrl.runExport({ fileName: 'a', engine: 'excel' });
    expect(env.st.state.exportState).toMatchObject({ phase: 'done', engine: 'excel', warnings: [] });
  });

  it('Excel engine failure falls back to the built-in engine with a toast', async () => {
    const env = await merged('?mock=excelfail');
    await env.ctrl.runExport({ fileName: 'a', engine: 'excel' });
    expect(env.st.state.exportState).toMatchObject({ phase: 'done', engine: 'js', fellBack: true });
    expect(
      env.st.actions.some((a) => a.type === 'toast/show' && a.kind === 'error' && a.text.includes('内蔵エンジンで出力します') && !!a.detail?.includes('タイムアウト')),
    ).toBe(true);
  });

  it('without Excel the Excel engine request uses the built-in engine', async () => {
    const env = await merged('?mock=noexcel');
    await env.ctrl.runExport({ fileName: 'a', engine: 'excel' });
    expect(env.st.state.exportState).toMatchObject({ phase: 'done', engine: 'js', fellBack: false });
  });

  it('errors of both engines end in an error state with a message', async () => {
    const env = await merged('?mock=excelfail', {
      exportWithJs: async () => {
        throw new Error('zip error');
      },
    });
    await env.ctrl.runExport({ fileName: 'a', engine: 'excel' });
    expect(env.st.state.exportState.phase).toBe('error');
    expect(env.st.state.exportState.error).toBe('出力できませんでした: 予期しないエラーが発生しました［技術情報: zip error］');
  });

  it('cancel stops the Excel job', async () => {
    const env = await merged();
    const p = env.ctrl.runExport({ fileName: 'a', engine: 'excel' });
    await new Promise((r) => setTimeout(r, 10));
    env.ctrl.cancelExport();
    await p;
    expect(env.st.state.exportState.phase).toBe('idle');
    expect(env.st.state.toast?.text).toBe('出力を中止しました');
  });
});

describe('small helpers', () => {
  it('output format and merged region', () => {
    expect(outputFormatOf(undefined)).toBe('xlsx');
    expect(mergedRegion({ headerRow: 2, firstDataRow: 3, lastDataRow: 32 }, [{ row: 33 }, { row: 34 }])).toEqual({ headerRow: 2, firstDataRow: 3, lastDataRow: 34 });
    expect(mergedRegion({ headerRow: 1, firstDataRow: 2, lastDataRow: 9 }, [])).toEqual({ headerRow: 1, firstDataRow: 2, lastDataRow: 9 });
    expect(regionProblem({ firstDataRow: 5, lastDataRow: 4 })).toContain('データ行がありません');
    expect(regionProblem({ firstDataRow: 2, lastDataRow: 2 })).toBeUndefined();
    expect(regionProblem(null)).toBeUndefined();
  });
});
