/**
 * Round 9 (plan A) flows of the controller against the ?mock=1 services and the real reducer: the two sheet options reach
 * buildPremerged, the merge result (names, report, layout), the toast, the 警告 notes, the preview list = the engine's list,
 * and that the settings are merge settings.
 */
import { describe, expect, it } from 'vitest';
import { createController } from '../../src/state/controller';
import { createMockServices } from '../../src/state/mock';
import { fidelityNotes } from '../../src/state/notes';
import { effectiveBaseId, excludedSourceIds } from '../../src/state/keysPlan';
import { outputSheetCounts, outputSheetRows, outputSheetsLine, plannedOutputSheets } from '../../src/state/sheetPlan';
import type { Services } from '../../src/state/services';
import { file, store } from './fixtures';

function setup(patch: Partial<Services> = {}) {
  const st = store();
  const services = { ...createMockServices('?mock=1&delay=0.05'), ...patch };
  const ctrl = createController({ services, dispatch: st.dispatch, getState: st.getState });
  return { st, ctrl, services };
}

async function loaded(names: string[], patch: Partial<Services> = {}) {
  const env = setup(patch);
  env.ctrl.init();
  await env.ctrl.addFiles(names.map((n) => file(n)));
  await env.ctrl.idle();
  await env.ctrl.ensureAnalysis();
  return env;
}

const THREE = ['学生回答.xlsx', '職場回答.xlsx', '人事回答.xlsx'];
const sheetNames = (env: { st: ReturnType<typeof store> }) => env.st.state.merged!.book.sheets.map((s) => s.name);

/** records the options buildPremerged was called with */
function spy() {
  const real = createMockServices('?mock=1&delay=0.05');
  const seen: Array<{ drop: boolean | undefined; source: boolean | undefined }> = [];
  return {
    seen,
    patch: {
      buildPremerged: (input: Parameters<Services['buildPremerged']>[0], p?: Parameters<Services['buildPremerged']>[1]) => {
        seen.push({ drop: input.options.dropOtherSheets, source: input.options.includeSourceSheets });
        return real.buildPremerged(input, p);
      },
    } satisfies Partial<Services>,
  };
}

describe('「他のファイルのシートも含める」 reaches the premerge', () => {
  it('default: single-sheet output, includeSourceSheets false, no copied sheet, no 「シート」 words in the toast', async () => {
    const s = spy();
    const env = await loaded(THREE, s.patch);
    await env.ctrl.merge();
    expect(s.seen).toEqual([{ drop: true, source: false }]);
    expect(sheetNames(env)).toEqual(['応募者一覧']);
    expect(env.st.state.merged!.report.addedSheets).toBeUndefined();
    expect(env.st.state.merged!.layout.addedSheets).toBeUndefined();
    expect(env.st.state.toast?.text).not.toContain('他のファイルのシート');
  });

  it('first checkbox on: both options reach the premerge (dropOtherSheets false, includeSourceSheets true)', async () => {
    const s = spy();
    const env = await loaded(THREE, s.patch);
    env.ctrl.setIncludeOtherSheets(true);
    await env.ctrl.merge();
    expect(s.seen).toEqual([{ drop: false, source: true }]);
    expect(sheetNames(env)).toEqual(['応募者一覧', '集計', '職場回答_回答', '人事回答_Sheet1']);
  });

  it('second checkbox alone ticks the first as well; then the user can clear just the second', async () => {
    const s = spy();
    const env = await loaded(THREE, s.patch);
    env.ctrl.setIncludeSourceSheets(true);
    expect(env.st.state).toMatchObject({ includeOtherSheets: true, includeSourceSheets: true });
    env.ctrl.setIncludeSourceSheets(false);
    expect(env.st.state).toMatchObject({ includeOtherSheets: true, includeSourceSheets: false });
    env.ctrl.setIncludeOtherSheets(false);
    expect(env.st.state).toMatchObject({ includeOtherSheets: false, includeSourceSheets: false });
    env.ctrl.setIncludeSourceSheets(true);
    await env.ctrl.merge();
    expect(s.seen).toEqual([{ drop: false, source: true }]);
  });

  it('the settings cannot change after the merge, and survive 結合をやり直す', async () => {
    const env = await loaded(THREE);
    env.ctrl.setIncludeSourceSheets(true);
    await env.ctrl.merge();
    env.ctrl.setIncludeSourceSheets(false);
    env.ctrl.setIncludeOtherSheets(false);
    expect(env.st.state).toMatchObject({ includeOtherSheets: true, includeSourceSheets: true });
    env.ctrl.requestRemerge();
    expect(env.st.state.phase).toBe('files');
    expect(env.st.state).toMatchObject({ includeOtherSheets: true, includeSourceSheets: true });
    env.ctrl.setIncludeOtherSheets(false);
    expect(env.st.state).toMatchObject({ includeOtherSheets: false, includeSourceSheets: false });
  });
});

describe('the merge result with the sheets of the other files', () => {
  it('names, report, layout, toast and side line agree', async () => {
    const env = await loaded(THREE);
    env.ctrl.setIncludeSourceSheets(true);
    await env.ctrl.merge();
    const m = env.st.state.merged!;
    expect(m.report.addedSheets).toEqual(['職場回答_回答', '人事回答_Sheet1']);
    expect(m.layout.addedSheets).toEqual([
      { sourceId: 's2', file: '職場回答.xlsx', original: '回答', name: '職場回答_回答', suffixed: false, complete: true },
      { sourceId: 's3', file: '人事回答.xlsx', original: 'Sheet1', name: '人事回答_Sheet1', suffixed: false, complete: true },
    ]);
    // the merged sheet is the first sheet, then the base's sheets, then the copies
    expect(sheetNames(env)).toEqual(['応募者一覧', '集計', '職場回答_回答', '人事回答_Sheet1']);
    expect(env.st.state.toast?.text).toContain('他のファイルのシート 2 枚を追加');
    expect(env.st.state.toast?.detail).toContain('他のファイルのシート 2 枚を追加しました（職場回答_回答 ほか）');
    const counts = outputSheetCounts(m.book.sheets.length, m.report.addedSheets!.length);
    expect(outputSheetsLine(counts)).toBe('出力は 4 シート（結合 1 + 基準のほか 1 + 他ファイル 2）');
    const rows = outputSheetRows(m.book.sheets, m.layout, m.report);
    expect(rows.map((r) => r.kind)).toEqual(['merged', 'base', 'copy', 'copy']);
  });

  it('the preview list before the merge is the list the merge produces (card order, conflicts)', async () => {
    const env = await loaded(THREE);
    env.ctrl.setIncludeSourceSheets(true);
    const st0 = env.st.state;
    const planned = plannedOutputSheets(st0.files, effectiveBaseId(st0.files, st0.baseId), excludedSourceIds(st0.analysis, st0.plan), st0);
    await env.ctrl.merge();
    expect(sheetNames(env)).toEqual(planned.sheets.map((s) => s.name));
    expect(env.st.state.merged!.layout.addedSheets!.map((a) => a.name)).toEqual(planned.sheets.filter((s) => s.kind === 'copy').map((s) => s.name));
  });

  it('reordering the cards before the merge reorders the planned list AND the copies', async () => {
    const env = await loaded(THREE);
    env.ctrl.setIncludeSourceSheets(true);
    const before = () => {
      const s = env.st.state;
      return plannedOutputSheets(s.files, effectiveBaseId(s.files, s.baseId), excludedSourceIds(s.analysis, s.plan), s).sheets.map((x) => x.name);
    };
    expect(before()).toEqual(['応募者一覧', '集計', '職場回答_回答', '人事回答_Sheet1']);
    expect(env.ctrl.reorderFile('s3', 's2')).toBe(true);
    expect(before()).toEqual(['応募者一覧', '集計', '人事回答_Sheet1', '職場回答_回答']);
    await env.ctrl.ensureAnalysis();
    await env.ctrl.merge();
    expect(sheetNames(env)).toEqual(['応募者一覧', '集計', '人事回答_Sheet1', '職場回答_回答']);
  });

  it('two files with the same name: the later one gets (1) and the report says so', async () => {
    const env = await loaded(['学生回答.xlsx', '職場回答.xlsx', '職場回答.xlsx']);
    // (the second drop of the same name + size + date is refused as a duplicate: use another timestamp)
    if (env.st.state.files.length < 3) {
      await env.ctrl.addFiles([file('職場回答.xlsx', 1001, 1_700_000_999_999)]);
      await env.ctrl.idle();
      await env.ctrl.ensureAnalysis();
    }
    expect(env.st.state.files.length).toBe(3);
    env.ctrl.setIncludeSourceSheets(true);
    await env.ctrl.merge();
    expect(sheetNames(env)).toEqual(['応募者一覧', '集計', '職場回答_回答', '職場回答_回答(1)']);
    const added = env.st.state.merged!.layout.addedSheets!;
    expect(added.map((a) => a.suffixed)).toEqual([false, true]);
    expect(env.st.state.merged!.report.items.some((i) => i.code === 'sheetRenamed' && i.message.includes('職場回答_回答(1)'))).toBe(true);
  });

  it('formulas of a copied sheet read the copy (the sheet reference is renamed), not the merged sheet', async () => {
    // 学生回答 (with the formula sheet 集計) is the SECOND file here: its sheets are copied
    const env = await loaded(['職場回答.xlsx', '学生回答.xlsx']);
    env.ctrl.setIncludeSourceSheets(true);
    await env.ctrl.merge();
    const m = env.st.state.merged!;
    const names = m.book.sheets.map((s) => s.name);
    expect(names).toContain('学生回答_集計');
    const agg = m.book.sheets.find((s) => s.name === '学生回答_集計')!;
    const f = agg.rows.flatMap((r) => (r ?? []).map((c) => c?.f ?? '')).filter(Boolean);
    expect(f.length).toBeGreaterThan(0);
    for (const x of f) expect(x).toContain("'学生回答_応募者一覧'!");
    // the copy of the data sheet is a full original copy
    expect(names).toContain('学生回答_応募者一覧');
  });

  it('a copy that had to leave something out: the warning item is in the report and in the 警告 notes with the sheet name', async () => {
    const env = await loaded(['職場回答.xlsx', '学生回答.xlsx']);
    env.ctrl.setIncludeSourceSheets(true);
    await env.ctrl.merge();
    const m = env.st.state.merged!;
    const limits = m.report.items.filter((i) => i.code === 'sheetCopyLimits');
    expect(limits).toHaveLength(1);
    expect(limits[0].message).toContain('学生回答_集計');
    expect(m.layout.addedSheets!.find((a) => a.name === '学生回答_集計')!.complete).toBe(false);
    const s = env.st.state;
    const notes = fidelityNotes(s.files, effectiveBaseId(s.files, s.baseId), true, m.report, { merged: true, includeSourceSheets: true, dropOtherSheets: false });
    const n = notes.find((x) => x.key.startsWith('sheetcopy'))!;
    expect(n.severity).toBe('warning');
    expect(n.text).toContain('学生回答_集計');
    const rows = outputSheetRows(m.book.sheets, m.layout, m.report);
    expect(rows.find((r) => r.name === '学生回答_集計')).toMatchObject({ kind: 'copy', complete: false });
  });

  it('a file without a key takes no part: its sheets are not copied', async () => {
    const env = await loaded([...THREE, 'キーなし.xlsx']);
    env.ctrl.setIncludeSourceSheets(true);
    await env.ctrl.merge();
    expect(env.st.state.merged!.layout.addedSheets!.map((a) => a.sourceId)).toEqual(['s2', 's3']);
    expect(sheetNames(env)).not.toContain('キーなし_参加状況');
  });

  it('the 警告 notes of the engine-less run: asked for the copies but none were reported -> a warning (no false promise)', async () => {
    const real = createMockServices('?mock=1&delay=0.05');
    const env = await loaded(THREE, {
      buildPremerged: async (input, p) => {
        // an engine that ignores includeSourceSheets
        const { includeSourceSheets: _ignored, ...rest } = input.options;
        void _ignored;
        return real.buildPremerged({ ...input, options: rest }, p);
      },
    });
    env.ctrl.setIncludeSourceSheets(true);
    await env.ctrl.merge();
    const s = env.st.state;
    const m = s.merged!;
    expect(m.report.addedSheets).toBeUndefined();
    const notes = fidelityNotes(s.files, effectiveBaseId(s.files, s.baseId), true, m.report, { merged: true, includeSourceSheets: s.includeSourceSheets, dropOtherSheets: !s.includeOtherSheets });
    expect(notes.map((n) => n.key)).toContain('sheets-not-added');
    // and the counts line tells the truth about the workbook
    expect(outputSheetsLine(outputSheetCounts(m.book.sheets.length, 0))).toBe('出力は 2 シート（結合 1 + 基準のほか 1 + 他ファイル 0）');
  });
});
