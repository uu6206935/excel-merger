import { beforeAll, describe, expect, it } from 'vitest';
import type { EditorSnapshot, SourceBook, TableRegion } from '../../src/core/types';
import { analysisInputKey, filesKey, othersOrderKey } from '../../src/state/keysPlan';
import { clampActive, currentStep, hasEdits, initialState, reducer, type Action } from '../../src/state/reducer';
import type { AppState } from '../../src/state/types';
import { SOURCE_COLORS } from '../../src/state/helpers';
import { book, editor, joined, type Joined } from './fixtures';

const run = (actions: Action[], start: AppState = initialState) => actions.reduce(reducer, start);

let b1: SourceBook;
let b2: SourceBook;
let j: Joined;
let snap: EditorSnapshot;

beforeAll(async () => {
  b1 = await book('学生回答.xlsx', 's1');
  b2 = await book('職場回答.xlsx', 's2');
  j = await joined();
  snap = (await editor()).snapshot;
});

const region = (b: SourceBook): TableRegion => ({ sheet: b.sheets[0].name, headerRow: 1, firstDataRow: 2, lastDataRow: 10, firstCol: 1, lastCol: 5 });

function twoFiles(): AppState {
  return run([
    { type: 'files/add', entries: [{ id: 's1', fileName: '学生回答.xlsx', size: 10 }, { id: 's2', fileName: '職場回答.xlsx', size: 20 }] },
    { type: 'file/loaded', id: 's1', book: b1, sheet: b1.sheets[0].name, region: region(b1) },
    { type: 'file/loaded', id: 's2', book: b2, sheet: b2.sheets[0].name, region: region(b2) },
  ]);
}

describe('files', () => {
  it('adds files with distinct colours, the first one is the base', () => {
    const s = run([{ type: 'files/add', entries: [{ id: 's1', fileName: 'a.xlsx', size: 1 }, { id: 's2', fileName: 'b.xlsx', size: 1 }] }]);
    expect(s.files.map((f) => f.status)).toEqual(['loading', 'loading']);
    expect(s.files.map((f) => f.color)).toEqual([SOURCE_COLORS[0], SOURCE_COLORS[1]]);
    expect(s.baseId).toBe('s1');
    expect(s.seq).toBe(3);
  });

  it('marks a file ready / error and keeps the other files', () => {
    let s = run([{ type: 'files/add', entries: [{ id: 's1', fileName: 'a.xlsx', size: 1 }, { id: 's2', fileName: 'b.xlsx', size: 1 }] }]);
    s = reducer(s, { type: 'file/loaded', id: 's1', book: b1, sheet: 'x', region: region(b1) });
    s = reducer(s, { type: 'file/error', id: 's2', error: '壊れています' });
    expect(s.files[0]).toMatchObject({ status: 'ready', sheet: 'x', book: b1 });
    expect(s.files[0].detected).toEqual(s.files[0].region);
    expect(s.files[1]).toMatchObject({ status: 'error', error: '壊れています' });
    expect(s.files[1].book).toBeUndefined();
  });

  it('ignores actions for unknown files (no new state object)', () => {
    const s = twoFiles();
    expect(reducer(s, { type: 'file/progress', id: 'zz', message: 'x' })).toBe(s);
  });

  it('removing the base picks the next ready file; fewer than 2 files clear the analysis', () => {
    let s = twoFiles();
    s = reducer(s, { type: 'analysis/done', key: analysisInputKey(s.files, s.baseId, s.options), analysis: j.analysis, plan: j.plan, joinResult: j.join });
    expect(s.analysis).not.toBeNull();
    s = reducer(s, { type: 'file/remove', id: 's1' });
    expect(s.baseId).toBe('s2');
    expect(s.analysis).toBeNull();
    expect(s.plan).toBeNull();
    expect(s.joinResult).toBeNull();
  });

  it('a user header row change updates the region and clears the card error', () => {
    let s = twoFiles();
    s = reducer(s, { type: 'file/region', id: 's1', region: { ...region(b1), headerRow: 2, firstDataRow: 3 } });
    expect(s.files[0].region?.headerRow).toBe(2);
    expect(s.files[0].detected?.headerRow).toBe(1);
    s = reducer(s, { type: 'file/region', id: 's1', region: { ...region(b1), headerRow: 9, firstDataRow: 10, lastDataRow: 9 }, error: 'データなし' });
    expect(s.files[0].error).toBe('データなし');
  });
});

describe('analysis', () => {
  it('accepts results only for the current inputs', () => {
    const s = twoFiles();
    const key = analysisInputKey(s.files, s.baseId, s.options);
    const stale = reducer(s, { type: 'analysis/done', key: 'old', analysis: j.analysis, plan: j.plan, joinResult: j.join });
    expect(stale).toBe(s);
    const ok = reducer(s, { type: 'analysis/done', key, analysis: j.analysis, plan: j.plan, joinResult: j.join });
    expect(ok.analysis).toBe(j.analysis);
    expect(ok.analysisKey).toBe(key);
    expect(ok.planFilesKey).not.toBeNull();
  });

  it('changing the base, options or a region invalidates the analysis; the base change also drops the plan', () => {
    let s = twoFiles();
    s = reducer(s, { type: 'analysis/done', key: analysisInputKey(s.files, s.baseId, s.options), analysis: j.analysis, plan: j.plan, joinResult: j.join });
    const opt = reducer(s, { type: 'options/set', options: { itaiji: false } });
    expect(opt.analysis).toBeNull();
    expect(opt.plan).toBe(j.plan); // kept for carry-over
    expect(opt.options.itaiji).toBe(false);
    const base = reducer(s, { type: 'base/set', id: 's2' });
    expect(base.baseId).toBe('s2');
    expect(base.analysis).toBeNull();
    expect(base.plan).toBeNull();
    const same = reducer(s, { type: 'base/set', id: 's1' });
    expect(same).toBe(s);
    const reg = reducer(s, { type: 'file/region', id: 's2', region: { ...region(b2), lastDataRow: 5 } });
    expect(reg.analysis).toBeNull();
  });

  it('progress messages do not invalidate the analysis', () => {
    let s = twoFiles();
    s = reducer(s, { type: 'analysis/done', key: analysisInputKey(s.files, s.baseId, s.options), analysis: j.analysis, plan: j.plan, joinResult: j.join });
    s = reducer(s, { type: 'files/add', entries: [{ id: 's3', fileName: 'c.xlsx', size: 1 }] });
    // a loading file is not part of the inputs yet
    expect(s.analysis).toBe(j.analysis);
    expect(reducer(s, { type: 'file/progress', id: 's3', message: '50%' }).analysis).toBe(j.analysis);
  });

  it('plan edits replace plan + join result; import overrides merge', () => {
    let s = twoFiles();
    s = reducer(s, { type: 'plan/set', plan: j.plan, joinResult: j.join });
    expect(s.plan).toBeNull(); // no analysis yet
    s = reducer(s, { type: 'analysis/done', key: analysisInputKey(s.files, s.baseId, s.options), analysis: j.analysis, plan: j.plan, joinResult: j.join });
    const plan2 = { ...j.plan, joinType: 'full' as const };
    s = reducer(s, { type: 'plan/set', plan: plan2, joinResult: null, error: 'x' });
    expect(s.plan).toBe(plan2);
    expect(s.keyError).toBe('x');
    s = reducer(s, { type: 'imports/set', changes: { a: true } });
    s = reducer(s, { type: 'imports/set', changes: { b: false } });
    expect(s.importOverrides).toEqual({ a: true, b: false });
  });

  it('manual keys come with the analysis, vanish with a new base or fewer than two files, and default to none', () => {
    let s = twoFiles();
    const key = analysisInputKey(s.files, s.baseId, s.options);
    expect(s.manualKeys).toEqual({});
    const manualKeys = { s2: [{ id: 'manual:1|1', baseCols: [1], otherCols: [1] }] };
    s = reducer(s, { type: 'analysis/done', key, analysis: j.analysis, plan: j.plan, joinResult: j.join, manualKeys });
    expect(s.manualKeys).toBe(manualKeys);
    // options / region changes keep them (the controller re-applies them after the re-analysis)
    expect(reducer(s, { type: 'options/set', options: { itaiji: false } }).manualKeys).toBe(manualKeys);
    // a stale analysis result does not touch them
    expect(reducer(s, { type: 'analysis/done', key: 'old', analysis: j.analysis, plan: j.plan, joinResult: j.join, manualKeys: {} }).manualKeys).toBe(manualKeys);
    // another analysis result without manual keys (older callers) resets them
    expect(reducer(s, { type: 'analysis/done', key, analysis: j.analysis, plan: j.plan, joinResult: j.join }).manualKeys).toEqual({});
    expect(reducer(s, { type: 'base/set', id: 's2' }).manualKeys).toEqual({});
    expect(reducer(s, { type: 'file/remove', id: 's2' }).manualKeys).toEqual({});
  });
});

function mergedState(): AppState {
  let s = twoFiles();
  s = reducer(s, { type: 'analysis/done', key: analysisInputKey(s.files, s.baseId, s.options), analysis: j.analysis, plan: j.plan, joinResult: j.join });
  return reducer(s, {
    type: 'merge/done',
    bytes: new Uint8Array([1]),
    layout: { sheet: '応募者一覧', insertAt: 12, imported: [], appendAt: null, appended: [] },
    report: { mergedRows: 30, matchedRows: 27, baseOnlyRows: 3, appendedRows: 0, importedColumns: 7, formulasKept: 0, formulasToValues: 0, items: [] },
    book: b1,
    snapshot: snap,
    baseFileName: '学生回答.xlsx',
    format: 'xlsx',
  });
}

describe('merge + history', () => {
  it('merge/done enters the editor with a fresh history', () => {
    const s = mergedState();
    expect(s.phase).toBe('merged');
    expect(s.merged?.history.present).toBe(snap);
    expect(s.active).toEqual({ row: snap.doc.firstDataRow, col: snap.doc.columns[0].id });
    expect(Object.keys(s.merged!.initialWidths)).toHaveLength(snap.doc.columns.length);
    expect(hasEdits(s)).toBe(false);
    expect(currentStep(s)).toBe(2);
  });

  it('commit / undo / redo', () => {
    let s = mergedState();
    const next: EditorSnapshot = { ...snap, doc: { ...snap.doc, columns: snap.doc.columns.slice(1), version: snap.doc.version + 1 }, log: [], label: '列を削除' };
    const gone = snap.doc.columns[0].id;
    s = reducer(s, { type: 'grid/selection', selection: { cols: new Set([gone, snap.doc.columns[1].id]), anchor: gone } });
    s = reducer(s, { type: 'grid/active', active: { row: 5, col: gone } });
    s = reducer(s, { type: 'doc/commit', snapshot: next });
    expect(s.merged?.history.past).toHaveLength(1);
    expect(hasEdits(s)).toBe(true);
    // the deleted column leaves the selection and the active cell moves to an existing column
    expect([...s.selection.cols]).toEqual([snap.doc.columns[1].id]);
    expect(s.selection.anchor).toBeNull();
    expect(s.active).toEqual({ row: 5, col: next.doc.columns[0].id });
    s = reducer(s, { type: 'doc/undo' });
    expect(s.merged?.history.present).toBe(snap);
    expect(reducer(s, { type: 'doc/undo' })).toBe(s); // nothing more to undo
    s = reducer(s, { type: 'doc/redo' });
    expect(s.merged?.history.present).toBe(next);
    expect(reducer(s, { type: 'doc/redo' })).toBe(s);
  });

  it('live widths are cleared by a commit', () => {
    let s = mergedState();
    s = reducer(s, { type: 'grid/liveWidth', id: 'b1', px: 120 });
    expect(s.liveWidths).toEqual({ b1: 120 });
    s = reducer(s, { type: 'grid/liveWidth', id: 'b1', px: null });
    expect(s.liveWidths).toEqual({});
  });

  it('merge/reset returns to the key screen keeping the analysis', () => {
    const s = reducer(mergedState(), { type: 'merge/reset' });
    expect(s.phase).toBe('files');
    expect(s.merged).toBeNull();
    expect(s.analysis).toBe(j.analysis);
    expect(currentStep(s)).toBe(1);
  });

  it('doc actions without a merge are ignored', () => {
    expect(reducer(initialState, { type: 'doc/undo' })).toBe(initialState);
    expect(reducer(initialState, { type: 'doc/commit', snapshot: snap })).toBe(initialState);
  });

  it('clampActive keeps the active cell inside the data rows', () => {
    const first = snap.doc.columns[0].id;
    expect(clampActive({ row: 1, col: first }, snap)).toEqual({ row: snap.doc.firstDataRow, col: first });
    expect(clampActive({ row: 99999, col: 'nope' }, snap)).toEqual({ row: snap.doc.lastDataRow, col: first });
    expect(clampActive(null, snap)).toBeNull();
  });
});

describe('ui chrome', () => {
  it('toasts get increasing ids; clearing an old id keeps the new toast', () => {
    let s = reducer(initialState, { type: 'toast/show', text: 'a', kind: 'info' });
    const first = s.toast!.id;
    s = reducer(s, { type: 'toast/show', text: 'b', kind: 'error' });
    expect(s.toast).toMatchObject({ text: 'b', kind: 'error' });
    expect(reducer(s, { type: 'toast/clear', id: first }).toast?.text).toBe('b');
    expect(reducer(s, { type: 'toast/clear' }).toast).toBeNull();
  });

  it('step indicator', () => {
    expect(currentStep(initialState)).toBe(0);
    expect(currentStep(twoFiles())).toBe(1);
    const m = mergedState();
    expect(currentStep(reducer(m, { type: 'modal/open', modal: { kind: 'export' } }))).toBe(3);
    expect(currentStep(reducer(m, { type: 'export/update', patch: { phase: 'done' } }))).toBe(3);
  });

  it('export state patches and resets', () => {
    let s = reducer(initialState, { type: 'export/update', patch: { phase: 'running', message: 'x', ratio: 0.5 } });
    expect(s.exportState).toMatchObject({ phase: 'running', ratio: 0.5 });
    s = reducer(s, { type: 'export/reset' });
    expect(s.exportState.phase).toBe('idle');
  });

  it('server status', () => {
    const s = reducer(initialState, { type: 'server/status', status: null });
    expect(s.server).toEqual({ checked: true, status: null, connected: true });
    const off = reducer(s, { type: 'server/connection', connected: false });
    expect(off.server.connected).toBe(false);
    expect(reducer(off, { type: 'server/connection', connected: false })).toBe(off);
  });
});

// ---------------------------------------------------------------------------------------------
// round 7: file card reorder, 「他のシートも出力に含める」

function threeFiles(): AppState {
  const b3 = b2; // any loaded book will do: the reducer does not look inside
  let s = twoFiles();
  s = reducer(s, { type: 'files/add', entries: [{ id: 's3', fileName: '人事回答.xlsx', size: 30 }] });
  s = reducer(s, { type: 'file/loaded', id: 's3', book: b3, sheet: b3.sheets[0].name, region: region(b3) });
  return s;
}

describe('file card reorder (files/move)', () => {
  const ids = (s: AppState) => s.files.map((f) => f.id);

  it('moves a card before another / to the end and keeps everything the user chose', () => {
    let s = threeFiles();
    s = reducer(s, { type: 'analysis/done', key: analysisInputKey(s.files, s.baseId, s.options), analysis: j.analysis, plan: j.plan, joinResult: j.join, manualKeys: { s2: [{ id: 'manual:1|1', baseCols: [1], otherCols: [1] }] } });
    s = reducer(s, { type: 'imports/set', changes: { 'k|x|1': false } });
    const plan = s.plan;
    const planFilesKey = s.planFilesKey;
    const manualKeys = s.manualKeys;

    const m = reducer(s, { type: 'files/move', id: 's3', beforeId: 's2' });
    expect(ids(m)).toEqual(['s1', 's3', 's2']);
    // a PURE reorder: the key plan, the manual keys, the region choices, the import choices and the base stay
    expect(m.plan).toBe(plan);
    expect(m.planFilesKey).toBe(planFilesKey);
    expect(m.manualKeys).toBe(manualKeys);
    expect(m.importOverrides).toEqual({ 'k|x|1': false });
    expect(m.baseId).toBe('s1');
    for (const f of s.files) expect(m.files.find((x) => x.id === f.id)).toBe(f); // the very same file entries (sheet, region, error …)
    // ... but what depends on the ORDER is recomputed: the analysis / join result (pairs, import order, preview rows, appended rows)
    expect(m.analysis).toBeNull();
    expect(m.joinResult).toBeNull();
    expect(m.analysisKey).toBeNull();

    expect(ids(reducer(s, { type: 'files/move', id: 's2', beforeId: null }))).toEqual(['s1', 's3', 's2']);
  });

  it('the plan key ignores the order of the files; the analysis key does not (order of the OTHER files)', () => {
    const s = threeFiles();
    const m = reducer(s, { type: 'files/move', id: 's3', beforeId: 's2' });
    expect(filesKey(m.files, m.baseId)).toBe(filesKey(s.files, s.baseId));
    expect(analysisInputKey(m.files, m.baseId, m.options)).not.toBe(analysisInputKey(s.files, s.baseId, s.options));
    expect(othersOrderKey(s.files, 's1')).toBe('s2,s3');
    expect(othersOrderKey(m.files, 's1')).toBe('s3,s2');
  });

  it('moving the BASE card among the others does not change the order of the imports: the analysis stays', () => {
    let s = threeFiles();
    s = reducer(s, { type: 'analysis/done', key: analysisInputKey(s.files, s.baseId, s.options), analysis: j.analysis, plan: j.plan, joinResult: j.join });
    const m = reducer(s, { type: 'files/move', id: 's1', beforeId: null });
    expect(ids(m)).toEqual(['s2', 's3', 's1']);
    expect(m.baseId).toBe('s1');
    expect(m.analysis).toBe(s.analysis);
    expect(m.joinResult).toBe(s.joinResult);
  });

  it('a no-op (itself, where it is, unknown id) returns the same state object', () => {
    const s = threeFiles();
    expect(reducer(s, { type: 'files/move', id: 's2', beforeId: 's2' })).toBe(s);
    expect(reducer(s, { type: 'files/move', id: 's2', beforeId: 's3' })).toBe(s);
    expect(reducer(s, { type: 'files/move', id: 's3', beforeId: null })).toBe(s);
    expect(reducer(s, { type: 'files/move', id: 'zz', beforeId: 's1' })).toBe(s);
  });

  it('files that are still loading / failed can be moved without invalidating the analysis (they are not part of the key)', () => {
    let s = twoFiles();
    s = reducer(s, { type: 'files/add', entries: [{ id: 's3', fileName: '壊れた.xlsx', size: 1 }] });
    s = reducer(s, { type: 'file/error', id: 's3', error: '壊れています' });
    s = reducer(s, { type: 'analysis/done', key: analysisInputKey(s.files, s.baseId, s.options), analysis: j.analysis, plan: j.plan, joinResult: j.join });
    const m = reducer(s, { type: 'files/move', id: 's3', beforeId: 's1' });
    expect(ids(m)).toEqual(['s3', 's1', 's2']);
    expect(m.analysis).toBe(s.analysis);
  });

  it('is ignored after the merge (the cards then move the column groups of the grid)', () => {
    const s = mergedState();
    expect(reducer(s, { type: 'files/move', id: 's2', beforeId: null })).toBe(s);
  });
});

describe('output sheets option', () => {
  it('defaults to the single-sheet output and is a merge setting: changeable before the merge only', () => {
    expect(initialState.includeOtherSheets).toBe(false);
    const s = reducer(initialState, { type: 'output/otherSheets', include: true });
    expect(s.includeOtherSheets).toBe(true);
    expect(reducer(s, { type: 'output/otherSheets', include: true })).toBe(s);
    const m = mergedState();
    expect(reducer(m, { type: 'output/otherSheets', include: true })).toBe(m);
    // the choice survives 結合をやり直す
    expect(reducer(reducer(mergedState(), { type: 'merge/reset' }), { type: 'output/otherSheets', include: true }).includeOtherSheets).toBe(true);
    expect(reducer(s, { type: 'merge/reset' }).includeOtherSheets).toBe(true);
  });
});
