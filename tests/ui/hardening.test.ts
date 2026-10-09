/**
 * Round 5 hardening (UI): friendly errors, folder drops, stale error cards, mislabelled files, lost server during export,
 * file-name rules. Pure logic + the controller against the ?mock=1 services.
 */
import { afterEach, describe, expect, it, vi } from 'vitest';
import { formatNote, createController } from '../../src/state/controller';
import { folderScanMessage, readDropPayload, scanFolders, type EntryLike } from '../../src/state/dropFiles';
import { errorLine, exportFileName, friendlyError, hasJapanese, headerRowIsNumeric, normalizeExportName, sanitizeFileName } from '../../src/state/helpers';
import { createMockServices } from '../../src/state/mock';
import type { Services } from '../../src/state/services';
import { editor, file, joined, store } from './fixtures';
import { effectiveImports, plannedColumns } from '../../src/state/keysPlan';
import { currentStep, initialState, reducer } from '../../src/state/reducer';

// ------------------------------------------------------------------------------------------------ fake file system

const fileEntry = (name: string, size = 100): EntryLike => ({
  isFile: true,
  isDirectory: false,
  name,
  file: (ok) => ok(file(name, size)),
});

/** a directory whose readEntries hands the children out in batches (like Chromium: a batch, then [] at the end) */
const dirEntry = (name: string, children: EntryLike[], batch = 2): EntryLike => ({
  isFile: false,
  isDirectory: true,
  name,
  createReader: () => {
    let i = 0;
    return {
      readEntries: (ok) => {
        const part = children.slice(i, i + batch);
        i += batch;
        queueMicrotask(() => ok(part));
      },
    };
  },
});

// ------------------------------------------------------------------------------------------------ errors

describe('friendlyError / errorLine', () => {
  it('keeps messages that are already Japanese', () => {
    expect(friendlyError(new Error('ディスクがいっぱいです'))).toEqual({ text: 'ディスクがいっぱいです' });
    expect(errorLine('保存できませんでした')).toBe('保存できませんでした');
  });

  it('turns English library errors into a Japanese sentence and keeps the original as detail', () => {
    const f = friendlyError(new TypeError("Cannot read properties of undefined (reading 'rows')"));
    expect(f.text).toBe('予期しないエラーが発生しました');
    expect(f.detail).toContain('Cannot read properties');
    expect(hasJapanese(f.text)).toBe(true);
    expect(errorLine(new Error('boom'))).toBe('予期しないエラーが発生しました［技術情報: boom］');
  });

  it('recognises out-of-memory and network failures', () => {
    expect(friendlyError(new RangeError('Invalid array length')).text).toContain('メモリが足りません');
    expect(friendlyError(new TypeError('Failed to fetch')).text).toContain('ローカルサーバーと通信できません');
  });

  it('never prints undefined / [object Object]', () => {
    for (const e of [undefined, null, {}, '', new Error(''), { message: undefined }, 42]) {
      const s = errorLine(e);
      expect(s).not.toMatch(/undefined|\[object|NaN|null/);
      expect(s.length).toBeGreaterThan(0);
    }
  });

  it('truncates very long technical notes', () => {
    const line = errorLine(new Error('x'.repeat(500)));
    expect(line.length).toBeLessThan(200);
    expect(line).toContain('…');
  });
});

// ------------------------------------------------------------------------------------------------ file names

describe('export file names', () => {
  it('reserved Windows device names get a leading underscore (same as the server)', () => {
    expect(normalizeExportName('CON', 'xlsx')).toBe('_CON.xlsx');
    expect(normalizeExportName('aux.xlsx', 'xlsx')).toBe('_aux.xlsx');
    expect(normalizeExportName('COM1', 'xlsm')).toBe('_COM1.xlsm');
    expect(normalizeExportName('console', 'xlsx')).toBe('console.xlsx');
  });

  it('dots-only / empty names fall back, long names are cut at 150 characters', () => {
    expect(sanitizeFileName('..')).toBe('統合結果');
    expect(sanitizeFileName('   ')).toBe('統合結果');
    expect(sanitizeFileName('x'.repeat(300))).toHaveLength(150);
    expect(normalizeExportName('x'.repeat(300), 'xlsx')).toHaveLength(150 + '.xlsx'.length);
  });

  it('forbidden characters become _, the default name keeps the date', () => {
    expect(normalizeExportName('a:b*c?', 'xlsx')).toBe('a_b_c_.xlsx');
    expect(normalizeExportName('sub/dir/x', 'xlsx')).toBe('sub_dir_x.xlsx');
    expect(exportFileName('名簿 (最新).xlsx', 'xlsx', new Date(2026, 9, 1))).toBe('名簿 (最新)_統合_20261001.xlsx');
  });
});

// ------------------------------------------------------------------------------------------------ dropped folders

describe('dropped folders', () => {
  it('separates folders from plain files (entries are read synchronously)', () => {
    const dir = dirEntry('フォルダ', []);
    const items = [
      { kind: 'file', getAsFile: () => file('a.xlsx'), webkitGetAsEntry: () => fileEntry('a.xlsx') },
      { kind: 'file', getAsFile: () => file('フォルダ', 0), webkitGetAsEntry: () => dir },
      { kind: 'string', getAsFile: () => null },
    ];
    const p = readDropPayload({ items, files: [file('a.xlsx'), file('フォルダ', 0)] });
    expect(p.files.map((f) => f.name)).toEqual(['a.xlsx']);
    expect(p.folders).toEqual([dir]);
  });

  it('falls back to the plain files list without items (older browsers)', () => {
    const p = readDropPayload({ files: [file('a.xlsx'), file('b.csv')], items: null });
    expect(p.files).toHaveLength(2);
    expect(p.folders).toEqual([]);
    expect(readDropPayload(null)).toEqual({ files: [], folders: [] });
  });

  it('collects the Excel files of a folder tree in natural order, in batches, ignoring the rest', async () => {
    const tree = dirEntry('root', [
      fileEntry('データ10.xlsx'),
      fileEntry('データ2.xlsx'),
      fileEntry('メモ.txt'),
      fileEntry('~$データ2.xlsx'),
      fileEntry('.hidden.xlsx'),
      fileEntry('Thumbs.db'),
      dirEntry('sub', [fileEntry('b.csv'), fileEntry('c.png'), dirEntry('deeper', [fileEntry('d.xls')])]),
      dirEntry('.git', [fileEntry('x.xlsx')]),
    ]);
    const r = await scanFolders([tree]);
    // one folder level after the other, each in natural order (データ2 before データ10); 'sub' sorts before 'データ…'
    expect(r.files.map((f) => f.name)).toEqual(['b.csv', 'd.xls', 'データ2.xlsx', 'データ10.xlsx']);
    expect(r.found).toBe(4);
    expect(r.ignored).toBe(2); // メモ.txt, c.png
    expect(r.truncated).toBe(false);
    expect(r.unreadable).toBe(0);
  });

  it('stops at the file limit and at the depth limit and says so', async () => {
    const many = dirEntry(
      'many',
      Array.from({ length: 40 }, (_, i) => fileEntry(`f${String(i).padStart(2, '0')}.xlsx`)),
      7,
    );
    const r = await scanFolders([many], { maxFiles: 30 });
    expect(r.files).toHaveLength(30);
    expect(r.found).toBe(40);
    expect(r.truncated).toBe(true);
    expect(folderScanMessage(['many'], r)).toMatchObject({ kind: 'error' });
    expect(folderScanMessage(['many'], r)!.text).toContain('多すぎる');

    const deep = dirEntry('l0', [dirEntry('l1', [dirEntry('l2', [fileEntry('x.xlsx')])])]);
    const d = await scanFolders([deep], { maxDepth: 1 });
    expect(d.tooDeep).toBe(true);
    expect(d.files).toHaveLength(0);
  });

  it('gives up on a folder with a huge number of entries instead of walking all of it', async () => {
    const huge = dirEntry(
      'Documents',
      Array.from({ length: 400 }, (_, i) => (i % 2 ? fileEntry(`f${i}.txt`) : dirEntry(`d${i}`, [fileEntry(`x${i}.xlsx`)]))),
      50,
    );
    const r = await scanFolders([huge], { maxEntries: 100 });
    expect(r.stopped).toBe(true);
    expect(r.files.length).toBeLessThan(60);
    const m = folderScanMessage(['Documents'], r)!;
    expect(m.kind).toBe('error');
    expect(m.text).toContain('途中までしか調べていません');
    const none = await scanFolders([dirEntry('x', Array.from({ length: 20 }, (_, i) => fileEntry(`n${i}.txt`)))], { maxEntries: 5 });
    expect(none.stopped).toBe(true);
    expect(folderScanMessage(['x'], none)!.text).toContain('途中までしか調べていません');
  });

  it('counts unreadable folders / files instead of failing', async () => {
    const broken: EntryLike = { isFile: false, isDirectory: true, name: 'locked', createReader: () => ({ readEntries: (_ok, fail) => fail?.(new Error('denied')) }) };
    const badFile: EntryLike = { isFile: true, isDirectory: false, name: 'bad.xlsx', file: (_ok, fail) => fail?.(new Error('denied')) };
    const r = await scanFolders([broken, dirEntry('ok', [badFile, fileEntry('good.xlsx')])]);
    expect(r.files.map((f) => f.name)).toEqual(['good.xlsx']);
    expect(r.unreadable).toBe(2);
  });

  it('messages: nothing found / found with remarks', async () => {
    const empty = await scanFolders([dirEntry('空', [fileEntry('memo.txt')])]);
    const m = folderScanMessage(['空'], empty)!;
    expect(m.kind).toBe('error');
    expect(m.text).toContain('フォルダ「空」');
    expect(m.text).toContain('見つかりませんでした');
    const some = await scanFolders([dirEntry('データ', [fileEntry('a.xlsx'), fileEntry('b.xlsx'), fileEntry('c.txt')])]);
    const ok = folderScanMessage(['データ'], some)!;
    expect(ok.kind).toBe('info');
    expect(ok.text).toContain('2 個のファイルを読み込みます');
    expect(ok.text).toContain('Excel以外の 1 個は無視');
    expect(folderScanMessage(['a', 'b'], empty)!.text).toContain('2 個のフォルダ');
  });
});

// ------------------------------------------------------------------------------------------------ controller

function setup(search = '?mock=1', patch: Partial<Services> = {}) {
  const st = store();
  const services = { ...createMockServices(`${search}&delay=0.05`), ...patch };
  const ctrl = createController({ services, dispatch: st.dispatch, getState: st.getState });
  return { st, ctrl, services };
}

describe('controller: drops', () => {
  it('a dropped folder adds its Excel files (alphabetical) and explains what it ignored', async () => {
    const { st, ctrl } = setup();
    const folder = dirEntry('回答', [fileEntry('職場回答.xlsx'), fileEntry('学生回答.xlsx'), fileEntry('メモ.txt')]);
    await ctrl.addDropped({ files: [], folders: [folder] });
    await ctrl.idle();
    expect(st.state.files.map((f) => f.fileName)).toEqual(['学生回答.xlsx', '職場回答.xlsx'].sort((a, b) => a.localeCompare(b, 'ja', { numeric: true, sensitivity: 'base' })));
    expect(st.state.files.every((f) => f.status === 'ready')).toBe(true);
    expect(st.state.busy).toBeNull();
    const toasts = st.actions.filter((a) => a.type === 'toast/show') as Array<{ text: string }>;
    expect(toasts.map((t) => t.text).join('\n')).toContain('フォルダ「回答」から 2 個のファイルを読み込みます');
  });

  it('an empty folder only produces a message', async () => {
    const { st, ctrl } = setup();
    await ctrl.addDropped({ files: [], folders: [dirEntry('空のフォルダ', [])] });
    expect(st.state.files).toHaveLength(0);
    expect(st.state.toast).toMatchObject({ kind: 'error' });
    expect(st.state.toast?.text).toContain('見つかりませんでした');
    expect(st.state.busy).toBeNull();
  });

  it('dropping only an Excel lock file explains what it is', async () => {
    const { st, ctrl } = setup();
    await ctrl.addFiles([file('~$名簿.xlsx')]);
    expect(st.state.files).toHaveLength(0);
    expect(st.state.toast).toMatchObject({ kind: 'error' });
    expect(st.state.toast?.text).toContain('一時ファイル');
  });

  it('dropping the same file again replaces its error card instead of piling up', async () => {
    const { st, ctrl } = setup();
    await ctrl.addFiles([file('壊れた.xlsx')]);
    await ctrl.idle();
    expect(st.state.files.map((f) => f.status)).toEqual(['error']);
    await ctrl.addFiles([file('壊れた.xlsx', 2000, 1_700_000_999_000)]);
    await ctrl.idle();
    expect(st.state.files).toHaveLength(1);
    expect(st.state.files[0].id).not.toBe('s1');
  });

  it('a file the browser cannot read (locked by Excel, online-only OneDrive file) explains what to do', async () => {
    const { st, ctrl } = setup();
    const unreadable = Object.assign(file('開けない.xlsx'), { arrayBuffer: () => Promise.reject(new DOMException('The requested file could not be read', 'NotReadableError')) });
    await ctrl.addFiles([unreadable]);
    await ctrl.idle();
    expect(st.state.files[0].status).toBe('error');
    expect(st.state.files[0].error).toContain('閉じてから、もう一度選んでください');
    expect(st.state.files[0].error).toContain('オンライン専用');
    expect(st.state.files[0].error).not.toMatch(/NotReadable|could not be read/);
  });

  it('mixed good and bad files: the good ones are usable, every bad one has its own Japanese card text', async () => {
    const { st, ctrl } = setup();
    await ctrl.addFiles([file('学生回答.xlsx'), file('壊れた.xlsx'), file('パスワード付き.xlsx'), file('空.xlsx', 0), file('職場回答.xlsx')]);
    await ctrl.idle();
    await ctrl.ensureAnalysis();
    expect(st.state.files.map((f) => f.status)).toEqual(['ready', 'error', 'error', 'error', 'ready']);
    for (const f of st.state.files.filter((x) => x.status === 'error')) {
      expect(f.error).toMatch(/[぀-ヿ一-鿿]/);
      expect(f.error).not.toMatch(/undefined|\[object/);
    }
    expect(st.state.analysis).not.toBeNull();
  });
});

describe('formatNote', () => {
  it('says so when an Excel extension holds plain text', () => {
    expect(formatNote('名簿.xlsx', { format: 'csv' })).toContain('中身はExcelではなくテキスト');
    expect(formatNote('名簿.xls', { format: 'csv' })).toContain('.xls');
    expect(formatNote('名簿.csv', { format: 'csv' })).toBeUndefined();
    expect(formatNote('名簿.xlsx', { format: 'xlsx' })).toBeUndefined();
  });
});

describe('controller: export while the local server is gone', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('uses the built-in engine and the browser download even though the server answered before', async () => {
    const downloads: string[] = [];
    vi.stubGlobal('document', {
      createElement: () => ({ style: {}, click() {}, remove() {}, set download(v: string) { downloads.push(v); } }),
      body: { appendChild() {} },
    });
    const { st, ctrl, services } = setup('?mock=1');
    let excelCalls = 0;
    let saveCalls = 0;
    services.runExcelJob = (async () => {
      excelCalls++;
      return { bytes: new Uint8Array([1]), warnings: [] };
    }) as Services['runExcelJob'];
    services.saveToOutput = (async () => {
      saveCalls++;
      return { path: 'x', name: 'x.xlsx' };
    }) as Services['saveToOutput'];
    ctrl.init();
    await ctrl.addFiles([file('学生回答.xlsx'), file('職場回答.xlsx'), file('人事回答.xlsx')]);
    await ctrl.idle();
    await ctrl.ensureAnalysis();
    await ctrl.merge();
    expect(st.state.phase).toBe('merged');
    expect(st.state.server.status).not.toBeNull();
    st.dispatch({ type: 'server/connection', connected: false });
    await ctrl.runExport({ fileName: '結果', engine: 'excel' });
    expect(excelCalls).toBe(0);
    expect(saveCalls).toBe(0);
    expect(st.state.exportState).toMatchObject({ phase: 'done', engine: 'js', delivered: 'download', fileName: '結果.xlsx' });
    expect(downloads).toEqual(['結果.xlsx']);
  });
});

describe('reducer: toasts around edits', () => {
  it('a successful edit clears the error toast of an earlier refused one, but keeps an info toast', async () => {
    const { snapshot, mergedBook } = await editor();
    const base = reducer(
      reducer(initialState, {
        type: 'merge/done',
        bytes: new Uint8Array(1),
        layout: {} as never,
        report: { importedColumns: 0, appendedRows: 0, formulasToValues: 0, items: [] } as never,
        book: mergedBook,
        snapshot,
        baseFileName: 'a.xlsx',
        format: 'xlsx',
      }),
      { type: 'toast/show', text: '数式のかっこ「(」が閉じられていません', kind: 'error' },
    );
    const next = { ...snapshot, doc: { ...snapshot.doc, version: snapshot.doc.version + 1 }, log: [], label: '列を削除' };
    expect(reducer(base, { type: 'doc/commit', snapshot: next }).toast).toBeNull();
    const info = reducer(base, { type: 'toast/show', text: '結合しました', kind: 'info' });
    expect(reducer(info, { type: 'doc/commit', snapshot: next }).toast?.text).toBe('結合しました');
  });

  it('a finished export goes back to idle when the document changes (commit / undo / redo), a running one does not', async () => {
    const { snapshot, mergedBook } = await editor();
    const merged = reducer(initialState, {
      type: 'merge/done',
      bytes: new Uint8Array(1),
      layout: {} as never,
      report: { importedColumns: 0, appendedRows: 0, formulasToValues: 0, items: [] } as never,
      book: mergedBook,
      snapshot,
      baseFileName: 'a.xlsx',
      format: 'xlsx',
    });
    const next = { ...snapshot, doc: { ...snapshot.doc, version: snapshot.doc.version + 1 }, log: [], label: '列を削除' };
    const done = reducer(merged, { type: 'export/update', patch: { phase: 'done', engine: 'js', message: '' } });
    expect(currentStep(done)).toBe(3);
    const edited = reducer(done, { type: 'doc/commit', snapshot: next });
    expect(edited.exportState.phase).toBe('idle');
    expect(currentStep(edited)).toBe(2);
    expect(reducer(reducer(done, { type: 'doc/commit', snapshot: next }), { type: 'export/update', patch: { phase: 'done' } }).exportState.phase).toBe('done');
    expect(reducer(reducer(edited, { type: 'export/update', patch: { phase: 'done' } }), { type: 'doc/undo' }).exportState.phase).toBe('idle');
    const running = reducer(merged, { type: 'export/update', patch: { phase: 'running', engine: 'js' } });
    expect(reducer(running, { type: 'doc/commit', snapshot: next }).exportState.phase).toBe('running');
    const failed = reducer(merged, { type: 'export/update', patch: { phase: 'error', error: '出力できませんでした' } });
    expect(reducer(failed, { type: 'doc/commit', snapshot: next }).exportState.phase).toBe('idle');
  });
});

describe('headerRowIsNumeric', () => {
  const sheet = (rows: Array<Array<{ t: string; v: unknown }>>) => ({ sheets: [{ name: 'S', rows, maxRow: rows.length, maxCol: 4 }] }) as never;
  const region = { sheet: 'S', headerRow: 1, firstDataRow: 2, lastDataRow: 3, firstCol: 1, lastCol: 4 };
  it('is true for a header row of numbers only (a table without headers)', () => {
    const b = sheet([[{ t: 'n', v: 1 }, { t: 'n', v: 37 }, { t: 'n', v: 11 }, { t: 'n', v: 5 }], [{ t: 'n', v: 2 }]]);
    expect(headerRowIsNumeric(b, region)).toBe(true);
  });
  it('is false with any text in the header row, for a single cell, or without a region', () => {
    expect(headerRowIsNumeric(sheet([[{ t: 'n', v: 1 }, { t: 's', v: '氏名' }]]), region)).toBe(false);
    expect(headerRowIsNumeric(sheet([[{ t: 'n', v: 1 }, { t: 'z', v: null }, { t: 'z', v: null }]]), region)).toBe(false);
    expect(headerRowIsNumeric(sheet([[{ t: 'n', v: 1 }]]), undefined)).toBe(false);
  });
});

describe('effectiveImports: header overrides of the keys module (two-row headers)', () => {
  it('a header the keys module gives a column (面接 1次) travels with the import to the premerge', async () => {
    const j = await joined();
    const first = j.join.defaultImports.find((d) => d.sourceId === 's2')!;
    const join = { ...j.join, defaultImports: j.join.defaultImports.map((d) => (d === first ? { ...d, header: '面接 1次' } : d)) };
    const eff = effectiveImports(j.analysis, j.plan, join, {});
    expect(eff.find((i) => i.sourceId === 's2' && i.col === first.col)).toEqual({ sourceId: 's2', col: first.col, header: '面接 1次' });
    // the other columns stay plain specs (no header key at all)
    const others = eff.filter((i) => !(i.sourceId === 's2' && i.col === first.col));
    expect(others.every((i) => !('header' in i))).toBe(true);
    // a column switched off is not imported, so its header does not travel either
    const sheet = j.analysis.pairs.s2.other.sheet;
    const off = effectiveImports(j.analysis, j.plan, join, { [`s2|${sheet}|${first.col}`]: false });
    expect(off.some((i) => i.sourceId === 's2' && i.col === first.col)).toBe(false);
    expect(plannedColumns(j.analysis, j.plan, join, {}).filter((c) => c.imported).length).toBe(eff.length);
  });
});
