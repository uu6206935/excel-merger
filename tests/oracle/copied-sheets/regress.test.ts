/**
 * Regression tests for the defects the copied-sheets oracle (task V9) found. Each test names the oracle class it came from; the real-Excel facts they encode
 * were measured in a hidden Excel of our own (the same runner job through server/excel/excel-runner.ps1 and replayLog, see RESULTS.md).
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { createEditor, moveColumns } from '../../../src/core/doc';
import { readBook, replayLog } from '../../../src/core/xlsx';
import { adjustSheetViews, frozenAfter } from '../../../src/core/xlsx/patch/views';
import { buildSamplePremerged, samplesAvailable } from '../../doc/sampleMerge';
import { runRunnerJob } from './excel';

describe.skipIf(!samplesAvailable())('engine-package-view: the doc and replayLog agree, and so does Excel (EXCEL_TESTS=1), when the whole frozen block is moved away', () => {
  it('the sample list freezes 2 columns (C4): moving 学生ID + 氏名 to the right leaves ONE frozen column in the doc and in the file', async () => {
    const p = await buildSamplePremerged({});
    const s0 = createEditor(p.book, p.layout, p.sourceInfo, p.region);
    expect(s0.doc.features.freeze).toEqual({ cols: 2, rows: 3 });
    const ids = s0.doc.columns.slice(0, 2).map((c) => c.id);
    const res = moveColumns(s0, ids, 8, { unmergeTitles: true });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.snapshot.doc.features.freeze).toEqual({ cols: 1, rows: 3 });
    const sheet = (await readBook(await replayLog(p.bytes, res.snapshot.log), 'after.xlsx')).sheets.find((x) => x.name === s0.doc.sheetName)!;
    expect(sheet.features.freeze).toEqual({ cols: 1, rows: 3 });
  });

  it.skipIf(process.env.EXCEL_TESTS !== '1')('real Excel: the runner leaves one frozen column too', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'v9-regress-'));
    try {
      const p = await buildSamplePremerged({});
      const s0 = createEditor(p.book, p.layout, p.sourceInfo, p.region);
      const res = moveColumns(s0, s0.doc.columns.slice(0, 2).map((c) => c.id), 8, { unmergeTitles: true });
      expect(res.ok).toBe(true);
      if (!res.ok) return;
      const pre = path.join(dir, 'pre.xlsx');
      fs.writeFileSync(pre, p.bytes);
      const xlFile = path.join(dir, 'xl.xlsx');
      const out = await runRunnerJob(pre, res.snapshot.log, xlFile, s0.doc.sheetName, { tempRoot: dir, cacheDir: path.join(dir, 'cache') });
      expect(out.ok, JSON.stringify(out)).toBe(true);
      const xl = (await readBook(new Uint8Array(fs.readFileSync(xlFile)), 'xl.xlsx')).sheets.find((x) => x.name === s0.doc.sheetName)!;
      expect(xl.features.freeze).toEqual({ cols: 1, rows: 3 });
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  }, 180_000);
});

describe('engine-package-view: deleting / moving out ALL frozen lines leaves ONE frozen line (Excel), not k', () => {
  const del = (at: number, count: number) => ({ type: 'delete', axis: 'col', at, count }) as const;
  const move = (from: number, count: number, to: number) => ({ type: 'move', axis: 'col', from, count, to }) as const;

  it('measured in Excel: k = 2, 3, 4 frozen columns deleted or moved away -> xSplit 1 (js used to keep k)', () => {
    for (const k of [1, 2, 3, 4]) {
      expect(frozenAfter(k, del(1, k)), `delete 1..${k}`).toBe(1);
      expect(frozenAfter(k, del(1, k + 2)), `delete 1..${k + 2}`).toBe(1);
      expect(frozenAfter(k, move(1, k, k + 6)), `move 1..${k} to the right`).toBe(1);
    }
  });

  it('what was right stays right: partial deletes / moves, inserts, ops behind the freeze', () => {
    expect(frozenAfter(3, move(2, 3, 9))).toBe(1); // columns 2..4 leave: one of the three frozen columns stays (measured)
    expect(frozenAfter(3, del(2, 1))).toBe(2);
    expect(frozenAfter(3, del(4, 2))).toBe(3);
    expect(frozenAfter(3, { type: 'insert', axis: 'col', at: 2, count: 2 })).toBe(5);
    expect(frozenAfter(3, { type: 'insert', axis: 'col', at: 4, count: 2 })).toBe(3);
    expect(frozenAfter(0, del(1, 3))).toBe(0);
    // a move inside / onto the block does nothing
    expect(frozenAfter(3, move(1, 2, 2))).toBe(3);
  });

  it('rows follow the same rule (measured: ySplit 3 -> 1 after deleting rows 1..3, 2 -> 1 after moving rows 1..2 away)', () => {
    expect(frozenAfter(3, { type: 'delete', axis: 'row', at: 1, count: 3 })).toBe(1);
    expect(frozenAfter(2, { type: 'move', axis: 'row', from: 1, count: 2, to: 8 })).toBe(1);
  });

  it('the <pane> element is rebuilt with the split of one column / its top-left cell', () => {
    const head = '<sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="C4" sqref="C4"/></sheetView></sheetViews>';
    const out = adjustSheetViews(head, move(1, 2, 5));
    expect(out).toContain('xSplit="1"');
    expect(out).toContain('ySplit="3"');
    expect(out).toContain('topLeftCell="B4"');
  });
});
