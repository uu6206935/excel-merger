/**
 * Regression tests for the defects the consistency oracle (task V1) found. Each test names the oracle class it came from; the real-Excel
 * facts they encode were measured in a hidden Excel of our own (probe scripts are described in RESULTS.md).
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { createEditor, insertColumns } from '../../../src/core/doc';
import { readBook, replayLog } from '../../../src/core/xlsx';
import { buildSamplePremerged, samplesAvailable } from '../../doc/sampleMerge';
import { PEOPLE, merge, type BookSpec } from '../../xlsx/formulaimport-helpers';
import { runRunnerJob } from './excel';

describe.skipIf(!samplesAvailable())('engine-package cols: a column inserted next to a hidden column is VISIBLE (Excel), in the doc and in replayLog', () => {
  it('insert right of the hidden helper column of the sample list: not hidden, width / customWidth of the left column kept', async () => {
    const p = await buildSamplePremerged({});
    const s0 = createEditor(p.book, p.layout, p.sourceInfo, p.region);
    const hiddenAt = s0.doc.columns.findIndex((c) => c.props.hidden);
    expect(hiddenAt).toBeGreaterThanOrEqual(0); // 重複チェック (L)
    const res = insertColumns(s0, hiddenAt + 1, 2);
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    const cols = res.snapshot.doc.columns;
    expect(cols[hiddenAt].props.hidden).toBe(true); // the hidden column stays hidden
    for (const k of [hiddenAt + 1, hiddenAt + 2]) {
      expect(cols[k].props.hidden).toBeUndefined();
      expect(cols[k].props.width).toBe(cols[hiddenAt].props.width); // Excel copies the width (stored value) of the left column
    }
    // the built-in engine writes the same
    const book = await readBook(await replayLog(p.bytes, res.snapshot.log), 'after.xlsx');
    const sheet = book.sheets.find((x) => x.name === s0.doc.sheetName)!;
    expect(sheet.colProps[hiddenAt]?.hidden).toBe(true);
    for (const k of [hiddenAt + 1, hiddenAt + 2]) {
      expect(sheet.colProps[k]?.hidden).toBeFalsy();
      expect(sheet.colProps[k]?.width).toBe(sheet.colProps[hiddenAt]?.width);
    }
  });
});

// The fact itself, measured with the app's Excel engine (EXCEL_TESTS=1): the runner inserts through Excel, so Excel decides.
describe.skipIf(!samplesAvailable() || process.env.EXCEL_TESTS !== '1')('real Excel: a column inserted next to a hidden column is visible, and the built-in engine agrees', () => {
  it('runner output == replayLog output for the hidden column and the two new columns (hidden, width, customWidth)', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cons-regress-'));
    try {
      const p = await buildSamplePremerged({});
      const s0 = createEditor(p.book, p.layout, p.sourceInfo, p.region);
      const hiddenAt = s0.doc.columns.findIndex((c) => c.props.hidden);
      const res = insertColumns(s0, hiddenAt + 1, 2);
      expect(res.ok).toBe(true);
      if (!res.ok) return;
      const pre = path.join(dir, 'pre.xlsx');
      fs.writeFileSync(pre, p.bytes);
      const xlFile = path.join(dir, 'xl.xlsx');
      const out = await runRunnerJob(pre, res.snapshot.log, xlFile, s0.doc.sheetName, { tempRoot: dir, cacheDir: path.join(dir, 'cache') });
      expect(out.ok, JSON.stringify(out)).toBe(true);
      const xl = (await readBook(new Uint8Array(fs.readFileSync(xlFile)), 'xl.xlsx')).sheets.find((x) => x.name === s0.doc.sheetName)!;
      const js = (await readBook(await replayLog(p.bytes, res.snapshot.log), 'js.xlsx')).sheets.find((x) => x.name === s0.doc.sheetName)!;
      for (const k of [hiddenAt, hiddenAt + 1, hiddenAt + 2]) {
        expect(!!xl.colProps[k]?.hidden, `hidden col ${k + 1}`).toBe(k === hiddenAt);
        expect(!!js.colProps[k]?.hidden, `hidden col ${k + 1} (built-in)`).toBe(k === hiddenAt);
        expect(Math.abs((js.colProps[k]?.width ?? 0) - (xl.colProps[k]?.width ?? 0)), `width col ${k + 1}`).toBeLessThan(0.01);
      }
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  }, 180_000);
});

// ---------------------------------------------------------------------------------------------------------------------
// OPEN defect found by the oracle (class append-range-running, premerge / patch/formula.ts extendArea — not mine to change, reported in
// RESULTS.md). `it.fails` documents it: the test passes while the defect exists and turns RED when somebody fixes it (then make it a plain `it`).
// ---------------------------------------------------------------------------------------------------------------------
describe('join full: appended rows and the formulas of the base (open defect)', () => {
  const base = (): BookSpec => ({
    sheet: '名簿',
    rows: [['ID', '氏名', '点', '累計'], ...PEOPLE.map((p, i) => [p, `名前${i + 1}`, 10 * (i + 1), { f: `SUM($C$2:C${i + 2})` }])],
  });
  const src = (): BookSpec => ({ sheet: '評価', rows: [['ID', '氏名', '評価'], ...[...PEOPLE, 'P9'].map((p, i) => [p, `名前${i + 1}`, i])] });
  const plan = () => [...PEOPLE.map((_p, i) => ({ base: i + 2, src: i + 2 })), { base: null, src: 7 }];
  const keyFill = [{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }];

  it('a running total in a row INSIDE the table is not touched by the appended row', async () => {
    const m = await merge({ base: base(), src: src(), imports: [3], rows: plan(), keyFill });
    expect(m.cell(3, 4)!.f).toBe('SUM($C$2:C3)');
    expect(m.cell(7, 1)!.v).toBe('P9'); // the appended row exists (row 7)
  });

  it.fails('a running total in the LAST base row ends at its own row: =SUM($C$2:C6) must not become =SUM($C$2:C7) (it sums the appended people too)', async () => {
    const m = await merge({ base: base(), src: src(), imports: [3], rows: plan(), keyFill });
    expect(m.cell(6, 4)!.f).toBe('SUM($C$2:C6)');
  });
});

// ---------------------------------------------------------------------------------------------------------------------
// class lookup-silent / new-error: swapping the KEY column of a VLOOKUP range with another column of the range
// ---------------------------------------------------------------------------------------------------------------------
import { lookupWarnings, type StructOp } from '../../../src/core/formula/excelRef';
import { moveColumns } from '../../../src/core/doc';
import { buildRichBase, buildRowSource } from './gen';
import { DEFAULT_PARAMS, loadKeySourceBytes, mergeAll } from './lib';

describe('VLOOKUP: the search column (first column of the range) is part of the lookup-by-column-number warning', () => {
  const ctx = { hostSheet: 'Data', targetSheet: 'Data' };
  const mv = (from: number, to: number, count = 1): StructOp => ({ type: 'move', axis: 'col', from, count, to });
  const w = (f: string, op: StructOp) => lookupWarnings(f, op, ctx);

  it('excelRef: A <-> B inside VLOOKUP(x,A:D,4,..) leaves column 4 alone but searches another column: warned (it was silent, the result is #N/A)', () => {
    const f = 'VLOOKUP(A2,$A$2:$D$31,4,FALSE)';
    const swapAB = w(f, mv(1, 3)); // A moves behind B
    expect(swapAB.map((x) => x.fn)).toEqual(['VLOOKUP']);
    expect(swapAB[0].message).toMatch(/先頭の列（検索する列）/);
    expect(w(f, mv(2, 1)).map((x) => x.fn)).toEqual(['VLOOKUP']); // B before A: the same swap
    // what stays silent: the search column and the returned column are where they were
    expect(w(f, mv(2, 4))).toEqual([]); // B and C swap
    expect(w(f, mv(2, 5, 1)).length).toBe(1); // B moved to the end of the range: D is now 3rd (the returned column moved: the old rule)
    expect(w(f, mv(5, 7))).toEqual([]); // outside the range
    expect(w('VLOOKUP(A2,$A$2:$D$31,2,FALSE)', mv(3, 4))).toEqual([]); // C and D swap, columns 1 and 2 stay
  });

  it('doc: swapping the ID and 氏名 columns of the rich base (VLOOKUP over A:D) warns, and Excel agrees that the lookup is broken', async () => {
    const base = await loadKeySourceBytes((await buildRichBase('plain')).bytes, '名簿_rich.xlsx', 'b');
    const other = await loadKeySourceBytes((await buildRowSource('same', 11)).bytes, '評価.xlsx', 's2');
    const m = await mergeAll(base, [other], { ...DEFAULT_PARAMS, dropOtherSheets: false });
    const r = moveColumns(m.editor, ['b1'], 2); // ID behind 氏名
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.warnings.join('\n')).toMatch(/列番号・位置で列を指定する関数/);
    expect(r.warnings.join('\n')).toMatch(/先頭の列（検索する列）/);
  });
});
