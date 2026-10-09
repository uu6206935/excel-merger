/** logToRunnerJob: op log -> excel-runner.ps1 job; isUnsafeFormula (DDE / external link / web request detector). */
import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { isUnsafeFormula, logToRunnerJob, toRunnerCell } from '../../src/core/export/comJob';
import type { LoggedOp } from '../../src/core/types';

const S = '名簿';

describe('logToRunnerJob', () => {
  it('maps struct ops 1:1 (1-based, move = clipFree with original-coordinate `to`)', () => {
    const log: LoggedOp[] = [
      { kind: 'struct', sheet: S, op: { type: 'move', axis: 'col', from: 2, count: 3, to: 9 } },
      { kind: 'struct', sheet: S, op: { type: 'move', axis: 'col', from: 7, count: 1, to: 1 } },
      { kind: 'struct', sheet: S, op: { type: 'insert', axis: 'col', at: 4, count: 2 } },
      { kind: 'struct', sheet: S, op: { type: 'delete', axis: 'col', at: 6, count: 1 } },
      { kind: 'struct', sheet: S, op: { type: 'insert', axis: 'row', at: 10, count: 1 } },
      { kind: 'struct', sheet: S, op: { type: 'delete', axis: 'row', at: 3, count: 2 } },
      { kind: 'struct', sheet: S, op: { type: 'move', axis: 'row', from: 5, count: 1, to: 2 } },
    ];
    const job = logToRunnerJob(log, { format: 'xlsx', sheet: S });
    expect(job.version).toBe(1);
    expect(job.format).toBe('xlsx');
    expect(job.options).toEqual({ recalc: true, clipboard: 'restoreEachOp' });
    expect(job.input).toBeUndefined();
    expect(job.ops).toEqual([
      { op: 'moveCols', sheet: S, start: 2, count: 3, to: 9, method: 'clipFree' },
      { op: 'moveCols', sheet: S, start: 7, count: 1, to: 1, method: 'clipFree' },
      { op: 'insertCols', sheet: S, at: 4, count: 2, formatFrom: 'left' },
      { op: 'deleteCols', sheet: S, at: 6, count: 1 },
      { op: 'insertRows', sheet: S, at: 10, count: 1, formatFrom: 'above' },
      { op: 'deleteRows', sheet: S, at: 3, count: 2 },
      { op: 'moveRows', sheet: S, start: 5, count: 1, to: 2, method: 'clipFree' },
    ]);
  });

  it('drops no-op moves (target inside or right after the block) and empty ops', () => {
    const log: LoggedOp[] = [
      { kind: 'struct', sheet: S, op: { type: 'move', axis: 'col', from: 3, count: 2, to: 3 } },
      { kind: 'struct', sheet: S, op: { type: 'move', axis: 'col', from: 3, count: 2, to: 5 } },
      { kind: 'struct', sheet: S, op: { type: 'move', axis: 'col', from: 3, count: 2, to: 4 } },
      { kind: 'struct', sheet: S, op: { type: 'insert', axis: 'col', at: 3, count: 0 } },
      { kind: 'cells', sheet: S, writes: [], reason: 'edit' },
    ];
    expect(logToRunnerJob(log, { format: 'xlsm', sheet: S }).ops).toEqual([]);
    expect(logToRunnerJob(log, { format: 'xlsm', sheet: S }).format).toBe('xlsm');
  });

  it('uses the op sheet, falling back to opts.sheet; widths are sent in the stored <col width> unit', () => {
    const log = [{ kind: 'colWidth', sheet: '', col: 3, width: 12.5 }] as LoggedOp[];
    expect(logToRunnerJob(log, { format: 'xlsx', sheet: 'Main' }).ops).toEqual([{ op: 'setColumnWidth', sheet: 'Main', col: 3, width: 12.5, unit: 'stored' }]);
    const log2: LoggedOp[] = [
      { kind: 'colWidth', sheet: '集計', col: 1, width: 300 },
      { kind: 'colWidth', sheet: '集計', col: 2, width: -1 },
      { kind: 'colWidth', sheet: '集計', col: 2, width: Number.NaN },
    ];
    expect(logToRunnerJob(log2, { format: 'xlsx', sheet: 'Main' }).ops).toEqual([{ op: 'setColumnWidth', sheet: '集計', col: 1, width: 256, unit: 'stored' }]);
  });

  it('maps unmerge (A1 refs only, centerAcross flag)', () => {
    const log: LoggedOp[] = [
      { kind: 'unmerge', sheet: S, refs: ['A1:F1', ' $B$2:$C$3 ', 'Sheet2!A1:B1', 'X'], centerAcross: true },
      { kind: 'unmerge', sheet: S, refs: ['G1:H1'], centerAcross: false },
      { kind: 'unmerge', sheet: S, refs: [], centerAcross: true },
    ];
    expect(logToRunnerJob(log, { format: 'xlsx', sheet: S }).ops).toEqual([
      { op: 'unmerge', sheet: S, refs: ['A1:F1', '$B$2:$C$3'], centerAcross: true },
      { op: 'unmerge', sheet: S, refs: ['G1:H1'], centerAcross: false },
    ]);
  });

  it('marks toValues writes (arrays under them are replaced by their values) and keeps numFmt / dyn', () => {
    const log: LoggedOp[] = [
      { kind: 'cells', sheet: S, reason: 'toValues', writes: [{ row: 3, col: 10, cell: { t: 'n', v: 20 } }] },
      {
        kind: 'cells',
        sheet: S,
        reason: 'edit',
        writes: [
          { row: 5, col: 2, cell: { t: 'd', v: 46000, numFmt: 'yyyy/m/d' } },
          { row: 6, col: 2, cell: { t: 'n', v: 0.5, numFmt: '0%' } },
          { row: 7, col: 2, cell: { t: 'n', v: 0, f: '_xlfn.SEQUENCE(3)', dyn: true } },
          { row: 8, col: 2, cell: { t: 'n', v: 0, f: 'SUM(A1:A3)' } },
        ],
      },
    ];
    expect(logToRunnerJob(log, { format: 'xlsx', sheet: S }).ops).toEqual([
      { op: 'setCells', sheet: S, at: 'J3', rows: [[20]], toValues: true },
      {
        op: 'setCells',
        sheet: S,
        at: 'B5',
        rows: [[{ t: 'n', v: 46000, numFmt: 'yyyy/m/d' }], [{ t: 'n', v: 0.5, numFmt: '0%' }], [{ f: '=SEQUENCE(3)', dyn: true }], [{ f: '=SUM(A1:A3)' }]],
      },
    ]);
  });

  it('groups cell writes into vertical runs (last write wins) with the runner cell encoding', () => {
    const log: LoggedOp[] = [
      {
        kind: 'cells',
        sheet: S,
        reason: 'toValues',
        writes: [
          { row: 5, col: 3, cell: { t: 'n', v: 1 } },
          { row: 2, col: 3, cell: { t: 's', v: '001' } },
          { row: 3, col: 3, cell: { t: 'b', v: true } },
          { row: 4, col: 3, cell: { t: 'e', v: '#N/A' } },
          { row: 2, col: 1, cell: { t: 'n', v: 0, f: 'SUM(B2:B9)' } },
          { row: 3, col: 1, cell: null },
          { row: 7, col: 3, cell: { t: 'd', v: 45000.5 } },
          { row: 5, col: 3, cell: { t: 's', v: '=not a formula' } },
          { row: 1, col: 28, cell: { t: 'z', v: null } },
        ],
      },
    ];
    const ops = logToRunnerJob(log, { format: 'xlsx', sheet: S }).ops;
    expect(ops).toEqual([
      { op: 'setCells', sheet: S, at: 'A2', rows: [[{ f: '=SUM(B2:B9)' }], [null]], toValues: true },
      { op: 'setCells', sheet: S, at: 'C2', rows: [['001'], [true], [{ e: '#N/A' }], ['=not a formula']], toValues: true },
      { op: 'setCells', sheet: S, at: 'C7', rows: [[45000.5]], toValues: true },
      { op: 'setCells', sheet: S, at: 'AB1', rows: [[null]], toValues: true },
    ]);
  });

  it('moves inside an Excel table, also to the gap right after it (to = last table column + 1), go to the runner as they are', () => {
    // table B2:I22 (columns 2..9): the doc / patcher accept these as in-table reorders; the runner recognises the table itself
    // (Excel's own Cut + Insert Cut Cells) and verifies the header order
    const log: LoggedOp[] = [
      { kind: 'struct', sheet: S, op: { type: 'move', axis: 'col', from: 6, count: 1, to: 5 } },
      { kind: 'struct', sheet: S, op: { type: 'move', axis: 'col', from: 3, count: 2, to: 9 } },
      { kind: 'struct', sheet: S, op: { type: 'move', axis: 'col', from: 2, count: 1, to: 10 } }, // to = c2 + 1
      { kind: 'struct', sheet: S, op: { type: 'move', axis: 'col', from: 9, count: 1, to: 10 } }, // the last column "after itself": no-op
      { kind: 'struct', sheet: S, op: { type: 'move', axis: 'col', from: 8, count: 2, to: 10 } }, // the last two columns to the end: no-op
    ];
    expect(logToRunnerJob(log, { format: 'xlsx', sheet: S }).ops).toEqual([
      { op: 'moveCols', sheet: S, start: 6, count: 1, to: 5, method: 'clipFree' },
      { op: 'moveCols', sheet: S, start: 3, count: 2, to: 9, method: 'clipFree' },
      { op: 'moveCols', sheet: S, start: 2, count: 1, to: 10, method: 'clipFree' },
    ]);
  });

  it('quote: true (strictly) travels as { t: "s", v, quote: true }; every other text stays a bare string (no guessing)', () => {
    const log: LoggedOp[] = [
      {
        kind: 'cells',
        sheet: S,
        reason: 'edit',
        writes: [
          { row: 2, col: 4, cell: { t: 's', v: '001', quote: true } },
          { row: 3, col: 4, cell: { t: 's', v: '03-1234-5678' } }, // phone number: Excel keeps it as text by itself
          { row: 4, col: 4, cell: { t: 's', v: '1/2', quote: false } }, // the doc said no: the runner must not add one
          { row: 5, col: 4, cell: { t: 's', v: '=text', quote: true, numFmt: '@' } },
          { row: 6, col: 4, cell: { t: 's', v: 'abc', numFmt: 'General' } },
        ],
      },
    ];
    expect(logToRunnerJob(log, { format: 'xlsx', sheet: S }).ops).toEqual([
      {
        op: 'setCells',
        sheet: S,
        at: 'D2',
        rows: [[{ t: 's', v: '001', quote: true }], ['03-1234-5678'], ['1/2'], [{ t: 's', v: '=text', quote: true, numFmt: '@' }], [{ t: 's', v: 'abc', numFmt: 'General' }]],
      },
    ]);
  });

  it('keeps the order of the log (struct ops and cell writes interleaved)', () => {
    const log: LoggedOp[] = [
      { kind: 'cells', sheet: S, reason: 'toValues', writes: [{ row: 2, col: 4, cell: { t: 'n', v: 3 } }] },
      { kind: 'struct', sheet: S, op: { type: 'delete', axis: 'col', at: 2, count: 1 } },
      { kind: 'cells', sheet: S, reason: 'edit', writes: [{ row: 2, col: 3, cell: { t: 's', v: 'x' } }] },
    ];
    expect(logToRunnerJob(log, { format: 'xlsx', sheet: S }).ops.map((o) => o.op)).toEqual(['setCells', 'deleteCols', 'setCells']);
  });
});

describe('toRunnerCell', () => {
  it('encodes every cell type', () => {
    expect(toRunnerCell(null)).toBeNull();
    expect(toRunnerCell({ t: 'z', v: null })).toBeNull();
    expect(toRunnerCell({ t: 's', v: '' })).toBeNull();
    expect(toRunnerCell({ t: 's', v: 'abc' })).toBe('abc');
    expect(toRunnerCell({ t: 's', v: 12 })).toBe('12');
    expect(toRunnerCell({ t: 'n', v: -0.5 })).toBe(-0.5);
    expect(toRunnerCell({ t: 'n', v: 'x' })).toBe('x');
    expect(toRunnerCell({ t: 'd', v: 44927 })).toBe(44927);
    expect(toRunnerCell({ t: 'b', v: false })).toBe(false);
    expect(toRunnerCell({ t: 'b', v: 'TRUE' })).toBe(true);
    expect(toRunnerCell({ t: 'e', v: '#div/0!' })).toEqual({ e: '#DIV/0!' });
    expect(toRunnerCell({ t: 'e', v: '#WHAT' })).toBe('#WHAT');
    expect(toRunnerCell({ t: 'n', v: 1, f: 'A1+1' })).toEqual({ f: '=A1+1' });
    expect(toRunnerCell({ t: 's', v: 'x', f: '"a"&"b"' })).toEqual({ f: '="a"&"b"' });
  });

  it('quote applies to non-empty text only, strictly === true', () => {
    expect(toRunnerCell({ t: 's', v: '001', quote: true })).toEqual({ t: 's', v: '001', quote: true });
    expect(toRunnerCell({ t: 's', v: '001', quote: true, numFmt: '@' })).toEqual({ t: 's', v: '001', quote: true, numFmt: '@' });
    expect(toRunnerCell({ t: 's', v: 12, quote: true })).toEqual({ t: 's', v: '12', quote: true });
    expect(toRunnerCell({ t: 's', v: '001', quote: false })).toBe('001');
    expect(toRunnerCell({ t: 's', v: '001', quote: 1 as unknown as boolean })).toBe('001');
    expect(toRunnerCell({ t: 's', v: '', quote: true })).toBeNull(); // nothing to prefix
    expect(toRunnerCell({ t: 'n', v: 5, quote: true })).toBe(5);
    expect(toRunnerCell({ t: 'z', v: null, quote: true })).toBeNull();
    expect(toRunnerCell({ t: 'n', v: 0, f: 'A1+1', quote: true })).toEqual({ f: '=A1+1' });
    expect(toRunnerCell({ t: 's', v: 'x', quote: true, numFmt: 'General' })).toEqual({ t: 's', v: 'x', quote: true, numFmt: 'General' });
  });

  it('legacy formulas (Range.Formula) vs dynamic-array formulas (Range.Formula2), UI form', () => {
    expect(toRunnerCell({ t: 'n', v: 0, f: '_xlfn.XLOOKUP(A1,B:B,C:C)' })).toEqual({ f: '=XLOOKUP(A1,B:B,C:C)' });
    expect(toRunnerCell({ t: 'n', v: 0, f: '_xlfn._xlws.SORT(A2:A9)', dyn: true })).toEqual({ f: '=SORT(A2:A9)', dyn: true });
    expect(toRunnerCell({ t: 'n', v: 0, f: 'SUM(_xlfn.ANCHORARRAY(D2))', dyn: false })).toEqual({ f: '=SUM(D2#)' });
    expect(toRunnerCell({ t: 'n', v: 0, f: '_xlfn.SINGLE(A1:A9)' })).toEqual({ f: '=@A1:A9' });
  });

  it('number formats travel with the value ({t, v, numFmt}) or the formula ({f, numFmt})', () => {
    expect(toRunnerCell({ t: 'd', v: 45000, numFmt: 'yyyy/m/d' })).toEqual({ t: 'n', v: 45000, numFmt: 'yyyy/m/d' });
    expect(toRunnerCell({ t: 'n', v: 1000, numFmt: '"¥"#,##0' })).toEqual({ t: 'n', v: 1000, numFmt: '"¥"#,##0' });
    expect(toRunnerCell({ t: 's', v: '001', numFmt: '@' })).toEqual({ t: 's', v: '001', numFmt: '@' });
    expect(toRunnerCell({ t: 'b', v: true, numFmt: 'General' })).toEqual({ t: 'b', v: true, numFmt: 'General' });
    expect(toRunnerCell({ t: 'e', v: '#n/a', numFmt: '0' })).toEqual({ t: 'e', v: '#N/A', numFmt: '0' });
    expect(toRunnerCell({ t: 'z', v: null, numFmt: '0%' })).toEqual({ t: 'z', v: null, numFmt: '0%' });
    expect(toRunnerCell({ t: 'n', v: 0, f: 'A1*2', numFmt: '0.0%' })).toEqual({ f: '=A1*2', numFmt: '0.0%' });
    expect(toRunnerCell({ t: 'n', v: 5, numFmt: '' })).toBe(5);
  });
});

describe('isUnsafeFormula (the same rule and vectors as XlRunner.Cells.UnsafeFormula in the runner)', () => {
  const V = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'unsafe-formulas.json'), 'utf8')) as { dangerous: string[]; harmless: string[] };
  it('flags DDE, external workbook references (local, UNC, URL) and web / RTD / XLM call-outs', () => {
    for (const f of V.dangerous) expect(isUnsafeFormula(f), f).toBe(true);
  });
  it('leaves ordinary formulas alone', () => {
    for (const f of V.harmless) expect(isUnsafeFormula(f), f).toBe(false);
  });
  it('also works on the file form (no leading =) that the op log carries', () => {
    expect(isUnsafeFormula("cmd|' /C calc'!A0")).toBe(true);
    expect(isUnsafeFormula("'[x.xlsx]Sheet1'!A1")).toBe(true);
    expect(isUnsafeFormula('SUM(A1:A3)')).toBe(false);
    expect(isUnsafeFormula('_xlfn.WEBSERVICE("x")')).toBe(true);
  });
});
