/**
 * Round 8 (user report: formulas of the 2nd and later files became plain / blank cells): the UI side.
 *   - the pure text helpers of 「数式の取り込み」 (src/state/formulaReport.ts)
 *   - keyPairs reach buildPremerged (restricted to the files that take part), also through withoutSources
 *   - the merge toast, the 警告 list and the 数式→値 rows of the check panel
 *   - the reducer keeps / drops MergeReport.formulaStats with the merge
 */
import { describe, expect, it } from 'vitest';
import type { JoinResult } from '../../src/core/contracts';
import type { CheckItem, DocColumn, MergeReport, SourceBook, TableRegion } from '../../src/core/types';
import { createController } from '../../src/state/controller';
import {
  exampleCellColumn,
  FORMULA_EMPTY_HINT,
  FORMULA_RECALC_LINE,
  formulaCheckRows,
  formulaColumnsOf,
  formulaReportViews,
  formulaToastLines,
  formulaToastParts,
  formulaTotals,
  formulaWarningText,
  isFormulaImportItem,
  reasonBreakdown,
} from '../../src/state/formulaReport';
import { keyPairsFor, withoutSources } from '../../src/state/keysPlan';
import { createMockServices } from '../../src/state/mock';
import { fidelityNotes } from '../../src/state/notes';
import { initialState, reducer } from '../../src/state/reducer';
import type { Services } from '../../src/state/services';
import type { FileEntry } from '../../src/state/types';
import { editor, file, store } from './fixtures';

const report = (patch: Partial<MergeReport> = {}): MergeReport => ({ mergedRows: 60, matchedRows: 60, baseOnlyRows: 0, appendedRows: 0, importedColumns: 3, formulasKept: 0, formulasToValues: 0, items: [], ...patch });
const stat = (sourceId: string, label: string, kept: number, toValues: number, empty: number, reasons: Record<string, number> = {}) => ({ sourceId, label, kept, toValues, empty, reasons });
const nameOf = (sid: string, label: string) => ({ s2: '02_面接評価.xlsx', s4: '05_研修出欠.xlsx' })[sid as 's2' | 's4'] ?? label;
const item = (code: string, sourceId: string, message: string, rows?: number[], severity: CheckItem['severity'] = 'warning'): CheckItem => ({ severity, code, sourceId, message, ...(rows ? { rows } : {}) });

describe('formulaReportViews: the plain-Japanese lines', () => {
  it('the example of the request: kept / values with the breakdown / empty', () => {
    const [v] = formulaReportViews(report({ formulaStats: [stat('s2', '02_面接評価', 147, 12, 0, { otherSheet: 8, externalCell: 4 })] }), nameOf);
    expect(v.sentence).toBe('02_面接評価.xlsx: 数式 147 件を数式のまま取り込みました／12 件は値にしました（内訳: 他のシートを参照 8、表の外のセル 4）');
    expect(v.tone).toBe('note');
    expect(v.hint).toBeNull();
    expect(v.emptyLine).toBeNull();
  });

  it('empty cells make it a warning with the hint; the breakdown is then not claimed for the values alone', () => {
    const [v] = formulaReportViews(report({ formulaStats: [stat('s2', '02_面接評価', 147, 12, 3, { otherSheet: 11, externalCell: 4 })] }), nameOf);
    expect(v.sentence).toBe('02_面接評価.xlsx: 数式 147 件を数式のまま取り込みました／12 件は値にしました／空白 3 件（元のファイルに計算結果がありません）');
    expect(v.tone).toBe('warn');
    expect(v.hint).toBe('元のExcelで一度開いて保存し直すと取り込めることがあります');
    expect(v.hint).toBe(FORMULA_EMPTY_HINT);
    // the contract's reasons count the values AND the empty cells together (11 + 4 = 12 + 3): the details say so
    expect(v.reasonsIncludeEmpty).toBe(true);
    expect(v.reasons.reduce((n, r) => n + r.count, 0)).toBe(15);
    expect(v.valuesLine).not.toContain('内訳');
  });

  it('everything kept: information only, no warning entry', () => {
    const [v] = formulaReportViews(report({ formulaStats: [stat('s2', '02_面接評価', 116, 0, 0)] }), nameOf);
    expect(v.sentence).toBe('02_面接評価.xlsx: 数式 116 件を数式のまま取り込みました');
    expect(v.tone).toBe('ok');
    expect(v.valuesLine).toBeNull();
    expect(v.emptyLine).toBeNull();
    expect(formulaWarningText(v)).toBeNull();
  });

  it('only values (nothing kept): the line says 数式, a lone empty line is a warning too', () => {
    const [a] = formulaReportViews(report({ formulaStats: [stat('s2', '02_面接評価', 0, 5, 0, { rowRelation: 5 })] }), nameOf);
    expect(a.sentence).toBe('02_面接評価.xlsx: 数式 5 件は値にしました（内訳: 前後の行との関係 5）');
    const [b] = formulaReportViews(report({ formulaStats: [stat('s2', '02_面接評価', 0, 0, 7, { otherSheet: 7 })] }), nameOf);
    expect(b.sentence).toBe('02_面接評価.xlsx: 空白 7 件（元のファイルに計算結果がありません）');
    expect(b.tone).toBe('warn');
    expect(formulaWarningText(b)).toContain('元のExcelで一度開いて保存し直すと取り込めることがあります');
  });

  it('reasons are sorted by count; an unknown code is described neutrally; the file name falls back to the label', () => {
    const [v] = formulaReportViews(report({ formulaStats: [stat('s9', 'X', 0, 6, 0, { externalCell: 1, otherSheet: 3, weird: 2 })] }));
    expect(v.name).toBe('X');
    expect(v.reasons.map((r) => [r.code, r.count])).toEqual([['otherSheet', 3], ['weird', 2], ['externalCell', 1]]);
    expect(reasonBreakdown(v.reasons)).toBe('他のシートを参照 3、対応していない書き方 2、表の外のセル 1');
  });

  it('numbers are written with thousands separators', () => {
    const [v] = formulaReportViews(report({ formulaStats: [stat('s2', 'A', 12345, 0, 0)] }));
    expect(v.keptLine).toBe('数式 12,345 件を数式のまま取り込みました');
  });

  it('no formulaStats (older reports, mock), no formulas, or an all-zero entry: no block', () => {
    expect(formulaReportViews(null)).toEqual([]);
    expect(formulaReportViews(report())).toEqual([]);
    expect(formulaReportViews(report({ formulaStats: [] }))).toEqual([]);
    expect(formulaReportViews(report({ formulaStats: [stat('s2', 'A', 0, 0, 0)] }))).toEqual([]);
  });

  it('kept formulas without a stored result get the "Excel shows the result" line; the report items of the file become the details', () => {
    const items: CheckItem[] = [
      item('formulasKept', 's2', '「02_面接評価」の数式 116 件は数式のまま残しました', undefined, 'info'),
      item('formulasRecalc', 's2', '「02_面接評価」の数式 116 件は、結合後の表に合わせて付け替えたため…', [4, 5], 'info'),
      item('formulasToValues', 's2', '「02_面接評価」の他のシートを参照している数式 3 件は値にしました。', [7, 8, 9]),
      item('formulasToValues', 's4', '別のファイルの項目', [1]),
      item('formulasNoCache', 's2', '「02_面接評価」の数式 2 件は、元のファイルに計算結果が保存されていない…', [11, 12]),
      item('duplicate', 's2', '重複', [3]),
    ];
    const [v] = formulaReportViews(report({ items, formulaStats: [stat('s2', '02_面接評価', 116, 3, 2, { otherSheet: 5 })] }), nameOf);
    expect(v.recalcLine).toBe(FORMULA_RECALC_LINE);
    expect(v.details.map((d) => [d.code, d.rows, d.cell])).toEqual([
      ['formulasRecalc', [4, 5], 'formula'],
      ['formulasToValues', [7, 8, 9], 'plain'],
      ['formulasNoCache', [11, 12], 'plain'],
    ]);
    expect(isFormulaImportItem(items[0])).toBe(true);
    expect(isFormulaImportItem(items[5])).toBe(false);
    const [w] = formulaReportViews(report({ formulaStats: [stat('s2', '02_面接評価', 116, 0, 0)] }), nameOf);
    expect(w.recalcLine).toBeNull();
  });
});

describe('toast, 警告, check panel', () => {
  const rep = report({
    formulasToValues: 17,
    items: [
      item('formulasToValues', 's2', 'm1', [7, 8]),
      item('formulasNoCache', 's2', 'm2', [11]),
      item('formulasToValues', 's4', 'm3', [8]),
      item('formulaRangesWidened', 's4', 'm4', [20], 'info'),
      item('formulasPlaceholderCache', 's4', '「05_研修出欠」は、数式の計算結果が保存されていない形式のファイルです', undefined, 'warning'),
      item('formulasKept', 's4', '「05_研修出欠」の数式 31 件は数式のまま', undefined, 'info'),
    ],
    formulaStats: [stat('s2', '02_面接評価', 116, 12, 3, { otherSheet: 15 }), stat('s4', '05_研修出欠', 31, 2, 0, { rowRelation: 2 })],
  });
  const views = formulaReportViews(rep, nameOf);

  it('toast parts count all files; the detail is one sentence per file (+ the hint when cells are empty)', () => {
    expect(formulaTotals(views)).toEqual({ kept: 147, toValues: 14, empty: 3, files: 2 });
    expect(formulaToastParts(views)).toEqual(['数式 147 件は数式のまま', '値にした数式 14 件', '空白になった数式 3 件']);
    expect(formulaToastLines(views)).toEqual([
      '02_面接評価.xlsx: 数式 116 件を数式のまま取り込みました／12 件は値にしました／空白 3 件（元のファイルに計算結果がありません）',
      '05_研修出欠.xlsx: 数式 31 件を数式のまま取り込みました／2 件は値にしました（内訳: 前後の行との関係 2）',
      '元のExcelで一度開いて保存し直すと取り込めることがあります',
    ]);
    expect(formulaToastParts([])).toEqual([]);
    expect(formulaToastLines([])).toEqual([]);
  });

  it('the 警告 list gets one entry per file with values / empty cells instead of the generic line and the formula items', () => {
    const files = [
      { id: 's1', fileName: '01_学生名簿.xlsx' },
      { id: 's2', fileName: '02_面接評価.xlsx' },
      { id: 's4', fileName: '05_研修出欠.xlsx' },
    ] as FileEntry[];
    const notes = fidelityNotes(files, 's1', true, rep, { labels: { s2: '02_面接評価', s4: '05_研修出欠' } });
    expect(notes.find((n) => n.key === 'toValues')).toBeUndefined();
    const f2 = notes.find((n) => n.key === 'formulas-s2')!;
    expect(f2.severity).toBe('warning');
    expect(f2.text).toContain('02_面接評価.xlsx: 12 件は値にしました／空白 3 件');
    expect(f2.text).toContain('元のExcelで一度開いて保存し直すと取り込めることがあります');
    expect(notes.find((n) => n.key === 'formulas-s4')!.text).toContain('2 件は値にしました');
    // the report's formula items are in the block, not repeated here — except the file-level warning about stored zeros
    const texts = notes.map((n) => n.text);
    expect(texts.some((t) => /^m[1-4]/.test(t))).toBe(false);
    expect(texts.some((t) => t.includes('数式 31 件は数式のまま'))).toBe(false);
    expect(texts.some((t) => t.includes('計算結果が保存されていない形式'))).toBe(true);
  });

  it('a report without formulaStats keeps the generic line and all its items', () => {
    const notes = fidelityNotes([{ id: 's1', fileName: 'a.xlsx' }] as FileEntry[], 's1', true, report({ formulasToValues: 3, items: [item('formulasToValues', 's2', 'm1', [7])] }));
    expect(notes.find((n) => n.key === 'toValues')?.text).toContain('3 件');
    expect(notes.some((n) => n.text.startsWith('m1'))).toBe(true);
  });

  it('everything kept adds no 警告 entry', () => {
    const r = report({ formulaStats: [stat('s2', '02_面接評価', 116, 0, 0)], items: [item('formulasKept', 's2', 'kept', undefined, 'info')] });
    const notes = fidelityNotes([{ id: 's1', fileName: 'a.xlsx' }] as FileEntry[], 's1', true, r);
    expect(notes.some((n) => n.key.startsWith('formulas-'))).toBe(false);
    expect(notes.some((n) => n.text === 'kept')).toBe(false);
  });

  it('check panel rows: the example rows of the values / empty items, once each, with the file name', () => {
    const rows = formulaCheckRows(rep, nameOf);
    expect(rows.map((r) => [r.row, r.text, r.sub])).toEqual([
      [7, '7行目', '02_面接評価.xlsx: 値にした数式'],
      [8, '8行目', '02_面接評価.xlsx: 値にした数式'],
      [11, '11行目', '02_面接評価.xlsx: 空白になった数式'],
      [8, '8行目', '05_研修出欠.xlsx: 値にした数式'],
    ]);
    expect(new Set(rows.map((r) => r.key)).size).toBe(rows.length);
    expect(formulaCheckRows(report())).toEqual([]);
  });
});

describe('where the formulas are (preview hint, jump target)', () => {
  const region: TableRegion = { sheet: 'S', headerRow: 1, firstDataRow: 2, lastDataRow: 4, firstCol: 1, lastCol: 4 };
  const cell = (f?: string) => ({ t: 'n' as const, v: 1, ...(f ? { f } : {}) });
  const book = (fidelity: 'full' | 'values' = 'full'): SourceBook =>
    ({
      fidelity,
      sheets: [
        {
          name: 'S',
          rows: [
            [cell(), cell(), cell(), cell()],
            [cell(), cell(), cell('C2*2'), cell()],
            [cell(), cell(), cell('C3*2'), cell('A3')],
            [cell(), cell(), undefined, cell()],
          ],
        },
      ],
    }) as unknown as SourceBook;

  it('formulaColumnsOf counts the formula cells per column of the data rows (cached per book + region)', () => {
    const b = book();
    const m = formulaColumnsOf(b, region);
    expect([...m]).toEqual([[3, 2], [4, 1]]);
    expect(formulaColumnsOf(b, region)).toBe(m);
    // the header row is not a data row
    expect(formulaColumnsOf(b, { ...region, firstDataRow: 3 }).get(3)).toBe(1);
    expect(formulaColumnsOf(book('values'), region).size).toBe(0);
    expect(formulaColumnsOf(undefined, region).size).toBe(0);
    expect(formulaColumnsOf(b, { ...region, sheet: 'nope' }).size).toBe(0);
  });

  it('exampleCellColumn: the file\'s formula column whose cell is (not) a formula; fallbacks; null without columns', () => {
    const col = (id: string, origin: DocColumn['origin'], cells: DocColumn['cells']): DocColumn => ({ id, origin, props: {}, cells });
    const columns: DocColumn[] = [
      col('b1', { kind: 'base', col: 1 }, [{ t: 'n', v: 1 }]),
      col('i:s2:1', { kind: 'import', sourceId: 's2', col: 1 }, [{ t: 'n', v: 1 }]),
      col('i:s2:3', { kind: 'import', sourceId: 's2', col: 3 }, [{ t: 'n', v: 1, f: 'A1' }]),
      col('i:s2:4', { kind: 'import', sourceId: 's2', col: 4 }, [{ t: 'n', v: 1 }]),
      col('i:s3:2', { kind: 'import', sourceId: 's3', col: 2 }, [undefined]),
    ];
    const formulaCols = new Set([3, 4]);
    // the formula was not kept in column 4 (a value) -> 'plain' picks it; the kept one is column 3
    expect(exampleCellColumn(columns, 's2', 1, formulaCols, 'plain')).toBe('i:s2:4');
    expect(exampleCellColumn(columns, 's2', 1, formulaCols, 'formula')).toBe('i:s2:3');
    // nothing fits: the first formula column; no formula column known: the first column of the file
    expect(exampleCellColumn(columns, 's2', 1, new Set([3]), 'plain')).toBe('i:s2:3');
    expect(exampleCellColumn(columns, 's2', 1, new Set(), 'plain')).toBe('i:s2:1');
    expect(exampleCellColumn(columns, 's3', 1, formulaCols, 'plain')).toBe('i:s3:2');
    expect(exampleCellColumn(columns, 's9', 1, formulaCols, 'plain')).toBeNull();
  });
});

describe('keyPairs reach the premerge', () => {
  const pairs = { s2: [{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }], s3: [{ baseCol: 3, sourceCol: 2 }], s4: [] as Array<{ baseCol: number; sourceCol: number }> };

  it('keyPairsFor keeps the files that take part (and drops empty lists); undefined without keyPairs', () => {
    expect(keyPairsFor({ keyPairs: pairs }, ['s2', 's4'])).toEqual({ s2: pairs.s2 });
    expect(keyPairsFor({ keyPairs: pairs }, ['s2', 's3'])).toEqual({ s2: pairs.s2, s3: pairs.s3 });
    expect(keyPairsFor({ keyPairs: pairs }, [])).toEqual({});
    expect(keyPairsFor({}, ['s2'])).toBeUndefined();
    expect(keyPairsFor(null, ['s2'])).toBeUndefined();
  });

  it('withoutSources carries keyPairs over for the remaining files only', async () => {
    const { j } = await editor();
    const join: JoinResult = { ...j.join, keyPairs: { s2: [{ baseCol: 1, sourceCol: 1 }], s3: [{ baseCol: 2, sourceCol: 2 }] } };
    expect(withoutSources(join, []).keyPairs).toEqual(join.keyPairs);
    expect(withoutSources(join, ['s3']).keyPairs).toEqual({ s2: [{ baseCol: 1, sourceCol: 1 }] });
    expect(withoutSources({ ...join, keyPairs: undefined }, ['s3'])).not.toHaveProperty('keyPairs');
  });

  async function loaded(patch: Partial<Services> = {}) {
    const st = store();
    const services = { ...createMockServices('?mock=1&delay=0.05'), ...patch };
    const ctrl = createController({ services, dispatch: st.dispatch, getState: st.getState });
    ctrl.init();
    await ctrl.addFiles([file('学生回答.xlsx'), file('職場回答.xlsx'), file('人事回答.xlsx')]);
    await ctrl.idle();
    await ctrl.ensureAnalysis();
    return { st, ctrl };
  }

  it('the merge hands JoinResult.keyPairs to buildPremerged (the mock services have the field too)', async () => {
    const real = createMockServices('?mock=1&delay=0.05');
    const seen: Array<Record<string, unknown> | undefined> = [];
    const env = await loaded({
      buildPremerged: (input, p) => {
        seen.push(input.keyPairs);
        return real.buildPremerged(input, p);
      },
    });
    expect(env.st.state.joinResult?.keyPairs).toBeDefined();
    await env.ctrl.merge();
    expect(seen).toHaveLength(1);
    expect(seen[0]).toEqual(env.st.state.joinResult!.keyPairs);
    expect(Object.keys(seen[0]!).sort()).toEqual(['s2', 's3']);
  });

  it('a file without a usable key takes no part: its keyPairs are not handed over', async () => {
    const real = createMockServices('?mock=1&delay=0.05');
    const seen: Array<Record<string, unknown> | undefined> = [];
    const st = store();
    const ctrl = createController({
      services: {
        ...real,
        buildPremerged: (input, p) => {
          seen.push(input.keyPairs);
          return real.buildPremerged(input, p);
        },
      },
      dispatch: st.dispatch,
      getState: st.getState,
    });
    ctrl.init();
    await ctrl.addFiles([file('学生回答.xlsx'), file('職場回答.xlsx'), file('参加状況（キーなし）.xlsx')]);
    await ctrl.idle();
    await ctrl.ensureAnalysis();
    await ctrl.merge();
    expect(st.state.phase).toBe('merged');
    expect(Object.keys(seen[0] ?? {})).not.toContain('s3');
    expect(Object.keys(seen[0] ?? {})).toContain('s2');
  });
});

describe('merge toast and state', () => {
  it('the toast counts the formulas and names each file in the detail; the report is kept in the state', async () => {
    const real = createMockServices('?mock=1&delay=0.05');
    const formulaStats = [stat('s2', '職場回答', 40, 3, 2, { otherSheet: 5 })];
    const st = store();
    const ctrl = createController({
      services: {
        ...real,
        buildPremerged: async (input, p) => {
          const r = await real.buildPremerged(input, p);
          return { ...r, report: { ...r.report, formulaStats, formulasToValues: 5 } };
        },
      },
      dispatch: st.dispatch,
      getState: st.getState,
    });
    ctrl.init();
    await ctrl.addFiles([file('学生回答.xlsx'), file('職場回答.xlsx'), file('人事回答.xlsx')]);
    await ctrl.idle();
    await ctrl.ensureAnalysis();
    await ctrl.merge();
    expect(st.state.merged!.report.formulaStats).toEqual(formulaStats);
    const t = st.state.toast!;
    expect(t.text).toContain('数式 40 件は数式のまま');
    expect(t.text).toContain('値にした数式 3 件');
    expect(t.text).toContain('空白になった数式 2 件');
    expect(t.text).not.toContain('数式→値');
    expect(t.detail).toContain('職場回答.xlsx: 数式 40 件を数式のまま取り込みました／3 件は値にしました／空白 2 件（元のファイルに計算結果がありません）');
    expect(t.detail).toContain(FORMULA_EMPTY_HINT);
  });

  it('a report without formulaStats keeps the old 「数式→値」 part and no formula detail', async () => {
    const real = createMockServices('?mock=1&delay=0.05');
    const st = store();
    const ctrl = createController({
      services: {
        ...real,
        buildPremerged: async (input, p) => {
          const r = await real.buildPremerged(input, p);
          const { formulaStats: _drop, ...rest } = r.report;
          return { ...r, report: { ...rest, formulasToValues: 4 } };
        },
      },
      dispatch: st.dispatch,
      getState: st.getState,
    });
    ctrl.init();
    await ctrl.addFiles([file('学生回答.xlsx'), file('職場回答.xlsx')]);
    await ctrl.idle();
    await ctrl.ensureAnalysis();
    await ctrl.merge();
    expect(st.state.toast!.text).toContain('数式→値 4 件');
    expect(st.state.toast!.detail ?? '').not.toContain('取り込みました');
  });

  it('the reducer keeps formulaStats with the merged state and drops it with 「結合をやり直す」', async () => {
    const { j, snapshot, mergedBook } = await editor();
    const rep = report({ formulaStats: [stat('s2', '職場回答', 3, 1, 0, { externalCell: 1 })] });
    const done = reducer(initialState, { type: 'merge/done', bytes: new Uint8Array(1), layout: { sheet: 'x', insertAt: 1, imported: [], appendAt: null, appended: [] }, report: rep, book: mergedBook, snapshot, baseFileName: 'a.xlsx', format: 'xlsx' });
    expect(done.phase).toBe('merged');
    expect(done.merged!.report.formulaStats).toEqual(rep.formulaStats);
    expect(formulaReportViews(done.merged!.report)[0].sentence).toContain('職場回答: 数式 3 件を数式のまま取り込みました／1 件は値にしました');
    const reset = reducer(done, { type: 'merge/reset' });
    expect(reset.phase).toBe('files');
    expect(reset.merged).toBeNull();
    expect(j.join.rows.length).toBeGreaterThan(0);
  });
});
