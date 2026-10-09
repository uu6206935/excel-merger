/**
 * Markup (server-side render) of the formula import report: the 「数式の取り込み」 block, the hint under 「追加される列」 of the merge
 * preview, the 数式→値 group of the check panel.
 */
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';
import { CheckPanel } from '../../src/components/CheckPanel';
import { FormulaReport } from '../../src/components/FormulaReport';
import { MergePreview } from '../../src/components/MergePreview';
import type { MergeReport, SourceBook } from '../../src/core/types';
import { createController } from '../../src/state/controller';
import { formulaReportViews, PREVIEW_FORMULA_HINT } from '../../src/state/formulaReport';
import { checkGroups, importChoices, importKey } from '../../src/state/keysPlan';
import { createMockServices } from '../../src/state/mock';
import { file, store } from './fixtures';

const noop = () => undefined;
const nameOf = (sid: string, label: string) => ({ s2: '02_面接評価.xlsx' })[sid as 's2'] ?? label;
const rep = (p: Partial<MergeReport>): MergeReport => ({ mergedRows: 60, matchedRows: 60, baseOnlyRows: 0, appendedRows: 0, importedColumns: 3, formulasKept: 0, formulasToValues: 0, items: [], ...p });

describe('FormulaReport block', () => {
  it('everything kept: information only (no amber), the kept line, no hint', () => {
    const views = formulaReportViews(rep({ formulaStats: [{ sourceId: 's2', label: '02_面接評価', kept: 116, toValues: 0, empty: 0, reasons: {} }] }), nameOf);
    const html = renderToStaticMarkup(<FormulaReport views={views} colors={{ s2: '#15803d' }} canJump onJump={noop} />);
    expect(html).toContain('data-testid="formula-report"');
    expect(html).toContain('data-tone="ok"');
    expect(html).toContain('数式の取り込み');
    expect(html).toContain('02_面接評価.xlsx');
    expect(html).toContain('数式 116 件を数式のまま取り込みました');
    expect(html).not.toContain('空白');
    expect(html).not.toContain('formula-hint');
    expect(html).not.toContain('badge-warn');
    expect(html).not.toContain('formula-details');
  });

  it('empty cells: amber block + badge, the three lines, the hint, a collapsible with reasons and clickable example rows', () => {
    const views = formulaReportViews(
      rep({
        formulaStats: [{ sourceId: 's2', label: '02_面接評価', kept: 147, toValues: 12, empty: 3, reasons: { otherSheet: 11, externalCell: 4 } }],
        items: [
          { severity: 'warning', code: 'formulasToValues', sourceId: 's2', message: '「02_面接評価」の他のシートを参照している数式 8 件は値にしました。', rows: [7, 8, 9, 10, 11, 12, 13] },
          { severity: 'warning', code: 'formulasNoCache', sourceId: 's2', message: '「02_面接評価」の数式 3 件は、元のファイルに計算結果が保存されていない数式です', rows: [20] },
        ],
      }),
      nameOf,
    );
    const html = renderToStaticMarkup(<FormulaReport views={views} canJump onJump={noop} />);
    expect(html).toContain('data-tone="warn"');
    expect(html).toContain('badge-warn');
    expect(html).toContain('数式 147 件を数式のまま取り込みました');
    expect(html).toContain('12 件は値にしました');
    expect(html).toContain('空白 3 件（元のファイルに計算結果がありません）');
    expect(html).toContain('元のExcelで一度開いて保存し直すと取り込めることがあります');
    expect(html).toContain('<details');
    expect(html).toContain('詳しく見る');
    expect(html).toContain('他のシートを参照 11 件');
    expect(html).toContain('表の外のセル 4 件');
    // example rows: the first five as buttons, the rest counted
    expect((html.match(/class="formula-row-btn"/g) ?? []).length).toBe(6);
    expect(html).toContain('data-row="7"');
    expect(html).toContain('ほか 2 件');
    expect(html).toContain('data-row="20"');
  });

  it('values only: a quiet note tone, the breakdown inline', () => {
    const views = formulaReportViews(rep({ formulaStats: [{ sourceId: 's2', label: '02_面接評価', kept: 100, toValues: 12, empty: 0, reasons: { otherSheet: 8, externalCell: 4 } }] }), nameOf);
    const html = renderToStaticMarkup(<FormulaReport views={views} canJump onJump={noop} />);
    expect(html).toContain('data-tone="note"');
    expect(html).toContain('12 件は値にしました（内訳: 他のシートを参照 8、表の外のセル 4）');
    expect(html).not.toContain('badge-warn');
    expect(html).not.toContain('formula-hint');
  });

  it('shows nothing without views; the rows are not buttons when the grid is not shown', () => {
    expect(renderToStaticMarkup(<FormulaReport views={[]} canJump onJump={noop} />)).toBe('');
    const views = formulaReportViews(rep({ formulaStats: [{ sourceId: 's2', label: 'A', kept: 1, toValues: 1, empty: 0, reasons: { otherSheet: 1 } }], items: [{ severity: 'warning', code: 'formulasToValues', sourceId: 's2', message: 'm', rows: [7] }] }));
    const html = renderToStaticMarkup(<FormulaReport views={views} canJump={false} onJump={noop} />);
    expect(html).toContain('is-static');
    expect(html).not.toContain('data-row=');
  });

  it('the recalc line (kept formulas without a stored result) is shown without opening the details', () => {
    const views = formulaReportViews(
      rep({
        formulaStats: [{ sourceId: 's2', label: 'A', kept: 10, toValues: 0, empty: 0, reasons: {} }],
        items: [{ severity: 'info', code: 'formulasRecalc', sourceId: 's2', message: 'recalc', rows: [4] }],
      }),
    );
    const html = renderToStaticMarkup(<FormulaReport views={views} canJump onJump={noop} />);
    expect(html).toContain('Excel で開いたときに表示されます');
    expect(html).toContain('data-testid="formula-recalc-s2"');
  });
});

describe('check panel: 数式→値', () => {
  it('lists the example rows (clickable) and explains the group in the new words', () => {
    const rows = [{ key: 'a', row: 7, text: '7行目', sub: '02_面接評価.xlsx: 値にした数式', tone: 'info' as const }];
    const groups = [{ id: 'toValues' as const, title: '数式→値', count: 12, rows }];
    const html = renderToStaticMarkup(<CheckPanel groups={groups} canJump onJump={noop} emptyText="" />);
    expect(html).toContain('数式→値');
    expect(html).toContain('空白になったものを含みます');
    expect(html).not.toContain('参照先の列が取り込まれていない');
  });

  it('checkGroups keeps the rows it is given', () => {
    const g = checkGroups({ rows: [], reports: {}, summary: { baseRows: 0, matchedAny: 0, matchedAll: 0, appended: 0 }, items: [], keyFill: {}, defaultImports: [] }, {}, { count: 2, rows: [{ key: 'a', row: 3, text: '3行目', tone: 'info' }] });
    expect(g.find((x) => x.id === 'toValues')?.rows[0].row).toBe(3);
  });
});

describe('merge preview: the hint under 「追加される列」', () => {
  async function env() {
    const st = store();
    const ctrl = createController({ services: createMockServices('?mock=1&delay=0.05'), dispatch: st.dispatch, getState: st.getState });
    ctrl.init();
    await ctrl.addFiles([file('学生回答.xlsx'), file('職場回答.xlsx'), file('人事回答.xlsx')]);
    await ctrl.idle();
    await ctrl.ensureAnalysis();
    return st;
  }

  const props = (st: Awaited<ReturnType<typeof env>>, files = st.state.files, overrides: Record<string, boolean> = {}) => ({
    files,
    analysis: st.state.analysis,
    plan: st.state.plan,
    joinResult: st.state.joinResult,
    overrides,
    keyError: null,
    analyzing: false,
    onImports: noop,
    onOpenDetail: noop,
  });

  /** the file with one formula in the first data row of an imported column (a copy; the original book stays as it is) */
  function withFormula(st: Awaited<ReturnType<typeof env>>, id: string, col: number) {
    return st.state.files.map((f) => {
      if (f.id !== id || !f.book || !f.region) return f;
      const sheet = f.book.sheets.find((s) => s.name === f.region!.sheet)!;
      const rows = sheet.rows.map((r) => (r ? r.slice() : r));
      const r0 = f.region.firstDataRow - 1;
      const row = rows[r0] ?? [];
      row[col - 1] = { t: 'n', v: 5, f: 'A1+1' };
      rows[r0] = row;
      const book: SourceBook = { ...f.book, sheets: f.book.sheets.map((s) => (s === sheet ? { ...sheet, rows } : s)) };
      return { ...f, book };
    });
  }

  it('is shown when an imported column holds formulas (the count per file is in its tooltip), not otherwise', async () => {
    const st = await env();
    expect(renderToStaticMarkup(<MergePreview {...props(st)} />)).not.toContain('preview-formulas-hint');
    const col = importChoices(st.state.analysis!, st.state.plan!, st.state.joinResult, {}, 's2').find((c) => c.on)!.col;
    const html = renderToStaticMarkup(<MergePreview {...props(st, withFormula(st, 's2', col))} />);
    expect(html).toContain('data-testid="preview-formulas-hint"');
    expect(html).toContain(PREVIEW_FORMULA_HINT);
    expect(html).toContain('数式を含む列があります。結合後も数式のまま残るよう付け替えます');
    expect(html).toContain('data-formula-columns="s2:1"');
    expect(html).toContain('職場回答.xlsx: 数式を含む列 1 列');
    // the hint goes with the columns that will be imported: switching that column off removes it
    const off = renderToStaticMarkup(<MergePreview {...props(st, withFormula(st, 's2', col), { [importKey('s2', st.state.analysis!.pairs.s2.other.sheet, col)]: false })} />);
    expect(off).not.toContain('preview-formulas-hint');
    expect(off).not.toContain('data-formula-columns');
  });
});
