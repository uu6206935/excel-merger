/**
 * Round 9 (plan A) markup (server-side render): the second checkbox, the planned sheet list of the merge preview, the 「シート一覧」
 * of the side panel, the Excel本体 hint of the export dialog.
 */
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';
import type { ServerStatus } from '../../src/api';
import { ExportDialog, type ExportDialogProps } from '../../src/components/ExportDialog';
import { MergePreview } from '../../src/components/MergePreview';
import { PlannedSheets, SheetListSection, SHORTENED_BADGE, SUFFIX_BADGE } from '../../src/components/OutputSheets';
import { KeyControls, OtherSheetsOption, SheetOptions, SOURCE_SHEETS_HINT, SOURCE_SHEETS_LABEL, SourceSheetsOption } from '../../src/components/SettingsBar';
import { SidePanel } from '../../src/components/SidePanel';
import { createController } from '../../src/state/controller';
import { createMockServices } from '../../src/state/mock';
import { plannedOutputSheets, type OutputSheetRow, type PlannedOutput } from '../../src/state/sheetPlan';
import type { ServerState } from '../../src/state/types';
import { file, store } from './fixtures';

const services = createMockServices('?mock=1&delay=0.05');
const noop = () => undefined;

async function env() {
  const st = store();
  const ctrl = createController({ services, dispatch: st.dispatch, getState: st.getState });
  ctrl.init();
  await ctrl.addFiles([file('学生回答.xlsx'), file('職場回答.xlsx'), file('人事回答.xlsx')]);
  await ctrl.idle();
  await ctrl.ensureAnalysis();
  return { st, ctrl };
}

describe('「他のファイルのシートも含める」 checkbox', () => {
  it('has the id, the label and the tooltip of the request; unchecked and enabled by default', () => {
    const html = renderToStaticMarkup(<SourceSheetsOption checked={false} onChange={noop} />);
    expect(html).toContain('data-testid="include-source-sheets"');
    expect(html).toContain('>他のファイルのシートも含める<');
    expect(SOURCE_SHEETS_LABEL).toBe('他のファイルのシートも含める');
    const tip = '2つ目以降のファイルの元のシートを、すべてそのまま入れます。シート名は「ファイル名_シート名」になり、名前が重なるときは(1)(2)…が付きます。元のシートを参照している数式もエラーにならずに動きます';
    expect(SOURCE_SHEETS_HINT).toBe(tip);
    expect(html).toContain(`title="${tip.replace(/"/g, '&quot;')}"`);
    expect(html).not.toMatch(/\schecked/);
    expect(html).not.toMatch(/\sdisabled/);
  });

  it('after the merge it is disabled and the tooltip tells how to change it', () => {
    const html = renderToStaticMarkup(<SourceSheetsOption checked disabled />);
    expect(html).toMatch(/\schecked/);
    expect(html).toMatch(/\sdisabled/);
    expect(html).toContain('結合をやり直す');
    expect(html).toContain('is-disabled');
  });

  it('is indented under 「他のシートも出力に含める」 (same group, second in the DOM)', () => {
    const html = renderToStaticMarkup(<SheetOptions includeOtherSheets includeSourceSheets={false} onIncludeOtherSheets={noop} onIncludeSourceSheets={noop} />);
    const first = html.indexOf('data-testid="include-other-sheets"');
    const second = html.indexOf('data-testid="include-source-sheets"');
    expect(first).toBeGreaterThan(0);
    expect(second).toBeGreaterThan(first);
    expect(html).toContain('settings-sheets');
    expect(html).toContain('settings-check-sub');
    // only the second one is the indented one
    expect(renderToStaticMarkup(<OtherSheetsOption checked={false} />)).not.toContain('settings-check-sub');
    // checked state comes from the props, separately
    const mixed = renderToStaticMarkup(<SheetOptions includeOtherSheets includeSourceSheets={false} />);
    expect(/<input[^>]*include-other-sheets[^>]*>/.exec(mixed)?.[0]).toMatch(/\schecked/);
    expect(/<input[^>]*include-source-sheets[^>]*>/.exec(mixed)?.[0]).not.toMatch(/\schecked/);
  });

  it('sits right after 「他のシートも出力に含める」 in the settings bar', async () => {
    const { st } = await env();
    const s = st.state;
    const html = renderToStaticMarkup(
      <KeyControls
        analysis={s.analysis!}
        plan={s.plan!}
        labels={{}}
        onPlan={noop}
        onToggleFamily={noop}
        onOpenDetail={noop}
        onJoinType={noop}
        includeOtherSheets
        onIncludeOtherSheets={noop}
        includeSourceSheets
        onIncludeSourceSheets={noop}
      />,
    );
    const join = html.indexOf('data-testid="join-type-select"');
    const other = html.indexOf('data-testid="include-other-sheets"');
    const source = html.indexOf('data-testid="include-source-sheets"');
    expect(join).toBeGreaterThan(0);
    expect(other).toBeGreaterThan(join);
    expect(source).toBeGreaterThan(other);
  });
});

describe('merge preview: 出力されるシート', () => {
  async function previewProps() {
    const { st, ctrl } = await env();
    const s = st.state;
    return {
      ctrl,
      props: { files: s.files, analysis: s.analysis, plan: s.plan, joinResult: s.joinResult, overrides: {}, keyError: null, analyzing: false, onImports: noop, onOpenDetail: noop },
    };
  }

  it('is shown only when 「他のファイルのシートも含める」 is on; then it replaces the one-sheet line', async () => {
    const { props } = await previewProps();
    const off = renderToStaticMarkup(<MergePreview {...props} />);
    expect(off).not.toContain('data-testid="preview-output-sheets"');
    expect(off).toContain('data-testid="preview-sheets"');
    const baseOnly = renderToStaticMarkup(<MergePreview {...props} includeOtherSheets />);
    expect(baseOnly).not.toContain('data-testid="preview-output-sheets"');
    const on = renderToStaticMarkup(<MergePreview {...props} includeOtherSheets includeSourceSheets />);
    expect(on).toContain('data-testid="preview-output-sheets"');
    expect(on).not.toContain('data-testid="preview-sheets"');
    expect(on).toContain('出力されるシート');
  });

  it('lists 1. the merged sheet (結合した表), the base sheets (基準ファイル), then <new name> ← 「<file>」の「<sheet>」', async () => {
    const { props } = await previewProps();
    const html = renderToStaticMarkup(<MergePreview {...props} includeOtherSheets includeSourceSheets />);
    const rows = [...html.matchAll(/<li class="sheet-plan-item[^"]*" data-testid="planned-sheet-(\d+)" data-kind="(\w+)" data-name="([^"]*)"/g)].map((m) => [m[1], m[2], m[3]]);
    expect(rows).toEqual([
      ['0', 'merged', '応募者一覧'],
      ['1', 'base', '集計'],
      ['2', 'copy', '職場回答_回答'],
      ['3', 'copy', '人事回答_Sheet1'],
    ]);
    expect(html).toContain('結合した表');
    expect(html).toContain('基準ファイル');
    expect(html).toContain('← 「職場回答.xlsx」の「回答」');
    expect(html).toContain('← 「人事回答.xlsx」の「Sheet1」');
    expect(html).toContain('data-sheet-count="4"');
  });

  it('follows the order of the cards', async () => {
    const { ctrl, props } = await previewProps();
    expect(ctrl.reorderFile('s3', 's2')).toBe(true);
    const html = renderToStaticMarkup(<MergePreview {...props} files={[props.files[0], props.files[2], props.files[1]]} includeOtherSheets includeSourceSheets />);
    const names = [...html.matchAll(/data-testid="planned-sheet-\d+" data-kind="\w+" data-name="([^"]*)"/g)].map((m) => m[1]);
    expect(names).toEqual(['応募者一覧', '集計', '人事回答_Sheet1', '職場回答_回答']);
  });

  it('a file without a key is listed as left out', async () => {
    const { props } = await previewProps();
    const html = renderToStaticMarkup(<MergePreview {...props} excluded={['s2']} includeOtherSheets includeSourceSheets />);
    expect(html).toContain('data-testid="planned-sheets-skipped"');
    expect(html).toContain('「職場回答.xlsx」のシートは入りません');
    expect(html).not.toContain('data-name="職場回答_回答"');
  });
});

describe('PlannedSheets', () => {
  const plan = (): PlannedOutput => ({
    sheets: [
      { kind: 'merged', name: '名簿', sourceId: 's1' },
      { kind: 'copy', name: '名簿_名簿', sourceId: 's2', file: '名簿.xlsx', original: '名簿', suffixed: false, shortened: false },
      { kind: 'copy', name: '名簿_名簿(1)', sourceId: 's3', file: '名簿.xlsx', original: '名簿', suffixed: true, suffixNo: 1, shortened: false },
      { kind: 'copy', name: '長い長いファイル_回答の一覧とまとめ', sourceId: 's3', file: '長い長いファイル名.xlsx', original: '回答の一覧とまとめ', suffixed: false, shortened: true, hidden: true },
    ],
    skipped: [],
  });
  const colors = { s1: '#111111', s2: '#222222', s3: '#333333' };

  it('badge 「名前が重なったため(1)を付けました」 only on the sheet that got the suffix', () => {
    const html = renderToStaticMarkup(<PlannedSheets plan={plan()} colors={colors} />);
    expect(SUFFIX_BADGE(1)).toBe('名前が重なったため(1)を付けました');
    expect(SUFFIX_BADGE(2)).toBe('名前が重なったため(2)を付けました');
    expect(html.match(/名前が重なったため\(1\)を付けました/g)).toHaveLength(1);
    expect(html).toContain('data-testid="planned-sheet-suffixed-2"');
    expect(html).not.toContain('data-testid="planned-sheet-suffixed-1"');
  });

  it('hint 「31文字までに短縮」 only on the shortened name; hidden sheets are marked', () => {
    const html = renderToStaticMarkup(<PlannedSheets plan={plan()} colors={colors} />);
    expect(SHORTENED_BADGE).toBe('31文字までに短縮');
    expect(html.match(/31文字までに短縮/g)).toHaveLength(1);
    expect(html).toContain('data-testid="planned-sheet-shortened-3"');
    expect(html).not.toContain('data-testid="planned-sheet-shortened-2"');
    expect(html).toContain('非表示');
  });

  it('numbers the rows from 1 and carries the file colour', () => {
    const html = renderToStaticMarkup(<PlannedSheets plan={plan()} colors={colors} />);
    expect([...html.matchAll(/<span class="sheet-no">(\d+)<\/span>/g)].map((m) => m[1])).toEqual(['1', '2', '3', '4']);
    expect(html).toContain('background:#222222');
  });

  it('renders nothing for an empty plan', () => {
    expect(renderToStaticMarkup(<PlannedSheets plan={{ sheets: [], skipped: [] }} colors={{}} />)).toBe('');
  });

  it('works with the real planner output', () => {
    const p = plannedOutputSheets([], null, [], { includeOtherSheets: true, includeSourceSheets: true });
    expect(renderToStaticMarkup(<PlannedSheets plan={p} colors={{}} />)).toBe('');
  });
});

describe('SheetListSection (side panel, after the merge)', () => {
  const rows: OutputSheetRow[] = [
    { kind: 'merged', name: '応募者一覧', hidden: false },
    { kind: 'base', name: '集計', hidden: false },
    { kind: 'copy', name: '職場回答_回答', hidden: false, sourceId: 's2', file: '職場回答.xlsx', original: '回答', complete: true },
    { kind: 'copy', name: '学生回答_集計', hidden: true, sourceId: 's3', file: '学生回答.xlsx', original: '集計', complete: false },
  ];

  it('is a collapsed <details> with the names, their origin and the count', () => {
    const html = renderToStaticMarkup(<SheetListSection rows={rows} colors={{ s2: '#2' }} />);
    expect(html).toContain('<details');
    expect(html).not.toMatch(/<details[^>]*\sopen/);
    expect(html).toContain('シート一覧');
    expect(html).toContain('count-badge">4<');
    expect([...html.matchAll(/data-testid="sheet-list-item-\d+"/g)]).toHaveLength(4);
    for (const r of rows) expect(html).toContain(`data-name="${r.name}"`);
    expect(html).toContain('結合した表');
    expect(html).toContain('基準ファイル');
    expect(html).toContain('コピー');
  });

  it('marks hidden sheets and copies that are not complete (the 警告 list has the details)', () => {
    const html = renderToStaticMarkup(<SheetListSection rows={rows} colors={{}} />);
    expect(html).toContain('非表示');
    expect(html.match(/一部のみ/g)).toHaveLength(1);
  });

  it('is not shown for a one-sheet workbook', () => {
    expect(renderToStaticMarkup(<SheetListSection rows={rows.slice(0, 1)} colors={{}} />)).toBe('');
  });

  it('sits between the summary and the rest of the side panel', () => {
    const html = renderToStaticMarkup(
      <SidePanel
        phase="merged"
        primaryDisabled={false}
        onMerge={noop}
        onExport={noop}
        exportRunning={false}
        canUndo={false}
        canRedo={false}
        onUndo={noop}
        onRedo={noop}
        onRemerge={noop}
        summary={<span id="sum">S</span>}
        sheets={<SheetListSection rows={rows} colors={{}} />}
        checks={<span id="chk">C</span>}
        columns={<span />}
        columnCount={0}
        columnHelp=""
        notes={[]}
      />,
    );
    expect(html.indexOf('id="sum"')).toBeLessThan(html.indexOf('data-testid="sheet-list-details"'));
    expect(html.indexOf('data-testid="sheet-list-details"')).toBeLessThan(html.indexOf('id="chk"'));
  });
});

describe('export dialog: hint for the copied sheets', () => {
  const okStatus: ServerStatus = { app: 'excel-merger', version: '1.0.0', excel: { installed: true, version: '16.0' }, outputDir: 'D:\\出力', diskLow: false, diskFreeMB: 20000, outputDiskLow: false };
  const server: ServerState = { checked: true, status: okStatus, connected: true };
  const dialog = (patch: Partial<ExportDialogProps> = {}) =>
    renderToStaticMarkup(
      <ExportDialog
        defaultName="学生回答_統合_20261010.xlsx"
        format="xlsx"
        server={server}
        exportState={{ phase: 'idle', engine: null, message: '', ratio: null }}
        editCount={0}
        onExport={noop}
        onCancelExport={noop}
        onOpenFile={noop}
        onOpenFolder={noop}
        onDownloadAgain={noop}
        onClose={noop}
        {...patch}
      />,
    );

  it('says that charts and pivot tables of the copied sheets are kept best by Excel本体 — only when the sheets are copied', () => {
    const on = dialog({ includeSourceSheets: true });
    expect(on).toContain('data-testid="export-engine-sheets-hint"');
    expect(on).toContain('グラフ・ピボットテーブルは、この方法がもっとも確実に保たれます');
    expect(dialog()).not.toContain('export-engine-sheets-hint');
    expect(dialog({ includeSourceSheets: false })).not.toContain('export-engine-sheets-hint');
  });

  it('the file name suggestion and the rest of the dialog stay', () => {
    const on = dialog({ includeSourceSheets: true });
    expect(on).toContain('value="学生回答_統合_20261010.xlsx"');
    expect(on).toContain('Excel本体で出力（推奨・最も忠実）');
    expect(on).toContain('内蔵エンジンで出力');
  });
});
