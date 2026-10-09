/**
 * Round 7 markup (server-side render): the file card as a drag / click target, the 「他のシートも出力に含める」 checkbox, the
 * single-sheet line of the merge preview.
 */
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';
import { FileCard } from '../../src/components/FileCard';
import { MergePreview } from '../../src/components/MergePreview';
import { KeyControls, OTHER_SHEETS_HINT, OtherSheetsOption } from '../../src/components/SettingsBar';
import { createController } from '../../src/state/controller';
import { createMockServices } from '../../src/state/mock';
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

const cardProps = (st: ReturnType<typeof store>, i: number, patch: Partial<Parameters<typeof FileCard>[0]> = {}) => ({
  index: i,
  file: st.state.files[i],
  isBase: i === 0,
  single: false,
  keyCols: [],
  analyzing: false,
  locked: false,
  compact: false,
  cellText: services.cellText,
  onSetBase: noop,
  onRemove: noop,
  onSheet: noop,
  onHeaderRow: noop,
  onLastRow: noop,
  ...patch,
});

describe('file card as a drag / click target', () => {
  it('is focusable (Alt+← / → moves it), carries its file id, and offers the grip + the gesture hint', async () => {
    const { st } = await env();
    const html = renderToStaticMarkup(<FileCard {...cardProps(st, 1, { gestureHint: 'ドラッグで並べ替え' })} />);
    expect(html).toContain('data-file-id="s2"');
    expect(html).toContain('data-reorderable=""');
    expect(html).toContain('tabindex="0"');
    expect(html).toContain('aria-keyshortcuts="Alt+ArrowLeft Alt+ArrowRight"');
    expect(html).toContain('class="card-grip"');
    expect(html).toContain('ドラッグで並べ替え');
  });

  it('a card that cannot be moved (no columns in the table) has no grip and no shortcut', async () => {
    const { st } = await env();
    const html = renderToStaticMarkup(<FileCard {...cardProps(st, 1, { reorderable: false })} />);
    expect(html).not.toContain('data-reorderable');
    expect(html).not.toContain('card-grip');
    expect(html).not.toContain('aria-keyshortcuts');
  });

  it('the active look is a class + data-active only while the file\'s columns are the selection; the picked look is separate', async () => {
    const { st } = await env();
    const active = renderToStaticMarkup(<FileCard {...cardProps(st, 1, { active: true, compact: true, locked: true })} />);
    expect(active).toMatch(/class="[^"]*\bis-active\b/);
    expect(active).toContain('data-active="true"');
    const plain = renderToStaticMarkup(<FileCard {...cardProps(st, 1, { compact: true, locked: true })} />);
    expect(plain).not.toContain('is-active');
    expect(plain).not.toContain('data-active');
    const picked = renderToStaticMarkup(<FileCard {...cardProps(st, 1, { picked: true })} />);
    expect(picked).toMatch(/class="[^"]*\bis-picked\b/);
    expect(picked).not.toContain('is-active');
  });

  it('the controls of a card are still there (base toggle, sheet / header row selects, クリア) in the files phase', async () => {
    const { st } = await env();
    const html = renderToStaticMarkup(<FileCard {...cardProps(st, 1)} />);
    for (const id of ['base-toggle-1', 'sheet-select-1', 'header-row-select-1', 'file-clear-1']) expect(html).toContain(`data-testid="${id}"`);
  });
});

describe('「他のシートも出力に含める」', () => {
  it('is an unchecked checkbox with the explanatory tooltip; disabled (with the hint to redo the merge) after the merge', () => {
    const html = renderToStaticMarkup(<OtherSheetsOption checked={false} onChange={noop} />);
    expect(html).toContain('data-testid="include-other-sheets"');
    expect(html).toContain('他のシートも出力に含める');
    expect(html).not.toContain('checked');
    expect(html).not.toContain('disabled');
    expect(html).toContain('通常は、結合した1つのシートだけを出力します。基準ファイルにあるほかのシート（集計・グラフなど）も残したいときにチェックします');
    expect(OTHER_SHEETS_HINT).toContain('通常は、結合した1つのシートだけを出力します');
    const locked = renderToStaticMarkup(<OtherSheetsOption checked disabled />);
    expect(locked).toContain('checked');
    expect(locked).toContain('disabled');
    expect(locked).toContain('「結合をやり直す」');
  });

  it('sits next to 結合方法 in the settings bar', async () => {
    const { st } = await env();
    const s = st.state;
    const html = renderToStaticMarkup(
      <KeyControls analysis={s.analysis!} plan={s.plan!} labels={{}} onPlan={noop} onToggleFamily={noop} onOpenDetail={noop} onJoinType={noop} includeOtherSheets={false} onIncludeOtherSheets={noop} />,
    );
    const join = html.indexOf('data-testid="join-type-select"');
    const other = html.indexOf('data-testid="include-other-sheets"');
    expect(join).toBeGreaterThan(0);
    expect(other).toBeGreaterThan(join);
  });
});

describe('merge preview: single-sheet line', () => {
  it('says which sheets of the base file are left out (default), or that they are kept', async () => {
    const { st } = await env();
    const s = st.state;
    const props = { files: s.files, analysis: s.analysis, plan: s.plan, joinResult: s.joinResult, overrides: {}, keyError: null, analyzing: false, onImports: noop, onOpenDetail: noop };
    const single = renderToStaticMarkup(<MergePreview {...props} />);
    expect(single).toContain('data-testid="preview-sheets"');
    expect(single).toContain('出力は結合した 1 シートのみです（他のシート 1 枚「集計」 は含めません）');
    const keep = renderToStaticMarkup(<MergePreview {...props} includeOtherSheets />);
    expect(keep).toContain('基準ファイルの他のシート 1 枚（「集計」）も、そのまま出力に含めます');
    expect(keep).not.toContain('出力は結合した 1 シートのみです');
  });
});
