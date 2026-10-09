/**
 * Markup of the round-4 UI pieces (server-side render, no DOM needed): the 結合対象外 badge of a file card, the merge
 * preview notice, the key detail modal (banner, match rate sentence, manual candidates with ×, the picker).
 */
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';
import { FileCard } from '../../src/components/FileCard';
import { KeyDetailModal, type KeyDetailModalProps } from '../../src/components/KeyDetailModal';
import { ManualKeyPicker } from '../../src/components/ManualKeyPicker';
import { MergePreview } from '../../src/components/MergePreview';
import { createController } from '../../src/state/controller';
import { DEFAULT_OPTIONS } from '../../src/state/helpers';
import { excludedSourceIds } from '../../src/state/keysPlan';
import { createMockServices } from '../../src/state/mock';
import { file, store } from './fixtures';

const services = createMockServices('?mock=1&delay=0.05');
const noop = () => undefined;

async function env() {
  const st = store();
  const ctrl = createController({ services, dispatch: st.dispatch, getState: st.getState });
  ctrl.init();
  await ctrl.addFiles([file('学生回答.xlsx'), file('職場回答.xlsx'), file('参加状況（キーなし）.xlsx')]);
  await ctrl.idle();
  await ctrl.ensureAnalysis();
  return { st, ctrl };
}

function modalProps(st: ReturnType<typeof store>, patch: Partial<KeyDetailModalProps> = {}): KeyDetailModalProps {
  const s = st.state;
  return {
    analysis: s.analysis!,
    plan: s.plan!,
    joinResult: s.joinResult,
    labels: { s1: '学生回答', s2: '職場回答', s3: '参加状況（キーなし）' },
    colors: { s1: '#2563eb', s2: '#15803d', s3: '#c2410c' },
    options: DEFAULT_OPTIONS,
    overrides: {},
    initialTab: null,
    excluded: excludedSourceIds(s.analysis, s.plan),
    manualKeys: s.manualKeys,
    onPlan: noop,
    onOptions: noop,
    onImports: noop,
    onAddManual: () => ({ ok: true, id: 'x', existing: false }),
    onRemoveManual: noop,
    onClose: noop,
    ...patch,
  };
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

describe('file card', () => {
  it('shows 結合対象外（結合キーが見つかりません） and 「キーを選ぶ…」 for a file without a key', async () => {
    const { st } = await env();
    const html = renderToStaticMarkup(<FileCard {...cardProps(st, 2, { excluded: true, onPickKey: noop })} />);
    expect(html).toContain('結合対象外（結合キーが見つかりません）');
    expect(html).toContain('data-testid="file-pick-key-2"');
    expect(html).toContain('キーを選ぶ…');
    expect(html).toContain('is-excluded');
    expect(html).not.toContain('data-testid="file-keys-2"');
    // the plain line 「結合キーなし」 is replaced by the badge
    expect(html).not.toContain('結合キーなし');
  });

  it('after the merge the badge stays but the button is gone; a file with a key / the base never shows it', async () => {
    const { st } = await env();
    const locked = renderToStaticMarkup(<FileCard {...cardProps(st, 2, { excluded: true, onPickKey: noop, locked: true, compact: true })} />);
    expect(locked).toContain('結合対象外（結合キーが見つかりません）');
    expect(locked).not.toContain('file-pick-key-2');
    const ok = renderToStaticMarkup(<FileCard {...cardProps(st, 1, { excluded: false, onPickKey: noop })} />);
    expect(ok).not.toContain('結合対象外');
    expect(ok).toContain('data-testid="file-keys-1"');
    const base = renderToStaticMarkup(<FileCard {...cardProps(st, 0, { excluded: true, onPickKey: noop })} />);
    expect(base).not.toContain('結合対象外');
  });
});

describe('merge preview', () => {
  it('says which file is out of the merge, offers the key picker, and lists no columns for it', async () => {
    const { st } = await env();
    const s = st.state;
    const html = renderToStaticMarkup(
      <MergePreview files={s.files} analysis={s.analysis} plan={s.plan} joinResult={s.joinResult} overrides={{}} keyError={null} analyzing={false} excluded={['s3']} onImports={noop} onOpenDetail={noop} />,
    );
    expect(html).toContain('data-testid="excluded-notice"');
    expect(html).toContain('結合対象外: 「参加状況（キーなし）.xlsx」');
    expect(html).toContain('他のファイルだけで結合できます');
    expect(html).toContain('data-testid="excluded-pick-key-s3"');
    expect(html).toContain('data-testid="preview-excluded-s3"');
    expect(html).toContain('結合キーが選ばれるまで、列は追加されません');
    expect(html).not.toContain('どのファイルにも結合キーが選ばれていません');
  });

  it('with no key at all the notice says so', async () => {
    const { st, ctrl } = await env();
    ctrl.updatePlan({ ...st.state.plan!, tiers: { s2: [], s3: [] } });
    const s = st.state;
    const html = renderToStaticMarkup(
      <MergePreview files={s.files} analysis={s.analysis} plan={s.plan} joinResult={s.joinResult} overrides={{}} keyError={null} analyzing={false} onImports={noop} onOpenDetail={noop} />,
    );
    expect(html).toContain('どのファイルにも結合キーが選ばれていません');
    expect(html).toContain('追加される列 <span class="count-badge">0</span>');
  });

  it('nothing excluded: no notice', async () => {
    const { st } = await env();
    const s = st.state;
    const html = renderToStaticMarkup(
      <MergePreview files={s.files} analysis={s.analysis} plan={{ ...s.plan!, tiers: { ...s.plan!.tiers, s3: ['x'] } }} joinResult={s.joinResult} overrides={{}} keyError={null} analyzing={false} excluded={[]} onImports={noop} onOpenDetail={noop} />,
    );
    expect(html).not.toContain('excluded-notice');
  });
});

describe('key detail modal', () => {
  it('the tab of the excluded file: banner, tag, empty-state sentence, the picker', async () => {
    const { st } = await env();
    const html = renderToStaticMarkup(<KeyDetailModal {...modalProps(st, { initialTab: 's3' })} />);
    expect(html).toContain('data-testid="detail-excluded"');
    expect(html).toContain('このファイルは今、結合の対象外です');
    expect(html).toContain('結合対象外</span>'); // the tag on the tab
    expect(html).toContain('data-testid="match-rate-empty"');
    // no candidate at all: the texts point to the picker, not to a 使う checkbox that does not exist
    expect(html).toContain('結合キーが選ばれていません。下の「列を選んで結合キーに追加」でキーを選ぶと');
    expect(html).not.toContain('「使う」にチェックを入れ');
    expect(html).toContain('キーになりそうな列が自動では見つかりませんでした');
    expect(html).toContain('data-testid="manual-key-picker"');
    expect(html).toContain('列を選んで結合キーに追加');
    expect(html).toContain('基準ファイルの列');
    expect(html).toContain('このファイルの列');
    expect(html).toContain('＋列を追加');
    expect(html).toContain('氏名（C列）'); // every column by header text with its letter
    expect(html).toContain('参加者（B列）');
  });

  it('an excluded file WITH candidates (all unchecked): the texts mention the 使う checkbox', async () => {
    const { st } = await env();
    const plan = { ...st.state.plan!, tiers: { ...st.state.plan!.tiers, s2: [] } };
    const html = renderToStaticMarkup(<KeyDetailModal {...modalProps(st, { initialTab: 's2', plan, joinResult: null, excluded: ['s2', 's3'] })} />);
    expect(html).toContain('data-testid="detail-excluded"');
    expect(html).toContain('下の表の「使う」にチェックを入れるか、「列を選んで結合キーに追加」');
    expect(html).toContain('結合キーが選ばれていません。「使う」にチェックを入れると');
  });

  it('the tab of a file with keys: the Power Query sentence under the table, live numbers', async () => {
    const { st } = await env();
    const html = renderToStaticMarkup(<KeyDetailModal {...modalProps(st, { initialTab: 's2' })} />);
    expect(html).not.toContain('data-testid="detail-excluded"');
    const m = /このキーで (\d+) \/ (\d+) 行が一致します（基準 (\d+) 行中 (\d+) 行が見つかりました）/.exec(html);
    expect(m).not.toBeNull();
    const rep = st.state.joinResult!.reports.s2;
    expect(Number(m![1])).toBe(30 - rep.unmatchedOther.length);
    expect(Number(m![2])).toBe(30);
    expect(Number(m![3])).toBe(30);
    expect(Number(m![4])).toBe(30 - rep.unmatchedBase.length);
    // after the table (the table comes first in the markup)
    expect(html.indexOf('data-testid="key-candidates"')).toBeLessThan(html.indexOf('data-testid="match-rate-text"'));
    expect(html.indexOf('data-testid="match-rate-text"')).toBeLessThan(html.indexOf('data-testid="manual-key-picker"'));
  });

  it('manual candidates carry the 手動 tag and a × button; automatic ones have no ×', async () => {
    const { st, ctrl } = await env();
    expect(ctrl.addManualKey('s3', [3], [2]).ok).toBe(true);
    const html = renderToStaticMarkup(<KeyDetailModal {...modalProps(st, { initialTab: 's3' })} />);
    expect(html).toContain('氏名 ⟷ 参加者');
    expect(html).toContain('data-testid="key-remove-manual:3|2"');
    expect(html).toContain('<span class="tag tag-new">手動</span>');
    expect(html).not.toContain('detail-excluded"'); // a key is selected: no longer excluded
    expect(html).toContain('このキーで 20 / 20 行が一致します');
    const s2 = renderToStaticMarkup(<KeyDetailModal {...modalProps(st, { initialTab: 's2' })} />);
    expect(s2).not.toContain('data-testid="key-remove-');
    // without the remove handler there is no ×
    expect(renderToStaticMarkup(<KeyDetailModal {...modalProps(st, { initialTab: 's3', onRemoveManual: undefined })} />)).not.toContain('data-testid="key-remove-');
  });
});

describe('manual key picker', () => {
  it('starts with one dropdown per side, a hint and a disabled 追加', async () => {
    const { st } = await env();
    const html = renderToStaticMarkup(<ManualKeyPicker sourceId="s3" analysis={st.state.analysis!} baseLabel="学生回答" otherLabel="参加状況" onAdd={() => ({ ok: true, id: 'x', existing: false })} />);
    expect(html).toContain('data-testid="manual-base-select-0"');
    expect(html).toContain('data-testid="manual-other-select-0"');
    expect(html).not.toContain('manual-base-select-1');
    expect(html).toContain('>選んでください</option>');
    expect(html).toContain('data-testid="manual-base-add-column"');
    expect(html).toContain('両方のファイルの列を選ぶと、検出した列の種類をここに表示します');
    expect(/<button[^>]*data-testid="manual-key-add"[^>]*disabled/.test(html)).toBe(true);
    // unknown source: nothing rendered
    expect(renderToStaticMarkup(<ManualKeyPicker sourceId="zz" analysis={st.state.analysis!} baseLabel="a" otherLabel="b" onAdd={() => ({ ok: false, error: 'x' })} />)).toBe('');
  });
});
