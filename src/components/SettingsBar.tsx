/**
 * Settings bar under the file cards: 結合キー chips (ID / 氏名 always first), extra key chips, 「＋ 他の候補」,
 * 「詳細設定…」, 結合方法, the live summary line and the toast (right end), like the reference app.
 */
import type { ReactNode } from 'react';
import type { JoinPlan, KeyAnalysis } from '../core/contracts';
import type { JoinType } from '../core/types';
import { chipViews, extraChips, otherCandidateCount, toggleCandidate, type ChipView } from '../state/keysPlan';

export const JOIN_TYPE_LABEL: Record<JoinType, string> = {
  left: '基準ファイルの行に合わせる',
  full: '他ファイルにしかない行も末尾に追加',
};

export const OTHER_SHEETS_LABEL = '他のシートも出力に含める';
export const OTHER_SHEETS_HINT = '通常は、結合した1つのシートだけを出力します。基準ファイルにあるほかのシート（集計・グラフなど）も残したいときにチェックします';
export const OTHER_SHEETS_LOCKED_HINT = '結合のやり直しが必要です。変えるには「結合をやり直す」を押してください';

/**
 * 「他のシートも出力に含める」 (default off: the output holds ONLY the merged sheet). A merge setting, like 結合方法: it can
 * only be changed before the merge; afterwards the checkbox stays visible but disabled, with the hint to redo the merge.
 */
export function OtherSheetsOption({ checked, disabled, onChange }: { checked: boolean; disabled?: boolean; onChange?: (include: boolean) => void }) {
  return (
    <label className={`settings-check${disabled ? ' is-disabled' : ''}`} title={disabled ? `${OTHER_SHEETS_HINT}\n（${OTHER_SHEETS_LOCKED_HINT}）` : OTHER_SHEETS_HINT}>
      <input type="checkbox" data-testid="include-other-sheets" checked={checked} disabled={disabled} onChange={(e) => onChange?.(e.target.checked)} />
      <span>{OTHER_SHEETS_LABEL}</span>
    </label>
  );
}

export interface KeyControlsProps {
  analysis: KeyAnalysis;
  plan: JoinPlan;
  labels: Record<string, string>;
  onPlan: (plan: JoinPlan) => void;
  onToggleFamily: (family: 'id' | 'name', on: boolean) => void;
  onOpenDetail: () => void;
  onJoinType: (t: JoinType) => void;
  /** 「他のシートも出力に含める」 (next to 結合方法) */
  includeOtherSheets: boolean;
  onIncludeOtherSheets: (include: boolean) => void;
}

function Chip({ chip, onToggle }: { chip: ChipView; onToggle: (on: boolean) => void }) {
  const title = chip.available ? chip.details.join('\n') : `${chip.label} にあたる列が見つかりません`;
  return (
    <button
      type="button"
      className={`key-chip${chip.on ? ' on' : ''}${chip.partial ? ' partial' : ''}`}
      data-testid={`key-chip-${chip.family}`}
      aria-pressed={chip.on || chip.partial}
      disabled={!chip.available}
      title={title}
      onClick={() => onToggle(!(chip.on || chip.partial))}
    >
      <span className="chip-check" aria-hidden>
        {chip.on ? '✓' : chip.partial ? '−' : ''}
      </span>
      {chip.label}
      {chip.partial && <span className="chip-note">一部</span>}
    </button>
  );
}

export function KeyControls(props: KeyControlsProps) {
  const { analysis, plan, labels } = props;
  const chips = chipViews(analysis, plan, labels);
  const extras = extraChips(analysis, plan);
  const more = otherCandidateCount(analysis, plan);
  const multi = Object.keys(analysis.pairs).length > 1;
  return (
    <>
      <span className="settings-label">結合キー</span>
      <div className="chip-row" role="group" aria-label="結合キー">
        {chips.map((c) => (
          <Chip key={c.family} chip={c} onToggle={(on) => props.onToggleFamily(c.family, on)} />
        ))}
        {extras.map((x) => (
          <button
            key={`${x.sourceId}:${x.candidateId}`}
            type="button"
            className={`key-chip${x.on ? ' on' : ''}`}
            aria-pressed={x.on}
            title={multi ? `${labels[x.sourceId] ?? x.sourceId}: ${x.label}` : x.label}
            onClick={() => props.onPlan(toggleCandidate(plan, x.sourceId, x.candidateId, !x.on))}
          >
            <span className="chip-check" aria-hidden>
              {x.on ? '✓' : ''}
            </span>
            {x.label}
            {multi && <span className="chip-note">{labels[x.sourceId] ?? ''}</span>}
          </button>
        ))}
        {more > 0 && (
          <button type="button" className="chip-more" onClick={props.onOpenDetail} title="他のキー候補を詳細設定で確認できます">
            ＋ 他の候補 {more}件
          </button>
        )}
      </div>
      <button type="button" className="btn btn-sm" data-testid="key-detail-button" onClick={props.onOpenDetail} aria-haspopup="dialog" title="結合キーの候補・照合のしかた・取り込む列を細かく設定します">
        詳細設定…
      </button>
      <label className="settings-field">
        <span>結合方法</span>
        <select
          data-testid="join-type-select"
          value={plan.joinType}
          title="基準ファイルの行に合わせる: 基準の行だけを残します。／他ファイルにしかない行も末尾に追加: 基準にない人も表の末尾に足します"
          onChange={(e) => props.onJoinType(e.target.value as JoinType)}
        >
          {(Object.keys(JOIN_TYPE_LABEL) as JoinType[]).map((t) => (
            <option key={t} value={t}>
              {JOIN_TYPE_LABEL[t]}
            </option>
          ))}
        </select>
      </label>
      <OtherSheetsOption checked={props.includeOtherSheets} onChange={props.onIncludeOtherSheets} />
    </>
  );
}

export function SettingsBar({ children, summary, summaryTitle, toast }: { children?: ReactNode; summary: ReactNode; summaryTitle?: string; toast: ReactNode }) {
  return (
    <div className="settings-bar" data-testid="settings-bar">
      {children}
      {summary !== null && summary !== undefined && summary !== '' && (
        <span className="settings-summary" data-testid="summary-line" title={summaryTitle}>
          {summary}
        </span>
      )}
      {toast}
    </div>
  );
}
