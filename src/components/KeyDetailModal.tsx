/**
 * 詳細設定 (key detail modal): one tab per other file with the key candidates (使う / 優先 / 一致率 …), the match rate
 * sentence (Power Query style), 「列を選んで結合キーに追加」, normalisation options, duplicate policy, 取り込む列 and
 * 「もしかして」 suggestions. A file without a selected key is 結合対象外 (it takes part in the merge as soon as a key is
 * selected here).
 */
import { useState } from 'react';
import type { CandidateDTO, JoinPlan, JoinResult, KeyAnalysis, NormalizeOptions } from '../core/contracts';
import type { DuplicatePolicy } from '../core/types';
import { OPTION_LABELS } from '../state/helpers';
import {
  hasManualLink,
  importChoices,
  importKey,
  isManualCandidate,
  matchRate,
  matchRateText,
  matchRateTone,
  moveCandidate,
  orderedCandidates,
  pairIds,
  setManualLink,
  toggleCandidate,
  verdictTone,
} from '../state/keysPlan';
import type { ManualKey } from '../state/types';
import { type AddManualResult, ManualKeyPicker } from './ManualKeyPicker';
import { Modal } from './Modal';

const pct = (r: number) => `${Math.round(Math.max(0, Math.min(1, r)) * 100)}%`;

const DUP_LABEL: Record<DuplicatePolicy, string> = { first: '先頭の行を使う', last: '最後の行を使う', skip: '結合しない（空欄のまま）' };

const KIND_BADGE: Partial<Record<CandidateDTO['kind'], string>> = { composite: '組み合わせ', discovered: '値から発見' };

export interface KeyDetailModalProps {
  analysis: KeyAnalysis;
  plan: JoinPlan;
  joinResult: JoinResult | null;
  labels: Record<string, string>;
  colors: Record<string, string>;
  options: NormalizeOptions;
  overrides: Record<string, boolean>;
  initialTab: string | null;
  /** ids of the files without a selected key (結合対象外) */
  excluded?: string[];
  /** manual key candidates per file (× removes them) */
  manualKeys?: Record<string, ManualKey[]>;
  /** remembers the open tab (the modal can re-mount while the analysis re-runs) */
  onTab?: (tab: string) => void;
  onPlan: (plan: JoinPlan) => void;
  onOptions: (o: Partial<NormalizeOptions>) => void;
  onImports: (changes: Record<string, boolean>) => void;
  /** 「追加」 of 列を選んで結合キーに追加 */
  onAddManual?: (sourceId: string, baseCols: number[], otherCols: number[]) => AddManualResult;
  /** × of a manual candidate */
  onRemoveManual?: (sourceId: string, candidateId: string) => void;
  onClose: () => void;
}

function MatchBar({ c }: { c: CandidateDTO }) {
  const tone = verdictTone(c.verdict);
  return (
    <span
      className="match-cell"
      title={`このファイルのキー ${c.nonBlankOther} 件中 ${c.matchedOther} 件が基準にあります／基準 ${c.nonBlankBase} 件中 ${c.matchedBase} 件が見つかります（${pct(c.coverage)}）`}
    >
      <span className="match-bar">
        <span className={`match-fill tone-${tone}`} style={{ width: pct(c.matchRate) }} />
      </span>
      <span className="match-pct">{pct(c.matchRate)}</span>
    </span>
  );
}

export function KeyDetailModal(props: KeyDetailModalProps) {
  const { analysis, plan, joinResult, labels, colors, options, overrides } = props;
  const ids = pairIds(analysis);
  const excluded = props.excluded ?? [];
  const [tab, setTab] = useState<string>(props.initialTab && ids.includes(props.initialTab) ? props.initialTab : (ids[0] ?? ''));
  const pair = analysis.pairs[tab];
  const report = joinResult?.reports[tab];
  const baseLabel = labels[analysis.baseId] ?? '基準';
  const isExcluded = excluded.includes(tab);
  const hasCandidates = (pair?.candidates.length ?? 0) > 0;
  const manual = props.manualKeys?.[tab] ?? [];
  const rate = matchRate(analysis, joinResult, tab);

  const choices = pair ? importChoices(analysis, plan, joinResult, overrides, tab) : [];
  const setAll = (on: boolean) => {
    const changes: Record<string, boolean> = {};
    for (const ch of choices) changes[importKey(tab, pair!.other.sheet, ch.col)] = on;
    props.onImports(changes);
  };

  return (
    <Modal
      title="結合キーの詳細設定"
      subtitle={`基準: ${baseLabel}（${analysis.base.sheet}・${analysis.base.rows.toLocaleString()} 行）`}
      testId="key-detail-modal"
      size="lg"
      onClose={props.onClose}
      initialFocus="[data-testid=key-detail-ok]"
      restoreFocus="[data-testid=key-detail-button]"
      footer={
        <>
          <span className="modal-footer-note muted">変更はすぐに反映されます</span>
          <button type="button" className="btn btn-primary btn-wide" data-testid="key-detail-ok" onClick={props.onClose}>
            OK
          </button>
        </>
      }
    >
      <div className="detail-options">
        <div className="detail-option-group" role="group" aria-label="照合のしかた">
          <span className="detail-option-title">照合のしかた</span>
          {OPTION_LABELS.map((o) => (
            <label key={o.key} className="check-label" title={o.title}>
              <input type="checkbox" checked={options[o.key]} onChange={(e) => props.onOptions({ [o.key]: e.target.checked })} />
              {o.label}
            </label>
          ))}
        </div>
        <div className="detail-option-group" role="radiogroup" aria-label="重複があるとき">
          <span className="detail-option-title">重複があるとき</span>
          {(Object.keys(DUP_LABEL) as DuplicatePolicy[]).map((p) => (
            <label key={p} className="check-label">
              <input type="radio" name="dup-policy" checked={plan.duplicatePolicy === p} onChange={() => props.onPlan({ ...plan, duplicatePolicy: p })} />
              {DUP_LABEL[p]}
            </label>
          ))}
        </div>
      </div>

      <div className="detail-tabs" role="tablist" aria-label="結合するファイル">
        {ids.map((s) => (
          <button
            key={s}
            type="button"
            role="tab"
            aria-selected={tab === s}
            className={`detail-tab${tab === s ? ' active' : ''}${excluded.includes(s) ? ' is-excluded' : ''}`}
            data-testid={`key-detail-tab-${s}`}
            title={`「${labels[s] ?? s}」を基準の表とつなぐ結合キーと、取り込む列を設定します`}
            onClick={() => {
              setTab(s);
              props.onTab?.(s);
            }}
          >
            <span className="src-dot" style={{ background: colors[s] }} aria-hidden />
            {labels[s] ?? s}
            {excluded.includes(s) && <span className="tag tag-warn tab-tag">結合対象外</span>}
          </button>
        ))}
      </div>

      {pair && (
        <div className="detail-body" role="tabpanel">
          {isExcluded && (
            <div className="detail-excluded" role="status" data-testid="detail-excluded">
              <b>このファイルは今、結合の対象外です。</b>
              結合キーが選ばれていないため、列は取り込まれません。
              {hasCandidates ? '下の表の「使う」にチェックを入れるか、「列を選んで結合キーに追加」でキーを選ぶと、そのまま結合に含まれます。' : '下の「列を選んで結合キーに追加」でキーを選ぶと、そのまま結合に含まれます。'}
            </div>
          )}

          <div className="detail-table-wrap">
            <table className="detail-table" data-testid="key-candidates">
              <thead>
                <tr>
                  <th>使う</th>
                  <th>優先</th>
                  <th>キー</th>
                  <th>基準の列</th>
                  <th>このファイルの列</th>
                  <th>一致率</th>
                  <th title="基準 / このファイル">重複</th>
                  <th title="基準 / このファイル">空欄</th>
                  <th>一意性</th>
                  <th>判定</th>
                </tr>
              </thead>
              <tbody>
                {orderedCandidates(pair, plan).map(({ c, priority }, _i, arr) => {
                  const used = priority !== null;
                  const usedCount = arr.filter((x) => x.priority !== null).length;
                  const isManual = isManualCandidate(c, manual);
                  return (
                    <tr key={c.id} className={used ? 'is-used' : undefined} data-testid={`key-row-${c.id}`}>
                      <td className="c-center">
                        <input
                          type="checkbox"
                          checked={used}
                          aria-label={`${c.label} を使う`}
                          disabled={!used && c.verdict === '使用不可' && !isManual}
                          onChange={(e) => props.onPlan(toggleCandidate(plan, tab, c.id, e.target.checked))}
                        />
                      </td>
                      <td className="c-prio">
                        {used ? (
                          <>
                            <span className="prio-num">{priority}</span>
                            <button type="button" className="icon-btn" aria-label="優先度を上げる" disabled={priority === 1} onClick={() => props.onPlan(moveCandidate(plan, tab, c.id, -1))}>
                              ↑
                            </button>
                            <button type="button" className="icon-btn" aria-label="優先度を下げる" disabled={priority === usedCount} onClick={() => props.onPlan(moveCandidate(plan, tab, c.id, 1))}>
                              ↓
                            </button>
                          </>
                        ) : (
                          <span className="muted">－</span>
                        )}
                      </td>
                      <td className="c-key">
                        <b>{c.label}</b>
                        {isManual && <span className="tag tag-new">手動</span>}
                        {KIND_BADGE[c.kind] && <span className="tag">{KIND_BADGE[c.kind]}</span>}
                        {isManual && props.onRemoveManual && (
                          <button
                            type="button"
                            className="icon-btn key-remove"
                            data-testid={`key-remove-${c.id}`}
                            aria-label={`${c.label} を削除`}
                            title="手動で追加したこの組み合わせを表から削除します（自動で見つけた候補は削除できません。チェックを外すと使いません）"
                            onClick={() => props.onRemoveManual?.(tab, c.id)}
                          >
                            ×
                          </button>
                        )}
                      </td>
                      <td>{c.baseColumns.map((x) => x.header).join(' ＋ ')}</td>
                      <td>{c.otherColumns.map((x) => x.header).join(' ＋ ')}</td>
                      <td>
                        <MatchBar c={c} />
                      </td>
                      <td className="c-num">
                        {c.dupRowsBase} / {c.dupRowsOther}
                      </td>
                      <td className="c-num">
                        {c.blankBase} / {c.blankOther}
                      </td>
                      <td className="c-num">{pct(Math.min(c.uniqBase, c.uniqOther))}</td>
                      <td>
                        <span className={`verdict tone-${verdictTone(c.verdict)}`}>{c.verdict}</span>
                      </td>
                    </tr>
                  );
                })}
                {pair.candidates.length === 0 && (
                  <tr>
                    <td colSpan={10} className="muted c-center">
                      キーになりそうな列が自動では見つかりませんでした。下の「列を選んで結合キーに追加」で選べます。
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          <div className="detail-summary detail-rate" data-testid="match-rate" role="status" aria-live="polite">
            {rate ? (
              <>
                <span className={`rate-sentence tone-${matchRateTone(rate)}`} data-testid="match-rate-text">
                  {matchRateText(rate)}
                </span>
                {report && (
                  <span className="rate-sub muted">
                    基準にだけある行 <b className={report.unmatchedBase.length ? 'text-warn' : ''}>{report.unmatchedBase.length}</b> ／ このファイルにだけある行 <b>{report.unmatchedOther.length}</b>
                    {report.mode === 'lookup' && (
                      <span className="tag tag-warn" title="基準ファイルに同じキーの行が複数あるため、このファイルの1行を、基準のそれぞれの行に当てはめます（VLOOKUP のような結合）">
                        基準側のキーに重複あり（それぞれの行に同じ内容を入れます）
                      </span>
                    )}
                  </span>
                )}
              </>
            ) : (
              <span className="muted" data-testid="match-rate-empty">
                {isExcluded
                  ? hasCandidates
                    ? '結合キーが選ばれていません。「使う」にチェックを入れると、一致する行数を表示します。'
                    : '結合キーが選ばれていません。下の「列を選んで結合キーに追加」でキーを選ぶと、一致する行数を表示します。'
                  : 'キーを選ぶと、一致する行数を表示します。'}
              </span>
            )}
          </div>
          <p className="detail-hint muted">上から順に試し、見つからなかった行だけ次のキーで探します。一致率はこのファイルの行のうち基準に見つかった割合です。</p>

          {props.onAddManual && (
            <ManualKeyPicker
              key={tab}
              sourceId={tab}
              analysis={analysis}
              baseLabel={baseLabel}
              otherLabel={labels[tab] ?? tab}
              onAdd={(baseCols, otherCols) => props.onAddManual!(tab, baseCols, otherCols)}
            />
          )}

          <div className="detail-columns">
            <section className="detail-section">
              <h3>
                取り込む列 <span className="count-badge">{choices.filter((c) => c.on).length}</span>
                <span className="detail-section-actions">
                  <button type="button" className="btn btn-sm" onClick={() => setAll(true)} title="このファイルの列をすべて取り込みます">
                    すべて選択
                  </button>
                  <button type="button" className="btn btn-sm" onClick={() => setAll(false)} title="取り込む列の選択をすべて外します">
                    すべて解除
                  </button>
                </span>
              </h3>
              {isExcluded && <p className="detail-hint muted">結合キーを選ぶまで、ここで選んだ列は取り込まれません。</p>}
              <ul className="import-list" data-testid="import-list">
                {choices.map((ch) => (
                  <li key={ch.col}>
                    <label className="check-label">
                      <input type="checkbox" checked={ch.on} onChange={(e) => props.onImports({ [importKey(tab, pair.other.sheet, ch.col)]: e.target.checked })} />
                      {ch.header}
                      {ch.isKey && <span className="tag">キー</span>}
                    </label>
                  </li>
                ))}
              </ul>
            </section>
            <section className="detail-section">
              <h3>
                もしかして（候補） <span className="count-badge">{report?.suggestions.length ?? 0}</span>
              </h3>
              {(plan.manualLinks[tab] ?? []).length > 0 && (
                <ul className="suggest-list manual-list" aria-label="手動で結合した行">
                  {(plan.manualLinks[tab] ?? []).map((l) => (
                    <li key={`${l.baseRow}-${l.otherRow}`}>
                      <div className="suggest-text">
                        <span className="tag tag-ok">手動</span> 基準 {l.baseRow}行目 ⟷ このファイル {l.otherRow}行目
                      </div>
                      <button type="button" className="btn btn-sm" onClick={() => props.onPlan(setManualLink(plan, tab, l.baseRow, l.otherRow, false))}>
                        取り消す
                      </button>
                    </li>
                  ))}
                </ul>
              )}
              {report && report.suggestions.length > 0 ? (
                <ul className="suggest-list">
                  {report.suggestions.map((sg, i) => {
                    const linked = hasManualLink(plan, tab, sg.base.excelRow, sg.other.excelRow);
                    return (
                      <li key={i}>
                        <div className="suggest-text">
                          <div>
                            <span className="muted">基準 {sg.base.excelRow}行目</span> {sg.base.preview.filter(Boolean).join(' ')}
                          </div>
                          <div>
                            <span className="muted">このファイル {sg.other.excelRow}行目</span> {sg.other.preview.filter(Boolean).join(' ')}
                          </div>
                          <div className="suggest-reason">{sg.reasons.join('・')}</div>
                        </div>
                        <button
                          type="button"
                          className={`btn btn-sm${linked ? '' : ' btn-primary'}`}
                          onClick={() => props.onPlan(setManualLink(plan, tab, sg.base.excelRow, sg.other.excelRow, !linked))}
                        >
                          {linked ? '結合を取り消す' : 'この組み合わせで結合'}
                        </button>
                      </li>
                    );
                  })}
                </ul>
              ) : (
                <p className="muted detail-empty">候補はありません</p>
              )}
            </section>
          </div>
        </div>
      )}
    </Modal>
  );
}
