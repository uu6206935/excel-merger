/**
 * 「列を選んで結合キーに追加」 (key detail modal, one per file tab): the user builds a key from columns of the base and of
 * this file — 1..3 columns each (「＋列を追加」), every column of the data table listed by header text + column letter —
 * sees the detected type of the chosen columns live, and adds it with 「追加」.
 */
import { type Ref, useRef, useState } from 'react';
import type { KeyAnalysis } from '../core/contracts';
import { columnOptions, manualKeyHint } from '../state/keysPlan';

const MAX_COLS = 3;

export type AddManualResult = { ok: true; id: string; existing: boolean } | { ok: false; error: string };

export interface ManualKeyPickerProps {
  sourceId: string;
  analysis: KeyAnalysis;
  baseLabel: string;
  otherLabel: string;
  onAdd: (baseCols: number[], otherCols: number[]) => AddManualResult;
}

interface SideProps {
  legend: string;
  fileLabel: string;
  testPrefix: string;
  options: ReturnType<typeof columnOptions>;
  value: string[];
  onChange: (v: string[]) => void;
  firstRef?: Ref<HTMLSelectElement>;
}

function Side({ legend, fileLabel, testPrefix, options, value, onChange, firstRef }: SideProps) {
  return (
    <fieldset className="picker-side" data-testid={`${testPrefix}-group`}>
      <legend>
        {legend}
        <span className="picker-file muted" title={fileLabel}>
          {fileLabel}
        </span>
      </legend>
      {value.map((v, i) => (
        <div className="picker-row" key={i}>
          <select
            ref={i === 0 ? firstRef : undefined}
            data-testid={`${testPrefix}-select-${i}`}
            aria-label={`${legend} ${value.length > 1 ? i + 1 : ''}`.trim()}
            value={v}
            onChange={(e) => onChange(value.map((x, j) => (j === i ? e.target.value : x)))}
          >
            <option value="">選んでください</option>
            {options.map((o) => (
              <option key={o.col} value={String(o.col)}>
                {o.label}
              </option>
            ))}
          </select>
          {i > 0 && (
            <button
              type="button"
              className="icon-btn"
              data-testid={`${testPrefix}-remove-${i}`}
              aria-label={`${legend} ${i + 1} 列目を外す`}
              title="この列を外す"
              onClick={() => onChange(value.filter((_, j) => j !== i))}
            >
              ×
            </button>
          )}
        </div>
      ))}
      {value.length < MAX_COLS && (
        <button type="button" className="btn btn-xs picker-add-col" data-testid={`${testPrefix}-add-column`} onClick={() => onChange([...value, ''])} title="氏名＋大学のように、複数の列を組み合わせてキーにします（最大 3 列）">
          ＋列を追加
        </button>
      )}
    </fieldset>
  );
}

export function ManualKeyPicker({ sourceId, analysis, baseLabel, otherLabel, onAdd }: ManualKeyPickerProps) {
  const pair = analysis.pairs[sourceId];
  const [baseSel, setBaseSel] = useState<string[]>(['']);
  const [otherSel, setOtherSel] = useState<string[]>(['']);
  const [msg, setMsg] = useState<{ tone: 'ok' | 'error'; text: string } | null>(null);
  const firstRef = useRef<HTMLSelectElement>(null);
  if (!pair) return null;

  const baseOpts = columnOptions(analysis.base);
  const otherOpts = columnOptions(pair.other);
  const colOf = (sheet: typeof analysis.base, v: string) => (v === '' ? undefined : sheet.columns.find((c) => c.col === Number(v)));
  const hint = manualKeyHint(
    baseSel.map((v) => colOf(analysis.base, v)),
    otherSel.map((v) => colOf(pair.other, v)),
  );

  const change = (set: (v: string[]) => void) => (v: string[]) => {
    setMsg(null);
    set(v);
  };

  const add = () => {
    if (!hint.ready) return;
    const res = onAdd(baseSel.map(Number), otherSel.map(Number));
    if (!res.ok) {
      setMsg({ tone: 'error', text: res.error });
      return;
    }
    setMsg({
      tone: 'ok',
      text: res.existing
        ? '同じ組み合わせがすでに表にあるため、「使う」にチェックを入れました。'
        : '追加しました。表の一番下（優先度は最後）に「使う」で入っています。',
    });
    setBaseSel(['']);
    setOtherSel(['']);
    requestAnimationFrame(() => firstRef.current?.focus());
  };

  return (
    <section className="detail-section manual-picker" aria-label="列を選んで結合キーに追加" data-testid="manual-key-picker">
      <h3>列を選んで結合キーに追加</h3>
      <p className="detail-hint muted picker-lead">自動で見つからなかったとき、両方のファイルで同じ値が入っている列を選んで、結合キーにできます。</p>
      <div className="picker-grid">
        <Side legend="基準ファイルの列" fileLabel={baseLabel} testPrefix="manual-base" options={baseOpts} value={baseSel} onChange={change(setBaseSel)} firstRef={firstRef} />
        <span className="picker-link" aria-hidden>
          ⟷
        </span>
        <Side legend="このファイルの列" fileLabel={otherLabel} testPrefix="manual-other" options={otherOpts} value={otherSel} onChange={change(setOtherSel)} />
      </div>
      <p className={`picker-hint tone-${hint.tone}`} data-testid="manual-key-hint" role="status" aria-live="polite">
        {hint.text}
      </p>
      <div className="picker-actions">
        <button type="button" className="btn btn-primary btn-sm" data-testid="manual-key-add" disabled={!hint.ready} onClick={add} title="選んだ列の組み合わせを、結合キーの候補に追加します（追加すると結合に使われます）">
          追加
        </button>
        {msg && (
          <span className={`picker-msg tone-${msg.tone}`} data-testid="manual-key-message" role={msg.tone === 'error' ? 'alert' : 'status'}>
            {msg.text}
          </span>
        )}
      </div>
    </section>
  );
}
