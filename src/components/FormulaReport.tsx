/**
 * 「数式の取り込み」: what happened to the formulas of the imported files, one compact block per file (side panel, after the merge).
 * Plain Japanese; amber when cells ended up EMPTY (the file stored no result for a formula that could not be kept), a quiet note when
 * some formulas became values, information only when everything stayed a formula. The collapsible 「詳しく」 lists the reasons and the
 * first example rows (click = the grid scrolls to that cell). Text logic: src/state/formulaReport.ts.
 */
import type { FormulaFileView } from '../state/formulaReport';
import { reasonBreakdown } from '../state/formulaReport';

const MAX_ROWS = 5;

export interface FormulaReportProps {
  views: readonly FormulaFileView[];
  /** source id -> file colour (the dot before the file name) */
  colors?: Record<string, string>;
  /** example rows are clickable only while the grid is shown */
  canJump: boolean;
  /** `cell`: 'plain' = the cell that holds a value / nothing instead of the formula, 'formula' = a kept formula */
  onJump: (sourceId: string, row: number, cell: 'plain' | 'formula') => void;
}

function RowButtons({ sourceId, rows, cell, canJump, onJump }: { sourceId: string; rows: readonly number[]; cell: 'plain' | 'formula'; canJump: boolean; onJump: FormulaReportProps['onJump'] }) {
  if (!rows.length) return null;
  const shown = rows.slice(0, MAX_ROWS);
  return (
    <span className="formula-rows" data-testid={`formula-rows-${sourceId}`}>
      {shown.map((r) =>
        canJump ? (
          <button key={r} type="button" className="formula-row-btn" data-row={r} onClick={() => onJump(sourceId, r, cell)} title={`${r}行目のセルへ移動`}>
            {r}行目
          </button>
        ) : (
          <span key={r} className="formula-row-btn is-static">
            {r}行目
          </span>
        ),
      )}
      {rows.length > shown.length && <span className="muted">ほか {rows.length - shown.length} 件</span>}
    </span>
  );
}

export function FormulaReport({ views, colors, canJump, onJump }: FormulaReportProps) {
  if (!views.length) return null;
  const warn = views.filter((v) => v.tone === 'warn').length;
  const tone = warn > 0 ? 'warn' : views.some((v) => v.tone === 'note') ? 'note' : 'ok';
  return (
    <section className="side-section formula-section" data-testid="formula-report" data-tone={tone} aria-label="数式の取り込み">
      <h3 title="結合したときの結果です（結合後の編集は含みません）">
        数式の取り込み
        <span className={`count-badge${warn > 0 ? ' badge-warn' : ''}`} title={warn > 0 ? `${warn} ファイルで数式が空白になりました` : `${views.length} ファイルに数式がありました`}>
          {views.length}
        </span>
      </h3>
      <div className="formula-files">
        {views.map((v) => (
          <div key={v.sourceId} className={`formula-file tone-${v.tone}`} data-testid={`formula-file-${v.sourceId}`} data-tone={v.tone}>
            <div className="formula-file-head">
              <span className="src-dot" style={colors?.[v.sourceId] ? { background: colors[v.sourceId] } : undefined} aria-hidden />
              <span className="formula-file-name" title={v.name}>
                {v.name}
              </span>
            </div>
            <ul className="formula-lines">
              {v.keptLine && (
                <li className="formula-kept" data-testid={`formula-kept-${v.sourceId}`}>
                  {v.keptLine}
                </li>
              )}
              {v.valuesLine && (
                <li className="formula-values" data-testid={`formula-values-${v.sourceId}`}>
                  {v.valuesLine}
                </li>
              )}
              {v.emptyLine && (
                <li className="formula-empty" data-testid={`formula-empty-${v.sourceId}`}>
                  {v.emptyLine}
                </li>
              )}
              {v.recalcLine && (
                <li className="formula-recalc" data-testid={`formula-recalc-${v.sourceId}`}>
                  {v.recalcLine}
                </li>
              )}
            </ul>
            {v.hint && (
              <p className="formula-hint" data-testid={`formula-hint-${v.sourceId}`}>
                {v.hint}
              </p>
            )}
            {(v.reasons.length > 0 || v.details.length > 0) && (
              <details className="formula-details" data-testid={`formula-details-${v.sourceId}`}>
                <summary>詳しく見る</summary>
                {v.reasons.length > 0 && (
                  <>
                    <p className="formula-sub" title={reasonBreakdown(v.reasons)}>
                      {v.reasonsIncludeEmpty ? '値にした・空白になった数式の内訳' : '値にした数式の内訳'}
                    </p>
                    <ul className="formula-reasons">
                      {v.reasons.map((r) => (
                        <li key={r.code}>
                          <b>
                            {r.label} {r.count.toLocaleString('ja-JP')} 件
                          </b>
                          <span className="muted">{r.note}</span>
                        </li>
                      ))}
                    </ul>
                  </>
                )}
                {v.details.length > 0 && (
                  <ul className="formula-items">
                    {v.details.map((d) => (
                      <li key={d.key} className={`sev-${d.severity}`}>
                        <span>{d.text}</span>
                        <RowButtons sourceId={v.sourceId} rows={d.rows} cell={d.cell} canJump={canJump} onJump={onJump} />
                      </li>
                    ))}
                  </ul>
                )}
              </details>
            )}
          </div>
        ))}
      </div>
    </section>
  );
}
