/**
 * Main area before the merge: 結合プレビュー (per-file check table) + the columns that will be added
 * (click a column chip to include / exclude it).
 */
import type { JoinPlan, JoinResult, KeyAnalysis } from '../core/contracts';
import { importChoices, importKey, pairIds, previewRows } from '../state/keysPlan';
import { formulaColumnsOf, PREVIEW_FORMULA_HINT } from '../state/formulaReport';
import { droppedSheetsLine, otherSheetNames } from '../state/notes';
import type { FileEntry } from '../state/types';

export interface MergePreviewProps {
  files: FileEntry[];
  analysis: KeyAnalysis | null;
  plan: JoinPlan | null;
  joinResult: JoinResult | null;
  overrides: Record<string, boolean>;
  keyError: string | null;
  analyzing: boolean;
  /** ids of the files without a key (they take no part in the merge) */
  excluded?: string[];
  /** 「他のシートも出力に含める」: false (default) = the output holds only the merged sheet */
  includeOtherSheets?: boolean;
  onImports: (changes: Record<string, boolean>) => void;
  onOpenDetail: (tab?: string) => void;
}

const num = (n: number | null) => (n === null ? '－' : n.toLocaleString());

export function MergePreview(props: MergePreviewProps) {
  const { files, analysis, plan, joinResult, overrides } = props;
  const byId = new Map(files.map((f) => [f.id, f]));
  const ready = files.filter((f) => f.status === 'ready');
  const loading = files.some((f) => f.status === 'loading');

  if (ready.length < 2) {
    return (
      <div className="table-wrap preview-wrap" data-testid="merge-preview">
        <div className="table-placeholder">
          {loading ? (
            <p>
              <span className="spinner spinner-sm" aria-hidden /> ファイルを読み込んでいます…
            </p>
          ) : (
            <>
              <p className="placeholder-title">もう1つ以上のファイルを追加してください</p>
              <p className="muted">基準のファイル（★）の表に、他のファイルの列を ID や氏名で突き合わせて追加します。</p>
            </>
          )}
        </div>
      </div>
    );
  }
  if (props.keyError && !joinResult) {
    return (
      <div className="table-wrap preview-wrap" data-testid="merge-preview">
        <div className="card-error preview-error" role="alert">
          {props.keyError}
        </div>
      </div>
    );
  }
  if (!analysis || !plan) {
    return (
      <div className="table-wrap preview-wrap" data-testid="merge-preview">
        <div className="table-placeholder">
          <p>
            <span className="spinner spinner-sm" aria-hidden /> {props.analyzing || loading ? '結合キーを探しています…' : '結合キーを確認できませんでした'}
          </p>
        </div>
      </div>
    );
  }
  const rows = previewRows(analysis, plan, joinResult);
  const labelOf = (id: string) => byId.get(id)?.fileName ?? id;
  // the base file's other sheets: what the single-sheet output leaves out (a prediction; the merge result reports the truth)
  const otherSheets = otherSheetNames(byId.get(analysis.baseId));
  const excluded = new Set(props.excluded ?? rows.filter((r) => r.excluded).map((r) => r.sourceId));
  const allExcluded = excluded.size > 0 && pairIds(analysis).every((s) => excluded.has(s));
  let total = 0;
  // columns that will be imported and hold formulas in their file (cheap scan of the data region, cached per book)
  const formulaColsOf = new Map<string, number>();
  const groups = pairIds(analysis).map((s) => {
    const choices = importChoices(analysis, plan, joinResult, overrides, s);
    if (!excluded.has(s)) {
      total += choices.filter((c) => c.on).length;
      const f = byId.get(s);
      const withFormulas = formulaColumnsOf(f?.book, f?.region);
      const n = choices.filter((c) => c.on && withFormulas.has(c.col)).length;
      if (n > 0) formulaColsOf.set(s, n);
    }
    return { s, choices, sheet: analysis.pairs[s].other.sheet };
  });

  return (
    <div className="table-wrap preview-wrap" data-testid="merge-preview">
      {excluded.size > 0 && (
        <div className="preview-notice" role="status" data-testid="excluded-notice">
          {allExcluded ? (
            <>
              <b>どのファイルにも結合キーが選ばれていません。</b>
              結合するには、少なくとも 1 つのファイルで結合キーを選んでください。
            </>
          ) : (
            <>
              <b>結合対象外: {[...excluded].map((s) => `「${labelOf(s)}」`).join('、')}</b>
              は結合キーが見つからないため、今回の結合には含まれません（列は取り込まれません）。他のファイルだけで結合できます。
            </>
          )}
          <span className="preview-notice-actions">
            {[...excluded].map((s) => (
              <button key={s} type="button" className="btn btn-sm" data-testid={`excluded-pick-key-${s}`} onClick={() => props.onOpenDetail(s)} aria-haspopup="dialog">
                「{labelOf(s)}」のキーを選ぶ…
              </button>
            ))}
          </span>
        </div>
      )}
      <section className="preview-section">
        <h3 className="preview-h">結合プレビュー</h3>
        <table className="check-table">
          <thead>
            <tr>
              <th>ファイル</th>
              <th className="c-num">行数</th>
              <th>キー</th>
              <th className="c-num">一致</th>
              <th className="c-num" title="基準: どのファイルにも該当しない行／他ファイル: 基準にない行">未一致</th>
              <th className="c-num">重複</th>
              <th className="c-num">空欄</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((r) => (
              <tr key={r.sourceId} className={`${r.isBase ? 'is-base' : ''}${r.excluded ? ' is-excluded' : ''}`.trim() || undefined}>
                <td>
                  <span className="src-dot" style={{ background: byId.get(r.sourceId)?.color }} aria-hidden />
                  <span className="file-ellip" title={labelOf(r.sourceId)}>
                    {labelOf(r.sourceId)}
                  </span>
                  {r.isBase && <span className="tag tag-base">★ 基準</span>}
                </td>
                <td className="c-num">{r.rows.toLocaleString()}</td>
                <td>
                  {r.isBase ? (
                    <span className="muted">―</span>
                  ) : r.excluded ? (
                    <button type="button" className="link-btn" data-testid={`preview-excluded-${r.sourceId}`} onClick={() => props.onOpenDetail(r.sourceId)} title="結合キーを選ぶ">
                      <span className="tag tag-warn">結合対象外</span> キーを選ぶ…
                    </button>
                  ) : (
                    <button type="button" className="link-btn" onClick={() => props.onOpenDetail(r.sourceId)} title="詳細設定で変更">
                      {r.keys}
                    </button>
                  )}
                </td>
                <td className="c-num">{num(r.matched)}</td>
                <td className={`c-num${r.unmatched ? ' text-warn' : ''}`}>{num(r.unmatched)}</td>
                <td className={`c-num${r.duplicates ? ' text-warn' : ''}`}>{num(r.duplicates)}</td>
                <td className="c-num">{num(r.blanks)}</td>
              </tr>
            ))}
          </tbody>
        </table>
        {props.keyError && (
          <div className="card-error" role="alert">
            {props.keyError}
          </div>
        )}
      </section>

      <section className="preview-section">
        <h3 className="preview-h">
          追加される列 <span className="count-badge">{total}</span>
          <span className="preview-h-note muted">クリックで取り込む／取り込まないを切り替え</span>
        </h3>
        {formulaColsOf.size > 0 && (
          <p
            className="preview-formula-hint"
            data-testid="preview-formulas-hint"
            data-formula-columns={[...formulaColsOf].map(([s, n]) => `${s}:${n}`).join(',')}
            title={[...formulaColsOf].map(([s, n]) => `${labelOf(s)}: 数式を含む列 ${n} 列`).join('\n')}
          >
            {PREVIEW_FORMULA_HINT}
          </p>
        )}
        {groups.map(({ s, choices, sheet }) => excluded.has(s) ? (
          <div key={s} className="import-group is-excluded">
            <div className="import-group-head">
              <span className="src-dot" style={{ background: byId.get(s)?.color }} aria-hidden />
              <span className="file-ellip" title={labelOf(s)}>
                {labelOf(s)}
              </span>
              <span className="tag tag-warn">結合対象外</span>
              <span className="muted">結合キーが選ばれるまで、列は追加されません</span>
            </div>
          </div>
        ) : (
          <div key={s} className="import-group">
            <div className="import-group-head">
              <span className="src-dot" style={{ background: byId.get(s)?.color }} aria-hidden />
              <span className="file-ellip" title={labelOf(s)}>
                {labelOf(s)}
              </span>
              <span className="muted">（{choices.filter((c) => c.on).length}/{choices.length} 列）</span>
            </div>
            <div className="col-chips">
              {choices.map((ch) => (
                <button
                  key={ch.col}
                  type="button"
                  className={`col-chip${ch.on ? ' on' : ''}${ch.isKey ? ' is-key' : ''}`}
                  style={ch.on ? { borderColor: byId.get(s)?.color } : undefined}
                  aria-pressed={ch.on}
                  title={ch.isKey ? 'キーの列（通常は取り込みません）' : ch.on ? '取り込みません' : '取り込みます'}
                  onClick={() => props.onImports({ [importKey(s, sheet, ch.col)]: !ch.on })}
                >
                  {ch.on ? '＋ ' : ''}
                  {ch.header}
                </button>
              ))}
            </div>
          </div>
        ))}
        <p className="preview-foot muted">
          「結合する」を押すと、基準ファイル（★）の表の右側にこれらの列を追加します。元のファイルは変更されません。
          {plan.joinType === 'full' && joinResult && joinResult.summary.appended > 0 && ` 他ファイルにしかない ${joinResult.summary.appended} 行は表の末尾に追加します。`}
        </p>
        {otherSheets.length > 0 && (
          <p className="preview-foot preview-sheets" data-testid="preview-sheets" data-single-sheet={props.includeOtherSheets ? 'false' : 'true'}>
            {props.includeOtherSheets ? `基準ファイルの他のシート ${otherSheets.length} 枚（${otherSheets.map((n) => `「${n}」`).join('')}）も、そのまま出力に含めます。` : `${droppedSheetsLine(otherSheets)}。`}
          </p>
        )}
      </section>
    </div>
  );
}
