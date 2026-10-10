/**
 * Round 9 (plan A): the sheets of the output.
 *   PlannedSheets   main area before the merge (MergePreview): 「出力されるシート」, the list planSheetNames will produce
 *   SheetListSection side panel after the merge: the collapsible 「シート一覧」 (names only; the grid shows the merged sheet)
 */
import type { OutputSheetRow, PlannedOutput } from '../state/sheetPlan';

export const SUFFIX_BADGE = (n: number | undefined) => `名前が重なったため(${n ?? 1})を付けました`;
export const SHORTENED_BADGE = '31文字までに短縮';

const KIND_TAG: Record<OutputSheetRow['kind'], string> = { merged: '結合した表', base: '基準ファイル', copy: 'コピー' };

function Dot({ color }: { color?: string }) {
  return color ? <span className="src-dot" style={{ background: color }} aria-hidden /> : null;
}

/** 「出力されるシート」: 1. merged sheet, then the base's sheets (when kept), then the copies of the other files in card order. */
export function PlannedSheets({ plan, colors }: { plan: PlannedOutput; colors: Record<string, string> }) {
  const { sheets, skipped } = plan;
  if (!sheets.length) return null;
  return (
    <section className="preview-section" data-testid="preview-output-sheets" data-sheet-count={sheets.length}>
      <h3 className="preview-h">
        出力されるシート <span className="count-badge">{sheets.length}</span>
        <span className="preview-h-note muted">他のファイルのシートは、ファイルカードの並び順に入ります</span>
      </h3>
      <ol className="sheet-plan">
        {sheets.map((s, i) => (
          <li key={i} className={`sheet-plan-item kind-${s.kind}`} data-testid={`planned-sheet-${i}`} data-kind={s.kind} data-name={s.name}>
            <span className="sheet-no">{i + 1}</span>
            <Dot color={colors[s.sourceId]} />
            <span className="sheet-name" title={s.name}>
              {s.name}
            </span>
            {s.kind === 'merged' && <span className="tag tag-new">結合した表</span>}
            {s.kind === 'base' && <span className="tag tag-muted">基準ファイル</span>}
            {s.kind === 'copy' && (
              <span className="sheet-from muted" title={`${s.file ?? ''} の「${s.original ?? ''}」`}>
                ← 「{s.file}」の「{s.original}」
              </span>
            )}
            {s.suffixed && (
              <span className="tag tag-warn" data-testid={`planned-sheet-suffixed-${i}`} title="同じ名前のシートがすでにあるため、あとから入れるほうに番号を付けます">
                {SUFFIX_BADGE(s.suffixNo)}
              </span>
            )}
            {s.shortened && (
              <span className="tag tag-muted" data-testid={`planned-sheet-shortened-${i}`} title="Excel のシート名は 31 文字までのため、「ファイル名_シート名」を短くしました">
                {SHORTENED_BADGE}
              </span>
            )}
            {s.hidden && (
              <span className="tag tag-muted" title="元のファイルで非表示のシートは、非表示のまま入ります">
                非表示
              </span>
            )}
          </li>
        ))}
      </ol>
      {skipped.length > 0 && (
        <p className="preview-foot muted" data-testid="planned-sheets-skipped">
          結合対象外（結合キーなし）の{skipped.map((n) => `「${n}」`).join('')}のシートは入りません。
        </p>
      )}
    </section>
  );
}

/** 「シート一覧」: collapsible, names only, with where each sheet comes from. */
export function SheetListSection({ rows, colors }: { rows: OutputSheetRow[]; colors: Record<string, string> }) {
  if (rows.length < 2) return null;
  return (
    <section className="side-section sheet-list-section" aria-label="シート一覧">
      <details className="side-details" data-testid="sheet-list-details">
        <summary title="出力されるすべてのシートの名前です。表に表示されるのは、結合した表のシートだけです">
          シート一覧 <span className="count-badge">{rows.length}</span>
        </summary>
        <ol className="side-list sheet-list" data-testid="sheet-list">
          {rows.map((r, i) => (
            <li key={i} className={`kind-${r.kind}`} data-kind={r.kind} data-name={r.name} data-testid={`sheet-list-item-${i}`}>
              <Dot color={r.sourceId ? colors[r.sourceId] : undefined} />
              <span className="name" title={r.kind === 'copy' && r.file ? `${r.name}\n「${r.file}」の「${r.original ?? ''}」` : r.name}>
                {r.name}
              </span>
              <span className="tag tag-muted">{KIND_TAG[r.kind]}</span>
              {r.hidden && <span className="tag tag-muted">非表示</span>}
              {r.complete === false && (
                <span className="tag tag-warn" title="コピーできなかったものがあります。くわしくは「警告」を見てください">
                  一部のみ
                </span>
              )}
            </li>
          ))}
        </ol>
      </details>
    </section>
  );
}
