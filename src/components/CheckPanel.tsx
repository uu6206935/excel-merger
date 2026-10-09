/**
 * 「チェック結果」: 未一致 / 重複 / 要確認 / 数式→値 … lists with counts. Clicking a row with an Excel row
 * scrolls the grid to it (after the merge). Check items of the keys module that point into another file
 * (CheckItem.otherRows) are shown as summary lines on top of their group: 「02_面接評価 の 12行目・13行目」.
 */
import { useState } from 'react';
import type { CheckGroup, CheckRow } from '../state/keysPlan';

const MAX_ITEMS = 300;

const GROUP_HELP: Record<CheckGroup['id'], string> = {
  unmatched: '基準の行のうち、どのファイルにも該当する行が見つからなかったもの',
  partial: '一部のファイルにだけ該当する行がなかった基準の行（その列は空欄になります）',
  duplicate: '同じキーの行が複数あるもの（設定に従って1行を選んでいます）',
  review: '値が食い違っている行、候補が複数あって決めきれなかった行など、目で確認してほしいもの',
  toValues: '数式のまま取り込めず、計算結果の値にした数式（元のファイルに計算結果がなく、空白になったものを含みます）。理由は「数式の取り込み」に表示します',
  info: '全角・半角などの表記の違いを補正して結合した内容（確認用です）',
  otherOnly: '基準の表にない行（「他ファイルにしかない行も末尾に追加」を選ぶと末尾に足します）',
};

const COUNT_TESTID: Partial<Record<CheckGroup['id'], string>> = {
  unmatched: 'check-unmatched-count',
  duplicate: 'check-duplicate-count',
};

function ItemLine({ r, canJump, onJump, summary }: { r: CheckRow; canJump: boolean; onJump: (row: number) => void; summary?: boolean }) {
  const jump = canJump && r.row !== null;
  const cls = `tone-${r.tone}${summary ? ' check-summary' : ''}`;
  return (
    <li className={cls} data-testid={summary ? 'check-summary' : undefined}>
      {jump ? (
        <button type="button" className="check-item-btn" onClick={() => onJump(r.row!)} title={`基準の ${r.row}行目へ移動`}>
          <span className="name">{r.text}</span>
          {r.sub && <span className="reason">{r.sub}</span>}
        </button>
      ) : (
        <div className="check-item">
          <span className="name">{r.text}</span>
          {r.sub && <span className="reason">{r.sub}</span>}
        </div>
      )}
    </li>
  );
}

export function CheckPanel({ groups, canJump, onJump, emptyText }: { groups: CheckGroup[] | null; canJump: boolean; onJump: (row: number) => void; emptyText: string }) {
  const firstWithItems = groups?.find((g) => g.count > 0 && (g.id === 'unmatched' || g.id === 'duplicate' || g.id === 'review'))?.id ?? null;
  const [open, setOpen] = useState<Record<string, boolean>>({});
  const isOpen = (id: string) => open[id] ?? id === firstWithItems;

  return (
    <section className="side-section check-section" data-testid="check-panel" aria-label="チェック結果">
      <h3>チェック結果</h3>
      {!groups ? (
        <p className="side-help">{emptyText}</p>
      ) : (
        <div className="check-groups">
          {groups.map((g) => {
            const summaries = g.summaries ?? [];
            const has = g.count > 0 || summaries.length > 0;
            const o = has && isOpen(g.id);
            const tone = g.count === 0 ? 'zero' : g.id === 'review' || g.id === 'duplicate' || g.id === 'unmatched' ? 'warn' : 'info';
            return (
              <div key={g.id} className={`check-group${o ? ' open' : ''}`} data-group={g.id}>
                <button type="button" className="check-group-head" aria-expanded={o} disabled={!has} title={GROUP_HELP[g.id]} onClick={() => setOpen({ ...open, [g.id]: !o })}>
                  <span className="check-caret" aria-hidden>
                    {!has ? '' : o ? '▾' : '▸'}
                  </span>
                  <span className="check-title">{g.title}</span>
                  <span className={`count-badge badge-${tone}`} data-testid={COUNT_TESTID[g.id]}>
                    {g.count}
                  </span>
                </button>
                {o && (
                  <ul className="side-list check-list">
                    {summaries.map((r) => (
                      <ItemLine key={r.key} r={r} canJump={canJump} onJump={onJump} summary />
                    ))}
                    {g.rows.slice(0, MAX_ITEMS).map((r) => (
                      <ItemLine key={r.key} r={r} canJump={canJump} onJump={onJump} />
                    ))}
                    {g.rows.length > MAX_ITEMS && <li className="muted">ほか {g.rows.length - MAX_ITEMS} 件</li>}
                    {g.rows.length === 0 && summaries.length === 0 && (
                      <li className="muted">{g.id === 'toValues' ? '数式のまま取り込めなかった数式を、計算結果の値に置き換えました' : '詳細はありません'}</li>
                    )}
                  </ul>
                )}
              </div>
            );
          })}
        </div>
      )}
    </section>
  );
}
