/**
 * First-run help: the four steps of the flow (the same names as the step indicator in the header) as a collapsible
 * 「使い方」 panel on the empty screen (open the first time, the choice is remembered), and the same text plus the keyboard
 * shortcuts in a dialog that the header's 「使い方」 button opens at any time.
 */
import { useState } from 'react';
import { Modal } from './Modal';

export const HELP_STEPS: ReadonlyArray<{ num: string; title: string; body: string }> = [
  { num: '①', title: 'ファイル', body: 'まとめたいExcel（.xlsx / .xls / .csv など）を2つ以上、ここにドロップします。最初のファイルが「基準」（出力の元になる表）です。' },
  { num: '②', title: '結合キー', body: 'ID・氏名など、同じ人を見分ける列を自動で探します。合っているか確認して「結合する」を押します（「詳細設定…」で変更できます）。' },
  { num: '③', title: '列の編集', body: '列見出しをクリックして選び、ドラッグで移動。右クリックで挿入・削除、Ctrl+Z で元に戻せます。セルの編集もできます。' },
  { num: '④', title: '出力', body: '「Excel出力」で、結合した表（1つのシート）を数式・書式を保ったまま保存します。元のファイルは変更されません。' },
];

export const SHORTCUTS: ReadonlyArray<{ keys: string; what: string }> = [
  { keys: 'Ctrl + O / Ctrl + S', what: 'ファイルを選ぶ / Excel出力の画面を開く（結合後）' },
  { keys: 'F6 / Shift + F6', what: '右のパネルの主なボタン（結合する・Excel出力）へ移動 / 表に戻る' },
  { keys: 'Tab（表の右端で）', what: '表の外へ移動（表の中では次のセルへ）' },
  { keys: 'Ctrl + Z / Ctrl + Y', what: '元に戻す / やり直す（何回でも）' },
  { keys: 'Ctrl + クリック / Shift + クリック', what: '列見出しを複数選ぶ（離れた列 / 範囲）' },
  { keys: 'Ctrl + Space', what: '今いるセルの列を選ぶ' },
  { keys: 'Shift + ← →', what: '選んだ列を左右に広げる' },
  { keys: 'Delete / Ctrl + -', what: '選んだ列を削除' },
  { keys: 'Ctrl + +', what: '選んだ列の左に列を挿入' },
  { keys: 'F2 / ダブルクリック', what: 'セルを編集（Enter で確定、Esc で取り消し）' },
  { keys: '列の一覧で Alt + ↑ ↓', what: '選んだ列をキーボードで移動' },
  { keys: 'ファイルのカードをドラッグ / Alt + ← →', what: '結合前: ファイルの並べ替え（取り込む列の並びも変わります）／結合後: そのファイルの列をまとめて移動' },
  { keys: 'ファイルのカード・色の帯をクリック', what: '結合後: そのファイルの列をすべて選ぶ（カードは、表をその列が中央に来る位置へ動かします）。色の帯はドラッグで列をまとめて移動' },
];

const STORAGE_KEY = 'excel-merger.help-open';

function readOpen(): boolean {
  try {
    return window.localStorage.getItem(STORAGE_KEY) !== '0';
  } catch {
    return true;
  }
}

function writeOpen(open: boolean): void {
  try {
    window.localStorage.setItem(STORAGE_KEY, open ? '1' : '0');
  } catch {
    /* storage can be blocked (private window, file://): the panel just starts open every time */
  }
}

/** Collapsible 「使い方」 panel of the empty screen. */
export function HelpPanel() {
  const [open, setOpen] = useState<boolean>(readOpen);
  return (
    <details
      className="help-panel"
      data-testid="help-panel"
      open={open}
      onToggle={(e) => {
        const next = (e.currentTarget as HTMLDetailsElement).open;
        if (next !== open) {
          setOpen(next);
          writeOpen(next);
        }
      }}
    >
      <summary title="作業の流れ（4ステップ）。閉じても、上の「使い方」ボタンでいつでも見られます">使い方（4ステップ）</summary>
      <ol className="help-steps">
        {HELP_STEPS.map((s) => (
          <li key={s.num}>
            <b>
              <span className="help-num" aria-hidden>
                {s.num}
              </span>
              {s.title}
            </b>
            <span>{s.body}</span>
          </li>
        ))}
      </ol>
    </details>
  );
}

/** 「使い方」 dialog (header button). */
export function HelpDialog({ onClose }: { onClose: () => void }) {
  return (
    <Modal title="使い方" subtitle="ファイルを投げる → 結合 → 列を並べる → 出力" testId="help-dialog" size="md" onClose={onClose} initialFocus="[data-testid=help-close]" footer={<button type="button" className="btn btn-primary btn-wide" data-testid="help-close" onClick={onClose}>閉じる</button>}>
      <ol className="help-steps help-steps-vertical">
        {HELP_STEPS.map((s) => (
          <li key={s.num}>
            <b>
              <span className="help-num" aria-hidden>
                {s.num}
              </span>
              {s.title}
            </b>
            <span>{s.body}</span>
          </li>
        ))}
      </ol>
      <h3 className="help-h">キー操作</h3>
      <table className="help-keys">
        <tbody>
          {SHORTCUTS.map((k) => (
            <tr key={k.keys}>
              <th scope="row">{k.keys}</th>
              <td>{k.what}</td>
            </tr>
          ))}
        </tbody>
      </table>
      <p className="muted help-foot">ファイルはこのPCの外に送信されません。元のファイルは変更されません。</p>
    </Modal>
  );
}
