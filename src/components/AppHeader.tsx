/**
 * Header bar: title · server status · Excel本体 availability · step indicator (segmented control, right).
 */
import { DISK_LOW_CHIP_TEXT, diskNotice } from '../state/helpers';
import type { ServerState } from '../state/types';

export const STEPS = ['ファイル', '結合キー', '列の編集', '出力'] as const;
const STEP_NUM = ['①', '②', '③', '④'];

export function StepIndicator({ step }: { step: number }) {
  return (
    <ol className="step-toggle" aria-label="作業の流れ" data-testid="step-indicator">
      {STEPS.map((label, i) => (
        <li
          key={label}
          className={`step-btn${i === step ? ' active' : ''}${i < step ? ' done' : ''}`}
          aria-current={i === step ? 'step' : undefined}
          data-testid={`step-${i + 1}`}
        >
          <span className="step-num">{STEP_NUM[i]}</span> {label}
        </li>
      ))}
    </ol>
  );
}

export function AppHeader({
  server,
  step,
  mock,
  onHelp,
  onAddFiles,
}: {
  server: ServerState;
  step: number;
  mock?: boolean;
  onHelp?: () => void;
  /** shown once files are loaded: the 「＋ ファイルを選択」 card scrolls out of sight when there are many files */
  onAddFiles?: () => void;
}) {
  const st = server.status;
  const online = !!st;
  const excel = !!st?.excel.installed;
  const disk = diskNotice(server);
  let statusText: string;
  let statusTitle: string;
  if (!server.checked) {
    statusText = '接続を確認中…';
    statusTitle = '';
  } else if (online && !server.connected) {
    statusText = 'ローカルサーバーに接続できません';
    statusTitle = 'サーバー（黒い画面）が閉じられた可能性があります。start.bat から起動し直してください。';
  } else if (online) {
    statusText = 'ローカル実行中';
    statusTitle = `このPCの中だけで動いています（ファイルは外部に送信されません）。出力先: ${st.outputDir}`;
  } else {
    statusText = 'スタンドアロン';
    statusTitle = 'ブラウザだけで動いています（ファイルは外部に送信されません）。Excel本体を使った出力はできません。';
  }
  return (
    <header className="app-header">
      <h1>Excel 統合ツール</h1>
      <div className="status-line" data-testid="server-status" data-online={online ? 'true' : 'false'} data-excel={excel ? 'true' : 'false'} title={statusTitle}>
        <span className={`status-dot ${!server.checked ? 'pending' : online ? (server.connected ? 'online' : 'lost') : 'offline'}`} aria-hidden />
        <span>{statusText}</span>
        {server.checked && (
          <span className={`tag ${excel ? 'tag-ok' : 'tag-muted'}`} title={excel ? `Excel ${st?.excel.version ?? ''} で出力できます（最も忠実）` : 'Excel本体がないため、内蔵エンジンで出力します'}>
            Excel本体: {excel ? '利用可能' : 'なし'}
          </span>
        )}
        {disk && (
          <span className="tag tag-warn disk-chip" data-testid="disk-low-chip" role="note" title={disk.summary} aria-label={disk.summary}>
            {DISK_LOW_CHIP_TEXT}
          </span>
        )}
        {mock && <span className="tag tag-warn">デモデータ（?mock）</span>}
      </div>
      <StepIndicator step={step} />
      {onAddFiles && (
        <button type="button" className="btn btn-sm header-add-btn" data-testid="header-add-files" onClick={onAddFiles} title="Excelファイルを選んで追加します（ウィンドウにドロップしても追加できます）">
          ＋ ファイルを追加
        </button>
      )}
      {onHelp && (
        <button type="button" className="btn btn-sm help-btn" data-testid="help-button" onClick={onHelp} aria-haspopup="dialog" title="使い方と、キー操作の一覧を表示します">
          使い方
        </button>
      )}
    </header>
  );
}
