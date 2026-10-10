/**
 * Excel出力: file name, engine (Excel本体 = recommended when the server has Excel, else 内蔵エンジン), progress,
 * result (「保存しました」 + Excelで開く / フォルダを開く, or the browser download in standalone mode).
 */
import { useEffect, useState } from 'react';
import { diskNotice, normalizeExportName } from '../state/helpers';
import type { ExportEngineChoice, ExportState, ServerState } from '../state/types';
import { ProgressBar } from './BusyOverlay';
import { Modal } from './Modal';

/** After this long (ms) an Excel本体 export that is still running gets a hint (window.__EM_SLOW_HINT_MS overrides it: tests). */
const SLOW_HINT_MS = 45_000;

function slowHintMs(): number {
  const v = (globalThis as { __EM_SLOW_HINT_MS?: unknown }).__EM_SLOW_HINT_MS;
  return typeof v === 'number' && v >= 0 ? v : SLOW_HINT_MS;
}

export interface ExportDialogProps {
  defaultName: string;
  format: 'xlsx' | 'xlsm';
  server: ServerState;
  exportState: ExportState;
  editCount: number;
  /** 「他のファイルのシートも含める」 was used: the copied sheets may hold charts / pivot tables (Excel本体 keeps them best) */
  includeSourceSheets?: boolean;
  onExport: (fileName: string, engine: ExportEngineChoice) => void;
  onCancelExport: () => void;
  onOpenFile: () => void;
  onOpenFolder: () => void;
  onDownloadAgain: () => void;
  onClose: () => void;
}

export function ExportDialog(props: ExportDialogProps) {
  const { server, exportState: ex, format } = props;
  // the server answered before but does not any more (its console window was closed): same as standalone mode
  const serverDown = !!server.status && !server.connected;
  const serverMode = !!server.status && !serverDown;
  const excelOk = serverMode && !!server.status?.excel.installed;
  const [name, setName] = useState(props.defaultName);
  // low disk space (server says so; refreshed when the dialog opens): the built-in engine is the default, Excel本体 stays selectable
  const disk = diskNotice(server);
  // the user's own pick; until then the default follows the (possibly just refreshed) status
  const [picked, setPicked] = useState<ExportEngineChoice | null>(null);
  const engine: ExportEngineChoice = picked ?? (excelOk && !disk?.tempLow ? 'excel' : 'js');
  const running = ex.phase === 'running';
  const done = ex.phase === 'done';
  const effectiveEngine: ExportEngineChoice = engine === 'excel' && excelOk ? 'excel' : 'js';
  const finalName = normalizeExportName(name || props.defaultName, format);
  const invalid = name.trim() === '';
  const [slow, setSlow] = useState(false);
  useEffect(() => {
    setSlow(false);
    if (!running || ex.engine !== 'excel') return;
    const t = setTimeout(() => setSlow(true), slowHintMs());
    return () => clearTimeout(t);
  }, [running, ex.engine]);

  return (
    <Modal
      title="Excel出力"
      testId="export-dialog"
      size="md"
      onClose={props.onClose}
      locked={running}
      focusKey={ex.phase}
      initialFocus={running ? '[data-testid=export-cancel]' : done ? '[data-testid=export-result] .btn-primary, [data-testid=export-close]' : '[data-testid=export-confirm]'}
      footer={
        running ? (
          <>
            <span className="modal-footer-note muted">出力が終わるまでお待ちください</span>
            <button type="button" className="btn" onClick={props.onCancelExport} data-testid="export-cancel">
              中止
            </button>
          </>
        ) : done ? (
          <button type="button" className="btn btn-primary btn-wide" onClick={props.onClose} data-testid="export-close">
            閉じる
          </button>
        ) : (
          <>
            <button type="button" className="btn" onClick={props.onClose}>
              キャンセル
            </button>
            <button type="button" className="btn btn-primary btn-wide" data-testid="export-confirm" disabled={invalid} onClick={() => props.onExport(finalName, effectiveEngine)}>
              {ex.phase === 'error' ? 'もう一度出力' : '出力'}
            </button>
          </>
        )
      }
    >
      <form
        className="export-form"
        onSubmit={(e) => {
          e.preventDefault();
          if (!running && !done && !invalid) props.onExport(finalName, effectiveEngine);
        }}
      >
        <label className="form-row">
          <span className="form-label">ファイル名</span>
          <span className="filename-box">
            <input
              data-testid="export-filename"
              value={name}
              disabled={running || done}
              onChange={(e) => setName(e.target.value)}
              spellCheck={false}
              aria-invalid={invalid}
            />
          </span>
        </label>
        {invalid && <div className="form-error">ファイル名を入力してください</div>}

        {disk && (
          <div className="disk-notice" role="status" data-testid="export-disk-notice" data-temp-low={disk.tempLow ? 'true' : 'false'} data-output-low={disk.outputLow ? 'true' : 'false'}>
            <span className="disk-notice-icon" aria-hidden>
              ⚠
            </span>
            <div className="disk-notice-body">
              {disk.tempMessage && <p data-testid="export-disk-temp">{disk.tempMessage}</p>}
              {disk.recommendation && (
                <p className="disk-notice-advice" data-testid="export-disk-advice">
                  {disk.recommendation}
                </p>
              )}
              {disk.outputMessage && <p data-testid="export-disk-output">{disk.outputMessage}</p>}
            </div>
          </div>
        )}

        <div className="form-row engine-set" role="radiogroup" aria-labelledby="export-engine-label">
          <span className="form-label" id="export-engine-label">
            出力方法
          </span>
          <div className="engine-options">
          <label className={`engine-option${effectiveEngine === 'excel' ? ' active' : ''}${excelOk ? '' : ' disabled'}`}>
            <input
              type="radio"
              name="engine"
              data-testid="export-engine-excel"
              checked={effectiveEngine === 'excel'}
              disabled={!excelOk || running || done}
              onChange={() => setPicked('excel')}
              onClick={() => setPicked('excel')} // clicking the radio that is already selected is a choice too (the default may still change)
            />
            <span>
              <b>Excel本体で出力（推奨・最も忠実）</b>
              <span className="engine-note">
                {excelOk
                  ? 'このPCの Excel で列の操作を再現して保存します。グラフ・ピボット・条件付き書式も確実に保持されます。'
                  : serverDown
                    ? 'ローカルサーバーに接続できないため使えません（start.bat から起動し直すと使えます）。'
                    : serverMode
                      ? 'Excel がインストールされていないため使えません。'
                      : 'スタンドアロンで開いているため使えません（start.bat から起動すると使えます）。'}
                {props.includeSourceSheets && (
                  <span className="engine-sheets-hint" data-testid="export-engine-sheets-hint">
                    {' '}
                    他のファイルのシートをコピーしているため、そのシートのグラフ・ピボットテーブルは、この方法がもっとも確実に保たれます。
                  </span>
                )}
              </span>
            </span>
          </label>
          <label className={`engine-option${effectiveEngine === 'js' ? ' active' : ''}`}>
            <input type="radio" name="engine" data-testid="export-engine-js" checked={effectiveEngine === 'js'} disabled={running || done} onChange={() => setPicked('js')} onClick={() => setPicked('js')} />
            <span>
              <b>内蔵エンジンで出力</b>
              <span className="engine-note">Excel なしで出力します。結合した表の数式・書式・列幅を保ちます。</span>
            </span>
          </label>
          </div>
        </div>

        <div className="form-row export-dest">
          <span className="form-label">保存先</span>
          <span className="dest-text">
            {serverMode ? (
              <>
                出力フォルダ <code title={server.status!.outputDir}>{server.status!.outputDir}</code>
              </>
            ) : serverDown ? (
              'ローカルサーバーに接続できないため、ブラウザの保存（保存先を選ぶ画面、または「ダウンロード」フォルダ）になります'
            ) : (
              '保存先を選ぶ画面が開きます（開かない場合はダウンロードフォルダに保存）'
            )}
          </span>
        </div>
        <p className="muted export-summary">
          結合後の変更 {props.editCount} 件を反映して <b>{finalName}</b> を作ります。元のファイルは変更されません。
        </p>
      </form>

      {running && (
        <div className="export-progress" data-testid="export-progress" aria-live="polite">
          <div className="export-progress-text">
            <span className="spinner spinner-sm" aria-hidden /> {ex.message || '出力しています…'}
          </div>
          <ProgressBar ratio={ex.ratio} />
          {slow && (
            <p className="export-slow" data-testid="export-slow" role="note">
              Excel本体での出力に時間がかかっています。大きなファイルでは数分かかることがあります。待てないときは「中止」を押し、「内蔵エンジンで出力」を選んでやり直せます。
            </p>
          )}
        </div>
      )}

      {ex.phase === 'error' && (
        <div className="card-error" role="alert" data-testid="export-error">
          {ex.error}
        </div>
      )}

      {done && (
        <div className="export-result" data-testid="export-result" role="status">
          <div className="result-title">
            <span className="result-check" aria-hidden>
              ✓
            </span>
            {ex.savedPath ? (
              <>
                保存しました: <code className="result-path">{ex.savedPath}</code>
              </>
            ) : ex.delivered === 'picker' ? (
              <>保存しました: {ex.fileName}</>
            ) : (
              <>ダウンロードしました: {ex.fileName}</>
            )}
          </div>
          <div className="result-engine muted" data-testid="export-result-engine">
            {ex.engine === 'excel' ? 'Excel本体で出力しました' : '内蔵エンジンで出力しました'}
          </div>
          {ex.fellBack && <div className="card-note">Excel本体での出力に失敗したため、内蔵エンジンで出力しました。</div>}
          {ex.warnings && ex.warnings.length > 0 && (
            <div className="export-warnings" data-testid="export-warnings" role="note">
              <div className="export-warnings-title">
                <span aria-hidden>⚠</span> 出力したファイルで確認してください（{ex.warnings.length} 件）
              </div>
              <ul>
                {ex.warnings.map((w, i) => (
                  <li key={i}>{w}</li>
                ))}
              </ul>
            </div>
          )}
          <div className="result-actions">
            {ex.savedPath ? (
              <>
                <button type="button" className="btn btn-primary" onClick={props.onOpenFile} data-testid="export-open-file">
                  Excelで開く
                </button>
                <button type="button" className="btn" onClick={props.onOpenFolder} data-testid="export-open-folder">
                  フォルダを開く
                </button>
              </>
            ) : (
              <button type="button" className="btn" onClick={props.onDownloadAgain} data-testid="export-download-again">
                もう一度ダウンロード
              </button>
            )}
          </div>
        </div>
      )}
    </Modal>
  );
}
