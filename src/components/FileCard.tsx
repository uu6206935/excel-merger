/**
 * One loaded file (like the reference UploadCard): colour strip, name + row count, ★基準 toggle, クリア,
 * シート / 見出し行 selectors + 最終行 input, format tags (xlsx / csv + 文字コード / 値のみ …), detected key columns
 * and a 5-row preview of the data table.
 */
import { memo, useEffect, useMemo, useState } from 'react';
import type { Cell, SourceBook, TableRegion } from '../core/types';
import { FORMAT_LABEL, colLetter, encodingLabel, formatSize, headerRowChoices, headerRowIsNumeric, lastRowBounds, regionRowCount, sheetByName } from '../state/helpers';
import { fileKeyLine, type FileKeyColumn } from '../state/keysPlan';
import { HelpPanel } from './HelpPanel';
import type { FileEntry } from '../state/types';

const PREVIEW_ROWS = 5;
const PREVIEW_COLS = 30;

export interface FileCardProps {
  index: number;
  file: FileEntry;
  isBase: boolean;
  /** only one file: the base toggle makes no sense yet */
  single: boolean;
  keyCols: FileKeyColumn[];
  /** no key selected: the file takes no part in the merge (badge + 「キーを選ぶ…」) */
  excluded?: boolean;
  /** 「キーを選ぶ…」: opens the key detail modal on this file's tab */
  onPickKey?: (id: string) => void;
  /** analysis still running for this file */
  analyzing: boolean;
  /** merged phase: settings are frozen */
  locked: boolean;
  compact: boolean;
  cellText: (book: SourceBook, cell: Cell | undefined) => string;
  onSetBase: (id: string) => void;
  onRemove: (id: string) => void;
  onSheet: (id: string, sheet: string) => void;
  onHeaderRow: (id: string, row: number) => void;
  /** 最終行; null = automatic */
  onLastRow: (id: string, row: number | null) => void;
  /**
   * The card can be dragged (FileCardRow) / moved with Alt+← →. Merged phase: false for a file without columns in the table.
   * Default true.
   */
  reorderable?: boolean;
  /** merged phase: the current column selection is exactly this file's column group (「選択中」 look) */
  active?: boolean;
  /** files phase: the card the user clicked last (highlight only) */
  picked?: boolean;
  /** tooltip of the title line: what a click / a drag does here */
  gestureHint?: string;
}

function safeText(fn: FileCardProps['cellText'], book: SourceBook, cell: Cell | undefined): string {
  try {
    return fn(book, cell);
  } catch {
    return cell?.v === null || cell?.v === undefined ? '' : String(cell.v);
  }
}

export const MiniPreview = memo(function MiniPreview(props: {
  book: SourceBook;
  region: TableRegion;
  keyCols: ReadonlySet<number>;
  cellText: FileCardProps['cellText'];
}) {
  const { book, region, keyCols, cellText } = props;
  const sheet = sheetByName(book, region.sheet);
  const model = useMemo(() => {
    if (!sheet) return null;
    const last = Math.min(region.lastCol, region.firstCol + PREVIEW_COLS - 1);
    const cols: number[] = [];
    for (let c = region.firstCol; c <= last; c++) cols.push(c);
    const hdr = sheet.rows[region.headerRow - 1] ?? [];
    const head = cols.map((c) => safeText(cellText, book, hdr[c - 1]) || colLetter(c));
    const body: Array<{ row: number; cells: string[] }> = [];
    for (let r = region.firstDataRow; r <= region.lastDataRow && body.length < PREVIEW_ROWS; r++) {
      const row = sheet.rows[r - 1] ?? [];
      body.push({ row: r, cells: cols.map((c) => safeText(cellText, book, row[c - 1])) });
    }
    return { cols, head, body, more: region.lastCol - last };
  }, [sheet, region, book, cellText]);
  if (!model) return null;
  return (
    <div className="matching-scroll mini-preview" aria-label="先頭5行のプレビュー">
      <table className="matching-table">
        <thead>
          <tr>
            <th className="m-no">{region.headerRow}</th>
            {model.head.map((h, i) => (
              <th key={model.cols[i]} className={keyCols.has(model.cols[i]) ? 'm-key' : undefined} title={h}>
                {h}
              </th>
            ))}
            {model.more > 0 && <th className="muted">…他{model.more}列</th>}
          </tr>
        </thead>
        <tbody>
          {model.body.map((r) => (
            <tr key={r.row}>
              <td className="m-no">{r.row}</td>
              {r.cells.map((t, i) => (
                <td key={model.cols[i]} className={keyCols.has(model.cols[i]) ? 'm-key' : undefined} title={t}>
                  {t}
                </td>
              ))}
              {model.more > 0 && <td />}
            </tr>
          ))}
          {model.body.length === 0 && (
            <tr>
              <td className="muted" colSpan={model.cols.length + 1}>
                データ行がありません
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
});

/** 最終行: the last data row of the table (Enter / leaving the field applies it; 「自動」 goes back to the detected one). */
function LastRowInput(props: {
  index: number;
  value: number;
  min: number;
  max: number;
  manual: boolean;
  disabled: boolean;
  lockTitle?: string;
  onCommit: (row: number | null) => void;
}) {
  const { index, value, min, max, manual, disabled, lockTitle, onCommit } = props;
  const [draft, setDraft] = useState(String(value));
  useEffect(() => setDraft(String(value)), [value]);
  const commit = () => {
    const t = draft.normalize('NFKC').replace(/[,\s行目]/g, '');
    const n = Number(t);
    if (!t || !Number.isInteger(n)) {
      setDraft(String(value));
      return;
    }
    const clamped = Math.max(min, Math.min(max, n));
    setDraft(String(clamped));
    if (clamped !== value) onCommit(clamped);
  };
  return (
    <label className="folder-select-label last-row-label">
      <span>最終行</span>
      <input
        type="text"
        inputMode="numeric"
        className="last-row-input"
        data-testid={`last-row-input-${index}`}
        value={draft}
        disabled={disabled}
        aria-label="最終行（表のデータの最後の行）"
        title={lockTitle ?? `表のデータの最後の行（${min}〜${max}）。合計・平均などの集計行は含めません。Enter で確定`}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={() => {
          if (draft !== String(value)) commit();
        }}
        onKeyDown={(e) => {
          if (e.nativeEvent.isComposing || e.keyCode === 229) return;
          if (e.key === 'Enter') {
            e.preventDefault();
            commit();
          } else if (e.key === 'Escape') {
            e.preventDefault();
            setDraft(String(value));
          } else if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
            e.preventDefault();
            const n = Number(draft.normalize('NFKC'));
            const base = Number.isInteger(n) ? n : value;
            setDraft(String(Math.max(min, Math.min(max, base + (e.key === 'ArrowUp' ? 1 : -1)))));
          }
        }}
      />
      {manual && !disabled && (
        <button type="button" className="btn btn-xs last-row-auto" data-testid={`last-row-auto-${index}`} onClick={() => onCommit(null)} title="最終行を自動の判定に戻す">
          自動
        </button>
      )}
    </label>
  );
}

/** Short text of a row for the 見出し行 options ('2 行目（No / 学生ID / 氏名 …）'). */
function rowLabel(book: SourceBook, sheetName: string, row: number, cellText: FileCardProps['cellText']): string {
  const sheet = sheetByName(book, sheetName);
  const cells = sheet?.rows[row - 1] ?? [];
  const texts: string[] = [];
  for (const c of cells) {
    const t = c ? safeText(cellText, book, c).trim() : '';
    if (t) texts.push(t.length > 8 ? t.slice(0, 8) + '…' : t);
    if (texts.length >= 3) break;
  }
  return texts.length ? `${row} 行目（${texts.join(' / ')}${texts.length >= 3 ? ' …' : ''}）` : `${row} 行目（空）`;
}

export const FileCard = memo(function FileCard(props: FileCardProps) {
  const { index, file, isBase, single, keyCols, analyzing, locked, compact, cellText } = props;
  const excluded = !!props.excluded && !isBase;
  const book = file.book;
  const rows = regionRowCount(file.region);
  const keySet = useMemo(() => new Set(keyCols.map((k) => k.col)), [keyCols]);
  const sheet = book && file.sheet ? sheetByName(book, file.sheet) : undefined;
  const headerChoices = useMemo(() => headerRowChoices(sheet, file.region?.headerRow), [sheet, file.region?.headerRow]);
  const headerLabels = useMemo(
    () => (book && file.sheet ? headerChoices.map((r) => rowLabel(book, file.sheet!, r, cellText)) : []),
    [book, file.sheet, headerChoices, cellText],
  );
  const visibleSheets = book ? book.sheets.filter((s) => s.state === 'visible' || s.name === file.sheet) : [];
  const lastBounds = lastRowBounds(sheet, file.region);
  const numericHeader = useMemo(() => !!book && file.status === 'ready' && headerRowIsNumeric(book, file.region), [book, file.status, file.region]);
  const encoding = book ? encodingLabel(book.encoding) : null;
  const reorderable = props.reorderable !== false;
  const cls = [
    'upload-card',
    'file-card',
    file.status === 'ready' && !file.error ? 'loaded' : '',
    file.status === 'error' || file.error ? 'has-error' : '',
    isBase ? 'is-base' : '',
    excluded ? 'is-excluded' : '',
    compact ? 'compact' : '',
    props.active ? 'is-active' : '',
    props.picked ? 'is-picked' : '',
  ]
    .filter(Boolean)
    .join(' ');
  const lockTitle = locked ? '結合後は変更できません（「結合をやり直す」で戻れます）' : undefined;

  return (
    <section
      className={cls}
      data-testid={`file-card-${index}`}
      data-file-id={file.id}
      data-reorderable={reorderable ? '' : undefined}
      data-active={props.active ? 'true' : undefined}
      // focusable so that Alt+← / Alt+→ can move the card (FileCardRow)
      tabIndex={0}
      aria-keyshortcuts={reorderable ? 'Alt+ArrowLeft Alt+ArrowRight' : undefined}
      aria-label={`${file.fileName}${isBase ? '（基準）' : ''}`}
    >
      <div className="card-strip" style={{ background: file.color }} aria-hidden />
      <header className="card-head">
        {reorderable && (
          <span className="card-grip" aria-hidden title={props.gestureHint}>
            ⋮⋮
          </span>
        )}
        <h2 title={props.gestureHint ? `${file.fileName}\n${props.gestureHint}` : file.fileName}>{file.fileName}</h2>
        {file.status === 'ready' && file.region && <span className="count-badge">{rows.toLocaleString()} 行</span>}
        {compact ? (
          isBase && (
            <span className="tag tag-base card-head-actions" title="結合結果の元になるブック">
              ★ 基準
            </span>
          )
        ) : (
        <span className="card-head-actions">
          <button
            type="button"
            className={`btn btn-sm base-toggle${isBase ? ' active' : ''}`}
            data-testid={`base-toggle-${index}`}
            aria-pressed={isBase}
            disabled={locked || file.status !== 'ready' || (single && isBase)}
            onClick={() => props.onSetBase(file.id)}
            title={
              lockTitle ??
              (isBase ? 'このファイルの表に、他のファイルの列を追加します（出力の元になるブック）' : 'このファイルを基準（出力の元になるブック）にします')
            }
          >
            {isBase ? '★ 基準' : '☆ 基準にする'}
          </button>
          <button
            type="button"
            className="btn btn-sm card-clear"
            data-testid={`file-clear-${index}`}
            onClick={() => props.onRemove(file.id)}
            disabled={locked}
            title={lockTitle ?? 'このファイルを外します'}
            aria-label={`${file.fileName} をクリア`}
          >
            クリア
          </button>
        </span>
        )}
      </header>

      {file.status === 'loading' && (
        <div className="card-loading">
          <span className="spinner spinner-sm" aria-hidden /> {file.message ?? '読み込み中…'}
        </div>
      )}

      {book && file.status === 'ready' && !compact && (
        <div className="card-selects">
          <label className="folder-select-label">
            <span>シート</span>
            <select
              data-testid={`sheet-select-${index}`}
              value={file.sheet ?? ''}
              disabled={locked}
              title={lockTitle ?? '表（データ）が入っているシート'}
              onChange={(e) => props.onSheet(file.id, e.target.value)}
            >
              {visibleSheets.map((s) => (
                <option key={s.name} value={s.name}>
                  {s.name}
                </option>
              ))}
            </select>
          </label>
          <label className="folder-select-label">
            <span>見出し行</span>
            <select
              data-testid={`header-row-select-${index}`}
              value={file.region?.headerRow ?? ''}
              disabled={locked}
              title={lockTitle ?? '表の見出し（項目名）が書かれている行'}
              onChange={(e) => props.onHeaderRow(file.id, Number(e.target.value))}
            >
              {!file.region && <option value="">選んでください</option>}
              {headerChoices.map((r, i) => (
                <option key={r} value={r}>
                  {headerLabels[i]}
                </option>
              ))}
            </select>
          </label>
          {file.region && lastBounds && (
            <LastRowInput
              index={index}
              value={file.region.lastDataRow}
              min={lastBounds.min}
              max={lastBounds.max}
              manual={!!file.lastRowManual}
              disabled={locked}
              lockTitle={lockTitle}
              onCommit={(row) => props.onLastRow(file.id, row)}
            />
          )}
        </div>
      )}

      {book && (
        <div className="file-info">
          <span className="tag">{FORMAT_LABEL[book.format] ?? book.format}</span>
          {encoding && (
            <span className="tag tag-muted" data-testid={`file-encoding-${index}`} title="文字コード（自動判定）">
              {encoding}
            </span>
          )}
          {book.fidelity === 'values' && (
            <span className="tag tag-warn" title="書式・数式は引き継がれず、値だけを使います">
              値のみ
            </span>
          )}
          {book.features.vba && <span className="tag">マクロ</span>}
          {book.features.charts > 0 && <span className="tag">グラフ {book.features.charts}</span>}
          <span className="muted">{formatSize(file.size || book.size)}</span>
          {sheet && compact && <span className="muted">シート「{sheet.name}」</span>}
        </div>
      )}

      {excluded && book && file.status === 'ready' && file.region && (
        <div className="card-excluded" data-testid={`file-excluded-${index}`} role="status">
          <span className="tag tag-warn card-excluded-badge" title="結合キーが選ばれていないため、このファイルの列は今回の結合に取り込まれません">
            結合対象外（結合キーが見つかりません）
          </span>
          {!locked && props.onPickKey && (
            <button
              type="button"
              className="btn btn-sm card-excluded-btn"
              data-testid={`file-pick-key-${index}`}
              onClick={() => props.onPickKey?.(file.id)}
              aria-haspopup="dialog"
              title="詳細設定を開いて、このファイルの結合キーを選びます（列を選んで追加もできます）"
            >
              キーを選ぶ…
            </button>
          )}
        </div>
      )}

      {book && file.status === 'ready' && file.region && !excluded && (
        <div className="card-meta card-keys" data-testid={`file-keys-${index}`} title={keyCols.length ? `キー ${fileKeyLine(keyCols)}` : undefined}>
          {keyCols.length ? (
            <>
              <span className="card-keys-label">キー</span> {fileKeyLine(keyCols)}
            </>
          ) : analyzing ? (
            <span className="muted">キーを確認中…</span>
          ) : single ? (
            <span className="muted">もう1つファイルを追加すると結合キーを探します</span>
          ) : (
            <span className="muted">結合キーなし</span>
          )}
        </div>
      )}

      {(file.error || file.status === 'error') && (
        <div className="card-error" role="alert">
          {file.error ?? '読み込めませんでした'}
        </div>
      )}
      {file.note && file.status === 'ready' && !compact && (
        <div className="card-note" data-testid={`file-note-${index}`} role="note">
          {file.note}
        </div>
      )}
      {numericHeader && !compact && (
        <div className="card-note" data-testid={`file-numeric-header-${index}`} role="note" title="見出しのない表では、1行目のデータは結合の対象になりません。「見出し行」で別の行を選ぶこともできます">
          {file.region?.headerRow} 行目が数字だけです。見出しのない表の場合、この行はデータとして結合されません。
        </div>
      )}

      {book && file.region && !compact && <MiniPreview book={book} region={file.region} keyCols={keySet} cellText={cellText} />}
    </section>
  );
});

/** The dashed 「＋ ここにExcelファイルをドロップ」 card at the end of the row. */
export function AddFileCard({ onPick, dragging, compact }: { onPick: () => void; dragging: boolean; compact: boolean }) {
  return (
    <section className={`upload-card add-card${dragging ? ' dragging' : ''}${compact ? ' compact' : ''}`} data-testid="drop-zone" aria-label="ファイルを追加">
      <div className="add-card-inner">
        <span className="add-plus" aria-hidden>
          ＋
        </span>
        <span className="dropzone-text">ここにExcelファイルをドロップ（複数可）</span>
        <button type="button" className="btn btn-primary btn-sm" onClick={onPick} title="Excelファイルを選んで追加します（ドラッグ＆ドロップでも追加できます）">
          ファイルを選択
        </button>
      </div>
    </section>
  );
}

/** Empty state: big dropzone with the collapsible 「使い方」 (4 steps). */
export function EmptyState({ onPick, dragging }: { onPick: () => void; dragging: boolean }) {
  return (
    <div className={`empty-drop${dragging ? ' dragging' : ''}`} data-testid="drop-zone">
      <svg className="empty-icon" viewBox="0 0 48 48" aria-hidden>
        <rect x="9" y="5" width="26" height="34" rx="3" fill="#fff" stroke="#9fb5d8" strokeWidth="2" />
        <rect x="15" y="11" width="26" height="34" rx="3" fill="#eef4ff" stroke="#1f6feb" strokeWidth="2" />
        <path d="M21 21h14M21 27h14M21 33h9" stroke="#1f6feb" strokeWidth="2" strokeLinecap="round" />
      </svg>
      <h2>ここにExcelファイルをドロップ（複数可）</h2>
      <p className="muted">.xlsx / .xlsm / .xls / .xlsb / .csv / .ods に対応。ファイルはこのPCの外に送信されません。</p>
      <button type="button" className="btn btn-primary btn-big" onClick={onPick} title="Excelファイルを選んで読み込みます（ドラッグ＆ドロップでも読み込めます）">
        ファイルを選択
      </button>
      <HelpPanel />
    </div>
  );
}
