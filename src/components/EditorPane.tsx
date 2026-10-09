/**
 * Main area after the merge: hint bar + formula bar (grid G's FormulaBar) + the grid (G's Grid), adapted from
 * the MergedDoc of the present snapshot (src/state/gridAdapter.ts).
 */
import { memo, useCallback, useMemo, useState, type Ref } from 'react';
import { Grid } from './grid/Grid';
import { FormulaBar } from './grid/FormulaBar';
import { cellEditText } from './grid/cellview';
import type { CellPos, EditingState, GridHandle } from './grid/types';
import { makeStyleCache, ptToPx } from '../core/display/styles';
import type { Cell, SourceBook } from '../core/types';
import { APP_SOURCE_COLOR } from '../state/helpers';
import { buildGridModel, cellAddress, makeGetCell } from '../state/gridAdapter';
import type { Services } from '../state/services';
import type { ColSelectionState, GridPos, MergedState } from '../state/types';

export interface EditorPaneProps {
  merged: MergedState;
  services: Services;
  baseId: string | null;
  keyRoles: Map<number, string>;
  selection: ColSelectionState;
  active: GridPos | null;
  liveWidths: Record<string, number>;
  highlightRows: ReadonlySet<number>;
  gridRef: Ref<GridHandle>;
  onFocusGrid: () => void;
  onSelection: (sel: ColSelectionState) => void;
  onActive: (pos: CellPos) => void;
  onMoveColumns: (ids: string[], gap: number) => void;
  onResizeColumn: (id: string, px: number, commit: { from: number } | null) => void;
  onCommitEdit: (pos: CellPos, text: string) => void;
  onDeleteColumns: (ids: string[]) => void;
  onInsertColumns: (index: number, count: number) => void;
  onResetColumnWidths: (ids: string[]) => void;
  onUndo: () => void;
  onRedo: () => void;
  onNotice?: (text: string) => void;
}

export const EditorPane = memo(function EditorPane(props: EditorPaneProps) {
  const { merged, services, baseId, keyRoles, selection, active, liveWidths } = props;
  const doc = merged.history.present.doc;
  const [editing, setEditing] = useState<EditingState | null>(null);

  const textBook: SourceBook = useMemo(() => ({ ...merged.book, styles: doc.styles }), [merged.book, doc.styles]);
  const headerText = useCallback(
    (cell: Cell | undefined) => {
      try {
        return services.cellText(textBook, cell);
      } catch {
        return cell?.v === null || cell?.v === undefined ? '' : String(cell.v);
      }
    },
    [services, textBook],
  );
  const defaultColWidth = useMemo(() => merged.book.sheets.find((s) => s.name === doc.sheetName)?.defaultColWidth, [merged.book, doc.sheetName]);
  // the doc-derived model; live widths (while a resize handle is dragged) only replace `columns`, so getCell
  // and its view cache survive a resize gesture
  const baseModel = useMemo(
    () => buildGridModel(doc, { baseId, keyRoles, headerText, defaultColWidth }),
    [doc, baseId, keyRoles, headerText, defaultColWidth],
  );
  const model = useMemo(() => {
    const ids = Object.keys(liveWidths);
    if (!ids.length) return baseModel;
    return { ...baseModel, columns: baseModel.columns.map((c) => (liveWidths[c.id] !== undefined ? { ...c, width: liveWidths[c.id] } : c)) };
  }, [baseModel, liveWidths]);
  const rows = useMemo(() => baseModel.rows.filter((r) => !doc.rowProps[r - 1]?.hidden), [baseModel.rows, doc.rowProps]);
  const getCell = useMemo(() => makeGetCell(doc, baseModel, (ci, row) => services.formulaForDisplay(doc, ci, row)), [doc, baseModel, services]);
  const cssOf = useMemo(() => makeStyleCache(doc.styles), [doc.styles]);
  const hasHeights = useMemo(() => doc.rowProps.some((p) => p?.height !== undefined), [doc.rowProps]);
  const rowHeight = useMemo(
    () =>
      hasHeights
        ? (row: number) => {
            const h = doc.rowProps[row - 1]?.height;
            return h === undefined ? undefined : Math.max(18, Math.min(120, ptToPx(h)));
          }
        : undefined,
    [hasHeights, doc.rowProps],
  );

  // formula bar
  const colIdx = active ? (model.colIndex.get(active.col) ?? -1) : -1;
  const activeCol = colIdx >= 0 ? model.columns[colIdx] : undefined;
  let fbValue = '';
  if (active && activeCol) {
    const view = getCell(active.row, activeCol);
    fbValue = cellEditText(view, cssOf(view?.styleId));
  }
  const address = active && colIdx >= 0 ? cellAddress(colIdx, active.row) : '';
  const nSel = selection.cols.size;

  return (
    <>
      <div className="table-hint blank-hint editor-hint" data-testid="grid-hint">
        <span>
          <b>列見出しをクリックで選択</b>（Ctrl・Shiftで複数）→ <b>ドラッグで移動</b>。見出しの上の<b>色の帯</b>やファイルの<b>カード</b>は、ファイルの列をまとめて選択（クリック）・移動（ドラッグ）できます。右クリックで挿入・削除、Delete キーで削除、Ctrl+Z で元に戻す。
        </span>
        {nSel > 0 && <span className="hint-sel">{nSel} 列を選択中</span>}
      </div>
      <FormulaBar
        pos={active && activeCol ? active : null}
        address={address}
        value={fbValue}
        mirror={editing && active && editing.pos.row === active.row && editing.pos.col === active.col ? editing.text : null}
        onCommitEdit={props.onCommitEdit}
        onDone={props.onFocusGrid}
        placeholder="セルを選ぶと内容を表示します"
      />
      <div className="grid-host">
        <Grid
          ref={props.gridRef}
          columns={model.columns}
          sources={model.sources}
          appSourceColor={APP_SOURCE_COLOR}
          appSourceName="アプリで追加"
          rows={rows}
          headerRow={doc.headerRow}
          frozenCols={model.frozenCols}
          getCell={getCell}
          cssOf={cssOf}
          dataVersion={doc.version}
          rowHeight={rowHeight}
          selection={selection}
          onSelectionChange={props.onSelection}
          active={active}
          onActiveChange={props.onActive}
          onMoveColumns={props.onMoveColumns}
          onResizeColumn={props.onResizeColumn}
          onCommitEdit={props.onCommitEdit}
          onDeleteColumns={props.onDeleteColumns}
          onInsertColumns={props.onInsertColumns}
          onResetColumnWidths={props.onResetColumnWidths}
          onUndo={props.onUndo}
          onRedo={props.onRedo}
          onEditingChange={setEditing}
          onNotice={props.onNotice}
          highlightRows={props.highlightRows}
          highlightLabel="未一致（どのファイルにも該当する行がありません）"
          ariaLabel={`結合結果「${doc.sheetName}」`}
          emptyText="データ行がありません"
        />
      </div>
    </>
  );
});
