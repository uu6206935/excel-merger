/**
 * 「列の一覧」: vertical list of the current columns with their source colour. Drag to reorder (= the same
 * moveColumns op as the grid; dragging a selected column moves the whole selection), click to select /
 * scroll the grid, double-click (or F2) to rename the header, Alt+↑/↓ to move with the keyboard.
 */
import { memo, useRef, useState, type DragEvent, type KeyboardEvent } from 'react';
import type { ColumnListItem } from '../state/gridAdapter';
import type { PlannedColumn } from '../state/keysPlan';

const DND_TYPE = 'application/x-excel-merger-column';

export interface ColumnListProps {
  items: ColumnListItem[];
  selected: ReadonlySet<string>;
  onMove: (ids: string[], gap: number) => void;
  onSelect: (id: string, mods: { ctrl: boolean; shift: boolean }) => void;
  onRename: (id: string, text: string) => void;
}

export const ColumnList = memo(function ColumnList({ items, selected, onMove, onSelect, onRename }: ColumnListProps) {
  const [drop, setDrop] = useState<number | null>(null);
  const [dragIds, setDragIds] = useState<string[] | null>(null);
  const [renaming, setRenaming] = useState<string | null>(null);
  // roving tabindex: Tab enters the list once (on the item last focused / the first selected one), the arrow keys move inside it
  const [tabStop, setTabStop] = useState<string | null>(null);
  const listRef = useRef<HTMLUListElement>(null);
  const tabStopId = items.some((it) => it.id === tabStop) ? tabStop : (items.find((it) => selected.has(it.id)) ?? items[0])?.id;

  const idsFor = (id: string) => (selected.has(id) ? items.filter((it) => selected.has(it.id)).map((it) => it.id) : [id]);

  const gapFromEvent = (e: DragEvent<HTMLElement>, index: number) => {
    const r = e.currentTarget.getBoundingClientRect();
    return e.clientY < r.top + r.height / 2 ? index : index + 1;
  };

  const onDragStart = (e: DragEvent<HTMLLIElement>, id: string) => {
    const ids = idsFor(id);
    setDragIds(ids);
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData(DND_TYPE, JSON.stringify(ids));
    e.dataTransfer.setData('text/plain', ids.join(','));
  };
  const onDragOver = (e: DragEvent<HTMLLIElement>, index: number) => {
    if (!dragIds) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    const g = gapFromEvent(e, index);
    if (g !== drop) setDrop(g);
  };
  const endDrag = () => {
    setDrop(null);
    setDragIds(null);
  };
  const onDrop = (e: DragEvent<HTMLLIElement>, index: number) => {
    if (!dragIds) return;
    e.preventDefault();
    e.stopPropagation();
    const gap = gapFromEvent(e, index);
    const ids = dragIds;
    endDrag();
    onMove(ids, gap);
  };

  const onKeyDown = (e: KeyboardEvent<HTMLLIElement>, it: ColumnListItem) => {
    if (renaming) return;
    if (e.altKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) {
      e.preventDefault();
      const ids = idsFor(it.id);
      const idx = items.filter((x) => ids.includes(x.id)).map((x) => x.index);
      const gap = e.key === 'ArrowUp' ? Math.max(0, Math.min(...idx) - 1) : Math.min(items.length, Math.max(...idx) + 2);
      onMove(ids, gap);
      requestAnimationFrame(() => listRef.current?.querySelector<HTMLElement>(`[data-col-id="${CSS.escape(it.id)}"]`)?.focus());
      return;
    }
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      const next = e.currentTarget[e.key === 'ArrowDown' ? 'nextElementSibling' : 'previousElementSibling'] as HTMLElement | null;
      next?.focus();
      return;
    }
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      onSelect(it.id, { ctrl: e.ctrlKey || e.metaKey, shift: e.shiftKey });
      return;
    }
    if (e.key === 'F2') {
      e.preventDefault();
      setRenaming(it.id);
    }
  };

  return (
    <ul className="side-list column-list" ref={listRef} aria-label="列の一覧（ドラッグで並べ替え）" onDragLeave={(e) => e.currentTarget === e.target && setDrop(null)}>
      {items.map((it) => {
        const sel = selected.has(it.id);
        const dragging = dragIds?.includes(it.id);
        const cls = ['col-item', sel ? 'is-sel' : '', dragging ? 'is-dragging' : '', drop === it.index ? 'drop-before' : '', drop === it.index + 1 && it.index === items.length - 1 ? 'drop-after' : '']
          .filter(Boolean)
          .join(' ');
        return (
          <li
            key={it.id}
            className={cls}
            data-col-id={it.id}
            data-testid={`column-list-item-${it.index}`}
            draggable={renaming !== it.id}
            tabIndex={it.id === tabStopId ? 0 : -1}
            onFocus={() => setTabStop(it.id)}
            aria-selected={sel}
            onDragStart={(e) => onDragStart(e, it.id)}
            onDragOver={(e) => onDragOver(e, it.index)}
            onDrop={(e) => onDrop(e, it.index)}
            onDragEnd={endDrag}
            onClick={(e) => onSelect(it.id, { ctrl: e.ctrlKey || e.metaKey, shift: e.shiftKey })}
            onKeyDown={(e) => onKeyDown(e, it)}
            title={`${it.letter}列: ${it.title}（${it.sourceLabel}）— ドラッグで移動／ダブルクリックで名前を変更`}
          >
            <span className="col-grip" aria-hidden>
              ⋮⋮
            </span>
            <span className="col-color" style={{ background: it.color }} aria-hidden />
            <span className="col-letter">{it.letter}</span>
            {renaming === it.id ? (
              <RenameInput
                initial={it.title.startsWith('（') ? '' : it.title}
                onDone={(text) => {
                  setRenaming(null);
                  if (text !== null && text !== it.title) onRename(it.id, text);
                }}
              />
            ) : (
              <span
                className="col-title"
                onDoubleClick={(e) => {
                  e.stopPropagation();
                  setRenaming(it.id);
                }}
              >
                {it.title}
              </span>
            )}
            {it.keyRole && <span className="key-badge">キー</span>}
            {it.hidden && <span className="tag">非表示</span>}
          </li>
        );
      })}
    </ul>
  );
});

function RenameInput({ initial, onDone }: { initial: string; onDone: (text: string | null) => void }) {
  const [text, setText] = useState(initial);
  const done = useRef(false);
  const finish = (v: string | null) => {
    if (done.current) return;
    done.current = true;
    onDone(v);
  };
  return (
    <input
      className="col-rename"
      autoFocus
      value={text}
      aria-label="見出しの名前"
      onChange={(e) => setText(e.target.value)}
      onClick={(e) => e.stopPropagation()}
      onKeyDown={(e) => {
        if (e.nativeEvent.isComposing || e.keyCode === 229) return;
        e.stopPropagation();
        if (e.key === 'Enter') finish(text);
        else if (e.key === 'Escape') finish(null);
      }}
      onBlur={() => finish(text)}
    />
  );
}

/** Before the merge: the planned column order (read-only). */
export function PlannedColumnList({ items, colors }: { items: PlannedColumn[]; colors: Record<string, string> }) {
  return (
    <ul className="side-list column-list is-readonly" aria-label="結合後の列（予定）">
      {items.map((c, i) => (
        <li key={c.key} className={`col-item${c.imported ? ' is-new' : ''}`}>
          <span className="col-color" style={{ background: colors[c.sourceId] ?? '#999' }} aria-hidden />
          <span className="col-letter">{i + 1}</span>
          <span className="col-title">{c.header}</span>
          {c.imported && <span className="tag tag-new">追加</span>}
        </li>
      ))}
      {items.length === 0 && <li className="muted">―</li>}
    </ul>
  );
}
