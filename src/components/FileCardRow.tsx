/**
 * The row of file cards (「upload-row」): drag a card to change the order, click a card, keyboard alternative.
 *
 *  - POINTER based drag, never HTML5 draggable: the window-level handler for files dropped from the OS (App.tsx) must never see
 *    a card drag, and a card drag must never look like a file drop (dataTransfer is not involved at all).
 *  - Grab a card anywhere except its controls (buttons, selects, inputs, links, labels, the preview's scrollbar). A movement of
 *    DRAG_START_PX separates a click from a drag. The ghost follows the pointer, a vertical bar shows where the card will go
 *    (before the card under the pointer's left half / after the last card — the dashed 「ここにドロップ」 card is neither
 *    draggable nor a target), the row scrolls horizontally near its edges, Esc cancels, releasing far outside the row cancels.
 *  - A drop calls onReorder(id, beforeId) — `beforeId` is the card the dragged one is put before (null = after the last
 *    card). What it means (reordering the files before the merge, moving the file's column group in the grid after it) is
 *    the caller's business.
 *  - A click (no drag, not on a control) calls onCardClick(id).
 *  - Keyboard: a focused card (or any control inside it) + Alt+←/→ moves it one place; the result is announced in an
 *    aria-live region and the focus stays on the card (React moves the DOM node, which blurs it).
 */
import { useCallback, useEffect, useLayoutEffect, useRef, useState, type KeyboardEvent, type PointerEvent as ReactPointerEvent, type ReactNode } from 'react';
import { createPortal } from 'react-dom';

/** movement (px) that turns a press into a drag */
export const CARD_DRAG_START_PX = 5;
const EDGE_PX = 56;
const MAX_SPEED = 26;
/** the pointer this far outside the row (px) = releasing cancels */
const OUTSIDE_PX = 140;
/** pressing / clicking these inside a card never starts a drag and is not a card click */
const CONTROL_SELECTOR = 'button, select, input, textarea, a, label, summary, [role="button"], [contenteditable="true"], [data-nodrag]';

export interface CardGhostInfo {
  name: string;
  color: string;
  base?: boolean;
}

export interface FileCardRowProps {
  className?: string;
  /** no drag / click / keyboard (a merge or an analysis is running, a dialog is open) */
  disabled?: boolean;
  /** which cards may be grabbed (default: all) */
  canDrag?: (id: string) => boolean;
  /** `id` was dropped before `beforeId` (null = after the last card). true = the order changed (for the screen reader message). */
  onReorder: (id: string, beforeId: string | null) => boolean;
  onCardClick?: (id: string) => void;
  /** name / colour of a card for the ghost and the announcement */
  describe: (id: string) => CardGhostInfo | undefined;
  /** the FileCards (elements with data-file-id) and the dashed add card */
  children: ReactNode;
}

interface Gesture {
  pointerId: number;
  id: string;
  startX: number;
  startY: number;
  x: number;
  y: number;
  started: boolean;
  /** drop target: the card to go before (null = the end), undefined = not computed yet */
  beforeId: string | null | undefined;
  outside: boolean;
  raf: number;
  /** last indicator position pushed to React (avoids a state update on every pointer move) */
  indicator: string;
  /** stops the gesture and removes everything it added (listeners, capture, body class, ghost, bar); safe to call twice */
  end: () => void;
}

interface Indicator {
  left: number;
  top: number;
  height: number;
}

export function FileCardRow(props: FileCardRowProps) {
  const { className, children } = props;
  const rowRef = useRef<HTMLDivElement>(null);
  const gestureRef = useRef<Gesture | null>(null);
  const suppressClickRef = useRef(false);
  const [ghost, setGhost] = useState<{ info: CardGhostInfo; cancel: boolean } | null>(null);
  const [indicator, setIndicator] = useState<Indicator | null>(null);
  const [live, setLive] = useState('');
  const ghostRef = useRef<HTMLDivElement>(null);
  const propsRef = useRef(props);
  propsRef.current = props;

  const cardEls = useCallback((): HTMLElement[] => Array.from(rowRef.current?.querySelectorAll<HTMLElement>(':scope > [data-file-id]') ?? []), []);

  const placeGhost = useCallback((x: number, y: number) => {
    const el = ghostRef.current;
    if (el) el.style.transform = `translate(${Math.round(x + 14)}px, ${Math.round(y + 12)}px)`;
  }, []);

  // the ghost mounts after the drag started: put it under the pointer before the first paint
  useLayoutEffect(() => {
    const g = gestureRef.current;
    if (ghost && g) placeGhost(g.x, g.y);
  }, [ghost, placeGhost]);

  /** Where would a drop at the pointer go? (index among the cards; the pointer's side of each card's middle decides) */
  const updateTarget = (g: Gesture) => {
    const row = rowRef.current;
    if (!row) return;
    const cards = cardEls();
    let idx = cards.length;
    for (let i = 0; i < cards.length; i++) {
      const r = cards[i].getBoundingClientRect();
      if (g.x < r.left + r.width / 2) {
        idx = i;
        break;
      }
    }
    g.beforeId = idx < cards.length ? (cards[idx].dataset.fileId ?? null) : null;
    const rr = row.getBoundingClientRect();
    const outside = g.y < rr.top - OUTSIDE_PX || g.y > rr.bottom + OUTSIDE_PX || g.x < rr.left - OUTSIDE_PX || g.x > rr.right + OUTSIDE_PX;
    if (outside !== g.outside) {
      g.outside = outside;
      setGhost((cur) => (cur ? { ...cur, cancel: outside } : cur));
    }
    if (cards.length === 0) return;
    // the bar sits in the gap in front of card `idx` (or behind the last card), in the row's scrolling content coordinates
    const ref = idx < cards.length ? cards[idx].getBoundingClientRect() : cards[cards.length - 1].getBoundingClientRect();
    const x = (idx < cards.length ? ref.left - 5 : ref.right + 5) - rr.left + row.scrollLeft;
    const first = cards[0].getBoundingClientRect();
    const next: Indicator = { left: Math.round(x - 2), top: Math.round(first.top - rr.top + row.scrollTop), height: Math.round(first.height) };
    const key = `${next.left}|${next.top}|${next.height}`;
    if (key !== g.indicator) {
      g.indicator = key;
      setIndicator(next);
    }
  };

  const onPointerDown = (e: ReactPointerEvent<HTMLDivElement>) => {
    const p = propsRef.current;
    if (p.disabled || gestureRef.current || e.button !== 0 || (e.pointerType !== 'mouse' && e.pointerType !== 'pen')) return;
    const row = rowRef.current;
    const target = e.target as HTMLElement;
    const card = target.closest<HTMLElement>('[data-file-id]');
    if (!row || !card || !row.contains(card) || card.parentElement !== row) return;
    if (target.closest(CONTROL_SELECTOR) || target.classList.contains('matching-scroll')) return; // controls / the preview's scrollbar
    const id = card.dataset.fileId!;
    if (p.canDrag && !p.canDrag(id)) return;

    const g: Gesture = {
      pointerId: e.pointerId,
      id,
      startX: e.clientX,
      startY: e.clientY,
      x: e.clientX,
      y: e.clientY,
      started: false,
      beforeId: undefined,
      outside: false,
      raf: 0,
      indicator: '',
      end: () => undefined,
    };
    gestureRef.current = g;

    const onMove = (ev: PointerEvent) => {
      if (ev.pointerId !== g.pointerId) return;
      g.x = ev.clientX;
      g.y = ev.clientY;
      if (!g.started) {
        if (Math.hypot(g.x - g.startX, g.y - g.startY) < CARD_DRAG_START_PX) return;
        start();
      }
      placeGhost(g.x, g.y);
      updateTarget(g);
    };
    const onUp = (ev: PointerEvent) => {
      if (ev.pointerId !== g.pointerId) return;
      g.x = ev.clientX;
      g.y = ev.clientY;
      if (!g.started) {
        g.end(); // a plain click: the click event follows and reaches onClick
        return;
      }
      updateTarget(g);
      const { beforeId, outside } = g;
      g.end();
      // the click that follows the release must not count as a card click
      suppressClickRef.current = true;
      setTimeout(() => (suppressClickRef.current = false), 0);
      if (!outside && beforeId !== undefined) {
        const changed = propsRef.current.onReorder(g.id, beforeId);
        if (changed) afterReorder(g.id);
      }
    };
    const onCancel = (ev: PointerEvent) => {
      if (ev.pointerId === g.pointerId) g.end();
    };
    const onKey = (ev: globalThis.KeyboardEvent) => {
      if (ev.key !== 'Escape' || !g.started) return;
      ev.preventDefault();
      ev.stopPropagation();
      g.end();
      // the button is still down: its release must not count as a click on the card
      suppressClickRef.current = true;
      window.addEventListener('pointerup', () => setTimeout(() => (suppressClickRef.current = false), 0), { once: true, capture: true });
    };
    const tick = () => {
      if (gestureRef.current !== g || !g.started) return;
      const r = row.getBoundingClientRect();
      let v = 0;
      if (!g.outside) {
        if (g.x > r.right - EDGE_PX) v = Math.min(MAX_SPEED, Math.ceil(((g.x - (r.right - EDGE_PX)) / EDGE_PX) * MAX_SPEED));
        else if (g.x < r.left + EDGE_PX) v = -Math.min(MAX_SPEED, Math.ceil(((r.left + EDGE_PX - g.x) / EDGE_PX) * MAX_SPEED));
      }
      if (v !== 0) row.scrollLeft += v;
      updateTarget(g); // the cards move under a still pointer while the row scrolls
      g.raf = requestAnimationFrame(tick);
    };
    const start = () => {
      g.started = true;
      try {
        row.setPointerCapture(g.pointerId);
      } catch {
        /* synthetic events */
      }
      document.body.classList.add('card-dragging');
      window.getSelection()?.removeAllRanges();
      card.setAttribute('data-card-dragging', '');
      const info = propsRef.current.describe(g.id) ?? { name: '', color: '#6b7280' };
      setGhost({ info, cancel: false });
      g.raf = requestAnimationFrame(tick);
    };
    g.end = () => {
      if (gestureRef.current !== g) return;
      gestureRef.current = null;
      cancelAnimationFrame(g.raf);
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      window.removeEventListener('pointercancel', onCancel);
      window.removeEventListener('lostpointercapture', onCancel);
      window.removeEventListener('keydown', onKey, true);
      try {
        if (row.hasPointerCapture(g.pointerId)) row.releasePointerCapture(g.pointerId);
      } catch {
        /* ignore */
      }
      document.body.classList.remove('card-dragging');
      card.removeAttribute('data-card-dragging');
      if (g.started) {
        setGhost(null);
        setIndicator(null);
      }
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
    window.addEventListener('pointercancel', onCancel);
    window.addEventListener('lostpointercapture', onCancel);
    window.addEventListener('keydown', onKey, true);
  };

  // a gesture still running when the row unmounts (the files were cleared meanwhile): stop it cleanly
  useEffect(
    () => () => {
      gestureRef.current?.end();
    },
    [],
  );

  /** after a reorder React has moved the DOM nodes (which blurs a focused card): tell the screen reader, put the focus back */
  const afterReorder = (id: string, refocus = false) => {
    requestAnimationFrame(() => {
      const row = rowRef.current;
      if (!row) return;
      const cards = cardEls();
      const pos = cards.findIndex((c) => c.dataset.fileId === id);
      if (pos < 0) return;
      if (refocus) cards[pos].focus({ preventScroll: false });
      const name = propsRef.current.describe(id)?.name ?? '';
      setLive(`「${name}」を ${pos + 1} 番目に移動しました（全 ${cards.length} 件）`);
    });
  };

  const onClick = (e: React.MouseEvent<HTMLDivElement>) => {
    if (suppressClickRef.current) {
      suppressClickRef.current = false;
      return;
    }
    const p = propsRef.current;
    if (p.disabled || !p.onCardClick) return;
    const t = e.target as HTMLElement;
    const card = t.closest<HTMLElement>('[data-file-id]');
    if (!card || card.parentElement !== rowRef.current) return;
    if (t.closest(CONTROL_SELECTOR)) return;
    p.onCardClick(card.dataset.fileId!);
  };

  const onKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
    const p = propsRef.current;
    if (p.disabled || !e.altKey || e.ctrlKey || e.metaKey || e.shiftKey) return;
    if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
    const t = e.target as HTMLElement;
    if (t.closest('input, textarea, [contenteditable="true"]')) return; // Alt+arrows belong to text editing there
    const card = t.closest<HTMLElement>('[data-file-id]');
    if (!card || card.parentElement !== rowRef.current) return;
    e.preventDefault(); // (Alt+← is also the browser's 「戻る」)
    const cards = cardEls();
    const i = cards.indexOf(card);
    const id = card.dataset.fileId!;
    if (p.canDrag && !p.canDrag(id)) {
      setLive('このカードは移動できません');
      return;
    }
    const left = e.key === 'ArrowLeft';
    if (left ? i <= 0 : i >= cards.length - 1) {
      setLive(left ? '先頭のカードです' : '末尾のカードです');
      return;
    }
    const beforeId = left ? (cards[i - 1].dataset.fileId ?? null) : (cards[i + 2]?.dataset.fileId ?? null);
    if (p.onReorder(id, beforeId)) afterReorder(id, true);
  };

  const dragging = ghost !== null;
  return (
    <div
      ref={rowRef}
      className={`${className ?? ''}${dragging ? ' is-card-dragging' : ''}`.trim()}
      data-testid="file-cards"
      onPointerDown={onPointerDown}
      onClick={onClick}
      onKeyDown={onKeyDown}
      onDragStart={(e) => e.preventDefault()}
    >
      {children}
      {indicator && ghost && !ghost.cancel && (
        <div className="card-drop" data-testid="card-drop-indicator" style={{ left: indicator.left, top: indicator.top, height: indicator.height }} aria-hidden>
          <span className="card-drop-cap" />
        </div>
      )}
      <div className="visually-hidden" role="status" aria-live="polite" data-testid="card-live">
        {live}
      </div>
      {ghost &&
        createPortal(
          <div ref={ghostRef} className={`card-ghost${ghost.cancel ? ' is-cancel' : ''}`} data-testid="card-ghost" aria-hidden>
            <span className="card-ghost-dot" style={{ background: ghost.info.color }} />
            {ghost.cancel ? 'ここで離すと取り消します' : `${ghost.info.name}${ghost.info.base ? '（基準）' : ''} を移動`}
          </div>,
          document.body,
        )}
    </div>
  );
}
