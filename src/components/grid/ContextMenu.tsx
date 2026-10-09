/**
 * Small accessible context menu (owner: grid G). Used by the grid for its column menu; generic enough for
 * other right-click menus.
 *
 *   <ContextMenu x={clientX} y={clientY} items={[{ label: '左に列を挿入', onSelect }, 'separator', …]} onClose={…} />
 *
 * - Rendered with position:fixed at the pointer, clamped into the viewport.
 * - Closes on: an item, Esc, Tab, a pointerdown outside, window blur / resize, any scroll outside the menu.
 * - Keyboard: ArrowUp/ArrowDown/Home/End move, Enter/Space activates; the first enabled item is focused on open.
 * - data-testid="context-menu" on the menu, `context-menu-item-<key>` on each item (key defaults to the index).
 */
import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import './grid.css';

export interface MenuItem {
  label: string;
  onSelect: () => void;
  disabled?: boolean;
  danger?: boolean;
  /** extra text on the right (shortcut / count); not part of the accessible name */
  hint?: string;
  /** stable key for data-testid="context-menu-item-<key>" */
  key?: string;
  icon?: ReactNode;
}

export type MenuEntry = MenuItem | 'separator';

export interface ContextMenuProps {
  x: number;
  y: number;
  items: MenuEntry[];
  onClose: () => void;
  /** aria-label of the menu */
  label?: string;
}

export function ContextMenu({ x, y, items, onClose, label }: ContextMenuProps) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({ left: x, top: y });
  const onCloseRef = useRef(onClose);
  onCloseRef.current = onClose;

  // clamp into the viewport once the size is known
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    const left = Math.max(4, Math.min(x, vw - r.width - 4));
    const top = y + r.height > vh - 4 ? Math.max(4, y - r.height) : y;
    setPos({ left, top });
    const first = el.querySelector<HTMLButtonElement>('button:not(:disabled)');
    first?.focus({ preventScroll: true });
  }, [x, y]);

  useEffect(() => {
    const close = () => onCloseRef.current();
    const onDown = (e: PointerEvent) => {
      if (ref.current && e.target instanceof Node && ref.current.contains(e.target)) return;
      close();
    };
    const onScroll = (e: Event) => {
      if (ref.current && e.target instanceof Node && ref.current.contains(e.target)) return;
      close();
    };
    window.addEventListener('pointerdown', onDown, true);
    window.addEventListener('blur', close);
    window.addEventListener('resize', close);
    window.addEventListener('scroll', onScroll, true);
    return () => {
      window.removeEventListener('pointerdown', onDown, true);
      window.removeEventListener('blur', close);
      window.removeEventListener('resize', close);
      window.removeEventListener('scroll', onScroll, true);
    };
  }, []);

  const onKeyDown = (e: React.KeyboardEvent<HTMLDivElement>) => {
    const buttons = Array.from(ref.current?.querySelectorAll<HTMLButtonElement>('button:not(:disabled)') ?? []);
    const i = buttons.indexOf(document.activeElement as HTMLButtonElement);
    const focusAt = (k: number) => buttons[(k + buttons.length) % buttons.length]?.focus({ preventScroll: true });
    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        focusAt(i + 1);
        break;
      case 'ArrowUp':
        e.preventDefault();
        focusAt(i < 0 ? buttons.length - 1 : i - 1);
        break;
      case 'Home':
        e.preventDefault();
        focusAt(0);
        break;
      case 'End':
        e.preventDefault();
        focusAt(buttons.length - 1);
        break;
      case 'Escape':
      case 'Tab':
        e.preventDefault();
        onClose();
        break;
      default:
        break;
    }
    e.stopPropagation();
  };

  return (
    <div
      ref={ref}
      className="g-menu"
      role="menu"
      aria-label={label}
      data-testid="context-menu"
      style={{ left: pos.left, top: pos.top }}
      onKeyDown={onKeyDown}
      onContextMenu={(e) => e.preventDefault()}
      onMouseDown={(e) => e.preventDefault()}
    >
      {items.map((it, i) =>
        it === 'separator' ? (
          <div key={'sep' + i} className="g-menu-sep" role="separator" />
        ) : (
          <button
            key={it.key ?? i}
            type="button"
            role="menuitem"
            className={`g-menu-item${it.danger ? ' is-danger' : ''}`}
            disabled={it.disabled}
            data-testid={`context-menu-item-${it.key ?? i}`}
            onClick={() => {
              onClose();
              it.onSelect();
            }}
          >
            {it.icon && <span className="g-menu-icon" aria-hidden>{it.icon}</span>}
            <span className="g-menu-label">{it.label}</span>
            {/* hint drawn by CSS (::after) so it is neither in textContent nor in the accessible name */}
            {it.hint && <span className="g-menu-hint" aria-hidden data-hint={it.hint} />}
          </button>
        ),
      )}
    </div>
  );
}
