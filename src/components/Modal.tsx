/**
 * Dialog shell (like the reference CalendarModal): backdrop, title, × button, Esc / backdrop click = cancel,
 * focus trapped inside the dialog (Tab / Shift+Tab cycle, focus that escapes is pulled back), initial focus on
 * the dialog's primary button (`initialFocus`), focus restored on close.
 */
import { type ReactNode, useEffect, useId, useRef } from 'react';

const FOCUSABLE = 'button:not([disabled]), [href], input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])';

/** open dialogs, innermost last (only the top one traps focus / reacts to Esc) */
const stack: symbol[] = [];

export interface ModalProps {
  title: ReactNode;
  testId: string;
  onClose: () => void;
  children: ReactNode;
  footer?: ReactNode;
  /** css width class: 'modal-sm' | 'modal-md' | 'modal-lg' */
  size?: 'sm' | 'md' | 'lg';
  /** element to focus first (selector inside the dialog) — the primary button */
  initialFocus?: string;
  /** when this value changes, the focus moves to `initialFocus` again (e.g. export: form -> progress -> result) */
  focusKey?: string | number;
  /**
   * selector (in the page) of the element that gets the focus on close when the element that had it on open is gone
   * (e.g. the 「キーを選ぶ…」 button disappears once the dialog selected a key)
   */
  restoreFocus?: string;
  /** Esc / backdrop / × do nothing (e.g. while exporting) */
  locked?: boolean;
  subtitle?: ReactNode;
  className?: string;
}

function focusables(root: HTMLElement): HTMLElement[] {
  return [...root.querySelectorAll<HTMLElement>(FOCUSABLE)].filter((el) => el.offsetParent !== null || el === document.activeElement);
}

export function Modal(props: ModalProps) {
  const { title, testId, onClose, children, footer, size = 'md', initialFocus, focusKey, locked, subtitle } = props;
  const ref = useRef<HTMLDivElement>(null);
  const titleId = useId();
  const closeRef = useRef(onClose);
  closeRef.current = onClose;
  const lockedRef = useRef(locked);
  lockedRef.current = locked;
  const initialRef = useRef(initialFocus);
  initialRef.current = initialFocus;
  const restoreRef = useRef(props.restoreFocus);
  restoreRef.current = props.restoreFocus;

  const focusInitial = () => {
    const root = ref.current;
    if (!root) return;
    const sel = initialRef.current;
    const first = (sel && root.querySelector<HTMLElement>(sel)) || focusables(root)[0];
    const target = first && !(first as HTMLButtonElement).disabled ? first : root;
    target.focus({ preventScroll: true });
  };

  // restore the focus of the page when the dialog closes
  useEffect(() => {
    const previous = document.activeElement as HTMLElement | null;
    return () => {
      if (previous && previous.isConnected) {
        previous.focus?.({ preventScroll: true });
        return;
      }
      const sel = restoreRef.current;
      if (!sel) return;
      // after the unmount: the page behind the dialog has re-rendered by then
      requestAnimationFrame(() => {
        const el = document.querySelector<HTMLElement>(sel);
        if (el && (document.activeElement === document.body || !document.activeElement)) el.focus({ preventScroll: true });
      });
    };
  }, []);

  // initial focus (again whenever focusKey changes)
  useEffect(() => {
    focusInitial();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [focusKey]);

  // Esc closes, Tab cycles inside, focus that escapes (e.g. after the focused button disappeared) comes back
  useEffect(() => {
    const me = Symbol('modal');
    stack.push(me);
    const isTop = () => stack[stack.length - 1] === me;
    const onKey = (e: KeyboardEvent) => {
      if (!isTop()) return;
      const root = ref.current;
      if (!root) return;
      if (e.key === 'Escape') {
        if (e.isComposing) return;
        e.preventDefault();
        e.stopPropagation();
        if (!lockedRef.current) closeRef.current();
        return;
      }
      if (e.key !== 'Tab') return;
      const items = focusables(root);
      if (!items.length) {
        e.preventDefault();
        root.focus({ preventScroll: true });
        return;
      }
      const firstEl = items[0];
      const lastEl = items[items.length - 1];
      const active = document.activeElement;
      const inside = !!active && root.contains(active);
      if (e.shiftKey && (!inside || active === firstEl || active === root)) {
        e.preventDefault();
        lastEl.focus();
      } else if (!e.shiftKey && (!inside || active === lastEl)) {
        e.preventDefault();
        firstEl.focus();
      }
    };
    const onFocusIn = (e: FocusEvent) => {
      if (!isTop()) return;
      const root = ref.current;
      const t = e.target as Node | null;
      if (!root || !t || root.contains(t)) return;
      focusInitial();
    };
    document.addEventListener('keydown', onKey, true);
    document.addEventListener('focusin', onFocusIn, true);
    return () => {
      document.removeEventListener('keydown', onKey, true);
      document.removeEventListener('focusin', onFocusIn, true);
      const i = stack.indexOf(me);
      if (i >= 0) stack.splice(i, 1);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  return (
    <div
      className="modal-backdrop"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget && !locked) onClose();
      }}
    >
      <div
        ref={ref}
        className={`modal modal-${size}${props.className ? ' ' + props.className : ''}`}
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        data-testid={testId}
        tabIndex={-1}
      >
        <button type="button" className="modal-close" onClick={onClose} aria-label="閉じる" disabled={locked}>
          ×
        </button>
        <div className="modal-head">
          <h2 id={titleId} className="modal-title">
            {title}
          </h2>
          {subtitle && <p className="modal-sub">{subtitle}</p>}
        </div>
        <div className="modal-body">{children}</div>
        {footer && <div className="modal-footer">{footer}</div>}
      </div>
    </div>
  );
}
