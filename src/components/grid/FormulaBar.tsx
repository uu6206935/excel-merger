/**
 * Formula bar (owner: grid G): address box ("C5") + fx + an input bound to the active cell.
 *
 *   <FormulaBar
 *     pos={active}                         // active cell (Excel row + column id) or null (input disabled)
 *     address="C5"                         // A1 address of the active cell ('' when none)
 *     value={editTextOfActiveCell}         // formula in UI form ('=SUM(A2:C2)') or the constant as Excel shows it
 *     mirror={editing?.text ?? null}       // Grid onEditingChange: text being typed IN the cell (shown instead of value)
 *     onCommitEdit={(pos, text) => …}      // same callback as the grid's (raw text, caller parses)
 *     onDone={() => gridRef.current?.focus()}   // after Enter / Esc: give the keyboard back to the grid
 *   />
 *
 * Behaviour (Excel-like, IME safe):
 *  - Typing edits a local draft. Enter commits (onCommitEdit(pos, draft)) when the draft differs from `value`,
 *    then calls onDone. Esc restores `value` and calls onDone. Leaving the input (blur) with a changed draft
 *    commits it to the cell that was active when typing started.
 *  - Enter / Esc are ignored while an IME composition is running (isComposing / keyCode 229), so the Enter
 *    that confirms a Japanese conversion never commits the cell.
 *  - Alt+Enter inserts nothing (single line); formulas are one line in this tool.
 *  - data-testid: formula-bar (root), formula-bar-address, formula-bar-input.
 */
import { useEffect, useRef, useState } from 'react';
import type { CellPos } from './types';
import './grid.css';

export interface FormulaBarProps {
  pos: CellPos | null;
  address: string;
  value: string;
  onCommitEdit: (pos: CellPos, text: string) => void;
  mirror?: string | null;
  onDone?: () => void;
  readOnly?: boolean;
  /** placeholder when no cell is active */
  placeholder?: string;
  className?: string;
}

function samePos(a: CellPos | null, b: CellPos | null): boolean {
  return a === b || (!!a && !!b && a.row === b.row && a.col === b.col);
}

export function FormulaBar(props: FormulaBarProps) {
  const { pos, address, value, mirror, readOnly } = props;
  const [draft, setDraft] = useState(value);
  const [dirty, setDirty] = useState(false);
  const [focused, setFocused] = useState(false);
  const editPosRef = useRef<CellPos | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const skipBlurCommitRef = useRef(false);

  // follow the active cell / its value while the user is not editing here
  const lastPosRef = useRef(pos);
  useEffect(() => {
    if (!samePos(lastPosRef.current, pos)) {
      lastPosRef.current = pos;
      setDirty(false);
      editPosRef.current = null;
      setDraft(value);
      return;
    }
    if (!dirty) setDraft(value);
  }, [pos, value, dirty]);

  const commit = () => {
    const at = editPosRef.current ?? pos;
    if (at && dirty && draft !== value) props.onCommitEdit(at, draft);
    setDirty(false);
    editPosRef.current = null;
  };

  const shown = !focused && !dirty && mirror !== null && mirror !== undefined ? mirror : draft;
  const disabled = !pos;

  return (
    <div className={`fb-bar${props.className ? ' ' + props.className : ''}`} data-testid="formula-bar">
      <div className="fb-addr" data-testid="formula-bar-address" aria-label="セル番地" title="選択中のセル">
        {address}
      </div>
      <div className="fb-fx" aria-hidden>
        fx
      </div>
      <input
        ref={inputRef}
        className="fb-input"
        data-testid="formula-bar-input"
        aria-label={address ? `数式バー（${address}）` : '数式バー'}
        spellCheck={false}
        autoComplete="off"
        disabled={disabled}
        readOnly={readOnly}
        placeholder={disabled ? (props.placeholder ?? '') : undefined}
        value={shown}
        onFocus={() => {
          setFocused(true);
          if (!dirty) setDraft(value);
        }}
        onBlur={() => {
          setFocused(false);
          if (skipBlurCommitRef.current) {
            skipBlurCommitRef.current = false;
            return;
          }
          commit();
        }}
        onChange={(e) => {
          if (!dirty) editPosRef.current = pos;
          setDraft(e.target.value);
          setDirty(true);
        }}
        onKeyDown={(e) => {
          // never act on keys that belong to an IME composition (the Enter that confirms a conversion)
          if (e.nativeEvent.isComposing || e.keyCode === 229) return;
          if (e.key === 'Enter') {
            e.preventDefault();
            commit();
            skipBlurCommitRef.current = true;
            props.onDone?.();
            if (document.activeElement !== inputRef.current) return;
            skipBlurCommitRef.current = false;
          } else if (e.key === 'Escape') {
            e.preventDefault();
            setDraft(value);
            setDirty(false);
            editPosRef.current = null;
            skipBlurCommitRef.current = true;
            props.onDone?.();
            if (document.activeElement !== inputRef.current) return;
            skipBlurCommitRef.current = false;
          }
          // Ctrl+Z etc. stay native (input undo) while typing here
          e.stopPropagation();
        }}
      />
    </div>
  );
}
