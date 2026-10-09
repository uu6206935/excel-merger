import { describe, expect, it } from 'vitest';
import {
  canRedo, canUndo, commit, deleteColumns, initHistory, insertColumns, moveColumns, redo, setCellInput, setColumnWidth, undo,
} from '../../src/core/doc';
import type { EditorSnapshot } from '../../src/core/types';
import { ext, fAt, ids, ok, roster } from './helpers';

describe('history', () => {
  it('undo / redo walk through immutable snapshots', () => {
    const s0 = roster();
    let h = initHistory(s0);
    expect(canUndo(h)).toBe(false);
    expect(canRedo(h)).toBe(false);
    const steps: EditorSnapshot[] = [s0];
    const run = (fn: (s: EditorSnapshot) => EditorSnapshot) => {
      const next = fn(h.present);
      h = commit(h, next);
      steps.push(next);
    };
    run((s) => ok(moveColumns(s, ['b4'], 0)));
    run((s) => ok(insertColumns(s, 2, 1)));
    run((s) => ok(setCellInput(s, 'n1', 2, '=B2*10')));
    run((s) => ok(setColumnWidth(s, 'n1', 18)));
    run((s) => ok(deleteColumns(s, ['b3'], { convertDependents: true })));
    expect(h.present.log.map((l) => l.kind)).toEqual(['struct', 'struct', 'cells', 'colWidth', 'cells', 'struct']);
    // undo everything: exactly the earlier snapshot objects come back
    for (let i = steps.length - 2; i >= 0; i--) {
      h = undo(h);
      expect(h.present).toBe(steps[i]);
    }
    expect(canUndo(h)).toBe(false);
    expect(undo(h)).toBe(h);
    expect(h.present.log).toEqual([]);
    expect(fAt(h.present.doc, 'E2')).toBe('D2*2');
    expect(ext(h.present.doc, '集計', 'B3')).toBe('名簿!C2');
    // redo everything
    for (let i = 1; i < steps.length; i++) {
      h = redo(h);
      expect(h.present).toBe(steps[i]);
    }
    expect(canRedo(h)).toBe(false);
    expect(redo(h)).toBe(h);
    // a new command after an undo drops the redo branch
    h = undo(undo(h));
    expect(canRedo(h)).toBe(true);
    h = commit(h, ok(moveColumns(h.present, ['b1'], 3)));
    expect(canRedo(h)).toBe(false);
    expect(h.past).toHaveLength(4);
  });

  it('undo after a refused op changes nothing (refusals never produce snapshots)', () => {
    const s0 = roster((sp) => {
      sp.sheets[0].features = { merges: [{ r1: 8, c1: 1, r2: 8, c2: 3 }], rules: [], tables: [], arrays: [], hyperlinks: [] };
    });
    const r = moveColumns(s0, ['b2'], 6);
    expect(r.ok).toBe(false);
    expect(ids(s0.doc)).toEqual(['b1', 'b2', 'b3', 'b4', 'b5', 'i:s2:3', 'i:s2:4', 'b6']);
    expect(s0.log).toEqual([]);
  });

  it('keeps a bounded number of steps', () => {
    let h = initHistory(roster());
    for (let i = 0; i < 230; i++) h = commit(h, ok(setColumnWidth(h.present, 'b2', 10 + (i % 7) + 0.5)));
    expect(h.past.length).toBe(200);
    expect(h.present.log).toHaveLength(230);
  });
});
