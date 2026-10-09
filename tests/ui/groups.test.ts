/**
 * File groups of the merged document (round 7): which columns belong to a file (by ORIGIN, not colour), the order of the cards
 * after the merge, the gap a card drop turns into, the "which card is active" selection test and the pure reorder before the merge.
 */
import { describe, expect, it } from 'vitest';
import type { DocColumn } from '../../src/core/types';
import {
  columnFileId,
  fileColumnIds,
  fileColumnIndexes,
  groupMove,
  moveBefore,
  orderAfterGroupMove,
  orderFilesByGroups,
  selectionOwner,
} from '../../src/state/groups';
import { editor } from './fixtures';

type Col = Pick<DocColumn, 'id' | 'origin'>;
const base = (n: number): Col => ({ id: `b${n}`, origin: { kind: 'base', col: n } });
const imp = (s: string, n: number): Col => ({ id: `i:${s}:${n}`, origin: { kind: 'import', sourceId: s, col: n } });
const fresh = (n: number): Col => ({ id: `n${n}`, origin: { kind: 'new' } });

// [b1 b2 b3 | s2:1 s2:2 | s3:1 s3:2 s3:3 | n1]
const cols: Col[] = [base(1), base(2), base(3), imp('s2', 1), imp('s2', 2), imp('s3', 1), imp('s3', 2), imp('s3', 3), fresh(1)];
const order = cols.map((c) => c.id);

describe('column groups by origin', () => {
  it('base columns belong to the base file, imports to their source, new columns to nobody', () => {
    expect(columnFileId(cols[0], 's1')).toBe('s1');
    expect(columnFileId(cols[3], 's1')).toBe('s2');
    expect(columnFileId(cols[8], 's1')).toBeNull();
    const groups = fileColumnIndexes(cols, 's1');
    expect([...groups.entries()]).toEqual([
      ['s1', [0, 1, 2]],
      ['s2', [3, 4]],
      ['s3', [5, 6, 7]],
    ]);
    expect(fileColumnIds(cols, 's1', 's3')).toEqual(['i:s3:1', 'i:s3:2', 'i:s3:3']);
    expect(fileColumnIds(cols, 's1', 'zz')).toEqual([]);
  });

  it('a file whose columns were split by moves keeps ALL its columns in the group (ascending indexes)', () => {
    const moved = [cols[3], cols[0], cols[1], cols[4], cols[2]]; // s2, s1, s1, s2, s1
    const g = fileColumnIndexes(moved, 's1');
    expect(g.get('s2')).toEqual([0, 3]);
    expect(g.get('s1')).toEqual([1, 2, 4]);
  });

  it('works on the real merged document of the mock services', async () => {
    const { snapshot } = await editor();
    const d = snapshot.doc;
    const g = fileColumnIndexes(d.columns, 's1');
    expect(g.get('s1')!.length).toBeGreaterThan(5);
    expect(g.get('s2')!.length).toBeGreaterThan(0);
    expect(g.get('s3')!.length).toBeGreaterThan(0);
    // every column is in exactly one group
    expect([...g.values()].reduce((a, v) => a + v.length, 0)).toBe(d.columns.length);
    expect(fileColumnIds(d.columns, 's1', 's2').every((id) => id.startsWith('i:s2:'))).toBe(true);
  });
});

describe('orderFilesByGroups (cards after the merge)', () => {
  const files = [{ id: 's1' }, { id: 's2' }, { id: 's3' }, { id: 's4' }];
  it('sorts by the first column of each group; files without columns go last in their original order', () => {
    const groups = new Map<string, number[]>([
      ['s1', [4, 5]],
      ['s2', [0, 1]],
      ['s3', [2, 3]],
    ]);
    expect(orderFilesByGroups(files, groups).map((f) => f.id)).toEqual(['s2', 's3', 's1', 's4']);
  });
  it('is the original order when the groups are in file order', () => {
    const groups = fileColumnIndexes(cols, 's1');
    expect(orderFilesByGroups(files, groups).map((f) => f.id)).toEqual(['s1', 's2', 's3', 's4']);
  });
});

describe('selectionOwner (the active card)', () => {
  it('is the file whose COMPLETE group is selected, else null', () => {
    expect(selectionOwner(cols, 's1', new Set(['i:s2:1', 'i:s2:2']))).toBe('s2');
    expect(selectionOwner(cols, 's1', new Set(['b1', 'b2', 'b3']))).toBe('s1');
    expect(selectionOwner(cols, 's1', new Set(['i:s2:1']))).toBeNull(); // only part of the group
    expect(selectionOwner(cols, 's1', new Set(['i:s2:1', 'i:s2:2', 'b1']))).toBeNull(); // two files
    expect(selectionOwner(cols, 's1', new Set(['n1']))).toBeNull(); // an app column is nobody's
    expect(selectionOwner(cols, 's1', new Set())).toBeNull();
    expect(selectionOwner(cols, 's1', new Set(['i:s2:1', 'i:s2:2', 'gone']))).toBeNull();
  });
  it('a split group is complete only with all its columns', () => {
    const split: Col[] = [imp('s2', 1), base(1), imp('s2', 2)];
    expect(selectionOwner(split, 's1', new Set(['i:s2:1', 'i:s2:2']))).toBe('s2');
    expect(selectionOwner(split, 's1', new Set(['i:s2:1']))).toBeNull();
  });
});

describe('moveBefore (files before the merge)', () => {
  const files = ['a', 'b', 'c', 'd'].map((id) => ({ id }));
  const ids = (x: Array<{ id: string }> | null) => x?.map((f) => f.id);
  it('moves right before the target, or to the end for beforeId = null', () => {
    expect(ids(moveBefore(files, 'a', 'c'))).toEqual(['b', 'a', 'c', 'd']);
    expect(ids(moveBefore(files, 'd', 'b'))).toEqual(['a', 'd', 'b', 'c']);
    expect(ids(moveBefore(files, 'a', null))).toEqual(['b', 'c', 'd', 'a']);
    expect(ids(moveBefore(files, 'c', 'a'))).toEqual(['c', 'a', 'b', 'd']);
  });
  it('returns null when nothing would change (itself, where it already is, unknown ids)', () => {
    expect(moveBefore(files, 'b', 'b')).toBeNull();
    expect(moveBefore(files, 'b', 'c')).toBeNull(); // right before its successor = where it is
    expect(moveBefore(files, 'd', null)).toBeNull();
    expect(moveBefore(files, 'zz', 'a')).toBeNull();
    expect(moveBefore(files, 'a', 'zz')).toBeNull();
  });
  it('never mutates the input', () => {
    const before = files.map((f) => f.id);
    moveBefore(files, 'a', null);
    expect(files.map((f) => f.id)).toEqual(before);
  });
});

describe('groupMove (card drop in the merged phase -> moveColumns(ids, gap))', () => {
  const groups = fileColumnIndexes(cols, 's1');
  const cards = ['s1', 's2', 's3'];
  const mv = (id: string, beforeId: string | null, over = groups, c = cards) => groupMove({ order, groups: over, cards: c, id, beforeId });

  it('before the target card: the gap is the target group\'s first column', () => {
    const m = mv('s3', 's2');
    expect(m).toEqual({ kind: 'move', ids: ['i:s3:1', 'i:s3:2', 'i:s3:3'], gap: 3 });
    expect(orderAfterGroupMove(order, (m as { ids: string[] }).ids, (m as { gap: number }).gap)).toEqual(['b1', 'b2', 'b3', 'i:s3:1', 'i:s3:2', 'i:s3:3', 'i:s2:1', 'i:s2:2', 'n1']);
    const first = mv('s3', 's1');
    expect(first).toMatchObject({ kind: 'move', gap: 0 });
    expect(orderAfterGroupMove(order, (first as { ids: string[] }).ids, 0)[0]).toBe('i:s3:1');
  });

  it('after the last card: right behind the furthest column of the other files (app columns stay after)', () => {
    const m = mv('s1', null);
    expect(m).toMatchObject({ kind: 'move', ids: ['b1', 'b2', 'b3'], gap: 8 });
    expect(orderAfterGroupMove(order, ['b1', 'b2', 'b3'], 8)).toEqual(['i:s2:1', 'i:s2:2', 'i:s3:1', 'i:s3:2', 'i:s3:3', 'b1', 'b2', 'b3', 'n1']);
  });

  it('dropping on itself or where it already is = noop (no undo entry)', () => {
    expect(mv('s2', 's2')).toEqual({ kind: 'noop' });
    expect(mv('s2', 's3')).toEqual({ kind: 'noop' }); // right before its successor
    expect(mv('s3', null)).toEqual({ kind: 'noop' }); // the last one, dropped at the end
    expect(mv('s1', 's2')).toEqual({ kind: 'noop' });
  });

  it('a file without columns cannot be moved; dropping before a card without columns goes to the end of the groups', () => {
    const cards4 = ['s1', 's2', 's3', 's4'];
    expect(mv('s4', 's1', groups, cards4)).toEqual({ kind: 'nocols' });
    const m = groupMove({ order, groups, cards: cards4, id: 's1', beforeId: 's4' });
    expect(m).toMatchObject({ kind: 'move', gap: 8 });
  });

  it('a split target group: before ITS FIRST column; the moved group is gathered whole', () => {
    // order: s2a, s1a, s1b, s2b, s3  (s2 is split around s1)
    const split: Col[] = [imp('s2', 1), base(1), base(2), imp('s2', 2), imp('s3', 1)];
    const o = split.map((c) => c.id);
    const g = fileColumnIndexes(split, 's1');
    const c = ['s2', 's1', 's3']; // by first column
    expect(orderFilesByGroups([{ id: 's1' }, { id: 's2' }, { id: 's3' }], g).map((f) => f.id)).toEqual(c);
    // s3 before s1: ahead of s1's first column (index 1)
    const m1 = groupMove({ order: o, groups: g, cards: c, id: 's3', beforeId: 's1' });
    expect(m1).toMatchObject({ kind: 'move', gap: 1 });
    expect(orderAfterGroupMove(o, ['i:s3:1'], 1)).toEqual(['i:s2:1', 'i:s3:1', 'b1', 'b2', 'i:s2:2']);
    // s2 (split) after the last card: both its columns end up together at the end
    const m2 = groupMove({ order: o, groups: g, cards: c, id: 's2', beforeId: null }) as { ids: string[]; gap: number };
    expect(m2.ids).toEqual(['i:s2:1', 'i:s2:2']);
    expect(orderAfterGroupMove(o, m2.ids, m2.gap)).toEqual(['b1', 'b2', 'i:s3:1', 'i:s2:1', 'i:s2:2']);
  });

  it('only one file in the table: nothing to move around', () => {
    const only = fileColumnIndexes([base(1), base(2)], 's1');
    expect(groupMove({ order: ['b1', 'b2'], groups: only, cards: ['s1'], id: 's1', beforeId: null })).toEqual({ kind: 'noop' });
  });
});
