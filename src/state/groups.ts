/**
 * Column GROUPS of the merged document = the columns that came from one file, and the file-card operations built on
 * them (pure, no React). Unit-tested in tests/ui/groups.test.ts.
 *
 * A group is found through the column ORIGIN (DocColumn.origin), never through colours or titles:
 *   origin.kind 'base'    -> the base file (every column of the base sheet, wherever the user moved it)
 *   origin.kind 'import'  -> origin.sourceId
 *   origin.kind 'new'     -> no file (columns added in the app)
 * The columns of a file need not be contiguous any more once the user has moved single columns: the helpers work
 * with index lists.
 */
import { computeMove, isNoopMove } from '../components/grid/colops';
import type { DocColumn } from '../core/types';

/** Id of the file a column belongs to (the base file for base columns), null for columns added in the app. */
export function columnFileId(col: Pick<DocColumn, 'origin'>, baseId: string | null): string | null {
  if (col.origin.kind === 'base') return baseId ?? 'base';
  if (col.origin.kind === 'import') return col.origin.sourceId;
  return null;
}

/** fileId -> ascending 0-based indexes of its columns in the current order (files without columns are absent). */
export function fileColumnIndexes(columns: ReadonlyArray<Pick<DocColumn, 'origin'>>, baseId: string | null): Map<string, number[]> {
  const out = new Map<string, number[]>();
  columns.forEach((c, i) => {
    const f = columnFileId(c, baseId);
    if (f === null) return;
    const list = out.get(f);
    if (list) list.push(i);
    else out.set(f, [i]);
  });
  return out;
}

/** Column ids of one file in current order ([] when the file has no column in the table). */
export function fileColumnIds(columns: ReadonlyArray<Pick<DocColumn, 'id' | 'origin'>>, baseId: string | null, fileId: string): string[] {
  return columns.filter((c) => columnFileId(c, baseId) === fileId).map((c) => c.id);
}

/**
 * The cards in the order of their columns in the table: files with columns sorted by their FIRST column, then the files
 * without any column (結合対象外, all columns deleted) in their original order. After the merge the card row is a
 * mini-map of the grid: moving a card moves its group and the card ends up where the group is.
 */
export function orderFilesByGroups<T extends { id: string }>(files: readonly T[], groups: ReadonlyMap<string, readonly number[]>): T[] {
  const withCols = files.filter((f) => (groups.get(f.id)?.length ?? 0) > 0);
  const without = files.filter((f) => (groups.get(f.id)?.length ?? 0) === 0);
  // Array.prototype.sort is stable: equal first columns cannot happen, but keep the original order anyway
  withCols.sort((a, b) => groups.get(a.id)![0] - groups.get(b.id)![0]);
  return [...withCols, ...without];
}

/**
 * The file whose complete column group is exactly the selected columns (the card that is "active"), else null.
 * An empty selection owns nothing.
 */
export function selectionOwner(columns: ReadonlyArray<Pick<DocColumn, 'id' | 'origin'>>, baseId: string | null, selected: ReadonlySet<string>): string | null {
  if (selected.size === 0) return null;
  let owner: string | null = null;
  let n = 0;
  for (const c of columns) {
    const f = columnFileId(c, baseId);
    if (selected.has(c.id)) {
      if (f === null || (owner !== null && f !== owner)) return null;
      owner = f;
      n++;
    }
  }
  if (owner === null || n !== selected.size) return null; // some selected id is not a column any more
  // every column of that file must be selected
  for (const c of columns) if (columnFileId(c, baseId) === owner && !selected.has(c.id)) return null;
  return owner;
}

// ---------------------------------------------------------------------------------------------
// reordering

/**
 * Files before the merge: move `id` so that it stands right before `beforeId` (null = after the last one).
 * Returns the new array, or null when nothing changes (unknown ids, dropped on itself / where it already is).
 */
export function moveBefore<T extends { id: string }>(items: readonly T[], id: string, beforeId: string | null): T[] | null {
  const from = items.findIndex((x) => x.id === id);
  if (from < 0) return null;
  const to = beforeId === null ? items.length : items.findIndex((x) => x.id === beforeId);
  if (to < 0 || to === from || to === from + 1) return null;
  const moved = items[from];
  const rest = items.filter((_, i) => i !== from);
  const at = to > from ? to - 1 : to;
  return [...rest.slice(0, at), moved, ...rest.slice(at)];
}

export type GroupMove =
  | { kind: 'noop' }
  /** the file has no column in the table: nothing to move */
  | { kind: 'nocols' }
  | { kind: 'move'; ids: string[]; gap: number };

/**
 * Card drop in the merged phase = move the file's whole column group in the grid.
 *  - `order`   current column ids (doc.columns order)
 *  - `groups`  fileId -> ascending column indexes (fileColumnIndexes)
 *  - `cards`   the file ids in the order the cards are shown (orderFilesByGroups)
 *  - dropped before the card `beforeId` (null = after the last card).
 * The group ends up directly BEFORE the target card's group (gap = its first column) or, dropped after the last card,
 * directly AFTER the furthest column of any other file's group. Returns the arguments of doc.moveColumns(ids, gap).
 */
export function groupMove(args: { order: readonly string[]; groups: ReadonlyMap<string, readonly number[]>; cards: readonly string[]; id: string; beforeId: string | null }): GroupMove {
  const { order, groups, cards, id, beforeId } = args;
  const mine = groups.get(id);
  if (!mine || mine.length === 0) return { kind: 'nocols' };
  const from = cards.indexOf(id);
  const to = beforeId === null ? cards.length : cards.indexOf(beforeId);
  if (from < 0 || to < 0 || to === from || to === from + 1) return { kind: 'noop' };
  let gap = -1;
  for (let k = to; k < cards.length; k++) {
    if (cards[k] === id) continue;
    const g = groups.get(cards[k]);
    if (g && g.length) {
      gap = g[0];
      break;
    }
  }
  if (gap < 0) {
    // after the last card that has columns: right behind the furthest column of the other files
    let last = -1;
    for (const [f, g] of groups) if (f !== id && g.length) last = Math.max(last, g[g.length - 1]);
    if (last < 0) return { kind: 'noop' };
    gap = last + 1;
  }
  const ids = mine.map((i) => order[i]);
  if (isNoopMove(order, new Set(ids), gap)) return { kind: 'noop' };
  return { kind: 'move', ids, gap };
}

/** The column order after a group move (for tests / previews). */
export function orderAfterGroupMove(order: readonly string[], ids: readonly string[], gap: number): string[] {
  return computeMove(order, new Set(ids), gap);
}
