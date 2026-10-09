/**
 * Title merges that block a column move (merged areas lying entirely ABOVE the header row, e.g. a report title
 * merged over A1:K1). Excel refuses to move a part of a merged area; the editor offers to unmerge such titles
 * WITHOUT changing their look: LoggedOp {kind:'unmerge', refs, centerAcross:true} = Range.UnMerge +
 * HorizontalAlignment = xlCenterAcrossSelection (「選択範囲内で中央」). Every cell of each former area gets its
 * style with hAlign 'centerContinuous'; missing cells are created (styled, empty) so the alignment covers the whole
 * former area. Since round 3 every other cell of the area also receives the top-left cell's fill and border sides
 * (types.ts LoggedOp 'unmerge'), so a band styled on the top-left cell only stays a band — exactly what the JS patcher
 * (Engine.unmerge) and the Excel runner do.
 */
import type { StructOp } from '../formula/excelRef';
import type { Cell, DocColumn, MergedDoc, RangeBox } from '../types';
import { InsertStyles } from './insertStyle';
import { blockingMerges, refusalFor } from './struct';
import { boxText, carryCompiled, newCellXf } from './util';

/** merged areas entirely above the header row */
export function titleMerges(doc: MergedDoc): RangeBox[] {
  return doc.features.merges.filter((m) => m.r2 < doc.headerRow);
}

export type TitleBlock =
  /** nothing refuses the ops */
  | { kind: 'none' }
  /** only title merges refuse the ops: unmerging `boxes` lets them through */
  | { kind: 'titles'; boxes: RangeBox[]; reason: string }
  /** something else refuses them (even without the title merges) */
  | { kind: 'other'; reason: string };

/** Why `ops` are refused, and whether unmerging title merges alone would allow them. */
export function titleBlock(doc: MergedDoc, ops: readonly StructOp[]): TitleBlock {
  const why = refusalFor(doc, ops);
  if (!why) return { kind: 'none' };
  const titles = titleMerges(doc);
  if (!titles.length) return { kind: 'other', reason: why };
  const blocking = blockingMerges(doc, ops, titles);
  if (!blocking.length) return { kind: 'other', reason: why };
  const without = new Set(blocking);
  const rest: MergedDoc = { ...doc, features: { ...doc.features, merges: doc.features.merges.filter((m) => !without.has(m)) } };
  const still = refusalFor(rest, ops);
  if (still) return { kind: 'other', reason: still };
  return { kind: 'titles', boxes: blocking, reason: why };
}

/** 「タイトル行の結合セル（A1:K1）が…」 (the question the UI asks before unmerging) */
export function unmergeQuestion(boxes: readonly RangeBox[]): string {
  const list = boxes.map(boxText).join('、');
  return `タイトル行の結合セル（${list}）が列の移動を妨げています。見た目はそのままで結合だけを解除（選択範囲内で中央揃え）して移動しますか？`;
}

/** Cells of a merge larger than this are only restyled where they exist (the patcher does the same). */
const MAX_FILL = 100000;

/**
 * The doc after {kind:'unmerge', refs: boxes, centerAcross}: the merges are removed and
 *  - ALWAYS (independent of centerAcross) every other cell of each former area receives the top-left cell's FILL and
 *    the border sides it has for the edges of the area the cell lies on (top: first row, bottom: last row, left: first
 *    column, right: last column — for a one-row band: top / bottom on every cell, left on the first, right on the last
 *    cell), so a title band whose style was stored on the top-left cell only (openpyxl-style) stays visible. Additive:
 *    only what the top-left cell has is copied, the cell's own style is otherwise kept (the patcher's rule,
 *    Engine.unmerge);
 *  - with centerAcross every cell of the area gets hAlign 'centerContinuous'.
 * Derived display styles; styled empty cells are created where a cell is missing (it takes the row format, else the
 * column format, as a typed cell would).
 */
export function unmergeInDoc(doc: MergedDoc, boxes: readonly RangeBox[], centerAcross = true): MergedDoc {
  const same = (a: RangeBox, b: RangeBox) => a.r1 === b.r1 && a.r2 === b.r2 && a.c1 === b.c1 && a.c2 === b.c2;
  const removed = doc.features.merges.filter((m) => boxes.some((b) => same(b, m)));
  if (!removed.length) return doc;
  const merges = doc.features.merges.filter((m) => !removed.includes(m));
  const columns = doc.columns.slice();
  const table = new InsertStyles(doc.styles);
  const copied = new Set<number>();
  for (const m of removed) {
    const fill = (m.r2 - m.r1 + 1) * (m.c2 - m.c1 + 1) <= MAX_FILL;
    // the look to spread: the top-left cell's own style (a missing cell takes the row / column format)
    const tlCol: DocColumn | undefined = columns[m.c1 - 1];
    const tlCell = tlCol?.cells[m.r1 - 1];
    const tl = doc.styles[tlCell ? tlCell.s ?? 0 : newCellXf(doc, m.r1) ?? tlCol?.props.s ?? 0] ?? {};
    const hasLook = tl.fill !== undefined || !!tl.border;
    if (!centerAcross && !hasLook) continue;
    for (let c = m.c1; c <= Math.min(m.c2, columns.length); c++) {
      let col: DocColumn = columns[c - 1];
      for (let r = m.r1; r <= Math.min(m.r2, doc.rowCount); r++) {
        const cell = col.cells[r - 1];
        if (!cell && !fill) continue;
        // an existing cell keeps its own xf; a new one takes the row format, else the column format
        const base = cell ? cell.s ?? 0 : newCellXf(doc, r) ?? col.props.s ?? 0;
        let s = base;
        if (hasLook && !(r === m.r1 && c === m.c1)) {
          // the sides a cell lies on: top = first row of the area, bottom = last row, left = first column, right = last
          // column (a one-row band: top / bottom on every cell, left on the first, right on the last — the patcher's rule)
          const sides: Array<'top' | 'bottom' | 'left' | 'right'> = [];
          if (r === m.r1) sides.push('top');
          if (r === m.r2) sides.push('bottom');
          if (c === m.c1) sides.push('left');
          if (c === m.c2) sides.push('right');
          s = table.withLookOf(s, tl, sides);
        }
        if (centerAcross) s = table.withHAlign(s, 'centerContinuous');
        if (s === base) continue;
        if (!copied.has(c)) {
          const cells = col.cells.slice();
          carryCompiled(col.cells, cells);
          col = { ...col, cells };
          columns[c - 1] = col;
          copied.add(c);
        }
        const next: Cell = cell ? { ...cell, s } : { t: 'z', v: null, s };
        col.cells[r - 1] = next;
      }
    }
  }
  return {
    ...doc,
    columns: copied.size ? columns : doc.columns,
    styles: table.styles,
    features: { ...doc.features, merges },
    version: doc.version + 1,
  };
}
