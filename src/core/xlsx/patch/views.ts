/**
 * <sheetViews> under structural ops (owner: patcher P).
 * Frozen panes follow R2b R10: only the insert and delete steps of an op change the number of frozen columns/rows
 * (a move = insert at the destination + delete of the source). Selections follow their cells (reset when deleted);
 * the frozen pane's topLeftCell is reset to the first unfrozen cell.
 */
import { adjustArea, mapIndex, type StructOp } from '../../formula/excelRef';
import { parseSqref, renderSqref } from './formula';
import { cellRef, parseCellRef } from './sheet';
import { findAll, getAttr, setAttr } from './xml';

/** New frozen count after an op on the same axis. */
export function frozenAfter(k: number, op: StructOp): number {
  if (k <= 0) return k;
  const ins = (at: number, n: number, kk: number) => (at <= kk ? kk + n : kk);
  const del = (at: number, n: number, kk: number) => {
    const s = at, e = at + n - 1;
    // the frozen block behaves like the range 1..k; a delete that removes it entirely leaves the split alone
    // (golden g-plan-ACE-F: moving column A out of a one-column freeze keeps one frozen column)
    if (s <= 1 && e >= kk) return kk;
    const hit = Math.max(0, Math.min(e, kk) - Math.max(s, 1) + 1);
    return kk - hit;
  };
  if (op.type === 'insert') return ins(op.at, op.count, k);
  if (op.type === 'delete') return del(op.at, op.count, k);
  const s = op.from, e = op.from + op.count - 1;
  if (op.to >= s && op.to <= e + 1) return k;
  const k1 = ins(op.to, op.count, k);
  const p = op.to > e ? op.from : op.from + op.count;
  return del(p, op.count, k1);
}

function mapCell(ref: string, op: StructOp): string | null {
  const p = parseCellRef(ref);
  if (!p) return ref;
  const v = mapIndex(op.axis === 'col' ? p.c : p.r, op);
  if (v == null) return null;
  return op.axis === 'col' ? cellRef(v, p.r) : cellRef(p.c, v);
}

export function adjustSheetViews(head: string, op: StructOp): string {
  const views = findAll(head, 'sheetView');
  if (!views.length) return head;
  let out = '', pos = 0;
  for (const v of views) {
    out += head.slice(pos, v.start) + adjustView(v.text, op);
    pos = v.end;
  }
  return out + head.slice(pos);
}

function adjustView(view: string, op: StructOp): string {
  const openM = /^<sheetView\b[^>]*>/.exec(view)!;
  let open = openM[0];
  const tl = getAttr(open, 'topLeftCell');
  if (tl) {
    const n = mapCell(tl, op);
    open = setAttr(open, 'topLeftCell', n);
  }
  if (open.endsWith('/>')) return open;
  let inner = view.slice(openM[0].length, view.length - '</sheetView>'.length);
  const pane = /<pane\b[^>]*\/>/.exec(inner);
  const frozen = pane && /state="frozen(?:Split)?"/.test(pane[0]);
  if (pane && frozen) {
    const xs = Number(getAttr(pane[0], 'xSplit') ?? 0) || 0;
    const ys = Number(getAttr(pane[0], 'ySplit') ?? 0) || 0;
    const nx = op.axis === 'col' ? frozenAfter(xs, op) : xs;
    const ny = op.axis === 'row' ? frozenAfter(ys, op) : ys;
    if (nx !== xs || ny !== ys) return open + rebuildFrozen(inner, pane[0], nx, ny) + '</sheetView>';
    // unchanged split: keep the pane, reset its topLeftCell to the first unfrozen cell if it moved away
    const want = cellRef(nx + 1, ny + 1);
    const cur = getAttr(pane[0], 'topLeftCell');
    if (cur && mapCell(cur, op) !== cur) inner = inner.replace(pane[0], setAttr(pane[0], 'topLeftCell', want));
  }
  inner = inner.replace(/<selection\b[^>]*\/>/g, (sel) => adjustSelection(sel, op));
  return open + inner + '</sheetView>';
}

function adjustSelection(sel: string, op: StructOp): string {
  const sq = getAttr(sel, 'sqref');
  const ac = getAttr(sel, 'activeCell');
  let areas = sq ? parseSqref(sq) : [];
  areas = areas.map((a) => adjustArea(a, op)).filter((a): a is NonNullable<typeof a> => !!a);
  let active = ac ? mapCell(ac, op) : null;
  if (!areas.length) {
    const first = active ?? 'A1';
    return setAttr(setAttr(sel, 'sqref', first), 'activeCell', ac ? first : null);
  }
  if (ac && !active) active = renderSqref([areas[0]]).split(':')[0];
  let s = setAttr(sel, 'sqref', renderSqref(areas));
  if (ac) s = setAttr(s, 'activeCell', active);
  return s;
}

/** Rebuild pane + selections for a new frozen split (nx columns, ny rows). */
function rebuildFrozen(inner: string, pane: string, nx: number, ny: number): string {
  const rest = inner.replace(pane, '').replace(/<selection\b[^>]*\/>/g, '');
  if (!nx && !ny) return '<selection activeCell="A1" sqref="A1"/>' + rest;
  const tl = cellRef(nx + 1, ny + 1);
  const active = nx && ny ? 'bottomRight' : nx ? 'topRight' : 'bottomLeft';
  let p = pane;
  p = setAttr(p, 'xSplit', nx ? String(nx) : null);
  p = setAttr(p, 'ySplit', ny ? String(ny) : null);
  p = setAttr(p, 'topLeftCell', tl);
  p = setAttr(p, 'activePane', active);
  let sels = '';
  if (nx && ny) {
    sels = `<selection pane="topRight" activeCell="${cellRef(nx + 1, 1)}" sqref="${cellRef(nx + 1, 1)}"/>` +
      `<selection pane="bottomLeft" activeCell="${cellRef(1, ny + 1)}" sqref="${cellRef(1, ny + 1)}"/>` +
      `<selection pane="bottomRight" activeCell="${tl}" sqref="${tl}"/>`;
  } else {
    sels = `<selection pane="${active}" activeCell="${tl}" sqref="${tl}"/>`;
  }
  return p + sels + rest;
}
