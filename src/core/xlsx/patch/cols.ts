/**
 * <cols> model (owner: patcher P): one attribute string per column (without min/max), so inserts, deletes and
 * permutations are plain array operations; runs of identical columns are re-compressed on output.
 */
import { MAX_COL, movePerm, type StructOp } from '../../formula/excelRef';
import { findFirst, getAttr, setAttr } from './xml';

export class ColsModel {
  /** attrs[c] (1-based) = attribute string (leading spaces) or undefined (no <col> entry) */
  attrs: Array<string | undefined> = [];
  changed = false;
  /** the original <cols> element text ('' when absent) */
  private orig = '';

  constructor(head: string) {
    const m = findFirst(head, 'cols');
    if (!m) return;
    this.orig = m.text;
    for (const cm of m.inner.matchAll(/<col\b([^>]*?)\/?>(?:<\/col>)?/g)) {
      let a = cm[1];
      const min = Number(getAttr(a, 'min') ?? 0), max = Number(getAttr(a, 'max') ?? 0);
      if (!(min >= 1) || !(max >= min)) continue;
      a = setAttr(setAttr(a, 'min', null), 'max', null).replace(/\s+$/, '');
      for (let c = min; c <= Math.min(max, MAX_COL); c++) this.attrs[c] = a;
    }
  }

  get(c: number): string | undefined {
    return this.attrs[c];
  }

  set(c: number, a: string | undefined): void {
    this.attrs[c] = a;
    this.changed = true;
  }

  apply(op: StructOp, inheritLeft: boolean): void {
    if (op.axis !== 'col') return;
    const a = this.attrs;
    if (!a.length) {
      if (op.type !== 'insert' || !inheritLeft) return;
    }
    const len = Math.max(a.length, 1);
    if (op.type === 'insert') {
      const left = inheritLeft && op.at > 1 ? a[op.at - 1] : undefined;
      const ins: Array<string | undefined> = new Array(op.count).fill(left !== undefined ? inheritAttrs(left) : undefined);
      if (op.at >= len && left === undefined) return;
      const next = a.slice(0, op.at).concat(ins, a.slice(op.at));
      next.length = Math.min(next.length, MAX_COL + 1);
      this.attrs = next;
    } else if (op.type === 'delete') {
      if (op.at >= len) return;
      const next = a.slice(0, op.at).concat(a.slice(op.at + op.count));
      // columns that entered from beyond the old end keep the default (a span ending at XFD stays at XFD)
      const last = a[MAX_COL];
      if (last !== undefined && next.length <= MAX_COL) for (let c = MAX_COL - op.count + 1; c <= MAX_COL; c++) next[c] = last;
      this.attrs = next;
    } else {
      const p = movePerm(op.from, op.count, op.to);
      const hi = Math.max(op.from + op.count - 1, op.to);
      if (op.from >= len && op.to >= len) return;
      const next = a.slice();
      next.length = Math.max(a.length, hi + 1);
      for (let c = Math.min(op.from, op.to); c <= hi; c++) next[p(c)] = a[c];
      this.attrs = next;
    }
    this.changed = true;
  }

  /** <cols> element text ('' when no entries). */
  render(): string {
    const a = this.attrs;
    let out = '';
    let c = 1;
    while (c < a.length) {
      const v = a[c];
      if (v === undefined) { c++; continue; }
      let e = c;
      while (e + 1 < a.length && a[e + 1] === v) e++;
      out += `<col min="${c}" max="${e}"${v}/>`;
      c = e + 1;
    }
    return out ? `<cols>${out}</cols>` : '';
  }

  /** Replace/insert/remove <cols> in the sheet head. */
  applyToHead(head: string): string {
    if (!this.changed) return head;
    const txt = this.render();
    if (this.orig && head.includes(this.orig)) return head.replace(this.orig, txt);
    const m = findFirst(head, 'cols');
    if (m) return head.slice(0, m.start) + txt + head.slice(m.end);
    if (!txt) return head;
    // <cols> goes right before <sheetData> (head ends there)
    return head + txt;
  }
}

/**
 * Attributes an inserted column takes from its left neighbour (Excel "format same as left"): width, style, outline level — but neither
 * `collapsed` (it belongs to the group button) nor `hidden`: Excel inserts a VISIBLE column next to a hidden one, also inside a
 * collapsed group (measured, Excel 365; the doc does the same: doc/struct.ts makeInsertedColumns).
 */
function inheritAttrs(a: string): string {
  return setAttr(setAttr(a, 'collapsed', null), 'hidden', null);
}
