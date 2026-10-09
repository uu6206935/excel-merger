/**
 * Display style of a cell created by an Excel column insert ("format same as left").
 *
 * Verified with Excel 365 by the patcher's probe (tests/golden/com/probe-insert-format.ps1, mirrored by
 * xlsx/patch/styles.ts insertedStyle): the new cell takes the xf of the LEFT neighbour, but a border side survives
 * only when the RIGHT neighbour (the cell the insert pushes right) has the identical side. A cell whose left
 * neighbour has no style stays unstyled.
 *
 * The doc only needs display styles (XStyle); a derived style is appended to a copy of doc.styles (or an identical
 * existing entry is reused). The package patcher creates the matching xf record in styles.xml on replay.
 * The same table derives the styles of typed dates / % / currency (withNumFmt) and of unmerged titles
 * (withHAlign 'centerContinuous', withLookOf: the top-left cell's fill and border sides).
 */
import type { XBorderSide, XStyle } from '../types';

type Border = NonNullable<XStyle['border']>;
const SIDES = ['left', 'right', 'top', 'bottom'] as const;

const sameSide = (a: XBorderSide | undefined, b: XBorderSide | undefined): boolean =>
  !!a && !!b && a.style === b.style && (a.color ?? '') === (b.color ?? '');

/** canonical JSON (sorted keys) */
export function styleKey(s: XStyle | undefined): string {
  const canon = (v: unknown): unknown => {
    if (!v || typeof v !== 'object' || Array.isArray(v)) return v;
    const o = v as Record<string, unknown>;
    const out: Record<string, unknown> = {};
    for (const k of Object.keys(o).sort()) if (o[k] !== undefined) out[k] = canon(o[k]);
    return out;
  };
  return JSON.stringify(canon(s ?? {}));
}

/**
 * Excel switches WrapText on by itself when a text constant that contains a line feed (LF, also as CR LF; a lone CR
 * does not) is written into a cell whose style has no wrap yet — measured through COM (`Value2 = "a" & vbLf & "b"`,
 * golden q-wrap) and mirrored by the package patcher (Engine.setCell). Formulas and numbers never do. A cleared cell
 * keeps its style.
 */
export const textWrapsOnWrite = (cell: { t: string; v: string | number | boolean | null; f?: string } | null): boolean =>
  !!cell && cell.t === 's' && cell.f === undefined && typeof cell.v === 'string' && cell.v.includes('\n');

export class InsertStyles {
  styles: XStyle[];
  private copied = false;
  private cache = new Map<string, number>();
  private index: Map<string, number> | null = null;

  constructor(styles: XStyle[]) {
    this.styles = styles;
  }

  /** xf for a cell inserted between xf `before` (left) and `after` (right; 0 = no cell / no style). 0 = none. */
  inserted(before: number, after: number): number {
    if (!before) return 0;
    const key = `${before}:${after}`;
    const hit = this.cache.get(key);
    if (hit !== undefined) return hit;
    const A = this.styles[before] ?? {};
    const B: XStyle = after ? this.styles[after] ?? {} : {};
    let res = before;
    const ab: Border = A.border ?? {};
    const bb: Border = B.border ?? {};
    const nb: Border = {};
    let dropped = false;
    for (const s of SIDES) {
      if (!ab[s]) continue;
      if (sameSide(ab[s], bb[s])) nb[s] = ab[s];
      else dropped = true;
    }
    if (dropped) {
      const ns: XStyle = { ...A };
      if (Object.keys(nb).length) ns.border = nb;
      else delete ns.border;
      res = this.find(ns);
    }
    this.cache.set(key, res);
    return res;
  }

  /**
   * xf `xf` with number format `code` ('General' = none): what the patcher's Styles.withNumFmt creates for a
   * CellWrite.numFmt (Excel sets the number format when a date / % / currency is typed into a General cell).
   */
  withNumFmt(xf: number, code: string): number {
    const c = code.trim();
    if (!c) return xf;
    const key = `nf:${xf}\u0000${c}`;
    const hit = this.cache.get(key);
    if (hit !== undefined) return hit;
    const A = this.styles[xf] ?? {};
    const general = /^general$/i.test(c);
    let res = xf;
    if (general ? A.numFmt !== undefined : A.numFmt !== c) {
      const ns: XStyle = { ...A };
      if (general) delete ns.numFmt;
      else ns.numFmt = c;
      res = this.find(ns);
    }
    this.cache.set(key, res);
    return res;
  }

  /**
   * xf `xf` with horizontal alignment `h` (other alignment kept; the indent survives only with left / right /
   * distributed, as in Excel): the patcher's Styles.withHorizontal, used by the 'unmerge' op (centerContinuous =
   * 「選択範囲内で中央」).
   */
  withHAlign(xf: number, h: NonNullable<XStyle['hAlign']>): number {
    const key = `ha:${xf}\u0000${h}`;
    const hit = this.cache.get(key);
    if (hit !== undefined) return hit;
    const A = this.styles[xf] ?? {};
    const keepIndent = h === 'left' || h === 'right' || h === 'distributed';
    const cur = A.hAlign ?? 'general';
    let res = xf;
    if (cur !== h || (!keepIndent && A.indent !== undefined)) {
      const ns: XStyle = { ...A };
      if (h === 'general') delete ns.hAlign;
      else ns.hAlign = h;
      if (!keepIndent) delete ns.indent;
      res = this.find(ns);
    }
    this.cache.set(key, res);
    return res;
  }

  /**
   * xf `xf` with wrap text on (everything else kept): the patcher's Styles.withWrapText. Excel turns WrapText on by itself
   * when a text CONSTANT containing a line feed is written into a cell (see textWrapsOnWrite); nothing to change = `xf`.
   */
  withWrap(xf: number): number {
    const A = this.styles[xf] ?? {};
    if (A.wrap) return xf;
    const key = `wt:${xf}`;
    const hit = this.cache.get(key);
    if (hit !== undefined) return hit;
    const res = this.find({ ...A, wrap: true });
    this.cache.set(key, res);
    return res;
  }

  /**
   * xf `xf` with the FILL of style `tl` and the border `sides` of `tl` copied onto it (what 'unmerge' gives every other
   * cell of a former merged area so that a title band styled on the top-left cell only stays visible). Additive, like
   * the patcher / the Excel runner: only what `tl` actually has is copied (a fill it lacks, or a side it lacks, leaves
   * the cell's own fill / side alone); everything else of the cell's style is kept.
   */
  withLookOf(xf: number, tl: XStyle, sides: ReadonlyArray<keyof Border>): number {
    const tlSides = sides.filter((s) => tl.border?.[s]);
    if (tl.fill === undefined && !tlSides.length) return xf;
    const key = `lk:${xf}\u0000${tl.fill ?? ''}\u0000${tlSides.map((s) => `${s}=${tl.border![s]!.style}/${tl.border![s]!.color ?? ''}`).join(',')}`;
    const hit = this.cache.get(key);
    if (hit !== undefined) return hit;
    const A = this.styles[xf] ?? {};
    const ns: XStyle = { ...A };
    if (tl.fill !== undefined) ns.fill = tl.fill;
    if (tlSides.length) {
      const nb: Border = { ...(A.border ?? {}) };
      for (const s of tlSides) nb[s] = tl.border![s];
      ns.border = nb;
    }
    const res = styleKey(ns) === styleKey(A) ? xf : this.find(ns);
    this.cache.set(key, res);
    return res;
  }

  private find(s: XStyle): number {
    if (!this.index) {
      this.index = new Map();
      this.styles.forEach((x, i) => {
        const k = styleKey(x);
        if (!this.index!.has(k)) this.index!.set(k, i);
      });
    }
    const k = styleKey(s);
    const i = this.index.get(k);
    if (i !== undefined) return i;
    if (!this.copied) {
      this.styles = this.styles.slice();
      this.copied = true;
    }
    this.styles.push(s);
    this.index.set(k, this.styles.length - 1);
    return this.styles.length - 1;
  }
}
