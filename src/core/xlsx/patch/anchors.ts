/**
 * Object anchors under structural ops (owner: patcher P): DrawingML anchors (xdr:twoCellAnchor / oneCellAnchor,
 * 0-based col/row + EMU offsets), form-control anchors inside the sheet XML (<anchor><from>...), and VML anchors
 * of legacy notes / controls ("LeftCol, dx, TopRow, dy, RightCol, dx, BottomRow, dy", 0-based).
 */
import { adjustArea, mapIndex, type Area, type StructOp } from '../../formula/excelRef';
import { nsPrefix, reEsc } from './xml';

interface Pt { col: number; colOff: string; row: number; rowOff: string }

/** any element prefix: drawing anchors are written as xdr:from, as from (default namespace: openpyxl) or with another prefix */
const reFromTo = /<((?:[\w.-]+:)?)(from|to)>([\s\S]*?)<\/\1\2>/g;

function parsePt(inner: string): Pt | null {
  const g = (t: string) => new RegExp(`<(?:[\\w.-]+:)?${t}>\\s*(-?\\d+)\\s*</(?:[\\w.-]+:)?${t}>`).exec(inner)?.[1];
  const col = g('col'), row = g('row');
  if (col == null || row == null) return null;
  return { col: Number(col), row: Number(row), colOff: g('colOff') ?? '0', rowOff: g('rowOff') ?? '0' };
}

function writePt(inner: string, p: Pt): string {
  const set = (x: string, t: string, v: string | number) => x.replace(new RegExp(`(<(?:[\\w.-]+:)?${t}>)\\s*-?\\d+\\s*(</(?:[\\w.-]+:)?${t}>)`), `$1${v}$2`);
  let x = set(inner, 'col', p.col);
  x = set(x, 'colOff', p.colOff);
  x = set(x, 'row', p.row);
  x = set(x, 'rowOff', p.rowOff);
  return x;
}

type Mode = 'twoCell' | 'oneCell' | 'absolute';

/** One step of Excel's move sequence: 'paste' = Range.Cut(Destination) of the (already shifted) source block. */
type Step = StructOp | { type: 'paste'; axis: StructOp['axis']; from: number; count: number; to: number };

/**
 * Objects follow Excel's own move sequence (insert blank at the target, cut/paste the source block onto it,
 * delete the emptied source): only objects lying entirely inside the moved block travel with it.
 */
function stepsOf(op: StructOp): Step[] {
  if (op.type !== 'move') return [op];
  const n = op.count;
  const p = op.to < op.from ? op.from + n : op.from;
  return [
    { type: 'insert', axis: op.axis, at: op.to, count: n },
    { type: 'paste', axis: op.axis, from: p, count: n, to: op.to },
    { type: 'delete', axis: op.axis, at: p, count: n },
  ];
}

/** Adjust a (from,to) pair on the op axis. Returns the new points. */
function adjustPair(from: Pt, to: Pt | null, op: Step, mode: Mode): [Pt, Pt | null] {
  if (mode === 'absolute') return [from, to];
  const isCol = op.axis === 'col';
  const get = (p: Pt) => (isCol ? p.col : p.row) + 1; // 1-based
  const put = (p: Pt, v1: number, zeroOff: boolean): Pt => {
    const q = { ...p };
    if (isCol) { q.col = v1 - 1; if (zeroOff) q.colOff = '0'; } else { q.row = v1 - 1; if (zeroOff) q.rowOff = '0'; }
    return q;
  };
  const lo = get(from);
  if (op.type === 'paste') {
    // only objects lying entirely inside the cut block travel with it
    const s = op.from, e = op.from + op.count - 1;
    const inside = lo >= s && lo <= e && (!to || get(to) <= e);
    if (!inside) return [from, to];
    const d = op.to - op.from;
    return [put(from, lo + d, false), to ? put(to, get(to) + d, false) : null];
  }
  // "move but don't size": an insert moves the whole object; a delete still clips it (verified: t-mov-HI-G)
  if (!to || (mode === 'oneCell' && op.type === 'insert')) {
    const nl = mapIndex(lo, op);
    const newLo = nl ?? (op.type === 'delete' ? op.at : lo);
    const nf = put(from, newLo, nl == null);
    if (!to) return [nf, null];
    const d = newLo - lo;
    return [nf, put(to, get(to) + d, false)];
  }
  const hi = get(to);
  const area: Area = isCol
    ? { kind: 'area', c1: lo, c2: Math.max(lo, hi), r1: 1, r2: 1, ac1: false, ac2: false, ar1: false, ar2: false }
    : { kind: 'area', r1: lo, r2: Math.max(lo, hi), c1: 1, c2: 1, ac1: false, ac2: false, ar1: false, ar2: false };
  const na = adjustArea(area, op);
  if (!na) {
    const at = op.type === 'delete' ? op.at : lo;
    return [put(from, at, true), put(to, at, true)];
  }
  const nlo = isCol ? na.c1 : na.r1, nhi = isCol ? na.c2 : na.r2;
  const loDel = mapIndex(lo, op) == null;
  const hiDel = mapIndex(hi, op) == null;
  // a deleted right/bottom edge ends at the boundary after the kept part (0-based index = 1-based nhi)
  return [put(from, nlo, loDel), hiDel ? put(to, nhi + 1, true) : put(to, nhi, false)];
}

function adjustAnchorBlock(block: string, op: StructOp, mode: Mode): string {
  const pts: Array<{ kind: string; pre: string; m: RegExpExecArray }> = [];
  reFromTo.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = reFromTo.exec(block))) pts.push({ kind: m[2], pre: m[1], m });
  const fromM = pts.find((p) => p.kind === 'from');
  if (!fromM) return block;
  const toM = pts.find((p) => p.kind === 'to' && p.m.index > fromM.m.index);
  const from = parsePt(fromM.m[3]);
  if (!from) return block;
  const to = toM ? parsePt(toM.m[3]) : null;
  let nf: Pt = from, nt: Pt | null = to;
  for (const st of stepsOf(op)) [nf, nt] = adjustPair(nf, nt, st, mode);
  let out = block;
  // replace `to` first (later in the string) so indices of `from` stay valid
  if (toM && nt) {
    const t = toM.m;
    out = out.slice(0, t.index) + `<${toM.pre}to>${writePt(t[3], nt)}</${toM.pre}to>` + out.slice(t.index + t[0].length);
  }
  const f = fromM.m;
  out = out.slice(0, f.index) + `<${fromM.pre}from>${writePt(f[3], nf)}</${fromM.pre}from>` + out.slice(f.index + f[0].length);
  return out;
}

const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const VML_NS = 'urn:schemas-microsoft-com:vml';
const XL_NS = 'urn:schemas-microsoft-com:office:excel';

/** Drawing part (xl/drawings/drawingN.xml); the anchors are prefixed xdr: or sit in the default namespace (openpyxl). */
export function adjustDrawingXml(xml: string, op: StructOp): string {
  const p = reEsc(nsPrefix(xml, XDR_NS, 'xdr:'));
  return xml.replace(new RegExp(`<${p}(twoCellAnchor|oneCellAnchor)\\b([^>]*)>([\\s\\S]*?)</${p}\\1>`, 'g'), (m0, kind: string, attrs: string) => {
    let mode: Mode = 'twoCell';
    if (kind === 'oneCellAnchor') mode = 'oneCell';
    else {
      const ea = /editAs="([^"]*)"/.exec(attrs)?.[1];
      if (ea === 'oneCell') mode = 'oneCell';
      else if (ea === 'absolute') mode = 'absolute';
    }
    return adjustAnchorBlock(m0, op, mode);
  });
}

/** Control / OLE object anchors inside a worksheet (<anchor moveWithCells sizeWithCells><from>..</from><to>..</to></anchor>). */
export function adjustSheetObjectAnchors(tail: string, op: StructOp): string {
  if (tail.indexOf('<anchor') < 0) return tail;
  return tail.replace(/<anchor\b([^>]*)>([\s\S]*?)<\/anchor>/g, (m0, attrs: string) => {
    const move = /moveWithCells="(1|true)"/.test(attrs);
    const size = /sizeWithCells="(1|true)"/.test(attrs);
    const mode: Mode = size ? 'twoCell' : move ? 'oneCell' : 'absolute';
    return adjustAnchorBlock(m0, op, mode);
  });
}

/**
 * VML (legacy notes and form controls). Notes (<x:Row>/<x:Column>) follow their cell; a note whose cell is
 * deleted is removed. Other shapes' anchors adjust like a two-cell anchor.
 */
export function adjustVmlXml(xml: string, op: StructOp): string {
  // the prefixes of the VML / Excel namespaces are whatever the writer declared (v: x: from Excel, ns0: ns1: ns2: from openpyxl)
  const V = reEsc(nsPrefix(xml, VML_NS, 'v:'));
  const X = nsPrefix(xml, XL_NS, 'x:');
  const x = reEsc(X);
  const anchorRe = new RegExp(`<${x}Anchor>([\\s\\S]*?)</${x}Anchor>`);
  const rowRe = new RegExp(`<${x}Row>\\s*(\\d+)\\s*</${x}Row>`);
  const colRe = new RegExp(`<${x}Column>\\s*(\\d+)\\s*</${x}Column>`);
  return xml.replace(new RegExp(`<${V}shape\\b[\\s\\S]*?</${V}shape>`, 'g'), (shape) => {
    const rowM = rowRe.exec(shape);
    const colM = colRe.exec(shape);
    const isCol = op.axis === 'col';
    if (rowM && colM) {
      const v = Number(isCol ? colM[1] : rowM[1]) + 1;
      const nv = mapIndex(v, op);
      if (nv == null) return '';
      const d = nv - v;
      if (!d) return shape;
      let s = isCol
        ? shape.replace(colRe, `<${X}Column>${nv - 1}</${X}Column>`)
        : shape.replace(rowRe, `<${X}Row>${nv - 1}</${X}Row>`);
      s = s.replace(anchorRe, (_m, body: string) => {
        const p = body.split(',').map((t) => t.trim());
        if (p.length !== 8) return _m;
        const idx = isCol ? [0, 4] : [2, 6];
        for (const i of idx) p[i] = String(Math.max(0, Number(p[i]) + d));
        return `<${X}Anchor>\n    ${p.join(', ')}</${X}Anchor>`;
      });
      return s;
    }
    return shape.replace(anchorRe, (_m, body: string) => {
      const p = body.split(',').map((t) => t.trim());
      if (p.length !== 8) return _m;
      const [li, ri] = isCol ? [0, 4] : [2, 6];
      const [loi, roi] = isCol ? [1, 5] : [3, 7];
      const from: Pt = { col: Number(p[li]), colOff: p[loi], row: 0, rowOff: '0' };
      const to: Pt = { col: Number(p[ri]), colOff: p[roi], row: 0, rowOff: '0' };
      // reuse the column logic: map the chosen axis onto the Pt.col field
      let nf: Pt = from, nt: Pt | null = to;
      for (const st of stepsOf({ ...op, axis: 'col' } as StructOp)) [nf, nt] = adjustPair(nf, nt, st, 'twoCell');
      p[li] = String(nf.col); p[loi] = nf.colOff; p[ri] = String(nt!.col); p[roi] = nt!.colOff;
      return `<${X}Anchor>\n    ${p.join(', ')}</${X}Anchor>`;
    });
  });
}
