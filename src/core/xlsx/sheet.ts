/**
 * Worksheet part reader (owner: reader R).
 *
 *  - scanSheetData(): fast string scanner for <sheetData> (<row>/<c> with r, s, t, cm; children f, v, is). No DOM,
 *    no per-cell regex. Shared formulas are expanded with excelRef.translateFormula (never ExcelJS slideFormula).
 *  - parseSheetRest(): everything outside <sheetData> (small) through the XML tree helper: cols, sheetFormatPr,
 *    panes, merges, CF / DV incl. x14 extLst (FeatureRule.type = <cfRule type> / <dataValidation type>, the latter
 *    'none' when absent as the schema defaults), hyperlinks, autoFilter, table parts, sparkline groups.
 *
 * Sparse storage: rows / cells are only allocated where the file has them. Formatted cells WITHOUT a value that lie
 * far away from the data (beyond ROW_SLACK rows / COL_SLACK columns of the last value/formula) are not kept, so a
 * stray formatted cell at XFD1048576 does not produce a 1,048,576-row grid. maxRow / maxCol = extent of the cells
 * with values / formulas plus the kept formatted cells.
 */
import type { Cell, ColProps, FeatureRule, RangeBox, RowProps, SheetFeatures } from '../types';
import { MAX_COL, MAX_ROW, parseAreaText, translateFormula } from '../formula/excelRef';
import type { Rel } from './zip';
import { StringItemScanner, type SharedStrings } from './strings';
import type { Palette } from './theme';
import { boolAttr, decodeOoxmlEscapes, decodeXmlText, descendants, innerText, kid, kids, numAttr, parseAttrString, relId, type XEl } from './xml';

/** Formatted empty cells more than this many rows below the last value/formula are ignored. */
export const ROW_SLACK = 1000;
/** Formatted empty cells more than this many columns right of the last value/formula are ignored. */
export const COL_SLACK = 26;

export interface SheetScanContext {
  sst: SharedStrings;
  /** per xf index: date/time number format */
  isDate: boolean[];
  /** per xf index: what the date format shows (1904 conversion only shifts calendar dates); default 'date' */
  dateMode?: Array<'none' | 'date' | 'time' | 'elapsed'>;
  date1904: boolean;
  styleCount: number;
  pal: Palette;
}

export interface ScannedSheetData {
  rows: Array<Array<Cell | undefined> | undefined>;
  rowProps: Array<RowProps | undefined>;
  maxRow: number;
  maxCol: number;
  arrays: RangeBox[];
  /** array / spill ranges as written (A1) with their master cell */
  arrayRefs: Array<{ ref: string; row: number; col: number; dynamic: boolean }>;
  counts: { dataTable: number; sharedMissing: number; badStyle: number; droppedFormatted: number; badValue: number; noResult: number };
}

/** Normalised RangeBox of an A1 ref ("A1", "A1:B2", "$A:$A", "3:5"); null when unparsable. */
export function boxOf(ref: string): RangeBox | null {
  const a = parseAreaText(ref.trim().replace(/\$/g, ''));
  if (!a) return null;
  let r1 = a.r1;
  let r2 = a.r2;
  let c1 = a.c1;
  let c2 = a.c2;
  if (a.kind === 'cols') {
    r1 = 1;
    r2 = MAX_ROW;
  } else if (a.kind === 'rows') {
    c1 = 1;
    c2 = MAX_COL;
  }
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

/** Excel serial (1900 system) of an ISO 8601 date/time string (t="d" cells); undefined when unparsable. */
export function isoToSerial(iso: string): number | undefined {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(\.\d+)?)?)?(Z|[+-]\d{2}:?\d{2})?$/.exec(iso.trim());
  if (m) {
    const ms = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0), m[7] ? Math.round(parseFloat(m[7]) * 1000) : 0);
    let serial = ms / 86400000 + 25569;
    if (serial < 61) serial -= 1; // Excel's fictitious 1900-02-29
    return serial;
  }
  const t = /^T?(\d{2}):(\d{2})(?::(\d{2})(\.\d+)?)?$/.exec(iso.trim());
  if (t) return (+t[1] * 3600 + +t[2] * 60 + +(t[3] ?? 0) + (t[4] ? parseFloat(t[4]) : 0)) / 86400;
  return undefined;
}

// Reused attribute holder for <c> (avoids one object per cell)
interface CAttrs {
  r: string | undefined;
  s: number;
  t: string | undefined;
  cm: boolean;
}

function isWs(c: number): boolean {
  return c === 32 || c === 9 || c === 10 || c === 13;
}

/** Scan attributes of a <c> start tag in xml[from, to). */
function scanCAttrs(xml: string, from: number, to: number, out: CAttrs): void {
  out.r = undefined;
  out.s = 0;
  out.t = undefined;
  out.cm = false;
  let i = from;
  while (i < to) {
    while (i < to && isWs(xml.charCodeAt(i))) i++;
    if (i >= to) break;
    const ns = i;
    while (i < to && xml.charCodeAt(i) !== 61 /* = */ && !isWs(xml.charCodeAt(i))) i++;
    const ne = i;
    while (i < to && xml.charCodeAt(i) !== 34 && xml.charCodeAt(i) !== 39) i++;
    const q = xml.charCodeAt(i);
    const vs = i + 1;
    let ve = xml.indexOf(q === 39 ? "'" : '"', vs);
    if (ve < 0 || ve > to) ve = to;
    i = ve + 1;
    const len = ne - ns;
    if (len === 1) {
      const ch = xml.charCodeAt(ns);
      if (ch === 114 /* r */) out.r = xml.slice(vs, ve);
      else if (ch === 115 /* s */) out.s = parseInt(xml.slice(vs, ve), 10) || 0;
      else if (ch === 116 /* t */) out.t = xml.slice(vs, ve);
    } else if (len === 2 && xml.charCodeAt(ns) === 99 /* c */ && xml.charCodeAt(ns + 1) === 109 /* m */) {
      out.cm = true;
    }
  }
}

/** "AB12" → [row, col]; [0, 0] when invalid. */
function parseRef(ref: string): [number, number] {
  let col = 0;
  let row = 0;
  let i = 0;
  const n = ref.length;
  if (i < n && ref.charCodeAt(i) === 36) i++;
  for (; i < n; i++) {
    const c = ref.charCodeAt(i);
    if (c >= 65 && c <= 90) col = col * 26 + (c - 64);
    else if (c >= 97 && c <= 122) col = col * 26 + (c - 96);
    else break;
  }
  if (i < n && ref.charCodeAt(i) === 36) i++;
  for (; i < n; i++) {
    const c = ref.charCodeAt(i);
    if (c >= 48 && c <= 57) row = row * 10 + (c - 48);
    else return [0, 0];
  }
  return [row, col];
}

/**
 * Scan <sheetData> in xml[from, to) (the content between the start and end tag). `p` = element prefix ('' / 'x:').
 */
export function scanSheetData(xml: string, from: number, to: number, p: string, ctx: SheetScanContext): ScannedSheetData {
  const rows: Array<Array<Cell | undefined> | undefined> = [];
  const rowPropsAll: Array<[number, RowProps]> = [];
  const arrays: RangeBox[] = [];
  const arrayRefs: ScannedSheetData['arrayRefs'] = [];
  const counts = { dataTable: 0, sharedMissing: 0, badStyle: 0, droppedFormatted: 0, badValue: 0, noResult: 0 };
  const masters = new Map<string, { f: string; row: number; col: number }>();
  const formatted: Array<[number, number, Cell]> = [];
  const { sst, isDate, date1904, styleCount } = ctx;
  let inlineScanner: StringItemScanner | undefined;

  const qRow = p + 'row';
  const qC = p + 'c';
  const qF = p + 'f';
  const qV = p + 'v';
  const qIs = p + 'is';
  const closeC = `</${p}c>`;
  const closeF = `</${p}f>`;
  const closeV = `</${p}v>`;
  const closeIs = `</${p}is>`;

  let dataMaxRow = 0;
  let dataMaxCol = 0;
  let curRow = 0;
  let curCol = 0;
  let rowArr: Array<Cell | undefined> | undefined;
  let rowArrRow = 0;
  const ca: CAttrs = { r: undefined, s: 0, t: undefined, cm: false };

  let pos = from;
  while (pos < to) {
    const lt = xml.indexOf('<', pos);
    if (lt < 0 || lt >= to) break;
    const c1 = xml.charCodeAt(lt + 1);
    if (c1 === 47 /* / */ || c1 === 33 /* ! */ || c1 === 63 /* ? */) {
      const gt = xml.indexOf('>', lt + 2);
      pos = gt < 0 ? to : gt + 1;
      continue;
    }
    let k = lt + 1;
    while (k < to) {
      const ch = xml.charCodeAt(k);
      if (ch === 32 || ch === 62 || ch === 47 || ch === 9 || ch === 10 || ch === 13) break;
      k++;
    }
    const gt = xml.indexOf('>', k);
    if (gt < 0) break;
    const selfClose = xml.charCodeAt(gt - 1) === 47;
    const attrEnd = selfClose ? gt - 1 : gt;
    const nameLen = k - lt - 1;

    if (nameLen === qC.length && xml.startsWith(qC, lt + 1)) {
      // ---------------- cell
      scanCAttrs(xml, k, attrEnd, ca);
      let row = curRow;
      let col = curCol + 1;
      if (ca.r) {
        const rc = parseRef(ca.r);
        if (rc[0] > 0) row = rc[0];
        if (rc[1] > 0) col = rc[1];
      }
      if (row < 1) row = 1;
      curRow = row;
      curCol = col;
      let s = ca.s;
      if (s < 0 || s >= styleCount) {
        if (s !== 0) counts.badStyle++;
        s = 0;
      }
      let vText: string | undefined;
      let fText: string | undefined;
      let fAttrs: Record<string, string> | undefined;
      let isInner: string | undefined;
      let next = gt + 1;
      if (!selfClose) {
        const cEnd = xml.indexOf(closeC, gt + 1);
        const innerEnd = cEnd < 0 ? to : cEnd;
        let q = gt + 1;
        while (q < innerEnd) {
          const lt2 = xml.indexOf('<', q);
          if (lt2 < 0 || lt2 >= innerEnd) break;
          if (xml.charCodeAt(lt2 + 1) === 47) {
            const g = xml.indexOf('>', lt2);
            q = g < 0 ? innerEnd : g + 1;
            continue;
          }
          let k2 = lt2 + 1;
          while (k2 < innerEnd) {
            const ch = xml.charCodeAt(k2);
            if (ch === 32 || ch === 62 || ch === 47 || ch === 9 || ch === 10 || ch === 13) break;
            k2++;
          }
          const gt2 = xml.indexOf('>', k2);
          if (gt2 < 0 || gt2 > innerEnd) break;
          const sc = xml.charCodeAt(gt2 - 1) === 47;
          const nl = k2 - lt2 - 1;
          if (nl === qV.length && xml.startsWith(qV, lt2 + 1)) {
            if (sc) {
              vText = '';
              q = gt2 + 1;
            } else {
              const e = xml.indexOf(closeV, gt2);
              vText = xml.slice(gt2 + 1, e < 0 ? innerEnd : e);
              q = e < 0 ? innerEnd : e + closeV.length;
            }
          } else if (nl === qF.length && xml.startsWith(qF, lt2 + 1)) {
            fAttrs = parseAttrString(xml, k2, sc ? gt2 - 1 : gt2);
            if (sc) {
              fText = '';
              q = gt2 + 1;
            } else {
              const e = xml.indexOf(closeF, gt2);
              fText = decodeXmlText(xml.slice(gt2 + 1, e < 0 ? innerEnd : e));
              q = e < 0 ? innerEnd : e + closeF.length;
            }
          } else if (nl === qIs.length && xml.startsWith(qIs, lt2 + 1)) {
            if (sc) {
              isInner = '';
              q = gt2 + 1;
            } else {
              const e = xml.indexOf(closeIs, gt2);
              isInner = xml.slice(gt2 + 1, e < 0 ? innerEnd : e);
              q = e < 0 ? innerEnd : e + closeIs.length;
            }
          } else {
            q = gt2 + 1; // extLst etc.: skip (children are skipped the same way)
          }
        }
        next = cEnd < 0 ? to : cEnd + closeC.length;
      }
      pos = next;
      // a writer that wraps the digits / the index in <![CDATA[ ]]> (or leaves a comment there): read the text, not the markup
      if (vText !== undefined && vText.indexOf('<') >= 0) vText = decodeXmlText(vText);

      // ---------------- value
      let cell: Cell;
      const t = ca.t;
      if (t === 's') {
        if (vText === undefined || vText === '') cell = { t: 'z', v: null };
        else {
          const idx = parseInt(vText, 10);
          const text = sst.texts[idx];
          if (text === undefined) {
            counts.badValue++;
            cell = { t: 's', v: '' };
          } else {
            cell = { t: 's', v: text };
            const rich = sst.rich.get(idx);
            if (rich) cell.rich = rich;
          }
        }
      } else if (t === 'str') {
        cell = vText === undefined ? { t: 'z', v: null } : { t: 's', v: decodeXmlText(vText) };
      } else if (t === 'inlineStr') {
        if (isInner !== undefined) {
          if (!inlineScanner) inlineScanner = new StringItemScanner(p, ctx.pal);
          const r = inlineScanner.scan(isInner);
          cell = { t: 's', v: r.text };
          if (r.rich) cell.rich = r.rich;
        } else if (vText !== undefined) cell = { t: 's', v: decodeXmlText(vText) };
        else cell = { t: 'z', v: null };
      } else if (t === 'b') {
        cell = vText === undefined || vText === '' ? { t: 'z', v: null } : { t: 'b', v: vText === '1' || vText === 'true' || vText === 'TRUE' };
      } else if (t === 'e') {
        cell = vText === undefined ? { t: 'z', v: null } : { t: 'e', v: decodeXmlText(vText) };
      } else if (t === 'd') {
        if (vText === undefined || vText === '') cell = { t: 'z', v: null };
        else {
          const serial = isoToSerial(decodeXmlText(vText));
          if (serial === undefined) {
            counts.badValue++;
            cell = { t: 's', v: decodeXmlText(vText) };
          } else cell = { t: 'd', v: serial };
        }
      } else {
        // 'n' or absent
        if (vText === undefined || vText === '') cell = { t: 'z', v: null };
        else {
          const num = Number(vText);
          if (Number.isFinite(num)) {
            if (isDate[s]) {
              const mode = ctx.dateMode ? ctx.dateMode[s] : 'date';
              const shift = date1904 && (mode === 'date' || (mode === 'time' && num >= 1));
              cell = { t: 'd', v: shift ? num + 1462 : num };
            }
            else cell = { t: 'n', v: num };
          } else {
            counts.badValue++;
            cell = { t: 's', v: decodeXmlText(vText) };
          }
        }
      }
      if (s) cell.s = s;

      // ---------------- formula
      if (fAttrs) {
        const ft = fAttrs.t;
        if (ft === 'shared') {
          const si = fAttrs.si ?? '';
          if (fText) {
            masters.set(si, { f: fText, row, col });
            cell.f = fText;
          } else {
            const m = masters.get(si);
            if (m) cell.f = translateFormula(m.f, row - m.row, col - m.col);
            else counts.sharedMissing++;
          }
        } else if (ft === 'array') {
          if (fText) {
            cell.f = fText;
            const ref = fAttrs.ref || ca.r || '';
            cell.arr = ref;
            if (ca.cm) cell.dyn = true;
            const b = boxOf(ref);
            if (b) arrays.push(b);
            arrayRefs.push({ ref, row, col, dynamic: ca.cm });
          }
        } else if (ft === 'dataTable') {
          counts.dataTable++;
        } else if (fText) {
          cell.f = fText;
        }
      }

      // a formula with no stored result (openpyxl / pandas / XlsxWriter without a value): Excel computes it on load, this reader cannot
      if (cell.f !== undefined && cell.v === null) counts.noResult++;
      const hasData = cell.v !== null || cell.f !== undefined;
      if (hasData) {
        if (row > dataMaxRow) dataMaxRow = row;
        if (col > dataMaxCol) dataMaxCol = col;
        if (!rowArr || row !== rowArrRow) {
          rowArr = rows[row - 1];
          if (!rowArr) rows[row - 1] = rowArr = [];
          rowArrRow = row;
        }
        rowArr[col - 1] = cell;
      } else if (s) {
        formatted.push([row, col, cell]);
      }
      continue;
    }

    if (nameLen === qRow.length && xml.startsWith(qRow, lt + 1)) {
      // ---------------- row
      const a = parseAttrString(xml, k, attrEnd);
      const r = a.r !== undefined ? parseInt(a.r, 10) : NaN;
      curRow = Number.isFinite(r) && r > 0 ? r : curRow + 1;
      curCol = 0;
      const props: RowProps = {};
      let any = false;
      const ht = numAttr(a.ht);
      if (ht !== undefined) {
        props.height = ht;
        any = true;
      }
      if (boolAttr(a.hidden)) {
        props.hidden = true;
        any = true;
      }
      const ol = numAttr(a.outlineLevel);
      if (ol) {
        props.outlineLevel = ol;
        any = true;
      }
      if (boolAttr(a.customFormat)) {
        const rs = numAttr(a.s);
        if (rs !== undefined && rs > 0 && rs < styleCount) {
          props.s = rs;
          any = true;
        }
      }
      if (any) rowPropsAll.push([curRow, props]);
      pos = gt + 1;
      continue;
    }
    pos = gt + 1; // unknown element inside sheetData
  }

  // ---------------- extents, formatted cells near the data, row props
  const rowLimit = dataMaxRow + ROW_SLACK;
  const colLimit = dataMaxCol + COL_SLACK;
  let maxRow = dataMaxRow;
  let maxCol = dataMaxCol;
  for (const [r, c, cell] of formatted) {
    if (r > rowLimit || c > colLimit) {
      counts.droppedFormatted++;
      continue;
    }
    let arr = rows[r - 1];
    if (!arr) rows[r - 1] = arr = [];
    arr[c - 1] = cell;
    if (r > maxRow) maxRow = r;
    if (c > maxCol) maxCol = c;
  }
  const rowProps: Array<RowProps | undefined> = [];
  const propLimit = Math.max(maxRow, rowLimit);
  for (const [r, pr] of rowPropsAll) if (r <= propLimit) rowProps[r - 1] = pr;
  // trim holes past maxRow is not needed: arrays are sparse (holey) by construction
  return { rows, rowProps, maxRow, maxCol, arrays, arrayRefs, counts };
}

// ---------------------------------------------------------------------------------------------
// Everything outside <sheetData>
// ---------------------------------------------------------------------------------------------

export interface SheetRest {
  features: Omit<SheetFeatures, 'tables' | 'arrays'>;
  colProps: Array<ColProps | undefined>;
  defaultColWidth?: number;
  defaultRowHeight?: number;
  /** sheetFormatPr baseColWidth (characters of the default font's digit width; Mac Excel writes 10, the schema default is 8) */
  baseColWidth?: number;
  /** resolved table part paths */
  tableParts: string[];
  sparklineGroups: number;
}

/** xsd:boolean true spellings only ("1" / "true"), exactly what the patcher's <col> attribute tests accept. */
function xsdTrue(v: string | undefined): boolean {
  return v === '1' || v === 'true';
}

function isPlainNumber(s: string): boolean {
  return /^\s*[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?\s*$/.test(s);
}

/** Formula text of an element parsed by parseXml (entities already decoded) → _xHHHH_ decoded. */
function ftext(el: XEl | undefined): string {
  return decodeOoxmlEscapes(innerText(el));
}

export function parseSheetRest(root: XEl | undefined, rels: Rel[]): SheetRest {
  const relById = new Map(rels.map((r) => [r.id, r] as const));
  const out: SheetRest = {
    features: { merges: [], rules: [], hyperlinks: [] },
    colProps: [],
    tableParts: [],
    sparklineGroups: 0,
  };
  if (!root) return out;
  const f = out.features;

  // cols
  for (const c of kids(kid(root, 'cols'), 'col')) {
    const min = Math.max(1, Math.min(MAX_COL, numAttr(c.attrs.min) ?? 0));
    const max = Math.max(min, Math.min(MAX_COL, numAttr(c.attrs.max) ?? min));
    if (!numAttr(c.attrs.min)) continue;
    const props: ColProps = {};
    const w = numAttr(c.attrs.width);
    if (w !== undefined) props.width = w;
    // ColProps.customWidth = a width the USER set: customWidth="1" WITHOUT bestFit. An absent / "0" customWidth is an
    // auto width, and so is an auto-fitted column (Excel writes bestFit="1" customWidth="1" after a double-click
    // auto-fit or a widening on entry): it keeps growing with wider values, so it reads as auto. The flag stays
    // undefined (never false) for auto widths. Only the xsd:boolean spellings "1" / "true" count (the patcher's
    // test of the same two attributes is the same), not boolAttr's lenient "on".
    if (xsdTrue(c.attrs.customWidth) && !xsdTrue(c.attrs.bestFit)) props.customWidth = true;
    if (boolAttr(c.attrs.hidden)) props.hidden = true;
    const ol = numAttr(c.attrs.outlineLevel);
    if (ol) props.outlineLevel = ol;
    const st = numAttr(c.attrs.style);
    if (st) props.s = st;
    for (let i = min; i <= max; i++) out.colProps[i - 1] = props; // shared (read-only) object per span
  }

  // sheetFormatPr
  const fmt = kid(root, 'sheetFormatPr');
  if (fmt) {
    out.defaultColWidth = numAttr(fmt.attrs.defaultColWidth);
    out.defaultRowHeight = numAttr(fmt.attrs.defaultRowHeight);
    out.baseColWidth = numAttr(fmt.attrs.baseColWidth);
  }

  // freeze panes (first sheet view)
  const view = kid(kid(root, 'sheetViews'), 'sheetView');
  const pane = kid(view, 'pane');
  if (pane && (pane.attrs.state === 'frozen' || pane.attrs.state === 'frozenSplit')) {
    const cols = Math.max(0, Math.round(numAttr(pane.attrs.xSplit) ?? 0));
    const rws = Math.max(0, Math.round(numAttr(pane.attrs.ySplit) ?? 0));
    if (cols || rws) f.freeze = { cols, rows: rws };
  }

  // merges
  for (const m of kids(kid(root, 'mergeCells'), 'mergeCell')) {
    const b = boxOf(m.attrs.ref ?? '');
    if (b) f.merges.push(b);
  }

  // autoFilter
  const af = kid(root, 'autoFilter');
  if (af?.attrs.ref) f.autoFilter = af.attrs.ref;

  // conditional formatting (main)
  for (const cf of kids(root, 'conditionalFormatting')) {
    const sqref = (cf.attrs.sqref ?? '').trim();
    if (!sqref) continue;
    for (const rule of kids(cf, 'cfRule')) f.rules.push(withType({ sqref, formulas: cfFormulas(rule, false), kind: 'cf' }, cfTypeOf(rule)));
  }
  // data validation (main)
  for (const dv of kids(kid(root, 'dataValidations'), 'dataValidation')) {
    const sqref = (dv.attrs.sqref ?? '').trim();
    if (!sqref) continue;
    const formulas: string[] = [];
    const f1 = kid(dv, 'formula1');
    const f2 = kid(dv, 'formula2');
    if (f1) formulas.push(ftext(f1));
    if (f2) formulas.push(ftext(f2));
    f.rules.push(withType({ sqref, formulas, kind: 'dv' }, dvTypeOf(dv)));
  }

  // hyperlinks
  for (const h of kids(kid(root, 'hyperlinks'), 'hyperlink')) {
    const ref = h.attrs.ref;
    if (!ref) continue;
    const link: { ref: string; target?: string; location?: string } = { ref };
    const id = relId(h);
    if (id) {
      const rel = relById.get(id);
      if (rel) link.target = rel.target;
    }
    if (h.attrs.location !== undefined) link.location = h.attrs.location;
    f.hyperlinks.push(link);
  }

  // table parts
  for (const tp of kids(kid(root, 'tableParts'), 'tablePart')) {
    const id = relId(tp);
    const rel = id ? relById.get(id) : undefined;
    if (rel && !rel.external) out.tableParts.push(rel.path);
  }

  // extLst: x14 conditional formatting / data validation / sparklines (any nesting, e.g. mc:AlternateContent)
  for (const ext of kids(root, 'extLst')) {
    for (const cf of descendants(ext, 'conditionalFormatting')) {
      const sqref = ftext(kid(cf, 'sqref')).trim();
      if (!sqref) continue;
      for (const rule of kids(cf, 'cfRule')) f.rules.push(withType({ sqref, formulas: cfFormulas(rule, true), kind: 'cf', ext: true }, cfTypeOf(rule)));
    }
    for (const dv of descendants(ext, 'dataValidation')) {
      const sqref = ftext(kid(dv, 'sqref')).trim();
      if (!sqref) continue;
      const formulas: string[] = [];
      const f1 = kid(dv, 'formula1');
      const f2 = kid(dv, 'formula2');
      if (f1) formulas.push(ftext(kid(f1, 'f') ?? f1));
      if (f2) formulas.push(ftext(kid(f2, 'f') ?? f2));
      f.rules.push(withType({ sqref, formulas, kind: 'dv', ext: true }, dvTypeOf(dv)));
    }
    out.sparklineGroups += descendants(ext, 'sparklineGroup').length;
  }
  return out;
}

/** `rule` with its `type` set when known. */
function withType(rule: FeatureRule, type: string | undefined): FeatureRule {
  if (type !== undefined) rule.type = type;
  return rule;
}

/** <cfRule type> (main and x14: 'cellIs', 'expression', 'colorScale', 'dataBar', …); undefined when absent / empty. */
function cfTypeOf(rule: XEl): string | undefined {
  const t = (rule.attrs.type ?? '').trim();
  return t || undefined;
}

/** <dataValidation type> (main and x14). The attribute's schema default is 'none' (any value; message only). */
function dvTypeOf(dv: XEl): string {
  return (dv.attrs.type ?? '').trim() || 'none';
}

/** Formulas of a cfRule: <formula> texts (main) / <xm:f> texts (x14), then formula-valued cfvo thresholds. */
function cfFormulas(rule: XEl, ext: boolean): string[] {
  const out: string[] = [];
  if (!ext) {
    for (const fm of kids(rule, 'formula')) out.push(ftext(fm));
    for (const cfvo of descendants(rule, 'cfvo')) {
      const val = cfvo.attrs.val;
      if (val !== undefined && val !== '' && (cfvo.attrs.type === 'formula' || !isPlainNumber(val))) out.push(decodeOoxmlEscapes(val));
    }
  } else {
    // x14: <xm:f> directly under cfRule, then under cfvo
    for (const fm of kids(rule, 'f')) out.push(ftext(fm));
    for (const cfvo of descendants(rule, 'cfvo')) {
      const fm = kid(cfvo, 'f');
      if (fm) {
        const t = ftext(fm);
        if (cfvo.attrs.type === 'formula' || !isPlainNumber(t)) out.push(t);
      }
    }
  }
  return out;
}

export type { FeatureRule };
