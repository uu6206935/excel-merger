/**
 * The core copy of ONE worksheet of a source workbook into the target package (see addsheets.ts for the overview).
 *
 * The sheet is parsed with the package patcher's own SheetModel (rows / cells as attribute and inner-XML strings, shared formulas expanded)
 * and converted cell by cell:
 *  - style: the source xf is imported into the target styles.xml (fonts / fills / borders / number formats merged and deduplicated, theme
 *    colours converted when the themes differ, named cell styles map to Normal); row and column styles likewise;
 *  - shared strings: re-indexed into the target sst (rich text runs and phonetic data kept);
 *  - dates: a 1904 <-> 1900 date-system difference shifts the date serials by 1462 days (calendar dates; times from 1 day on);
 *  - formulas: sheet references to sheets of the same source follow the renamed copies; a formula that would reach outside the copies
 *    (external workbook, 3D reference, call-out function, table / name that is not copied) keeps its cached value instead;
 *    arrays and dynamic arrays stay arrays (cm re-pointed at the target's metadata).
 * The parts around the cells (sheet properties, views, columns, merges, conditional formats, validations, hyperlinks, page setup,
 * sparklines ...) are rewritten element by element. Drawings, comments, tables, pivot tables and other objects are NOT copied here: their
 * elements and relationships are removed cleanly and reported (ObjectItem); copySheetObjects can add them afterwards.
 */
import { parseAreaText, type Area } from '../../formula/excelRef';
import { normalizeEol } from '../xml';
import { convertNamespaces, isStrict } from './bookinfo';
import { FormulaCopier, type CopyFailReason, type NameView } from './copyformula';
import { copyCfBlock, copyDvContainer, copySparklineGroups, copyX14CfBlock, type RuleCounts, type RuleEnv } from './copyrules';
import type { CopySource, CopyTarget, ObjectItem, SheetCopy } from './copytypes';
import { detachCell } from './dropsheets';
import { to1904Serial, type Engine } from './engine';
import { isRel, type Rel } from './pkg';
import { convertRich } from './richtext';
import { SheetModel, parseF } from './sheet';
import type { Sst } from './sst';
import { convertColorTags } from './styles';
import { defaultColumnPx, excelWidthToPx, pxToExcelWidth } from '../../display/styles';
import { decodeXEsc, esc, findAll, getAttr, replaceElems, setAttr, unesc } from './xml';

/** what copyWorksheet needs to know about one source sheet */
export interface SheetJob {
  src: CopySource;
  target: CopyTarget;
  /** the name of the copy (diagnostics only) */
  newName: string;
  /** the defined names as seen from this sheet */
  names: NameView;
  copier: FormulaCopier;
  dxf(id: number): number | null;
  guid(old: string): string;
  /** the dynamic-array cell-metadata index of the TARGET (cm="n"), resolved by the caller when the source has any dynamic array */
  dynCm: number | null;
  /** lower-cased source name of the sheet (host of its formulas) */
  hostLower: string;
}

/** uri of the extension elements of a worksheet that are copied (x14) or recognised and dropped */
const EXT_CF = '{78c0d931-6437-407d-a8ee-f0aad7539e65}';
const EXT_DV = '{cce6a557-97bc-4b89-adb6-d9c93caab3df}';
const EXT_SPARK = '{05c60535-1f16-4fd2-b633-f4f36f0b64e0}';
const EXT_SLICERS = new Set(['{a8765ba9-456a-4dab-b4f3-acf838c121de}', '{3a4cf648-6aed-40f4-86ff-dc5316d8aed3}', '{bbe1a952-aa13-448e-aadc-164f8a28a991}']);
const EXT_TIMELINES = new Set(['{7e03d99c-dc04-49d9-9315-930204a7b6e9}']);

interface CellAttr {
  /** the cell has an s attribute */
  hasS: boolean;
  t: string | undefined;
  cm: boolean;
  vm: boolean;
  /** other attributes verbatim (ph ...) */
  rest: string;
  /** target xf */
  ns: number;
  mode: 'none' | 'date' | 'time' | 'elapsed';
  /** the attributes of the converted cell (no cm) */
  out: string;
  /** the same for a cell without content */
  empty: string;
}

const ATTR = /\s([\w:.-]+)\s*=\s*"([^"]*)"/g;

export interface NewRelIds {
  newRelId(): string;
}

/**
 * Convert one sheet. `xml`: the source sheet part; `rels`: its relationships (resolved). Returns the converted sheet, or { problem } when the
 * sheet cannot be written back safely (duplicate cells, cells outside the sheet ...).
 */
export async function copyWorksheet(job: SheetJob, xml: string, rels: Array<Rel & { path: string }>, ids: NewRelIds): Promise<SheetCopy | { problem: string }> {
  const { src, target } = job;
  // a Strict OOXML source is read in the transitional namespaces (the sheet is written back in the target's flavour at the end)
  const model = new SheetModel(isStrict(xml) ? convertNamespaces(xml, false) : xml);
  if (model.problem) return { problem: model.problem };
  model.unshare();
  const bsst: Sst = target.sst;
  const out: SheetCopy = {
    xml: '', rels: [], usedNames: new Set(), failures: [], noCache: 0, left: [], vmCells: 0, xfOverflow: 0,
    dropped: { cf: 0, dv: 0, sparklines: 0, hyperlinks: 0, other: [] }, notes: [], formulaCount: 0,
  };
  const failures = new Map<CopyFailReason, { count: number; examples: string[] }>();
  const fail = (reason: CopyFailReason, ref: string) => {
    let f = failures.get(reason);
    if (!f) failures.set(reason, (f = { count: 0, examples: [] }));
    f.count++;
    if (f.examples.length < 3) f.examples.push(ref);
  };

  // ---- style / string maps -------------------------------------------------------------------------------------------------
  const mapXf = (s: number): number => {
    const hit = src.xfCache.get(s);
    if (hit !== undefined) return hit;
    const r = target.styles.importXfLimited(src.styles, s, src.themeRgb, src.xfCache, src.indexedRgb, target.xfMax);
    if (r < 0) { out.xfOverflow++; return 0; }
    return r;
  };
  const sstMap = new Int32Array(src.sst.size).fill(-1);
  const mapSst = (i: number): number => {
    if (i < sstMap.length && sstMap[i] >= 0) return sstMap[i];
    const rich = i < sstMap.length ? src.sst.richXml(i) : null;
    const v = rich ? bsst.addRich(convertRich(rich, src, target.styles)) : bsst.add(src.sst.text(i));
    if (i < sstMap.length) sstMap[i] = v;
    return v;
  };
  const attrCache = new Map<string, CellAttr>();
  const parseAttr = (a: string): CellAttr => {
    let s = 0, hasS = false, t: string | undefined, cm = false, vm = false;
    let rest = '';
    ATTR.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = ATTR.exec(a))) {
      const k = m[1];
      if (k === 's') { s = Number(m[2]) || 0; hasS = true; }
      else if (k === 't') t = m[2];
      else if (k === 'cm') cm = true;
      else if (k === 'vm') vm = true;
      else rest += m[0];
    }
    const ns = mapXf(s);
    return {
      hasS, t, cm, vm, rest, ns, mode: src.styles.dateMode(s),
      out: (ns || hasS ? ` s="${ns}"` : '') + (t !== undefined ? ` t="${t}"` : '') + rest,
      empty: (hasS ? ` s="${ns}"` : '') + rest,
    };
  };

  // ---- cells -------------------------------------------------------------------------------------------------------------------
  /** array ranges whose master formula became a value: their spilled text results are plain strings now */
  const detachedArrays: Area[] = [];
  const shift = src.dateShift;
  let formulas = 0;
  for (const row of model.rows) {
    if (row.a) {
      const rs = getAttr(row.a, 's');
      if (rs !== undefined) row.a = setAttr(row.a, 's', String(mapXf(Number(rs) || 0)));
    }
    const rowNo = row.r;
    for (const cell of row.cells) {
      let info = attrCache.get(cell.a);
      if (!info) { info = parseAttr(cell.a); attrCache.set(cell.a, info); }
      const x = cell.x;
      if (x === null || x === '') {
        cell.a = info.empty;
        cell.x = null;
        continue;
      }
      if (info.vm) out.vmCells++;
      cell.a = info.out;
      if (x.indexOf('<f') >= 0) {
        const f = parseF(x);
        if (f) {
          formulas++;
          const ft = getAttr(f.attrs, 't');
          if (ft === 'dataTable') {
            if (shift && info.t === undefined) cell.x = shiftValue(x, info.mode, shift);
            continue;
          }
          if (!f.text && ft !== 'array') {
            // a formula without text (a shared dependent whose master is missing): the value stays
            detachCell(cell, f, bsst);
            continue;
          }
          const res = job.copier.rewrite(f.text, job.hostLower, job.names);
          if (res.ok) {
            for (const n of res.names) out.usedNames.add(n);
            let nx = res.changed ? x.slice(0, f.start) + `<f${f.attrs}>${esc(res.text)}</f>` + x.slice(f.end) : x;
            if (shift && (info.t === undefined || info.t === 'n')) nx = shiftValue(nx, info.mode, shift);
            cell.x = nx;
            // a dynamic-array master keeps its array nature: cm points at the target's metadata
            if (info.cm && ft === 'array' && job.dynCm !== null) cell.a = info.out + ` cm="${job.dynCm}"`;
            continue;
          }
          fail(res.reason, colRef(cell.c, rowNo));
          if (!/<v[\s>]|<is>/.test(x.slice(0, f.start) + x.slice(f.end))) out.noCache++;
          detachCell(cell, f, bsst);
          if (ft === 'array') {
            const a = parseAreaText((getAttr(f.attrs, 'ref') ?? '').replace(/\$/g, ''));
            if (a) detachedArrays.push(a);
          }
          if (cell.x !== null && shift && (info.t === undefined || info.t === 'n')) cell.x = shiftValue(cell.x, info.mode, shift);
          continue;
        }
      }
      // ---- a constant ----------------------------------------------------------------------------------------------------------
      const t = info.t;
      if (t === 's') {
        const m = /<v>\s*(\d+)\s*<\/v>/.exec(x);
        if (m) cell.x = `<v>${mapSst(Number(m[1]))}</v>`;
        else { cell.x = null; cell.a = info.empty; }
      } else if (t === 'inlineStr') {
        if ((src.themeRgb || src.indexedRgb) && x.indexOf('<r>') >= 0) {
          cell.x = x.replace(/<is>([\s\S]*?)<\/is>/, (_m, inner: string) => `<is>${convertRich(inner, src, target.styles)}</is>`);
        }
      } else if (t === 'str') {
        // a formula-less formula string: a cell of a spill range whose master became a value
        if (detachedArrays.length && detachedArrays.some((a) => inArea(a, rowNo, cell.c))) {
          const v = /<v>([\s\S]*?)<\/v>/.exec(x);
          cell.a = (info.ns || info.hasS ? ` s="${info.ns}"` : '') + ' t="s"' + info.rest;
          if (v) cell.x = `<v>${bsst.add(decodeXEsc(unesc(normalizeEol(v[1]))))}</v>`;
          else { cell.x = null; cell.a = info.empty; }
        }
      } else if (shift && (t === undefined || t === 'n')) {
        cell.x = shiftValue(x, info.mode, shift);
      }
    }
  }
  out.formulaCount = formulas;

  // ---- the part around the cells ----------------------------------------------------------------------------------------
  const counts: RuleCounts = { cfDropped: 0, dvDropped: 0, sparklinesDropped: 0, dvDateShifted: 0 };
  const env: RuleEnv = {
    formula: (f: string): string | null => {
      const r = job.copier.rewrite(f, job.hostLower, job.names);
      if (!r.ok) { fail(r.reason, '条件付き書式・入力規則など'); return null; }
      for (const n of r.names) out.usedNames.add(n);
      return r.text;
    },
    base: target.styles,
    src: src.styles,
    themeRgb: src.themeRgb,
    indexedRgb: src.indexedRgb,
    dateShift: shift,
    dxf: job.dxf,
    guid: job.guid,
    counts,
  };
  model.head = copyHead(model.head, job, mapXf).replace(/<dimension\b[^>]*>/, `<dimension ref="${model.usedRange()}"/>`);
  model.tail = await copyTail(model.tail, job, env, rels, out, ids);
  out.dropped.cf = counts.cfDropped;
  out.dropped.dv = counts.dvDropped;
  out.dropped.sparklines = counts.sparklinesDropped;
  out.failures = [...failures].map(([reason, v]) => ({ reason, count: v.count, examples: v.examples }));
  out.xml = model.serialize(job.newName);
  if (target.strict) out.xml = convertNamespaces(out.xml, true);
  return out;
}

function colRef(c: number, r: number): string {
  let s = '';
  for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s + r;
}

function inArea(a: Area, row: number, col: number): boolean {
  return row >= Math.min(a.r1, a.r2) && row <= Math.max(a.r1, a.r2) && col >= Math.min(a.c1, a.c2) && col <= Math.max(a.c1, a.c2);
}

/** `<v>number</v>` of a cell with a date serial moved to the other date system (the reader's rule, see premerge toBase) */
function shiftValue(x: string, mode: 'none' | 'date' | 'time' | 'elapsed', shift: number): string {
  if (mode === 'none' || mode === 'elapsed') return x;
  return x.replace(/<v>([^<]*)<\/v>/, (m0, v: string) => {
    const n = Number(v);
    if (!Number.isFinite(n) || v.trim() === '') return m0;
    const r = shift > 0 ? (mode === 'date' || n >= 1 ? n + shift : n) : to1904Serial(n, mode);
    return r === n ? m0 : `<v>${r}</v>`;
  });
}

// ---------------------------------------------------------------------------------------------------------------------
// head: sheetPr / dimension / sheetViews / sheetFormatPr / cols
// ---------------------------------------------------------------------------------------------------------------------

function copyHead(head0: string, job: SheetJob, mapXf: (s: number) => number): string {
  const { src } = job;
  let head = head0;
  // the revision uid of the worksheet belongs to the source sheet
  head = head.replace(/\sxr\d*:uid="[^"]*"/g, '');
  // sheetPr: no VBA code name (the target has other code modules); tab colour in the target's theme
  head = head.replace(/<sheetPr\b[^>]*?(?:\/>|>[\s\S]*?<\/sheetPr>)/, (el) => convertColorTags(el.replace(/^<sheetPr\b[^>]*>/, (o) => setAttr(o, 'codeName', null)), src.themeRgb, src.indexedRgb));
  // views: no selected tab (the target decides which sheet is shown), no pivot selections
  head = head.replace(/<sheetView\b[^>]*>/g, (o) => setAttr(o, 'tabSelected', null));
  head = head.replace(/<pivotSelection\b[^>]*?(?:\/>|>[\s\S]*?<\/pivotSelection>)/g, '');
  // columns: the column format is a cell format of the source
  head = head.replace(/<col\b[^>]*>/g, (o) => {
    const s = getAttr(o, 'style');
    return s === undefined ? o : setAttr(o, 'style', String(mapXf(Number(s) || 0)));
  });
  // Column widths are stored in characters of the workbook's default font: a sheet whose default font is narrower / wider than the target's
  // (Calibri 7 px per digit, 游ゴシック 8 px) keeps its PIXEL widths — the copy looks like the original, not 14 % wider
  const sm = src.mdw, tm = job.target.mdw;
  if (sm !== undefined && tm !== undefined && sm !== tm) {
    const w = (chars: number) => String(Math.round(pxToExcelWidth(excelWidthToPx(chars, sm), tm) * 1e6) / 1e6);
    head = head.replace(/<col\b[^>]*>/g, (o) => {
      const v = Number(getAttr(o, 'width'));
      return v > 0 ? setAttr(o, 'width', w(v)) : o;
    });
    const fp = /<sheetFormatPr\b[^>]*>/.exec(head)?.[0];
    const dcw = Number(getAttr(fp ?? '', 'defaultColWidth'));
    const base = Number(getAttr(fp ?? '', 'baseColWidth') ?? 8);
    const px = dcw > 0 ? excelWidthToPx(dcw, sm) : defaultColumnPx(sm, Number.isFinite(base) && base >= 0 ? base : 8);
    const width = String(Math.round(pxToExcelWidth(px, tm) * 1e6) / 1e6);
    if (fp) head = head.replace(fp, setAttr(fp, 'defaultColWidth', width));
    else {
      // no <sheetFormatPr>: the default column width follows the default font, which differs in the target -> state it
      const el = `<sheetFormatPr defaultRowHeight="15" defaultColWidth="${width}"/>`;
      const at = /<\/sheetViews>|<sheetViews\b[^>]*\/>/.exec(head) ?? /<dimension\b[^>]*\/>/.exec(head) ?? /<sheetPr\b[^>]*\/>|<\/sheetPr>/.exec(head);
      if (at) head = head.slice(0, at.index + at[0].length) + el + head.slice(at.index + at[0].length);
      else head = head.replace(/(<worksheet\b[^>]*>)/, `$1${el}`);
    }
  }
  return head;
}

// ---------------------------------------------------------------------------------------------------------------------
// tail
// ---------------------------------------------------------------------------------------------------------------------

interface Top {
  tag: string;
  text: string;
}

/** the top-level elements of the text after <sheetData> (none of them nests an element of its own name) */
export function splitTop(xml: string): Top[] {
  const out: Top[] = [];
  const re = /<([A-Za-z_][\w.:-]*)(?=[\s/>])[^>]*?(\/?)>/y;
  let i = 0;
  while (i < xml.length) {
    const lt = xml.indexOf('<', i);
    if (lt < 0) break;
    if (xml.startsWith('</', lt)) break; // </worksheet>
    if (xml.startsWith('<!--', lt)) { const e = xml.indexOf('-->', lt); i = e < 0 ? xml.length : e + 3; continue; }
    re.lastIndex = lt;
    const m = re.exec(xml);
    if (!m) break;
    if (m[2] === '/') { out.push({ tag: m[1], text: m[0] }); i = lt + m[0].length; continue; }
    const close = `</${m[1]}>`;
    const e = xml.indexOf(close, lt + m[0].length);
    if (e < 0) break;
    out.push({ tag: m[1], text: xml.slice(lt, e + close.length) });
    i = e + close.length;
  }
  return out;
}

/** elements that are copied as they are (no sheet reference, no style, no relationship in them) */
const KEEP = new Set(['sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'mergeCells', 'printOptions', 'pageMargins', 'headerFooter', 'rowBreaks', 'colBreaks', 'cellWatches', 'ignoredErrors']);
/** elements that are removed because the objects / relationships they point at are not copied (reported through describeObjects) */
const OBJECT_ELEMENTS = new Set(['mc:AlternateContent', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'tableParts', 'webPublishItems', 'customSheetViews', 'customProperties', 'smartTags']);

async function copyTail(tail: string, job: SheetJob, env: RuleEnv, srcRels: Array<Rel & { path: string }>, out: SheetCopy, ids: NewRelIds): Promise<string> {
  const top = splitTop(tail);
  const kept: string[] = [];
  const droppedCf = new Set<string>();
  out.left.push(...(await describeObjects(top, srcRels, job.src.eng)));
  for (const el of top) {
    switch (el.tag) {
      case 'conditionalFormatting': {
        const y = copyCfBlock(el.text, env, droppedCf);
        if (y) kept.push(y);
        break;
      }
      case 'dataValidations': {
        const y = copyDvContainer(el.text, env, false);
        if (y) kept.push(y);
        break;
      }
      case 'hyperlinks': {
        const y = copyHyperlinks(el.text, job, srcRels, out, ids);
        if (y) kept.push(y);
        break;
      }
      case 'phoneticPr': {
        const id = getAttr(el.text, 'fontId');
        if (id === undefined) { kept.push(el.text); break; }
        const nid = job.target.styles.importFont(job.src.styles, Number(id), job.src.themeRgb, job.src.indexedRgb);
        kept.push(nid === null ? setAttr(el.text, 'fontId', null) : setAttr(el.text, 'fontId', String(nid)));
        break;
      }
      case 'pageSetup':
        // the printer settings are a binary part of the source (r:id): not copied; paper size, orientation, scale ... stay
        kept.push(setAttr(el.text, 'r:id', null));
        break;
      case 'dataConsolidate':
        out.dropped.other.push('データの統合の設定');
        break;
      case 'extLst': {
        const y = copyExtLst(el.text, env, droppedCf, out);
        if (y) kept.push(y);
        break;
      }
      default:
        if (KEEP.has(el.tag)) kept.push(el.text);
        else if (!OBJECT_ELEMENTS.has(el.tag)) out.dropped.other.push(el.tag);
    }
  }
  // the revision uids (xr:uid, xr2:uid ...) identify elements of the SOURCE file
  return (kept.join('') + '</worksheet>').replace(/\sxr\d*:uid="[^"]*"/g, '');
}

/** hyperlinks: external targets get a new relationship, locations follow the renamed copies */
function copyHyperlinks(container: string, job: SheetJob, srcRels: Array<Rel & { path: string }>, out: SheetCopy, ids: NewRelIds): string {
  const open = /^<hyperlinks\b[^>]*>/.exec(container)?.[0];
  if (!open || open.endsWith('/>')) return '';
  const inner = container.slice(open.length, container.length - '</hyperlinks>'.length);
  let kept = 0;
  const body = replaceElems(inner, findAll(inner, 'hyperlink'), (h) => {
    let x = h.text;
    const rid = getAttr(h.open, 'r:id');
    const loc = getAttr(h.open, 'location');
    let ok = false;
    if (rid !== undefined) {
      const r = srcRels.find((q) => q.id === rid);
      if (r && r.external && r.target) {
        const id = ids.newRelId();
        out.rels.push({ id, type: convertNamespaces(convertNamespaces(r.type, false), job.target.strict), target: r.target, external: true });
        x = setAttr(x, 'r:id', id);
        ok = true;
      } else x = setAttr(x, 'r:id', null);
    }
    if (loc !== undefined) {
      const res = job.copier.rewrite(loc, job.hostLower, job.names);
      if (res.ok) {
        for (const n of res.names) out.usedNames.add(n);
        if (res.changed) x = setAttr(x, 'location', res.text);
        ok = true;
      } else x = setAttr(x, 'location', null);
    }
    if (!ok) { out.dropped.hyperlinks++; return ''; }
    kept++;
    return x;
  });
  return kept ? open + body + '</hyperlinks>' : '';
}

const setExtOpen = (o: string): string => (o.endsWith('/>') ? o.slice(0, -2).trimEnd() + '>' : o);

/** the extension list: conditional formats / validations / sparklines of the x14 family are copied, the rest is dropped */
function copyExtLst(extLst: string, env: RuleEnv, droppedCf: Set<string>, out: SheetCopy): string {
  const open = /^<extLst\b[^>]*>/.exec(extLst)?.[0];
  if (!open || open.endsWith('/>')) return '';
  const inner = extLst.slice(open.length, extLst.length - '</extLst>'.length);
  const parts: string[] = [];
  for (const e of findAll(inner, 'ext')) {
    const uri = (getAttr(e.open, 'uri') ?? '').toLowerCase();
    const body = e.selfClosing ? '' : e.inner;
    let y: string;
    if (uri === EXT_CF) {
      y = replaceElems(body, findAll(body, 'x14:conditionalFormatting'), (b) => copyX14CfBlock(b.text, env, droppedCf));
      y = y.replace(/<x14:conditionalFormattings\b[^>]*>\s*<\/x14:conditionalFormattings>/g, '');
    } else if (uri === EXT_DV) {
      y = replaceElems(body, findAll(body, 'x14:dataValidations'), (b) => copyDvContainer(b.text, env, true));
    } else if (uri === EXT_SPARK) {
      y = replaceElems(body, findAll(body, 'x14:sparklineGroups'), (b) => copySparklineGroups(b.text, env));
    } else if (EXT_SLICERS.has(uri) || EXT_TIMELINES.has(uri)) {
      if (!out.left.some((p) => p.kind === 'slicers')) {
        const n = Math.max(1, (body.match(/<x14:slicer\b|<x15:timelineRef\b/g) ?? []).length);
        out.left.push({ kind: 'slicers', label: `スライサー・タイムライン ${n} 個`, count: n });
      }
      continue;
    } else {
      out.dropped.other.push(uri || 'ext');
      continue;
    }
    if (y.trim()) parts.push(setExtOpen(e.open) + y + '</ext>');
  }
  return parts.length ? open + parts.join('') + '</extLst>' : '';
}

// ---------------------------------------------------------------------------------------------------------------------
// the objects the sheet owns (reported; copySheetObjects can copy them)
// ---------------------------------------------------------------------------------------------------------------------

const prefixed = (name: string) => new RegExp(`<(?:[\\w.-]+:)?${name}[\\s>/]`, 'g');

async function describeObjects(top: Top[], rels: Array<Rel & { path: string }>, eng: Engine): Promise<ObjectItem[]> {
  const items: ObjectItem[] = [];
  const has = (tag: string) => top.find((t) => t.tag === tag);
  const relOfType = (kind: string) => rels.filter((r) => !r.external && isRel(r.type, kind));
  const drawing = has('drawing');
  if (drawing) {
    const rid = getAttr(drawing.text, 'r:id');
    const rel = rels.find((r) => r.id === rid);
    const x = rel && !rel.external ? await eng.pkg.read(rel.path) : null;
    const parts: string[] = [];
    let n = 0;
    if (x) {
      const charts = (x.match(/<(?:[\w.-]+:)?chart\b[^>]*\sr:id=/g) ?? []).length;
      const pics = x.match(prefixed('pic'))?.length ?? 0;
      const shapes = (x.match(prefixed('sp'))?.length ?? 0) + (x.match(prefixed('cxnSp'))?.length ?? 0) + (x.match(prefixed('grpSp'))?.length ?? 0);
      const diagrams = (x.match(/<dgm:relIds\b/g) ?? []).length;
      const slicers = (x.match(/<(?:[\w.-]+:)?(?:slicer|timeslicer)\b/g) ?? []).length;
      if (charts) { parts.push(`グラフ ${charts} 個`); n += charts; }
      if (pics) { parts.push(`画像 ${pics} 個`); n += pics; }
      if (shapes) { parts.push(`図形 ${shapes} 個`); n += shapes; }
      if (diagrams) { parts.push(`SmartArt ${diagrams} 個`); n += diagrams; }
      if (slicers) { parts.push(`スライサー ${slicers} 個`); n += slicers; }
    }
    items.push({ kind: 'drawing', label: parts.length ? parts.join('、') : '図形・画像・グラフ', count: Math.max(1, n) });
  }
  const comments = relOfType('comments');
  if (comments.length) {
    let n = 0;
    for (const c of comments) n += ((await eng.pkg.read(c.path)) ?? '').match(/<comment\b/g)?.length ?? 0;
    items.push({ kind: 'comments', label: `コメント ${n} 件`, count: n });
  } else if (relOfType('threadedComment').length) {
    items.push({ kind: 'comments', label: 'コメント', count: 1 });
  }
  if (has('legacyDrawing') || relOfType('vmlDrawing').length) {
    let controls = 0;
    for (const v of relOfType('vmlDrawing')) {
      const x = (await eng.pkg.read(v.path)) ?? '';
      controls += (x.match(/ObjectType="(?!Note)[^"]*"/g) ?? []).length;
    }
    if (controls) items.push({ kind: 'legacyDrawing', label: `フォームコントロール ${controls} 個`, count: controls });
    else if (!items.some((i) => i.kind === 'comments')) items.push({ kind: 'legacyDrawing', label: 'コメント', count: 1 });
  }
  const tp = has('tableParts');
  if (tp) {
    const n = Number(getAttr(tp.text, 'count') ?? (tp.text.match(/<tablePart\b/g) ?? []).length) || 1;
    items.push({ kind: 'tables', label: `テーブル ${n} 個（テーブルの機能は外し、セルの値と書式は残しました）`, count: n });
  }
  const pv = relOfType('pivotTable').length;
  if (pv) items.push({ kind: 'pivotTables', label: `ピボットテーブル ${pv} 個（セルの値だけを残しました）`, count: pv });
  if (has('picture')) items.push({ kind: 'background', label: '背景画像', count: 1 });
  const ole = has('oleObjects');
  if (ole) { const n = (ole.text.match(/<oleObject\b/g) ?? []).length || 1; items.push({ kind: 'oleObjects', label: `埋め込みオブジェクト ${n} 個`, count: n }); }
  const ctl = has('controls') ?? has('mc:AlternateContent');
  if (ctl) { const n = (ctl.text.match(/<control\b/g) ?? []).length || 1; items.push({ kind: 'controls', label: `コントロール ${n} 個`, count: n }); }
  if (has('legacyDrawingHF') || has('drawingHF')) items.push({ kind: 'headerFooterImages', label: 'ヘッダー・フッターの画像', count: 1 });
  if (has('customSheetViews')) items.push({ kind: 'customViews', label: 'ユーザー設定のビュー', count: 1 });
  if (has('customProperties')) items.push({ kind: 'customProperties', label: 'シートのユーザー設定プロパティ', count: 1 });
  return items;
}
