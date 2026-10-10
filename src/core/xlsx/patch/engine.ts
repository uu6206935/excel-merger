/**
 * The package patcher engine (owner: patcher P).
 *
 * Applies Excel structural ops (excelRef StructOp: insert / delete / move of columns or rows), cell writes and
 * column widths to an .xlsx package the way Excel does, touching only the XML that carries coordinates or
 * references; every other byte is kept.
 *
 * Cost model: the op's sheet is parsed once into a cell model (SheetModel); per op, cells are remapped and the
 * small per-sheet features (merges, CF/DV, tables, hyperlinks, autofilter, panes, cols, anchors ...) are rewritten.
 * Formulas (cells of every sheet, names, charts, pivot sources ...) are rewritten ONCE at the end with the whole
 * op list (one tokenization per formula; a written cell only receives the ops logged after its write).
 */
import {
  adjustArea,
  MAX_COL,
  MAX_ROW,
  mapIndex,
  movePerm,
  parseAreaText,
  quoteSheetName,
  renderArea,
  tableMoveKind,
  translateFormula,
  type Area,
  type StructOp,
} from '../../formula/excelRef';
import type { CellWrite } from '../../types';
import { isUnsafeFormula } from '../../export/comJob';
import { toUiForm } from '../../formula/forms';
import { formatValue } from '../../display/numfmt';
import { defaultColumnPx, excelWidthToPx } from '../../display/styles';
import { adjustDrawingXml, adjustSheetObjectAnchors, adjustVmlXml } from './anchors';
import { fontMdw, numberTextPx } from './autowidth';
import { ColsModel } from './cols';
import {
  adjustAreaEntry,
  adjustFormulaOps,
  adjustPlainArea,
  escapeSrefItem,
  parseSqref,
  renameTableColumnInFormula,
  renderSqref,
  stripAbs,
  type OpEntry,
  type TableHit,
} from './formula';
import { METADATA_CT, METADATA_REL, ensureDynamicArrayMetadata } from './metadata';
import { Pkg, PatchError, REL_TYPE, dirOf, isRel, parseRels, relsPathOf, resolveTarget } from './pkg';
import { adjustRulesInTail, mapRuleFormulas, removeEmptyExt, renumberPriorities } from './rules';
import { SheetModel, cellRef, cellStyle, findCell, parseF, type CellM, type RowM } from './sheet';
import { Sst } from './sst';
import { Styles } from './styles';
import { adjustSheetViews } from './views';
import { esc, escAttr, escText, findAll, findFirst, getAttr, nsPrefix, reEsc, setAttr, unesc } from './xml';

export interface SheetInfo {
  name: string;
  path: string;
  rid: string;
  /** position in <sheets> (localSheetId) */
  index: number;
  worksheet: boolean;
}

export interface TableState {
  path: string;
  rid: string;
  xml: string;
  /** displayName (used in formulas) */
  name: string;
  ref: Area;
  header: number;
  totals: number;
  removed: boolean;
  dirty: boolean;
}

export interface SheetState {
  info: SheetInfo;
  model: SheetModel;
  cols: ColsModel;
  tables: TableState[];
  /** legacy (CSE) multi-cell array ranges, null = recompute */
  arrays: Area[] | null;
  splitCf: boolean;
  changed: boolean;
  /** the sheet may hold a dynamic-array (cm="n") master: only then does a structural op have spills to collapse (a 100 MB sheet is not scanned per op otherwise) */
  spillPossible: boolean;
}

export interface StructOptions {
  /** premerge: no "format same as left/above" (no style, width or CF inheritance) */
  clean?: boolean;
  /** premerge row append: ranges ending at lastDataRow (starting at/above firstDataRow) grow */
  ext?: { firstDataRow: number; lastDataRow: number };
  /** with `ext`: false = tables / autofilters do NOT grow over the inserted rows (PremergeOptions.extendTable) */
  extTables?: boolean;
  /** inserted rows copy the row above exactly (premerge appended rows: no border filtering) */
  exactFormat?: boolean;
}

const REL = (t: string) => (r: { type: string }) => r.type.endsWith('/' + t);

const MSG = {
  noSheet: (n: string) => `シート「${n}」が見つかりません。`,
  mergeSplit: 'この操作は結合したセルには行えません（結合セルの一部だけを移動することはできません）。',
  array: '配列の一部を変更することはできません（配列数式の範囲の一部だけを挿入・削除・移動することはできません）。',
  tableMove: 'テーブル内の他のセルに予想外の影響を与える可能性があるため、テーブルの列をテーブルの外へ移動したり、外からテーブル内へ移動したりすることはできません。',
  tableHeader: 'テーブルの見出し行は削除できません。',
  pushOff: 'シートの端にあるデータが押し出されてしまうため、挿入できません（最終列・最終行付近にデータがあります）。',
  badOp: '列・行の操作の指定が正しくありません。',
};

/**
 * The serial as stored in a 1904 workbook for a doc value in the 1900 system (the inverse of the reader's conversion: dates
 * always shift by 1462 days, a time of day only when it carries a day count, durations / other formats never). Values
 * below the 1904 epoch cannot be represented and stay as they are.
 */
export function to1904Serial(v: number, mode: 'none' | 'date' | 'time' | 'elapsed'): number {
  if (mode === 'date') return v >= 1462 ? v - 1462 : v;
  if (mode === 'time') return v >= 1463 ? v - 1462 : v;
  return v;
}

export class Engine {
  pkg: Pkg;
  wbPath = '';
  wbXml = '';
  sheets: SheetInfo[] = [];
  states = new Map<string, SheetState>();
  entries: OpEntry[] = [];
  private sstObj: Sst | null = null;
  private sstPath: string | null = null;
  private stylesObj: Styles | null = null;
  private stylesPath: string | null = null;
  /** dynamic-array metadata: cm index + the part (written at finish when changed) */
  private dynMeta: { cm: number; path: string; xml: string; changed: boolean } | null = null;
  /**
   * Formulas that are adjusted ONCE at finish with the whole op list may have been brought up to date earlier (a
   * table column rename rewrites every structured reference at the time of the rename): entries before these
   * indexes are already applied. Defined names: namesEp; table part formulas: tableEp (by part path).
   */
  private namesEp = 0;
  private tableEp = new Map<string, number>();
  /**
   * Warnings of the replay (the Excel runner's warn lines): a formula / text that would make Excel reach outside the workbook was
   * written as text (UNSAFE_FORMULA). At most 20 cells are named, then one summary line, like XlRunner (replayLog hands them on).
   */
  readonly warnings: Array<{ code: 'UNSAFE_FORMULA'; message: string }> = [];
  private neutralized = 0;

  private constructor(pkg: Pkg) {
    this.pkg = pkg;
  }

  /** The workbook stores dates in the 1904 system (Mac Excel 2011 and older default; <workbookPr date1904>). */
  get date1904(): boolean {
    return /<workbookPr\b[^>]*\sdate1904\s*=\s*"(1|true)"/.test(this.wbXml);
  }

  static async open(bytes: Uint8Array): Promise<Engine> {
    const e = new Engine(await Pkg.load(bytes));
    await e.init();
    return e;
  }

  private async init(): Promise<void> {
    this.wbPath = await this.pkg.workbookPath();
    const wb = await this.pkg.read(this.wbPath);
    if (!wb) throw new PatchError('ブックの構造（workbook.xml）が見つかりません。ファイルが壊れている可能性があります。');
    this.wbXml = wb;
    const rels = await this.pkg.relTargets(this.wbPath);
    const sheetsEl = findFirst(wb, 'sheets');
    let i = 0;
    for (const m of (sheetsEl?.inner ?? '').matchAll(/<sheet\b[^>]*>/g)) {
      const name = getAttr(m[0], 'name') ?? '';
      const rid = getAttr(m[0], 'r:id') ?? '';
      const rel = rels.find((r) => r.id === rid);
      this.sheets.push({ name, rid, path: rel ? rel.path : '', index: i++, worksheet: !!rel && isRel(rel.type, 'worksheet') });
    }
  }

  sheet(name: string): SheetInfo {
    const l = name.toLowerCase();
    const s = this.sheets.find((x) => x.name.toLowerCase() === l);
    if (!s || !s.worksheet || !s.path) throw new PatchError(MSG.noSheet(name));
    return s;
  }

  async state(name: string): Promise<SheetState> {
    const info = this.sheet(name);
    const ex = this.states.get(info.path);
    if (ex) return ex;
    const xml = await this.pkg.read(info.path);
    if (xml == null) throw new PatchError(MSG.noSheet(name));
    const model = new SheetModel(xml);
    model.unshare();
    const st: SheetState = { info, model, cols: new ColsModel(model.head), tables: [], arrays: null, splitCf: false, changed: false, spillPossible: xml.indexOf('cm="') >= 0 };
    st.tables = await this.loadTables(info);
    this.states.set(info.path, st);
    return st;
  }

  private async loadTables(info: SheetInfo): Promise<TableState[]> {
    const out: TableState[] = [];
    for (const r of (await this.pkg.relTargets(info.path)).filter(REL('table'))) {
      const xml = await this.pkg.read(r.path);
      if (!xml) continue;
      const open = /<table\b[^>]*>/.exec(xml)?.[0] ?? '';
      const ref = parseAreaText(getAttr(open, 'ref') ?? '');
      if (!ref) continue;
      out.push({
        path: r.path,
        rid: r.id,
        xml,
        name: getAttr(open, 'displayName') ?? getAttr(open, 'name') ?? '',
        ref: norm(ref),
        header: Number(getAttr(open, 'headerRowCount') ?? 1),
        totals: Number(getAttr(open, 'totalsRowCount') ?? 0),
        removed: false,
        dirty: false,
      });
    }
    return out;
  }

  async sst(): Promise<Sst> {
    if (this.sstObj) return this.sstObj;
    this.sstPath = await this.pkg.ensureSharedStrings(this.wbPath);
    this.sstObj = new Sst((await this.pkg.read(this.sstPath)) ?? '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>');
    return this.sstObj;
  }

  async styles(): Promise<Styles> {
    if (this.stylesObj) return this.stylesObj;
    const rel = (await this.pkg.relTargets(this.wbPath)).find((r) => isRel(r.type, 'styles'));
    this.stylesPath = rel ? rel.path : null;
    this.stylesObj = new Styles(rel ? await this.pkg.read(rel.path) : null);
    return this.stylesObj;
  }

  // -------------------------------------------------------------------------------------------------
  // structural ops
  // -------------------------------------------------------------------------------------------------

  async struct(sheetName: string, op: StructOp, opts: StructOptions = {}): Promise<void> {
    validateOp(op);
    const st = await this.state(sheetName);
    st.model.assertWritable(st.info.name);
    if (op.type === 'move' && op.to >= op.from && op.to <= op.from + op.count) return; // no-op move
    this.check(st, op);
    this.collapseSpills(st, op);
    const entry: OpEntry = { sheet: st.info.name, op };
    if (opts.ext) entry.ext = opts.ext;
    if (op.type === 'delete') {
      const hits = new Map<string, TableHit>();
      for (const t of st.tables) {
        if (t.removed) continue;
        const h = tableHit(t, op);
        if (h) hits.set(t.name.toLowerCase(), h);
      }
      if (hits.size) entry.tableHits = hits;
    }
    const inherit = !opts.clean;
    this.moveCells(st, op, inherit, inherit && op.type === 'insert' && !opts.exactFormat ? await this.styles() : null);
    if (op.axis === 'col') st.cols.apply(op, inherit);
    const tableEntry: OpEntry = opts.ext && opts.extTables === false ? { sheet: entry.sheet, op, tableHits: entry.tableHits } : entry;
    await this.applyTables(st, op, tableEntry);
    this.applyHead(st, op);
    this.applyTail(st, op, entry, inherit, tableEntry);
    await this.applyRelated(st, op);
    if (st.arrays) st.arrays = st.arrays.map((a) => adjustArea(a, op)).filter((a): a is Area => !!a);
    st.changed = true;
    this.entries.push(entry);
  }

  /**
   * Dynamic-array (spill) formulas: the spilled cells hold only cached values that Excel recomputes on load. When
   * an op cuts through a spill range (insert inside it, delete / move part of it) those values must not travel as
   * constants: they are cleared and the master's ref is collapsed to the master cell (Excel re-spills on load,
   * fullCalcOnLoad).
   */
  private collapseSpills(st: SheetState, op: StructOp): void {
    if (!st.spillPossible) return;
    const isCol = op.axis === 'col';
    const host = st.info.name;
    for (const row of st.model.rows) {
      for (const cell of row.cells) {
        const x = cell.x;
        if (!x || x.indexOf('t="array"') < 0 || !/\scm="/.test(cell.a)) continue;
        const f = parseF(x);
        const ref = f ? getAttr(f.attrs, 'ref') : undefined;
        const a0 = ref ? parseAreaText(ref) : null;
        if (!f || !a0) continue;
        const a = norm(adjustPlainArea(a0, this.entries, host, cell.ep) ?? a0);
        if (a.r1 === a.r2 && a.c1 === a.c2) continue;
        const lo = isCol ? a.c1 : a.r1, hi = isCol ? a.c2 : a.r2;
        let cut: boolean;
        if (op.type === 'insert') cut = op.at > lo && op.at <= hi;
        else if (op.type === 'delete') cut = !(op.at + op.count - 1 < lo || op.at > hi);
        else {
          const s = op.from, e = op.from + op.count - 1;
          const inter = !(e < lo || s > hi), covers = s <= lo && e >= hi;
          cut = (inter && !covers) || (!inter && op.to > lo && op.to <= hi);
        }
        if (!cut) continue;
        this.collapseSpill(st, row.r, cell, f, a);
      }
    }
  }

  /**
   * Collapse the spill of the dynamic-array master `cell` (at `rowNo`): the master formula is brought up to date, its ref
   * becomes the master cell alone and the cached spilled values of the rest of the range are cleared (Excel re-spills on
   * load, fullCalcOnLoad; a spill that is blocked by a value there becomes #SPILL! like in Excel).
   */
  private collapseSpill(st: SheetState, rowNo: number, cell: CellM, f: NonNullable<ReturnType<typeof parseF>>, a: Area): void {
    const x = cell.x!;
    const r = f.text ? adjustFormulaOps(f.text, this.entries, st.info.name, 'cell', cell.ep) : null;
    const text = r && r.changed ? r.text : f.text;
    const attrs = setAttr(f.attrs, 'ref', cellRef(cell.c, rowNo));
    cell.x = x.slice(0, f.start) + `<f${attrs}>${esc(text)}</f>` + x.slice(f.end);
    cell.ep = this.entries.length;
    for (let rr = a.r1; rr <= a.r2; rr++) {
      for (let cc = a.c1; cc <= a.c2; cc++) {
        if (rr === rowNo && cc === cell.c) continue;
        const sc = st.model.cell(rr, cc);
        if (!sc || !sc.x || sc.x.indexOf('<f') >= 0) continue;
        sc.x = null;
        sc.a = setAttr(sc.a, 't', null);
      }
    }
  }

  /**
   * The dynamic-array masters of a sheet whose CURRENT spill range covers more than the master cell (cell + row + range).
   * A write into such a range (not at the master) collapses that spill (see writeCells).
   */
  private spillMasters(st: SheetState): Array<{ row: number; cell: CellM; area: Area }> {
    const out: Array<{ row: number; cell: CellM; area: Area }> = [];
    for (const row of st.model.rows) {
      for (const cell of row.cells) {
        if (!cell.x || cell.x.indexOf('t="array"') < 0 || !/\scm="/.test(cell.a)) continue;
        const a = this.arrayOfMaster(st, cell);
        if (a) out.push({ row: row.r, cell, area: a });
      }
    }
    return out;
  }

  /** Refuse what Excel refuses (Japanese messages). */
  private check(st: SheetState, op: StructOp): void {
    const isCol = op.axis === 'col';
    const span = (a: Area): [number, number] => (isCol ? [Math.min(a.c1, a.c2), Math.max(a.c1, a.c2)] : [Math.min(a.r1, a.r2), Math.max(a.r1, a.r2)]);
    // content pushed off the sheet
    if (op.type === 'insert') {
      const max = isCol ? MAX_COL : MAX_ROW;
      const limit = max - op.count;
      for (const row of st.model.rows) {
        if (isCol) {
          const last = row.cells[row.cells.length - 1];
          if (last && last.c > limit && last.x) throw new PatchError(MSG.pushOff);
        } else if (row.r > limit && row.cells.some((c) => c.x)) throw new PatchError(MSG.pushOff);
      }
    }
    // merged cells: a move may not split a merged area
    if (op.type === 'move') {
      const s = op.from, e = op.from + op.count - 1;
      for (const m of this.merges(st)) {
        const [lo, hi] = span(m);
        const inter = !(e < lo || s > hi);
        const contains = s <= lo && e >= hi;
        if (inter && !contains) throw new PatchError(MSG.mergeSplit);
      }
    }
    // legacy array formulas
    for (const a of this.arrays(st)) {
      const [lo, hi] = span(a);
      if (op.type === 'insert') { if (op.at > lo && op.at <= hi) throw new PatchError(MSG.array); continue; }
      const s = op.type === 'delete' ? op.at : op.from, e = s + op.count - 1;
      const inter = !(e < lo || s > hi);
      const contains = s <= lo && e >= hi;
      if (inter && !contains) throw new PatchError(MSG.array);
      if (op.type === 'move' && !inter && op.to > lo && op.to <= hi) throw new PatchError(MSG.array);
    }
    // tables
    for (const t of st.tables) {
      if (t.removed) continue;
      const [lo, hi] = span(t.ref);
      if (op.type === 'delete' && !isCol && t.header) {
        const hr = Math.min(t.ref.r1, t.ref.r2);
        if (op.at <= hr && op.at + op.count - 1 >= hr && !(op.at <= lo && op.at + op.count - 1 >= hi)) throw new PatchError(MSG.tableHeader);
      }
      if (op.type !== 'move') continue;
      const s = op.from, e = op.from + op.count - 1;
      // same rules as the editor (doc/struct.ts): a table moves whole, or its columns are reordered inside it.
      // Columns: excelRef.tableMoveKind (verified in Excel 365) — a block from inside the table may land anywhere in
      // [first column, last column + 1]: right after the last column is still a reorder (the block becomes the
      // table's last column(s)).
      if (isCol) {
        if (tableMoveKind(lo, hi, op) !== 'refused') continue;
        throw new PatchError(MSG.tableMove);
      }
      const covers = s <= lo && e >= hi;
      const inside = s >= lo && e <= hi;
      const disjoint = e < lo || s > hi;
      if (covers) continue;
      // rows: only DATA rows are reordered, and only between the header row and the end: a data row dropped at / above the
      // header row would leave the table with a data row in its header (Excel repairs such a table; found by the op fuzz)
      const dataFirst = lo + t.header, dataLast = hi - t.totals;
      if (inside && s >= dataFirst && e <= dataLast && op.to >= dataFirst && op.to <= hi) continue;
      if (disjoint && !(op.to > lo && op.to <= hi)) continue;
      throw new PatchError(MSG.tableMove);
    }
  }

  merges(st: SheetState): Area[] {
    const out: Area[] = [];
    for (const m of st.model.tail.matchAll(/<mergeCell\b[^>]*>/g)) {
      const a = parseAreaText(getAttr(m[0], 'ref') ?? '');
      if (a) out.push(norm(a));
    }
    return out;
  }

  private arrays(st: SheetState): Area[] {
    if (st.arrays) return st.arrays;
    const out: Area[] = [];
    for (const row of st.model.rows) {
      for (const c of row.cells) {
        if (!c.x || c.x.indexOf('t="array"') < 0 || /\scm="/.test(c.a)) continue;
        const f = parseF(c.x);
        const ref = f ? getAttr(f.attrs, 'ref') : undefined;
        const a0 = ref ? parseAreaText(ref) : null;
        // the ref attribute is still in the coordinates of the cell's last write: bring it up to date
        const a = a0 ? adjustPlainArea(a0, this.entries, st.info.name, c.ep) : null;
        if (a && a.kind === 'area' && !(a.c1 === a.c2 && a.r1 === a.r2)) out.push(norm(a));
      }
    }
    st.arrays = out;
    return out;
  }

  /** Remap cell positions (+ Excel "format same as left/above" for inserts). */
  private moveCells(st: SheetState, op: StructOp, inherit: boolean, styles: Styles | null): void {
    const m = st.model;
    m.dirty = true;
    if (op.axis === 'col') {
      for (const row of m.rows) row.cells = mapRowCells(row, op, inherit ? styles : null);
      return;
    }
    // rows
    const rows = m.rows;
    if (op.type === 'insert') {
      const above = inherit && op.at > 1 ? m.row(op.at - 1) : undefined;
      const below = above ? m.row(op.at) : undefined;
      for (const r of rows) if (r.r >= op.at) r.r += op.count;
      let next = rows.filter((r) => r.r <= MAX_ROW);
      if (above) {
        const add: RowM[] = [];
        const attrs = inheritRowAttrs(above.a);
        const cells: CellM[] = [];
        for (const c of above.cells) {
          const s = cellStyle(c);
          if (!s) continue;
          const b = below ? findCell(below, c.c) : undefined;
          const ns = styles ? styles.insertedStyle(s, b ? cellStyle(b) : 0) : s;
          if (ns) cells.push({ c: c.c, a: ` s="${ns}"`, x: null, ep: 0 });
        }
        for (let k = 0; k < op.count; k++) {
          if (!attrs && !cells.length) break;
          add.push({ r: op.at + k, a: attrs, cells: cells.map((c) => ({ ...c })) });
        }
        if (add.length) next = next.concat(add).sort((a, b) => a.r - b.r);
      }
      m.rows = next;
    } else if (op.type === 'delete') {
      const s = op.at, e = op.at + op.count - 1;
      m.rows = rows.filter((r) => r.r < s || r.r > e);
      for (const r of m.rows) if (r.r > e) r.r -= op.count;
    } else {
      const p = movePerm(op.from, op.count, op.to);
      for (const r of rows) r.r = p(r.r);
      rows.sort((a, b) => a.r - b.r);
    }
    m.invalidateIndex();
  }

  private applyHead(st: SheetState, op: StructOp): void {
    st.model.head = adjustSheetViews(st.model.head, op);
  }

  /** Per-op rewrite of the sheet XML after <sheetData>. */
  private applyTail(st: SheetState, op: StructOp, entry: OpEntry, inherit: boolean, afEntry: OpEntry): void {
    let tail = st.model.tail;
    const plain = (a: Area) => adjustArea(a, op);
    const ext = (a: Area) => adjustAreaEntry(a, afEntry, 'cell');
    // merged cells
    tail = rewriteCounted(tail, 'mergeCells', 'mergeCell', (el) => {
      const a = parseAreaText(getAttr(el, 'ref') ?? '');
      if (!a) return el;
      if (op.type === 'move') {
        const [lo, hi] = op.axis === 'col' ? [Math.min(a.c1, a.c2), Math.max(a.c1, a.c2)] : [Math.min(a.r1, a.r2), Math.max(a.r1, a.r2)];
        const s = op.from, e = op.from + op.count - 1;
        const disjoint = e < lo || s > hi;
        if (disjoint && op.to > lo && op.to <= hi) return ''; // Excel silently un-merges (R2b R12)
      }
      const na = plain(a);
      if (!na || (na.c1 === na.c2 && na.r1 === na.r2)) return '';
      return na === a ? el : setAttr(el, 'ref', renderSqref([na]));
    });
    // CF / DV (anchored rules)
    const rr = adjustRulesInTail(tail, op, st.info.name, { leak: inherit, ext: entry.ext });
    tail = rr.tail;
    if (rr.split) st.splitCf = true;
    // hyperlinks
    tail = rewriteContainer(tail, 'hyperlinks', 'hyperlink', (el) => {
      const ref = getAttr(el, 'ref');
      const areas = ref ? parseSqref(ref) : [];
      if (!areas.length) return el;
      const na = areas.map(plain).filter((x): x is Area => !!x);
      if (!na.length) return '';
      const txt = renderSqref(na);
      return txt === ref ? el : setAttr(el, 'ref', txt);
    });
    // sheet autoFilter (+ filterColumn ids, sortState)
    tail = mapFirst(tail, 'autoFilter', (el) => adjustAutoFilter(el, op, ext));
    // a sheet-level sortState outside the autoFilter (the one inside is handled with the filter)
    const afEl = findFirst(tail, 'autoFilter');
    const ss = findFirst(tail, 'sortState');
    if (ss && !(afEl && ss.start > afEl.start && ss.end <= afEl.end)) tail = tail.slice(0, ss.start) + adjustSortState(ss.text, plain) + tail.slice(ss.end);
    // protected ranges / ignored errors
    tail = rewriteContainer(tail, 'protectedRanges', 'protectedRange', (el) => sqrefAttr(el, plain));
    tail = rewriteContainer(tail, 'ignoredErrors', 'ignoredError', (el) => sqrefAttr(el, plain));
    // page breaks
    if (op.axis === 'row') tail = adjustBreaks(tail, 'rowBreaks', op);
    else tail = adjustBreaks(tail, 'colBreaks', op);
    // sparklines (location cells)
    tail = adjustSparklineLocations(tail, op);
    // control / OLE anchors
    tail = adjustSheetObjectAnchors(tail, op);
    tail = removeEmptyExt(tail);
    st.model.tail = tail;
  }

  /** Tables of the op's sheet. */
  private async applyTables(st: SheetState, op: StructOp, entry: OpEntry): Promise<void> {
    const isCol = op.axis === 'col';
    for (const t of st.tables) {
      if (t.removed) continue;
      const a = t.ref;
      const [lo, hi] = isCol ? [a.c1, a.c2] : [a.r1, a.r2];
      let na: Area | null;
      let cols = tableColumns(t.xml);
      let newColNames: Array<{ pos: number; name: string }> = [];
      /** a row delete takes the totals row (not the whole table): like Excel the table loses its totals row and definitions */
      let dropTotals = false;
      if (op.type === 'move') {
        const s = op.from, e = s + op.count - 1;
        const reorder = isCol ? tableMoveKind(lo, hi, op) === 'reorder' : s >= lo && e <= hi && op.to >= lo && op.to <= hi && !(s <= lo && e >= hi);
        if (reorder) {
          // reorder inside the table (columns only; rows just move)
          na = a;
          if (isCol) {
            const p = movePerm(op.from, op.count, op.to);
            const next: string[] = new Array(cols.length);
            cols.forEach((c, i) => { next[p(lo + i) - lo] = c; });
            if (next.some((x) => x === undefined)) na = a; else cols = next;
          }
        } else na = adjustArea(a, op);
      } else if (op.type === 'insert' && op.at > lo && op.at <= hi) {
        na = adjustAreaEntry(a, entry, 'cell');
        if (isCol) {
          const idx = op.at - lo;
          const all = cols.map(colName);
          const names = excelNewColumnNames(all, all[idx - 1] ?? null, op.count);
          const newEls = names.map((n) => `<tableColumn id="__NEW__" name="${escAttr(escTableName(n))}"/>`);
          cols = cols.slice(0, idx).concat(newEls, cols.slice(idx));
          newColNames = names.map((n, k) => ({ pos: op.at + k, name: n }));
        }
      } else if (op.type === 'delete') {
        const s = op.at, e = op.at + op.count - 1;
        if (s <= lo && e >= hi) { await this.removeTable(st, t); continue; }
        na = adjustArea(a, op);
        if (isCol && na) {
          cols = cols.filter((_c, i) => lo + i < s || lo + i > e);
        }
        if (!isCol && na) {
          // the totals row (last row of the table) goes: totalsRowCount, the totals functions / labels / formulas of every
          // column and the filter range follow (measured: golden h2-del-totals-row, h2-del-last-data-and-totals; without
          // this the shrunk ref would make the LAST DATA row the totals row and Excel stops responding on the file)
          dropTotals = t.totals > 0 && a.r2 >= s && a.r2 <= e;
          // keep at least one data row
          const minRows = t.header + (dropTotals ? 0 : t.totals) + 1;
          if (na.r2 - na.r1 + 1 < minRows) na = { ...na, r2: na.r1 + minRows - 1 };
        }
      } else {
        na = adjustAreaEntry(a, entry, 'cell');
      }
      if (!na) { await this.removeTable(st, t); continue; }
      na = norm(na);
      if (dropTotals) cols = cols.map(stripTotalsDefinition);
      const sameRef = na.c1 === a.c1 && na.c2 === a.c2 && na.r1 === a.r1 && na.r2 === a.r2;
      const colsChanged = cols.join('') !== tableColumns(t.xml).join('');
      if (sameRef && !colsChanged && !dropTotals) continue;
      let x = t.xml;
      x = x.replace(/<table\b[^>]*>/, (open) => setAttr(open, 'ref', renderSqref([na!])));
      if (dropTotals) x = x.replace(/<table\b[^>]*>/, (open) => setAttr(setAttr(open, 'totalsRowCount', null), 'totalsRowShown', '0'));
      // autoFilter = table without the totals row
      const afArea: Area = { ...na, r2: na.r2 - (dropTotals ? 0 : t.totals), kind: 'area' };
      x = x.replace(/<autoFilter\b[^>]*?(\/>|>[\s\S]*?<\/autoFilter>)/, (el) => {
        let e2 = setAttr(el.replace(/^<autoFilter\b[^>]*>/, (o) => o), 'ref', renderSqref([afArea]));
        e2 = remapFilterColumns(e2, op, a, na!);
        return e2;
      });
      x = x.replace(/<sortState\b[^>]*?(\/>|>[\s\S]*?<\/sortState>)/, (el) => adjustSortState(el, (ar) => adjustArea(ar, op)));
      if (colsChanged) {
        let maxId = 0;
        for (const c of cols) { const id = Number(/\sid="(\d+)"/.exec(c)?.[1] ?? 0); if (id > maxId) maxId = id; }
        cols = cols.map((c) => (c.includes('id="__NEW__"') ? c.replace('id="__NEW__"', `id="${++maxId}"`) : c));
        x = x.replace(/<tableColumns\b[^>]*?(?:\/>|>[\s\S]*?<\/tableColumns>)/, (el) => {
          const open = /^<tableColumns\b[^>]*?(?=\/?>)/.exec(el)![0];
          return `${setAttr(open + '>', 'count', String(cols.length))}${cols.join('')}</tableColumns>`;
        });
      }
      t.xml = x;
      t.ref = na;
      if (dropTotals) t.totals = 0;
      t.dirty = true;
      // header cells of new table columns (Excel writes 列1, 列2 ...; without them Excel repairs the table)
      if (newColNames.length && t.header) {
        const sst = await this.sst();
        const hr = na.r1;
        for (const nc of newColNames) {
          const cell = st.model.ensureCell(hr, nc.pos);
          cell.a = setAttr(setAttr(cell.a, 't', 's'), 'cm', null);
          cell.x = `<v>${sst.add(nc.name)}</v>`;
        }
      }
      // rows inserted into the data area: Excel fills the calculated columns
      if (!isCol && op.type === 'insert' && na.r2 - na.r1 > a.r2 - a.r1) {
        const firstData = na.r1 + t.header;
        const lastData = na.r2 - t.totals;
        const calcCols = cols.map((c) => /<calculatedColumnFormula(?:\s[^>]*)?>([\s\S]*?)<\/calculatedColumnFormula>/.exec(c)?.[1]);
        const ep = this.entries.length + 1;
        calcCols.forEach((raw, i) => {
          if (raw == null) return;
          const cur = adjustFormulaOps(unesc(raw), this.entries.concat([entry]), st.info.name).text;
          for (let r = Math.max(op.at, firstData); r <= Math.min(op.at + op.count - 1, lastData); r++) {
            const cell = st.model.ensureCell(r, na!.c1 + i);
            const f = translateFormula(cur, r - firstData, 0);
            cell.a = setAttr(setAttr(setAttr(cell.a, 't', null), 'cm', null), 'vm', null);
            cell.x = `<f>${esc(f)}</f>`;
            cell.ep = ep;
          }
        });
      }
    }
  }

  /**
   * premerge: grow a table to the right over `names.length` columns that were just inserted right after it
   * (new tableColumns named `names`; the caller writes the same texts into the header cells).
   */
  extendTableColumns(t: TableState, names: string[]): void {
    const na: Area = { ...t.ref, c2: t.ref.c2 + names.length };
    let cols = tableColumns(t.xml);
    let maxId = 0;
    for (const c of cols) { const id = Number(/\sid="(\d+)"/.exec(c)?.[1] ?? 0); if (id > maxId) maxId = id; }
    cols = cols.concat(names.map((n) => `<tableColumn id="${++maxId}" name="${escAttr(escTableName(n))}"/>`));
    let x = t.xml.replace(/<table\b[^>]*>/, (open) => setAttr(open, 'ref', renderSqref([na])));
    const afArea: Area = { ...na, r2: na.r2 - t.totals, kind: 'area' };
    x = x.replace(/<autoFilter\b[^>]*>/, (open) => setAttr(open, 'ref', renderSqref([afArea])));
    x = x.replace(/<tableColumns\b[^>]*?(?:\/>|>[\s\S]*?<\/tableColumns>)/, (el) => {
      const open = /^<tableColumns\b[^>]*?(?=\/?>)/.exec(el)![0];
      return `${setAttr(open + '>', 'count', String(cols.length))}${cols.join('')}</tableColumns>`;
    });
    t.xml = x;
    t.ref = na;
    t.dirty = true;
  }

  private async removeTable(st: SheetState, t: TableState): Promise<void> {
    t.removed = true;
    this.pkg.remove(t.path);
    await this.pkg.removeRel(st.info.path, (r) => r.id === t.rid);
    await this.pkg.setContentTypeOverride(t.path, null);
    st.model.tail = rewriteCounted(st.model.tail, 'tableParts', 'tablePart', (el) => (getAttr(el, 'r:id') === t.rid ? '' : el));
  }

  /** Parts hanging off the op's sheet: drawing anchors, notes (+VML), threaded comments, pivot table locations. */
  private async applyRelated(st: SheetState, op: StructOp): Promise<void> {
    const rels = await this.pkg.relTargets(st.info.path);
    for (const r of rels) {
      if (REL('drawing')(r)) {
        const x = await this.pkg.read(r.path);
        if (x) this.pkg.write(r.path, adjustDrawingXml(x, op));
      } else if (REL('comments')(r)) {
        const x = await this.pkg.read(r.path);
        if (x) this.pkg.write(r.path, adjustRefElements(x, 'comment', op));
      } else if (REL('threadedComment')(r)) {
        const x = await this.pkg.read(r.path);
        if (x) this.pkg.write(r.path, adjustRefElements(x, 'threadedComment', op));
      } else if (REL('pivotTable')(r)) {
        const x = await this.pkg.read(r.path);
        if (x) {
          this.pkg.write(r.path, x.replace(/<location\b[^>]*>/, (el) => {
            const a = parseAreaText(getAttr(el, 'ref') ?? '');
            const na = a ? adjustArea(a, op) : null;
            return na && na !== a ? setAttr(el, 'ref', renderSqref([na])) : el;
          }));
        }
      }
    }
    // legacy drawing (notes / form controls): only the part referenced by <legacyDrawing>
    const ld = /<legacyDrawing\b[^>]*>/.exec(st.model.tail);
    const rid = ld ? getAttr(ld[0], 'r:id') : undefined;
    if (rid) {
      const r = rels.find((x) => x.id === rid);
      if (r) {
        const x = await this.pkg.read(r.path);
        if (x) this.pkg.write(r.path, adjustVmlXml(x, op));
      }
    }
  }

  // -------------------------------------------------------------------------------------------------
  // cell writes / widths
  // -------------------------------------------------------------------------------------------------

  /**
   * Cell writes (positions as of this point of the log). Excel rules:
   *  - 'edit' inside a legacy (CSE) array range is refused; 'toValues' over an array / spill MASTER replaces the
   *    whole array by its cached values (the other cells of the range already hold theirs);
   *  - an 'edit' over a spill master clears the old spilled values (Excel removes them; the new formula re-spills
   *    on load);
   *  - a write into a table's header row renames the table column (unique name, 列N for an empty header) and every
   *    structured reference to it;
   *  - a write into a table's totals row redefines the column's totals entry like Excel: a constant becomes the
   *    column's totalsRowLabel (text), SUBTOTAL(1xx, Table[col]) of the column itself a built-in totalsRowFunction,
   *    any other formula a custom totalsRowFormula, a cleared cell removes the definition;
   *  - a new cell takes the row / column format (like typing into an empty cell); `numFmt` changes only the
   *    number format; `dyn` stores a dynamic-array (spilling) formula; `quote` (text constants only) stores the
   *    cell with quotePrefix, any other write leaves the cell without it (see setCell).
   */
  async writeCells(sheetName: string, writes: CellWrite[], reason: 'edit' | 'toValues' = 'edit'): Promise<void> {
    const st = await this.state(sheetName);
    st.model.assertWritable(st.info.name);
    const ep = this.entries.length;
    for (const w of writes) {
      if (!w || !(Number.isInteger(w.row) && Number.isInteger(w.col) && w.row >= 1 && w.row <= MAX_ROW && w.col >= 1 && w.col <= MAX_COL)) throw new PatchError('セルの位置が正しくありません。');
      checkWriteValue(w, sheetName);
    }
    writes = this.neutralizeUnsafe(writes);
    // arrays converted to values by this batch
    const converting: Area[] = [];
    if (reason === 'toValues') {
      for (const w of writes) {
        const a = this.arrayOfMaster(st, st.model.cell(w.row, w.col));
        if (a) converting.push(a);
      }
    }
    const inside = (a: Area, w: CellWrite) => w.row >= a.r1 && w.row <= a.r2 && w.col >= a.c1 && w.col <= a.c2;
    let spills: Array<{ row: number; cell: CellM; area: Area }> | null = null;
    for (const w of writes) {
      if (!converting.some((a) => inside(a, w))) {
        for (const a of this.arrays(st)) if (inside(a, w)) throw new PatchError(MSG.array);
        // A value / formula written INTO the spill range of a dynamic array (not at its master): the spill is blocked — Excel
        // shows #SPILL! and re-spills when the cell is free. The file must not keep an array range over a cell that holds
        // something else (Excel repairs it: "cell information" removed), so the spill collapses like when an op cuts it.
        if (w.cell) {
          spills ??= this.spillMasters(st);
          for (let i = 0; i < spills.length; i++) {
            const sm = spills[i];
            if (!inside(sm.area, w) || (sm.row === w.row && sm.cell.c === w.col)) continue;
            const f = parseF(sm.cell.x);
            if (f) this.collapseSpill(st, sm.row, sm.cell, f, sm.area);
            spills.splice(i--, 1);
          }
        }
      }
      const hdr = this.tableHeaderAt(st, w.row, w.col);
      if (hdr) { await this.writeTableHeader(st, hdr, w, ep); continue; }
      const tot = this.tableTotalsAt(st, w.row, w.col);
      if (tot && await this.writeTotalsCell(st, tot, w, ep)) continue;
      const existing = st.model.cell(w.row, w.col);
      if (!existing && !w.cell) continue; // clearing a cell that does not exist
      const cell = existing ?? this.newCell(st, w.row, w.col);
      const old = this.arrayOfMaster(st, existing);
      if (old) {
        if (reason === 'toValues') await this.freezeArrayRange(st, old, cell);
        else if (/\scm="/.test(cell.a)) clearSpill(st.model, old, cell);
        if (spills) spills = spills.filter((s) => s.cell !== cell); // this master is replaced: its spill is gone
      }
      await this.setCell(cell, w.cell, ep, w.row);
      if (w.cell?.dyn) st.spillPossible = true;
      const v = w.cell;
      if (v && typeof v.numFmt === 'string' && v.numFmt.trim() && !(v.f != null && v.f !== '') && (v.t === 'n' || v.t === 'd') && typeof v.v === 'number' && Number.isFinite(v.v)) {
        await this.autoWiden(st, w.col, cell, v.v, v.numFmt.trim());
      }
    }
    st.arrays = null;
    st.model.dirty = true;
    st.changed = true;
  }

  /**
   * Text that came from an untrusted file must never become a live call-out when it is written to a cell (the Excel runner does the
   * same, XlRunner.Cells.Build, and warns UNSAFE_FORMULA): a formula that `isUnsafeFormula` (DDE `|`, [book.xlsx] references, UNC / URL
   * paths, WEBSERVICE / IMAGE / RTD / CALL / REGISTER / EXEC) is written as TEXT with the quote prefix — the formula as the user sees it,
   * "=" included, no <f> — and a text constant that Excel would take for a formula ("= + - @" first, NFKC) and that is just as unsafe
   * gets the quote prefix, so editing the cell later cannot turn it into one. Ordinary formulas and HYPERLINK are left alone.
   */
  private neutralizeUnsafe(writes: CellWrite[]): CellWrite[] {
    let out: CellWrite[] | null = null;
    for (let i = 0; i < writes.length; i++) {
      const w = writes[i];
      const v = w.cell;
      if (!v) continue;
      let repl: NonNullable<CellWrite['cell']>;
      if (typeof v.f === 'string' && v.f !== '') {
        const ui = toUiForm(v.f.replace(/^=/, '')); // the file form has no '=' (the engine accepts both spellings)
        const text = ui.startsWith('=') ? ui : '=' + ui;
        if (!isUnsafeFormula(text)) continue;
        repl = { t: 's', v: text, quote: true, ...(v.numFmt ? { numFmt: v.numFmt } : {}) };
      } else if (v.t === 's' && typeof v.v === 'string' && v.quote !== true && /^[=+\-@]/.test(v.v.slice(0, 1).normalize('NFKC')) && isUnsafeFormula(v.v)) {
        repl = { ...v, quote: true };
      } else continue;
      out ??= writes.slice();
      out[i] = { ...w, cell: repl };
      this.neutralized++;
      if (this.neutralized <= 20) {
        this.warnings.push({ code: 'UNSAFE_FORMULA', message: `セル ${cellRef(w.col, w.row)} の内容は、外部のプログラム・ファイル・Web を呼び出す数式の形（DDE・外部ブックの参照・WEBSERVICE など）のため、数式にせず文字列として書き込みました` });
      } else if (this.neutralized === 21) {
        this.warnings.push({ code: 'UNSAFE_FORMULA', message: '同様に文字列として書き込んだセルが他にもあります' });
      }
    }
    return out ?? writes;
  }

  /** A cell for a write into an empty position: takes the row format (customFormat) or else the column format. */
  private newCell(st: SheetState, r: number, c: number): CellM {
    const s = this.defaultStyle(st, r, c);
    const cell = st.model.ensureCell(r, c);
    if (s > 0) cell.a = ` s="${s}"`;
    return cell;
  }

  /** The xf an absent cell has: its row format (customFormat) or else its column format, 0 = none. */
  private defaultStyle(st: SheetState, r: number, c: number): number {
    const row = st.model.row(r);
    let s = row && /\scustomFormat="(1|true)"/.test(row.a) ? Number(getAttr(row.a, 's') ?? 0) : 0;
    if (!s) s = Number(getAttr(st.cols.get(c) ?? '', 'style') ?? 0);
    return s > 0 ? s : 0;
  }

  /** Current (multi-cell) range of the array / spill formula whose master is `cell`, else null. */
  private arrayOfMaster(st: SheetState, cell: CellM | undefined): Area | null {
    if (!cell || !cell.x || cell.x.indexOf('t="array"') < 0) return null;
    const f = parseF(cell.x);
    if (!f || getAttr(f.attrs, 't') !== 'array') return null;
    const ref = getAttr(f.attrs, 'ref');
    const a0 = ref ? parseAreaText(ref) : null;
    const a = a0 ? adjustPlainArea(a0, this.entries, st.info.name, cell.ep) : null;
    if (!a || a.kind !== 'area' || (a.c1 === a.c2 && a.r1 === a.r2)) return null;
    return norm(a);
  }

  /** toValues over an array master: the other cells of the range keep their cached values as constants. */
  private async freezeArrayRange(st: SheetState, a: Area, master: CellM): Promise<void> {
    let sst: Sst | null = null;
    for (let r = a.r1; r <= a.r2; r++) {
      const row = st.model.row(r);
      if (!row) continue;
      for (const sc of row.cells) {
        if (sc.c < a.c1 || sc.c > a.c2 || sc === master || !sc.x || sc.x.indexOf('<f') >= 0) continue;
        if (getAttr(sc.a, 't') !== 'str') continue;
        // "formula string" results become ordinary strings
        const m = /<v>([\s\S]*?)<\/v>/.exec(sc.x);
        sst ??= await this.sst();
        sc.a = setAttr(sc.a, 't', m ? 's' : null);
        sc.x = m ? `<v>${sst.add(unesc(m[1]))}</v>` : null;
      }
    }
  }

  /**
   * Write a value / formula into cell (`row`, cell.c) keeping its style (+ numFmt / quotePrefix / wrap text /
   * dynamic-array metadata).
   *
   * quotePrefix is honoured EXACTLY (types.ts CellWrite.cell.quote): `quote === true` on a text constant gives the
   * cell an xf with quotePrefix="1" (derived from its own xf, so a number format such as '@' set by the same write
   * stays); every other write adds no prefix. DELIBERATE DIFFERENCE TO THE EXCEL ENGINE: a prefix the cell already had is
   * CLEARED here by a non-quote write (what typing in the Excel UI does: a new entry resets the prefix character),
   * while Excel through COM keeps the flag when `Value2` / `Formula` / `ClearContents` write over the cell (only
   * ClearFormats removes it; measured, golden q-overwrite / q-text, `QUOTE_COM`). The flag is invisible for numbers,
   * booleans, formulas and empty cells (Excel shows a prefix only on text), so the two engines differ in that one
   * invisible xf flag for those writes; for a text written over a prefixed cell the COM result keeps the apostrophe
   * in the formula bar, this engine does not (the contract says "no prefix", types.ts). A cleared cell keeps its style
   * (prefix included) in both engines.
   *
   * Wrap text: Excel turns WrapText on by itself when a text CONSTANT containing a line feed (LF, also as CR LF) is
   * written (measured: `Value2 = "a" & vbLf & "b"` and "a" & vbCrLf & "b", golden q-wrap) — only when the cell's xf has
   * no wrap yet; formulas and numbers never do. A lone CR does NOT switch it on in Excel (golden q-wrap, cell C7), so
   * neither does this engine.
   */
  async setCell(cell: CellM, v: CellWrite['cell'], ep: number, row: number): Promise<void> {
    const needSst = !!v && v.t === 's' && !(v.f != null && v.f !== '') && v.v != null;
    const hasF = !!v && v.f != null && v.f !== '';
    const dyn = hasF && v!.dyn ? { cm: await this.dynamicArrayCm(), ref: cellRef(cell.c, row) } : undefined;
    const wantNumFmt = !!v && typeof v.numFmt === 'string' && !!v.numFmt;
    const wantQuote = !!v && v.quote === true && v.t === 's' && !hasF && v.v != null;
    const wantWrap = needSst && String(v!.v).indexOf('\n') >= 0;
    // a cell without a style attribute is xf 0: nothing to derive unless the write asks for a number format / prefix /
    // wrap (a cleared cell keeps its style as is: types.ts CellWrite, `null clears the value / formula (style kept)`)
    if (v && (wantNumFmt || wantQuote || wantWrap || /\ss\s*=/.test(cell.a))) {
      const styles = await this.styles();
      const s = cellStyle(cell);
      let ns = s;
      if (wantNumFmt) ns = styles.withNumFmt(ns, v!.numFmt!);
      ns = styles.withQuotePrefix(ns, wantQuote);
      if (wantWrap) ns = styles.withWrapText(ns);
      if (ns !== s) cell.a = setAttr(cell.a, 's', ns ? String(ns) : null);
    }
    // The doc / the reader hold every date serial in the 1900 system (a 1904 book is converted on read, styles.ts dateModeOf):
    // a number written into a 1904 book goes back to its system by the same rule (calendar dates always, times from 1 day on)
    let vv = v;
    if (v && (v.t === 'd' || v.t === 'n') && typeof v.v === 'number' && Number.isFinite(v.v) && this.date1904) {
      const mode = (await this.styles()).dateMode(cellStyle(cell));
      const n = to1904Serial(v.v, mode);
      if (n !== v.v) vv = { ...v, v: n };
    }
    this.setCellSync(cell, vv, ep, needSst ? await this.sst() : null, dyn);
  }

  /**
   * setCell with the shared strings already loaded (`sst` may be null when no plain string is written).
   * `dyn`: store the formula as a dynamic array (cell-metadata index + the master's own A1 ref).
   */
  setCellSync(cell: CellM, v: CellWrite['cell'], ep: number, sst: Sst | null, dyn?: { cm: number; ref: string }): void {
    let a = cell.a;
    a = setAttr(setAttr(setAttr(setAttr(a, 't', null), 'cm', null), 'vm', null), 'ph', null);
    cell.ep = ep;
    if (!v) { cell.a = a.replace(/\s+$/, ''); cell.x = null; return; }
    const f = v.f != null && v.f !== '' ? v.f.replace(/^=/, '') : null;
    const val = v.v;
    let t: string | null = null;
    let vx = '';
    if (val !== null && val !== undefined && v.t !== 'z') {
      switch (v.t) {
        case 's': {
          const s = String(val).slice(0, 32767);
          if (f) { t = 'str'; vx = `<v>${escText(s)}</v>`; }
          else if (sst) { t = 's'; vx = `<v>${sst.add(s)}</v>`; }
          else { t = 'inlineStr'; vx = `<is><t xml:space="preserve">${escText(s)}</t></is>`; }
          break;
        }
        case 'b': t = 'b'; vx = `<v>${val === true || val === 1 || val === 'TRUE' ? 1 : 0}</v>`; break;
        case 'e': t = 'e'; vx = `<v>${esc(String(val))}</v>`; break;
        default: {
          const n = typeof val === 'number' ? val : Number(val);
          if (Number.isFinite(n)) vx = `<v>${numText(n)}</v>`;
          else { t = 'e'; vx = '<v>#NUM!</v>'; }
        }
      }
    }
    if (t) a = setAttr(a, 't', t);
    let fx = f ? `<f>${esc(f)}</f>` : '';
    if (f && dyn) {
      // dynamic array: <c cm="n"><f t="array" ref="A1">…</f> — Excel computes the spill range on load
      a = setAttr(a, 'cm', String(dyn.cm));
      fx = `<f t="array" ref="${dyn.ref}">${esc(f)}</f>`;
    }
    cell.a = a.replace(/\s+$/, '');
    cell.x = fx + vx || null;
  }

  /**
   * CellWrite.numFmt on a number: like Excel's Range.NumberFormat (the Excel runner) / typing a date, a column of
   * standard width (no customWidth, not hidden) whose width no longer fits the formatted text (it would show #####)
   * is widened to fit it (<col width bestFit customWidth>, see ./autowidth.ts). User-set widths are kept.
   */
  private async autoWiden(st: SheetState, col: number, cell: CellM, value: number, code: string): Promise<void> {
    const a = st.cols.get(col) ?? '';
    // a width set by the user (customWidth without bestFit) or a hidden column is never touched; an auto-fitted
    // column (bestFit) keeps growing
    if ((/\scustomWidth="(1|true)"/.test(a) && !/\sbestFit="(1|true)"/.test(a)) || /\shidden="(1|true)"/.test(a)) return;
    const text = formatValue(value, code).text;
    if (!text || /^#+$/.test(text) || text.includes('\n')) return;
    const styles = await this.styles();
    const def = styles.fontOf(0);
    const mdw = fontMdw(def.name, def.size);
    const f = styles.fontOf(cellStyle(cell));
    const need = numberTextPx(text, fontMdw(f.name, f.size));
    const cur = Number(getAttr(a, 'width'));
    let px: number;
    if (a && Number.isFinite(cur) && cur > 0) px = excelWidthToPx(cur, mdw);
    else {
      const fp = /<sheetFormatPr\b[^>]*>/.exec(st.model.head)?.[0] ?? '';
      const dw = Number(getAttr(fp, 'defaultColWidth'));
      const base = Number(getAttr(fp, 'baseColWidth') ?? 8);
      px = Number.isFinite(dw) && dw > 0 ? excelWidthToPx(dw, mdw) : defaultColumnPx(mdw, Number.isFinite(base) && base >= 0 ? base : 8);
    }
    if (need <= px - 5) return;
    const width = Math.min(255, Math.round(((need + 5) / mdw) * 1e6) / 1e6);
    st.cols.set(col, setAttr(setAttr(setAttr(a, 'width', String(width)), 'bestFit', '1'), 'customWidth', '1'));
    st.changed = true;
  }

  async setColWidth(sheetName: string, col: number, width: number): Promise<void> {
    if (!(col >= 1 && col <= MAX_COL) || !(width >= 0 && width <= 255)) throw new PatchError('列幅の指定が正しくありません。');
    const st = await this.state(sheetName);
    st.model.assertWritable(st.info.name);
    let a = st.cols.get(col) ?? '';
    // a width the user sets replaces an auto-fit one: Excel drops bestFit (golden w-after-auto), so a later wider date
    // no longer widens the column (autoWiden keeps only user widths that are not bestFit)
    a = setAttr(setAttr(setAttr(a, 'width', String(Math.round(width * 1e6) / 1e6)), 'customWidth', '1'), 'bestFit', null);
    if (width > 0 && /\shidden="(1|true)"/.test(a)) a = setAttr(a, 'hidden', null);
    st.cols.set(col, a);
    st.changed = true;
  }

  /** cm="…" index of the dynamic-array cell metadata (xl/metadata.xml created / merged on first use). */
  async dynamicArrayCm(): Promise<number> {
    if (this.dynMeta) return this.dynMeta.cm;
    const rel = (await this.pkg.relTargets(this.wbPath)).find((r) => isRel(r.type, 'sheetMetadata'));
    const old = rel ? await this.pkg.read(rel.path) : null;
    const res = ensureDynamicArrayMetadata(old);
    let path = rel?.path;
    if (!path || old == null) {
      path = path ?? dirOf(this.wbPath) + '/metadata.xml';
      if (!rel) {
        let p = path;
        for (let n = 2; this.pkg.has(p); n++) p = dirOf(this.wbPath) + `/metadata${n}.xml`;
        path = p;
        await this.pkg.addRel(this.wbPath, METADATA_REL, path);
      }
      await this.pkg.setContentTypeOverride(path, METADATA_CT);
    }
    this.dynMeta = { cm: res.cm, path, xml: res.xml, changed: res.xml !== old };
    return res.cm;
  }

  // -------------------------------------------------------------------------------------------------
  // unmerge (LoggedOp 'unmerge')
  // -------------------------------------------------------------------------------------------------

  /**
   * Remove merged areas (A1 refs, current coordinates).
   *  - ALWAYS (types.ts 'unmerge'): the top-left cell's look spreads over the rest of each former area, so a title band
   *    whose style sits on the top-left cell only (openpyxl-style files) stays visible: its FILL goes to every other
   *    cell, each border side it has goes to the cells that lie on that edge of the area (top: first row, bottom:
   *    last row, left: first column, right: last column — for a one-row band: top / bottom on every cell, left on
   *    the first, right on the last cell). A side the top-left cell does not have leaves the other cell's own side
   *    alone; other style parts (font, number format, alignment ...) of the cells are kept; a missing cell is created
   *    (empty, with the derived xf) when its style changes.
   *  - centerAcross: like Range.UnMerge + HorizontalAlignment = xlCenterAcrossSelection, every cell of each former
   *    area gets its own xf with horizontal="centerContinuous" (other alignment settings kept); missing cells are
   *    created with that style so the alignment covers the area.
   */
  async unmerge(sheetName: string, refs: string[], centerAcross: boolean): Promise<void> {
    const st = await this.state(sheetName);
    st.model.assertWritable(st.info.name);
    const areas: Area[] = [];
    for (const ref of refs ?? []) {
      const a = typeof ref === 'string' ? parseAreaText(ref.replace(/\$/g, '').replace(/^.*!/, '')) : null;
      if (!a || (a.kind !== 'area' && a.kind !== 'cell')) throw new PatchError('結合を解除するセル範囲の指定が正しくありません。');
      areas.push(norm(a));
    }
    const same = (x: Area, y: Area) => x.r1 === y.r1 && x.r2 === y.r2 && x.c1 === y.c1 && x.c2 === y.c2;
    const removed: Area[] = [];
    st.model.tail = rewriteCounted(st.model.tail, 'mergeCells', 'mergeCell', (el) => {
      const a = parseAreaText(getAttr(el, 'ref') ?? '');
      if (!a) return el;
      const n = norm(a);
      if (!areas.some((x) => same(x, n))) return el;
      removed.push(n);
      return '';
    });
    if (removed.length) {
      const styles = await this.styles();
      for (const a of removed) {
        const n = (a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1);
        const tl = st.model.cell(a.r1, a.c1);
        const look = styles.mergeLook(tl ? cellStyle(tl) : this.defaultStyle(st, a.r1, a.c1));
        if (!look && !centerAcross) continue;
        for (let r = a.r1; r <= a.r2; r++) {
          for (let c = a.c1; c <= a.c2; c++) {
            const existing = st.model.cell(r, c);
            // a pathological merge over whole rows / columns: only the cells that exist
            if (!existing && n > 100000) continue;
            const s = existing ? cellStyle(existing) : this.defaultStyle(st, r, c);
            let ns = s;
            if (look && !(r === a.r1 && c === a.c1)) {
              ns = styles.withMergeLook(ns, look, { top: r === a.r1, bottom: r === a.r2, left: c === a.c1, right: c === a.c2 });
            }
            if (centerAcross) ns = styles.withHorizontal(ns, 'centerContinuous');
            if (ns === s && !existing) continue;
            const cell = existing ?? this.newCell(st, r, c);
            if (ns !== s) cell.a = setAttr(cell.a, 's', String(ns));
          }
        }
      }
      st.model.dirty = true;
    }
    st.changed = true;
  }

  // -------------------------------------------------------------------------------------------------
  // table headers (a write into a header cell renames the table column)
  // -------------------------------------------------------------------------------------------------

  private tableHeaderAt(st: SheetState, row: number, col: number): TableState | null {
    for (const t of st.tables) {
      if (!t.removed && t.header && row === t.ref.r1 && col >= t.ref.c1 && col <= t.ref.c2) return t;
    }
    return null;
  }

  /**
   * Excel: the header cell always holds the column name as text; the name is unique inside the table
   * (「氏名」 again -> 「氏名2」), an empty header becomes 列N; the table column and every structured reference to
   * it (formulas of all sheets, calculated-column / totals formulas, defined names, CF / DV) are renamed.
   */
  private async writeTableHeader(st: SheetState, t: TableState, w: CellWrite, ep: number): Promise<void> {
    const cols = tableColumns(t.xml);
    const pos = w.col - t.ref.c1;
    if (pos < 0 || pos >= cols.length) return;
    const oldName = colName(cols[pos]);
    const others = cols.filter((_c, i) => i !== pos).map(colName);
    const name = uniqueColumnName(headerText(w.cell), others);
    const cell = st.model.cell(w.row, w.col) ?? this.newCell(st, w.row, w.col);
    await this.setCell(cell, { t: 's', v: name }, ep, w.row);
    if (name !== oldName) await this.renameTableColumn(st, t, pos, oldName, name);
  }

  // -------------------------------------------------------------------------------------------------
  // table totals row (a write into a totals cell redefines the column's totals entry)
  // -------------------------------------------------------------------------------------------------

  private tableTotalsAt(st: SheetState, row: number, col: number): TableState | null {
    for (const t of st.tables) {
      if (!t.removed && t.totals && row === t.ref.r2 && col >= t.ref.c1 && col <= t.ref.c2) return t;
    }
    return null;
  }

  /**
   * Excel (probe-totals.ps1, golden h2-totals*): the totals row is part of the table definition.
   *  - a constant (text / number / boolean / date; also `Value2 = Value2` over a formula) becomes the column's
   *    totalsRowLabel and the cell holds that text; the function / formula of the column is dropped;
   *  - SUBTOTAL(n, Table[thisColumn]) with n = 101 average, 102 countNums, 103 count, 104 max, 105 min, 107 stdDev,
   *    109 sum, 110 var is stored as the built-in totalsRowFunction (no totalsRowFormula);
   *  - any other formula: totalsRowFunction="custom" + totalsRowFormula (the typed text);
   *  - clearing the cell (or typing nothing) removes label / function / formula.
   * Returns false when the table part does not describe this column (the plain write is used).
   */
  private async writeTotalsCell(st: SheetState, t: TableState, w: CellWrite, ep: number): Promise<boolean> {
    const cols = tableColumns(t.xml);
    const pos = w.col - t.ref.c1;
    if (pos < 0 || pos >= cols.length) return false;
    const v = w.cell;
    const hasF = !!v && v.f != null && v.f !== '';
    let def: TotalsDef = {};
    let value: CellWrite['cell'] = null;
    if (v && hasF) {
      const f = v.f!.replace(/^=/, '');
      const fn = totalsFunctionOf(f, t.name, colName(cols[pos]));
      def = fn ? { fn } : { fn: 'custom', formula: f };
      value = { ...v, f, dyn: false };
    } else if (v) {
      const text = totalsLabelText(v);
      // a quote flag on a text constant is honoured here too (Excel: the cell keeps the apostrophe prefix, h2-totals-q)
      if (text !== '') { def = { label: text }; value = v.t === 's' && v.quote === true ? { t: 's', v: text, quote: true } : { t: 's', v: text }; }
    }
    const existing = st.model.cell(w.row, w.col);
    if (value || existing) {
      const cell = existing ?? this.newCell(st, w.row, w.col);
      await this.setCell(cell, value, ep, w.row);
    }
    // earlier ops must not be applied to the formula text written now
    this.syncTableFormulas(t, st.info.name);
    const next = mapTableColumnAt(t.xml, pos, (el) => withTotalsDef(el, def));
    if (next !== t.xml) { t.xml = next; t.dirty = true; }
    return true;
  }

  /** Bring the calculated-column / totals formulas of one table up to date with the ops logged so far. */
  private syncTableFormulas(t: TableState, host: string): void {
    const entries = this.entries;
    const from = this.tableEp.get(t.path) ?? 0;
    if (from < entries.length) {
      const x = t.xml.replace(/<(calculatedColumnFormula|totalsRowFormula)(\s[^>]*)?>([\s\S]*?)<\/\1>/g, (m0, tag: string, at: string | undefined, body: string) => {
        const r = adjustFormulaOps(unesc(body), entries, host, 'cell', from);
        return r.changed ? `<${tag}${at ?? ''}>${esc(r.text)}</${tag}>` : m0;
      });
      if (x !== t.xml) { t.xml = x; t.dirty = true; }
    }
    this.tableEp.set(t.path, entries.length);
  }

  private async renameTableColumn(st: SheetState, t: TableState, pos: number, oldName: string, newName: string): Promise<void> {
    // 1) the table part
    let i = 0;
    t.xml = t.xml.replace(/<tableColumns\b[^>]*>[\s\S]*?<\/tableColumns>/, (block) => block.replace(/<tableColumn\b[^>]*>/g, (open) => (i++ === pos ? setAttr(open, 'name', escTableName(newName)) : open)));
    t.dirty = true;
    const tl = t.name.toLowerCase();
    const entries = this.entries;
    const rename = (f: string, host: string | null) => renameTableColumnInFormula(f, tl, oldName, newName, host);
    // 2) every worksheet: cells, CF / DV, tables' formulas
    for (const info of this.sheets) {
      if (!info.worksheet || !info.path) continue;
      let s = this.states.get(info.path);
      if (!s) {
        const xml = await this.pkg.read(info.path);
        if (xml && mentions(xml, tl)) s = await this.state(info.name);
      }
      if (s) {
        let changed = false;
        const own = s === st;
        for (const row of s.model.rows) {
          for (const cell of row.cells) {
            if (!cell.x || cell.x.indexOf('<f') < 0 || cell.x.indexOf('[') < 0) continue;
            // bring the formula up to date first (a later rename must not meet an older delete's #REF! rule)
            this.adjustCell(cell, info.name, entries, row.r);
            const f = parseF(cell.x);
            if (!f || !f.text) continue;
            const host = own && row.r >= t.ref.r1 && row.r <= t.ref.r2 && cell.c >= t.ref.c1 && cell.c <= t.ref.c2 ? tl : null;
            const nt = rename(f.text, host);
            if (nt === f.text) continue;
            cell.x = cell.x.slice(0, f.start) + `<f${f.attrs}>${esc(nt)}</f>` + cell.x.slice(f.end);
            changed = true;
          }
        }
        const tail = mapRuleFormulas(s.model.tail, (f) => rename(f, null));
        if (tail !== s.model.tail) { s.model.tail = tail; changed = true; }
        if (changed) { s.model.dirty = true; s.changed = true; }
      }
      const tables = s ? s.tables.filter((x) => !x.removed) : await this.loadTables(info);
      for (const tt of tables) {
        const from = this.tableEp.get(tt.path) ?? 0;
        const x = tt.xml.replace(/<(calculatedColumnFormula|totalsRowFormula)(\s[^>]*)?>([\s\S]*?)<\/\1>/g, (m0, tag: string, at: string | undefined, body: string) => {
          let text = unesc(body);
          const r = adjustFormulaOps(text, entries, info.name, 'cell', from);
          if (r.changed) text = r.text;
          const nt = rename(text, tt.name.toLowerCase());
          return nt !== unesc(body) ? `<${tag}${at ?? ''}>${esc(nt)}</${tag}>` : m0;
        });
        this.tableEp.set(tt.path, entries.length);
        if (x !== tt.xml) {
          tt.xml = x;
          if (s) tt.dirty = true; else this.pkg.write(tt.path, x);
        }
      }
    }
    // 3) defined names
    const nfrom = this.namesEp;
    this.wbXml = this.wbXml.replace(/(<definedName\b[^>]*?(?<!\/)>)([\s\S]*?)(<\/definedName>)/g, (m0, a: string, body: string, c: string) => {
      let text = unesc(body);
      const r = adjustFormulaOps(text, entries, null, 'name', nfrom);
      if (r.changed) text = r.text;
      const nt = rename(text, null);
      return nt !== unesc(body) ? a + esc(nt) + c : m0;
    });
    this.namesEp = entries.length;
  }

  // -------------------------------------------------------------------------------------------------
  // finish: formulas everywhere, workbook-level parts, calcChain
  // -------------------------------------------------------------------------------------------------

  async finish(): Promise<Uint8Array> {
    const entries = this.entries;
    const opSheets = new Set(entries.map((e) => e.sheet.toLowerCase()));
    const tableNames = new Set<string>();
    for (const e of entries) if (e.tableHits) for (const k of e.tableHits.keys()) tableNames.add(k);
    // 1) every worksheet: cell formulas + formula-bearing features
    for (const info of this.sheets) {
      if (!info.worksheet || !info.path) continue;
      const st = this.states.get(info.path);
      if (st) {
        this.adjustCells(st.model, info.name, entries);
        st.model.tail = this.adjustTailFormulas(st.model.tail, info.name, entries);
        if (st.splitCf) st.model.tail = renumberPriorities(st.model.tail);
        st.model.head = st.cols.applyToHead(st.model.head);
        st.model.head = setDimension(st.model.head, st.model.usedRange());
        this.pkg.write(info.path, st.model.serialize(info.name));
        for (const t of st.tables) if (t.dirty && !t.removed) this.pkg.write(t.path, t.xml);
        continue;
      }
      if (!entries.length) continue;
      const xml = await this.pkg.read(info.path);
      if (!xml || !mayReference(xml, opSheets, tableNames)) continue;
      const model = new SheetModel(xml);
      model.unshare();
      const changedCells = this.adjustCells(model, info.name, entries);
      const tail = this.adjustTailFormulas(model.tail, info.name, entries);
      if (!changedCells && tail === model.tail) continue;
      model.tail = tail;
      this.pkg.write(info.path, model.serialize(info.name));
    }
    if (entries.length) {
      // 2) tables of every sheet: calculated column / totals formulas
      for (const info of this.sheets) {
        if (!info.worksheet || !info.path) continue;
        const st = this.states.get(info.path);
        const tables = st ? st.tables.filter((t) => !t.removed) : await this.loadTables(info);
        for (const t of tables) {
          const from = this.tableEp.get(t.path) ?? 0;
          const x = t.xml.replace(/<(calculatedColumnFormula|totalsRowFormula)(\s[^>]*)?>([\s\S]*?)<\/\1>/g, (m0, tag: string, at: string | undefined, body: string) => {
            const r = adjustFormulaOps(unesc(body), entries, info.name, 'cell', from);
            return r.changed ? `<${tag}${at ?? ''}>${esc(r.text)}</${tag}>` : m0;
          });
          if (x !== t.xml) { t.xml = x; this.pkg.write(t.path, x); }
        }
        // form controls (ctrlProps) and VML control formulas
        for (const r of (await this.pkg.relTargets(info.path))) {
          if (REL('ctrlProp')(r)) {
            const x = await this.pkg.read(r.path);
            if (!x) continue;
            const y = x.replace(/\s(fmlaLink|fmlaRange|fmlaTxbx)="([^"]*)"/g, (m0, n: string, v: string) => {
              const res = adjustFormulaOps(unesc(v), entries, info.name);
              return res.changed ? ` ${n}="${escAttr(res.text)}"` : m0;
            });
            this.pkg.write(r.path, y);
          } else if (REL('vmlDrawing')(r)) {
            const x = await this.pkg.read(r.path);
            // the Excel VML namespace has whatever prefix the writer chose (x: from Excel, ns2: from openpyxl)
            const xp = nsPrefix(x ?? '', 'urn:schemas-microsoft-com:office:excel', 'x:');
            if (!x || x.indexOf(`<${xp}Fmla`) < 0) continue;
            const y = x.replace(new RegExp(`<${reEsc(xp)}(FmlaLink|FmlaRange|FmlaTxbx)>([\\s\\S]*?)</${reEsc(xp)}\\1>`, 'g'), (m0, n: string, v: string) => {
              const res = adjustFormulaOps(unesc(v), entries, info.name);
              return res.changed ? `<${xp}${n}>${esc(res.text)}</${xp}${n}>` : m0;
            });
            this.pkg.write(r.path, y);
          }
        }
      }
      // 3) charts
      for (const p of this.pkg.list()) {
        if (!/(^|\/)charts\/chart[^/]*\.xml$/i.test(p)) continue;
        const x = await this.pkg.read(p);
        if (!x) continue;
        // chart formulas: <c:f> (Excel, XlsxWriter), <f> in the default namespace (openpyxl), c15:f / c15:sqref / cx:f extensions
        const y = x.replace(/<((?:[\w.-]+:)?(?:f|sqref))(\s[^>]*)?>([^<]*)<\/\1>/g, (m0, tag: string, at: string | undefined, body: string) => {
          const r = adjustFormulaOps(unesc(body), entries, null);
          return r.changed ? `<${tag}${at ?? ''}>${esc(r.text)}</${tag}>` : m0;
        });
        this.pkg.write(p, y);
      }
      // 4) pivot caches
      for (const p of this.pkg.list()) {
        if (!/pivotCacheDefinition[^/]*\.xml$/i.test(p)) continue;
        const x = await this.pkg.read(p);
        if (!x) continue;
        let changed = false;
        let y = x.replace(/<worksheetSource\b[^>]*>/, (el) => {
          const sh = getAttr(el, 'sheet');
          const ref = getAttr(el, 'ref');
          if (!sh || !ref || getAttr(el, 'r:id')) return el;
          const a = parseAreaText(ref);
          if (!a) return el;
          const na = adjustPlainArea(a, entries, sh);
          if (!na || na === a) return el;
          const txt = renderSqref([na]);
          if (txt === ref) return el;
          changed = true;
          return setAttr(el, 'ref', txt);
        });
        if (changed) y = y.replace(/<pivotCacheDefinition\b[^>]*>/, (o) => setAttr(o, 'refreshOnLoad', '1'));
        this.pkg.write(p, y);
      }
      // 5) defined names
      const nfrom = this.namesEp;
      this.wbXml = this.wbXml.replace(/(<definedName\b[^>]*?(?<!\/)>)([\s\S]*?)(<\/definedName>)/g, (m0, a: string, body: string, c: string) => {
        const r = adjustFormulaOps(unesc(body), entries, null, 'name', nfrom);
        return r.changed ? a + esc(r.text) + c : m0;
      });
    }
    if (this.dynMeta && this.dynMeta.changed) this.pkg.write(this.dynMeta.path, this.dynMeta.xml);
    // _xlnm._FilterDatabase follows the sheet's autoFilter range (Excel keeps them in sync)
    for (const st of this.states.values()) {
      if (!st.changed) continue;
      const af = /<autoFilter\b[^>]*>/.exec(st.model.tail);
      const ref = af ? getAttr(af[0], 'ref') : undefined;
      const a = ref ? parseAreaText(ref) : null;
      if (!a) continue;
      const abs = quoteSheetName(st.info.name) + '!' + renderArea({ ...a, kind: 'area', ac1: true, ac2: true, ar1: true, ar2: true });
      this.wbXml = this.wbXml.replace(/(<definedName\b[^>]*?(?<!\/)>)([\s\S]*?)(<\/definedName>)/g, (m0, o: string, _body: string, c: string) => {
        if (getAttr(o, 'name') !== '_xlnm._FilterDatabase' || getAttr(o, 'localSheetId') !== String(st.info.index)) return m0;
        return o + esc(abs) + c;
      });
    }
    // 6) calcChain, full recalculation on load
    await this.dropCalcChain();
    this.wbXml = setFullCalc(this.wbXml);
    this.pkg.write(this.wbPath, this.wbXml);
    if (this.sstObj && this.sstPath && this.sstObj.changed) this.pkg.write(this.sstPath, this.sstObj.serialize());
    if (this.stylesObj && this.stylesObj.changed) await this.writeStyles();
    return this.pkg.generate();
  }

  /** Adjust every cell formula (+ array / data-table refs) of a sheet model. Returns true if something changed. */
  adjustCells(model: SheetModel, host: string, entries: readonly OpEntry[]): boolean {
    if (!entries.length) return false;
    let changed = false;
    for (const row of model.rows) {
      for (const cell of row.cells) {
        if (this.adjustCell(cell, host, entries, row.r)) changed = true;
      }
    }
    if (changed) model.dirty = true;
    return changed;
  }

  /**
   * Apply the pending entries (from cell.ep) to one cell's formula and array / data-table refs. `row`: the cell's row (the premerge row
   * extension needs to know whether the formula sits inside the data rows, see patch/formula.ts extendArea).
   */
  adjustCell(cell: CellM, host: string, entries: readonly OpEntry[], row?: number): boolean {
    const x = cell.x;
    if (!x || cell.ep >= entries.length || x.indexOf('<f') < 0) return false;
    const from = cell.ep;
    cell.ep = entries.length;
    const f = parseF(x);
    if (!f) return false;
    let attrs = f.attrs;
    const r = f.text ? adjustFormulaOps(f.text, entries, host, 'cell', from, null, row) : null;
    const ref = getAttr(attrs, 'ref');
    if (ref) {
      const a = parseAreaText(ref);
      const na = a ? adjustPlainArea(a, entries, host, from) : null;
      if (na && a && na !== a) attrs = setAttr(attrs, 'ref', renderSqref([na]));
    }
    for (const k of ['r1', 'r2']) {
      const v = getAttr(attrs, k);
      const a = v ? parseAreaText(v) : null;
      if (!a) continue;
      const na = adjustPlainArea(a, entries, host, from);
      if (na && na !== a) attrs = setAttr(attrs, k, renderSqref([na]));
    }
    if ((!r || !r.changed) && attrs === f.attrs) return false;
    const text = r && r.changed ? r.text : f.text;
    cell.x = x.slice(0, f.start) + (text ? `<f${attrs}>${esc(text)}</f>` : `<f${attrs}/>`) + x.slice(f.end);
    return true;
  }

  /** Formula-bearing features of a sheet tail (CF / DV formulas refer to OTHER sheets here, hyperlink locations, sparklines). */
  adjustTailFormulas(tail: string, host: string, entries: readonly OpEntry[]): string {
    if (!entries.length) return tail;
    let x = mapRuleFormulas(tail, (f) => adjustFormulaOps(f, entries, host, 'cell', 0, host).text);
    x = x.replace(/<hyperlink\b[^>]*>/g, (el) => {
      const loc = getAttr(el, 'location');
      if (!loc) return el;
      const r = adjustFormulaOps(loc, entries, host);
      return r.changed ? setAttr(el, 'location', r.text) : el;
    });
    x = x.replace(/<x14:sparklineGroup\b[\s\S]*?<\/x14:sparklineGroup>/g, (g) => g.replace(/<xm:f>([\s\S]*?)<\/xm:f>/g, (m0, body: string) => {
      const r = adjustFormulaOps(unesc(body), entries, host);
      return r.changed ? `<xm:f>${esc(r.text)}</xm:f>` : m0;
    }));
    return x;
  }

  private async writeStyles(): Promise<void> {
    const st = this.stylesObj!;
    if (this.stylesPath) { this.pkg.write(this.stylesPath, st.serialize()); return; }
    const path = dirOf(this.wbPath) + '/styles.xml';
    this.pkg.write(path, st.serialize());
    await this.pkg.addRel(this.wbPath, REL_TYPE.styles, path);
    await this.pkg.setContentTypeOverride(path, 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml');
    this.stylesPath = path;
  }

  private async dropCalcChain(): Promise<void> {
    const rels = await this.pkg.relTargets(this.wbPath);
    for (const r of rels.filter((x) => isRel(x.type, 'calcChain'))) {
      this.pkg.remove(r.path);
      await this.pkg.setContentTypeOverride(r.path, null);
    }
    await this.pkg.removeRel(this.wbPath, (r) => isRel(r.type, 'calcChain'));
    const cc = dirOf(this.wbPath) + '/calcChain.xml';
    if (this.pkg.has(cc)) { this.pkg.remove(cc); await this.pkg.setContentTypeOverride(cc, null); }
  }
}

// ---------------------------------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------------------------------

/** Excel's limit for the length of a formula ("数式が長すぎます"), in characters. */
export const MAX_FORMULA_LENGTH = 8192;

/** Refuse what Excel refuses or what would produce an invalid file (a formula that is too long / not text). */
function checkWriteValue(w: CellWrite, sheet: string): void {
  const v = w.cell;
  if (v === null) return;
  if (typeof v !== 'object') throw new PatchError('セルに書き込む内容が正しくありません。');
  if (v.f !== undefined && v.f !== null) {
    if (typeof v.f !== 'string') throw new PatchError('数式の指定が正しくありません。');
    // the limit counts what the user types: the _xlfn. / _xlws. / _xlpm. prefixes of the file form do not count
    const len = v.f.replace(/_xl(?:fn|ws|pm)\./g, '').length;
    if (len > MAX_FORMULA_LENGTH) {
      throw new PatchError(`シート「${sheet}」の ${cellRef(w.col, w.row)} に入力した数式が長すぎます（${len.toLocaleString('ja-JP')} 文字。Excel の上限は ${MAX_FORMULA_LENGTH.toLocaleString('ja-JP')} 文字です）。`);
    }
    // eslint-disable-next-line no-control-regex
    if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/.test(v.f)) throw new PatchError(`シート「${sheet}」の ${cellRef(w.col, w.row)} の数式に使用できない文字が含まれています。`);
  }
  if (v.numFmt !== undefined && v.numFmt !== null && typeof v.numFmt !== 'string') throw new PatchError('表示形式の指定が正しくありません。');
}

function validateOp(op: StructOp): void {
  if (!op || typeof op !== 'object' || (op.axis !== 'col' && op.axis !== 'row') || (op.type !== 'insert' && op.type !== 'delete' && op.type !== 'move')) throw new PatchError(MSG.badOp);
  const max = op.axis === 'col' ? MAX_COL : MAX_ROW;
  const ok = (n: number) => Number.isInteger(n) && n >= 1;
  if (!ok(op.count) || op.count > max) throw new PatchError(MSG.badOp);
  if (op.type === 'move') {
    if (!ok(op.from) || !ok(op.to) || op.from + op.count - 1 > max || op.to > max + 1) throw new PatchError(MSG.badOp);
  } else if (!ok(op.at) || op.at > max || (op.type === 'delete' && op.at + op.count - 1 > max)) throw new PatchError(MSG.badOp);
}

export function norm(a: Area): Area {
  if (a.kind === 'cols' || a.kind === 'rows') {
    return a.kind === 'cols'
      ? { ...a, kind: 'area', r1: 1, r2: MAX_ROW, c1: Math.min(a.c1, a.c2), c2: Math.max(a.c1, a.c2) }
      : { ...a, kind: 'area', c1: 1, c2: MAX_COL, r1: Math.min(a.r1, a.r2), r2: Math.max(a.r1, a.r2) };
  }
  return { ...a, c1: Math.min(a.c1, a.c2), c2: Math.max(a.c1, a.c2), r1: Math.min(a.r1, a.r2), r2: Math.max(a.r1, a.r2) };
}

/** Cells of one row under a column op (`styles` given = Excel "format same as left" for inserts). */
function mapRowCells(row: RowM, op: StructOp, styles: Styles | null): CellM[] {
  const cells = row.cells;
  if (!cells.length) return cells;
  if (op.type === 'insert') {
    const left = styles && op.at > 1 ? findCell(row, op.at - 1) : undefined;
    let i = 0;
    while (i < cells.length && cells[i].c < op.at) i++;
    const head = cells.slice(0, i);
    const rest = cells.slice(i);
    const right = rest.length && rest[0].c === op.at ? rest[0] : undefined;
    const add: CellM[] = [];
    const ls = left ? cellStyle(left) : 0;
    if (ls && styles) {
      const ns = styles.insertedStyle(ls, right ? cellStyle(right) : 0);
      if (ns) for (let k = 0; k < op.count; k++) add.push({ c: op.at + k, a: ` s="${ns}"`, x: null, ep: 0 });
    }
    for (const c of rest) c.c += op.count;
    return head.concat(add, rest.filter((c) => c.c <= MAX_COL));
  }
  if (op.type === 'delete') {
    const s = op.at, e = op.at + op.count - 1;
    const out: CellM[] = [];
    for (const c of cells) {
      if (c.c < s) out.push(c);
      else if (c.c > e) { c.c -= op.count; out.push(c); }
    }
    return out;
  }
  const p = movePerm(op.from, op.count, op.to);
  const lo = Math.min(op.from, op.to), hi = Math.max(op.from + op.count - 1, op.to);
  let moved = false;
  for (const c of cells) if (c.c >= lo && c.c <= hi) { c.c = p(c.c); moved = true; }
  if (moved) cells.sort((a, b) => a.c - b.c);
  return cells;
}

/** Row attributes an inserted row takes from the row above (format, height, outline level). */
function inheritRowAttrs(a: string): string {
  let x = a;
  x = setAttr(x, 'hidden', null);
  x = setAttr(x, 'collapsed', null);
  x = setAttr(x, 'thickBot', null);
  x = setAttr(x, 'thickTop', null);
  return /\S/.test(x) ? x.replace(/\s+$/, '') : '';
}

/** Table columns touched by a delete (excelRef deletedTableParts). */
function tableHit(t: TableState, op: StructOp): TableHit | null {
  if (op.type !== 'delete') return null;
  const a = t.ref;
  const s1 = op.at, s2 = op.at + op.count - 1;
  const lo = op.axis === 'col' ? a.c1 : a.r1;
  const hi = op.axis === 'col' ? a.c2 : a.r2;
  if (s2 < lo || s1 > hi) return null;
  const whole = s1 <= lo && s2 >= hi;
  const cols = new Set<string>();
  if (op.axis === 'col') {
    const names = tableColumns(t.xml).map(colName);
    for (let c = Math.max(lo, s1); c <= Math.min(hi, s2); c++) { const n = names[c - lo]; if (n != null) cols.add(n.toLowerCase()); }
  }
  // (a structured reference to the totals row, Tbl[[#Totals],...], stays as it is when the totals row is deleted: measured,
  // golden h2-del-totals-ref)
  return { whole, cols };
}

/** <tableColumn> element texts in order. */
export function tableColumns(xml: string): string[] {
  const m = findFirst(xml, 'tableColumns');
  if (!m) return [];
  return findAll(m.inner, 'tableColumn').map((e) => e.text);
}

/** A <tableColumn> without its totals row definition (function / label / formula), as Excel saves it once the totals row is gone. */
function stripTotalsDefinition(el: string): string {
  let x = el.replace(/<totalsRowFormula(?:\s[^>]*)?(?:\/>|>[\s\S]*?<\/totalsRowFormula>)/g, '');
  x = x.replace(/^<tableColumn\b[^>]*>/, (open) => setAttr(setAttr(setAttr(open, 'totalsRowFunction', null), 'totalsRowLabel', null), 'totalsRowDxfId', null));
  return x.replace(/^(<tableColumn\b[^>]*?)\s*><\/tableColumn>$/, '$1/>'); // nothing left inside: Excel writes <tableColumn …/>
}

export function colName(el: string): string {
  return unTableName(getAttr(el, 'name') ?? '');
}

/** Table column names in XML use _xHHHH_ for line breaks etc. */
function unTableName(s: string): string {
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, h: string) => String.fromCharCode(parseInt(h, 16)));
}
export function escTableName(s: string): string {
  // eslint-disable-next-line no-control-regex
  return s.replace(/[\u0000-\u001F]/g, (ch) => '_x' + ch.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0') + '_');
}

/**
 * The names Excel 365 gives to `count` table columns inserted right of the column named `leftName` (measured with
 * Excel through COM; the same rule as the editor's doc/tables.ts excelNewColumnNames — tests/xlsx/patch-round3.test.ts
 * keeps the two identical): a left neighbour ending in a number "h1" -> "h2" when that name is free (case-insensitive),
 * else "h12", "h13" ...; a left neighbour without a trailing number (or none) -> "列1", "列2" ... the first free one.
 * `existing` = the current column names of the table.
 */
export function excelNewColumnNames(existing: string[], leftName: string | null, count: number): string[] {
  const taken = new Set(existing.map((s) => s.toLowerCase()));
  const out: string[] = [];
  const take = (n: string) => {
    taken.add(n.toLowerCase());
    out.push(n);
  };
  const m = leftName != null ? /^(.*?)(\d+)$/su.exec(leftName) : null;
  for (let k = 0; k < count; k++) {
    if (m && leftName != null) {
      const next = m[1] + String(Number(m[2]) + 1);
      if (!taken.has(next.toLowerCase())) {
        take(next);
        continue;
      }
      for (let i = 2; ; i++) {
        const c = leftName + String(i);
        if (!taken.has(c.toLowerCase())) {
          take(c);
          break;
        }
      }
      continue;
    }
    for (let i = 1; ; i++) {
      const c = `列${i}`;
      if (!taken.has(c.toLowerCase())) {
        take(c);
        break;
      }
    }
  }
  return out;
}

/** The first free 列N (an empty table header). */
function uniqueNames(existing: string[], count: number): string[] {
  return excelNewColumnNames(existing, null, count);
}

/**
 * The column name Excel stores for a typed header (unique case-insensitively within the table: 「名前」 taken ->
 * 「名前2」, 「名前3」 ...; an empty header -> the first free 列N).
 */
export function uniqueColumnName(text: string, others: string[]): string {
  const used = new Set(others.map((s) => s.toLowerCase()));
  if (text === '') return uniqueNames(others, 1)[0];
  if (!used.has(text.toLowerCase())) return text;
  for (let n = 2; ; n++) {
    const c = text + n;
    if (!used.has(c.toLowerCase())) return c;
  }
}

/** Text a table header cell receives for a write (header cells always hold text). */
function headerText(v: CellWrite['cell']): string {
  if (!v || v.v === null || v.v === undefined || v.t === 'z') return '';
  if (v.t === 'b') return v.v === true || v.v === 1 || v.v === 'TRUE' ? 'TRUE' : 'FALSE';
  if (typeof v.v === 'number') return numText(v.v);
  return String(v.v).slice(0, 255);
}

/** totals-row definition of one table column (attributes / child of <tableColumn>) */
interface TotalsDef {
  label?: string;
  fn?: string;
  formula?: string;
}

/** SUBTOTAL function numbers Excel maps to the built-in totals functions (probe-totals.ps1). */
const TOTALS_FN: Record<string, string> = { '101': 'average', '102': 'countNums', '103': 'count', '104': 'max', '105': 'min', '107': 'stdDev', '109': 'sum', '110': 'var' };

/**
 * The built-in totalsRowFunction a totals-row formula stands for: exactly SUBTOTAL(n,Table[column]) of the column
 * it sits in (canonical text, as typed without extra spaces), else null (= custom).
 */
function totalsFunctionOf(formula: string, table: string, column: string): string | null {
  const m = /^SUBTOTAL\((\d+),(.+)\)$/i.exec(formula);
  if (!m) return null;
  const fn = TOTALS_FN[m[1]];
  if (!fn) return null;
  const item = `[${escapeSrefItem(column)}]`.toLowerCase();
  const ref = m[2].toLowerCase();
  return ref === `${table}${item}`.toLowerCase() || ref === item ? fn : null;
}

/** The text a totals-row cell keeps when a constant is typed into it (the column's totalsRowLabel). */
function totalsLabelText(v: NonNullable<CellWrite['cell']>): string {
  if (v.v === null || v.v === undefined || v.t === 'z') return '';
  if (v.t === 'b') return v.v === true || v.v === 1 || v.v === 'TRUE' ? 'TRUE' : 'FALSE';
  if (typeof v.v === 'number') {
    if (typeof v.numFmt === 'string' && v.numFmt.trim()) {
      const text = formatValue(v.v, v.numFmt.trim()).text;
      if (text && !/^#+$/.test(text)) return text.slice(0, 255);
    }
    return numText(v.v);
  }
  return String(v.v).slice(0, 255);
}

/** Replace the pos-th <tableColumn> element of a table part. */
function mapTableColumnAt(xml: string, pos: number, fn: (el: string) => string): string {
  const tc = findFirst(xml, 'tableColumns');
  if (!tc || tc.selfClosing) return xml;
  const el = findAll(tc.inner, 'tableColumn')[pos];
  if (!el) return xml;
  const at = tc.start + tc.open.length + el.start;
  const repl = fn(el.text);
  return repl === el.text ? xml : xml.slice(0, at) + repl + xml.slice(at + el.text.length);
}

/**
 * A <tableColumn> element with the given totals definition: totalsRowLabel / totalsRowFunction attributes and the
 * <totalsRowFormula> child (schema order: calculatedColumnFormula, totalsRowFormula, xmlColumnPr, extLst).
 */
function withTotalsDef(el: string, def: TotalsDef): string {
  let open = /^<tableColumn\b[^>]*>/.exec(el)![0];
  const selfClosing = open.endsWith('/>');
  let inner = selfClosing ? '' : el.slice(open.length, el.length - '</tableColumn>'.length);
  open = setAttr(setAttr(open, 'totalsRowLabel', null), 'totalsRowFunction', null);
  inner = inner.replace(/<totalsRowFormula\b(?:[^>]*\/>|[^>]*>[\s\S]*?<\/totalsRowFormula>)/g, '');
  if (def.fn) open = setAttr(open, 'totalsRowFunction', def.fn);
  if (def.label !== undefined) open = setAttr(open, 'totalsRowLabel', def.label);
  if (def.formula !== undefined) {
    const f = `<totalsRowFormula>${esc(def.formula)}</totalsRowFormula>`;
    const cc = /<calculatedColumnFormula\b[^>]*>[\s\S]*?<\/calculatedColumnFormula>/.exec(inner);
    inner = cc ? inner.slice(0, cc.index + cc[0].length) + f + inner.slice(cc.index + cc[0].length) : f + inner;
  }
  if (!inner) return selfClosing ? open : open.replace(/\s*>$/, '/>'); // Excel writes childless columns self-closed
  return (selfClosing ? open.replace(/\s*\/>$/, '>') : open) + inner + '</tableColumn>';
}

/** The values spilled by a dynamic-array master (formula-less cells of its range) are removed, styles kept. */
function clearSpill(model: SheetModel, a: Area, master: CellM): void {
  for (let r = a.r1; r <= a.r2; r++) {
    const row = model.row(r);
    if (!row) continue;
    for (const sc of row.cells) {
      if (sc.c < a.c1 || sc.c > a.c2 || sc === master || !sc.x || sc.x.indexOf('<f') >= 0) continue;
      sc.x = null;
      sc.a = setAttr(sc.a, 't', null);
    }
  }
}

/** Rewrite children of a counted container; drops the container when it becomes empty. */
function rewriteCounted(xml: string, container: string, child: string, fn: (el: string) => string): string {
  const c = findFirst(xml, container);
  if (!c || c.selfClosing) return xml;
  const newInner = c.inner.replace(new RegExp(`<${child}\\b[^>]*?(?:/>|>[\\s\\S]*?</${child}>)`, 'g'), (el) => fn(el));
  if (newInner === c.inner) return xml;
  const n = (newInner.match(new RegExp(`<${child}\\b`, 'g')) ?? []).length;
  const repl = n ? setAttr(c.open, 'count', /\scount="/.test(c.open) ? String(n) : null) + newInner + `</${container}>` : '';
  return xml.slice(0, c.start) + repl + xml.slice(c.end);
}

function rewriteContainer(xml: string, container: string, child: string, fn: (el: string) => string): string {
  return rewriteCounted(xml, container, child, fn);
}

function mapFirst(xml: string, tag: string, fn: (el: string) => string): string {
  const m = findFirst(xml, tag);
  if (!m) return xml;
  const r = fn(m.text);
  if (r === m.text) return xml;
  return xml.slice(0, m.start) + r + xml.slice(m.end);
}

function sqrefAttr(el: string, fn: (a: Area) => Area | null): string {
  const sq = getAttr(el, 'sqref');
  if (!sq) return el;
  const areas = parseSqref(sq).map(fn).filter((a): a is Area => !!a);
  if (!areas.length) return '';
  const txt = renderSqref(areas);
  return txt === sq ? el : setAttr(el, 'sqref', txt);
}

/** Sheet-level autoFilter: ref (with premerge extension) + filterColumn colIds. Removed when its range is deleted. */
function adjustAutoFilter(el: string, op: StructOp, fn: (a: Area) => Area | null): string {
  const open = /^<autoFilter\b[^>]*>/.exec(el)![0];
  const ref = getAttr(open, 'ref');
  const a = ref ? parseAreaText(ref) : null;
  if (!a) return el;
  const na = fn(a);
  if (!na) return '';
  // filter column ids and the sort state change even when the range itself does not (moves inside it)
  let x = na === a ? el : el.replace(open, setAttr(open, 'ref', renderSqref([stripAbs(na)])));
  x = remapFilterColumns(x, op, norm(a), norm(na));
  x = x.replace(/<sortState\b[^>]*?(\/>|>[\s\S]*?<\/sortState>)/, (s) => adjustSortState(s, (ar) => adjustArea(ar, op)));
  return x;
}

/** filterColumn colId is relative to the filter's first column. */
function remapFilterColumns(x: string, op: StructOp, before: Area, after: Area): string {
  if (op.axis !== 'col' || x.indexOf('<filterColumn') < 0) return x;
  return x.replace(/<filterColumn\b[^>]*?(?:\/>|>[\s\S]*?<\/filterColumn>)/g, (fc) => {
    const id = Number(getAttr(fc, 'colId') ?? -1);
    if (id < 0) return fc;
    const abs = before.c1 + id;
    const nabs = mapIndex(abs, op);
    if (nabs == null || nabs < after.c1 || nabs > after.c2) return '';
    const nid = nabs - after.c1;
    return nid === id ? fc : fc.replace(/colId="\d+"/, `colId="${nid}"`);
  });
}

function adjustSortState(el: string, fn: (a: Area) => Area | null): string {
  let x = el.replace(/^<sortState\b[^>]*>/, (o) => {
    const a = parseAreaText(getAttr(o, 'ref') ?? '');
    const na = a ? fn(a) : null;
    return na && na !== a ? setAttr(o, 'ref', renderSqref([na])) : o;
  });
  x = x.replace(/<sortCondition\b[^>]*>/g, (o) => {
    const a = parseAreaText(getAttr(o, 'ref') ?? '');
    const na = a ? fn(a) : null;
    return na && na !== a ? setAttr(o, 'ref', renderSqref([na])) : o;
  });
  return x;
}

/** Manual page breaks: <brk id="n"> = break after row/column n (the next page starts at n+1). */
function adjustBreaks(tail: string, container: string, op: StructOp): string {
  const x = rewriteCounted(tail, container, 'brk', (el) => {
    const id = Number(getAttr(el, 'id') ?? 0);
    if (!id) return el;
    const nv = mapIndex(id + 1, op);
    if (nv == null || nv - 1 < 1) return '';
    return nv - 1 === id ? el : setAttr(el, 'id', String(nv - 1));
  });
  if (x === tail) return x;
  const c = findFirst(x, container);
  if (!c || !/\smanualBreakCount="/.test(c.open)) return x;
  const n = (c.inner.match(/<brk\b[^>]*\sman="1"/g) ?? []).length;
  return x.slice(0, c.start) + setAttr(c.open, 'manualBreakCount', String(n)) + x.slice(c.start + c.open.length);
}

function adjustSparklineLocations(tail: string, op: StructOp): string {
  if (tail.indexOf('<x14:sparkline>') < 0) return tail;
  let x = tail.replace(/<x14:sparkline>([\s\S]*?)<\/x14:sparkline>/g, (m0, inner: string) => {
    const sq = /<xm:sqref>([\s\S]*?)<\/xm:sqref>/.exec(inner);
    if (!sq) return m0;
    const areas = parseSqref(unesc(sq[1])).map((a) => adjustArea(a, op)).filter((a): a is Area => !!a);
    if (!areas.length) return '';
    return m0.replace(sq[0], `<xm:sqref>${renderSqref(areas)}</xm:sqref>`);
  });
  x = x.replace(/<x14:sparklineGroup\b[^>]*>[\s\S]*?<\/x14:sparklineGroup>/g, (g) => (/<x14:sparkline>/.test(g) ? g : ''));
  x = x.replace(/<x14:sparklineGroups\b[^>]*>\s*<\/x14:sparklineGroups>/g, '');
  return x;
}

/** comments / threadedComments: elements with a cell `ref`; removed when their cell is deleted. */
function adjustRefElements(xml: string, tag: string, op: StructOp): string {
  const re = new RegExp(`<${tag}\\b[^>]*?(?:/>|>[\\s\\S]*?</${tag}>)`, 'g');
  return xml.replace(re, (el) => {
    const open = /^<[^>]*>/.exec(el)![0];
    const ref = getAttr(open, 'ref');
    const a = ref ? parseAreaText(ref) : null;
    if (!a) return el;
    const na = adjustArea(a, op);
    if (!na) return '';
    return na === a ? el : el.replace(open, setAttr(open, 'ref', renderSqref([na])));
  });
}

function setDimension(head: string, ref: string): string {
  if (/<dimension\b/.test(head)) return head.replace(/<dimension\b[^>]*>/, (d) => setAttr(d, 'ref', ref));
  return head;
}

function setFullCalc(wb: string): string {
  if (/<calcPr\b/.test(wb)) return wb.replace(/<calcPr\b[^>]*>/, (c) => setAttr(c, 'fullCalcOnLoad', '1'));
  // CT_Workbook is a sequence: ... sheets, functionGroups, externalReferences, definedNames, calcPr, oleSize ... — calcPr goes after the
  // last of these that exists (a calcPr before <externalReferences> makes Excel refuse the whole file: workbooks with external links, hc-37)
  for (const tag of ['definedNames', 'externalReferences', 'functionGroups', 'sheets']) {
    const m = new RegExp(`</${tag}>|<${tag}\\b[^>]*/>`).exec(wb);
    if (m) return wb.slice(0, m.index + m[0].length) + '<calcPr fullCalcOnLoad="1"/>' + wb.slice(m.index + m[0].length);
  }
  return wb;
}

/** `needleLower` (decoded, lower case) occurs in the XML text, also when the writer spelled it with numeric character references / CDATA. */
function mentions(xml: string, needleLower: string): boolean {
  const low = xml.toLowerCase();
  if (low.includes(needleLower)) return true;
  return (low.indexOf('&#') >= 0 || low.indexOf('<![cdata[') >= 0) && unesc(low).includes(needleLower);
}

/** the text mentions `sheet` as a reference prefix ("name!", "'name'!", first sheet of a 3-D range "name:other!"), also in the XML-escaped spellings of the name */
function namesSheet(hay: string, sheetLower: string): boolean {
  for (const b of [sheetLower, sheetLower.replace(/'/g, "''")]) {
    for (const v of new Set([b, esc(b).toLowerCase(), escAttr(b).toLowerCase()])) {
      if (hay.includes(v + '!') || hay.includes(v + "'!") || hay.includes(v + '&apos;!') || hay.includes(v + ':')) return true;
    }
  }
  return false;
}

/**
 * Cheap pre-filter: can this sheet XML contain references affected by the ops? A sheet that is not an op sheet can only reach an op
 * sheet through a qualified reference ("名簿!A1", "'My Sheet'!A1": the name followed by "!") or a structured reference to one of the tables
 * the ops touched; a sheet whose name merely contains the op sheet's name (a copy "名簿_名簿" next to the merged sheet "名簿") is not a hit.
 */
function mayReference(xml: string, sheetsLower: Set<string>, tablesLower: Set<string>): boolean {
  if (xml.indexOf('<f') < 0 && xml.indexOf('<formula') < 0 && xml.indexOf('<xm:f') < 0 && xml.indexOf('location=') < 0) return false;
  const low = xml.toLowerCase();
  for (const s of sheetsLower) if (namesSheet(low, s)) return true;
  for (const t of tablesLower) if (low.includes(t)) return true;
  // a writer that spells text with numeric character references (&#x540d;) or CDATA: look at the decoded text
  if (low.indexOf('&#') >= 0 || low.indexOf('<![cdata[') >= 0) {
    const dec = unesc(low);
    for (const s of sheetsLower) if (namesSheet(dec, s)) return true;
    for (const t of tablesLower) if (dec.includes(t)) return true;
  }
  return false;
}

function numText(n: number): string {
  if (Object.is(n, -0)) return '0';
  return String(n);
}

export { cellRef, relsPathOf, parseRels, resolveTarget, rewriteCounted };
