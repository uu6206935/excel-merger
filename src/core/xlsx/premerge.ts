/**
 * buildPremerged (owner: patcher P): builds the merged package from the base package BEFORE any user edit.
 * See ./index.ts and docs/ARCHITECTURE.md §2 decisions 1, 6, 7, 8.
 *
 *  1. rows that exist only in other files (join 'full') are inserted after the base region's last data row with
 *     Excel insert-row semantics PLUS the range extension (ranges ending at the old last data row and starting at /
 *     above the first data row grow, so 合計 formulas include the new rows); they copy the last data row's format;
 *     their key cells are filled from `keyFill`;
 *  2. the imported columns are inserted right after the base region (a CLEAN insert: no format inheritance, no CF
 *     leak) and filled with header + data cells carrying their SOURCE styles (merged into the base styles.xml);
 *  3. imported formulas stay live formulas (round 8): the formula remapper (formula/remap.ts) rewrites every reference to the
 *     merged sheet — rows through the join plan (a matched source row lands on the base row of its key, not on the same row
 *     number), ranges over the whole source table become the whole merged table, key columns that are not imported read the
 *     base's partner column, defined names and table structured references are resolved. What cannot be remapped (other sheets,
 *     cells outside the table, positional functions …) is imported as the cached value, counted per reason in
 *     MergeReport.formulaStats; a formula that cannot be kept AND has no stored result is the only way to get an empty cell;
 *  4. a base ListObject / AutoFilter over the region grows over the imported columns (and the appended rows).
 */
import { numToCol, parseAreaText, type Area } from '../formula/excelRef';
import { isUnsafeFormula } from '../export/comJob';
import { FormulaRemapper, type RemapName, type RemapReason, type RemapResult, type RemapTable } from '../formula/remap';
import type { CellType, CheckItem, MergeReport, MergedRowPlan, SourceBook } from '../types';
import { cellText } from './cellText';
import type { KeyFillEntry, KeySource, PremergeInput, PremergeLayout, PremergeResult, ProgressFn } from '../contracts';
import { dropCheckItems, dropOtherSheets, type DropResult } from './patch/dropsheets';
import { Engine, colName, tableColumns, to1904Serial, type SheetState, type TableState } from './patch/engine';
import { PatchError, isRel, toPatchError } from './patch/pkg';
import { cellRef, parseF, type CellM } from './patch/sheet';
import { isoToSerial } from './sheet';
import { isRichOrPhonetic, siText, type Sst } from './patch/sst';
import { Styles, convertColors, themeColors } from './patch/styles';
import { escAttr, getAttr, setAttr, unesc } from './patch/xml';

/**
 * What happened to ONE formula cell of an imported column (the optional third argument of buildPremerged; the findings / oracle tools
 * and the tests use it). row / col = the cell in the merged sheet, srcRow / srcCol = where it came from.
 */
export interface FormulaTrace {
  sourceId: string;
  srcRow: number;
  srcCol: number;
  row: number;
  col: number;
  /** the formula of the source (FILE form, no '=') */
  source: string;
  outcome: 'kept' | 'value' | 'empty';
  /** kept: the formula written (FILE form); */
  formula?: string;
  /** not kept: why (RemapReason or 'policy') */
  reason?: RemapReason | 'policy';
  /** kept: the source's cached result was written next to it (false = Excel recalculates it on load) */
  cached?: boolean;
  /** kept: every reference reads the very same source values (false: a whole-table range over other rows, a renamed header cell …) */
  exact?: boolean;
  /** kept: a whole-table range now reads the merged table (not the same set of rows as in the source) */
  widened?: boolean;
  /** kept: a non-imported key column was replaced by the base's partner column */
  substituted?: boolean;
}

export type FormulaTraceFn = (t: FormulaTrace) => void;

interface SrcVal {
  t: CellType;
  v: string | number | boolean | null;
  f?: string;
  /** array / dynamic-array formula (kept only when it is a single-cell array: see arrRef) */
  arr?: boolean;
  /** the `ref` of an array master (t="array"); a single-cell array may stay a (CSE / dynamic-array) formula */
  arrRef?: string;
  /** dynamic-array formula (cm="n") */
  dyn?: boolean;
  s: number;
  /** a string with rich text runs / phonetic data: the <si> / <is> body of the SOURCE package */
  rich?: string;
  /** an ISO 8601 date cell (t="d"): `v` is its serial in the 1900 system whatever date system the source book uses */
  iso?: boolean;
}

interface SourceCtx {
  ks: KeySource;
  /** the source package (workbook.xml: defined names when the SourceBook has no sheets) */
  eng: Engine;
  /**
   * The results stored next to the formulas are placeholders (XlsxWriter writes 0 for every formula and sets fullCalcOnLoad): they are
   * treated as "no result stored" — importing a 0 as the value of a formula that Excel would compute as 78 is worse than an empty cell.
   */
  placeholderCache: boolean;
  st: SheetState;
  sst: Sst;
  styles: Styles;
  /** source theme colours when they differ from the base theme (theme refs are then converted to RGB) */
  themeRgb: string[] | null;
  /** the source's legacy indexed palette when it differs from the base's (indexed colour refs are then converted to RGB) */
  indexedRgb: string[] | null;
  xfCache: Map<number, number>;
  valuesOnly: boolean;
  /** source column -> merged column, for the imported columns of this source */
  colMap: Map<number, number>;
  /** source column -> most common data xf (already imported into the base styles) */
  commonXf: Map<number, number>;
  /** days to add to date serials (1904 <-> 1900 date systems), 0 when both books use the same system */
  dateShift: number;
  /** the base book uses the 1904 date system */
  base1904: boolean;
  /** source rich <si> body -> the same body converted for the base package */
  richCache: Map<string, string>;
}

const MAX_COL = 16384;
const MAX_ROW = 1048576;

/**
 * Failure contract: a correct package or a PatchError with a Japanese message (hostile / damaged files included:
 * anything unexpected is converted by toPatchError, XML parts the patcher wrote are checked for well-formedness).
 */
export async function buildPremerged(input: PremergeInput, onProgress?: ProgressFn, trace?: FormulaTraceFn): Promise<PremergeResult> {
  try {
    return await build(input, onProgress, trace);
  } catch (e) {
    throw toPatchError(e, '結合ファイルを作成できませんでした');
  }
}

async function build(input: PremergeInput, onProgress?: ProgressFn, trace?: FormulaTraceFn): Promise<PremergeResult> {
  const report = (ratio: number | null, message: string) => { try { onProgress?.({ ratio, message }); } catch { /* ignore UI errors */ } };
  const { base, sources, rows, imports, keyFill } = input;
  const o: Partial<PremergeInput['options']> = input.options ?? {};
  const options: PremergeInput['options'] = {
    formulaPolicy: o.formulaPolicy ?? 'keep-if-possible',
    extendTable: o.extendTable ?? true,
    headerConflict: o.headerConflict ?? 'suffix',
    dropOtherSheets: o.dropOtherSheets === true,
  };
  const region = base.region;
  validateRegion(region.sheet, region);
  validatePlan(rows, imports);
  report(0, '基準ファイルを読み込み中…');
  const eng = await Engine.open(base.book.bytes);
  let st: SheetState;
  try {
    st = await eng.state(region.sheet);
  } catch {
    throw new PatchError(`基準ファイル「${base.label}」にシート「${region.sheet}」が見つかりません。`);
  }
  const sheet = st.info.name;
  const items: CheckItem[] = [];
  // ---- 0. single-sheet output: the other sheets and everything that only they use go FIRST, so that every later step (and the
  // replay / both export engines) sees a one-sheet workbook; formulas that read a removed sheet were turned into values
  let drop: DropResult | null = null;
  if (options.dropOtherSheets) {
    report(0.02, '他のシートを取り除いています…');
    drop = await dropOtherSheets(eng, st);
  }
  const baseStyles = await eng.styles();
  const baseTheme = await themeXml(eng);
  const base1904 = is1904(eng);

  // ---- sources --------------------------------------------------------------------------------------------
  const byId = new Map<string, KeySource>();
  for (const s of sources) byId.set(s.id, s);
  for (const imp of imports) if (!byId.has(imp.sourceId)) throw new PatchError('取り込む列のファイルが見つかりません（ファイルが削除された可能性があります）。もう一度結合してください。');
  const ctxs = new Map<string, SourceCtx>();
  const needed = new Set<string>(imports.map((i) => i.sourceId));
  for (const r of rows) if (r.baseRow == null && r.appendedFrom) needed.add(r.appendedFrom.sourceId);
  let k = 0;
  for (const id of needed) {
    const ks = byId.get(id);
    if (!ks) throw new PatchError('追加する行のファイルが見つかりません。もう一度結合してください。');
    report(0.05 + (0.15 * k++) / Math.max(1, needed.size), `「${ks.label}」を読み込み中…`);
    ctxs.set(id, await openSource(ks, baseTheme, base1904, baseStyles));
  }
  for (const c of ctxs.values()) if (c.valuesOnly && imports.some((i) => i.sourceId === c.ks.id)) {
    items.push({ severity: 'info', code: 'valuesOnly', sourceId: c.ks.id, message: `「${c.ks.label}」は値だけを読み込める形式のため、書式なし（標準）で取り込みました。` });
  }

  // ---- 1. appended rows -------------------------------------------------------------------------------------
  const appendedPlans = rows.filter((r) => r.baseRow == null);
  const N = appendedPlans.length;
  const appendAt = N ? region.lastDataRow + 1 : null;
  if (N && region.lastDataRow + N > MAX_ROW) throw new PatchError('追加する行が多すぎます（Excel の行数の上限を超えます）。');
  const tablesBefore = st.tables.filter((t) => !t.removed).map((t) => ({ t, ref: { ...t.ref } }));
  const afBefore = autoFilterArea(st.model.tail);
  if (N) {
    report(0.25, `行を追加中…（${N} 行）`);
    await eng.struct(sheet, { type: 'insert', axis: 'row', at: appendAt!, count: N }, {
      ext: { firstDataRow: region.firstDataRow, lastDataRow: region.lastDataRow },
      extTables: options.extendTable,
      exactFormat: true,
    });
    items.push({ severity: 'info', code: 'appendedRows', message: `基準ファイルにない ${N} 行を表の末尾（${appendAt} 行目から）に追加しました。`, rows: range(appendAt!, Math.min(N, 20)) });
  }

  // ---- 2. imported columns (clean insert) -------------------------------------------------------------------
  const M = imports.length;
  const insertAt = region.lastCol + 1;
  if (M && insertAt + M - 1 > MAX_COL) throw new PatchError('取り込む列が多すぎます（Excel の列数の上限 16,384 列を超えます）。');
  if (M) {
    report(0.35, `列を挿入中…（${M} 列）`);
    await eng.struct(sheet, { type: 'insert', axis: 'col', at: insertAt, count: M }, { clean: true });
  }
  imports.forEach((imp, j) => ctxs.get(imp.sourceId)!.colMap.set(imp.col, insertAt + j));

  // ---- headers ------------------------------------------------------------------------------------------------
  const bsst = await eng.sst();
  const existing = new Set<string>();
  for (let c = region.firstCol; c <= region.lastCol; c++) {
    const cell = st.model.cell(region.headerRow, c);
    const t = headerText(base.book, region.sheet, region.headerRow, c, cell ? textOf(readCell(cell, bsst)) : '');
    if (t) existing.add(normHeader(t));
  }
  const layoutImported: PremergeLayout['imported'] = [];
  const headers: string[] = [];
  /** the header cell text of each imported column in its source (a formula that reads the header cell is exact only while the merged header is the same) */
  const srcHeaderText: string[] = [];
  for (let j = 0; j < M; j++) {
    const imp = imports[j];
    const c = ctxs.get(imp.sourceId)!;
    const srcHeader = c.st.model.cell(c.ks.region.headerRow, imp.col);
    const rawHeader = headerText(c.ks.book, c.ks.region.sheet, c.ks.region.headerRow, imp.col, textOf(readCell(srcHeader, c.sst))).trim();
    srcHeaderText.push(rawHeader);
    let h = (imp.header ?? rawHeader).trim();
    if (h && options.headerConflict === 'suffix' && existing.has(normHeader(h))) {
      let nh = `${h}（${c.ks.label}）`;
      for (let n = 2; existing.has(normHeader(nh)); n++) nh = `${h}（${c.ks.label} ${n}）`;
      items.push({ severity: 'warning', code: 'headerRenamed', sourceId: c.ks.id, message: `見出し「${h}」が既存の列と同じため「${nh}」にしました。` });
      h = nh;
    }
    if (h) existing.add(normHeader(h));
    headers.push(h);
  }

  // ---- 4. table / autofilter extension (before writing the header cells: table names must equal them) --------
  const table = findRegionTable(tablesBefore, region);
  if (table && options.extendTable && M) {
    const t = table.t;
    if (!t.removed) {
      const used = new Set(tableColumns(t.xml).map((x) => colName(x).toLowerCase()));
      for (let j = 0; j < M; j++) {
        let name = headers[j] || `列${j + 1}`;
        const baseName = name;
        for (let n = 2; used.has(name.toLowerCase()); n++) name = `${baseName}${n}`;
        used.add(name.toLowerCase());
        if (name !== headers[j] && headers[j]) items.push({ severity: 'info', code: 'headerRenamed', message: `テーブルの列名は重複できないため、見出し「${headers[j]}」を「${name}」にしました。` });
        headers[j] = name;
      }
      if (t.header) eng.extendTableColumns(t, headers.slice());
      items.push({ severity: 'info', code: 'tableExtended', message: `テーブル「${t.name}」を取り込んだ列${N ? 'と追加した行' : ''}まで広げました。` });
    }
  } else if (table && M && !options.extendTable) {
    items.push({ severity: 'info', code: 'tableNotExtended', message: `テーブル「${table.t.name}」は広げていません（取り込んだ列はテーブルの外になります）。` });
  }
  // a table the insertion landed INSIDE (the region is narrower than the table) got Excel's 列1.. columns: rename
  for (const t of st.tables) {
    if (t.removed || !t.header || !M) continue;
    if (!(t.ref.c1 < insertAt && t.ref.c2 >= insertAt + M - 1) || t === table?.t) continue;
    renameTableColumns(t, insertAt, headers);
  }
  const af = autoFilterArea(st.model.tail);
  if (afBefore && af && options.extendTable && M && !table && afBefore.r1 === region.headerRow && afBefore.c2 === region.lastCol && afBefore.c1 <= region.firstCol) {
    st.model.tail = st.model.tail.replace(/<autoFilter\b[^>]*>/, (o) => setAttr(o, 'ref', `${numToCol(af.c1)}${af.r1}:${numToCol(af.c2 + M)}${af.r2}`));
    items.push({ severity: 'info', code: 'filterExtended', message: `フィルターの範囲を取り込んだ列${N ? 'と追加した行' : ''}まで広げました。` });
  }

  // ---- 3. header + data cells --------------------------------------------------------------------------------
  report(0.45, 'データを書き込み中…');
  const ep = eng.entries.length;
  for (let j = 0; j < M; j++) {
    const imp = imports[j];
    const c = ctxs.get(imp.sourceId)!;
    const col = insertAt + j;
    const hsrc = c.st.model.cell(c.ks.region.headerRow, imp.col);
    const hxf = c.valuesOnly ? 0 : importXf(c, hsrc ? styleOf(hsrc) : 0, baseStyles);
    const hcell = st.model.ensureCell(region.headerRow, col);
    setStyle(hcell, hxf);
    eng.setCellSync(hcell, headers[j] ? { t: 's', v: headers[j] } : null, ep, bsst);
    layoutImported.push({ col, sourceId: imp.sourceId, sourceCol: imp.col, header: headers[j] });
    // width
    const wa = c.st.cols.get(imp.col);
    const w = wa ? Number(getAttr(wa, 'width')) : NaN;
    if (Number.isFinite(w) && w > 0 && !c.valuesOnly) await eng.setColWidth(sheet, col, w);
  }
  // ---- formula remappers (one per source, created when its first formula is met) --------------------------------------
  /** merged row of every plan row (appended rows follow the base data rows) */
  const targets: number[] = [];
  for (let i = 0, a = 0; i < rows.length; i++) targets.push(rows[i].baseRow == null ? appendAt! + a++ : rows[i].baseRow!);
  // A formula the remapper cannot handle (even by a defect in it) becomes a value: it must never fail the whole merge
  const GIVE_UP: RemapResult = { ok: false, reason: 'unsupported', detail: 'internal error' };
  const remappers = new Map<string, FormulaRemapper | null>();
  const remapFormula = (sid: string, f: string, srcRow: number, targetRow: number, srcCol: number): RemapResult => {
    try {
      let r = remappers.get(sid);
      if (r === undefined) {
        try {
          r = createRemapper({
            c: ctxs.get(sid)!, ctxs, rows, targets, region, N, insertAt, M, imports, headers, srcHeaderText,
            keyPairs: input.keyPairs?.[sid] ?? [], keyFill, baseKs: base, st, bsst,
          });
        } catch {
          r = null;
        }
        remappers.set(sid, r);
      }
      return r ? r.remap(f, srcRow, targetRow, srcCol) : GIVE_UP;
    } catch {
      return GIVE_UP;
    }
  };
  const stats = new Map<string, FStat>();
  const statOf = (sid: string): FStat => {
    let s = stats.get(sid);
    if (!s) { s = { kept: 0, toValues: 0, empty: 0, reasons: {}, valueReasons: {}, emptyReasons: {}, reasonRows: new Map(), emptyRows: [], widened: 0, recalc: 0, recalcRows: [], widenedRows: [] }; stats.set(sid, s); }
    return s;
  };
  let kept = 0, toValues = 0;
  let matchedRows = 0, baseOnly = 0;
  const layoutAppended: PremergeLayout['appended'] = [];
  for (let i = 0; i < rows.length; i++) {
    if (i % 500 === 0) report(0.45 + (0.45 * i) / Math.max(1, rows.length), `データを書き込み中…（${i}/${rows.length} 行）`);
    const plan = rows[i];
    const target = targets[i];
    if (plan.baseRow == null) {
      if (plan.appendedFrom) {
        layoutAppended.push({ row: target, sourceId: plan.appendedFrom.sourceId, sourceRow: plan.appendedFrom.row });
        const c = ctxs.get(plan.appendedFrom.sourceId);
        for (const kf of keyFill[plan.appendedFrom.sourceId] ?? []) {
          if (!c) break;
          if (!kf || !Number.isInteger(kf.baseCol) || kf.baseCol < 1 || kf.baseCol > MAX_COL) continue;
          const v = keyFillValue(c, kf, plan.appendedFrom.row);
          if (!v) continue;
          const cell = st.model.ensureCell(target, kf.baseCol);
          // the key cell keeps the base column's format: runs relying on the source cell font get it explicitly
          writeValue(eng, cell, v, ep, bsst, c, baseStyles, true);
        }
      }
    } else {
      if (Object.values(plan.matches ?? {}).some((m) => m)) matchedRows++; else baseOnly++;
    }
    for (let j = 0; j < M; j++) {
      const imp = imports[j];
      const c = ctxs.get(imp.sourceId)!;
      const col = insertAt + j;
      const m = plan.matches?.[imp.sourceId];
      const sr = m ? m.row : plan.appendedFrom && plan.appendedFrom.sourceId === imp.sourceId ? plan.appendedFrom.row : null;
      const cell = st.model.ensureCell(target, col);
      const scell = sr != null ? c.st.model.cell(sr, imp.col) : undefined;
      const v0 = scell ? readCell(scell, c.sst) : null;
      const v = v0 ? toBase(c, v0) : null;
      if (v && v.f && c.placeholderCache) { v.t = 'z'; v.v = null; }
      let xf = c.valuesOnly ? 0 : v && scell && (v.v != null || v.f || styleOf(scell)) ? importXf(c, styleOf(scell), baseStyles) : commonXf(c, imp.col, baseStyles);
      // an ISO date cell whose source format is not a date format would show its serial: give it a date format
      if (v?.iso && typeof v.v === 'number' && baseStyles.dateMode(xf) === 'none') xf = baseStyles.withNumFmt(xf, isoFormat(v.v));
      setStyle(cell, xf);
      if (!v || (v.v == null && !v.f)) continue;
      if (v.f) {
        const stat = statOf(c.ks.id);
        const at = { sourceId: c.ks.id, srcRow: sr!, srcCol: imp.col, row: target, col, source: v.f };
        let res: RemapResult | null = null;
        let reason: RemapReason | 'policy' | null = null;
        // a formula that would make Excel reach outside the workbook (DDE, [book.xlsx] / UNC / URL references, WEBSERVICE ...) is never kept
        // live: it comes from a file of unknown origin, the cached value is imported instead (same rule as the Excel engine's typed formulas)
        if (options.formulaPolicy !== 'keep-if-possible') reason = 'policy';
        else if (isUnsafeFormula(v.f)) reason = 'unsafe';
        else if (v.arr && !isSingleCellArray(v.arrRef)) reason = 'array';
        else {
          res = remapFormula(c.ks.id, v.f, sr!, target, imp.col);
          if (!res.ok) reason = res.reason;
        }
        if (res && res.ok) {
          // the source's result stays next to the formula only while every reference still reads the very same values; else Excel recalculates
          // it on load (fullCalcOnLoad) — never a stale value next to a changed formula
          const cached = res.exact && v.v != null && v.t !== 'z';
          await putFormula(eng, cell, v, res.formula, cached, target, col, ep, bsst);
          kept++;
          stat.kept++;
          if (!cached) { stat.recalc++; if (stat.recalcRows.length < 20 && !stat.recalcRows.includes(target)) stat.recalcRows.push(target); }
          if (res.widened) { stat.widened++; if (stat.widenedRows.length < 20 && !stat.widenedRows.includes(target)) stat.widenedRows.push(target); }
          trace?.({ ...at, outcome: 'kept', formula: res.formula, cached, exact: res.exact, widened: res.widened, substituted: res.substituted });
          continue;
        }
        toValues++;
        const why = reason ?? 'unsupported';
        stat.reasons[why] = (stat.reasons[why] ?? 0) + 1;
        if (v.v == null) {
          // nothing to import: the source stored no result and the formula cannot be kept
          stat.empty++;
          stat.emptyReasons[why] = (stat.emptyReasons[why] ?? 0) + 1;
          if (stat.emptyRows.length < 20 && !stat.emptyRows.includes(target)) stat.emptyRows.push(target);
          trace?.({ ...at, outcome: 'empty', reason: why });
          continue;
        }
        stat.toValues++;
        stat.valueReasons[why] = (stat.valueReasons[why] ?? 0) + 1;
        const rr = stat.reasonRows.get(why) ?? [];
        if (rr.length < 20 && !rr.includes(target)) rr.push(target);
        stat.reasonRows.set(why, rr);
        trace?.({ ...at, outcome: 'value', reason: why });
        eng.setCellSync(cell, { t: v.t, v: v.v }, ep, bsst);
        continue;
      }
      writeValue(eng, cell, v, ep, bsst, c, baseStyles);
    }
  }
  // ---- report ------------------------------------------------------------------------------------------------------
  const formulaStats: NonNullable<MergeReport['formulaStats']> = [];
  for (const [sid, s] of stats) {
    const c = ctxs.get(sid)!;
    const label = c.ks.label;
    formulaStats.push({ sourceId: sid, label, kept: s.kept, toValues: s.toValues, empty: s.empty, reasons: { ...s.reasons } });
    if (c.placeholderCache) {
      items.push({
        severity: 'warning', code: 'formulasPlaceholderCache', sourceId: sid,
        message: `「${label}」は、数式の計算結果が保存されていない形式のファイルです（プログラムで作られたファイルでは、結果が 0 のまま保存されていることがあります）。0 は取り込まず、数式のまま残せるものは数式にして Excel で計算されるようにしました。元のファイルを Excel で開いて保存し直すと、結果も取り込めます。`,
      });
    }
    if (s.kept) {
      items.push({ severity: 'info', code: 'formulasKept', sourceId: sid, message: `「${label}」の数式 ${s.kept} 件は数式のまま残しました（参照先を結合後の列・行に付け替えました）。` });
    }
    if (s.widened) {
      items.push({
        severity: 'info', code: 'formulaRangesWidened', sourceId: sid, rows: s.widenedRows,
        message: `「${label}」の表全体を参照している数式 ${s.widened} 件は、結合後の表全体（基準ファイルの行と追加した行）を参照するようにしました。結合されなかった行（未一致）は計算に入りません。`,
      });
    }
    if (s.recalc) {
      items.push({
        severity: 'info', code: 'formulasRecalc', sourceId: sid, rows: s.recalcRows,
        message: `「${label}」の数式 ${s.recalc} 件は、結合後の表に合わせて付け替えたため、計算結果は Excel で開いたときに表示されます（この画面では空欄に見えることがあります）。`,
      });
    }
    // one item per reason, in plain words (values only: the empty cells have their own warning)
    for (const [why, n] of Object.entries(s.valueReasons)) {
      if (why === 'policy') {
        items.push({ severity: 'warning', code: 'formulasToValues', sourceId: sid, rows: s.reasonRows.get(why), message: `「${label}」の数式 ${n} 件を値に変換して取り込みました。` });
        continue;
      }
      items.push({
        severity: 'warning', code: 'formulasToValues', sourceId: sid, rows: s.reasonRows.get(why),
        message: `「${label}」の${REASON_JA[why] ?? '対応していない書き方の数式'} ${n} 件は値にしました。${REASON_HINT[why] ?? ''}`,
      });
    }
    if (s.empty) {
      const why = Object.entries(s.emptyReasons).map(([k, n]) => `${REASON_JA[k] ?? '対応していない書き方の数式'} ${n} 件`).join('、');
      items.push({
        severity: 'warning', code: 'formulasNoCache', sourceId: sid, rows: s.emptyRows,
        message: `「${label}」の数式 ${s.empty} 件は、元のファイルに計算結果が保存されていない数式です（Excelで開いて保存し直してください）。数式のままでは取り込めない形（${why}）のため、空欄になりました。元のファイルを Excel で開いて保存し直してから、もう一度結合してください。`,
      });
    }
  }

  if (drop) {
    // rows below the table move down with the appended rows
    const shifted = (r: number) => (N && appendAt != null && r >= appendAt ? r + N : r);
    items.unshift(...dropCheckItems(drop, shifted));
  }
  report(0.92, 'ファイルを作成中…');
  const bytes = await eng.finish();
  report(1, '完了');
  const reportOut: MergeReport = {
    mergedRows: rows.length,
    matchedRows,
    baseOnlyRows: baseOnly,
    appendedRows: N,
    importedColumns: M,
    formulasKept: kept,
    formulasToValues: toValues,
    formulaStats,
    ...(drop ? { droppedSheets: drop.dropped, formulasDetached: drop.detached } : {}),
    items,
  };
  const layout: PremergeLayout = { sheet, insertAt, imported: layoutImported, appendAt, appended: layoutAppended };
  return { bytes, layout, report: reportOut };
}

// ---------------------------------------------------------------------------------------------------------------

function validateRegion(sheet: string, r: PremergeInput['base']['region']): void {
  const ok = (n: number, max: number) => Number.isInteger(n) && n >= 1 && n <= max;
  if (!sheet || !ok(r.headerRow, MAX_ROW) || !ok(r.firstDataRow, MAX_ROW) || !ok(r.lastDataRow, MAX_ROW) || !ok(r.firstCol, MAX_COL) || !ok(r.lastCol, MAX_COL) ||
    r.firstCol > r.lastCol || r.firstDataRow > r.lastDataRow + 1) {
    throw new PatchError('基準ファイルの表の範囲（見出し行・データ行・列）が正しくありません。');
  }
}

/** row / column numbers of the plan are written into the sheet: they must be real positions */
function validatePlan(rows: PremergeInput['rows'], imports: PremergeInput['imports']): void {
  if (!Array.isArray(rows) || !Array.isArray(imports)) throw new PatchError('結合の指定が正しくありません。もう一度結合してください。');
  for (const imp of imports) {
    if (!imp || !Number.isInteger(imp.col) || imp.col < 1 || imp.col > MAX_COL) throw new PatchError('取り込む列の指定が正しくありません。もう一度結合してください。');
  }
  for (const r of rows) {
    if (!r || (r.baseRow != null && (!Number.isInteger(r.baseRow) || r.baseRow < 1 || r.baseRow > MAX_ROW))) throw new PatchError('結合する行の指定が正しくありません。もう一度結合してください。');
  }
}

async function themeXml(eng: Engine): Promise<string | null> {
  const rel = (await eng.pkg.relTargets(eng.wbPath)).find((r) => isRel(r.type, 'theme'));
  return rel ? eng.pkg.read(rel.path) : null;
}

function fontScheme(theme: string | null): string {
  return /<a:fontScheme\b[\s\S]*?<\/a:fontScheme>/.exec(theme ?? '')?.[0].replace(/\sname="[^"]*"/, '') ?? '';
}

function is1904(eng: Engine): boolean {
  return /<workbookPr\b[^>]*\sdate1904="(1|true)"/.test(eng.wbXml);
}

function samePalette(a: Styles, b: Styles): boolean {
  const pa = a.effectivePalette(), pb = b.effectivePalette();
  return pa.every((c, i) => c === pb[i]);
}

async function openSource(ks: KeySource, baseTheme: string | null, base1904: boolean, baseStyles: Styles): Promise<SourceCtx> {
  let eng: Engine;
  try {
    eng = await Engine.open(ks.book.bytes);
  } catch (e) {
    // the package guards (zip.ts: expands beyond 1.5 GB, a broken part) say what is wrong: pass their text on
    const why = e instanceof PatchError ? e.message : '';
    throw new PatchError(`「${ks.label}」を開けませんでした（ファイルが壊れている可能性があります）。${why}`);
  }
  let st: SheetState;
  try {
    st = await eng.state(ks.region.sheet);
  } catch {
    throw new PatchError(`「${ks.label}」にシート「${ks.region.sheet}」が見つかりません。`);
  }
  const srcTheme = await themeXml(eng);
  const srcStyles = await eng.styles();
  const same = JSON.stringify(themeColors(srcTheme)) === JSON.stringify(themeColors(baseTheme)) && fontScheme(srcTheme) === fontScheme(baseTheme);
  return {
    ks,
    eng,
    placeholderCache: /<calcPr\b[^>]*\sfullCalcOnLoad="(1|true)"/.test(eng.wbXml) && hasPlaceholderCache(st, ks.region),
    st,
    sst: await eng.sst(),
    styles: srcStyles,
    themeRgb: same ? null : themeColors(srcTheme),
    indexedRgb: samePalette(srcStyles, baseStyles) ? null : srcStyles.effectivePalette(),
    xfCache: new Map(),
    valuesOnly: ks.book.fidelity === 'values',
    colMap: new Map(),
    commonXf: new Map(),
    dateShift: is1904(eng) === base1904 ? 0 : is1904(eng) ? 1462 : -1462,
    base1904,
    richCache: new Map(),
  };
}

/**
 * Key cells of an appended row (KeyFillEntry): one source column (value kept as typed, rich text kept), several
 * columns joined with `joiner` (default full-width space: base 氏名 <- source 姓 + 名), and / or part `splitPart`
 * of the text split on white space (base 姓 <- part 0 of source 氏名). Empty parts are skipped.
 */
function keyFillValue(c: SourceCtx, kf: KeyFillEntry, row: number): SrcVal | null {
  const okCol = (n: unknown): n is number => typeof n === 'number' && Number.isInteger(n) && n >= 1 && n <= MAX_COL;
  const cols = Array.isArray(kf.sourceCols) && kf.sourceCols.length ? kf.sourceCols.filter(okCol) : okCol(kf.sourceCol) ? [kf.sourceCol] : [];
  if (!cols.length) return null;
  const read = (col: number) => {
    const v = readCell(c.st.model.cell(row, col), c.sst);
    return v && v.v != null ? toBase(c, v) : null;
  };
  const split = typeof kf.splitPart === 'number' && Number.isInteger(kf.splitPart) && kf.splitPart >= 0 ? kf.splitPart : null;
  if (typeof kf.splitPart === 'number' && split === null) return null;
  if (cols.length === 1 && split === null) {
    const v = read(cols[0]);
    return v ? { t: v.t, v: v.v, s: v.s, rich: v.rich } : null;
  }
  const single = cols.length === 1 ? read(cols[0]) : null;
  let text = cols.length === 1
    ? textOf(single).trim()
    : cols.map((col) => textOf(read(col)).trim()).filter((s) => s !== '').join(typeof kf.joiner === 'string' ? kf.joiner : '　');
  if (split !== null) {
    const parts = text.split(/[\s　]+/).filter((s) => s !== '');
    // one part only: the value itself (a number stays a number)
    if (single && parts.length === 1 && split === 0) return { t: single.t, v: single.v, s: single.s, rich: single.rich };
    text = parts[split] ?? '';
  }
  return text ? { t: 's', v: text, s: 0 } : null;
}

/**
 * Write a (non-formula) source value; rich text / phonetic strings keep their runs (copied into the base sst).
 * `ownFont`: the target cell does NOT carry the source cell's style (key cells of appended rows keep the base
 * format), so runs without <rPr> — which Excel renders with the CELL font — get the source cell's font explicitly.
 */
function writeValue(eng: Engine, cell: CellM, v: SrcVal, ep: number, bsst: Sst, c: SourceCtx, base: Styles, ownFont = false): void {
  if (v.t !== 's' || !v.rich || c.valuesOnly) {
    eng.setCellSync(cell, { t: v.t, v: v.v }, ep, bsst);
    return;
  }
  const key = ownFont ? `${v.s}\u0000${v.rich}` : v.rich;
  let conv = c.richCache.get(key);
  if (conv === undefined) {
    conv = convertRich(ownFont ? explicitRuns(v.rich, c, v.s) : v.rich, c, base);
    c.richCache.set(key, conv);
  }
  const idx = bsst.addRich(conv);
  eng.setCellSync(cell, null, ep, bsst);
  cell.a = setAttr(cell.a, 't', 's');
  cell.x = `<v>${idx}</v>`;
}

/**
 * A source <si> body in the base package: run fonts are inline (<rPr>), theme colours / theme font schemes are
 * converted when the themes differ (like imported cell styles); <phoneticPr fontId> refers to a font of the source
 * styles.xml and is re-pointed at the same font imported into the base (dropped when it does not exist).
 */
function convertRich(xml: string, c: SourceCtx, base: Styles): string {
  let x = xml;
  if (c.indexedRgb && !c.themeRgb) x = x.replace(/<rPr>([\s\S]*?)<\/rPr>/g, (_m, inner: string) => `<rPr>${convertColors(inner, null, c.indexedRgb)}</rPr>`);
  if (c.themeRgb) {
    const rgb = c.themeRgb;
    x = x.replace(/<rPr>([\s\S]*?)<\/rPr>/g, (_m, inner: string) => `<rPr>${convertColors(inner, rgb, c.indexedRgb)}</rPr>`);
  }
  x = x.replace(/<phoneticPr\b[^>]*\/>/g, (el) => {
    const id = getAttr(el, 'fontId');
    if (id === undefined) return el;
    const nid = base.importFont(c.styles, Number(id), c.themeRgb, c.indexedRgb);
    return nid === null ? '' : setAttr(el, 'fontId', String(nid));
  });
  return x;
}

/** Runs without <rPr> get the font of source xf `xf` as <rPr> (<name> becomes <rFont>); theme refs are converted later. */
function explicitRuns(xml: string, c: SourceCtx, xf: number): string {
  if (!/<r>(?!\s*<rPr\b)/.test(xml)) return xml;
  const fontId = Number(getAttr(/^<[^>]*>/.exec(c.styles.xf(xf))?.[0] ?? '', 'fontId') ?? 0);
  const font = c.styles.get('fonts', fontId) ?? c.styles.get('fonts', 0) ?? '';
  const inner = /^<font\b[^>]*>([\s\S]*)<\/font>$/.exec(font)?.[1] ?? '';
  const rpr = `<rPr>${inner.replace(/<name\b/g, '<rFont')}</rPr>`;
  return xml.replace(/<r>(?!\s*<rPr\b)/g, `<r>${rpr}`);
}

/**
 * A source value as written into the base (date serials moved to the base's date system). The rule is the reader's
 * (styles.ts dateModeOf, used by both conversions): calendar dates always move by 1462 days, a time of day only when it
 * carries a day count, durations and non-date formats never. ISO date cells are absolute: their serial only depends on the base.
 */
function toBase(c: SourceCtx, v: SrcVal): SrcVal {
  if (v.t !== 'n' || typeof v.v !== 'number') return v;
  if (v.iso) return c.base1904 && v.v >= 1462 ? { ...v, v: v.v - 1462 } : v;
  if (!c.dateShift) return v;
  const mode = c.styles.dateMode(v.s);
  if (mode === 'none' || mode === 'elapsed') return v;
  if (c.dateShift > 0) return mode === 'date' || v.v >= 1 ? { ...v, v: v.v + c.dateShift } : v;
  const n = to1904Serial(v.v, mode);
  return n === v.v ? v : { ...v, v: n };
}

/** display format for an ISO date that arrives in a cell without a date format */
function isoFormat(serial: number): string {
  if (serial < 1) return 'h:mm:ss';
  return Math.abs(serial - Math.floor(serial)) > 1e-9 ? 'yyyy/m/d h:mm' : 'yyyy/m/d';
}

function styleOf(c: CellM): number {
  const s = getAttr(c.a, 's');
  return s ? Number(s) || 0 : 0;
}

function setStyle(cell: CellM, xf: number): void {
  cell.a = setAttr(cell.a, 's', xf ? String(xf) : null);
}

function importXf(c: SourceCtx, srcXf: number, base: Styles): number {
  if (c.valuesOnly) return 0;
  return base.importXf(c.styles, srcXf, c.themeRgb, c.xfCache, c.indexedRgb);
}

/** the column's most common data style (unmatched rows keep borders / number formats consistent) */
function commonXf(c: SourceCtx, col: number, base: Styles): number {
  if (c.valuesOnly) return 0;
  const hit = c.commonXf.get(col);
  if (hit !== undefined) return hit;
  const counts = new Map<number, number>();
  const r = c.ks.region;
  for (let row = r.firstDataRow; row <= r.lastDataRow; row++) {
    const cell = c.st.model.cell(row, col);
    const s = cell ? styleOf(cell) : -1;
    if (s < 0) continue;
    counts.set(s, (counts.get(s) ?? 0) + 1);
  }
  let best = -1, bestN = 0;
  for (const [s, n] of counts) if (n > bestN) { best = s; bestN = n; }
  if (best < 0) {
    const colStyle = Number(getAttr(c.st.cols.get(col) ?? '', 'style') ?? 0);
    best = colStyle || 0;
  }
  const res = best ? importXf(c, best, base) : 0;
  c.commonXf.set(col, res);
  return res;
}

/** Cell value (cached value for formulas) of a source / base cell. */
function readCell(cell: CellM | undefined, sst: Sst | null): SrcVal | null {
  if (!cell) return null;
  const s = styleOf(cell);
  const t = getAttr(cell.a, 't');
  const x = cell.x ?? '';
  const f = parseF(x);
  // Excel writes the empty string result of a formula as <c t="str"><f>…</f><v/></c>: an empty text, not a missing result
  const vm = /<v>([\s\S]*?)<\/v>/.exec(x);
  const raw = vm ? unesc(vm[1]) : f && getAttr(cell.a, 't') === 'str' && /<v\s*\/>/.test(x) ? '' : null;
  const out: SrcVal = { t: 'z', v: null, s };
  if (f) {
    const ft = getAttr(f.attrs, 't');
    out.f = f.text;
    const isDyn = /\scm="/.test(cell.a);
    out.arr = ft === 'array' || ft === 'dataTable' || isDyn || !f.text;
    if (out.arr && ft === 'array' && f.text) { out.arrRef = getAttr(f.attrs, 'ref'); if (isDyn) out.dyn = true; }
  }
  if (t === 's') {
    const i = raw == null ? NaN : Number(raw);
    if (sst && Number.isInteger(i)) {
      out.t = 's'; out.v = sst.text(i);
      const rich = sst.richXml(i);
      if (rich && !f) out.rich = rich;
    }
  } else if (t === 'inlineStr') {
    const is = /<is>([\s\S]*?)<\/is>/.exec(x);
    out.t = 's'; out.v = is ? siText(is[1]) : '';
    if (is && !f && isRichOrPhonetic(is[1])) out.rich = is[1];
  } else if (t === 'str') {
    if (raw != null) { out.t = 's'; out.v = raw; }
  } else if (t === 'b') {
    if (raw != null) { out.t = 'b'; out.v = raw === '1' || raw.toLowerCase() === 'true'; }
  } else if (t === 'e') {
    if (raw != null) { out.t = 'e'; out.v = raw; }
  } else if (t === 'd') {
    // ISO 8601 date / time (Strict OOXML, openpyxl iso_dates, SheetJS cellDates): a serial, never the text
    const serial = raw != null ? isoToSerial(raw) : undefined;
    if (serial !== undefined) { out.t = 'n'; out.v = serial; out.iso = true; }
    else if (raw != null) { out.t = 's'; out.v = raw; }
  } else if (raw != null && raw !== '') {
    const n = Number(raw);
    if (Number.isFinite(n)) { out.t = 'n'; out.v = n; } else { out.t = 's'; out.v = raw; }
  }
  return out;
}

function textOf(v: SrcVal | null): string {
  if (!v || v.v == null) return '';
  if (typeof v.v === 'boolean') return v.v ? 'TRUE' : 'FALSE';
  return String(v.v);
}

/**
 * The header as Excel shows it: a date header (2024/04/01, 4月) or a number with a format reads as its DISPLAY text, not as the serial
 * (45383). The reader's SourceBook has the cell and the styles (cellText); a book without sheets (tests build bare ones) falls back to
 * the raw text of the package cell.
 */
function headerText(book: SourceBook, sheet: string, headerRow: number, col: number, raw: string): string {
  const cell = book.sheets.find((s) => s.name === sheet)?.rows[headerRow - 1]?.[col - 1];
  return cell && cell.t !== 's' ? cellText(book, cell) : raw;
}

function normHeader(s: string): string {
  return s.normalize('NFKC').trim().toLowerCase();
}

// ---------------------------------------------------------------------------------------------------------------
// formula import (round 8)
// ---------------------------------------------------------------------------------------------------------------

/** counters of ONE source's imported formulas */
interface FStat {
  kept: number;
  /** not kept, imported as the cached value */
  toValues: number;
  /** not kept and no cached value: the cell stays empty */
  empty: number;
  /** reason -> number of formulas (values and empty cells together) */
  reasons: Record<string, number>;
  /** reason -> number of formulas that became VALUES / that stayed EMPTY */
  valueReasons: Record<string, number>;
  emptyReasons: Record<string, number>;
  reasonRows: Map<string, number[]>;
  emptyRows: number[];
  /** kept, whole-table range now reading the merged table */
  widened: number;
  widenedRows: number[];
  /** kept without a stored result (Excel recalculates on load) */
  recalc: number;
  recalcRows: number[];
}

/** the reasons in plain Japanese ("<phrase> N 件は値にしました") */
const REASON_JA: Record<string, string> = {
  otherSheet: '他のシートを参照している数式',
  externalCell: '表の外のセル（例: 税率）を参照している数式',
  rowRelation: '前の行との差・途中までの合計など、結合後は行の並びが変わって同じ範囲を指せない数式',
  unresolvedName: '定義名（名前）を使っていて、置き換えられない数式',
  unresolvedTable: 'テーブル参照（[@列名] など）を置き換えられない数式',
  positional: 'ROW・COLUMN・OFFSET・INDIRECT など、セルの位置に結果が左右される数式',
  notImported: '取り込まない列を参照している数式',
  unsafe: '外部のブックなどを参照している数式',
  array: '配列数式（スピルを含む）',
  unsupported: '対応していない書き方の数式',
};

const REASON_HINT: Record<string, string> = {
  notImported: '（その列も取り込むと、数式のまま残せることがあります）',
  rowRelation: '（行の並びが元のファイルと同じなら数式のまま残せます）',
};

/**
 * Every formula of the data table stores a result of exactly 0 (and none a real one) in a book that asks Excel to recalculate on load: the
 * signature of XlsxWriter / similar writers that cannot compute — the zeros mean "not computed".
 */
function hasPlaceholderCache(st: SheetState, r: PremergeInput['base']['region']): boolean {
  let zeros = 0;
  for (const row of st.model.rows) {
    if (row.r < r.headerRow || row.r > r.lastDataRow) continue;
    for (const cell of row.cells) {
      const x = cell.x;
      if (!x || x.indexOf('<f') < 0) continue;
      const f = parseF(x);
      if (!f) continue;
      const vm = /<v>([\s\S]*?)<\/v>/.exec(x);
      if (!vm || vm[1] === '') continue; // no result stored
      const t = getAttr(cell.a, 't');
      if ((t == null || t === 'n') && Number(vm[1]) === 0) { zeros++; continue; }
      return false; // a real result
    }
  }
  return zeros > 0;
}

function isSingleCellArray(ref: string | undefined): boolean {
  if (!ref) return false;
  const a = parseAreaText(ref.replace(/\$/g, ''));
  return !!a && a.kind !== 'cols' && a.kind !== 'rows' && a.c1 === a.c2 && a.r1 === a.r2;
}

/**
 * Write a kept formula. `cached`: the source's result stays next to it (same inputs); else the formula is written without a result
 * (Excel computes it on load: fullCalcOnLoad is always set by Engine.finish). A single-cell array formula stays one (CSE: t="array",
 * dynamic array: cm + the metadata part), so that its array evaluation is preserved.
 */
async function putFormula(eng: Engine, cell: CellM, v: SrcVal, f: string, cached: boolean, row: number, col: number, ep: number, bsst: Sst): Promise<void> {
  const val = cached ? { t: v.t, v: v.v, f } : { t: 'z' as CellType, v: null, f };
  if (!v.arr) {
    eng.setCellSync(cell, val, ep, bsst);
    return;
  }
  const ref = cellRef(col, row);
  if (v.dyn) {
    eng.setCellSync(cell, val, ep, bsst, { cm: await eng.dynamicArrayCm(), ref });
    return;
  }
  eng.setCellSync(cell, val, ep, bsst);
  cell.x = (cell.x ?? '').replace(/^<f>/, `<f t="array" ref="${ref}">`);
}

interface RemapperArgs {
  c: SourceCtx;
  ctxs: Map<string, SourceCtx>;
  rows: MergedRowPlan[];
  targets: number[];
  region: PremergeInput['base']['region'];
  /** rows appended to the base table */
  N: number;
  insertAt: number;
  M: number;
  imports: PremergeInput['imports'];
  /** final merged header text per import (index = import order) */
  headers: string[];
  srcHeaderText: string[];
  keyPairs: Array<{ baseCol: number; sourceCol: number }>;
  keyFill: PremergeInput['keyFill'];
  /** the base file (its SourceBook shows the base cells' text) */
  baseKs: KeySource;
  st: SheetState;
  bsst: Sst;
}

function hasContent(cell: CellM): boolean {
  const x = cell.x;
  return x != null && (x.indexOf('<v>') >= 0 || x.indexOf('<f') >= 0 || x.indexOf('<is>') >= 0);
}

function createRemapper(a: RemapperArgs): FormulaRemapper {
  const { c, region } = a;
  const sid = c.ks.id;
  const sreg = c.ks.region;
  const rowPairs: Array<[number, number]> = [];
  for (let i = 0; i < a.rows.length; i++) {
    const plan = a.rows[i];
    const m = plan.matches?.[sid];
    const sr = m ? m.row : plan.appendedFrom && plan.appendedFrom.sourceId === sid ? plan.appendedFrom.row : null;
    if (sr != null) rowPairs.push([sr, a.targets[i]]);
  }
  // what is in the source sheet outside the data table / which data rows hold data (one pass over the sheet's rows)
  const outside = new Map<number, number[]>();
  const dataRows: number[] = [];
  for (const row of c.st.model.rows) {
    const inTable = row.r >= sreg.headerRow && row.r <= sreg.lastDataRow;
    if (inTable) {
      if (row.r >= sreg.firstDataRow && row.cells.some(hasContent)) dataRows.push(row.r);
      continue;
    }
    for (const cell of row.cells) {
      if (!hasContent(cell)) continue;
      let l = outside.get(cell.c);
      if (!l) outside.set(cell.c, (l = []));
      l.push(row.r);
    }
  }
  const hasContentOutside = (col: number, r1: number, r2: number): boolean => {
    const l = outside.get(col);
    if (!l) return false;
    let lo = 0, hi = l.length;
    while (lo < hi) { const mid = (lo + hi) >> 1; if (l[mid] < r1) lo = mid + 1; else hi = mid; }
    return lo < l.length && l[lo] <= r2;
  };
  const headerOf = new Map<number, boolean>();
  a.imports.forEach((imp, j) => {
    if (imp.sourceId === sid) headerOf.set(imp.col, !imp.header && a.headers[j] === a.srcHeaderText[j]);
  });
  const partnerCols = decidePartners(a);
  return new FormulaRemapper({
    sheet: sreg.sheet,
    region: { headerRow: sreg.headerRow, firstDataRow: sreg.firstDataRow, lastDataRow: sreg.lastDataRow, firstCol: sreg.firstCol, lastCol: sreg.lastCol },
    target: { headerRow: region.headerRow, firstDataRow: region.firstDataRow, lastDataRow: region.lastDataRow + a.N },
    colMap: c.colMap,
    partnerCols,
    rowPairs,
    dataRows,
    names: sourceNames(c),
    tables: sourceTables(c),
    hasContentOutside,
    headerSame: (col) => headerOf.get(col) === true,
  });
}

/** defined names of the source: from the SourceBook, or (a book without sheets: tests build bare ones) from workbook.xml */
function sourceNames(c: SourceCtx): RemapName[] {
  const b = c.ks.book;
  if (b.sheets.length > 0) return b.names.map((n) => ({ name: n.name, scope: n.scope, refersTo: n.refersTo }));
  const out: RemapName[] = [];
  for (const m of c.eng.wbXml.matchAll(/<definedName\b([^>]*?)(?:\/>|>([\s\S]*?)<\/definedName>)/g)) {
    const name = getAttr(m[1], 'name');
    if (!name) continue;
    const lsi = getAttr(m[1], 'localSheetId');
    out.push({ name: unesc(name), scope: lsi != null ? c.eng.sheets[Number(lsi)]?.name ?? null : null, refersTo: unesc(m[2] ?? '').trim() });
  }
  return out;
}

function sourceTables(c: SourceCtx): RemapTable[] {
  const b = c.ks.book;
  if (b.sheets.length > 0) return b.sheets.flatMap((sh) => sh.features.tables.map((t) => ({ name: t.name, sheet: sh.name, ref: t.ref, headerRow: t.headerRow, totalsRow: t.totalsRow, columns: t.columns })));
  return c.st.tables.filter((t) => !t.removed).map((t) => ({
    name: t.name,
    sheet: c.st.info.name,
    ref: `${numToCol(t.ref.c1)}${t.ref.r1}:${numToCol(t.ref.c2)}${t.ref.r2}`,
    headerRow: t.header > 0,
    totalsRow: t.totals > 0,
    columns: tableColumns(t.xml).map(colName),
  }));
}

/** the text a book shows for a cell of its sheet (null: the book has no cells, e.g. a test fixture) */
function shownText(ks: KeySource, row: number, col: number): string | null {
  const g = ks.book.sheets.find((s) => s.name === ks.region.sheet);
  if (!g) return null;
  const cell = g.rows[row - 1]?.[col - 1];
  return cell ? cellText(ks.book, cell) : '';
}

/**
 * Key columns that are not imported but are read by formulas (=A2&"様", COUNTIF($A$2:$A$61,A2), VLOOKUP(A2,…)): JoinResult.keyPairs says which
 * base column is aligned with such a source column. A formula may read the base column instead ONLY when both columns hold the same thing on
 * every matched row (same type, same value, same displayed text; for appended rows: what keyFill writes into the base cell) — a column-level
 * decision, so that one differently written ID (全角, 先頭ゼロ, 敬称) keeps the whole column out and the formula becomes a value.
 */
function decidePartners(a: RemapperArgs): Map<number, number> {
  const out = new Map<number, number>();
  const { c } = a;
  const sid = c.ks.id;
  for (const p of a.keyPairs) {
    if (!p || !Number.isInteger(p.baseCol) || !Number.isInteger(p.sourceCol) || p.baseCol < 1 || p.baseCol > MAX_COL || p.sourceCol < 1 || p.sourceCol > MAX_COL) continue;
    if (c.colMap.has(p.sourceCol) || out.has(p.sourceCol)) continue;
    let n = 0;
    let ok = true;
    for (let i = 0; i < a.rows.length && ok; i++) {
      const plan = a.rows[i];
      const m = plan.matches?.[sid];
      const sr = m ? m.row : plan.appendedFrom && plan.appendedFrom.sourceId === sid ? plan.appendedFrom.row : null;
      if (sr == null) continue;
      const sv = readCell(c.st.model.cell(sr, p.sourceCol), c.sst);
      let bv: SrcVal | null;
      let shownBase: string | null = null;
      if (plan.baseRow != null) {
        bv = readCell(a.st.model.cell(plan.baseRow, p.baseCol), a.bsst);
        shownBase = shownText(a.baseKs, plan.baseRow, p.baseCol);
      } else {
        // an appended row: its base key cells are filled from the file the row came from
        const from = plan.appendedFrom;
        const kf = from ? (a.keyFill[from.sourceId] ?? []).find((k) => k && k.baseCol === p.baseCol) : undefined;
        const fc = from ? a.ctxs.get(from.sourceId) : undefined;
        bv = kf && from && fc ? keyFillValue(fc, kf, from.row) : null;
      }
      const sEmpty = !sv || sv.v == null, bEmpty = !bv || bv.v == null;
      if (sEmpty || bEmpty) { ok = sEmpty && bEmpty; n++; continue; }
      if (sv!.t !== bv!.t || sv!.v !== bv!.v) { ok = false; break; }
      const shownSrc = shownText(c.ks, sr, p.sourceCol);
      if (shownBase != null && shownSrc != null && shownBase !== shownSrc) ok = false;
      n++;
    }
    if (ok && n > 0) out.set(p.sourceCol, p.baseCol >= a.insertAt ? p.baseCol + a.M : p.baseCol);
  }
  return out;
}

/** the ListObject whose header row is the region's header row and whose last column is the region's last column */
function findRegionTable(tables: Array<{ t: TableState; ref: Area }>, r: PremergeInput['base']['region']): { t: TableState; ref: Area } | null {
  for (const x of tables) {
    const a = x.ref;
    if (a.r1 === r.headerRow && a.c2 === r.lastCol && a.c1 <= r.firstCol) return x;
  }
  return null;
}

/** rename the tableColumns sitting on [from, from + names.length) to `names` (made unique) */
function renameTableColumns(t: TableState, from: number, names: string[]): void {
  const cols = tableColumns(t.xml);
  const used = new Set(cols.map((x, i) => (t.ref.c1 + i >= from && t.ref.c1 + i < from + names.length ? '' : colName(x).toLowerCase())));
  const next = cols.map((el, i) => {
    const pos = t.ref.c1 + i;
    if (pos < from || pos >= from + names.length) return el;
    let n = names[pos - from] || `列${pos - from + 1}`;
    const b = n;
    for (let k = 2; used.has(n.toLowerCase()); k++) n = `${b}${k}`;
    used.add(n.toLowerCase());
    names[pos - from] = n;
    return el.replace(/\sname="[^"]*"/, ` name="${escAttr(n)}"`);
  });
  t.xml = t.xml.replace(/<tableColumns\b[^>]*>[\s\S]*?<\/tableColumns>/, (m) => m.replace(/(<tableColumns\b[^>]*>)[\s\S]*(<\/tableColumns>)/, (_x, a: string, b: string) => a + next.join('') + b));
  t.dirty = true;
}

function autoFilterArea(tail: string): { c1: number; r1: number; c2: number; r2: number } | null {
  const m = /<autoFilter\b[^>]*>/.exec(tail);
  const ref = m ? getAttr(m[0], 'ref') : undefined;
  const p = ref ? /^\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)$/.exec(ref) : null;
  if (!p) return null;
  const col = (s: string) => { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
  return { c1: col(p[1]), r1: Number(p[2]), c2: col(p[3]), r2: Number(p[4]) };
}

function range(from: number, n: number): number[] {
  return Array.from({ length: n }, (_x, i) => from + i);
}

