/**
 * FROZEN COPY of src/core/xlsx/premerge.ts as it was BEFORE round 8 (task F1). Used only by the findings / oracle tooling
 * (tests/oracle/import-values/diagnose.ts) to print before / after numbers. Do not edit; do not import from src.
 */
import { numToCol, type Area } from '../../../../src/core/formula/excelRef';
import { isUnsafeFormula } from '../../../../src/core/export/comJob';
import { rewriteRefs } from '../../../../src/core/formula/forms';
import type { CellType, CheckItem, MergeReport, SourceBook } from '../../../../src/core/types';
import { cellText } from '../../../../src/core/xlsx/cellText';
import type { KeyFillEntry, KeySource, PremergeInput, PremergeLayout, PremergeResult, ProgressFn } from '../../../../src/core/contracts';
import { dropCheckItems, dropOtherSheets, type DropResult } from '../../../../src/core/xlsx/patch/dropsheets';
import { Engine, colName, tableColumns, to1904Serial, type SheetState, type TableState } from '../../../../src/core/xlsx/patch/engine';
import { PatchError, isRel, toPatchError } from '../../../../src/core/xlsx/patch/pkg';
import { parseF, type CellM } from '../../../../src/core/xlsx/patch/sheet';
import { isoToSerial } from '../../../../src/core/xlsx/sheet';
import { isRichOrPhonetic, siText, type Sst } from '../../../../src/core/xlsx/patch/sst';
import { Styles, convertColors, themeColors } from '../../../../src/core/xlsx/patch/styles';
import { escAttr, getAttr, setAttr, unesc } from '../../../../src/core/xlsx/patch/xml';

interface SrcVal {
  t: CellType;
  v: string | number | boolean | null;
  f?: string;
  /** array / dynamic-array formula (never kept live) */
  arr?: boolean;
  s: number;
  /** a string with rich text runs / phonetic data: the <si> / <is> body of the SOURCE package */
  rich?: string;
  /** an ISO 8601 date cell (t="d"): `v` is its serial in the 1900 system whatever date system the source book uses */
  iso?: boolean;
}

interface SourceCtx {
  ks: KeySource;
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
export async function buildPremerged(input: PremergeInput, onProgress?: ProgressFn): Promise<PremergeResult> {
  try {
    return await build(input, onProgress);
  } catch (e) {
    throw toPatchError(e, '結合ファイルを作成できませんでした');
  }
}

async function build(input: PremergeInput, onProgress?: ProgressFn): Promise<PremergeResult> {
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
  for (let j = 0; j < M; j++) {
    const imp = imports[j];
    const c = ctxs.get(imp.sourceId)!;
    const srcHeader = c.st.model.cell(c.ks.region.headerRow, imp.col);
    let h = (imp.header ?? headerText(c.ks.book, c.ks.region.sheet, c.ks.region.headerRow, imp.col, textOf(readCell(srcHeader, c.sst)))).trim();
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
  let kept = 0, toValues = 0;
  const valueRows = new Map<string, number[]>();
  const valueCount = new Map<string, number>();
  /** formulas that could not be kept AND whose source file stored no result (openpyxl, pandas, XlsxWriter without a value): the cell stays empty */
  const noCache = new Map<string, number>();
  const noCacheRows = new Map<string, number[]>();
  let matchedRows = 0, baseOnly = 0;
  const layoutAppended: PremergeLayout['appended'] = [];
  let ai = 0;
  for (let i = 0; i < rows.length; i++) {
    if (i % 500 === 0) report(0.45 + (0.45 * i) / Math.max(1, rows.length), `データを書き込み中…（${i}/${rows.length} 行）`);
    const plan = rows[i];
    let target: number;
    if (plan.baseRow == null) {
      target = appendAt! + ai++;
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
      target = plan.baseRow;
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
      let xf = c.valuesOnly ? 0 : v && scell && (v.v != null || v.f || styleOf(scell)) ? importXf(c, styleOf(scell), baseStyles) : commonXf(c, imp.col, baseStyles);
      // an ISO date cell whose source format is not a date format would show its serial: give it a date format
      if (v?.iso && typeof v.v === 'number' && baseStyles.dateMode(xf) === 'none') xf = baseStyles.withNumFmt(xf, isoFormat(v.v));
      setStyle(cell, xf);
      if (!v || (v.v == null && !v.f)) continue;
      if (v.f) {
        // a formula that would make Excel reach outside the workbook (DDE, [book.xlsx] / UNC / URL references, WEBSERVICE ...) is never kept
        // live: it comes from a file of unknown origin, the cached value is imported instead (same rule as the Excel engine's typed formulas)
        const f = options.formulaPolicy === 'keep-if-possible' && !v.arr && !isUnsafeFormula(v.f) ? remapFormula(v.f, c, sr!, target) : null;
        if (f != null) {
          kept++;
          eng.setCellSync(cell, { t: v.t, v: v.v, f }, ep, bsst);
          continue;
        }
        toValues++;
        valueCount.set(c.ks.id, (valueCount.get(c.ks.id) ?? 0) + 1);
        const l = valueRows.get(c.ks.id) ?? [];
        if (l.length < 20 && !l.includes(target)) l.push(target);
        valueRows.set(c.ks.id, l);
        if (v.v == null) {
          noCache.set(c.ks.id, (noCache.get(c.ks.id) ?? 0) + 1);
          const nl = noCacheRows.get(c.ks.id) ?? [];
          if (nl.length < 20 && !nl.includes(target)) nl.push(target);
          noCacheRows.set(c.ks.id, nl);
          continue;
        }
        eng.setCellSync(cell, { t: v.t, v: v.v }, ep, bsst);
        continue;
      }
      writeValue(eng, cell, v, ep, bsst, c, baseStyles);
    }
  }
  if (kept) items.push({ severity: 'info', code: 'formulasKept', message: `取り込んだ数式 ${kept} 件は数式のまま残しました（参照先を結合後の列に付け替えました）。` });
  for (const [sid, rs] of valueRows) {
    const c = ctxs.get(sid)!;
    const empty = noCache.get(sid) ?? 0;
    const n = (valueCount.get(sid) ?? 0) - empty;
    if (empty) {
      items.push({
        severity: 'warning', code: 'formulasNoCache', sourceId: sid, rows: noCacheRows.get(sid),
        message: `「${c.ks.label}」の数式 ${empty} 件は、計算結果が保存されていないため空欄になりました（取り込まない列・別のシート・名前などを参照していて数式のままでは取り込めません）。元のファイルを Excel で開いて保存し直してから、もう一度結合してください。`,
      });
    }
    if (n <= 0) continue;
    items.push({
      severity: 'warning', code: 'formulasToValues', sourceId: sid, rows: rs,
      message: options.formulaPolicy === 'values'
        ? `「${c.ks.label}」の数式 ${n} 件を値に変換して取り込みました。`
        : `「${c.ks.label}」の数式 ${n} 件は、取り込まない列・別のシート・名前などを参照しているため、値に変換して取り込みました。`,
    });
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
  const vm = /<v>([\s\S]*?)<\/v>/.exec(x);
  const raw = vm ? unesc(vm[1]) : null;
  const out: SrcVal = { t: 'z', v: null, s };
  if (f) {
    const ft = getAttr(f.attrs, 't');
    out.f = f.text;
    out.arr = ft === 'array' || ft === 'dataTable' || /\scm="/.test(cell.a) || !f.text;
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

/**
 * Keep a source formula live when every reference is a same-row reference to an imported column of the same
 * source sheet (or a same-row absolute reference); references are rewritten to the merged column / row.
 */
function remapFormula(f: string, c: SourceCtx, srcRow: number, targetRow: number): string | null {
  const srcSheet = c.ks.region.sheet.toLowerCase();
  return rewriteRefs(f, c.ks.region.sheet, (ref) => {
    if (!ref.sheet || ref.sheet.toLowerCase() !== srcSheet) return null;
    const a = ref.area;
    if (a.kind !== 'cell' && a.kind !== 'area') return null;
    if (a.r1 !== srcRow || a.r2 !== srcRow) return null;
    const lo = Math.min(a.c1, a.c2), hi = Math.max(a.c1, a.c2);
    const n1 = c.colMap.get(lo);
    if (n1 == null) return null;
    for (let col = lo; col <= hi; col++) if (c.colMap.get(col) !== n1 + (col - lo)) return null;
    const d = n1 - lo;
    const na: Area = { ...a, c1: a.c1 + d, c2: a.c2 + d, r1: targetRow, r2: targetRow };
    return { area: na, prefixText: '' };
  });
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

