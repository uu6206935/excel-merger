/**
 * includeSourceSheets (round 9, plan A) — the sheet copier of the package patcher.
 *
 * PremergeOptions.includeSourceSheets copies EVERY sheet (visible, hidden, very hidden; the source's data sheet too) of every other file of
 * the merge into the premerged package, after the merged sheet and the base's own other sheets. Names come from planSheetNames
 * (src/core/xlsx/sheetnames.ts: "<file name>_<sheet name>", 31 characters, "(1)" for the one added later). Formulas of a copy that read other
 * sheets of their own file read the renamed copies, so a copy never reads the merged sheet or a sheet of the base.
 *
 * Pieces (all under patch/):
 *   addsheets.ts   this file: planning (names), workbook-level registration (sheets, relationships, content types, names, app.xml),
 *                  the report, and the two extension points of the object phase (planSheetObjects / copySheetObjects)
 *   copysheet.ts   the core copy of ONE worksheet: cells, styles, shared strings, formulas, rows / columns, sheet properties, rules, links
 *   copyformula.ts sheet-reference rewriting of formulas (tokens, never regular expressions)
 *   copyrules.ts   conditional formats / data validations / sparklines (classic and x14)
 *   copynames.ts   defined names of the source for the copies (sheet-scoped on each copy)
 *
 * What the core copy leaves out (drawings, charts, pictures, comments, tables, pivot tables, controls ...) is removed cleanly (no dangling
 * r:id, no orphan part); copySheetObjects(ctx) (addsheets-objects.ts, task A2) adds drawings / charts / pictures / shapes / notes / tables /
 * pivot tables / the background back, and the chart sheets are copied as chart sheets; whatever is still left out (slicers, SmartArt, form
 * controls ...) is named in a 'sheetCopyLimits' warning per sheet.
 *
 * The copies are written straight into the package (they are not registered in Engine.sheets): the structural ops of the premerge never
 * touch them, and Engine.finish() writes the workbook parts. replayLog opens the premerged package afresh and sees them as ordinary sheets.
 */
import type { KeySource, PremergeLayout } from '../../contracts';
import type { CheckItem } from '../../types';
import { planSheetNames, type SheetNameRequest } from '../sheetnames';
import { collectTakenNames, copySheetObjects as copyObjects, chartSheetRids, planTables, prepareChartSheet, takeObjectInfo } from './addsheets-objects';
import { FormulaCopier } from './copyformula';
import { SourceNames } from './copynames';
import { copyWorksheet, splitTop, type SheetJob } from './copysheet';
import type { CopyPlan, CopySource, CopyTarget, NewRel, ObjectItem, ObjectKind, PlannedSheet, SheetCopy, SheetObjectsContext, SourcePlan } from './copytypes';
import type { Engine, SheetInfo, SheetState } from './engine';
import { convertNamespaces, is1904, isStrict, knownMdw, normalFontOf, samePalette, sameTheme, themeXml } from './bookinfo';
import { convertRich } from './richtext';
import { PatchError, REL_TYPE, dirOf, relativeTarget, relsPathOf, resolveTarget, type Rel } from './pkg';
import type { Sst } from './sst';
import { themeColors, type Styles } from './styles';
import { escAttr, esc, findFirst, getAttr, setAttr, unesc } from './xml';
import { FAIL_JA } from './copyformula';

export type { CopyPlan, SheetObjectsContext, ObjectItem, ObjectKind, NewRel } from './copytypes';

const WS_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const CHARTSHEET_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml';
const CHARTSHEET_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet';
/** Excel's limit of cell formats in a workbook */
export const XF_MAX = 64000;

// ---------------------------------------------------------------------------------------------------------------------
// extension points of the object phase (task A2)
// ---------------------------------------------------------------------------------------------------------------------

/** What A2 decides about the tables of a source BEFORE the formulas are copied (so that structured references can stay formulas). */
export interface ObjectsPlan {
  /** lower-cased table name of the source -> the name of its copy (tables that are not listed are not copied) */
  tables: ReadonlyMap<string, string>;
  /** lower-cased names of every table of the source */
  allTables: ReadonlySet<string>;
  /** lower-cased source sheet names whose tables are copied */
  sheetsWithTables: ReadonlySet<string>;
}

/** Which tables of the source are copied under which names (addsheets-objects-table.ts). */
export async function planSheetObjects(source: SourcePlan, taken: ReadonlySet<string>): Promise<ObjectsPlan> {
  return planTables(source, taken);
}

/**
 * Called for every copied sheet AFTER the core copy, with the new sheet XML and relationships in `ctx`: drawings, charts, pictures, notes,
 * tables, pivot tables (addsheets-objects.ts). What it copies is marked in ctx.handled; the rest stays reported as left out.
 */
export const copySheetObjects = copyObjects;

// ---------------------------------------------------------------------------------------------------------------------
// planning
// ---------------------------------------------------------------------------------------------------------------------

/** the `state` attribute of every <sheet> of workbook.xml, in order */
function sheetStates(wbXml: string): Array<'visible' | 'hidden' | 'veryHidden'> {
  const el = findFirst(wbXml, 'sheets');
  const out: Array<'visible' | 'hidden' | 'veryHidden'> = [];
  for (const m of (el?.inner ?? '').matchAll(/<sheet\b[^>]*>/g)) {
    const s = getAttr(m[0], 'state');
    out.push(s === 'hidden' || s === 'veryHidden' ? s : 'visible');
  }
  return out;
}

/** the sheets (SheetInfo) that are chart sheets: planSourceSheets finds them, addSourceSheets copies them as chart sheets */
const CHART_SHEETS = new WeakSet<SheetInfo>();
const copyable = (info: SheetInfo): boolean => (info.worksheet || CHART_SHEETS.has(info)) && !!info.path;

/**
 * The names of the copies: planSheetNames over every worksheet of every source (in source order, then workbook order) next to the sheets
 * that are in the target already (`existing`: the merged sheet + the base's other sheets that are kept). Chart / dialog / macro sheets are
 * not copied (name = null). `engineOf` opens (or reuses) the package of a source.
 */
export async function planSourceSheets(existing: readonly string[], sources: readonly KeySource[], engineOf: (ks: KeySource) => Promise<Engine>): Promise<CopyPlan> {
  const entries: SourcePlan[] = [];
  const requests: SheetNameRequest[] = [];
  for (const ks of sources) {
    const eng = await engineOf(ks);
    const states = sheetStates(eng.wbXml);
    // chart sheets are copied as chart sheets: they take a name like the worksheets (macro / dialog sheets are not copied)
    const charts = await chartSheetRids(eng);
    const sheets: PlannedSheet[] = eng.sheets.map((info, i) => ({ info, state: states[i] ?? 'visible', name: null, original: info.name, suffixed: false }));
    for (const s of sheets) if (!s.info.worksheet && charts.has(s.info.rid)) CHART_SHEETS.add(s.info);
    entries.push({ ks, eng, sheets });
    requests.push({ file: ks.book.fileName, sheets: sheets.filter((s) => copyable(s.info)).map((s) => s.info.name) });
  }
  const planned = planSheetNames(existing, requests);
  let k = 0;
  for (const e of entries) {
    for (const sh of e.sheets) {
      if (!copyable(sh.info)) continue;
      const p = planned[k++];
      sh.name = p.name;
      sh.suffixed = p.suffixed;
    }
  }
  return { sources: entries };
}

// ---------------------------------------------------------------------------------------------------------------------
// the copy
// ---------------------------------------------------------------------------------------------------------------------

export interface AddSheetsArgs {
  /** the target (premerged base) engine and the state of the merged sheet */
  eng: Engine;
  merged: SheetState;
  plan: CopyPlan;
  baseStyles: Styles;
  bsst: Sst;
  baseTheme: string | null;
  base1904: boolean;
  /** source id -> the xf cache premerge uses for the imported columns of that source (same mapping, no double work) */
  xfCaches?: Map<string, Map<number, number>>;
  progress?: (ratio: number, message: string) => void;
}

export interface AddSheetsResult {
  added: NonNullable<PremergeLayout['addedSheets']>;
  items: CheckItem[];
  /** the merged sheet had to move to the first tab */
  movedMergedFirst: boolean;
}

export async function addSourceSheets(a: AddSheetsArgs): Promise<AddSheetsResult> {
  const { eng } = a;
  const pkg = eng.pkg;
  const items: CheckItem[] = [];
  const added: NonNullable<PremergeLayout['addedSheets']> = [];
  const worksheetsBefore = eng.sheets.filter((s) => s.worksheet).map((s) => s.name);
  const movedMergedFirst = await moveMergedFirst(eng, a.merged);
  const strict = isStrict(eng.wbXml);
  const target: CopyTarget = { eng, styles: a.baseStyles, sst: a.bsst, xfMax: XF_MAX, strict, mdw: defaultFontMdw(a.baseStyles, a.baseTheme) };
  let nextIndex = eng.sheets.length;
  let sheetId = maxSheetId(eng.wbXml);
  const sheetEls: string[] = [];
  const nameEls: string[] = [];
  const outNames: string[] = []; // worksheets in output order (for app.xml)
  const total = a.plan.sources.reduce((n, s) => n + s.sheets.filter((x) => x.name !== null).length, 0);
  let done = 0;
  // the names of the tables and defined names of the target: a copied table must not take one of them (nor one of an earlier copy)
  const takenTables = await collectTakenNames(eng);

  for (const sp of a.plan.sources) {
    const ks = sp.ks;
    const copyable = sp.sheets.filter((s) => s.name !== null);
    for (const sh of sp.sheets) {
      if (sh.name === null) {
        items.push({ severity: 'warning', code: 'sheetCopyLimits', sourceId: ks.id, message: `「${ks.book.fileName}」のシート「${sh.original}」は、マクロシート・ダイアログシートのためコピーできませんでした。` });
      }
    }
    if (!copyable.length) continue;
    const cs = await openCopySource(sp, a);
    cs.rename = new Map(copyable.map((s) => [s.info.name.toLowerCase(), s.name!]));
    const op = await planSheetObjects(sp, takenTables);
    for (const n of op.tables.values()) takenTables.add(n.toLowerCase());
    const copier = new FormulaCopier({ sheets: cs.rename, tables: op.tables, allTables: op.allTables, sheetsWithTables: op.sheetsWithTables });
    const names = new SourceNames(sp.eng, (i) => sp.eng.sheets[i]?.name.toLowerCase() ?? null);
    names.analyse(copier);
    const dxfMap = new Map<number, number | null>();
    const dxf = (id: number): number | null => {
      let v = dxfMap.get(id);
      if (v === undefined) { v = a.baseStyles.importDxf(cs.styles, id, cs.themeRgb, cs.indexedRgb); dxfMap.set(id, v); }
      return v;
    };
    if (cs.dateShift) {
      const from = cs.dateShift > 0 ? '1904' : '1900', to = cs.dateShift > 0 ? '1900' : '1904';
      items.push({
        severity: 'warning', code: 'sheetCopyLimits', sourceId: ks.id,
        message: `「${ks.book.fileName}」は日付の基準（${from} 年方式）が基準ファイル（${to} 年方式）と異なるため、コピーしたシートの日付の値を ${to} 年方式に変換しました。数式の中に直接書いた日付の数値は変換されません。`,
      });
    }
    if (cs.valuesOnly) {
      items.push({
        severity: 'warning', code: 'sheetCopyLimits', sourceId: ks.id,
        message: `「${ks.book.fileName}」は値だけを読み込める形式（xls・csv など）のため、コピーしたシートには値と基本的な表示形式だけが入ります（数式・罫線・図などは含まれません）。`,
      });
    }
    for (const sh of copyable) {
      a.progress?.(0.9 + (0.08 * done++) / Math.max(1, total), `シートをコピー中…（${sh.name}）`);
      const index = nextIndex++;
      let relN = 0;
      const ids = { newRelId: () => `rId${++relN}` };
      let rels: Array<Rel & { path: string }> = [];
      let sheetXml: string | null = null;
      let job: SheetJob | null = null;
      let res: SheetCopy | null = null;
      let why = '';
      const chart = CHART_SHEETS.has(sh.info);
      try {
        sheetXml = await sp.eng.pkg.read(sh.info.path);
        if (sheetXml == null) why = 'シートのデータを読み込めませんでした。';
        else {
          rels = (await sp.eng.pkg.rels(sh.info.path)).map((r) => ({ ...r, path: r.external ? '' : resolveTarget(dirOf(sh.info.path), r.target) }));
          job = {
            src: cs, target, newName: sh.name!, names: names.view(sh.info.index), copier,
            dxf, guid: (old) => makeGuid(`${sh.name}\u0000${old}`),
            dynCm: sheetXml.indexOf(' cm="') >= 0 ? await eng.dynamicArrayCm() : null,
            hostLower: sh.info.name.toLowerCase(),
          };
          const r = chart ? prepareChartSheet(sheetXml) : await copyWorksheet(job, sheetXml, rels, ids);
          if ('problem' in r) why = `シートの内容に、そのまま書き戻せない部分があります（${r.problem}）。`;
          else res = r;
        }
      } catch (e) {
        // a damaged part of the source (inflate error, size guard ...) is a problem of this sheet; a defect of the copier is not hidden
        if (e instanceof PatchError && !e.internal) why = e.message;
        else throw e;
      }
      sp.eng.pkg.release(sh.info.path);
      // ---- where the new sheet lives ----
      let kind: 'worksheet' | 'chartsheet' = chart ? 'chartsheet' : 'worksheet';
      let path = freshPartName(pkg, kind === 'chartsheet' ? 'xl/chartsheets/sheet' : 'xl/worksheets/sheet', '.xml');
      let left: ObjectItem[] = [];
      if (res && job && sheetXml !== null) {
        const ctx = makeObjectsContext(sp, sh, sheetXml, rels, res, ids, job, path, index, a);
        await copySheetObjects(ctx);
        items.push(...takeObjectInfo(eng));
        res.xml = ctx.xml;
        res.rels = ctx.rels;
        left = res.left.filter((o) => !ctx.handled.has(o.kind));
        // a chart sheet without its chart is nothing: it becomes an empty sheet like any sheet that cannot be copied
        if (chart && !/<drawing\b/.test(res.xml)) { res = null; why = 'グラフを読み込めませんでした。'; }
      }
      if (!res || !job || sheetXml === null) {
        // A sheet that cannot be copied stays as an EMPTY sheet of the planned name: the formulas of the other copies that read it keep a valid
        // target (they read nothing instead of becoming #REF!), and the warning says what happened
        res = emptyCopy();
        kind = 'worksheet';
        path = freshPartName(pkg, 'xl/worksheets/sheet', '.xml');
        left = [];
        items.push({ severity: 'warning', code: 'sheetCopyLimits', sourceId: ks.id, message: `「${ks.book.fileName}」のシート「${sh.original}」はコピーできませんでした。${why}「${sh.name}」は空のシートになります。Excel で開いて保存し直してから、もう一度お試しください。` });
      }
      pkg.write(path, res.xml);
      if (res.rels.length) pkg.write(relsPathOf(path), relsXml(res.rels, path));
      await pkg.setContentTypeOverride(path, kind === 'chartsheet' ? CHARTSHEET_CT : WS_CT);
      const rid = await pkg.addRel(eng.wbPath, convertNamespaces(kind === 'chartsheet' ? CHARTSHEET_REL : REL_TYPE.worksheet, strict), path);
      sheetEls.push(`<sheet name="${escAttr(sh.name!)}" sheetId="${++sheetId}"${sh.state !== 'visible' ? ` state="${sh.state}"` : ''} r:id="${rid}"/>`);
      if (kind === 'worksheet') outNames.push(sh.name!);
      // ---- names ----
      const em = names.emit(sh.info.index, index, res.usedNames);
      nameEls.push(...em.xml);
      // ---- report ----
      const message = why ? '' : limitsMessage(ks, sh, res, left, em.dropped);
      if (message) items.push({ severity: 'warning', code: 'sheetCopyLimits', sourceId: ks.id, message });
      added.push({ sourceId: ks.id, file: ks.book.fileName, original: sh.original, name: sh.name!, suffixed: sh.suffixed, complete: !why && !message && !cs.valuesOnly && !cs.dateShift });
    }
  }

  // ---- workbook.xml: <sheets>, <definedNames> ----
  if (sheetEls.length) {
    let wb = eng.wbXml;
    const el = findFirst(wb, 'sheets');
    if (el && !el.selfClosing) wb = wb.slice(0, el.end - '</sheets>'.length) + sheetEls.join('') + '</sheets>' + wb.slice(el.end);
    if (nameEls.length) wb = addDefinedNames(wb, nameEls.join(''));
    eng.wbXml = wb;
    await updateAppSheets(eng, worksheetsBefore, eng.sheets.filter((s) => s.worksheet).map((s) => s.name).concat(outNames));
  }

  // ---- the summary (first in the list) ----
  const head: CheckItem[] = [];
  if (added.length) {
    head.push({ severity: 'info', code: 'sheetsCopied', message: `他のファイルのシート ${added.length} 枚をコピーしました（${added[0].name}${added.length > 1 ? ' ほか' : ''}）` });
    const renamed = added.filter((s) => s.suffixed);
    for (const s of renamed.slice(0, 5)) head.push({ severity: 'info', code: 'sheetRenamed', sourceId: s.sourceId, message: `シート名が重なったため「${s.name}」にしました。` });
    if (renamed.length > 5) head.push({ severity: 'info', code: 'sheetRenamed', message: `シート名が重なったため、名前に (1) などを付けたシートが他に ${renamed.length - 5} 枚あります。` });
  }
  return { added, items: head.concat(items), movedMergedFirst };
}

function emptyCopy(): SheetCopy {
  return {
    xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="18.75"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`,
    rels: [], usedNames: new Set(), failures: [], noCache: 0, left: [], vmCells: 0, xfOverflow: 0,
    dropped: { cf: 0, dv: 0, sparklines: 0, hyperlinks: 0, other: [] }, notes: [], formulaCount: 0,
  };
}

// ---------------------------------------------------------------------------------------------------------------------
// source context
// ---------------------------------------------------------------------------------------------------------------------

/** max digit width (pixels) of the workbook's default font (the font of cell format 0, theme fonts resolved): what a stored column width is measured in; undefined for a font whose width is not known */
function defaultFontMdw(styles: Styles, theme: string | null): number | undefined {
  const f = normalFontOf(styles, theme);
  return knownMdw(f.name, f.size);
}

async function openCopySource(sp: SourcePlan, a: AddSheetsArgs): Promise<CopySource> {
  const eng = sp.eng;
  const styles = await eng.styles();
  const srcTheme = await themeXml(eng);
  const src1904 = is1904(eng);
  return {
    ks: sp.ks,
    eng,
    styles,
    sst: await eng.sst(),
    themeRgb: sameTheme(srcTheme, a.baseTheme) ? null : themeColors(srcTheme),
    indexedRgb: samePalette(styles, a.baseStyles) ? null : styles.effectivePalette(),
    dateShift: src1904 === a.base1904 ? 0 : src1904 ? 1462 : -1462,
    valuesOnly: sp.ks.book.fidelity === 'values',
    xfCache: a.xfCaches?.get(sp.ks.id) ?? new Map(),
    mdw: defaultFontMdw(styles, srcTheme),
    rename: new Map(),
  };
}

// ---------------------------------------------------------------------------------------------------------------------
// the context of copySheetObjects
// ---------------------------------------------------------------------------------------------------------------------

/** worksheet children in schema order (CT_Worksheet) */
const WS_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

/** insert a worksheet child element where the schema wants it */
export function insertWorksheetElement(xml: string, tag: string, fragment: string): string {
  const at = WS_ORDER.indexOf(tag);
  const sd = /<\/sheetData>|<sheetData\b[^>]*\/>/.exec(xml);
  const from = sd ? sd.index + sd[0].length : 0;
  const tail = xml.slice(from);
  let pos = tail.lastIndexOf('</worksheet>');
  if (pos < 0) pos = tail.length;
  for (const el of splitTopAt(tail)) {
    if (WS_ORDER.indexOf(el.tag) > at) { pos = el.start; break; }
  }
  let out = xml.slice(0, from + pos) + fragment + xml.slice(from + pos);
  if (/\sr:id=/.test(fragment) && !/<worksheet\b[^>]*\sxmlns:r=/.test(out)) {
    out = out.replace(/<worksheet\b/, '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
  }
  return out;
}

function splitTopAt(tail: string): Array<{ tag: string; start: number }> {
  // splitTop returns the texts; their positions follow from the order (only whitespace / comments can lie between them)
  const out: Array<{ tag: string; start: number }> = [];
  let cursor = 0;
  for (const t of splitTop(tail)) {
    const at = tail.indexOf(t.text, cursor);
    if (at < 0) break;
    out.push({ tag: t.tag, start: at });
    cursor = at + t.text.length;
  }
  return out;
}

function makeObjectsContext(sp: SourcePlan, sh: PlannedSheet, srcXml: string, rels: Array<Rel & { path: string }>, res: SheetCopy, ids: { newRelId(): string }, job: SheetJob, path: string, index: number, a: AddSheetsArgs): SheetObjectsContext {
  const ctx: SheetObjectsContext = {
    source: { ks: sp.ks, eng: sp.eng, info: sh.info, xml: srcXml, rels, styles: job.src.styles, sst: job.src.sst },
    target: { eng: a.eng, name: sh.name!, index, path, relsPath: relsPathOf(path), styles: a.baseStyles, sst: a.bsst },
    xml: res.xml,
    rels: res.rels,
    newRelId: ids.newRelId,
    insertElement(tag: string, fragment: string) {
      ctx.xml = insertWorksheetElement(ctx.xml, tag, fragment);
    },
    present: res.left,
    handled: new Set<ObjectKind>(),
    note(text: string) { res.notes.push(text); },
    names: {
      rename: job.src.rename,
      copier: job.copier,
      rewrite: (f: string) => job.copier.rewrite(f, job.hostLower, job.names),
      nameView: job.names,
    },
    maps: {
      xf: (i: number) => {
        const hit = job.src.xfCache.get(i);
        if (hit !== undefined) return hit;
        const r = a.baseStyles.importXfLimited(job.src.styles, i, job.src.themeRgb, job.src.xfCache, job.src.indexedRgb, XF_MAX);
        return r < 0 ? 0 : r;
      },
      dxf: job.dxf,
      sst: (i: number) => {
        const rich = job.src.sst.richXml(i);
        return rich ? a.bsst.addRich(convertRich(rich, job.src, a.baseStyles)) : a.bsst.add(job.src.sst.text(i));
      },
      themeRgb: job.src.themeRgb,
      indexedRgb: job.src.indexedRgb,
    },
  };
  return ctx;
}

// ---------------------------------------------------------------------------------------------------------------------
// report
// ---------------------------------------------------------------------------------------------------------------------

function limitsMessage(ks: KeySource, sh: PlannedSheet, res: SheetCopy, left: ObjectItem[], droppedNames: string[]): string {
  const parts: string[] = [];
  if (left.length) parts.push(`次のものはコピーできませんでした: ${left.map((o) => o.label).join('、')}`);
  for (const f of res.failures) {
    parts.push(`${FAIL_JA[f.reason]} ${f.count} 件は値にしました（${f.examples.join('、')}）`);
  }
  if (res.noCache) parts.push(`そのうち ${res.noCache} 件は元のファイルに計算結果が保存されていないため空欄になりました`);
  const rules: string[] = [];
  if (res.dropped.cf) rules.push(`条件付き書式 ${res.dropped.cf} 個`);
  if (res.dropped.dv) rules.push(`入力規則 ${res.dropped.dv} 個`);
  if (res.dropped.sparklines) rules.push(`スパークライン ${res.dropped.sparklines} 個`);
  if (res.dropped.hyperlinks) rules.push(`ハイパーリンク ${res.dropped.hyperlinks} 個`);
  if (rules.length) parts.push(`${rules.join('、')}は外しました`);
  if (res.dropped.other.length) parts.push(`未対応の設定 ${res.dropped.other.length} 件は含めませんでした`);
  if (res.vmCells) parts.push(`セル内の特殊なデータ（画像・データ型など）${res.vmCells} 個は通常の値にしました`);
  if (res.xfOverflow) parts.push(`書式の種類が Excel の上限（${XF_MAX.toLocaleString('ja-JP')}）に達したため、${res.xfOverflow} 個のセルの書式を標準にしました`);
  if (droppedNames.length) parts.push(`コピーできなかった名前: ${droppedNames.slice(0, 5).join('、')}${droppedNames.length > 5 ? ' ほか' : ''}`);
  parts.push(...res.notes);
  if (!parts.length) return '';
  return `シート「${sh.name}」（${ks.book.fileName} の「${sh.original}」）: ${parts.join('。')}。`;
}

// ---------------------------------------------------------------------------------------------------------------------
// workbook-level helpers
// ---------------------------------------------------------------------------------------------------------------------

function maxSheetId(wb: string): number {
  let m = 0;
  for (const s of (findFirst(wb, 'sheets')?.inner ?? '').matchAll(/<sheet\b[^>]*>/g)) m = Math.max(m, Number(getAttr(s[0], 'sheetId')) || 0);
  return m;
}

function freshPartName(pkg: Engine['pkg'], prefix: string, suffix: string): string {
  for (let n = 1; ; n++) {
    const p = `${prefix}${n}${suffix}`;
    if (!pkg.has(p)) return p;
  }
}

function relsXml(rels: NewRel[], ownerPath: string): string {
  const dir = dirOf(ownerPath);
  const body = rels.map((r) => r.external
    ? `<Relationship Id="${r.id}" Type="${escAttr(r.type)}" Target="${escAttr(r.target)}" TargetMode="External"/>`
    : `<Relationship Id="${r.id}" Type="${escAttr(r.type)}" Target="${escAttr(relativeTarget(dir, r.target))}"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
}

/** <definedNames> gets the elements (created in its schema position when the workbook has none) */
function addDefinedNames(wb: string, elements: string): string {
  const dn = findFirst(wb, 'definedNames');
  if (dn && !dn.selfClosing) return wb.slice(0, dn.end - '</definedNames>'.length) + elements + '</definedNames>' + wb.slice(dn.end);
  const el = `<definedNames>${elements}</definedNames>`;
  if (dn) return wb.slice(0, dn.start) + el + wb.slice(dn.end);
  for (const tag of ['externalReferences', 'functionGroups', 'sheets']) {
    const m = new RegExp(`</${tag}>|<${tag}\\b[^>]*/>`).exec(wb);
    if (m) return wb.slice(0, m.index + m[0].length) + el + wb.slice(m.index + m[0].length);
  }
  return wb;
}

/**
 * The deterministic GUID a rule link gets in a copy (x14 conditional formats: {XXXXXXXX-XXXX-4XXX-AXXX-XXXXXXXXXXXX}); the same input
 * always gives the same output so that the copy is reproducible.
 */
export function makeGuid(seed: string): string {
  let h1 = 1779033703, h2 = 3144134277, h3 = 1013904242, h4 = 2773480762;
  for (let i = 0; i < seed.length; i++) {
    const k = seed.charCodeAt(i);
    h1 = h2 ^ Math.imul(h1 ^ k, 597399067);
    h2 = h3 ^ Math.imul(h2 ^ k, 2869860233);
    h3 = h4 ^ Math.imul(h3 ^ k, 951274213);
    h4 = h1 ^ Math.imul(h4 ^ k, 2716044179);
  }
  h1 = Math.imul(h3 ^ (h1 >>> 18), 597399067);
  h2 = Math.imul(h4 ^ (h2 >>> 22), 2869860233);
  h3 = Math.imul(h1 ^ (h3 >>> 17), 951274213);
  h4 = Math.imul(h2 ^ (h4 >>> 19), 2716044179);
  const hex = (n: number, len: number) => (n >>> 0).toString(16).toUpperCase().padStart(8, '0').slice(0, len);
  return `{${hex(h1 ^ h2, 8)}-${hex(h2 ^ h3, 4)}-4${hex(h3 ^ h4, 3)}-A${hex(h4 ^ h1, 3)}-${hex(h1 ^ h3, 8)}${hex(h2 ^ h4, 4)}}`;
}

// ---------------------------------------------------------------------------------------------------------------------
// the merged sheet moves to the first tab
// ---------------------------------------------------------------------------------------------------------------------

/**
 * The merged sheet is the first sheet of the output (the user's request). When the base workbook has it elsewhere, it moves to the front:
 * <sheets> order, localSheetId of every defined name, the first tab shown, the selected tab (only the merged sheet stays selected),
 * Engine.sheets. Nothing happens when it is first already.
 */
async function moveMergedFirst(eng: Engine, merged: SheetState): Promise<boolean> {
  const info = merged.info;
  const at = eng.sheets.indexOf(info);
  if (at <= 0) return false;
  let wb = eng.wbXml;
  const sheets = findFirst(wb, 'sheets');
  if (!sheets || sheets.selfClosing) return false;
  const items = [...sheets.inner.matchAll(/<sheet\b[^>]*?\/>|<sheet\b[^>]*?>[\s\S]*?<\/sheet>/g)].map((m) => m[0]);
  if (items.length !== eng.sheets.length) return false;
  const order = [at, ...items.map((_x, i) => i).filter((i) => i !== at)];
  wb = wb.slice(0, sheets.start) + sheets.open + order.map((i) => items[i]).join('') + '</sheets>' + wb.slice(sheets.end);
  // defined names: localSheetId follows the sheets
  const newIndex = new Map<number, number>(order.map((old, i) => [old, i]));
  wb = wb.replace(/<definedName\b[^>]*>/g, (o) => {
    const ls = getAttr(o, 'localSheetId');
    if (ls === undefined) return o;
    const n = newIndex.get(Number(ls));
    return n === undefined ? o : setAttr(o, 'localSheetId', String(n));
  });
  // the first tab is shown and selected
  wb = wb.replace(/<workbookView\b[^>]*>/g, (o) => setAttr(setAttr(o, 'activeTab', null), 'firstSheet', null));
  eng.wbXml = wb;
  const reordered = order.map((i) => eng.sheets[i]);
  reordered.forEach((s, i) => { s.index = i; });
  eng.sheets = reordered;
  // only the merged sheet is selected (two selected tabs = grouped sheets: edits would hit both)
  merged.model.head = merged.model.head.replace(/<sheetView\b[^>]*>/, (o) => setAttr(o, 'tabSelected', '1'));
  merged.model.dirty = true;
  for (const s of reordered) {
    if (s === info || !s.path) continue;
    const st = eng.states.get(s.path);
    if (st) {
      st.model.head = st.model.head.replace(/<sheetView\b[^>]*>/g, (o) => setAttr(o, 'tabSelected', null));
      st.model.dirty = true;
      continue;
    }
    const x = await eng.pkg.read(s.path);
    if (x == null || !/tabSelected="(1|true)"/.test(x.slice(0, 20000))) continue;
    eng.pkg.write(s.path, x.replace(/<sheetView\b[^>]*>/g, (o) => setAttr(o, 'tabSelected', null)));
  }
  return true;
}

// ---------------------------------------------------------------------------------------------------------------------
// docProps/app.xml
// ---------------------------------------------------------------------------------------------------------------------

const VT_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';

/**
 * HeadingPairs / TitlesOfParts list the worksheets (first group) and the named ranges: the first group becomes `after`; when the part does
 * not look like that (or the first group is not the list of worksheets `before`), both lists are removed (they are optional, Excel rebuilds
 * them on save).
 */
async function updateAppSheets(eng: Engine, before: string[], after: string[]): Promise<void> {
  const pkg = eng.pkg;
  const rel = (await pkg.relTargets('')).find((r) => /\/extended-properties$/.test(r.type));
  if (!rel) return;
  const xml = await pkg.read(rel.path);
  if (!xml) return;
  const hp = /<HeadingPairs>([\s\S]*?)<\/HeadingPairs>/.exec(xml);
  const tp = /<TitlesOfParts>([\s\S]*?)<\/TitlesOfParts>/.exec(xml);
  if (!hp && !tp) return;
  const strip = () => pkg.write(rel.path, xml.replace(/<HeadingPairs>[\s\S]*?<\/HeadingPairs>/, '').replace(/<TitlesOfParts>[\s\S]*?<\/TitlesOfParts>/, ''));
  const prefix = new RegExp(`xmlns:([\\w.-]+)="${VT_NS.replace(/\./g, '\\.')}"`).exec(xml)?.[1];
  if (!hp || !tp || !prefix) { strip(); return; }
  const vt = prefix + ':';
  const lp = new RegExp(`<${vt}lpstr>([\\s\\S]*?)</${vt}lpstr>`, 'g');
  const i4 = new RegExp(`<${vt}i4>\\s*(\\d+)\\s*</${vt}i4>`, 'g');
  const labels = [...hp[1].matchAll(lp)].map((m) => unesc(m[1]));
  const counts = [...hp[1].matchAll(i4)].map((m) => Number(m[1]));
  const titles = [...tp[1].matchAll(lp)].map((m) => unesc(m[1]));
  if (!labels.length || labels.length !== counts.length || counts.reduce((x, y) => x + y, 0) !== titles.length || counts[0] !== before.length) { strip(); return; }
  const group0 = titles.slice(0, counts[0]);
  if (group0.map((t) => t.toLowerCase()).sort().join('\u0000') !== before.map((t) => t.toLowerCase()).sort().join('\u0000')) { strip(); return; }
  const newCounts = counts.slice();
  newCounts[0] = after.length;
  const all = after.concat(titles.slice(counts[0]));
  const variants = labels.map((l, i) => `<${vt}variant><${vt}lpstr>${esc(l)}</${vt}lpstr></${vt}variant><${vt}variant><${vt}i4>${newCounts[i]}</${vt}i4></${vt}variant>`).join('');
  const heading = `<HeadingPairs><${vt}vector size="${labels.length * 2}" baseType="variant">${variants}</${vt}vector></HeadingPairs>`;
  const parts = `<TitlesOfParts><${vt}vector size="${all.length}" baseType="lpstr">${all.map((t) => `<${vt}lpstr>${esc(t)}</${vt}lpstr>`).join('')}</${vt}vector></TitlesOfParts>`;
  pkg.write(rel.path, xml.replace(/<HeadingPairs>[\s\S]*?<\/HeadingPairs>/, heading).replace(/<TitlesOfParts>[\s\S]*?<\/TitlesOfParts>/, parts));
}
