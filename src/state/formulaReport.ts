/**
 * 「数式の取り込み」 (pure, no React): what happened to the formulas of the imported files (2nd file and later), in plain Japanese.
 *
 * Sources: MergeReport.formulaStats (counts per file: kept / toValues / empty / reasons) and the report items of the formula import
 * (codes in FORMULA_ITEM_CODES: their messages are already plain Japanese and carry example rows of the merged sheet).
 * Used by the side panel block (FormulaReport), the merge toast, the 警告 list and the 数式→値 group of the check panel.
 *
 * Counting: `reasons` of the contract holds ONE count per reason for every formula that was not kept, i.e. the formulas turned into
 * values AND the ones left empty together (sum(reasons) = toValues + empty). A per-reason breakdown of the values alone is therefore
 * only exact when no cell ended up empty; otherwise the combined breakdown is labelled as such.
 */
import type { CheckItem, DocColumn, MergeReport, Severity, SourceBook, TableRegion } from '../core/types';
import type { CheckRow } from './keysPlan';

/** the report-item codes written by the formula import (premerge) */
export const FORMULA_ITEM_CODES: readonly string[] = [
  'formulasKept',
  'formulaRangesWidened',
  'formulasRecalc',
  'formulasPlaceholderCache',
  'formulasToValues',
  'formulasNoCache',
];

export function isFormulaImportItem(it: Pick<CheckItem, 'code'>): boolean {
  return FORMULA_ITEM_CODES.includes(it.code);
}

/** short noun phrases for the breakdown 「内訳: 他のシートを参照 8、表の外のセル 4」 (reason codes of MergeReport.formulaStats) */
export const FORMULA_REASON_LABEL: Record<string, string> = {
  otherSheet: '他のシートを参照',
  externalCell: '表の外のセル',
  rowRelation: '前後の行との関係',
  notImported: '取り込まない列を参照',
  unresolvedName: '名前の定義',
  unresolvedTable: '表の列名での参照',
  positional: 'セルの位置で決まる式',
  array: '配列数式',
  unsafe: '外部のブック',
  unsupported: '対応していない書き方',
  policy: '設定による',
};

/** one explaining sentence per reason (details list) */
export const FORMULA_REASON_NOTE: Record<string, string> = {
  otherSheet: '元のファイルの別のシートを見ている数式です。結合後の表にはそのシートがないため、計算結果の値にしました。',
  externalCell: '表の外にあるセル（税率など）を見ている数式です。結合後は同じ場所にないため、計算結果の値にしました。',
  rowRelation: '前の行との差や、途中までの合計など。結合後は行の並びが変わるため、同じ範囲を指せません。行の並びが元のファイルと同じなら数式のまま残せます。',
  notImported: '取り込んでいない列を見ている数式です。その列も取り込むと、数式のまま残せることがあります。',
  unresolvedName: '名前（定義名）を使った数式で、置き換えられませんでした。',
  unresolvedTable: 'テーブルの列名（[@列名] など）を使った数式で、置き換えられませんでした。',
  positional: 'ROW・COLUMN・OFFSET・INDIRECT など、セルの位置によって結果が変わる数式です。',
  array: '複数のセルにまたがる配列数式（スピルを含む）です。',
  unsafe: '外部のブックなどを見ている数式です。安全のため数式にはしていません。',
  unsupported: 'このツールでは付け替えられない書き方の数式です。',
  policy: '設定により、計算結果の値にしました。',
};

const REASON_ORDER = Object.keys(FORMULA_REASON_LABEL);

export const FORMULA_EMPTY_HINT = '元のExcelで一度開いて保存し直すと取り込めることがあります';
export const FORMULA_EMPTY_WHY = '元のファイルに計算結果がありません';
export const FORMULA_RECALC_LINE = '一部の計算結果は、Excel で開いたときに表示されます（この画面では式を灰色で表示します）';

export interface FormulaReasonView {
  code: string;
  label: string;
  note: string;
  count: number;
}

export interface FormulaDetailItem {
  key: string;
  code: string;
  severity: Severity;
  /** the report item's own message (plain Japanese) */
  text: string;
  /** example rows of the merged sheet (≤ 20) */
  rows: number[];
  /** which cell the rows point at: 'plain' = a cell that holds a value / nothing (the formula was not kept), 'formula' = a kept formula */
  cell: 'plain' | 'formula';
}

export interface FormulaFileView {
  sourceId: string;
  /** file name as shown (with the extension when the file is known) */
  name: string;
  kept: number;
  toValues: number;
  empty: number;
  /** 'ok' = everything was kept (info only), 'note' = some formulas became values, 'warn' = cells ended up EMPTY (amber) */
  tone: 'ok' | 'note' | 'warn';
  keptLine: string | null;
  valuesLine: string | null;
  emptyLine: string | null;
  /** shown when kept formulas have no stored result (the report item 'formulasRecalc'): Excel calculates them when the file is opened */
  recalcLine: string | null;
  /** 「02_面接評価.xlsx: 数式 147 件を数式のまま取り込みました／12 件は値にしました（内訳: …）／空白 3 件（…）」 */
  sentence: string;
  /** what to try when cells are empty (null otherwise) */
  hint: string | null;
  reasons: FormulaReasonView[];
  /** true when the reason counts include formulas left empty (see the file comment) */
  reasonsIncludeEmpty: boolean;
  details: FormulaDetailItem[];
}

const num = (n: number) => n.toLocaleString('ja-JP');

/** 「他のシートを参照 8、表の外のセル 4」 */
export function reasonBreakdown(reasons: readonly FormulaReasonView[]): string {
  return reasons.map((r) => `${r.label} ${num(r.count)}`).join('、');
}

function reasonViews(reasons: Record<string, number>): FormulaReasonView[] {
  const out: FormulaReasonView[] = [];
  for (const [code, count] of Object.entries(reasons ?? {})) {
    if (!(count > 0)) continue;
    out.push({ code, label: FORMULA_REASON_LABEL[code] ?? '対応していない書き方', note: FORMULA_REASON_NOTE[code] ?? FORMULA_REASON_NOTE.unsupported, count });
  }
  const rank = (c: string) => {
    const i = REASON_ORDER.indexOf(c);
    return i < 0 ? REASON_ORDER.length : i;
  };
  return out.sort((a, b) => b.count - a.count || rank(a.code) - rank(b.code));
}

function detailItems(report: MergeReport, sourceId: string): FormulaDetailItem[] {
  const out: FormulaDetailItem[] = [];
  report.items.forEach((it, i) => {
    if (it.sourceId !== sourceId || !isFormulaImportItem(it) || it.code === 'formulasKept') return;
    out.push({
      key: `fi${i}`,
      code: it.code,
      severity: it.severity,
      text: it.message,
      rows: (it.rows ?? []).filter((r) => Number.isInteger(r) && r > 0),
      cell: it.code === 'formulasToValues' || it.code === 'formulasNoCache' ? 'plain' : 'formula',
    });
  });
  return out;
}

/**
 * One view per file that had formulas in its imported columns (MergeReport.formulaStats has an entry only for those). [] when the report
 * has no formulaStats (older / mock reports) or no file had formulas.
 * @param nameOf shown name of a file (default: the report's label)
 */
export function formulaReportViews(report: MergeReport | null | undefined, nameOf?: (sourceId: string, label: string) => string | undefined): FormulaFileView[] {
  const stats = report?.formulaStats;
  if (!report || !Array.isArray(stats)) return [];
  const views: FormulaFileView[] = [];
  for (const s of stats) {
    const kept = Math.max(0, s.kept | 0);
    const toValues = Math.max(0, s.toValues | 0);
    const empty = Math.max(0, s.empty | 0);
    if (kept + toValues + empty === 0) continue;
    const name = nameOf?.(s.sourceId, s.label) || s.label || s.sourceId;
    const reasons = reasonViews(s.reasons);
    const keptLine = kept > 0 ? `数式 ${num(kept)} 件を数式のまま取り込みました` : null;
    // the breakdown is exact only when nothing ended up empty (the reason counts cover values + empty together)
    const inline = empty === 0 && reasons.length > 0 ? `（内訳: ${reasonBreakdown(reasons)}）` : '';
    const valuesLine = toValues > 0 ? `${kept > 0 ? '' : '数式 '}${num(toValues)} 件は値にしました${inline}` : null;
    const emptyLine = empty > 0 ? `空白 ${num(empty)} 件（${FORMULA_EMPTY_WHY}）` : null;
    const recalcLine = report.items.some((it) => it.sourceId === s.sourceId && it.code === 'formulasRecalc') ? FORMULA_RECALC_LINE : null;
    const lines = [keptLine, valuesLine, emptyLine].filter((l): l is string => !!l);
    views.push({
      sourceId: s.sourceId,
      name,
      kept,
      toValues,
      empty,
      tone: empty > 0 ? 'warn' : toValues > 0 ? 'note' : 'ok',
      keptLine,
      valuesLine,
      emptyLine,
      recalcLine,
      sentence: `${name}: ${lines.join('／')}`,
      hint: empty > 0 ? FORMULA_EMPTY_HINT : null,
      reasons,
      reasonsIncludeEmpty: empty > 0,
      details: detailItems(report, s.sourceId),
    });
  }
  return views;
}

export interface FormulaTotals {
  kept: number;
  toValues: number;
  empty: number;
  files: number;
}

export function formulaTotals(views: readonly FormulaFileView[]): FormulaTotals {
  const t: FormulaTotals = { kept: 0, toValues: 0, empty: 0, files: views.length };
  for (const v of views) {
    t.kept += v.kept;
    t.toValues += v.toValues;
    t.empty += v.empty;
  }
  return t;
}

/** The count parts of the merge toast (「数式 147 件は数式のまま」 …); [] without formulas. */
export function formulaToastParts(views: readonly FormulaFileView[]): string[] {
  if (!views.length) return [];
  const t = formulaTotals(views);
  const parts: string[] = [];
  if (t.kept > 0) parts.push(`数式 ${num(t.kept)} 件は数式のまま`);
  if (t.toValues > 0) parts.push(`値にした数式 ${num(t.toValues)} 件`);
  if (t.empty > 0) parts.push(`空白になった数式 ${num(t.empty)} 件`);
  return parts;
}

/** The lines of the toast detail: one sentence per file (「02_面接評価.xlsx: 数式 147 件を数式のまま取り込みました／…」). */
export function formulaToastLines(views: readonly FormulaFileView[]): string[] {
  const lines = views.map((v) => v.sentence);
  if (views.some((v) => v.empty > 0)) lines.push(FORMULA_EMPTY_HINT);
  return lines;
}

/**
 * The 警告 entry of a file whose formulas did not all stay formulas (null = all kept: that is information for the block, no warning).
 * The details (reasons, example rows) are in the 「数式の取り込み」 block.
 */
export function formulaWarningText(v: FormulaFileView): string | null {
  const lines = [v.valuesLine, v.emptyLine].filter((l): l is string => !!l);
  if (!lines.length) return null;
  return `${v.name}: ${lines.join('／')}。${v.hint ? `${v.hint}。` : ''}詳しくは「数式の取り込み」を見てください。`;
}

/**
 * Rows of the 数式→値 group of the check panel: the example rows of the report items (formulas turned into values / left empty).
 * `row` is the Excel row of the merged sheet (click = jump).
 */
export function formulaCheckRows(report: MergeReport | null | undefined, nameOf?: (sourceId: string, label: string) => string | undefined): CheckRow[] {
  if (!report || !Array.isArray(report.formulaStats)) return [];
  const labels = new Map(report.formulaStats.map((s) => [s.sourceId, s.label]));
  const out: CheckRow[] = [];
  const seen = new Set<string>();
  report.items.forEach((it, i) => {
    if (it.code !== 'formulasToValues' && it.code !== 'formulasNoCache') return;
    const sid = it.sourceId ?? '';
    const name = nameOf?.(sid, labels.get(sid) ?? sid) || labels.get(sid) || sid;
    const sub = it.code === 'formulasNoCache' ? `${name}: 空白になった数式` : `${name}: 値にした数式`;
    for (const row of it.rows ?? []) {
      const k = `${it.code}:${sid}:${row}`;
      if (seen.has(k)) continue;
      seen.add(k);
      out.push({ key: `fv${i}-${row}`, row, text: `${row}行目`, sub, tone: it.code === 'formulasNoCache' ? 'warning' : 'info' });
    }
  });
  return out;
}

// ---------------------------------------------------------------------------------------------
// where the formulas are (preview hint, jump target)

const colCache = new WeakMap<SourceBook, Map<string, Map<number, number>>>();

/**
 * Per source column of the data table: how many cells hold a formula (data rows only; empty map for a 'values' book). Cached per book.
 * Used for the hint under 「追加される列」 and to find the cell an example row of the report points at.
 */
export function formulaColumnsOf(book: SourceBook | undefined, region: TableRegion | undefined): Map<number, number> {
  const out = new Map<number, number>();
  if (!book || !region || book.fidelity === 'values') return out;
  const key = `${region.sheet}|${region.headerRow}|${region.firstDataRow}|${region.lastDataRow}|${region.firstCol}|${region.lastCol}`;
  let perBook = colCache.get(book);
  const hit = perBook?.get(key);
  if (hit) return hit;
  const sheet = book.sheets.find((s) => s.name === region.sheet);
  if (sheet) {
    for (let r = region.firstDataRow; r <= region.lastDataRow; r++) {
      const row = sheet.rows[r - 1];
      if (!row) continue;
      for (let c = region.firstCol; c <= region.lastCol; c++) {
        if (row[c - 1]?.f) out.set(c, (out.get(c) ?? 0) + 1);
      }
    }
  }
  if (!perBook) {
    perBook = new Map();
    colCache.set(book, perBook);
  }
  perBook.set(key, out);
  return out;
}

/** The hint under 「追加される列」 (MergePreview). */
export const PREVIEW_FORMULA_HINT = '数式を含む列があります。結合後も数式のまま残るよう付け替えます';

/**
 * Column id of the cell to show for an example row of a file: among the file's columns that held formulas in the source
 * (`srcFormulaCols` = source column numbers) the first one whose merged cell is not a formula ('plain': the formula was turned into
 * a value / left empty) or is one ('formula': a kept formula); else the file's first formula column; else the file's first column.
 * null when the file has no column in the table.
 */
export function exampleCellColumn(
  columns: readonly Pick<DocColumn, 'id' | 'origin' | 'cells'>[],
  sourceId: string,
  row: number,
  srcFormulaCols: ReadonlySet<number>,
  want: 'plain' | 'formula',
): string | null {
  const mine = columns.filter((c) => c.origin.kind === 'import' && c.origin.sourceId === sourceId);
  if (!mine.length) return null;
  const withFormulas = mine.filter((c) => c.origin.kind === 'import' && srcFormulaCols.has(c.origin.col));
  const fits = withFormulas.find((c) => !!c.cells[row - 1]?.f === (want === 'formula'));
  return (fits ?? withFormulas[0] ?? mine[0]).id;
}
