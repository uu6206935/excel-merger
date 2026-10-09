/**
 * 警告 section: fidelity notes about the input books and the merge (pure).
 */
import type { MergeReport, Severity } from '../core/types';
import type { FileEntry } from './types';
import { formulaReportViews, formulaWarningText, isFormulaImportItem } from './formulaReport';
import { FORMAT_LABEL } from './helpers';
import { checkItemWhere } from './keysPlan';

export interface Note {
  key: string;
  severity: Severity;
  text: string;
}

const MAX_SHEET_NAMES = 3;

/** 「出力は結合した 1 シートのみです（他のシート N 枚「集計」「マスタ」… は含めません）」 */
export function droppedSheetsLine(names: readonly string[]): string {
  const shown = names
    .slice(0, MAX_SHEET_NAMES)
    .map((n) => `「${n.length > 16 ? `${n.slice(0, 15)}…` : n}」`)
    .join('');
  return `出力は結合した 1 シートのみです（他のシート ${names.length} 枚${shown}${names.length > MAX_SHEET_NAMES ? '…' : ''} は含めません）`;
}

/** 「他のシートを参照していた数式 N 個は値にしました」 */
export function formulasDetachedLine(count: number): string {
  return `他のシートを参照していた数式 ${count} 個は値にしました`;
}

/**
 * Names of the sheets of the base book that are not the data sheet = what 「他のシートは出力に含めない」 would leave out
 * (a prediction before the merge; after it MergeReport.droppedSheets is the truth). [] without a base book.
 */
export function otherSheetNames(base: Pick<FileEntry, 'book' | 'sheet' | 'region'> | undefined): string[] {
  if (!base?.book) return [];
  const data = base.region?.sheet ?? base.sheet;
  return base.book.sheets.map((s) => s.name).filter((n) => n !== data);
}

/**
 * @param excelAvailable server says Excel is installed (Excel本体 engine can be used)
 * @param extra.labels source id -> file label (rows of check items: 「02_面接評価 の 12行目」)
 * @param extra.exportWarnings warnings of the last export (Excel本体 runner, Japanese)
 * @param extra.excluded ids of the files that take no part in the merge (no key selected)
 * @param extra.merged the merge is done (the wording changes from 「取り込まれません」 to 「取り込んでいません」)
 * @param extra.dropOtherSheets the output holds only the merged sheet (the default). The objects of the base book (charts,
 *        pivot tables, slicers, images, comments …) are counted for the WHOLE book and cannot be told apart by sheet, so while
 *        the other sheets are dropped (and there are some) the notes that promise their preservation are not shown: they may
 *        live on a dropped sheet. The report says what was dropped (droppedSheets / formulasDetached).
 */
export function fidelityNotes(
  files: readonly FileEntry[],
  baseId: string | null,
  excelAvailable: boolean,
  report?: MergeReport | null,
  extra?: { labels?: Record<string, string>; exportWarnings?: readonly string[]; excluded?: readonly string[]; merged?: boolean; dropOtherSheets?: boolean },
): Note[] {
  const out: Note[] = [];
  // files without a key: listed first, they are the thing the user can still act on
  for (const id of extra?.excluded ?? []) {
    const name = extra?.labels?.[id] ?? files.find((f) => f.id === id)?.fileName ?? id;
    out.push({
      key: `excluded-${id}`,
      severity: 'warning',
      text: extra?.merged
        ? `「${name}」は結合キーが見つからないため結合の対象外です（列は取り込んでいません）。キーを選んでからもう一度結合すると取り込めます。`
        : `「${name}」は結合キーが見つからないため結合の対象外です（列は取り込まれません）。ファイルカードの「キーを選ぶ…」でキーを選ぶと結合に含まれます。`,
    });
  }
  const base = files.find((f) => f.id === baseId && f.book) ?? files.find((f) => f.book);
  const engineHint = excelAvailable ? 'Excel本体モードで確実に保持されます' : 'Excel本体モード（Excel のある PC でアプリを起動）で出力すると確実です';

  if (base?.book) {
    const b = base.book;
    const ft = b.features;
    const add = (key: string, severity: Severity, text: string) => out.push({ key: `base-${key}`, severity, text });
    if (b.fidelity === 'values') {
      add('values', 'warning', `基準ファイルが ${FORMAT_LABEL[b.format]} 形式のため、書式・数式のない新しいブックとして出力されます`);
    }
    // Single-sheet output: with other sheets in the base book the objects may sit on a sheet that is dropped
    // (after the merge the report is the truth: nothing dropped = every object is still there; before it: any other sheet in the book)
    const sheetsDropped = extra?.merged && report ? (report.droppedSheets?.length ?? 0) > 0 : b.sheets.length > 1;
    const mayBeDropped = extra?.dropOtherSheets === true && sheetsDropped;
    if (ft.charts > 0 && !mayBeDropped) add('charts', 'info', `グラフ ${ft.charts} 個: ${engineHint}`);
    if (ft.pivotTables > 0 && !mayBeDropped) add('pivots', 'warning', `ピボットテーブル ${ft.pivotTables} 個: 元データ範囲は ${engineHint}（出力後に「更新」してください）`);
    if (ft.vba) add('vba', 'info', 'マクロ（VBA）を含むため .xlsm で出力します（マクロ内のセル番地は自動では変わりません）');
    if (ft.slicers > 0 && !mayBeDropped) add('slicers', 'warning', `スライサー ${ft.slicers} 個: ${engineHint}`);
    if (ft.sparklineGroups > 0 && !mayBeDropped) add('sparklines', 'info', `スパークライン ${ft.sparklineGroups} 組: ${engineHint}`);
    if (ft.externalLinks > 0) add('ext', 'warning', `外部ブックへのリンク ${ft.externalLinks} 件: リンク先のファイルは変更されません`);
    if (ft.formControls > 0 && !mayBeDropped) add('controls', 'warning', `フォームコントロール ${ft.formControls} 個: ${engineHint}`);
    if (ft.threadedComments > 0 && !mayBeDropped) add('threaded', 'info', `スレッド形式のコメント ${ft.threadedComments} 件は保持されます`);
    const pictures = ft.images + Math.max(0, ft.drawings - ft.charts);
    if (ft.images > 0 && !mayBeDropped) add('images', 'info', `図・画像 ${pictures > 0 ? pictures : ft.images} 個: 列の移動に合わせて位置を調整します`);
    b.warnings.forEach((w, i) => add(`w${i}`, 'warning', `${base.fileName}: ${w}`));
  }

  for (const f of files) {
    if (!f.book || f === base) continue;
    const b = f.book;
    // the reader's own warning often says the same (「xls形式のファイルは値…だけを読み込みました」): one note is enough
    const readerSaysValues = b.warnings.some((w) => /値/.test(w));
    if (b.fidelity === 'values' && !readerSaysValues) {
      out.push({ key: `${f.id}-values`, severity: 'info', text: `${f.fileName}: ${FORMAT_LABEL[b.format]} 形式のため値のみ取り込みます（書式・数式は引き継がれません）` });
    }
    const ft = b.features;
    const objs = ft.charts + ft.pivotTables + ft.images + ft.slicers;
    if (objs > 0) {
      out.push({ key: `${f.id}-objs`, severity: 'info', text: `${f.fileName}: グラフ・図・ピボットは取り込まれません（取り込むのはセルの値と書式です）` });
    }
    b.warnings.forEach((w, i) => out.push({ key: `${f.id}-w${i}`, severity: 'warning', text: `${f.fileName}: ${w}` }));
  }

  if (report) {
    if (report.droppedSheets && report.droppedSheets.length > 0) {
      out.push({ key: 'dropped-sheets', severity: 'info', text: droppedSheetsLine(report.droppedSheets) });
    }
    if (report.formulasDetached && report.formulasDetached > 0) {
      out.push({ key: 'formulas-detached', severity: 'info', text: formulasDetachedLine(report.formulasDetached) });
    }
    // The formula import reports per file (formulaStats): one short entry for the files whose formulas did not all stay formulas — the
    // reasons, counts and example rows are in the side panel's 「数式の取り込み」, so the report's own formula items are not repeated here.
    // Reports without formulaStats (older / mock) keep the generic line.
    const hasStats = Array.isArray(report.formulaStats);
    if (hasStats) {
      const nameOf = (sid: string, label: string) => files.find((f) => f.id === sid)?.fileName ?? label;
      for (const v of formulaReportViews(report, nameOf)) {
        const text = formulaWarningText(v);
        if (text) out.push({ key: `formulas-${v.sourceId}`, severity: 'warning', text });
      }
    } else if (report.formulasToValues > 0) {
      out.push({
        key: 'toValues',
        severity: 'warning',
        text: `取り込んだ数式 ${report.formulasToValues} 件を値に変換しました（参照先の列が取り込まれていないため）`,
      });
    }
    report.items.forEach((it, i) => {
      if (hasStats && isFormulaImportItem(it) && it.code !== 'formulasPlaceholderCache') return;
      const where = checkItemWhere(it, extra?.labels ?? {});
      out.push({ key: `r${i}`, severity: it.severity, text: where ? `${it.message}（${where}）` : it.message });
    });
  }
  (extra?.exportWarnings ?? []).forEach((w, i) => out.push({ key: `export-w${i}`, severity: 'warning', text: `出力: ${w}` }));
  return out;
}

export function noteCounts(notes: readonly Note[]): { error: number; warning: number; info: number } {
  const c = { error: 0, warning: 0, info: 0 };
  for (const n of notes) c[n.severity]++;
  return c;
}
