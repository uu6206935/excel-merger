/**
 * Round 9 (plan A): the two sheet options and the sheet list of the output (pure; no React, no module calls).
 *
 *   「他のシートも出力に含める」        includeOtherSheets  = keep the BASE file's other sheets (PremergeOptions.dropOtherSheets = !it)
 *   「他のファイルのシートも含める」    includeSourceSheets = copy every sheet of the 2nd, 3rd … files (PremergeOptions.includeSourceSheets)
 *
 * The two checkboxes are coupled (the user wants either "only the merged sheet" or "all sheets"):
 *   - turning the first ON turns the second ON (the user may switch it off again);
 *   - turning the second ON while the first is off turns the first ON as well;
 *   - turning the first OFF turns the second OFF.
 *
 * `plannedOutputSheets` predicts the sheet list of the output BEFORE the merge with the very function the engine uses
 * (planSheetNames), so the preview and the merge result agree; after the merge `outputSheetRows` lists what the premerge reported.
 */
import type { PremergeLayout } from '../core/contracts';
import type { MergeReport } from '../core/types';
import { fileBaseName, planSheetNames, sanitizeSheetName, SHEET_NAME_MAX } from '../core/xlsx/sheetnames';
import { readyFiles } from './keysPlan';
import type { FileEntry } from './types';

// ---------------------------------------------------------------------------------------------
// the two options

export interface SheetOptions {
  includeOtherSheets: boolean;
  includeSourceSheets: boolean;
}

/** A state where the second option is on without the first cannot be reached by the UI; this makes it consistent anyway. */
export function normalizeSheetOptions(o: SheetOptions): SheetOptions {
  return !o.includeOtherSheets && o.includeSourceSheets ? { includeOtherSheets: false, includeSourceSheets: false } : o;
}

/** The user (un)ticked 「他のシートも出力に含める」. Ticking it also ticks the second one; clearing it clears the second one. */
export function withOtherSheets(cur: SheetOptions, include: boolean): SheetOptions {
  const c = normalizeSheetOptions(cur);
  if (c.includeOtherSheets === include) return c;
  return { includeOtherSheets: include, includeSourceSheets: include };
}

/** The user (un)ticked 「他のファイルのシートも含める」. Ticking it while the first one is off ticks the first one too. */
export function withSourceSheets(cur: SheetOptions, include: boolean): SheetOptions {
  const c = normalizeSheetOptions(cur);
  if (include) return c.includeOtherSheets && c.includeSourceSheets ? c : { includeOtherSheets: true, includeSourceSheets: true };
  return c.includeSourceSheets ? { includeOtherSheets: c.includeOtherSheets, includeSourceSheets: false } : c;
}

export function sameSheetOptions(a: SheetOptions, b: SheetOptions): boolean {
  return a.includeOtherSheets === b.includeOtherSheets && a.includeSourceSheets === b.includeSourceSheets;
}

// ---------------------------------------------------------------------------------------------
// the planned output (before the merge)

export type SheetKind = 'merged' | 'base' | 'copy';

export interface PlannedSheet {
  kind: SheetKind;
  /** sheet name in the output */
  name: string;
  /** the file the sheet comes from (base: the base file; copy: the other file) */
  sourceId: string;
  /** copy only: the file name as given and the sheet name in that file */
  file?: string;
  original?: string;
  /** copy only: "(n)" had to be appended because the plain name was taken */
  suffixed?: boolean;
  /** copy only: the n of "(n)" */
  suffixNo?: number;
  /** copy only: "<file>_<sheet>" was longer than Excel's 31 characters and had to be shortened */
  shortened?: boolean;
  /** the sheet is hidden in its file (a copy keeps that) */
  hidden?: boolean;
}

export interface PlannedOutput {
  sheets: PlannedSheet[];
  /** names of the files that take part with no sheet because they have no join key (結合対象外) */
  skipped: string[];
}

/** "(n)" at the end of a planned name -> n (the engine appends it to the plain name) */
export function suffixNumber(name: string): number | undefined {
  const m = /\((\d+)\)$/.exec(name);
  return m ? Number(m[1]) : undefined;
}

/** true when "<file>_<sheet>" is longer than Excel allows (the sheet name is then shortened by joinSheetName) */
export function isNameShortened(file: string, sheet: string): boolean {
  return sanitizeSheetName(fileBaseName(file), 'Book').length + 1 + sanitizeSheetName(sheet, 'Sheet').length > SHEET_NAME_MAX;
}

/**
 * The sheets of the output in order: the merged sheet (= the base data sheet), the base's other sheets when they are kept, then
 * (includeSourceSheets) every sheet of every participating other file, file by file in CARD order, sheet by sheet in workbook order.
 * `excluded`: ids of the files without a join key — they take no part in the merge, so their sheets are not copied either.
 * [] while the base file is not loaded.
 */
export function plannedOutputSheets(files: readonly FileEntry[], baseId: string | null, excluded: readonly string[], opts: SheetOptions): PlannedOutput {
  const ready = readyFiles(files);
  const base = ready.find((f) => f.id === baseId) ?? ready[0];
  if (!base?.book) return { sheets: [], skipped: [] };
  const merged = base.region?.sheet ?? base.sheet ?? base.book.sheets[0]?.name ?? '';
  const sheets: PlannedSheet[] = [{ kind: 'merged', name: merged, sourceId: base.id, hidden: false }];
  if (opts.includeOtherSheets) {
    for (const s of base.book.sheets) if (s.name !== merged) sheets.push({ kind: 'base', name: s.name, sourceId: base.id, hidden: s.state !== 'visible' });
  }
  const skipped: string[] = [];
  if (opts.includeSourceSheets) {
    const others = ready.filter((f) => f.id !== base.id && f.book);
    const taking = others.filter((f) => !excluded.includes(f.id));
    for (const f of others) if (excluded.includes(f.id)) skipped.push(f.fileName);
    const planned = planSheetNames(
      sheets.map((s) => s.name),
      // book.fileName is what the engine passes (KeySource.book.fileName); it equals the dropped file's name
      taking.map((f) => ({ file: f.book!.fileName || f.fileName, sheets: f.book!.sheets.map((s) => s.name) })),
    );
    let i = 0;
    for (const f of taking) {
      for (const sh of f.book!.sheets) {
        const p = planned[i++];
        sheets.push({
          kind: 'copy',
          name: p.name,
          sourceId: f.id,
          file: p.file,
          original: p.original,
          suffixed: p.suffixed,
          ...(p.suffixed ? { suffixNo: suffixNumber(p.name) } : {}),
          shortened: isNameShortened(p.file, p.original),
          hidden: sh.state !== 'visible',
        });
      }
    }
  }
  return { sheets, skipped };
}

// ---------------------------------------------------------------------------------------------
// after the merge

export interface OutputSheetCounts {
  total: number;
  /** the base's own other sheets that are kept */
  base: number;
  /** sheets copied from the other files */
  added: number;
}

/** `total`: the sheets of the merged workbook (as read back); `added`: MergeReport.addedSheets / PremergeLayout.addedSheets length */
export function outputSheetCounts(total: number, added: number): OutputSheetCounts {
  const a = Math.max(0, added);
  const t = Math.max(1, total);
  return { total: t, base: Math.max(0, t - 1 - a), added: a };
}

/** 「出力は 7 シート（結合 1 + 基準のほか 1 + 他ファイル 5）」 */
export function outputSheetsLine(c: OutputSheetCounts): string {
  return `出力は ${c.total} シート（結合 1 + 基準のほか ${c.base} + 他ファイル ${c.added}）`;
}

/** how many sheets the premerge copied: the layout when it has them, else the report's name list */
export function addedSheetCount(layout: Pick<PremergeLayout, 'addedSheets'> | null | undefined, report: Pick<MergeReport, 'addedSheets'> | null | undefined): number {
  return layout?.addedSheets?.length ?? report?.addedSheets?.length ?? 0;
}

export interface OutputSheetRow {
  kind: SheetKind;
  name: string;
  hidden: boolean;
  sourceId?: string;
  file?: string;
  original?: string;
  suffixed?: boolean;
  /** copy only: false = the copy had to leave things out (see the 警告 list) */
  complete?: boolean;
}

/**
 * The sheet list of the merged workbook, names only (the grid shows the merged sheet): `sheets` = the workbook as read back
 * (name + visibility, workbook order); the copies are told apart by PremergeLayout.addedSheets (else MergeReport.addedSheets).
 */
export function outputSheetRows(
  sheets: ReadonlyArray<{ name: string; state: string }>,
  layout: Pick<PremergeLayout, 'sheet' | 'addedSheets'> | null | undefined,
  report: Pick<MergeReport, 'addedSheets'> | null | undefined,
): OutputSheetRow[] {
  const detail = new Map((layout?.addedSheets ?? []).map((a) => [a.name, a] as const));
  const names = new Set((layout?.addedSheets?.length ? layout.addedSheets.map((a) => a.name) : (report?.addedSheets ?? [])));
  return sheets.map((s) => {
    const hidden = s.state !== 'visible';
    if (layout && s.name === layout.sheet) return { kind: 'merged', name: s.name, hidden };
    const a = detail.get(s.name);
    if (a) return { kind: 'copy', name: s.name, hidden, sourceId: a.sourceId, file: a.file, original: a.original, suffixed: a.suffixed, complete: a.complete };
    if (names.has(s.name)) return { kind: 'copy', name: s.name, hidden };
    return { kind: 'base', name: s.name, hidden };
  });
}

/** 「他のファイルのシート 7 枚を追加しました（02_面接評価_面接評価 ほか）」 for the toast detail */
export function addedSheetsLine(names: readonly string[]): string {
  if (!names.length) return '';
  const first = names[0].length > 20 ? `${names[0].slice(0, 19)}…` : names[0];
  return `他のファイルのシート ${names.length} 枚を追加しました（${first}${names.length > 1 ? ' ほか' : ''}）`;
}
