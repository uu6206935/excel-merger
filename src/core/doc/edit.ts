/**
 * Cell edits, column widths and formula display.
 *
 *  - setCellInput: Excel-like parsing (parse.ts) of what the user typed, written into ONE cell of the merged sheet.
 *    The cell style is kept, except that a date / % / currency typed into a General (or non-matching) cell changes
 *    the cell's number format like Excel (CellWrite.cell.numFmt; the doc derives the display style). A new cell
 *    takes the row format (row with customFormat) like the patcher / Excel. Typed formulas whose legacy
 *    evaluation could differ are stored as dynamic arrays (formula/dynamic.ts; CellWrite.cell.dyn; the doc cell
 *    becomes a spill master with arr = its own address). Refused like Excel inside array formulas.
 *    Editing the header cell of an Excel table renames the table column: the header gets the unique name Excel
 *    would store (duplicates -> "名前2", empty -> "列1") and every structured reference to the column is renamed
 *    (Excel does this itself when the header cell changes; the JS patcher mirrors it since round 2).
 *    A constant typed into a table's TOTALS row cell becomes text (t 's') = the typed text, which Excel stores as the
 *    column's totalsRowLabel (numbers, dates, TRUE ... all stay text, no number format change); formulas stay formulas
 *    (SUBTOTAL(n,Tbl[own column]) becomes the built-in totals function in the file).
 *    A text result that Excel would re-parse when the cell is edited again (leading apostrophe typed, digit-only input
 *    kept as text in an ID column, > 15 digits) is logged with CellWrite.cell.quote = true (parse.ts needsQuotePrefix);
 *    the doc cell shows the text without the apostrophe (Cell / XStyle have no quotePrefix field).
 *  - setColumnWidth: props.width (+ unhide) and customWidth = true (a user-set width is never auto-widened again),
 *    logged as {kind:'colWidth'}.
 *  - formulaForDisplay: FILE form -> UI form ('=' ...), with the cell's table as context ([@列] like Excel).
 */
import { parseAreaText, tokenize, MAX_ROW } from '../formula/excelRef';
import { isDynamicArrayFormula } from '../formula/dynamic';
import { toUiForm } from '../formula/forms';
import type { FormOptions } from '../formula/forms';
import type { Cell, CellWrite, DefinedName, DocColumn, EditorSnapshot, ExternalFormulaCell, FeatureRule, MergedDoc, RangeBox, TableInfo } from '../types';
import type { OpResult } from '../contracts';
import { autoWidenedWidth } from './autowiden';
import { InsertStyles, textWrapsOnWrite } from './insertStyle';
import { parseInput } from './parse';
import type { ParseContext } from './parse';
import { renameSrefColumn, tableBox, uniqueHeaderName } from './tables';
import { addr, cachedOf, carryCompiled, clearSpillValues, colIndexOf, newCellXf, setCached, styleOf } from './util';

// ---------------------------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------------------------
/** Table containing (row, col) of the merged sheet, if any. */
export function tableAt(doc: MergedDoc, row: number, col: number): TableInfo | undefined {
  for (const t of doc.features.tables) {
    const b = tableBox(t);
    if (b && row >= b.r1 && row <= b.r2 && col >= b.c1 && col <= b.c2) return t;
  }
  return undefined;
}

const DIGITS = /^[0-9]+$/;
const textIdsCache = new WeakMap<Array<Cell | undefined>, boolean>();

/**
 * The column holds IDs stored as text digits ('00123', '2024001' typed as text): digit-only input must stay
 * text there (leading zeros, key matching). True when text-digit cells are at least as common as numbers.
 */
export function holdsTextIds(doc: MergedDoc, col: DocColumn): boolean {
  const cached = textIdsCache.get(col.cells);
  if (cached !== undefined) return cached;
  let textDigits = 0;
  let numbers = 0;
  let zeroLed = false;
  const from = Math.max(1, doc.firstDataRow);
  const to = Math.min(col.cells.length, Math.max(doc.lastDataRow, from));
  for (let r = from; r <= to; r++) {
    const c = col.cells[r - 1];
    if (!c || c.f !== undefined) continue;
    if (c.t === 's' && typeof c.v === 'string' && DIGITS.test(c.v.trim())) {
      textDigits++;
      if (c.v.trim().length > 1 && c.v.trim()[0] === '0') zeroLed = true;
    } else if (c.t === 'n') numbers++;
  }
  const out = textDigits > 0 && (zeroLed || textDigits >= numbers);
  textIdsCache.set(col.cells, out);
  return out;
}

interface ArrayHit {
  text: string;
  /** dynamic-array (spill) formula */
  dyn: boolean;
  /** (row, col) is the formula's own cell */
  master: boolean;
  box: RangeBox | null;
}

function arrayAt(doc: MergedDoc, row: number, col: number): ArrayHit | null {
  for (const a of doc.features.arrays) {
    if (row >= a.r1 && row <= a.r2 && col >= a.c1 && col <= a.c2) {
      const m = doc.columns[a.c1 - 1]?.cells[a.r1 - 1];
      const text = a.r1 === a.r2 && a.c1 === a.c2 ? addr(a.c1, a.r1) : `${addr(a.c1, a.r1)}:${addr(a.c2, a.r2)}`;
      return { text, dyn: !!m?.dyn, master: row === a.r1 && col === a.c1, box: a };
    }
  }
  // masters whose range the reader did not list
  const master = doc.columns[col - 1]?.cells[row - 1];
  if (master?.arr) {
    const a = parseAreaText(master.arr.replace(/\$/g, ''));
    if (a && !(a.c1 === a.c2 && a.r1 === a.r2)) return { text: master.arr, dyn: !!master.dyn, master: true, box: null };
  }
  return null;
}

function mergeCovering(doc: MergedDoc, row: number, col: number) {
  return doc.features.merges.find((m) => row >= m.r1 && row <= m.r2 && col >= m.c1 && col <= m.c2);
}

/** Every sheet-name prefix of a formula must name a sheet of the book (Excel would ask for an external file). */
function unknownSheetIn(f: string, doc: MergedDoc): string | null {
  if (!f.includes('!')) return null;
  const known = new Set(doc.sheetNames.map((s) => s.toLowerCase()));
  for (const t of tokenize(f)) {
    const p = t.prefix;
    if (!p || p.book) continue;
    const sheets = p.sheet2 !== undefined ? [(p as unknown as { sheet1?: string }).sheet1 ?? '', p.sheet2] : [p.sheet ?? ''];
    for (const s of sheets) if (s && !known.has(s.toLowerCase())) return s;
  }
  return null;
}

const sameContent = (a: Cell | undefined, b: Cell | undefined): boolean => {
  const ea = !a || (a.f === undefined && (a.v === null || a.v === '') && a.arr === undefined);
  const eb = !b || (b.f === undefined && (b.v === null || b.v === '') && b.arr === undefined);
  if (ea || eb) return ea && eb;
  return a!.t === b!.t && a!.v === b!.v && a!.f === b!.f && a!.arr === b!.arr && !a!.rich === !b!.rich;
};

/** refersTo of a defined name as a formula of the merged sheet sees it (sheet scope first), null when unknown. */
export function resolveNameIn(doc: MergedDoc, name: string): string | null {
  const n = name.toLowerCase();
  const sheet = doc.sheetName.toLowerCase();
  let wb: string | null = null;
  for (const d of doc.names) {
    if (d.name.toLowerCase() !== n) continue;
    if (d.scope === null) wb = d.refersTo;
    else if (d.scope.toLowerCase() === sheet) return d.refersTo;
  }
  return wb;
}

function replaceCell(doc: MergedDoc, index: number, row: number, cell: Cell | undefined): DocColumn[] {
  const col = doc.columns[index];
  const cells = col.cells.slice();
  carryCompiled(col.cells, cells);
  cells[row - 1] = cell;
  const columns = doc.columns.slice();
  columns[index] = { ...col, cells };
  return columns;
}

// ---------------------------------------------------------------------------------------------
// table header rename
// ---------------------------------------------------------------------------------------------
function renameInDoc(doc: MergedDoc, columns: DocColumn[], table: TableInfo, oldName: string, newName: string): Pick<MergedDoc, 'columns' | 'externalFormulas' | 'names' | 'features' | 'externalRules'> {
  const tl = table.name.toLowerCase();
  const needle = table.name.toLowerCase();
  const ren = (f: string) => (f.toLowerCase().includes(needle) ? renameSrefColumn(f, tl, oldName, newName, null) : f);
  // merged sheet
  let outCols = columns;
  columns.forEach((col, i) => {
    let cells = col.cells;
    for (let r = 0; r < cells.length; r++) {
      const c = cells[r];
      if (!c || c.f === undefined) continue;
      const nf = ren(c.f);
      if (nf === c.f) continue;
      if (cells === col.cells) {
        cells = cells.slice();
        carryCompiled(col.cells, cells);
      }
      cells[r] = { ...c, f: nf };
    }
    if (cells !== col.cells) {
      if (outCols === columns) outCols = columns.slice();
      outCols[i] = { ...col, cells };
    }
  });
  let ext: ExternalFormulaCell[] = doc.externalFormulas;
  doc.externalFormulas.forEach((e, i) => {
    const nf = ren(e.f);
    if (nf === e.f) return;
    if (ext === doc.externalFormulas) ext = ext.slice();
    const ne: ExternalFormulaCell = { sheet: e.sheet, row: e.row, col: e.col, f: nf };
    setCached(ne, cachedOf(e));
    ext[i] = ne;
  });
  const names: DefinedName[] = doc.names.map((n) => {
    const nf = ren(n.refersTo);
    return nf === n.refersTo ? n : { ...n, refersTo: nf };
  });
  const renRule = (r: FeatureRule): FeatureRule => {
    const fs = r.formulas.map(ren);
    return fs.every((f, i) => f === r.formulas[i]) ? r : { ...r, formulas: fs };
  };
  const rules = doc.features.rules.map(renRule);
  const externalRules = doc.externalRules.map((er) => {
    const r = renRule(er.rule);
    return r === er.rule ? er : { sheet: er.sheet, rule: r };
  });
  const tables = doc.features.tables.map((t) => (t === table ? { ...t, columns: t.columns.map((c) => (c === oldName ? newName : c)) } : t));
  return {
    columns: outCols,
    externalFormulas: ext,
    names: names.every((n, i) => n === doc.names[i]) ? doc.names : names,
    features: { ...doc.features, tables, rules: rules.every((r, i) => r === doc.features.rules[i]) ? doc.features.rules : rules },
    externalRules: externalRules.every((r, i) => r === doc.externalRules[i]) ? doc.externalRules : externalRules,
  };
}

// ---------------------------------------------------------------------------------------------
// setCellInput
// ---------------------------------------------------------------------------------------------
export function setCellInputImpl(s: EditorSnapshot, colId: string, row: number, text: string, now?: Date): OpResult {
  const doc = s.doc;
  const index = colIndexOf(doc, colId);
  if (index < 0) return { ok: false, reason: '指定された列が見つかりません。画面を更新してからやり直してください。' };
  if (!Number.isInteger(row) || row < 1 || row > Math.min(doc.rowCount, MAX_ROW)) return { ok: false, reason: 'セルの位置が正しくありません。' };
  if (typeof text !== 'string') return { ok: false, reason: '入力内容が正しくありません。' };
  const colNo = index + 1;
  const arr = arrayAt(doc, row, colNo);
  if (arr && arr.dyn && !arr.master) {
    return { ok: false, reason: `このセルはスピル範囲（${arr.text}）の一部です。左上のセルの数式を編集してください。` };
  }
  if (arr && !arr.dyn) return { ok: false, reason: `配列数式（${arr.text}）の一部は変更できません。（Excel と同じ制限です）` };
  // editing the formula cell of a dynamic array replaces the whole spill (Excel recalculates the new spill)
  const dropArray = arr && arr.dyn && arr.master ? arr.box : null;
  const merge = mergeCovering(doc, row, colNo);
  if (merge && (merge.r1 !== row || merge.c1 !== colNo)) {
    return { ok: false, reason: `結合されたセル（${addr(merge.c1, merge.r1)}:${addr(merge.c2, merge.r2)}）には左上のセル（${addr(merge.c1, merge.r1)}）にだけ入力できます。` };
  }
  const col = doc.columns[index];
  const old = col.cells[row - 1];
  // the style the written cell keeps: its own (no s = xf 0, as read); a new cell takes the row format (customFormat),
  // else the column format — exactly like Excel / the patcher
  const cellXf = old ? old.s : newCellXf(doc, row) ?? col.props.s;
  const xf = cellXf ?? 0;
  const numFmt = styleOf(doc, xf)?.numFmt;
  const table = tableAt(doc, row, colNo);
  const box = table ? tableBox(table) : null;
  const isTableHeader = !!(table && box && table.headerRow && row === box.r1);
  const isTotalsCell = !!(table && box && table.totalsRow && !isTableHeader && row === box.r2);

  const formOptions: FormOptions = { tables: doc.features.tables.map((t) => t.name) };
  if (table && !isTableHeader) formOptions.hostTable = table.name;
  const ctx: ParseContext = { numFmt, textIds: row >= doc.firstDataRow && holdsTextIds(doc, col), formOptions };
  if (now) ctx.now = now;
  const parsed = parseInput(text, ctx);
  if (parsed.kind === 'error') return { ok: false, reason: parsed.reason };

  const warnings: string[] = [];
  let write: CellWrite['cell'];
  let next: Cell | undefined;
  const keepStyle = (c: Cell): Cell => {
    if (cellXf !== undefined) c.s = cellXf;
    return c;
  };

  if (isTableHeader && table && box) {
    // table header: always text, unique within the table; renames the table column (and its structured refs)
    if (parsed.kind === 'formula') return { ok: false, reason: 'テーブルの見出しには数式を入力できません。' };
    const pos = colNo - box.c1;
    const oldName = table.columns[pos] ?? '';
    // header cells are always text: a leading apostrophe is only Excel's quote prefix (no stored prefix needed here)
    const typed = parsed.kind === 'clear' ? '' : parsed.kind === 'value' && parsed.t === 's' && parsed.quote ? String(parsed.v) : text;
    const others = table.columns.filter((_, i) => i !== pos);
    const name = uniqueHeaderName(typed, others);
    if (name !== typed && typed !== '') warnings.push(`テーブルの見出しは重複できないため「${name}」にしました。`);
    if (name === oldName && old?.t === 's' && old.v === name && old.f === undefined) return { ok: true, snapshot: s, warnings: [] };
    next = keepStyle({ t: 's', v: name });
    write = { t: 's', v: name };
    let hstyles = doc.styles;
    if (textWrapsOnWrite(write)) {
      const d = new InsertStyles(doc.styles);
      const ns = d.withWrap(next.s ?? 0);
      if (ns !== (next.s ?? 0)) next.s = ns;
      hstyles = d.styles;
    }
    const columns = replaceCell(doc, index, row, next);
    const renamed = name === oldName ? null : renameInDoc(doc, columns, table, oldName, name);
    const ndoc: MergedDoc = renamed ? { ...doc, ...renamed, styles: hstyles, version: doc.version + 1 } : { ...doc, columns, styles: hstyles, version: doc.version + 1 };
    return {
      ok: true,
      snapshot: { doc: ndoc, log: [...s.log, { kind: 'cells', sheet: doc.sheetName, writes: [{ row, col: colNo, cell: write }], reason: 'edit' }], label: '見出しを変更' },
      warnings,
    };
  }

  let styles = doc.styles;
  let derived: InsertStyles | null = null;
  const derive = (): InsertStyles => (derived ??= new InsertStyles(doc.styles));
  let spill: RangeBox | null = null;
  if (parsed.kind === 'clear') {
    next = old?.s !== undefined ? { t: 'z', v: null, s: old.s } : undefined;
    write = null;
  } else if (parsed.kind === 'formula') {
    const bad = unknownSheetIn(parsed.f, doc);
    if (bad) return { ok: false, reason: `数式のシート「${bad}」がこのブックにありません。シート名を確認してください。` };
    next = keepStyle({ t: 'z', v: null, f: parsed.f });
    write = { t: 'z', v: null, f: parsed.f };
    // Excel 365 stores a typed formula as a dynamic array when legacy evaluation could differ (=A1:A3, =SEQUENCE(3)) —
    // except in a table's totals row, where the formula is the column's totals entry (the patcher never spills there)
    if (!isTotalsCell && isDynamicArrayFormula(parsed.f, { resolveName: (n) => resolveNameIn(doc, n) })) {
      write.dyn = true;
      next.dyn = true;
      next.arr = addr(colNo, row); // the spill range is computed by Excel on load
      spill = { r1: row, c1: colNo, r2: row, c2: colNo };
    }
  } else if (isTotalsCell) {
    // Excel: a constant typed into a totals-row cell is the column's totalsRowLabel — the cell holds the typed TEXT
    // (a number, a date, TRUE ... stay text; the cell format is not touched)
    const label = parsed.t === 's' ? String(parsed.v) : text;
    next = keepStyle({ t: 's', v: label });
    write = { t: 's', v: label };
    if (parsed.t !== 's') warnings.push('テーブルの集計行に入力した値は、計算結果ではなく文字列（集計行のラベル）として保存されます。合計などは数式（SUBTOTAL など）で入力してください。');
    else if (parsed.hint) warnings.push(parsed.hint);
  } else {
    next = keepStyle({ t: parsed.t, v: parsed.v });
    write = { t: parsed.t, v: parsed.v };
    // text Excel would re-parse on the next edit is stored with a quote prefix (both engines honour the flag exactly)
    if (parsed.quote) write.quote = true;
    if (parsed.numFmt !== undefined) {
      // Excel gives the cell the number format of what was typed (2026/10/20 -> yyyy/m/d, 50% -> 0%, ¥1,000 ...)
      write.numFmt = parsed.numFmt;
      const d = derive();
      const ns = d.withNumFmt(xf, parsed.numFmt);
      styles = d.styles;
      if (ns) next.s = ns;
      else delete next.s;
    }
    if (parsed.hint) warnings.push(parsed.hint);
  }
  // a text constant with a line feed turns the cell's wrap text on (Excel does it by itself; the patcher mirrors it)
  if (next && textWrapsOnWrite(write)) {
    const d = derive();
    const cur = next.s ?? 0;
    const ns = d.withWrap(cur);
    if (ns !== cur) next.s = ns;
    styles = d.styles;
  }
  // a quote-prefixed write is never a no-op: the doc cell cannot show whether the stored cell already has the prefix
  if (!write?.quote && sameContent(old, next) && next?.s === old?.s && !!next?.dyn === !!old?.dyn) return { ok: true, snapshot: s, warnings: [] };
  let columns = replaceCell(doc, index, row, next);
  // like Excel (and the patcher), a standard-width column grows when the new number format no longer fits
  if (write?.numFmt && write.f === undefined && typeof write.v === 'number') {
    const width = autoWidenedWidth(doc, columns[index], next?.s, write.v, write.numFmt, styles);
    if (width !== null) columns[index] = { ...columns[index], props: { ...columns[index].props, width } };
  }
  // the old spill disappears (Excel re-spills the new formula on load), like the patcher's clearSpill
  if (dropArray) columns = clearSpillValues(columns, dropArray, row, colNo);
  let arrays = dropArray ? doc.features.arrays.filter((a) => a !== dropArray) : doc.features.arrays;
  if (spill) arrays = [...arrays, spill];
  const features = arrays !== doc.features.arrays ? { ...doc.features, arrays } : doc.features;
  const ndoc: MergedDoc = { ...doc, columns, features, styles, version: doc.version + 1 };
  const label = row === doc.headerRow ? '見出しを変更' : 'セルを編集';
  return {
    ok: true,
    snapshot: { doc: ndoc, log: [...s.log, { kind: 'cells', sheet: doc.sheetName, writes: [{ row, col: colNo, cell: write }], reason: 'edit' }], label },
    warnings,
  };
}

// ---------------------------------------------------------------------------------------------
// column width
// ---------------------------------------------------------------------------------------------
/** width: the value stored in <col width> (ColProps.width unit, incl. Excel's padding), capped at 255 like the patcher. */
export function setColumnWidthImpl(s: EditorSnapshot, colId: string, width0: number): OpResult {
  const doc = s.doc;
  const index = colIndexOf(doc, colId);
  if (index < 0) return { ok: false, reason: '指定された列が見つかりません。画面を更新してからやり直してください。' };
  if (typeof width0 !== 'number' || !Number.isFinite(width0) || width0 < 0) return { ok: false, reason: '列幅の指定が正しくありません。' };
  const width = Math.round(Math.min(255, width0) * 1e6) / 1e6;
  const col = doc.columns[index];
  // the width of a column the user sets is a custom width, even when it equals the stored one (Excel: customWidth="1")
  if (col.props.width === width && col.props.customWidth && !(width > 0 && col.props.hidden)) return { ok: true, snapshot: s, warnings: [] };
  const props = { ...col.props, width, customWidth: true };
  if (width > 0 && props.hidden) delete props.hidden;
  const columns = doc.columns.slice();
  columns[index] = { ...col, props };
  const warnings = width0 > 255 ? ['列幅は Excel の上限（255）にしました。'] : [];
  return {
    ok: true,
    snapshot: { doc: { ...doc, columns, version: doc.version + 1 }, log: [...s.log, { kind: 'colWidth', sheet: doc.sheetName, col: index + 1, width }], label: '列幅を変更' },
    warnings,
  };
}

// ---------------------------------------------------------------------------------------------
// display
// ---------------------------------------------------------------------------------------------
const displayCache = new WeakMap<Cell, { key: string; text: string }>();

export function formulaForDisplayImpl(doc: MergedDoc, colIndex: number, row: number): string | null {
  const cell = doc.columns[colIndex]?.cells[row - 1];
  if (!cell || cell.f === undefined) return null;
  const table = doc.features.tables.length ? tableAt(doc, row, colIndex + 1) : undefined;
  const key = table?.name ?? '';
  const hit = displayCache.get(cell);
  if (hit && hit.key === key) return hit.text;
  let text: string;
  try {
    text = toUiForm(cell.f, table ? { hostTable: table.name } : {});
  } catch {
    text = '=' + cell.f;
  }
  // legacy (CSE) array formulas are shown in braces by Excel
  if (cell.arr !== undefined && !cell.dyn) text = `{${text}}`;
  displayCache.set(cell, { key, text });
  return text;
}
