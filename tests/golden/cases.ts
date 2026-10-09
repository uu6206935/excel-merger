/**
 * Golden cases: op sequences replayed both by real Excel (tests/golden/com/run-ops.ps1, results saved as
 * tests/fixtures/golden/out/<name>.xlsx) and by replayLog(); tests/golden/golden.test.ts compares the two.
 * Moves are replayed in Excel with the clipboard-free sequence (insert blank at the target, Range.Cut(Destination),
 * delete the source), which gives the same references as Insert Cut Cells (R2b).
 */
import { planMove } from '../../src/core/formula/excelRef';
import type { CellWrite, LoggedOp, StructOp } from '../../src/core/types';

export interface GoldenCase {
  name: string;
  /** fixture file in tests/fixtures/golden */
  fixture: string;
  ops: LoggedOp[];
  /** snapshot keys that are not compared for this case (documented deviations), with the reason */
  skip?: Record<string, string>;
  /** diff paths (prefix match) that are known deviations, with the reason */
  known?: Record<string, string>;
  /** Excel refuses the op sequence (then replayLog must throw a Japanese error) */
  expectError?: boolean;
  /** Excel's result file relative to tests/fixtures/golden (default out/<name>.xlsx) */
  golden?: string;
}

const s = (sheet: string, op: StructOp): LoggedOp => ({ kind: 'struct', sheet, op });
const insC = (sheet: string, at: number, count = 1) => s(sheet, { type: 'insert', axis: 'col', at, count });
const delC = (sheet: string, at: number, count = 1) => s(sheet, { type: 'delete', axis: 'col', at, count });
const movC = (sheet: string, from: number, count: number, to: number) => s(sheet, { type: 'move', axis: 'col', from, count, to });
const insR = (sheet: string, at: number, count = 1) => s(sheet, { type: 'insert', axis: 'row', at, count });
const delR = (sheet: string, at: number, count = 1) => s(sheet, { type: 'delete', axis: 'row', at, count });
const plan = (sheet: string, sel: number[], target: number) => planMove(sel, target).map((op) => s(sheet, op));

const CF_INSIDE_MOVE = 'a block moved INSIDE the range of a CF rule: Excel (insert + Range.Cut(Destination) + delete) shrinks the rule to the columns between the destination and the source (probe-cf-move.ps1: A2:F10 with D moved before B -> B2:D10); replayLog keeps cell-set semantics (every cell keeps its rule, excelRef per-cell model)';
const AUTO_NUMFMT = 'Excel gives a General cell that receives a formula the number format of the first precedent of the formula (UI auto-format); CellWrite never changes styles (types.ts), so replayLog keeps the cell format';
const TABLE_FIRST_ROW = 'row inserted at the FIRST data row of a ListObject: Excel grows ranges that exactly cover the table data rows (A4:F13 -> A4:F14 instead of A5:F14) and formats the new row like the data rows (probe-table-row.ps1); replayLog applies the plain insert rule (shift). The app never inserts rows there (premerge appends after the last data row)';
const QUOTE_COM = 'through COM a cell keeps its quotePrefix flag when a value is written over it (Value2 / Formula / ClearContents: PrefixCharacter reads empty for a number / empty cell but the flag stays in the xf and acts again when text is written; only ClearFormats / Clear remove it, together with every other format); typing in the Excel UI resets it and so does CellWrite.cell.quote (types.ts: anything but quote === true leaves no prefix), which replayLog implements — the Excel runner has to reset it explicitly';
const TABLE_CLIPFREE ='moving columns inside a table: the clipboard-free sequence used to build the golden deletes and re-creates the table column, so structured refs become #REF!; native Insert Cut Cells reorders the table columns and keeps the refs (R2b R11), which is what replayLog does';

export const CASES: GoldenCase[] = [
  // ---- torture.xlsx (R1): Data has merges A1:F1, frozen C4, CF/DV (x14 list), names, chart, pivot, notes ----
  { name: 't-ins-C', fixture: 'torture.xlsx', ops: [insC('Data', 3)] },
  { name: 't-ins-A2', fixture: 'torture.xlsx', ops: [insC('Data', 1, 2)] },
  { name: 't-ins-K3', fixture: 'torture.xlsx', ops: [insC('Data', 11, 3)] },
  { name: 't-del-C', fixture: 'torture.xlsx', ops: [delC('Data', 3)] },
  { name: 't-del-EF', fixture: 'torture.xlsx', ops: [delC('Data', 5, 2)] },
  { name: 't-mov-HI-G', fixture: 'torture.xlsx', ops: [movC('Data', 8, 2, 7)] },
  { name: 't-mov-G-K', fixture: 'torture.xlsx', ops: [movC('Data', 7, 1, 11)] },
  { name: 't-mov-L-H', fixture: 'torture.xlsx', ops: [movC('Data', 12, 1, 8)] },
  { name: 't-ins-row10', fixture: 'torture.xlsx', ops: [insR('Data', 10, 2)] },
  { name: 't-del-row5', fixture: 'torture.xlsx', ops: [delR('Data', 5)] },
  { name: 't-del-rows20', fixture: 'torture.xlsx', ops: [delR('Data', 20, 4)] },
  { name: 't-plan-GIJ-H', fixture: 'torture.xlsx', ops: plan('Data', [7, 9, 10], 8) },
  {
    name: 't-seq', fixture: 'torture.xlsx',
    ops: [insC('Data', 5), delC('Data', 12), movC('Data', 9, 1, 14), insR('Data', 6), { kind: 'colWidth', sheet: 'Data', col: 5, width: 20 }],
  },
  { name: 't-del-B', fixture: 'torture.xlsx', ops: [delC('Data', 2)] }, // note B4 + threaded comment B5 deleted, merge shrinks
  { name: 't-ins-row1', fixture: 'torture.xlsx', ops: [insR('Data', 1, 1)] },
  { name: 't-mov-rows', fixture: 'torture.xlsx', ops: [s('Data', { type: 'move', axis: 'row', from: 5, count: 2, to: 12 })] },
  { name: 't-del-rows3', fixture: 'torture.xlsx', ops: [delR('Data', 15, 3)] },
  {
    name: 't-cells', fixture: 'torture.xlsx',
    ops: [
      { kind: 'cells', sheet: 'Data', reason: 'edit', writes: [
        { row: 6, col: 5, cell: { t: 'n', v: 12345 } },
        { row: 7, col: 2, cell: { t: 's', v: '書き換え' } },
        { row: 25, col: 1, cell: { t: 'z', v: null, f: 'SUM(E4:E23)' } },
        { row: 8, col: 11, cell: null },
      ] },
      insC('Data', 3),
      { kind: 'cells', sheet: 'Data', reason: 'edit', writes: [{ row: 26, col: 1, cell: { t: 'z', v: null, f: 'A25*2+F4' } }] },
    ],
    known: { '[Data].styles.A25': AUTO_NUMFMT },
  },
  // ---- g-table.xlsx: base region is a ListObject with a 合計 row below it (premerge-like layout) ----
  { name: 'g-ins-D', fixture: 'g-table.xlsx', ops: [insC('名簿', 4)] },
  { name: 'g-ins-G', fixture: 'g-table.xlsx', ops: [insC('名簿', 7, 2)] },
  { name: 'g-del-E', fixture: 'g-table.xlsx', ops: [delC('名簿', 5)] },
  {
    name: 'g-mov-D-B', fixture: 'g-table.xlsx', ops: [movC('名簿', 4, 1, 2)],
    known: { '[名簿].cf': CF_INSIDE_MOVE, '[名簿].formulas.F': TABLE_CLIPFREE, '[名簿].tables': TABLE_CLIPFREE, '[Summary].formulas.A2': TABLE_CLIPFREE, '[Summary].formulas.A9': TABLE_CLIPFREE },
  },
  {
    name: 'g-plan-ACE-F', fixture: 'g-table.xlsx', ops: plan('名簿', [1, 3, 5], 6),
    known: { '[Summary].formulas.A9': TABLE_CLIPFREE },
  },
  { name: 'g-ins-row14', fixture: 'g-table.xlsx', ops: [insR('名簿', 14, 2)] },
  { name: 'g-ins-row8', fixture: 'g-table.xlsx', ops: [insR('名簿', 8, 3)] },
  { name: 'g-del-rows6', fixture: 'g-table.xlsx', ops: [delR('名簿', 6, 2)] },
  { name: 'g-mov-I-A', fixture: 'g-table.xlsx', ops: [movC('名簿', 9, 1, 1)], expectError: true }, // I is part of merge I1:K1
  { name: 'g-mov-L-A', fixture: 'g-table.xlsx', ops: [movC('名簿', 12, 1, 1)] },
  { name: 'g-mov-L-C', fixture: 'g-table.xlsx', ops: [movC('名簿', 12, 1, 3)], expectError: true }, // into the table
  // a data row dropped AT the header row (tblMain A3:G13, header row 3) would put a data row into the table header: refused (op fuzz, round 4)
  { name: 'g-mov-row-to-header', fixture: 'g-table.xlsx', ops: [s('名簿', { type: 'move', axis: 'row', from: 6, count: 1, to: 3 })], expectError: true },
  { name: 'g-mov-header-row', fixture: 'g-table.xlsx', ops: [s('名簿', { type: 'move', axis: 'row', from: 3, count: 1, to: 8 })], expectError: true }, // the header row itself
  // out of the table: the table is A:G, so H (8, right after its last column) is still a reorder (Excel keeps the block as
  // the last table column, round 3: tests/xlsx/patch-round3.test.ts); one column further (I) leaves the table
  { name: 'g-mov-B-I', fixture: 'g-table.xlsx', ops: [movC('名簿', 2, 1, 9)], expectError: true },
  { name: 'g-del-HJ', fixture: 'g-table.xlsx', ops: [delC('名簿', 8, 3)] },
  {
    name: 'g-ins-row4', fixture: 'g-table.xlsx', ops: [insR('名簿', 4, 1)],
    known: { names: TABLE_FIRST_ROW, charts: TABLE_FIRST_ROW, '[名簿].formulas': TABLE_FIRST_ROW, '[名簿].styles': TABLE_FIRST_ROW, '[名簿].cf': TABLE_FIRST_ROW, '[名簿].dv': TABLE_FIRST_ROW, '[名簿].tables': TABLE_FIRST_ROW, '[Summary].formulas': TABLE_FIRST_ROW },
  },
  { name: 'g-ins-row13', fixture: 'g-table.xlsx', ops: [insR('名簿', 13, 1)] }, // last data row: table grows
  { name: 'g-del-B-row5', fixture: 'g-table.xlsx', ops: [delC('名簿', 2), delR('名簿', 5)] }, // note on B5
  { name: 'g-ins-A', fixture: 'g-table.xlsx', ops: [insC('名簿', 1)] },
  // ---- probes (tests/golden/com/probe-*.ps1): Excel's before / after files ----
  // inserted cells: left / upper format, border sides only where both neighbours agree
  { name: 'p-borders', fixture: 'probe/probe-before.xlsx', golden: 'probe/probe-after.xlsx', ops: [insC('Cols', 3), insR('Rows', 3)] },
  // every CF type (and DV) grows over an inserted column
  { name: 'p-cf-ins', fixture: 'probe/cf-before.xlsx', golden: 'probe/cf-ins.xlsx', ops: [insC('P', 4)] },
  // a moved block landing right after a rule: cellIs / expression / text / blanks rules and DV leak onto it,
  // colorScale / dataBar / iconSet / top10 / unique / aboveAverage do not
  { name: 'p-cf-mov', fixture: 'probe/cf-before.xlsx', golden: 'probe/cf-mov.xlsx', ops: [movC('P', 6, 1, 4)] },
  { name: 'p-cf-inside', fixture: 'probe/cfm-before.xlsx', golden: 'probe/cfm-after.xlsx', ops: [movC('Sheet1', 4, 1, 2)], known: { '[Sheet1].cf': CF_INSIDE_MOVE } },

  // ---- round 2 ------------------------------------------------------------------------------------------
  // r2-arrays.xlsx (com/make-fixtures-r2.ps1): A!B2:B4 legacy CSE array, A!D2 =SEQUENCE(3,2) spill, A!G2 string spill
  // dynamic-array formulas written by the user (Excel: Range.Formula2), then a column insert
  {
    name: 'r-dyn', fixture: 'r2-arrays.xlsx',
    ops: [
      cells('A', [w(8, 4, { t: 'z', v: null, f: '_xlfn.SEQUENCE(3)', dyn: true }), w(8, 6, { t: 'z', v: null, f: '_xlfn.ANCHORARRAY(D8)*2', dyn: true }), w(12, 4, { t: 'z', v: null, f: 'SUM(A2:A4)' })]),
      insC('A', 3),
    ],
  },
  // a new dynamic array in a book without xl/metadata.xml (part, relationship and content type are created)
  // (Excel 365 crashes (RPC failure) when Formula2 is set anywhere in g-table.xlsx through COM: pm-base-plain is used)
  {
    name: 'r-dyn-new', fixture: 'pm-base-plain.xlsx', ops: [cells('Report', [w(1, 3, { t: 'z', v: null, f: '_xlfn.SEQUENCE(2,2)', dyn: true })])],
  },
  // 'toValues' over array / spill masters: the whole array becomes values (Excel: range.Value2 = range.Value2)
  {
    name: 'r-tovalues', fixture: 'r2-arrays.xlsx', ops: [cells('A', [w(2, 2, { t: 'n', v: 10 }), w(2, 4, { t: 'n', v: 1 }), w(2, 7, { t: 's', v: '1x' })], 'toValues')],
  },
  // a value typed over a spill master: the spilled values disappear
  { name: 'r-edit-spill', fixture: 'r2-arrays.xlsx', ops: [cells('A', [w(2, 4, { t: 'n', v: 5 })])] },
  // numFmt with the value (Range.NumberFormat after the value): built-in ids for ja-JP built-in codes, custom
  // otherwise; the standard-width column I is widened for the date (2026/9/30 would show #####)
  {
    name: 'r-numfmt', fixture: 'r2-arrays.xlsx',
    ops: [cells('A', [
      w(2, 9, { t: 'd', v: 46295, numFmt: 'yyyy/m/d' }), w(3, 9, { t: 'n', v: 0.5, numFmt: '0%' }), w(4, 9, { t: 'n', v: 1000, numFmt: '"¥"#,##0' }),
      w(5, 9, { t: 'n', v: 1234.5, numFmt: '#,##0.0' }), w(6, 9, { t: 'd', v: 0.375, numFmt: 'h:mm' }),
    ])],
  },
  // column widening by number formats: dates / times / grouped numbers / percent / Japanese literals; a column with a
  // user width (W) and a narrow value (S) are not widened
  {
    name: 'r-numfmt-w', fixture: 'r2-arrays.xlsx',
    ops: [
      { kind: 'colWidth', sheet: 'A', col: 23, width: 6 },
      cells('A', [
        w(6, 12, { t: 'd', v: 46023, numFmt: 'yyyy/m/d' }), w(6, 13, { t: 'd', v: 46305, numFmt: 'yyyy/m/d' }),
        w(6, 14, { t: 'd', v: 46295.3958333333, numFmt: 'yyyy/m/d h:mm' }), w(6, 15, { t: 'n', v: 123456789, numFmt: '#,##0' }),
        w(6, 16, { t: 'n', v: 1234567, numFmt: '"$"#,##0' }), w(6, 17, { t: 'n', v: 123.456, numFmt: '0.00%' }),
        w(6, 18, { t: 'n', v: 12345678901, numFmt: '0' }), w(6, 19, { t: 'd', v: 0.523842592593, numFmt: 'h:mm:ss' }),
        w(6, 20, { t: 'd', v: 46295, numFmt: 'yyyy"年"m"月"d"日"' }), w(6, 21, { t: 'n', v: -1234567.891, numFmt: '#,##0.00' }),
        w(6, 22, { t: 'd', v: 46295, numFmt: 'yyyy-mm-dd' }), w(6, 23, { t: 'd', v: 46295, numFmt: 'yyyy/m/d' }),
        w(7, 12, { t: 'd', v: 46295, numFmt: 'yyyy/m/d' }),
      ]),
    ],
  },
  // unmerge a title (Range.UnMerge + HorizontalAlignment = xlCenterAcrossSelection), then a move that the merge blocked
  { name: 'u-torture', fixture: 'torture.xlsx', ops: [unmerge('Data', ['A1:F1'], true), movC('Data', 3, 1, 8)] },
  { name: 'u-plain', fixture: 'torture.xlsx', ops: [unmerge('Data', ['A1:F1'], false)] },
  { name: 'u-gtable', fixture: 'g-table.xlsx', ops: [unmerge('名簿', ['I1:K1'], true), movC('名簿', 10, 1, 13)] },
  // typing over a table header cell renames the table column and every structured reference to it
  { name: 'h-rename', fixture: 'g-table.xlsx', ops: [cells('名簿', [w(3, 4, { t: 's', v: '得点' })])] },
  { name: 'h-dup', fixture: 'g-table.xlsx', ops: [cells('名簿', [w(3, 5, { t: 's', v: '点数1' })])] },
  { name: 'h-clear', fixture: 'g-table.xlsx', ops: [cells('名簿', [w(3, 5, null)])] },
  { name: 'h-special', fixture: 'g-table.xlsx', ops: [cells('名簿', [w(3, 5, { t: 's', v: 'a-b#c' })])] },
  { name: 'h-space', fixture: 'g-table.xlsx', ops: [cells('名簿', [w(3, 4, { t: 's', v: '得点 (前期)' })])] },
  { name: 'h-number', fixture: 'g-table.xlsx', ops: [cells('名簿', [w(3, 1, { t: 'n', v: 2024 })])] },
  { name: 'h-case', fixture: 'g-table.xlsx', ops: [cells('名簿', [w(3, 1, { t: 's', v: 'id' })])] },
  // rename, insert a table column before it, then delete the renamed column: its references become #REF!
  { name: 'h-then-ops', fixture: 'g-table.xlsx', ops: [cells('名簿', [w(3, 4, { t: 's', v: '得点' })]), insC('名簿', 2), delC('名簿', 5)] },
  // delete a column, then rename another one to the deleted column's name: references to the old one stay #REF!
  { name: 'h-after-del', fixture: 'g-table.xlsx', ops: [delC('名簿', 4), cells('名簿', [w(3, 4, { t: 's', v: '点数1' })])] },
  // r2-table.xlsx (com/make-fixtures-r2.ps1): totals row with a custom totalsRowFormula, calculated column, a defined
  // name and another sheet referring to the table columns
  { name: 'h2-rename', fixture: 'r2-table.xlsx', ops: [cells('T', [w(1, 2, { t: 's', v: 'Pts' })])] },
  { name: 'h2-dup', fixture: 'r2-table.xlsx', ops: [cells('T', [w(1, 3, { t: 's', v: 'score' })])] },
  { name: 'h2-special', fixture: 'r2-table.xlsx', ops: [cells('T', [w(1, 2, { t: 's', v: "x'y[1]#" })])] },
  // typing into the totals row: a label, a formula, a constant
  { name: 'h2-totals', fixture: 'r2-table.xlsx', ops: [cells('T', [w(5, 1, { t: 's', v: '計' }), w(5, 2, { t: 'z', v: null, f: 'MAX(Tbl[Score])' }), w(5, 3, { t: 'n', v: 5 })])] },
  { name: 'h2-totals-clear', fixture: 'r2-table.xlsx', ops: [cells('T', [w(5, 1, null), w(5, 3, null), w(5, 2, { t: 's', v: 'x' })])] },
  // SUBTOTAL(n, Tbl[own column]) becomes a built-in totals function; the same text in another column stays custom
  {
    name: 'h2-totals-fn', fixture: 'r2-table.xlsx',
    ops: [cells('T', [w(5, 1, { t: 'z', v: null, f: 'SUBTOTAL(109,Tbl[Score])' }), w(5, 2, { t: 'z', v: null, f: 'SUBTOTAL(104,Tbl[Score])' }), w(5, 3, { t: 'z', v: null, f: 'SUBTOTAL(103,Tbl[Twice])' })])],
  },
  // a constant over the totals formulas (the doc's delete-with-conversion path, Excel: Value2 = Value2)
  { name: 'h2-totals-values', fixture: 'r2-table.xlsx', ops: [cells('T', [w(5, 2, { t: 'n', v: 120 }), w(5, 3, { t: 'n', v: 120 })], 'toValues')] },
  // the written totals formula follows later ops; ops before the write are not applied to it
  {
    name: 'h2-totals-ops', fixture: 'r2-table.xlsx',
    ops: [cells('T', [w(5, 2, { t: 'z', v: null, f: 'SUM(B2:B4)' })]), insC('T', 1), insR('T', 3), cells('T', [w(6, 4, { t: 'z', v: null, f: 'SUM(D2:D5)/2' })])],
  },
  { name: 'h2-then-move', fixture: 'r2-table.xlsx', ops: [cells('T', [w(1, 2, { t: 's', v: '得点' })]), insC('T', 2), cells('T', [w(1, 2, { t: 's', v: '新' })])] },

  // ---- round 3 (r3-misc.xlsx, built by tools/make-r3-fixture.ts) ------------------------------------------------
  // 'unmerge' spreads the top-left cell's look: Excel's own UnMerge leaves the other cells empty, so the Excel side of
  // these cases (run-ops.ps1) copies Interior + border sides from the top-left cell after UnMerge, as the runner does
  // yellow band with a thin box on A1 only
  { name: 'u-band', fixture: 'r3-misc.xlsx', ops: [unmerge('Band', ['A1:F1'], false)] },
  { name: 'u-band-center', fixture: 'r3-misc.xlsx', ops: [unmerge('Band', ['A1:F1'], true)] },
  // 3 x 3, B4 date / C5 number format of their own
  { name: 'u-area', fixture: 'r3-misc.xlsx', ops: [unmerge('Band', ['A3:C5'], false)] },
  { name: 'u-theme', fixture: 'r3-misc.xlsx', ops: [unmerge('Band', ['E7:G7'], true)] }, // theme colour + tint fill, centred
  { name: 'u-nolook', fixture: 'r3-misc.xlsx', ops: [unmerge('Band', ['A9:C9'], false)] }, // top-left cell without fill / border
  // C11 has a border of its own
  { name: 'u-ownborder', fixture: 'r3-misc.xlsx', ops: [unmerge('Band', ['A11:C11'], false)] },
  // outlines as Excel stores them (left / top / bottom on the top-left cell of a band, left / top of a 3-row area) and a
  // theme-filled band with a medium red underline on its top-left cell only
  { name: 'u-outline', fixture: 'r3-misc.xlsx', ops: [unmerge('Outline', ['A1:F1', 'A3:C5'], false)] },
  { name: 'u-underline', fixture: 'r3-misc.xlsx', ops: [unmerge('Outline', ['A7:C7'], true)] },
  // round 4 (random op-log fuzz + Excel as the oracle: op-191-r2-table made Excel itself stop responding): deleting the totals
  // row of a table (r2-table.xlsx: Tbl A1:C5 = header, 3 data rows, totals row 5), alone and together with the last data row
  { name: 'h2-del-totals-row', fixture: 'r2-table.xlsx', ops: [delR('T', 5)] },
  { name: 'h2-del-last-data-and-totals', fixture: 'r2-table.xlsx', ops: [delR('T', 4, 2)] },
  // a structured reference to the totals row of a table that loses it; typing into the former totals row afterwards is an ordinary cell
  { name: 'h2-del-totals-ref', fixture: 'r2-table.xlsx', ops: [cells('U', [w(10, 1, { t: 'z', v: null, f: 'Tbl[[#Totals],[Score]]' })]), delR('T', 5)] },
  { name: 'h2-del-totals-then-write', fixture: 'r2-table.xlsx', ops: [delR('T', 5), cells('T', [w(4, 1, { t: 's', v: 'x' }), w(4, 2, { t: 'n', v: 5 })])] },
  // round 4: a value / a formula typed INTO the spill range of a dynamic array (not at its master): the spill is blocked
  // (#SPILL!), the file must not keep an array range over the occupied cell (unit tests: patch-round4-spill.test.ts)
  { name: 'r-spill-write', fixture: 'r2-arrays.xlsx', ops: [cells('A', [w(3, 5, { t: 'n', v: 7 }), w(3, 7, { t: 'z', v: null, f: '1+1' })])] },
  // CellWrite.quote honoured exactly: quote === true -> quotePrefix, anything else -> none (also over a prefixed cell)
  {
    name: 'q-text', fixture: 'r3-misc.xlsx',
    ops: [
      cells('Q', [
        w(2, 1, { t: 's', v: '00123', quote: true }), w(3, 1, { t: 's', v: '=text', quote: true }),
        w(4, 1, { t: 's', v: '03-1234-5678' }), w(5, 1, { t: 's', v: '00123', quote: true }), w(6, 1, { t: 's', v: '00123', quote: true }),
      ]),
      cells('Q', [w(6, 1, { t: 's', v: 'x' })]),
    ],
    known: { '[Q].styles.A6': QUOTE_COM },
  },
  {
    name: 'q-overwrite', fixture: 'r3-misc.xlsx',
    ops: [cells('Q', [w(1, 2, { t: 's', v: 'abc' }), w(2, 2, { t: 'n', v: 5 }), w(3, 2, null), w(1, 5, { t: 's', v: 'abc' })])],
    known: { '[Q].styles.B1': QUOTE_COM, '[Q].styles.E1': QUOTE_COM }, // B2 (a number over a prefixed cell) is no deviation any more: the prefix is invisible there, the snapshot only counts it on text
  },
  { name: 'q-numfmt', fixture: 'r3-misc.xlsx', ops: [cells('Q', [w(1, 3, { t: 's', v: '00123', numFmt: '@' }), w(1, 4, { t: 's', v: '00123', numFmt: '@', quote: true })])] },
  // round 4: Excel turns WrapText on by itself when a text constant with a line break is written (Value2 = "a" & vbLf & "b");
  // a formula with a line break in its text result, a text without a break and a cell that already wraps do not change
  {
    name: 'q-wrap', fixture: 'r3-misc.xlsx',
    ops: [
      cells('Q', [
        w(7, 1, { t: 's', v: 'a\nb' }), // LF (Alt+Enter), a cell without a style
        w(7, 2, { t: 's', v: 'a\r\nb' }), // CRLF
        w(7, 3, { t: 's', v: 'a\rb' }), // lone CR: Excel does NOT wrap (measured)
        w(5, 1, { t: 's', v: 'x\ny' }), // A5: bold, yellow, thin box — everything stays, wrap comes on top
        w(1, 3, { t: 's', v: 'a\nb' }), // C1: text format @
        w(2, 2, { t: 's', v: 'q\nr', quote: true }), // B2: quote prefix
        w(7, 4, { t: 's', v: 'a\nb', f: '"a"&CHAR(10)&"b"' }), // a formula: no wrap
        w(7, 5, { t: 's', v: 'a\nb', numFmt: '@' }), // text format set by the same write
        w(9, 1, { t: 's', v: 'plain' }), // no line break: no wrap
        w(10, 1, { t: 's', v: 'one\ntwo' }),
      ]),
      cells('Q', [w(10, 1, { t: 's', v: 'three\nfour' }), w(11, 1, { t: 'n', v: 12 })]), // already wrapping: unchanged; a number: no wrap
    ],
  },
  // a width the user sets after a number format widened the column: Excel drops bestFit, the next (wider) date does not widen it
  {
    name: 'w-after-auto', fixture: 'r2-arrays.xlsx',
    ops: [
      cells('A', [w(2, 9, { t: 'd', v: 46295, numFmt: 'yyyy/m/d' })]),
      { kind: 'colWidth', sheet: 'A', col: 9, width: 9 },
      cells('A', [w(3, 9, { t: 'd', v: 46305, numFmt: 'yyyy/m/d' })]),
    ],
  },
  // a constant typed into a totals cell is the column's label; a quote flag keeps the prefix on that cell like Excel
  { name: 'h2-totals-q', fixture: 'r2-table.xlsx', ops: [cells('T', [w(5, 1, { t: 's', v: '5', quote: true }), w(5, 3, { t: 's', v: '計' })])] },
  // names of inserted table columns (Excel 365: left neighbour 'Q1' -> 'Q2' when free, else 'Q12', 'Q13' ...; no number -> 列N)
  { name: 'n-q1', fixture: 'r3-misc.xlsx', ops: [insC('Tb', 3)] }, // right of Q1, Q2 exists
  { name: 'n-q3', fixture: 'r3-misc.xlsx', ops: [insC('Tb', 5)] }, // right of Q3, Q4 free
  { name: 'n-id2', fixture: 'r3-misc.xlsx', ops: [insC('Tb', 2, 2)] }, // right of ID
  { name: 'n-q1x3', fixture: 'r3-misc.xlsx', ops: [insC('Tb', 3, 3)] },
  { name: 'n-h1', fixture: 'g-table.xlsx', ops: [insC('名簿', 5)] }, // right of 点数1 (点数2 exists)
  { name: 'n-h1x2', fixture: 'g-table.xlsx', ops: [insC('名簿', 5, 2)] },
];

function w(row: number, col: number, cell: CellWrite['cell']): CellWrite {
  return { row, col, cell };
}
function cells(sheet: string, writes: CellWrite[], reason: 'edit' | 'toValues' = 'edit'): LoggedOp {
  return { kind: 'cells', sheet, writes, reason };
}
function unmerge(sheet: string, refs: string[], centerAcross: boolean): LoggedOp {
  return { kind: 'unmerge', sheet, refs, centerAcross };
}
