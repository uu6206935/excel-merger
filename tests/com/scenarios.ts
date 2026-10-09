/**
 * Fixtures (authored by Excel itself, tests/com/ps/make-fixture.ps1) and op logs shared by the real-Excel tests:
 *   features.test.ts  detailed assertions on what the runner does for one log (scenario "r2")
 *   invariant.test.ts the JS-vs-Excel invariant (same log through the runner and through replayLog) over every scenario
 */
import { mapIndex, numToCol, type StructOp } from '../../src/core/formula/excelRef';
import type { LoggedOp } from '../../src/core/types';
import { cellAt, type Dump, type FixtureSpec, type SpecCell } from './lib';

export interface Scenario {
  name: string;
  /** what the log exercises (printed with the report) */
  covers: string;
  /** sheet the log mainly works on */
  main: string;
  spec: () => FixtureSpec;
  build: (before: Dump) => { log: LoggedOp[] };
}

/** Builds a log step by step and keeps track of where an original / inserted column of a sheet is now. */
export class LogBuilder {
  log: LoggedOp[] = [];
  private ops: Array<{ sheet: string; op: StructOp }> = [];
  constructor(private readonly main: string) {}
  push(e: LoggedOp) {
    this.log.push(e);
    if (e.kind === 'struct') this.ops.push({ sheet: e.sheet, op: e.op });
  }
  struct(sheet: string, op: StructOp) {
    this.push({ kind: 'struct', sheet, op });
  }
  /** mark = number of struct ops at the time the column position `col` was valid */
  mark() {
    return this.ops.length;
  }
  /** where column `col` (valid after `mark` struct ops) is now; throws when it was deleted */
  col(col: number, mark = 0, sheet = this.main): number {
    let x: number | null = col;
    for (const { sheet: s, op } of this.ops.slice(mark)) if (x != null && s === sheet && op.axis === 'col') x = mapIndex(x, op);
    if (x == null) throw new Error(`column ${col} was deleted`);
    return x;
  }
  /** same for rows */
  row(row: number, mark = 0, sheet = this.main): number {
    let x: number | null = row;
    for (const { sheet: s, op } of this.ops.slice(mark)) if (x != null && s === sheet && op.axis === 'row') x = mapIndex(x, op);
    if (x == null) throw new Error(`row ${row} was deleted`);
    return x;
  }
}

// ---------------------------------------------------------------------------------------------------------
// Scenario r2: unmerge + centre across, numFmt / Formula2 / Formula writes, toValues over arrays, table move, widths
// ---------------------------------------------------------------------------------------------------------

export const MAIN = '名簿';
export const SUM = '集計';
export const TSHEET = '商品';

/** Original layout of 名簿 (row 1 = title merged A1:F1, row 2 = headers, rows 3..12 data, row 13 totals). */
export const HDR = ['ID', '氏名', '点数1', '点数2', '合計', '比率', '係数', '区分', '入社日', '倍', '順位', '係数x10'];

export function r2Spec(): FixtureSpec {
  const cells: SpecCell[] = [['A1', 's', '2026年度 社員名簿']];
  HDR.forEach((h, i) => cells.push([`${numToCol(i + 1)}2`, 's', h]));
  for (let r = 3; r <= 12; r++) {
    cells.push([`A${r}`, 'n', 1000 + r], [`B${r}`, 's', `氏名${r}`], [`C${r}`, 'n', ((r * 37) % 90) + 5], [`D${r}`, 'n', r * 3]);
    cells.push([`E${r}`, 'f', `=C${r}+D${r}`], [`F${r}`, 'f', `=E${r}/$E$13`], [`G${r}`, 'n', (r % 3) + 1], [`H${r}`, 's', r % 2 ? '営業' : '開発']);
    cells.push([`I${r}`, 'n', 45000 + r * 31], [`L${r}`, 'f', `=G${r}*10`]);
  }
  cells.push(['J3:J12', 'a', '=C3:C12*G3:G12']);
  cells.push(['K3', 'd', '=SORT(C3:C12,1,-1)']);
  cells.push(['A13', 's', '計'], ['C13', 'f', '=SUM(C3:C12)'], ['D13', 'f', '=SUM(D3:D12)'], ['E13', 'f', '=SUM(E3:E12)']);
  cells.push(['J13', 'f', '=SUM(J3:J12)'], ['K13', 'f', '=SUM(K3#)'], ['L13', 'f', '=SUMPRODUCT(G3:G12,C3:C12)']);
  const sum: SpecCell[] = [
    ['A1', 's', '集計'],
    ['B1', 'f', `=SUM(${MAIN}!C3:C12)`],
    ['B2', 'f', `=${MAIN}!E13`],
    ['B3', 'f', '=SUM(点数範囲)'],
    ['B4', 'f', `=SUM(${MAIN}!J3:J12)`],
    ['B5', 'f', `=INDEX(${MAIN}!A3:L12,2,5)`],
    ['B6', 'f', `=COUNTIF(${MAIN}!H3:H12,"営業")`],
    ['B7', 'f', '=SUM(Tbl[金額])'],
    ['B8', 'f', '=INDEX(Tbl[単価],2)'],
    ['B9', 'f', `=${MAIN}!I3`],
  ];
  const tcells: SpecCell[] = [];
  ['コード', '品名', '単価', '数量', '金額'].forEach((h, i) => tcells.push([`${numToCol(i + 1)}1`, 's', h]));
  for (let r = 2; r <= 6; r++) {
    tcells.push([`A${r}`, 's', `P-${r}`], [`B${r}`, 's', `品名${r}`], [`C${r}`, 'n', r * 100], [`D${r}`, 'n', r]);
    tcells.push([`E${r}`, 'f', '=[@単価]*[@数量]']);
  }
  tcells.push(['G1', 's', '合計'], ['G2', 'f', '=SUM(Tbl[金額])'], ['G3', 'f', '=SUM(Tbl[[単価]:[数量]])'], ['G4', 'f', '=COUNTA(Tbl[品名])']);
  return {
    sheets: [
      {
        name: MAIN,
        cells,
        merges: ['A1:F1'],
        align: [['A1:F1', 'center']],
        widths: [[2, 14]],
        formats: [['I3:I12', 'yyyy/m/d']],
        cf: [
          { range: 'C3:C12', type: 'expression', formula: '=C3>60' },
          { range: 'E3:E12', type: 'cellIs', operator: 5, formula: '=50' },
          { range: 'D3:D12', type: 'colorScale' },
        ],
        dv: [
          { range: 'H3:H12', type: 'list', formula: '営業,開発' },
          { range: 'C3:C12', type: 'whole', formula: '0', formula2: '1000' },
        ],
      },
      { name: SUM, cells: sum },
      { name: TSHEET, cells: tcells, tables: [{ name: 'Tbl', range: 'A1:E6' }] },
    ],
    names: [
      { name: '点数範囲', refersTo: `=${MAIN}!$C$3:$C$12` },
      { name: '係数列', refersTo: `=${MAIN}!$G$3:$G$12` },
    ],
  };
}

export function buildR2Log(before: Dump) {
  const b = new LogBuilder(MAIN);
  const v = (r: number, c: number) => cellAt(before, MAIN, r, c)?.v as number;
  b.push({ kind: 'unmerge', sheet: MAIN, refs: ['A1:F1'], centerAcross: true });
  b.struct(MAIN, { type: 'move', axis: 'col', from: 3, count: 1, to: 7 }); // 点数1 -> before 係数 (crosses the former title)
  const insMark = b.mark();
  b.struct(MAIN, { type: 'insert', axis: 'col', at: 13, count: 3 }); // three new columns after 係数x10
  const M = () => b.col(13, insMark + 1);
  const N = () => b.col(14, insMark + 1);
  const P = () => b.col(15, insMark + 1);
  b.push({
    kind: 'cells',
    sheet: MAIN,
    reason: 'edit',
    writes: [
      { row: 3, col: M(), cell: { t: 'd', v: 46000, numFmt: 'yyyy/m/d' } },
      { row: 4, col: M(), cell: { t: 'n', v: 0.25, numFmt: '0%' } },
      { row: 5, col: M(), cell: { t: 'n', v: 1234, numFmt: '"¥"#,##0' } },
      { row: 6, col: M(), cell: { t: 'n', v: 1, f: '_xlfn.SEQUENCE(3)', dyn: true } },
      { row: 10, col: M(), cell: { t: 'n', v: 0, f: 'C3:C12' } }, // legacy: implicit intersection (=@C3:C12)
      { row: 11, col: M(), cell: { t: 's', v: '001', numFmt: '@' } },
      { row: 3, col: N(), cell: { t: 'n', v: 0, f: `SUM(${numToCol(M())}3:${numToCol(M())}5)` } },
      { row: 4, col: N(), cell: { t: 'n', v: 0, f: `_xlfn.XLOOKUP(A4,A3:A12,B3:B12)` } },
      { row: 5, col: N(), cell: { t: 'b', v: true } },
      { row: 6, col: N(), cell: { t: 'e', v: '#N/A' } },
      { row: 7, col: N(), cell: { t: 's', v: '=text', quote: true } }, // the doc flags text Excel would parse: quotePrefix in both engines
      // one column, one number format over a vertical run (one Range.NumberFormat call): dates widen the standard-width
      // column like typing them; 123,456,789 needs more than the date
      ...[3, 4, 5, 6].map((row, i) => ({ row, col: P(), cell: { t: 'd' as const, v: 46000 + i * 31, numFmt: 'yyyy/m/d' } })),
      ...[7, 8].map((row) => ({ row, col: P(), cell: { t: 'n' as const, v: 123456789, numFmt: '#,##0' } })),
    ],
  });
  // delete 係数 with 「参照している数式を値に変換」: the CSE array 倍 (J3:J12), the spill 順位 (K3#, master only) and
  // 係数x10 are converted first (the doc logs the array / spill MASTER only; the whole array becomes values)
  const toValues: LoggedOp = {
    kind: 'cells',
    sheet: MAIN,
    reason: 'toValues',
    writes: [
      // the CSE array 倍: every cell of it (a multi-cell run inside an array) ...
      ...Array.from({ length: 10 }, (_, i) => ({ row: 3 + i, col: b.col(10), cell: { t: 'n' as const, v: v(3 + i, 10) } })),
      // ... the spill 順位: its master only (the whole spill becomes values)
      { row: 3, col: b.col(11), cell: { t: 'n', v: v(3, 11) } },
      ...Array.from({ length: 10 }, (_, i) => ({ row: 3 + i, col: b.col(12), cell: { t: 'n' as const, v: v(3 + i, 12) } })),
    ],
  };
  b.push(toValues);
  b.struct(MAIN, { type: 'delete', axis: 'col', at: b.col(7), count: 1 });
  b.struct(MAIN, { type: 'move', axis: 'col', from: 1, count: 2, to: b.col(9) }); // ID, 氏名 -> before 入社日
  b.struct(TSHEET, { type: 'move', axis: 'col', from: 4, count: 1, to: 3 }); // inside the table: 数量 before 単価
  b.push({ kind: 'colWidth', sheet: MAIN, col: b.col(2), width: 20 });
  b.push({ kind: 'colWidth', sheet: MAIN, col: b.col(5), width: 12.625 });
  b.struct(MAIN, { type: 'insert', axis: 'col', at: 1, count: 1 });
  return { log: b.log, b, M: M(), N: N(), P: P() };
}

// ---------------------------------------------------------------------------------------------------------
// Scenario cfdv: CF / DV / autofilter / hidden column / row ops / #REF! on delete / block moves to the edges
// ---------------------------------------------------------------------------------------------------------

export const DATA = 'データ';
export const SUMMARY = 'まとめ';
/** headers of the cfdv sheet (row 1); data rows 2..21; totals row 22 */
export const CF_HDR = ['ID', '氏名', '点数', '評価', '得点率', '区分', '備考', '合計', '日付', '倍率'];

export function cfdvSpec(): FixtureSpec {
  const cells: SpecCell[] = [];
  CF_HDR.forEach((h, i) => cells.push([`${numToCol(i + 1)}1`, 's', h]));
  for (let r = 2; r <= 21; r++) {
    cells.push([`A${r}`, 'n', 2000 + r], [`B${r}`, 's', `氏名${r}`], [`C${r}`, 'n', ((r * 53) % 100) + 1], [`D${r}`, 'n', (r * 7) % 5]);
    cells.push([`E${r}`, 'f', `=C${r}/100`], [`F${r}`, 's', ['A', 'B', 'C'][r % 3]], [`G${r}`, 's', r % 4 ? `メモ${r}` : '']);
    cells.push([`H${r}`, 'f', `=C${r}+D${r}*10`], [`I${r}`, 'n', 45900 + r], [`J${r}`, 'f', `=E${r}*2`]);
  }
  cells.push(['A22', 's', '計'], ['C22', 'f', '=SUM(C2:C21)'], ['D22', 'f', '=SUM(D2:D21)'], ['H22', 'f', '=SUM(H2:H21)'], ['J22', 'f', '=AVERAGE(J2:J21)']);
  const sum: SpecCell[] = [
    ['A1', 's', 'まとめ'],
    ['B1', 'f', `=SUM(${DATA}!C2:C21)`],
    ['B2', 'f', `=${DATA}!H22`],
    ['B3', 'f', `=SUM(${DATA}!C:C)`],
    ['B4', 'f', '=SUM(得点)'],
    ['B5', 'f', `=INDEX(${DATA}!A2:J21,3,3)`],
    ['B6', 'f', `=COUNTIF(${DATA}!F2:F21,"A")`],
    ['B7', 'f', `=SUMPRODUCT(${DATA}!C2:C21,${DATA}!D2:D21)`],
    ['B8', 'f', `=${DATA}!$G$2&"/"&${DATA}!$B$2`],
    ['B9', 'f', `=SUM(${DATA}!$C$2:$E$21)`],
    ['B10', 'f', `=MAX(${DATA}!J2:J21)`],
  ];
  return {
    sheets: [
      {
        name: DATA,
        cells,
        widths: [[2, 14], [7, 20]],
        formats: [['E2:E21', '0%'], ['I2:I21', 'yyyy/m/d']],
        hidden: [7],
        autoFilter: 'A1:J22', // Excel's AutoFilter adopts the whole current region (the totals row is adjacent)
        cf: [
          { range: 'C2:C21', type: 'expression', formula: '=C2>80' },
          { range: 'D2:D21', type: 'cellIs', operator: 5, formula: '=3' },
          { range: 'E2:E21', type: 'colorScale' },
          { range: 'J2:J21', type: 'dataBar' },
          { range: 'C2:E21', type: 'expression', formula: '=$C2>90' },
        ],
        dv: [
          { range: 'F2:F21', type: 'list', formula: 'A,B,C' },
          { range: 'C2:C21', type: 'whole', formula: '0', formula2: '100' },
        ],
      },
      { name: SUMMARY, cells: sum },
    ],
    names: [
      { name: '得点', refersTo: `=${DATA}!$C$2:$C$21` },
      { name: '備考列', refersTo: `=${DATA}!$G$2:$G$21` },
      { name: '全表', refersTo: `=${DATA}!$A$1:$J$22` },
    ],
  };
}

export function buildCfdvLog(_before: Dump) {
  const b = new LogBuilder(DATA);
  // a block of two columns to the right across CF / DV areas and the hidden 備考 (点数, 評価 -> before 合計)
  b.struct(DATA, { type: 'move', axis: 'col', from: 3, count: 2, to: 8 });
  // the last column to the very start
  b.struct(DATA, { type: 'move', axis: 'col', from: 10, count: 1, to: 1 });
  // insert INSIDE the moved block (between 点数 and 評価: inside the C:E conditional format), and right after the
  // DV column 区分 (format inheritance / "leak" of CF and DV areas)
  b.struct(DATA, { type: 'insert', axis: 'col', at: b.col(4), count: 1 });
  b.struct(DATA, { type: 'insert', axis: 'col', at: b.col(6) + 1, count: 1 });
  // a written value (+ formula) and widths; the inserted column is the one just left of 評価
  const ins = () => b.col(4) - 1;
  b.push({ kind: 'cells', sheet: DATA, reason: 'edit', writes: [{ row: 2, col: ins(), cell: { t: 's', v: '挿入' } }, { row: 3, col: ins(), cell: { t: 'n', v: 7, f: 'C3*2' } }] });
  b.push({ kind: 'colWidth', sheet: DATA, col: b.col(2), width: 30.5 });
  b.push({ kind: 'colWidth', sheet: DATA, col: b.col(9), width: 11 });
  // rows: insert two rows inside the data, delete one, move two
  b.struct(DATA, { type: 'insert', axis: 'row', at: 5, count: 2 });
  b.struct(DATA, { type: 'delete', axis: 'row', at: 3, count: 1 });
  b.struct(DATA, { type: 'move', axis: 'row', from: 12, count: 2, to: 4 });
  // delete 得点率 (倍率 / names / the summary sheet refer to it: #REF! in both engines) and the hidden 備考
  b.struct(DATA, { type: 'delete', axis: 'col', at: b.col(5), count: 1 });
  b.struct(DATA, { type: 'delete', axis: 'col', at: b.col(7), count: 1 });
  // a last block move of three columns to the end (the used range has 10 columns now)
  b.struct(DATA, { type: 'move', axis: 'col', from: 2, count: 3, to: 11 });
  return { log: b.log, b };
}

export const SCENARIOS: Scenario[] = [
  {
    name: 'r2',
    covers: 'unmerge + centre across, numFmt / Formula2 / Formula writes, toValues over a CSE array and a spill, delete, table move, inserts, widths',
    main: MAIN,
    spec: r2Spec,
    build: buildR2Log,
  },
  {
    name: 'cfdv',
    covers: 'CF (expression / cellIs / colorScale / dataBar) + DV + autofilter + hidden column across block moves, inserts inside / after the areas, row insert / delete / move, #REF! on delete, names',
    main: DATA,
    spec: cfdvSpec,
    build: buildCfdvLog,
  },
];
