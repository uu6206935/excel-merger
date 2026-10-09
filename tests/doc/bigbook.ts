// 10,000 rows x 60 columns with 50,000 formulas (performance fixture, shared by perf tests / profiling).
import { numToCol } from '../../src/core/formula/excelRef';
import type { Cell, SheetGrid } from '../../src/core/types';
import { makeBook } from './helpers';

export const ROWS = 10_000;
export const COLS = 60;

export function bigBook() {
  const rows: SheetGrid['rows'] = [];
  const header: Cell[] = [];
  for (let c = 1; c <= COLS; c++) header.push({ t: 's', v: `列${c}`, s: 1 });
  rows.push(header);
  for (let r = 2; r <= ROWS + 1; r++) {
    const row: Cell[] = [];
    for (let c = 1; c <= 55; c++) row.push(c % 3 === 0 ? { t: 's', v: `v${r}-${c}` } : { t: 'n', v: r * c, s: 2 });
    row.push({ t: 'n', v: 0, f: `A${r}*2` });
    row.push({ t: 'n', v: 0, f: `SUM(B${r}:E${r})` });
    row.push({ t: 'n', v: 0, f: `IF(C${r}>0,D${r},"")` });
    row.push({ t: 'n', v: 0, f: `VLOOKUP(A${r},$A$2:$F$${ROWS + 1},3,FALSE)` });
    row.push({ t: 's', v: '', f: `F${r}&G${r}` });
    rows.push(row);
  }
  const data: SheetGrid = {
    name: 'Data', state: 'visible', maxRow: ROWS + 1, maxCol: COLS, rows, colProps: [], rowProps: [],
    features: { merges: [], rules: [{ sqref: `C2:C${ROWS + 1}`, formulas: ['C2>100'], kind: 'cf' }], tables: [], arrays: [], hyperlinks: [], freeze: { cols: 1, rows: 1 }, autoFilter: `A1:${numToCol(COLS)}${ROWS + 1}` },
  };
  const book = makeBook({ sheets: [{ name: '集計' }], styles: [{}, { bold: true }, { numFmt: '#,##0' }] });
  const sum: SheetGrid['rows'] = [];
  for (let r = 1; r <= 1000; r++) sum.push([{ t: 'n', v: 0, f: `SUM(Data!${numToCol((r % 50) + 1)}2:${numToCol((r % 50) + 1)}${ROWS + 1})` }]);
  book.sheets = [data, { ...book.sheets[0], rows: sum, maxRow: 1000, maxCol: 1 }];
  book.names = [{ name: 'データ', scope: null, refersTo: `Data!$A$1:$BH$${ROWS + 1}` }];
  return book;
}

