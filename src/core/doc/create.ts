/**
 * createEditor: the MergedDoc (column-major) built from the premerged package as READ BACK by readBook.
 *
 * Column ids (stable across moves):
 *  - imported columns (layout.imported): 'i:<sourceId>:<sourceCol>'
 *  - base columns: 'b<n>', n = column number in the ORIGINAL base sheet (columns right of the imported block map
 *    back by subtracting the number of imported columns left of them)
 * Other sheets contribute their formula cells (externalFormulas, with the cached value kept aside for
 * "convert to values") and their CF / DV rules (externalRules).
 */
import type { Cell, DocColumn, EditorSnapshot, ExternalFormulaCell, MergedDoc, SheetGrid, SourceBook } from '../types';
import type { PremergeLayout } from '../contracts';
import { setDefaultColWidth } from './autowiden';
import { precompile, setCached } from './util';

function findSheet(book: SourceBook, name: string): SheetGrid | undefined {
  return book.sheets.find((s) => s.name === name) ?? book.sheets.find((s) => s.name.toLowerCase() === name.toLowerCase());
}

export function buildEditor(
  premerged: SourceBook,
  layout: PremergeLayout,
  sourceInfo: MergedDoc['sourceInfo'],
  region: { headerRow: number; firstDataRow: number; lastDataRow: number },
): EditorSnapshot {
  const sheet = findSheet(premerged, layout.sheet);
  if (!sheet) throw new Error(`結合結果のシート「${layout.sheet}」が見つかりません。`);

  // ---- geometry
  const importedByCol = new Map<number, { sourceId: string; sourceCol: number }>();
  let maxImported = 0;
  for (const im of layout.imported) {
    if (!(im.col >= 1)) continue;
    importedByCol.set(im.col, { sourceId: im.sourceId, sourceCol: im.sourceCol });
    maxImported = Math.max(maxImported, im.col);
  }
  const importedCols = [...importedByCol.keys()].sort((a, b) => a - b);
  const nCols = Math.max(sheet.maxCol, maxImported, 0);
  let lastDataRow = Math.max(region.lastDataRow, region.firstDataRow - 1);
  for (const a of layout.appended) lastDataRow = Math.max(lastDataRow, a.row);
  const rowCount = Math.max(sheet.maxRow, lastDataRow, region.headerRow, 0);

  // ---- columns (column-major copy of the row-major grid; holey arrays keep sparse sheets cheap)
  const cellsByCol: Array<Array<Cell | undefined>> = [];
  for (let c = 0; c < nCols; c++) {
    const arr = new Array<Cell | undefined>(rowCount);
    cellsByCol.push(arr);
  }
  for (let r = 0; r < rowCount; r++) {
    const row = sheet.rows[r];
    if (!row) continue;
    const lim = Math.min(row.length, nCols);
    for (let c = 0; c < lim; c++) {
      const cell = row[c];
      if (cell) cellsByCol[c][r] = cell;
    }
  }
  const used = new Set<string>();
  const unique = (id: string) => {
    let out = id;
    for (let k = 2; used.has(out); k++) out = `${id}#${k}`;
    used.add(out);
    return out;
  };
  let impIdx = 0;
  const columns: DocColumn[] = [];
  for (let c = 1; c <= nCols; c++) {
    while (impIdx < importedCols.length && importedCols[impIdx] < c) impIdx++;
    const imp = importedByCol.get(c);
    const props = { ...(sheet.colProps[c - 1] ?? {}) };
    if (imp) {
      columns.push({
        id: unique(`i:${imp.sourceId}:${imp.sourceCol}`),
        origin: { kind: 'import', sourceId: imp.sourceId, col: imp.sourceCol },
        props,
        cells: cellsByCol[c - 1],
      });
    } else {
      const baseCol = c - impIdx; // imported columns strictly left of c
      columns.push({ id: unique(`b${baseCol}`), origin: { kind: 'base', col: baseCol }, props, cells: cellsByCol[c - 1] });
    }
  }

  // ---- other sheets: formulas (column-major per sheet, so fill-down neighbours share tokenization) and rules
  const externalFormulas: ExternalFormulaCell[] = [];
  const externalRules: MergedDoc['externalRules'] = [];
  for (const sh of premerged.sheets) {
    if (sh === sheet) continue;
    const list: ExternalFormulaCell[] = [];
    sh.rows.forEach((row, ri) => {
      if (!row) return;
      row.forEach((cell, ci) => {
        if (!cell || cell.f === undefined) return;
        const e: ExternalFormulaCell = { sheet: sh.name, row: ri + 1, col: ci + 1, f: cell.f };
        setCached(e, cell.arr !== undefined ? { t: cell.t, v: cell.v, arr: cell.arr } : { t: cell.t, v: cell.v });
        list.push(e);
      });
    });
    list.sort((a, b) => a.col - b.col || a.row - b.row);
    externalFormulas.push(...list);
    for (const rule of sh.features?.rules ?? []) externalRules.push({ sheet: sh.name, rule });
  }

  const f = sheet.features;
  const doc: MergedDoc = {
    sheetName: sheet.name,
    sheetNames: premerged.sheets.map((s) => s.name),
    rowCount,
    headerRow: region.headerRow,
    firstDataRow: region.firstDataRow,
    lastDataRow,
    columns,
    rowProps: sheet.rowProps.slice(0, rowCount),
    styles: premerged.styles,
    features: {
      merges: f?.merges ?? [],
      rules: f?.rules ?? [],
      tables: f?.tables ?? [],
      arrays: f?.arrays ?? [],
      hyperlinks: f?.hyperlinks ?? [],
      ...(f?.autoFilter !== undefined ? { autoFilter: f.autoFilter } : {}),
      ...(f?.freeze !== undefined ? { freeze: f.freeze } : {}),
    },
    externalFormulas,
    externalRules,
    names: premerged.names,
    sourceInfo,
    version: 1,
  };
  setDefaultColWidth(doc, sheet.defaultColWidth);
  // tokenize every formula once now (fill-down copies reuse their neighbour's tokens), so the first op is fast
  precompile(doc);
  return { doc, log: [], label: '結合' };
}
