// Synthetic inspect dumps for unit-testing tools/lib/{compare-core,verdict}.mjs without Excel.
import type { InspectResult, InspectSheet } from '../../tools/lib/inspect.mjs';

export interface SheetSpec {
  name: string;
  /** A1 -> value (number / string / boolean) or '#REF!'-like error when it starts with '#' and ends with '!' / '?' / 'A' */
  cells?: Record<string, string | number | boolean>;
  formulas?: Record<string, string>;
  extra?: Partial<InspectSheet>;
}

const isError = (v: unknown) => typeof v === 'string' && /^#(REF!|NAME\?|N\/A|VALUE!|DIV\/0!|NUM!|NULL!|SPILL!|CALC!)$/.test(v);

export function sheet(spec: SheetSpec): InspectSheet {
  const values: Record<string, string | number | boolean> = {};
  const errors: Record<string, string> = {};
  const errorCounts: Record<string, number> = {};
  let maxRow = 1;
  let maxCol = 1;
  const rowText = new Array<number>(40).fill(0);
  const all = { ...(spec.cells ?? {}) };
  for (const a of Object.keys(spec.formulas ?? {})) if (!(a in all)) all[a] = 0;
  for (const [a, v] of Object.entries(all)) {
    const m = /^([A-Z]+)(\d+)$/.exec(a)!;
    const row = Number(m[2]);
    let col = 0;
    for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64;
    maxRow = Math.max(maxRow, row);
    maxCol = Math.max(maxCol, col);
    if (isError(v)) {
      errors[a] = v as string;
      errorCounts[v as string] = (errorCounts[v as string] ?? 0) + 1;
    } else {
      values[a] = v;
      if (typeof v === 'string' && row <= 40) rowText[row - 1]++;
    }
  }
  const formulas = spec.formulas ?? {};
  return {
    name: spec.name,
    index: 1,
    visible: -1,
    usedRange: `A1:${String.fromCharCode(64 + Math.min(maxCol, 26))}${maxRow}`,
    maxRow,
    maxCol,
    capped: false,
    formulaCount: Object.keys(formulas).length,
    refTextCount: Object.values(formulas).filter((f) => f.includes('#REF!')).length,
    errorCounts,
    rowTextCounts: rowText,
    firstRow: 1,
    firstCol: 1,
    formulas,
    values,
    errors,
    probes: {},
    arrays: [],
    spills: [],
    merges: [],
    cols: [],
    hiddenColumns: [],
    freeze: { frozen: false, splitRow: 0, splitCol: 0 },
    autoFilter: null,
    cf: [],
    dv: [],
    tables: [],
    charts: [],
    pivots: [],
    shapes: 0,
    comments: 0,
    threadedComments: 0,
    hyperlinks: 0,
    ...(spec.extra ?? {}),
  };
}

export function book(sheets: InspectSheet[], extra: Partial<InspectResult> = {}): InspectResult {
  return {
    tool: 'excel-merger/inspect',
    version: 1,
    file: 'test.xlsx',
    path: 'C:\\test.xlsx',
    openMode: 'normal',
    names: [],
    sheets,
    ...extra,
  };
}
