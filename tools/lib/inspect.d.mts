// Types for tools/lib/inspect.mjs and the JSON written by tools/inspect.ps1 (owner: T).

export interface InspectCf {
  type: number | null;
  appliesTo: string | null;
  operator: number | null;
  formula1: string | null;
  formula2: string | null;
  priority: number | null;
}

export interface InspectDv {
  range: string;
  type: number | null;
  operator: number | null;
  formula1: string | null;
  formula2: string | null;
  alertStyle: number | null;
}

export interface InspectSheet {
  name: string;
  index: number;
  /** -1 visible, 0 hidden, 2 very hidden */
  visible: number;
  usedRange: string;
  maxRow: number;
  maxCol: number;
  /** true when only the first rows were dumped (maxCells) */
  capped: boolean;
  formulaCount: number;
  /** formulas whose text contains '#REF!' */
  refTextCount: number;
  /** error VALUES per kind ('#REF!': 3 …) */
  errorCounts: Record<string, number>;
  /** text cells per row, first 40 rows of the used range (header-row guess) */
  rowTextCounts: number[];
  firstRow: number;
  firstCol: number;
  /** A1 -> Formula2 text ('=SUM(A1:A3)') */
  formulas: Record<string, string>;
  /** A1 -> Value2 (number / string / boolean), errors excluded */
  values: Record<string, string | number | boolean>;
  /** A1 -> error text ('#REF!') */
  errors: Record<string, string>;
  /** 'A1' -> 'numFmt|b1|fc0|bg<color or none>|h<align>|sz<size>' for row 1, the header row and the row below */
  probes: Record<string, string>;
  arrays: string[];
  spills: string[];
  merges: string[];
  cols: Array<{ col: number; letter: string; width: number; hidden: boolean }>;
  hiddenColumns: string[];
  freeze: { frozen: boolean; splitRow: number; splitCol: number };
  autoFilter: string | null;
  cf: InspectCf[];
  dv: InspectDv[];
  tables: Array<{ name: string; range: string; showTotals: boolean; columns: string[] }>;
  charts: Array<{ name: string; type: number | null; title: string | null; topLeft: string | null; series: Array<{ name: string | null; formula: string | null }> }>;
  pivots: Array<{ name: string; range: string | null; source: string | null; rowFields: string[]; columnFields: string[]; dataFields: string[] }>;
  shapes: number;
  comments: number;
  threadedComments: number | null;
  hyperlinks: number;
}

/** tools/inspect.ps1 Probe-Cells: one cell as Excel shows it */
export interface CellProbe {
  /** Range.NumberFormat (English codes, e.g. 'yyyy/m/d', '0%', 'General') */
  numFmt?: string;
  numFmtLocal?: string | null;
  /** Range.Text: what the cell displays ('2026/4/1', '50%', '####' when too narrow) */
  text?: string;
  value?: string | number | boolean | null;
  error?: string;
  /** Formula2 ('=UNIQUE(D4:D63)') or null */
  formula?: string | null;
  /** HorizontalAlignment (-4108 center, 7 centerContinuous = 選択範囲内で中央, 1 general …) */
  hAlign?: number;
  bold?: boolean;
  fontSize?: number;
  /** Font.Color (BGR integer as COM returns it: 16777215 = white) */
  fontColor?: number | null;
  /** Interior.Color (BGR integer) or null when the cell has no fill (Interior.Pattern = xlNone) */
  fill?: number | null;
  /** MergeArea address when the cell is merged */
  merge?: string | null;
  hasSpill?: boolean | null;
  /** SpillingToRange address when this cell is a spill master */
  spill?: string | null;
  /** PrefixCharacter ("'" when the cell has a quote prefix) */
  prefix?: string | null;
  probeError?: string;
}

export interface InspectResult {
  tool: 'excel-merger/inspect';
  version: 1;
  file: string;
  path: string;
  openMode: 'normal' | 'repair' | 'extract' | 'failed';
  openError?: string;
  /** removedRecord / repairedRecord lines of Excel's repair log for this file */
  repairLog?: string[];
  /** set when the inspection itself failed */
  error?: string;
  /** Excel starts used for this file (1 = the first attempt was consistent with the package; > 1 = it was inspected again) */
  attempts?: number;
  /** reasons why the dump still disagrees with the package after the last attempt (tools/lib/dump-check.mjs) */
  suspicious?: string[];
  /** reads Excel kept rejecting ("busy") in the last attempt */
  transientFailures?: number;
  /** Excel never reported Ready / calculation finished within the wait */
  notReady?: boolean;
  /** Excel opened another workbook than the one asked for (its full name) */
  wrongWorkbook?: string;
  /** 'read/expected': the defined names read differ from Names.Count */
  nameCountMismatch?: string;
  recalculated?: boolean;
  excelVersion?: string;
  date1904?: boolean;
  fileFormat?: number;
  names?: Array<{ name: string; scope: string | null; refersTo: string | null; visible: boolean | null }>;
  sheets?: InspectSheet[];
  chartSheets?: string[];
  /** only with opts.probeCells: 'Sheet!A1' -> probe */
  cellProbes?: Record<string, CellProbe>;
}

export const INSPECT_PS1: string;
export function excelAvailable(): boolean;
export function inspectFiles(
  files: string[],
  opts?: { recalc?: boolean; maxCells?: number; timeoutMs?: number; noRetry?: boolean; maxAttempts?: number; noCrossCheck?: boolean; probeCells?: string[]; onLine?: (o: Record<string, unknown>) => void },
): Promise<InspectResult[]>;
export function inspectFile(
  file: string,
  opts?: { recalc?: boolean; maxCells?: number; timeoutMs?: number; probeCells?: string[]; onLine?: (o: Record<string, unknown>) => void },
): Promise<InspectResult>;
