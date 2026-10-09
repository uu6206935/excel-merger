// Hand-built SourceBook / premerged fixtures for the doc (editor) tests. No reader, no package: the doc only
// depends on the readBook output shape (SourceBook), so the fixtures are built directly in that shape.
import { expect } from 'vitest';
import { colToNum, numToCol } from '../../src/core/formula/excelRef';
import type { Cell, CellType, ColProps, DefinedName, EditorSnapshot, MergedDoc, SheetFeatures, SheetGrid, SourceBook, XStyle } from '../../src/core/types';
import type { OpResult, PremergeLayout } from '../../src/core/contracts';
import { createEditor } from '../../src/core/doc';

export type Scalar = string | number | boolean | null;
export type CellSpec = Scalar | Cell;

/** "C5" -> {row: 5, col: 3} */
export function parseA1(a1: string): { row: number; col: number } {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a1);
  if (!m) throw new Error(`bad cell ${a1}`);
  return { row: Number(m[2]), col: colToNum(m[1]) };
}

export const a1 = (col: number, row: number): string => `${numToCol(col)}${row}`;

/** A formula cell (file form, no '=') with its cached value. */
export function F(f: string, v: Scalar = 0, extra: Partial<Cell> = {}): Cell {
  const t: CellType = v === null ? 'z' : typeof v === 'number' ? 'n' : typeof v === 'boolean' ? 'b' : typeof v === 'string' && v.startsWith('#') ? 'e' : 's';
  return { t, v, f, ...extra };
}

export function toCell(spec: CellSpec): Cell | undefined {
  if (spec === undefined) return undefined;
  if (spec !== null && typeof spec === 'object') return spec;
  if (spec === null) return { t: 'z', v: null };
  if (typeof spec === 'number') return { t: 'n', v: spec };
  if (typeof spec === 'boolean') return { t: 'b', v: spec };
  return { t: 's', v: spec };
}

export interface SheetSpec {
  name: string;
  cells?: Record<string, CellSpec>;
  /** by column letter or number */
  colProps?: Record<string, ColProps>;
  features?: Partial<SheetFeatures>;
  maxCol?: number;
  maxRow?: number;
}

export interface BookSpec {
  sheets: SheetSpec[];
  names?: DefinedName[];
  styles?: XStyle[];
}

export function makeSheet(s: SheetSpec): SheetGrid {
  const rows: SheetGrid['rows'] = [];
  let maxRow = s.maxRow ?? 0;
  let maxCol = s.maxCol ?? 0;
  for (const [addr, spec] of Object.entries(s.cells ?? {})) {
    const { row, col } = parseA1(addr);
    const cell = toCell(spec);
    if (!cell) continue;
    (rows[row - 1] ??= [])[col - 1] = cell;
    maxRow = Math.max(maxRow, row);
    maxCol = Math.max(maxCol, col);
  }
  const colProps: SheetGrid['colProps'] = [];
  for (const [k, p] of Object.entries(s.colProps ?? {})) {
    const c = /^\d+$/.test(k) ? Number(k) : colToNum(k);
    colProps[c - 1] = p;
  }
  return {
    name: s.name,
    state: 'visible',
    maxRow,
    maxCol,
    rows,
    colProps,
    rowProps: [],
    features: { merges: [], rules: [], tables: [], arrays: [], hyperlinks: [], ...(s.features ?? {}) },
  };
}

export function makeBook(spec: BookSpec): SourceBook {
  return {
    id: 'premerged',
    fileName: 'premerged.xlsx',
    size: 0,
    format: 'xlsx',
    bytes: new Uint8Array(),
    fidelity: 'full',
    styles: spec.styles ?? [{}],
    sheets: spec.sheets.map(makeSheet),
    names: spec.names ?? [],
    features: { charts: 0, pivotTables: 0, vba: false, drawings: 0, slicers: 0, sparklineGroups: 0, externalLinks: 0, threadedComments: 0, formControls: 0, tables: 0, images: 0 },
    date1904: false,
    warnings: [],
  };
}

export interface EditorSpec extends BookSpec {
  /** merged sheet (default: first sheet) */
  sheet?: string;
  imported?: PremergeLayout['imported'];
  appended?: PremergeLayout['appended'];
  region?: { headerRow: number; firstDataRow: number; lastDataRow: number };
}

export function makeEditor(spec: EditorSpec): EditorSnapshot {
  const book = makeBook(spec);
  const sheet = spec.sheet ?? spec.sheets[0].name;
  const imported = spec.imported ?? [];
  const layout: PremergeLayout = {
    sheet,
    insertAt: imported.length ? Math.min(...imported.map((i) => i.col)) : 0,
    imported,
    appendAt: spec.appended?.length ? Math.min(...spec.appended.map((a) => a.row)) : null,
    appended: spec.appended ?? [],
  };
  const g = book.sheets.find((s) => s.name === sheet)!;
  const region = spec.region ?? { headerRow: 1, firstDataRow: 2, lastDataRow: Math.max(2, g.maxRow) };
  return createEditor(book, layout, { s2: { label: '成績', color: '#e8a33d' } }, region);
}

// ---------------------------------------------------------------------------------------------
// accessors (CURRENT positions)
// ---------------------------------------------------------------------------------------------
export function at(doc: MergedDoc, addr: string): Cell | undefined {
  const { row, col } = parseA1(addr);
  return doc.columns[col - 1]?.cells[row - 1];
}

export const fAt = (doc: MergedDoc, addr: string): string | undefined => at(doc, addr)?.f;
export const vAt = (doc: MergedDoc, addr: string) => at(doc, addr)?.v;

export function ext(doc: MergedDoc, sheet: string, addr: string): string | undefined {
  const { row, col } = parseA1(addr);
  return doc.externalFormulas.find((e) => e.sheet === sheet && e.row === row && e.col === col)?.f;
}

export const ids = (doc: MergedDoc): string[] => doc.columns.map((c) => c.id);

export function rowValues(doc: MergedDoc, row: number): Scalar[] {
  return doc.columns.map((c) => c.cells[row - 1]?.v ?? null);
}

export function ok(r: OpResult): EditorSnapshot {
  if (!r.ok) throw new Error(`op refused: ${r.reason}`);
  return r.snapshot;
}

export function refused(r: OpResult): string {
  expect(r.ok).toBe(false);
  return r.ok ? '' : r.reason;
}

/** Every formula text of the doc (merged sheet, other sheets, names, rules). */
export function allFormulas(doc: MergedDoc): string[] {
  const out: string[] = [];
  for (const c of doc.columns) for (const cell of c.cells) if (cell?.f !== undefined) out.push(cell.f);
  for (const e of doc.externalFormulas) out.push(e.f);
  for (const n of doc.names) out.push(n.refersTo);
  for (const r of doc.features.rules) out.push(...r.formulas);
  for (const er of doc.externalRules) out.push(...er.rule.formulas);
  return out;
}

/** Row-1 marker per current column (tests put 'hA', 'hB' ... in row 1). */
export const markers = (doc: MergedDoc): Scalar[] => rowValues(doc, 1);

// ---------------------------------------------------------------------------------------------
// a small realistic merged book
// ---------------------------------------------------------------------------------------------
/** xf indices of ROSTER_STYLES */
export const XF = { general: 0, text: 1, int: 2, date: 3, pct: 4, header: 5 } as const;
export const ROSTER_STYLES: XStyle[] = [{}, { numFmt: '@' }, { numFmt: '0' }, { numFmt: 'yyyy/m/d' }, { numFmt: '0%' }, { bold: true, fill: '#DDEBF7' }];

/**
 * 名簿 (merged sheet, header row 1, data rows 2..6, total in D8):
 *   A 学生ID (text '00123'…) | B 氏名 | C 大学 | D 点数 | E 倍 (=D*2) | F 学部 [imported s2:3] | G メール [imported s2:4] | H 備考
 * 集計: formulas on 名簿. Names 点数範囲 / 大学セル.
 */
export function rosterSpec(): EditorSpec {
  const cells: Record<string, CellSpec> = {
    A1: { t: 's', v: '学生ID', s: XF.header }, B1: { t: 's', v: '氏名', s: XF.header }, C1: { t: 's', v: '大学', s: XF.header },
    D1: { t: 's', v: '点数', s: XF.header }, E1: { t: 's', v: '倍', s: XF.header }, F1: { t: 's', v: '学部', s: XF.header },
    G1: { t: 's', v: 'メール', s: XF.header }, H1: { t: 's', v: '備考', s: XF.header },
    D8: F('SUM(D2:D6)', 350, { s: XF.int }),
  };
  const names = ['佐藤', '鈴木', '高橋', '田中', '伊藤'];
  const univ = ['東大', '京大', '阪大', '東北大', '九大'];
  const pts = [80, 70, 90, 60, 50];
  for (let i = 0; i < 5; i++) {
    const r = i + 2;
    cells[`A${r}`] = { t: 's', v: `0012${3 + i}`, s: XF.text };
    cells[`B${r}`] = names[i];
    cells[`C${r}`] = univ[i];
    cells[`D${r}`] = { t: 'n', v: pts[i], s: XF.int };
    cells[`E${r}`] = F(`D${r}*2`, pts[i] * 2, { s: XF.int });
    cells[`F${r}`] = ['法', '文', '工', '理', '医'][i];
    cells[`G${r}`] = `u${i}@x.jp`;
  }
  return {
    styles: ROSTER_STYLES,
    sheets: [
      {
        name: '名簿',
        cells,
        // widths the user set (customWidth="1"): typed dates never widen them; tests that need an auto width replace a column's props
        colProps: {
          A: { width: 10, customWidth: true }, B: { width: 12, customWidth: true }, C: { width: 14, customWidth: true }, D: { width: 8, customWidth: true },
          E: { width: 9, customWidth: true }, F: { width: 11, customWidth: true }, G: { width: 20, customWidth: true }, H: { width: 15, customWidth: true },
        },
        maxCol: 8,
      },
      {
        name: '集計',
        cells: {
          B2: F('SUM(名簿!D2:D6)', 350),
          B3: F('名簿!C2', '東大'),
          B4: F('VLOOKUP("00123",名簿!A2:E6,4,FALSE)', 80),
          B5: F('COUNTA(名簿!$B:$B)', 6),
          B6: F('SUM(点数範囲)', 350),
        },
      },
    ],
    names: [
      { name: '点数範囲', scope: null, refersTo: '名簿!$D$2:$D$6' },
      { name: '大学セル', scope: null, refersTo: '名簿!$C$2' },
    ],
    imported: [
      { col: 6, sourceId: 's2', sourceCol: 3, header: '学部' },
      { col: 7, sourceId: 's2', sourceCol: 4, header: 'メール' },
    ],
    region: { headerRow: 1, firstDataRow: 2, lastDataRow: 6 },
  };
}

export const roster = (patch?: (s: EditorSpec) => void): EditorSnapshot => {
  const spec = rosterSpec();
  patch?.(spec);
  return makeEditor(spec);
};

/** current column id at a letter */
export const idAt = (doc: MergedDoc, letter: string): string => doc.columns[colToNum(letter) - 1].id;
