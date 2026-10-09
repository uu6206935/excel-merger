/**
 * Helpers of the formula-import tests (premerge round 8): build small source / base workbooks from a grid, run buildPremerged with a
 * hand-made join plan and read the merged cells back.
 */
import type { KeySource, PremergeInput } from '../../src/core/contracts';
import type { Cell, SheetGrid, SourceBook, TableRegion } from '../../src/core/types';
import { readBook } from '../../src/core/xlsx/read';
import { buildPremerged, type FormulaTrace } from '../../src/core/xlsx/premerge';
import { buildXlsx, type SheetSpec } from '../golden/build';
import { fakeBook } from '../golden/premerge-scenario';

/** a cell of a hand-built grid: number / text / boolean / null, or a formula with its stored result */
export type G = number | string | boolean | null | F;
export interface F {
  f: string;
  /** stored result (omit = no result stored) */
  v?: number | string | boolean;
  /** error result like '#N/A' */
  e?: string;
  /** shared formula: master has `ref`, dependents only `si` (no text) */
  shared?: { si: number; ref?: string };
  /** array formula: t="array" ref=... ; cm = dynamic array */
  array?: { ref: string; cm?: boolean };
}

export const col = (n: number): string => {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
};
const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

export function cellXml(r: number, c: number, g: G): string {
  const ref = `${col(c)}${r}`;
  if (g === null || g === undefined) return '';
  if (typeof g === 'number') return `<c r="${ref}"><v>${g}</v></c>`;
  if (typeof g === 'boolean') return `<c r="${ref}" t="b"><v>${g ? 1 : 0}</v></c>`;
  if (typeof g === 'string') return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(g)}</t></is></c>`;
  let t = '';
  let v = '';
  if (g.e !== undefined) { t = ' t="e"'; v = `<v>${g.e}</v>`; }
  // Excel writes an empty text result as <v/>
  else if (typeof g.v === 'string') { t = ' t="str"'; v = g.v === '' ? '<v/>' : `<v>${esc(g.v)}</v>`; }
  else if (typeof g.v === 'boolean') { t = ' t="b"'; v = `<v>${g.v ? 1 : 0}</v>`; }
  else if (typeof g.v === 'number') v = `<v>${g.v}</v>`;
  let fa = '';
  if (g.shared) fa = ` t="shared" si="${g.shared.si}"${g.shared.ref ? ` ref="${g.shared.ref}"` : ''}`;
  if (g.array) fa = ` t="array" ref="${g.array.ref}"`;
  const cm = g.array?.cm ? ' cm="1"' : '';
  const fx = g.f === '' ? `<f${fa}/>` : `<f${fa}>${esc(g.f)}</f>`;
  return `<c r="${ref}"${t}${cm}>${fx}${v}</c>`;
}

export function sheetData(rows: G[][]): string {
  return rows.map((row, i) => `<row r="${i + 1}">${row.map((g, j) => cellXml(i + 1, j + 1, g)).join('')}</row>`).join('');
}

export interface TableSpec {
  name: string;
  ref: string;
  columns: string[];
  totals?: boolean;
}

export interface BookSpec {
  sheet: string;
  rows: G[][];
  names?: Array<{ name: string; refersTo: string; localSheetId?: number }>;
  tables?: TableSpec[];
  /** more sheets (name + grid) after the data sheet */
  extra?: Array<{ name: string; rows: G[][] }>;
  /** workbook calcPr */
  calcPr?: string;
}

export async function buildBook(b: BookSpec): Promise<Uint8Array> {
  const main: SheetSpec = { name: b.sheet, data: sheetData(b.rows) };
  if (b.tables?.length) {
    main.rels = b.tables.map((t, i) => ({
      type: 'table', id: `rId${i + 1}`, path: `xl/tables/table${i + 1}.xml`, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
      xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${i + 1}" name="${t.name}" displayName="${t.name}" ref="${t.ref}"${t.totals ? ' totalsRowCount="1"' : ''}><tableColumns count="${t.columns.length}">${t.columns.map((c, k) => `<tableColumn id="${k + 1}" name="${esc(c)}"/>`).join('')}</tableColumns></table>`,
    }));
    main.tail = `<tableParts count="${b.tables.length}">${b.tables.map((_t, i) => `<tablePart r:id="rId${i + 1}"/>`).join('')}</tableParts>`;
  }
  const sheets: SheetSpec[] = [main, ...(b.extra ?? []).map((e) => ({ name: e.name, data: sheetData(e.rows) }))];
  const names = b.names?.map((n) => `<definedName name="${esc(n.name)}"${n.localSheetId != null ? ` localSheetId="${n.localSheetId}"` : ''}>${esc(n.refersTo)}</definedName>`).join('');
  let bytes = await buildXlsx(sheets, { names });
  if (b.calcPr) {
    const JSZip = (await import('jszip')).default;
    const z = await JSZip.loadAsync(bytes);
    z.file('xl/workbook.xml', (await z.file('xl/workbook.xml')!.async('string')).replace('<calcPr calcId="191029"/>', b.calcPr));
    bytes = await z.generateAsync({ type: 'uint8array' });
  }
  return bytes;
}

/** a KeySource over a hand-built book: the region is given (not detected); `real` = a SourceBook read by the app's reader, else a bare one */
export async function keySource(id: string, label: string, bytes: Uint8Array, region: TableRegion, real = true): Promise<KeySource> {
  const book: SourceBook = real ? await readBook(bytes, `${label}.xlsx`, { id }) : fakeBook(id, `${label}.xlsx`, bytes);
  return { id, label, book, region };
}

export interface RowPlan {
  /** base sheet row, null = appended */
  base: number | null;
  /** source row matched (or the row the appended row comes from) */
  src: number | null;
}

export interface MergeSpec {
  base: BookSpec;
  src: BookSpec;
  /** source columns to import (1-based) */
  imports: number[];
  rows: RowPlan[];
  keyPairs?: Array<{ baseCol: number; sourceCol: number }>;
  keyFill?: Array<{ baseCol: number; sourceCol: number }>;
  options?: Partial<PremergeInput['options']>;
  /** bare SourceBooks (no sheets / names / tables): the package decides */
  bare?: boolean;
  /** region overrides */
  srcRegion?: Partial<TableRegion>;
  baseRegion?: Partial<TableRegion>;
  /** change the assembled PremergeInput before the run (hostile inputs) */
  tweak?: (input: PremergeInput) => void;
}

export async function merge(m: MergeSpec) {
  const baseRows = m.base.rows.length;
  const srcRows = m.src.rows.length;
  const bytesB = await buildBook(m.base);
  const bytesS = await buildBook(m.src);
  const colsB = Math.max(...m.base.rows.map((r) => r.length));
  const colsS = Math.max(...m.src.rows.map((r) => r.length));
  const baseKs = await keySource('b', 'base', bytesB, { sheet: m.base.sheet, headerRow: 1, firstDataRow: 2, lastDataRow: baseRows, firstCol: 1, lastCol: colsB, ...m.baseRegion }, !m.bare);
  const srcKs = await keySource('s', 'src', bytesS, { sheet: m.src.sheet, headerRow: 1, firstDataRow: 2, lastDataRow: srcRows, firstCol: 1, lastCol: colsS, ...m.srcRegion }, !m.bare);
  const plan = m.rows.map((r) => (r.base == null
    ? { baseRow: null, appendedFrom: { sourceId: 's', row: r.src! }, matches: { s: { row: r.src!, via: 'ID', flags: [] as never[] } } }
    : { baseRow: r.base, matches: { s: r.src == null ? null : { row: r.src, via: 'ID', flags: [] as never[] } } }));
  const input: PremergeInput = {
    base: baseKs,
    sources: [srcKs],
    rows: plan,
    imports: m.imports.map((c) => ({ sourceId: 's', col: c })),
    keyFill: { s: m.keyFill ?? [] },
    ...(m.keyPairs ? { keyPairs: { s: m.keyPairs } } : {}),
    options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix', ...m.options },
  };
  m.tweak?.(input);
  const trace: FormulaTrace[] = [];
  const res = await buildPremerged(input, undefined, (t) => trace.push(t));
  const book = await readBook(res.bytes, 'merged.xlsx', { id: 'm' });
  const sheet: SheetGrid = book.sheets.find((s) => s.name === res.layout.sheet)!;
  const cell = (row: number, c: number): Cell | undefined => sheet.rows[row - 1]?.[c - 1];
  const find = (row: number, c: number) => trace.find((t) => t.row === row && t.col === c);
  return { res, trace, book, sheet, cell, find, input, baseKs, srcKs, report: res.report, items: res.report.items };
}

/** the standard persons: base order P1..P5 on rows 2..6 */
export const PEOPLE = ['P1', 'P2', 'P3', 'P4', 'P5'];
export const baseBook = (sheet = '名簿'): BookSpec => ({ sheet, rows: [['ID', '氏名'], ...PEOPLE.map((p, i) => [p, `名前${i + 1}`])] });
