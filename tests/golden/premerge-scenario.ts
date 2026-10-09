/**
 * Premerge scenario on Excel-authored fixtures (tests/golden/com/make-fixtures.ps1):
 *   base  pm-base-plain.xlsx / pm-base-table.xlsx : sheet 学生, header row 1 (ID 氏名 大学 点数), data rows 2..6
 *         (S001..S005), 合計 row 8 (B8 =COUNTA(B2:B6), D8 =SUM(D2:D6)), AutoFilter A1:D6 (plain) or ListObject
 *         tblStudents A1:D6 (table), CF D2:D6, name Scores, sheet Report with formulas + a chart on 学生.
 *   source pm-src.xlsx : sheet 回答, header row 1 (学生ID 名前 志望 評価 備考 点数 Calc Total 氏名), rows 2..6 =
 *         S003 S001 S009 S002 S005; Calc = F*2+D (same row), Total = SUM($F$2:$F$6), styled columns.
 * Join: S001..S005 base rows, S004 unmatched, S009 appended (join 'full').
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import type { KeySource, PremergeInput } from '../../src/core/contracts';
import type { SourceBook, TableRegion } from '../../src/core/types';

const here = path.dirname(fileURLToPath(import.meta.url));
export const GOLDEN_FIX = path.resolve(here, '../fixtures/golden');

export function fakeBook(id: string, fileName: string, bytes: Uint8Array, fidelity: 'full' | 'values' = 'full'): SourceBook {
  return {
    id, fileName, size: bytes.length, format: 'xlsx', bytes, fidelity, styles: [], sheets: [], names: [], date1904: false, warnings: [],
    features: { charts: 0, pivotTables: 0, vba: false, drawings: 0, slicers: 0, sparklineGroups: 0, externalLinks: 0, threadedComments: 0, formControls: 0, tables: 0, images: 0 },
  };
}

export function source(id: string, file: string, region: TableRegion, label: string, fidelity: 'full' | 'values' = 'full', bytes?: Uint8Array): KeySource {
  const b = bytes ?? new Uint8Array(fs.readFileSync(path.join(GOLDEN_FIX, file)));
  return { id, label, region, book: fakeBook(id, file, b, fidelity) };
}

/**
 * Base region narrower than the ListObject that contains it (g-table.xlsx: table A3:G13, region A..E): the
 * imported columns land INSIDE the table, so they become table columns named after their headers.
 */
export function insideTableScenario(): PremergeInput {
  const base = source('base', 'g-table.xlsx', { sheet: '名簿', headerRow: 3, firstDataRow: 4, lastDataRow: 13, firstCol: 1, lastCol: 5 }, 'g-table');
  const src = source('src', 'pm-src.xlsx', { sheet: '回答', headerRow: 1, firstDataRow: 2, lastDataRow: 6, firstCol: 1, lastCol: 9 }, 'pm-src');
  return {
    base,
    sources: [src],
    rows: Array.from({ length: 10 }, (_x, i) => ({ baseRow: 4 + i, matches: { src: i < 5 ? { row: 2 + i, via: 'ID', flags: [] } : null } })),
    imports: [{ sourceId: 'src', col: 3 }, { sourceId: 'src', col: 7 }],
    keyFill: {},
    options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix' },
  };
}

/**
 * Rich text source (r2-rich.xlsx, com/make-fixtures-r2.ps1): NAME B2 = "Bold normal" (bold red first run), B3 = 山田
 * with the reading ヤマダ. Imported into pm-base-plain (NAME -> E, MEMO -> F); S001 / S002 matched, one appended row
 * whose key cells come from the source (B7 = the rich NAME).
 */
export function richScenario(): PremergeInput {
  return {
    base: source('base', 'pm-base-plain.xlsx', { sheet: '学生', headerRow: 1, firstDataRow: 2, lastDataRow: 6, firstCol: 1, lastCol: 4 }, 'pm-base-plain'),
    sources: [source('r', 'r2-rich.xlsx', { sheet: 'S', headerRow: 1, firstDataRow: 2, lastDataRow: 3, firstCol: 1, lastCol: 3 }, 'r2-rich')],
    rows: [
      { baseRow: 2, matches: { r: { row: 2, via: 'ID', flags: [] } } },
      { baseRow: 3, matches: { r: { row: 3, via: 'ID', flags: [] } } },
      { baseRow: 4, matches: { r: null } }, { baseRow: 5, matches: { r: null } }, { baseRow: 6, matches: { r: null } },
      { baseRow: null, appendedFrom: { sourceId: 'r', row: 2 }, matches: { r: { row: 2, via: 'ID', flags: [] } } },
    ],
    imports: [{ sourceId: 'r', col: 2 }, { sourceId: 'r', col: 3 }],
    keyFill: { r: [{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 2 }] },
    options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix' },
  };
}

/** Appended rows whose key cells combine / split source columns (pm-src.xlsx 名前 + 氏名 ...). */
export function keyFillScenario(): PremergeInput {
  const s = scenario('plain');
  s.keyFill = { src: [{ baseCol: 1, sourceCols: [1, 2], joiner: '-' }, { baseCol: 2, sourceCols: [9, 2] }, { baseCol: 3, sourceCol: 9, splitPart: 0 }] };
  return s;
}

export function scenario(variant: 'plain' | 'table', opts: Partial<PremergeInput['options']> = {}): PremergeInput {
  const base = source('base', `pm-base-${variant}.xlsx`, { sheet: '学生', headerRow: 1, firstDataRow: 2, lastDataRow: 6, firstCol: 1, lastCol: 4 }, `pm-base-${variant}`);
  const src = source('src', 'pm-src.xlsx', { sheet: '回答', headerRow: 1, firstDataRow: 2, lastDataRow: 6, firstCol: 1, lastCol: 9 }, 'pm-src');
  const m = (row: number) => ({ src: { row, via: 'ID', flags: [] } });
  return {
    base,
    sources: [src],
    rows: [
      { baseRow: 2, matches: m(3) }, // S001
      { baseRow: 3, matches: m(5) }, // S002
      { baseRow: 4, matches: m(2) }, // S003
      { baseRow: 5, matches: { src: null } }, // S004 unmatched
      { baseRow: 6, matches: m(6) }, // S005
      { baseRow: null, appendedFrom: { sourceId: 'src', row: 4 }, matches: { src: { row: 4, via: 'ID', flags: [] } } }, // S009
    ],
    imports: [3, 4, 5, 6, 7, 8, 9].map((col) => ({ sourceId: 'src', col })),
    keyFill: { src: [{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 9 }] },
    options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix', ...opts },
  };
}
