/**
 * Test helpers for the keys module: SourceBooks built with SheetJS (the app's reader may not be ready) or from
 * plain arrays, and KeySources with auto-detected sheet / region.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as XLSX from 'xlsx';
import type { Cell, SheetGrid, SourceBook, XStyle, RangeBox } from '../../src/core/types';
import type { KeySource } from '../../src/core/contracts';
import { detectRegion, pickDataSheet } from '../../src/core/keys';
import { normEmail, normName } from '../../src/core/keys/normalize';

export const FIXTURES = fileURLToPath(new URL('../fixtures/keys', import.meta.url));

const EMPTY_FEATURES = () => ({ merges: [] as RangeBox[], rules: [], tables: [], arrays: [], hyperlinks: [] });
const NO_COUNTS = { charts: 0, pivotTables: 0, vba: false, drawings: 0, slicers: 0, sparklineGroups: 0, externalLinks: 0, threadedComments: 0, formControls: 0, tables: 0, images: 0 };

class StyleTable {
  styles: XStyle[] = [{}];
  private byFmt = new Map<string, number>();
  index(fmt: string | undefined): number | undefined {
    if (!fmt || fmt === 'General') return undefined;
    let i = this.byFmt.get(fmt);
    if (i === undefined) {
      i = this.styles.length;
      this.styles.push({ numFmt: fmt });
      this.byFmt.set(fmt, i);
    }
    return i;
  }
}

/** Excel serial (1900 system) of a UTC date */
export function serialOf(d: Date): number {
  return d.getTime() / 86400000 + 25569;
}

/** SourceBook from xlsx/csv bytes via SheetJS (values, number formats, formulas, merges, sheet visibility). */
export function bookFromBytes(bytes: Uint8Array, fileName: string, id = fileName): SourceBook {
  let wb: XLSX.WorkBook;
  if (/\.csv$/i.test(fileName)) {
    let text = new TextDecoder('utf-8').decode(bytes);
    text = text.replace(/^﻿/, '');
    wb = XLSX.read(text, { type: 'string', raw: true });
  } else {
    wb = XLSX.read(bytes, { type: 'array', cellNF: true, cellFormula: true, cellDates: false });
  }
  const st = new StyleTable();
  const sheets: SheetGrid[] = wb.SheetNames.map((name, si) => {
    const ws = wb.Sheets[name];
    const hidden = wb.Workbook?.Sheets?.[si]?.Hidden ?? 0;
    const grid: SheetGrid = {
      name, state: hidden === 1 ? 'hidden' : hidden === 2 ? 'veryHidden' : 'visible',
      maxRow: 0, maxCol: 0, rows: [], colProps: [], rowProps: [], features: EMPTY_FEATURES(),
    };
    if (!ws['!ref']) return grid;
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const x = ws[XLSX.utils.encode_cell({ r, c })] as XLSX.CellObject | undefined;
        if (!x) continue;
        let cell: Cell | undefined;
        const fmt = typeof x.z === 'string' ? x.z : undefined;
        switch (x.t) {
          case 'n': {
            const isDate = !!fmt && fmt !== 'General' && XLSX.SSF.is_date(fmt);
            cell = { t: isDate ? 'd' : 'n', v: x.v as number };
            break;
          }
          case 's':
            cell = { t: 's', v: String(x.v ?? '') };
            break;
          case 'b':
            cell = { t: 'b', v: !!x.v };
            break;
          case 'e':
            cell = { t: 'e', v: x.w ?? '#N/A' };
            break;
          case 'd':
            cell = { t: 'd', v: serialOf(x.v as Date) };
            break;
          default:
            cell = x.f ? { t: 'z', v: null } : undefined;
        }
        if (!cell) continue;
        if (x.f) cell.f = x.f;
        const s = st.index(fmt);
        if (s !== undefined) cell.s = s;
        (grid.rows[r] ??= [])[c] = cell;
        grid.maxRow = Math.max(grid.maxRow, r + 1);
        grid.maxCol = Math.max(grid.maxCol, c + 1);
      }
    }
    for (const m of ws['!merges'] ?? []) grid.features.merges.push({ r1: m.s.r + 1, c1: m.s.c + 1, r2: m.e.r + 1, c2: m.e.c + 1 });
    return grid;
  });
  return {
    id, fileName, size: bytes.length, format: /\.csv$/i.test(fileName) ? 'csv' : 'xlsx', bytes, fidelity: 'full',
    styles: st.styles, sheets, names: [], features: NO_COUNTS, date1904: false, warnings: [],
  };
}

export function bookFromFile(file: string, id?: string): SourceBook {
  const bytes = new Uint8Array(fs.readFileSync(file));
  return bookFromBytes(bytes, path.basename(file), id ?? path.basename(file));
}

/** A cell spec for bookFromRows: plain value, or { v, fmt, f, t } */
export type CellSpec = string | number | boolean | Date | null | undefined | { v: string | number | boolean | Date | null; fmt?: string; f?: string; t?: Cell['t'] };

export interface SheetSpec {
  name: string;
  rows: CellSpec[][];
  merges?: RangeBox[];
  state?: SheetGrid['state'];
}

/** SourceBook from arrays (row 1 = rows[0]). Date values become 'd' cells with format yyyy/m/d. */
export function bookFromRows(fileName: string, sheets: SheetSpec[], id = fileName): SourceBook {
  const st = new StyleTable();
  const grids = sheets.map((sp): SheetGrid => {
    const grid: SheetGrid = { name: sp.name, state: sp.state ?? 'visible', maxRow: 0, maxCol: 0, rows: [], colProps: [], rowProps: [], features: EMPTY_FEATURES() };
    sp.rows.forEach((row, r) => {
      row.forEach((spec, c) => {
        if (spec === null || spec === undefined) return;
        const o = typeof spec === 'object' && !(spec instanceof Date) ? spec : { v: spec };
        let cell: Cell;
        const v = o.v;
        let fmt = o.fmt;
        if (v instanceof Date) {
          cell = { t: 'd', v: serialOf(v) };
          fmt ??= 'yyyy/m/d';
        } else if (v === null) cell = { t: o.t ?? 'z', v: null };
        else if (typeof v === 'number') cell = { t: o.t ?? (fmt && XLSX.SSF.is_date(fmt) ? 'd' : 'n'), v };
        else if (typeof v === 'boolean') cell = { t: 'b', v };
        else cell = { t: o.t ?? (/^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!)$/.test(v) ? 'e' : 's'), v };
        if (o.f) cell.f = o.f;
        const s = st.index(fmt);
        if (s !== undefined) cell.s = s;
        (grid.rows[r] ??= [])[c] = cell;
        grid.maxRow = Math.max(grid.maxRow, r + 1);
        grid.maxCol = Math.max(grid.maxCol, c + 1);
      });
    });
    if (sp.merges) grid.features.merges.push(...sp.merges);
    return grid;
  });
  return {
    id, fileName, size: 0, format: 'xlsx', bytes: new Uint8Array(0), fidelity: 'full', styles: st.styles, sheets: grids,
    names: [], features: NO_COUNTS, date1904: false, warnings: [],
  };
}

/** KeySource with the auto-detected data sheet and region (what the UI does by default). */
export function sourceOf(book: SourceBook, id = book.id, label = book.fileName.replace(/\.[^.]+$/, ''), sheet?: string): KeySource {
  const s = sheet ?? pickDataSheet(book);
  return { id, book, region: detectRegion(book, s), label };
}

export function sourceFromRows(id: string, rows: CellSpec[][], sheetName = 'Sheet1'): KeySource {
  return sourceOf(bookFromRows(`${id}.xlsx`, [{ name: sheetName, rows }], id), id, id);
}

/**
 * Runs fn and returns its result with the time it took in ms: min(wall clock, CPU time of this process). The work is CPU-bound,
 * so on a machine shared with other builds / test runs the wall clock only measures the contention — the CPU time still says
 * whether the work fits its budget.
 */
export function timeCpu<T>(fn: () => T): [T, number] {
  const c0 = process.cpuUsage();
  const t0 = performance.now();
  const r = fn();
  const wall = performance.now() - t0;
  const c = process.cpuUsage(c0);
  return [r, Math.min(wall, (c.user + c.system) / 1000)];
}

/** ms the calibration workload below takes on an idle development machine (measured; see loadFactor) */
const CALIBRATION_IDLE_MS = 430;

function calibrationWork(): number {
  // the kind of work the key detector does: normalise names, count them in Maps, concatenate composite keys
  const m = new Map<string, number>();
  let n = 0;
  for (let i = 0; i < 120000; i++) {
    const k = normName(`山田　太郎${i % 4001}`).folded + '\u0001' + normEmail(`User${i % 977}@Example.com `);
    m.set(k, (m.get(k) ?? 0) + 1);
    n += k.length;
  }
  return n + m.size;
}

let cachedFactor: number | null = null;

/**
 * How much slower than an idle machine this process runs right now (1 = idle, capped at 4). The performance budgets are
 * stated for an idle machine; the whole suite runs ~100 test files in parallel, which slows every one of them down, so the
 * budgets are multiplied by this calibrated factor (best of 3 runs of a fixed workload, measured once per test file).
 */
export function loadFactor(): number {
  if (cachedFactor !== null) return cachedFactor;
  let best = Infinity;
  for (let k = 0; k < 3; k++) {
    const t0 = performance.now();
    calibrationWork();
    best = Math.min(best, performance.now() - t0);
  }
  cachedFactor = Math.max(1, Math.min(4, best / CALIBRATION_IDLE_MS));
  return cachedFactor;
}

/** a time budget (ms) stated for an idle machine, scaled to the current load */
export const budgetMs = (idleMs: number): number => Math.round(idleMs * loadFactor());
