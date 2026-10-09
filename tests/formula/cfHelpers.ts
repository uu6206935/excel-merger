// Shared helpers of cfSpread.test.ts / cfProbe.test.ts: per-cell coverage of CF / DV rules (Excel fragments rules
// arbitrarily, so rules are compared by the effective rule of every cell, never by their sqref).
import { adjustAnchoredRule, parseAreaText, relKey, sqrefTopLeft, translateFormula, type Area, type StructOp } from '../../src/core/formula/excelRef';

export const areasOf = (sq: string): Area[] => sq.split(/[\s,]+/).filter(Boolean).map((s) => {
  const a = parseAreaText(s);
  if (!a) throw new Error(`bad sqref part ${s}`);
  return a;
});

export interface CoverRule { kind: string; type: string; sqref: Area[]; formulas: string[] }

export const colName = (c: number): string => { let s = ''; for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };

/**
 * cell "col,row" -> the rules covering it, as a sorted "kind:type|R1C1 formulas" list joined by " || ". A cell can be
 * covered twice (a moved cell and a cell that joined a block by format inheritance); identical rules count once.
 * Formulas are written for the top-left cell of the sqref bounding box (the xlsx convention).
 */
export function cover(rules: CoverRule[]): Map<string, string> {
  const m = new Map<string, Set<string>>();
  for (const r of rules) {
    const tl = sqrefTopLeft(r.sqref);
    for (const a of r.sqref) {
      if (a.kind === 'cols' || a.kind === 'rows') throw new Error('whole-axis area in a probe');
      for (let c = Math.min(a.c1, a.c2); c <= Math.max(a.c1, a.c2); c++) for (let row = Math.min(a.r1, a.r2); row <= Math.max(a.r1, a.r2); row++) {
        const dr = row - tl.row, dc = c - tl.col;
        const fs = r.formulas.map((f) => relKey(dr === 0 && dc === 0 ? f : translateFormula(f, dr, dc), row, c));
        const k = `${c},${row}`;
        (m.get(k) ?? m.set(k, new Set()).get(k)!).add(`${r.kind}:${r.type}|${fs.join(' | ')}`);
      }
    }
  }
  return new Map([...m].map(([k, v]) => [k, [...v].sort().join(' || ')]));
}

/** The differences of two coverages as readable lines (row-major order). */
export function diffMaps(want: Map<string, string>, got: Map<string, string>, only?: (col: number, row: number) => boolean): string[] {
  const out: string[] = [];
  const keys = [...new Set([...want.keys(), ...got.keys()])].sort((x, y) => {
    const [xc, xr] = x.split(',').map(Number), [yc, yr] = y.split(',').map(Number);
    return xr - yr || xc - yc;
  });
  for (const k of keys) {
    const [c, r] = k.split(',').map(Number);
    if (only && !only(c, r)) continue;
    if (want.get(k) !== got.get(k)) out.push(`${colName(c)}${r}: excel [${want.get(k) ?? ''}] ours [${got.get(k) ?? ''}]`);
  }
  return out;
}

// ---------------------------------------------------------------------------------------------------------------
// probe scenarios (excel-cf-probe.ps1 / extract-cf-probe.mjs): rules + ONE column op + what Excel left behind
// ---------------------------------------------------------------------------------------------------------------
export interface ExRule { kind: 'cf' | 'dv'; type: string; sqref: string; formulas: string[] }
export interface ProbeScenario {
  id: string;
  rules: Array<{ type: string; addr: string; active: string; f: string }>;
  ops: Array<{ kind: 'mov' | 'ins' | 'del'; from?: number; count: number; to?: number; at?: number }>;
  excel: ExRule[];
}
const CF_TYPE: Record<string, string> = { expr: 'expression', cellIs: 'cellIs', colorScale: 'colorScale' };
const toOp = (o: ProbeScenario['ops'][number]): StructOp =>
  o.kind === 'ins' ? { type: 'insert', axis: 'col', at: o.at!, count: o.count }
    : o.kind === 'del' ? { type: 'delete', axis: 'col', at: o.at!, count: o.count }
      : { type: 'move', axis: 'col', from: o.from!, count: o.count, to: o.to! };

/** our model: every rule of the scenario through its ops; the other rules on the same range are its siblings */
export function probeOurs(sc: ProbeScenario): CoverRule[] {
  let cur: CoverRule[] = sc.rules.map((r) => ({
    kind: r.type.startsWith('dv') ? 'dv' : 'cf', type: CF_TYPE[r.type] ?? r.type, sqref: areasOf(r.addr.replace(/,/g, ' ')), formulas: r.f ? [r.f.replace(/^=/, '')] : [],
  }));
  for (const o of sc.ops) {
    const op = toOp(o);
    const addr = (r: CoverRule) => r.sqref.map((a) => `${a.c1}:${a.r1}:${a.c2}:${a.r2}`).join(' ');
    const next: CoverRule[] = [];
    for (const r of cur) {
      const siblings = cur.filter((x) => x !== r && x.kind === r.kind && addr(x) === addr(r)).map((x) => x.formulas);
      for (const x of adjustAnchoredRule({ sqref: r.sqref, formulas: r.formulas }, op, 'P', { kind: r.kind as 'cf' | 'dv', type: r.type, siblings })) {
        next.push({ ...r, sqref: x.sqref, formulas: x.formulas });
      }
    }
    cur = next;
  }
  return cur;
}

/** differences between Excel's cells and ours for a scenario (empty = equal) */
export const probeDiffs = (sc: ProbeScenario): string[] => diffMaps(cover(sc.excel.map((r) => ({ ...r, sqref: areasOf(r.sqref) }))), cover(probeOurs(sc)));
