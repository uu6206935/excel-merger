// Replays a sample of the R2b Excel ground truth (Excel 365 via COM, research/r2b) THROUGH THE DOC: the op sheet
// becomes the merged sheet of an editor, every formula of the fixture (hosted on the op sheet and on the other
// sheets), names, table, merges, array, autofilter, frozen panes, widths and row-1 markers are loaded, the op is
// applied with moveColumns / insertColumns / deleteColumns, and the doc must equal Excel's result.
// Cases where excelRef itself documents a deviation from Excel (tests/formula/DEVIATIONS.md: sheet-edge refs during
// moves, legacy implicit intersection, table quirks) are listed in `dev`; the doc must then equal excelRef.
// Fixture: tests/doc/fixtures/r2b-doc-sample.json (built by build-r2b-sample.mjs).
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { adjustFormula, mapIndex, numToCol, parseAreaText, type StructOp } from '../../src/core/formula/excelRef';
import { deleteColumns, insertColumns, moveColumns } from '../../src/core/doc';
import type { Cell, EditorSnapshot, RangeBox } from '../../src/core/types';
import { ext, makeEditor, markers, parseA1, type CellSpec } from './helpers';

interface SFormula { id: string; where: string; cell: string; f: string; arr?: string }
interface SSuite {
  dataSheet: string;
  sheets: string[];
  formulas: SFormula[];
  names: Array<{ key: string; name: string; scope: string | null; refersTo: string }>;
  tables: Array<{ name: string; ref: string; columns: string[] }>;
  merges: string[];
  arrays: string[];
  autofilter: string | null;
  freeze: number;
  widths: number[] | null;
  markers: string[] | null;
}
interface SOp {
  opId: string;
  suite: string;
  op: StructOp;
  refused: boolean;
  silent: boolean;
  formulas: Array<[string, string | null, string | null]>;
  names: Array<[string, string | null]>;
  merges?: string[];
  widths?: number[];
  markers?: Array<string | null>;
  autofilter?: string | null;
  freeze?: number;
  table?: { ref: string; columns: string[] } | null;
}
interface Sample { suites: Record<string, SSuite>; ops: SOp[]; dev: string[] }

const here = path.dirname(fileURLToPath(import.meta.url));
const SAMPLE: Sample = JSON.parse(fs.readFileSync(path.join(here, 'fixtures/r2b-doc-sample.json'), 'utf8'));
const DEV = new Set(SAMPLE.dev);

/**
 * Ops the doc deliberately refuses although Excel performs them (safer than Excel, documented):
 *  g03  B before D: Excel moves column B INTO the middle of the merged C3:E3 and silently un-merges it.
 *  g09  O before A: part of the legacy array N2:P2 — Excel's Insert Cut Cells silently degraded into a blank insert.
 */
const DOC_REFUSES = new Set(['g03', 'g09']);

const box = (s: string): RangeBox => {
  const a = parseAreaText(s.replace(/\$/g, ''))!;
  return { r1: Math.min(a.r1, a.r2), c1: Math.min(a.c1, a.c2), r2: Math.max(a.r1, a.r2), c2: Math.max(a.c1, a.c2) };
};

function editorFor(S: SSuite): EditorSnapshot {
  const data = S.dataSheet;
  const cellsBy: Record<string, Record<string, CellSpec>> = {};
  for (const sh of S.sheets) cellsBy[sh] = {};
  let maxCol = Math.max(30, S.widths?.length ?? 0);
  // row-1 markers hA, hB ... over all 30 columns (the fixture lists the first 26)
  for (let c = 1; c <= 30; c++) cellsBy[data][`${numToCol(c)}1`] = `h${numToCol(c)}`;
  for (const f of S.formulas) {
    const c: Cell = { t: 'n', v: 0, f: f.f };
    if (f.arr) {
      c.arr = f.arr;
      if (/SEQUENCE|FILTER|SORT|UNIQUE/i.test(f.f)) c.dyn = true;
    }
    cellsBy[f.where][f.cell] = c;
    if (f.where === data) maxCol = Math.max(maxCol, parseA1(f.cell).col);
  }
  const colProps: Record<string, { width: number }> = {};
  (S.widths ?? []).forEach((w, i) => (colProps[String(i + 1)] = { width: w }));
  return makeEditor({
    sheet: data,
    sheets: S.sheets.map((name) =>
      name === data
        ? {
            name,
            cells: cellsBy[name],
            colProps,
            maxCol,
            features: {
              merges: S.merges.map(box),
              arrays: S.arrays.map(box),
              tables: S.tables.map((t) => ({ name: t.name, ref: t.ref, headerRow: true, totalsRow: false, columns: t.columns })),
              ...(S.autofilter ? { autoFilter: S.autofilter } : {}),
              ...(S.freeze ? { freeze: { cols: S.freeze, rows: 1 } } : {}),
            },
          }
        : { name, cells: cellsBy[name] },
    ),
    names: S.names.map((n) => ({ name: n.name, scope: n.scope, refersTo: n.refersTo })),
    region: { headerRow: 1, firstDataRow: 2, lastDataRow: 20 },
  });
}

function apply(s: EditorSnapshot, op: StructOp) {
  const idsOf = (from: number, count: number) => s.doc.columns.slice(from - 1, from - 1 + count).map((c) => c.id);
  if (op.type === 'move') return moveColumns(s, idsOf(op.from, op.count), op.to - 1);
  if (op.type === 'insert') return insertColumns(s, op.at - 1, op.count);
  return deleteColumns(s, idsOf(op.at, op.count), { convertDependents: false });
}

const bySuite = new Map<string, SOp[]>();
for (const o of SAMPLE.ops) bySuite.set(o.suite, [...(bySuite.get(o.suite) ?? []), o]);

describe('R2b ground truth replayed through the doc', () => {
  for (const [suite, ops] of bySuite) {
    const S = SAMPLE.suites[suite];
    const base = editorFor(S);
    const before = new Map(S.formulas.map((f) => [f.id, f]));
    for (const o of ops) {
      it(`${o.opId} ${JSON.stringify(o.op)}`, () => {
        const r = apply(base, o.op);
        const degenerate = o.op.type === 'move' && o.op.to >= o.op.from && o.op.to <= o.op.from + o.op.count;
        if (degenerate) {
          // Excel refuses (target inside the block) or does nothing (target right after it): the doc does nothing
          expect(r.ok && r.snapshot === base).toBe(true);
          return;
        }
        if (o.refused || DOC_REFUSES.has(o.opId)) {
          expect(r.ok, r.ok ? 'doc accepted an op Excel refuses' : '').toBe(false);
          if (!r.ok) expect(r.reason).toMatch(/[ぁ-んァ-ン一-龥]/);
          return;
        }
        if (!r.ok) throw new Error(`doc refused ${o.opId}: ${r.reason}`);
        const doc = r.snapshot.doc;
        expect(r.snapshot.log).toEqual(expect.arrayContaining([{ kind: 'struct', sheet: S.dataSheet, op: o.op }]));
        const fails: string[] = [];
        for (const [id, cellAfter, fileAfter] of o.formulas) {
          const b = before.get(id)!;
          const caseId = `${o.opId}:${id}`;
          let want = fileAfter;
          if (DEV.has(caseId) && want !== null) {
            const hc = b.where === S.dataSheet ? parseA1(b.cell) : null;
            want = adjustFormula(b.f, o.op, { hostSheet: b.where, targetSheet: S.dataSheet, tables: S.tables.map((t) => ({ ...t, sheet: S.dataSheet })), hostCell: hc ?? undefined }).formula;
          }
          let got: string | undefined | null;
          let host = cellAfter;
          if (DEV.has(caseId) && b.where === S.dataSheet) {
            const hc = parseA1(b.cell);
            const nc = mapIndex(hc.col, o.op);
            host = nc === null ? null : `${numToCol(nc)}${hc.row}`;
          }
          if (b.where === S.dataSheet) {
            if (host === null) continue; // host cell deleted
            const { row, col } = parseA1(host);
            got = doc.columns[col - 1]?.cells[row - 1]?.f;
          } else got = ext(doc, b.where, b.cell);
          if (want === null) continue;
          if (got !== want) fails.push(`${caseId} ${b.where}!${b.cell} ${b.f}: excel ${want} doc ${got}`);
        }
        for (const [key, after] of o.names) {
          const n = S.names.find((x) => x.key === key)!;
          const caseId = `${o.opId}:name:${key}`;
          let want = after;
          if (DEV.has(caseId)) want = adjustFormula(n.refersTo, o.op, { hostSheet: null, targetSheet: S.dataSheet, isName: true, tables: S.tables.map((t) => ({ ...t, sheet: S.dataSheet })) }).formula;
          const got = doc.names.find((x) => x.name === n.name && x.scope === n.scope)?.refersTo;
          if (want !== null && got !== want) fails.push(`${caseId}: excel ${want} doc ${got}`);
        }
        if (o.merges) {
          const got = doc.features.merges.map((m) => `${String.fromCharCode(64 + m.c1)}${m.r1}:${String.fromCharCode(64 + m.c2)}${m.r2}`).sort();
          if (got.join() !== [...o.merges].sort().join()) fails.push(`merges excel ${o.merges} doc ${got}`);
        }
        if (o.autofilter !== undefined && (doc.features.autoFilter ?? null) !== o.autofilter) fails.push(`autofilter excel ${o.autofilter} doc ${doc.features.autoFilter}`);
        if (o.freeze !== undefined && (doc.features.freeze?.cols ?? 0) !== o.freeze) fails.push(`freeze excel ${o.freeze} doc ${doc.features.freeze?.cols}`);
        if (o.markers) {
          const all = markers(doc);
          const got = o.markers.map((_, i) => all[i] ?? null);
          if (got.join('|') !== o.markers.map((m) => m ?? '').join('|')) fails.push(`markers excel ${o.markers} doc ${got}`);
        }
        // Excel quirk (R2b table suite): after a move INSIDE a table, COM ColumnWidth reports the widths right of the
        // moved block shifted by one column although cells and row-1 markers did not shift — not a doc concern
        if (o.widths && !(suite === 'table' && o.op.type === 'move')) {
          o.widths.forEach((w, i) => {
            const got = doc.columns[i]?.props.width;
            if (got !== undefined && got !== w) fails.push(`width of column ${i + 1}: excel ${w} doc ${got}`);
            if (got === undefined && w !== 8.38) fails.push(`width of column ${i + 1}: excel ${w} doc (default)`);
          });
        }
        if (o.table !== undefined) {
          const t = doc.features.tables[0];
          const got = t ? { ref: t.ref, columns: t.columns } : null;
          if (JSON.stringify(got) !== JSON.stringify(o.table)) fails.push(`table excel ${JSON.stringify(o.table)} doc ${JSON.stringify(got)}`);
        }
        expect(fails).toEqual([]);
      });
    }
  }

  it('covers thousands of Excel formula results', () => {
    const n = SAMPLE.ops.reduce((k, o) => k + o.formulas.length, 0);
    expect(n).toBeGreaterThan(10000);
    // the documented engine deviations are the only tolerated differences
    expect(DEV.size).toBeLessThan(80);
  });
});
