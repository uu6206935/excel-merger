// Conditional formats / data validations of EVERY rule type under column insert / delete / move, against Excel 365
// (fixtures/formula/cf-spread-excel-results.json; generator: fixtures/formula/com/excel-cf-spread.ps1 +
// extract-cf-spread.mjs). One band of two rows per rule (area C or C:E, plus replicas of R2b cases and formulas that
// point across the op boundary); for every op Excel's saved rules are compared with adjustAnchoredRule PER CELL
// (effective rule type + R1C1-like formula of every cell), because Excel fragments rules arbitrarily.
//
// What this pins down (see AnchoredRuleOptions in excelRef.ts):
//  - data validations inherit / grow over inserted cells and spread onto a moved block exactly like CF;
//  - the CF types colorScale, dataBar, iconSet, top10, uniqueValues, duplicateValues, aboveAverage do NOT spread onto
//    a moved block (cellIs, expression, containsText, ... do).
// Every mismatch must be listed in KNOWN with its reason (and must still differ). CFS_REPORT=1 prints them all.
import { describe, expect, it } from 'vitest';
import {
  adjustAnchoredRule, ruleGrowsOnInsert, ruleSpreadsOnMove, NO_SPREAD_CF_TYPES, type StructOp,
} from '../../src/core/formula/excelRef';
import { areasOf, cover, diffMaps } from './cfHelpers';
import { readFixture } from './helpers';

interface ExRule { kind: 'cf' | 'dv'; type: string; sqref: string; formulas: string[]; x14?: boolean }
interface Fixture { ops: Record<string, StructOp | null>; files: Record<string, ExRule[]> }
const FX = readFixture<Fixture>('cf-spread-excel-results.json');

const before = FX.files.before;
const opNames = Object.keys(FX.ops).filter((o) => FX.ops[o]);
const rowsOf = (sq: string): Set<number> => {
  const rows = new Set<number>();
  for (const a of areasOf(sq)) for (let r = Math.min(a.r1, a.r2); r <= Math.max(a.r1, a.r2); r++) rows.add(r);
  return rows;
};

/** "<op>|<kind>:<type>|<first sqref of the rule before the op>" -> why Excel's result is not reproduced */
const KNOWN: Record<string, string> = {
  // A rule whose formula refers with an ABSOLUTE column to the moved column: Excel takes the cells left of the
  // destination out of the rule and lets it cover the columns between destination and source instead (the
  // CF_INSIDE_MOVE family of Excel quirks, tests/golden/RESULTS.md). We keep "every cell keeps its rule".
  'movGtoD|cf:expression|C209:E210': 'abs-ref-to-moved-column',
  'movGtoD|cf:cellIs|C260:E261': 'abs-ref-to-moved-column',
  // a formula mixing an absolute column before the moved one with a relative reference to it: Excel takes the cell
  // that joined the rule at the destination out again, we keep it (see cfProbe.test.ts p4)
  'movDtoH|cf:expression|C296:G297': 'dest-piece',
};

/** per rule of the "before" workbook: differences between Excel's cells in the rule's rows and ours */
function ruleDiffs(opName: string): Array<{ id: string; diffs: string[] }> {
  const op = FX.ops[opName]!;
  const want = cover(FX.files[opName].map((r) => ({ ...r, sqref: areasOf(r.sqref) })));
  const out: Array<{ id: string; diffs: string[] }> = [];
  for (const r of before) {
    const res = adjustAnchoredRule({ sqref: areasOf(r.sqref), formulas: r.formulas }, op, 'P', { kind: r.kind, type: r.type });
    const got = cover(res.map((x) => ({ kind: r.kind, type: r.type, sqref: x.sqref, formulas: x.formulas })));
    const rows = rowsOf(r.sqref);
    out.push({ id: `${opName}|${r.kind}:${r.type}|${r.sqref.split(' ')[0]}`, diffs: diffMaps(want, got, (_c, row) => rows.has(row)) });
  }
  return out;
}

describe('CF / DV of every type under insert / delete / move vs Excel 365', () => {
  it('uses the whole fixture', () => {
    expect(before.length).toBeGreaterThan(90);
    expect(new Set(before.filter((r) => r.kind === 'cf').map((r) => r.type)).size).toBeGreaterThanOrEqual(16);
    expect(new Set(before.filter((r) => r.kind === 'dv').map((r) => r.type))).toEqual(new Set(['whole', 'list', 'custom']));
    expect(opNames.length).toBeGreaterThanOrEqual(15);
  });

  for (const opName of opNames) {
    it(`op ${opName}: every rule equals Excel per cell except the listed quirks`, () => {
      const bad: string[] = [];
      for (const { id, diffs } of ruleDiffs(opName)) {
        if (!diffs.length) continue;
        if (process.env.CFS_REPORT) console.log(`--- ${id}\n  ${diffs.slice(0, 4).join('\n  ')}`);
        if (!KNOWN[id]) bad.push(`${id}: ${diffs.slice(0, 3).join(' | ')}`);
      }
      expect(bad).toEqual([]);
    });
  }

  it('every listed quirk still differs', () => {
    const all = new Map(opNames.flatMap((o) => ruleDiffs(o)).map((x) => [x.id, x.diffs.length] as const));
    const stale = Object.keys(KNOWN).filter((id) => !all.get(id));
    expect(stale).toEqual([]);
  });

  it('data validations grow over an inserted column and spread onto a moved block like conditional formats', () => {
    // insD: column inserted right after the one-column areas (C): validations of every type grow (dvWhole / dvList /
    // dvCustom are the dv rules of the fixture) - asserted through the comparison above; here the model itself
    const dv = before.find((r) => r.kind === 'dv' && r.type === 'whole')!;
    const op = FX.ops.insD!;
    const grown = adjustAnchoredRule({ sqref: areasOf(dv.sqref), formulas: dv.formulas }, op, 'P', { kind: 'dv', type: 'whole' });
    const cells = cover(grown.map((x) => ({ kind: 'dv', type: 'whole', sqref: x.sqref, formulas: x.formulas })));
    expect(cells.has('4,53')).toBe(true); // D53: the inserted column took the validation of C53
    const clean = adjustAnchoredRule({ sqref: areasOf(dv.sqref), formulas: dv.formulas }, op, 'P', { kind: 'dv', type: 'whole', inheritFormat: false });
    expect(cover(clean.map((x) => ({ kind: 'dv', type: 'whole', sqref: x.sqref, formulas: x.formulas }))).has('4,53')).toBe(false);
  });
});

describe('AnchoredRuleOptions helpers', () => {
  it('ruleGrowsOnInsert / ruleSpreadsOnMove follow the rule type', () => {
    for (const t of NO_SPREAD_CF_TYPES) {
      expect(ruleSpreadsOnMove({ kind: 'cf', type: t })).toBe(false);
      expect(ruleGrowsOnInsert({ kind: 'cf', type: t })).toBe(true);
    }
    expect([...NO_SPREAD_CF_TYPES].sort()).toEqual(['aboveAverage', 'colorScale', 'dataBar', 'duplicateValues', 'iconSet', 'top10', 'uniqueValues']);
    for (const t of ['cellIs', 'expression', 'containsText', 'containsBlanks', 'notContainsBlanks', 'containsErrors', 'beginsWith', 'timePeriod']) {
      expect(ruleSpreadsOnMove({ kind: 'cf', type: t })).toBe(true);
    }
    // a data validation of any type spreads / grows; a CF type name means nothing for kind 'dv'
    for (const t of ['whole', 'list', 'custom', 'decimal', 'date', 'textLength', 'colorScale']) {
      expect(ruleSpreadsOnMove({ kind: 'dv', type: t })).toBe(true);
      expect(ruleGrowsOnInsert({ kind: 'dv', type: t })).toBe(true);
    }
    expect(ruleSpreadsOnMove({ kind: 'cf', type: 'cellIs', inheritFormat: false })).toBe(false);
    expect(ruleGrowsOnInsert({ kind: 'dv', type: 'list', inheritFormat: false })).toBe(false);
    // unknown / missing type spreads; the round-1 option names still work
    expect(ruleSpreadsOnMove({})).toBe(true);
    expect(ruleSpreadsOnMove({ kind: 'cf', cfType: 'colorScale' })).toBe(false);
    expect(ruleSpreadsOnMove({ kind: 'cf', cfType: 'cellValue' })).toBe(false);
    expect(ruleSpreadsOnMove({ kind: 'cf', cfType: 'cellIs' })).toBe(true);
  });
});
