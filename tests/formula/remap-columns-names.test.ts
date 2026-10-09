/**
 * FormulaRemapper: columns (rule c: imported columns, key columns read through the base partner) and defined names (rule d).
 * Helpers and the source layout: remapHelpers.ts (A 学生ID, B 氏名 not imported; C..F imported to E..H).
 */
import { describe, expect, it } from 'vitest';
import type { RemapReason } from '../../src/core/formula/remap';
import { identity, names, ok, reversed, run, why } from './remapHelpers';

describe('remap: columns (rule c)', () => {
  it.each<[string]>([['A5'], ['B5&"様"'], ['SUM(B5:D5)'], ['VLOOKUP(A5,$A$2:$F$11,3,FALSE)'], ['COUNTIF($A$2:$A$11,A5)'], ['SUM(A:C)'], ['LEFT(B5,1)']])(
    'a non-imported column cannot be read: %s',
    (f) => {
      expect(why(f)).toBe('notImported');
    },
  );

  it('imported columns that are not consecutive in the merged sheet cannot be read as one area', () => {
    const gap = { colMap: new Map([[3, 5], [4, 9], [5, 10]]) };
    expect(ok('C5+D5', 5, 5, gap)).toBe('E5+I5');
    expect(why('SUM(C5:D5)', 5, 5, gap)).toBe('notImported');
    expect(ok('SUM(D5:E5)', 5, 5, gap)).toBe('SUM(I5:J5)');
    const swapped = { colMap: new Map([[3, 6], [4, 5]]) };
    expect(why('SUM(C5:D5)', 5, 5, swapped)).toBe('notImported');
    expect(ok('C5-D5', 5, 5, swapped)).toBe('F5-E5');
  });

  it('a lookup keeps its column index right because the area keeps the column order', () => {
    expect(ok('VLOOKUP(C5,$C$2:$F$11,4,FALSE)')).toBe('VLOOKUP(E5,$E$2:$H$11,4,FALSE)');
    expect(ok('INDEX($C$2:$F$11,3,2)')).toBe('INDEX($E$2:$H$11,3,2)');
  });

  it('when the key columns are imported too, a lookup over the whole source table keeps working', () => {
    const all = { colMap: new Map([[1, 3], [2, 4], [3, 5], [4, 6], [5, 7], [6, 8]]) };
    expect(ok('VLOOKUP(A5,$A$2:$F$11,3,FALSE)', 5, 5, all)).toBe('VLOOKUP(C5,$C$2:$H$11,3,FALSE)');
    expect(ok('COUNTIF($A$2:$A$11,A5)', 5, 5, all)).toBe('COUNTIF($C$2:$C$11,C5)');
  });

  describe('key columns read through the base partner column (the caller proved the texts equal)', () => {
    const partner = { partnerCols: new Map([[1, 1], [2, 2]]) };
    it.each<[string, string]>([
      ['A5', 'A5'],
      ['B5&"様"', 'B5&"様"'],
      ['A5&"様"', 'A5&"様"'],
      ['$A5', '$A5'],
      ['IF(COUNTIF($A$2:$A$11,A5)>1,"重複","")', 'IF(COUNTIF($A$2:$A$11,A5)>1,"重複","")'],
      ['COUNTIF($B:$B,B5)', 'COUNTIF($B:$B,B5)'],
      ['INDEX($D$2:$D$11,MATCH(A5,$A$2:$A$11,0))', 'INDEX($F$2:$F$11,MATCH(A5,$A$2:$A$11,0))'],
      ['LEFT(B5,1)', 'LEFT(B5,1)'],
      ['COUNTA($A$2:A5)', 'COUNTA($A$2:A5)'],
    ])('%s -> %s', (f, want) => {
      expect(ok(f, 5, 5, partner)).toBe(want);
    });

    it('is reported as substituted; a multi-column area with a partner column is not remapped', () => {
      const r = run('A5&C5', 5, 5, partner);
      expect(r.ok && r.substituted).toBe(true);
      const q = run('C5', 5, 5, partner);
      expect(q.ok && q.substituted).toBe(false);
      expect(why('SUM(A5:C5)', 5, 5, partner)).toBe('notImported');
      expect(why('VLOOKUP(A5,$A$2:$F$11,3,FALSE)', 5, 5, partner)).toBe('notImported');
    });

    it('follows the row mapping for a reference to another row', () => {
      expect(ok('A4', 5, 8, { ...partner, rowPairs: reversed() })).toBe('A9');
    });

    it('the partner column can be any merged column', () => {
      expect(ok('A5&B5', 5, 5, { partnerCols: new Map([[1, 3], [2, 4]]) })).toBe('C5&D5');
    });

    it('a whole-column reference to a key column does not need empty cells above / below in the SOURCE', () => {
      expect(ok('COUNTIF($A:$A,A5)', 5, 5, { ...partner, hasContentOutside: () => true })).toBe('COUNTIF($A:$A,A5)');
    });
  });

  it('a column beside the table is a cell outside the table', () => {
    expect(why('H5')).toBe('externalCell');
    expect(why('SUM(H2:H11)')).toBe('externalCell');
    expect(why('$J$1')).toBe('externalCell');
    expect(why('C5*$H$1')).toBe('externalCell');
  });

  it('whole-column references: the imported column must be empty above and below the table', () => {
    expect(ok('SUM(C:C)')).toBe('SUM(E:E)');
    expect(why('SUM(C:C)', 5, 5, { hasContentOutside: (c) => c === 3 })).toBe('externalCell');
    expect(ok('SUM(D:D)', 5, 5, { hasContentOutside: (c) => c === 3 })).toBe('SUM(F:F)');
    expect(why('SUM(C:E)', 5, 5, { hasContentOutside: (c) => c === 4 })).toBe('externalCell');
  });
});

describe('remap: defined names (rule d)', () => {
  it.each<[string, string, Array<[string, string | null, string]>, string]>([
    ['SUM(評価列)', 'SUM($E$2:$E$11)', [['評価列', null, "'評価'!$C$2:$C$11"]], 'column name'],
    ['SUM(Scores)', 'SUM($E$2:$E$11)', [['scores', null, '評価!$C$2:$C$11']], 'case-insensitive'],
    ['C5*税率', 'E5*0.08', [['税率', null, '0.08']], 'number constant'],
    ['C5*税率', 'E5*0.08', [['税率', null, '=0.08']], 'leading ='],
    ['C5&ラベル', 'E5&"点"', [['ラベル', null, '"点"']], 'text constant'],
    ['C5*係数', 'E5*(-0.5)', [['係数', null, '-0.5']], 'negative constant is parenthesised'],
    ['C5*Rate', 'E5*(8%)', [['Rate', null, '8%']], 'percent constant'],
    ['IF(Flag,C5,0)', 'IF(TRUE,E5,0)', [['Flag', null, 'TRUE']], 'boolean constant'],
    ['SUM(Scores)', 'SUM($E$2:$G$11)', [['Scores', null, '評価!$C$2:$E$11']], 'multi-column name'],
    ['AVERAGE(Scores)', 'AVERAGE($F$2:$F$11)', [['Scores', '評価', '評価!$D$2:$D$11'], ['Scores', null, '評価!$C$2:$C$11']], 'sheet scope wins over workbook scope'],
    ['SUM(Scores)', 'SUM($E$2:$E$11)', [['Scores', 'Other', '評価!$D$2:$D$11'], ['Scores', null, '評価!$C$2:$C$11']], 'a name scoped to another sheet is invisible'],
    ['SUM(評価!Scores)', 'SUM($F$2:$F$11)', [['Scores', '評価', '評価!$D$2:$D$11']], 'qualified with the sheet'],
    ['Pick+C5', '$E$5+E5', [['Pick', null, '評価!$C$5']], 'single cell name inside the table'],
    ['SUM(Rng)', 'SUM($E:$E)', [['Rng', null, '評価!$C:$C']], 'whole-column name'],
  ])('%s -> %s (%s)', (f, want, list) => {
    expect(ok(f, 5, 5, names(list))).toBe(want);
  });

  it.each<[string, RemapReason, Array<[string, string | null, string]>, string]>([
    ['C5*係数', 'externalCell', [['係数', null, '評価!$H$1']], 'a rate in a side cell'],
    ['C5*税率', 'externalCell', [['税率', null, '評価!$P$1']], 'a rate far beside the table'],
    ['VLOOKUP(C5,区分表,2,FALSE)', 'otherSheet', [['区分表', null, 'Master!$A$2:$B$100']], 'a range on another sheet'],
    ['SUM(動的)', 'unresolvedName', [['動的', null, "OFFSET('評価'!$C$1,1,0,10,1)"]], 'a formula name'],
    ['SUM(動的)', 'unresolvedName', [['動的', null, 'SUM(評価!$C$2:$C$11)']], 'a formula name with a range'],
    ['C5*Undefined', 'unresolvedName', [], 'an undefined name'],
    ['SUM(Scores)', 'unresolvedName', [['Scores', null, 'C2:C11']], 'a relative name without a sheet'],
    ['SUM(Scores)', 'unresolvedName', [['Scores', null, '評価!C2:C11']], 'a relative name (offsets from the cell that uses it)'],
    ['SUM(Scores)', 'unresolvedName', [['Scores', null, 'Other']], 'a name that refers to a name'],
    ['SUM(Scores)', 'unresolvedName', [['Scores', 'Other', '評価!$C$2:$C$11']], 'a name scoped to another sheet only'],
    ['C5*Ext', 'unsafe', [['Ext', null, '[1]Sheet1!$A$1']], 'a name that reads another workbook'],
    ['C5*Err', 'unresolvedName', [['Err', null, '#REF!']], 'a broken name'],
    ['C5*Arr', 'unresolvedName', [['Arr', null, '{1,2,3}']], 'an array constant name'],
    ['C5*Two', 'unresolvedName', [['Two', null, '1+1']], 'an expression name'],
  ])('%s -> %s (%s)', (f, want, list) => {
    expect(why(f, 5, 5, names(list))).toBe(want);
  });

  it('a name qualified with another sheet or another workbook', () => {
    expect(why('Other!Scores', 5, 5, names([['Scores', 'Other', '評価!$C$2:$C$11']]))).toBe('otherSheet');
    expect(why('[1]!Scores', 5, 5, names([['Scores', null, '評価!$C$2:$C$11']]))).toBe('unsafe');
  });

  it('a name over the whole table follows the same rules as the range it stands for (widened / partial)', () => {
    const miss = identity().filter(([r]) => r !== 7);
    const r = run('SUM(Scores)', 5, 5, { ...names([['Scores', null, '評価!$C$2:$C$11']]), rowPairs: miss });
    expect(r.ok && r.widened && !r.exact).toBe(true);
    expect(why('SUM(Part)', 5, 5, { ...names([['Part', null, '評価!$C$2:$C$5']]), rowPairs: reversed() })).toBe('rowRelation');
  });

  it('LET / LAMBDA parameters and eta-reduced functions are not names', () => {
    expect(ok('LET(_xlpm.x,C5*2,_xlpm.x+1)')).toBe('LET(_xlpm.x,E5*2,_xlpm.x+1)');
    expect(ok('BYROW(C2:E11,_xleta.SUM)')).toBe('BYROW(E2:G11,_xleta.SUM)');
    expect(why('BYROW(C2:E11,_xleta.ROW)')).toBe('positional');
    expect(ok('_xlfn.LAMBDA(_xlpm.a,_xlpm.a+1)(C5)')).toBe('_xlfn.LAMBDA(_xlpm.a,_xlpm.a+1)(E5)');
  });

  it('a named LAMBDA called like a function cannot be resolved; a built-in function is never a name', () => {
    expect(why('Double(C5)', 5, 5, names([['Double', null, '_xlfn.LAMBDA(_xlpm.x,_xlpm.x*2)']]))).toBe('unresolvedName');
    expect(ok('SUM(C5)', 5, 5, names([['Sum', null, '1']]))).toBe('SUM(E5)');
  });
});
