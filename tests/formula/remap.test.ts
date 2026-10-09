/**
 * FormulaRemapper (src/core/formula/remap.ts, round 8): how a formula of a SOURCE sheet is rewritten for the merged sheet — own row, other
 * rows through T(r), ranges over the whole table, partial ranges (helpers and the source layout: remapHelpers.ts).
 */
import { describe, expect, it } from 'vitest';
import { identity, ok, range, reversed, run, shifted, why } from './remapHelpers';

describe('remap: the formula\'s own row (T(sr) = target) and column mapping', () => {
  it.each<[string, string]>([
    ['C5', 'E5'],
    ['C5+D5', 'E5+F5'],
    ['$C$5', '$E$5'],
    ['C$5', 'E$5'],
    ['$C5', '$E5'],
    ['C5:E5', 'E5:G5'],
    ['E5:C5', 'G5:E5'],
    ['SUM(C5:F5)', 'SUM(E5:H5)'],
    ['AVERAGE(C5:E5)', 'AVERAGE(E5:G5)'],
    ['IF(C5>=80,"優","可")', 'IF(E5>=80,"優","可")'],
    ['ROUND(C5/D5,2)', 'ROUND(E5/F5,2)'],
    ['"C5"&C5', '"C5"&E5'],
    ['C5 + D5', 'E5 + F5'],
    ['c5+d5', 'E5+F5'],
    ['(C5+D5)*E5/F5', '(E5+F5)*G5/H5'],
    ['-C5', '-E5'],
    ['C5%', 'E5%'],
    ['C5^2', 'E5^2'],
    ['C5&"点"', 'E5&"点"'],
    ['C5=D5', 'E5=F5'],
    ['C5<>D5', 'E5<>F5'],
    ['IF(AND(C5>0,D5>0),C5*D5,0)', 'IF(AND(E5>0,F5>0),E5*F5,0)'],
    ['IFERROR(C5/D5,"")', 'IFERROR(E5/F5,"")'],
    ['TEXT(C5,"0.0")&"点"', 'TEXT(E5,"0.0")&"点"'],
    ['1+1', '1+1'],
    ['"abc"', '"abc"'],
    ['TRUE', 'TRUE'],
    ['NA()', 'NA()'],
    ['IF(C5>1,#N/A,1)', 'IF(E5>1,#N/A,1)'],
    ['_xlfn.CONCAT(C5,"-",D5)', '_xlfn.CONCAT(E5,"-",F5)'],
    ['_xlfn.IFS(C5>1,"a",TRUE,"b")', '_xlfn.IFS(E5>1,"a",TRUE,"b")'],
    ['SUM(C5,D5,E5)', 'SUM(E5,F5,G5)'],
    ['{1,2,3}', '{1,2,3}'],
    ['C5:C5', 'E5:E5'],
  ])('%s -> %s', (f, want) => {
    expect(ok(f)).toBe(want);
  });

  it('a row that moved: the own row is the target row', () => {
    expect(ok('C5+D5', 5, 8, { rowPairs: reversed() })).toBe('E8+F8');
    expect(ok('SUM(C5:E5)', 5, 8, { rowPairs: reversed() })).toBe('SUM(E8:G8)');
    expect(ok('$C$5', 5, 8, { rowPairs: reversed() })).toBe('$E$8');
  });

  it('the result is FILE form without "=" and unqualified (the merged sheet is the formula\'s own sheet)', () => {
    expect(ok("'評価'!C5")).toBe('E5');
    expect(ok('評価!C5*2')).toBe('E5*2');
    expect(ok("'評価'!C$2:C$11")).toBe('E$2:E$11');
    expect(ok("SUM('評価'!C5:E5)")).toBe('SUM(E5:G5)');
  });
});

describe('remap: references to OTHER rows go through T(r) (rows matched by key land on other rows)', () => {
  const rev = { rowPairs: reversed() };
  it.each<[string, string]>([
    ['C4', 'E9'],
    ['C6', 'E7'],
    ['C4-C5', 'E9-E8'],
    ['C5-C6', 'E8-E7'],
    ['C$2', 'E$11'],
    ['$C$11', '$E$2'],
    ['C6/C4', 'E7/E9'],
    ['D4', 'F9'],
    ['C4:E4', 'E9:G9'],
    ['IF(C5>C4,"up","down")', 'IF(E8>E9,"up","down")'],
  ])('reversed rows: %s -> %s (formula on source row 5, merged row 8)', (f, want) => {
    expect(ok(f, 5, 8, rev)).toBe(want);
  });

  it('identity mapping leaves rows alone', () => {
    expect(ok('C4-C5')).toBe('E4-E5');
    expect(ok('C6/C4')).toBe('E6/E4');
  });

  it('the header row maps to the base header row', () => {
    expect(ok('C$1')).toBe('E$1');
    expect(ok('C1&":"&C5')).toBe('E1&":"&E5');
    expect(ok('C$3', 5, 5, { region: { headerRow: 3, firstDataRow: 4, lastDataRow: 11, firstCol: 1, lastCol: 6 }, target: { headerRow: 2, firstDataRow: 3, lastDataRow: 10 }, rowPairs: range(4, 11).map((r) => [r, r - 1]) })).toBe('E$2');
  });

  it('a data row that was not placed has no merged row: rowRelation', () => {
    const miss = identity().filter(([r]) => r !== 7);
    expect(why('C7', 5, 5, { rowPairs: miss })).toBe('rowRelation');
    expect(why('C6-C7', 6, 6, { rowPairs: miss })).toBe('rowRelation');
    expect(ok('C6', 5, 5, { rowPairs: miss })).toBe('E6');
  });

  it('cells outside the table (footer, side cells, title rows) are externalCell', () => {
    expect(why('C14')).toBe('externalCell');
    expect(why('C$34')).toBe('externalCell');
    expect(why('C5*$H$1')).toBe('externalCell');
    expect(why('C5*$H5')).toBe('externalCell');
    expect(why('$P$1')).toBe('externalCell');
    expect(why('C5', 5, 5, { region: { headerRow: 3, firstDataRow: 4, lastDataRow: 11, firstCol: 1, lastCol: 6 } })).toBe('ok');
    expect(why('C1', 5, 5, { region: { headerRow: 3, firstDataRow: 4, lastDataRow: 11, firstCol: 1, lastCol: 6 } })).toBe('externalCell');
    expect(why('C2:E2', 5, 5, { region: { headerRow: 3, firstDataRow: 4, lastDataRow: 11, firstCol: 1, lastCol: 6 } })).toBe('externalCell');
  });

  it('whole-row references are externalCell', () => {
    expect(why('SUM(5:5)')).toBe('externalCell');
    expect(why('MATCH("x",$1:$1,0)')).toBe('externalCell');
  });

  it('a source row placed twice (duplicate key in the base) has no unique merged row, but the own row is always the target', () => {
    const dup: Array<[number, number]> = [...identity(), [4, 12]];
    expect(ok('C5+D5', 5, 5, { rowPairs: dup })).toBe('E5+F5');
    expect(why('C4', 5, 5, { rowPairs: dup })).toBe('rowRelation');
    expect(ok('C4+D4', 4, 4, { rowPairs: dup })).toBe('E4+F4');
    expect(ok('C4+D4', 4, 12, { rowPairs: dup })).toBe('E12+F12');
  });
});

describe('remap: ranges over the whole data table become the whole merged data table (rule b)', () => {
  it.each<[string, string]>([
    ['SUM(C2:C11)', 'SUM(E2:E11)'],
    ['SUM(C$2:C$11)', 'SUM(E$2:E$11)'],
    ['SUM($C$2:$C$11)', 'SUM($E$2:$E$11)'],
    ['AVERAGE(D2:D11)', 'AVERAGE(F2:F11)'],
    ['AVERAGE(D1:D11)', 'AVERAGE(F1:F11)'],
    ['MAX(C:C)', 'MAX(E:E)'],
    ['MAX($C:$C)', 'MAX($E:$E)'],
    ['SUM(C:E)', 'SUM(E:G)'],
    ['SUM(C2:C1000)', 'SUM(E2:E11)'],
    ['COUNT(C2:C100)', 'COUNT(E2:E11)'],
    ['SUM(C2:E11)', 'SUM(E2:G11)'],
    ['RANK(C5,$C$2:$C$11)', 'RANK(E5,$E$2:$E$11)'],
    ['RANK(C5,C$2:C$11,0)', 'RANK(E5,E$2:E$11,0)'],
    ['LARGE($C$2:$C$11,3)', 'LARGE($E$2:$E$11,3)'],
    ['SMALL($C$2:$C$11,1)', 'SMALL($E$2:$E$11,1)'],
    ['COUNTIF($C$2:$C$11,">="&C5)', 'COUNTIF($E$2:$E$11,">="&E5)'],
    ['SUMIF($C$2:$C$11,">"&C5,$D$2:$D$11)', 'SUMIF($E$2:$E$11,">"&E5,$F$2:$F$11)'],
    ['COUNTIFS($C$2:$C$11,C5,$D$2:$D$11,D5)', 'COUNTIFS($E$2:$E$11,E5,$F$2:$F$11,F5)'],
    ['SUMPRODUCT(($C$2:$C$11>C5)*1)', 'SUMPRODUCT(($E$2:$E$11>E5)*1)'],
    ['C5/SUM(C$2:C$11)', 'E5/SUM(E$2:E$11)'],
    ['C5-AVERAGE($C$2:$C$11)', 'E5-AVERAGE($E$2:$E$11)'],
    ['VLOOKUP(C5,$C$2:$F$11,3,FALSE)', 'VLOOKUP(E5,$E$2:$H$11,3,FALSE)'],
    ['VLOOKUP(C5,$C$2:$F$11,4,0)', 'VLOOKUP(E5,$E$2:$H$11,4,0)'],
    ['HLOOKUP(C$1,$C$1:$F$11,2,FALSE)', 'HLOOKUP(E$1,$E$1:$H$11,2,FALSE)'],
    ['INDEX($D$2:$D$11,MATCH(C5,$C$2:$C$11,0))', 'INDEX($F$2:$F$11,MATCH(E5,$E$2:$E$11,0))'],
    ['_xlfn.XLOOKUP(C5,$C$2:$C$11,$D$2:$D$11,"-")', '_xlfn.XLOOKUP(E5,$E$2:$E$11,$F$2:$F$11,"-")'],
    ['LOOKUP(C5,$C$2:$C$11,$D$2:$D$11)', 'LOOKUP(E5,$E$2:$E$11,$F$2:$F$11)'],
    ['MEDIAN(C2:C11)', 'MEDIAN(E2:E11)'],
    ['STDEV(C2:C11)', 'STDEV(E2:E11)'],
    ['(C5-AVERAGE(C2:C11))/STDEV(C2:C11)*10+50', '(E5-AVERAGE(E2:E11))/STDEV(E2:E11)*10+50'],
    ['TEXTJOIN(",",TRUE,C2:C11)', 'TEXTJOIN(",",TRUE,E2:E11)'],
  ])('%s -> %s', (f, want) => {
    expect(ok(f)).toBe(want);
  });

  it('the merged table includes the appended rows; the header row maps to the base header row', () => {
    const o = { target: { headerRow: 1, firstDataRow: 2, lastDataRow: 14 } };
    expect(ok('SUM(C2:C11)', 5, 5, o)).toBe('SUM(E2:E14)');
    expect(ok('SUM(C:C)', 5, 5, o)).toBe('SUM(E:E)');
    expect(ok('AVERAGE(C1:C11)', 5, 5, o)).toBe('AVERAGE(E1:E14)');
    // base header on another row than the source's
    const o2 = { target: { headerRow: 3, firstDataRow: 4, lastDataRow: 13 } };
    expect(ok('AVERAGE(C1:C11)', 5, 6, o2)).toBe('AVERAGE(E3:E13)');
    expect(ok('SUM(C2:C11)', 5, 6, o2)).toBe('SUM(E4:E13)');
  });

  it('a reversed order does not matter for a whole-table range', () => {
    expect(ok('SUM(C2:C11)', 5, 8, { rowPairs: reversed() })).toBe('SUM(E2:E11)');
    expect(ok('RANK(C5,$C$2:$C$11)', 5, 8, { rowPairs: reversed() })).toBe('RANK(E8,$E$2:$E$11)');
  });

  it('a range that overshoots the data but covers only empty cells below is the whole table; with content below it is not', () => {
    const below = (col: number, r1: number, r2: number) => col === 3 && r1 <= 14 && r2 >= 14; // C14 holds a footer total
    expect(why('SUM(C2:C20)', 5, 5, { hasContentOutside: below })).toBe('externalCell');
    expect(why('SUM(C2:C13)', 5, 5, { hasContentOutside: below })).toBe('ok');
    expect(why('SUM(D2:D20)', 5, 5, { hasContentOutside: below })).toBe('ok');
    expect(why('SUM(C:C)', 5, 5, { hasContentOutside: below })).toBe('externalCell');
    expect(why('SUM(D:D)', 5, 5, { hasContentOutside: below })).toBe('ok');
    expect(why('SUM(C:D)', 5, 5, { hasContentOutside: below })).toBe('externalCell');
  });

  it('title rows above the header: a range starting above it needs those cells to be empty', () => {
    const r = { region: { headerRow: 3, firstDataRow: 4, lastDataRow: 11, firstCol: 1, lastCol: 6 }, target: { headerRow: 3, firstDataRow: 4, lastDataRow: 11 }, rowPairs: range(4, 11).map((x) => [x, x] as [number, number]) };
    expect(ok('SUM(C1:C11)', 5, 5, r)).toBe('SUM(E3:E11)');
    expect(why('SUM(C1:C11)', 5, 5, { ...r, hasContentOutside: (c, a, b) => c === 3 && a <= 1 && b >= 1 })).toBe('externalCell');
    expect(ok('SUM(C3:C11)', 5, 5, r)).toBe('SUM(E3:E11)');
    expect(ok('SUM(C4:C11)', 5, 5, r)).toBe('SUM(E4:E11)');
  });

  it('a 1-row table: D2:D2 is one cell, D2:D3 runs past it', () => {
    const one = { region: { headerRow: 1, firstDataRow: 2, lastDataRow: 2, firstCol: 1, lastCol: 6 }, target: { headerRow: 1, firstDataRow: 2, lastDataRow: 2 }, rowPairs: [[2, 2]] as Array<[number, number]> };
    expect(ok('SUM(C2:C2)', 2, 2, one)).toBe('SUM(E2:E2)');
    expect(ok('SUM(C2:C3)', 2, 2, one)).toBe('SUM(E2:E2)');
  });

  it('a range partly inside, partly beyond the table with content: externalCell / rowRelation', () => {
    expect(why('SUM(C5:C14)', 5, 5, { hasContentOutside: () => false })).toBe('externalCell'); // rows 12..14 are not part of the table
    expect(why('SUM(C13:C20)')).toBe('externalCell');
  });
});

describe('remap: partial ranges only when T(a)..T(b) are consecutive and ascending', () => {
  it.each<[string, string]>([
    ['SUM(C$2:C5)', 'SUM(E$2:E5)'],
    ['SUM(C$2:C2)', 'SUM(E$2:E2)'],
    ['SUM(C3:C7)', 'SUM(E3:E7)'],
    ['AVERAGE(C$2:C5)', 'AVERAGE(E$2:E5)'],
    ['COUNTA(C$2:C5)', 'COUNTA(E$2:E5)'],
    ['SUM(C1:C5)', 'SUM(E1:E5)'],
    ['SUM(C5:C11)', 'SUM(E5:E11)'],
    ['SUM(D3:E7)', 'SUM(F3:G7)'],
  ])('identity: %s -> %s', (f, want) => {
    expect(ok(f)).toBe(want);
  });

  it('a run that moved as a block stays a range', () => {
    expect(ok('SUM(C3:C7)', 5, 15, { rowPairs: shifted(), target: { headerRow: 1, firstDataRow: 12, lastDataRow: 21 } })).toBe('SUM(E13:E17)');
  });

  it.each<[string, string]>([
    ['SUM(C$2:C5)', 'rowRelation'],
    ['SUM(C3:C7)', 'rowRelation'],
    ['SUM(C5:C11)', 'rowRelation'],
  ])('reversed order: %s is %s', (f, want) => {
    expect(why(f, 5, 8, { rowPairs: reversed() })).toBe(want);
  });

  it('a shuffled order or a gap breaks the run', () => {
    const shuf: Array<[number, number]> = [[2, 5], [3, 2], [4, 9], [5, 3], [6, 4], [7, 6], [8, 7], [9, 8], [10, 10], [11, 11]];
    expect(why('SUM(C$2:C5)', 5, 3, { rowPairs: shuf })).toBe('rowRelation');
    expect(why('SUM(C6:C9)', 6, 4, { rowPairs: shuf })).toBe('rowRelation');
    expect(ok('SUM(C7:C9)', 7, 6, { rowPairs: shuf })).toBe('SUM(E6:E8)'); // 7 -> 6, 8 -> 7, 9 -> 8 : a consecutive block
    expect(ok('SUM(C5:C6)', 5, 3, { rowPairs: shuf })).toBe('SUM(E3:E4)');
    const gap = identity().filter(([r]) => r !== 7);
    expect(why('SUM(C5:C9)', 5, 5, { rowPairs: gap })).toBe('rowRelation');
    expect(ok('SUM(C2:C6)', 5, 5, { rowPairs: gap })).toBe('SUM(E2:E6)');
  });

  it('a partial range outside the table is not remapped', () => {
    expect(why('SUM(C9:C14)')).toBe('externalCell');
    expect(why('SUM(C12:C14)')).toBe('externalCell');
  });

  it('a reversed written area keeps its orientation', () => {
    expect(ok('SUM(C7:C3)')).toBe('SUM(E7:E3)');
    expect(ok('SUM(C11:C2)', 5, 8, { rowPairs: reversed() })).toBe('SUM(E11:E2)');
  });
});

describe('remap: exactness of the cached result (rule g) and the widened flag', () => {
  const miss = identity().filter(([r]) => r !== 7);
  const flags = (f: string, over: Parameters<typeof run>[3] = {}, sr = 5, tr = sr) => {
    const r = run(f, sr, tr, over);
    if (!r.ok) throw new Error(f + ' ' + r.reason);
    return { exact: r.exact, widened: r.widened };
  };

  it('same-row, other-row and consecutive partial references read the very same values', () => {
    expect(flags('C5+D5')).toEqual({ exact: true, widened: false });
    expect(flags('C4+D6', {}, 5, 8)).toEqual({ exact: true, widened: false });
    expect(flags('SUM(C$2:C5)')).toEqual({ exact: true, widened: false });
    expect(flags('C4-C5', { rowPairs: reversed() }, 5, 8)).toEqual({ exact: true, widened: false });
  });

  it('a whole-table range over exactly the same rows keeps the result (reorder or not)', () => {
    expect(flags('SUM(C2:C11)')).toEqual({ exact: true, widened: false });
    expect(flags('SUM(C2:C11)', { rowPairs: reversed() }, 5, 8)).toEqual({ exact: true, widened: false });
    expect(flags('RANK(C5,$C$2:$C$11)', { rowPairs: reversed() }, 5, 8)).toEqual({ exact: true, widened: false });
  });

  it('source rows that were not placed change what a whole-table range sums: not exact, widened', () => {
    expect(flags('SUM(C2:C11)', { rowPairs: miss })).toEqual({ exact: false, widened: true });
    expect(flags('C5/SUM(C$2:C$11)', { rowPairs: miss })).toEqual({ exact: false, widened: true });
    expect(flags('MAX(C:C)', { rowPairs: miss })).toEqual({ exact: false, widened: true });
  });

  it('rows added to the table (appended rows, other files) widen the range but not an aggregate\'s result', () => {
    const bigger = { target: { headerRow: 1, firstDataRow: 2, lastDataRow: 14 } };
    expect(flags('SUM(C2:C11)', bigger)).toEqual({ exact: true, widened: true });
    expect(flags('AVERAGE(C2:C11)', bigger)).toEqual({ exact: true, widened: true });
    expect(flags('RANK(C5,$C$2:$C$11)', bigger)).toEqual({ exact: true, widened: true });
  });

  it('a position (MATCH on its own, INDEX with a literal position) is exact only while the layout is the same', () => {
    expect(flags('MATCH(C5,$C$2:$C$11,0)')).toEqual({ exact: true, widened: false });
    expect(flags('MATCH(C5,$C$2:$C$11,0)', { rowPairs: reversed() }, 5, 8)).toEqual({ exact: false, widened: false });
    expect(flags('INDEX($C$2:$C$11,3)', { rowPairs: reversed() }, 5, 8)).toEqual({ exact: false, widened: false });
    expect(flags('INDEX($C$2:$C$11,3)')).toEqual({ exact: true, widened: false });
    // the position is consumed by an INDEX: value based
    expect(flags('INDEX($D$2:$D$11,MATCH(C5,$C$2:$C$11,0))', { rowPairs: reversed() }, 5, 8)).toEqual({ exact: true, widened: false });
    expect(flags('INDEX($D$2:$D$11,MATCH(C5,$C$2:$C$11,0))', { rowPairs: miss })).toEqual({ exact: false, widened: true });
  });

  it('a row count (ROWS, COUNTBLANK) is exact only for a table of the same size', () => {
    expect(flags('ROWS($C$2:$C$11)')).toEqual({ exact: true, widened: false });
    expect(flags('ROWS($C$2:$C$11)', { target: { headerRow: 1, firstDataRow: 2, lastDataRow: 14 } })).toEqual({ exact: false, widened: true });
    expect(flags('COUNTBLANK(C2:C11)', { target: { headerRow: 1, firstDataRow: 2, lastDataRow: 14 } })).toEqual({ exact: false, widened: true });
  });

  it('a header cell is exact only while the merged header text is the source\'s', () => {
    expect(flags('C$1&C5', { headerSame: () => true })).toEqual({ exact: true, widened: false });
    expect(flags('C$1&C5', { headerSame: () => false })).toEqual({ exact: false, widened: false });
    expect(flags('C$1&C5', {})).toEqual({ exact: true, widened: false });
    expect(flags('A$1&C5', { partnerCols: new Map([[1, 1]]) })).toEqual({ exact: false, widened: false });
  });

  it('allPlaced / sameCount / layoutIdentical / orderPreserved describe the mapping', async () => {
    const { FormulaRemapper } = await import('../../src/core/formula/remap');
    const { spec } = await import('./remapHelpers');
    const a = new FormulaRemapper(spec());
    expect([a.allPlaced, a.sameCount, a.layoutIdentical, a.orderPreserved]).toEqual([true, true, true, true]);
    const b = new FormulaRemapper(spec({ rowPairs: reversed() }));
    expect([b.allPlaced, b.sameCount, b.layoutIdentical, b.orderPreserved]).toEqual([true, true, false, false]);
    const c = new FormulaRemapper(spec({ rowPairs: miss }));
    expect([c.allPlaced, c.sameCount, c.layoutIdentical, c.orderPreserved]).toEqual([false, true, false, true]);
    const d = new FormulaRemapper(spec({ target: { headerRow: 1, firstDataRow: 2, lastDataRow: 14 } }));
    expect([d.allPlaced, d.sameCount, d.layoutIdentical, d.orderPreserved]).toEqual([true, false, false, true]);
    const e = new FormulaRemapper(spec({ rowPairs: miss, dataRows: range(2, 11).filter((r) => r !== 7) }));
    expect(e.allPlaced).toBe(true); // row 7 holds no data
  });
});
