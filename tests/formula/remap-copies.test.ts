/**
 * FormulaRemapper with the copies of the source's sheets (round 9, includeSourceSheets, RemapSpec.sheetCopies): what the merged table cannot
 * keep reads the COPY of the sheet at the original coordinates. Source layout and helpers: remapHelpers.ts (sheet 評価, header row 1, data rows
 * 2..11, columns A..F; the default import is C..F -> E..H; PARTIAL imports only C, D -> E, F; COPIES names the copies).
 */
import { describe, expect, it } from 'vitest';
import { FormulaRemapper } from '../../src/core/formula/remap';
import { COPIES, KEYS, PARTIAL, identity, names, okC, reversed, run, runC, spec, viaC, why, whyC, without } from './remapHelpers';

describe('copies: no copies = round 8, and what the merged table keeps is never touched', () => {
  const FORMULAS = [
    'C5', 'C5+D5', '$C$5', 'SUM(C5:F5)', 'IF(C5>=80,"優","可")', 'C4-C5', 'C6/C4', 'C$1', 'C1&":"&C5', 'SUM(C2:C11)', 'SUM(C$2:C$11)', 'C5/SUM(C2:C11)',
    'AVERAGE(C:C)', 'RANK(C5,C$2:C$11)', 'COUNTIF(C2:C11,">"&C5)', 'SUM(C$2:C5)', 'VLOOKUP(C5,C2:F11,2,FALSE)', 'INDEX(D2:D11,MATCH(C5,C2:C11,0))',
    'SUMIF(C2:C11,">5",D2:D11)', 'SUMPRODUCT(C2:C11,D2:D11)', 'MAX(C5:F5)-MIN(C5:F5)', 'TEXT(C5,"0.0")&"点"', 'C5%', '1+1', '"abc"', 'NA()',
    '_xlfn.CONCAT(C5,"-",D5)', 'C5:C5', 'ROWS(C2:C11)', 'MATCH(C5,C2:C11,0)',
  ];
  it.each(FORMULAS)('%s: the same result with and without copies, no viaCopy', (f) => {
    const a = run(f);
    expect(a.ok).toBe(true);
    expect(runC(f)).toEqual(a);
    expect(viaC(f)).toBe(false);
  });
  it.each(['C14', 'C5*$H$1', 'E5*2', 'レート!A1', 'ROW()-1', 'SUM(C$2:C5)', 'INDIRECT("C5")'])('%s: an empty / absent sheetCopies map is round 8', (f) => {
    const a = run(f, 5, 8, { rowPairs: reversed() });
    for (const sheetCopies of [undefined, new Map<string, string>()]) {
      expect(new FormulaRemapper(spec({ rowPairs: reversed(), sheetCopies })).remap(f, 5, 8)).toEqual(a);
    }
  });
  it('the failure reported is the one of the merged-table rules (round 8), also when the copies cannot help either', () => {
    expect(whyC('INDIRECT("C5")')).toBe('positional');
    expect(whyC('[1]Sheet1!A1')).toBe('unsafe');
    expect(whyC('SUM(レート:Master!B2)')).toBe('otherSheet');
    expect(whyC('C5+#REF!')).toBe('unsupported');
    expect(whyC('C5#')).toBe('array');
    expect(whyC('ROW()')).toBe('positional');
  });
});

describe('copies: references to OTHER sheets of the source read their copies', () => {
  it.each<[string, string]>([
    ['C5*レート!$B$1', 'E5*src_レート!$B$1'],
    ["C5*'レート'!$B$1", 'E5*src_レート!$B$1'],
    ['レート!B1', 'src_レート!B1'],
    ['レート!$B$1+レート!$B$2', 'src_レート!$B$1+src_レート!$B$2'],
    ['VLOOKUP(C5,レート!$A$2:$B$10,2,FALSE)', 'VLOOKUP(E5,src_レート!$A$2:$B$10,2,FALSE)'],
    ['VLOOKUP(C5,レート!$A:$B,2,TRUE)', 'VLOOKUP(E5,src_レート!$A:$B,2,TRUE)'],
    ['INDEX(Master!$B:$B,MATCH(C5,Master!$A:$A,0))', 'INDEX(src_Master!$B:$B,MATCH(E5,src_Master!$A:$A,0))'],
    ['SUMIF(Master!A:A,"x",Master!B:B)', 'SUMIF(src_Master!A:A,"x",src_Master!B:B)'],
    ['COUNTIF(Master!$A$2:$A$100,C5)', 'COUNTIF(src_Master!$A$2:$A$100,E5)'],
    ['SUM(レート!B2:B10)*C5', 'SUM(src_レート!B2:B10)*E5'],
    ["'My Sheet'!B1+C5", "'src_My Sheet'!B1+E5"],
    ["'O''Brien'!A1", "'src_O''Brien'!A1"],
    ['Sheet2!A1', "'src_Sheet2(1)'!A1"],
    ['MASTER!A1', 'src_Master!A1'],
    ['master!$A$1', 'src_Master!$A$1'],
    ['IFERROR(VLOOKUP(C5,Master!A:C,3,0),"")', 'IFERROR(VLOOKUP(E5,src_Master!A:C,3,0),"")'],
    ['IF(C5>レート!$B$1,"高","低")', 'IF(E5>src_レート!$B$1,"高","低")'],
    ['COUNTIFS(Master!A:A,"a",Master!B:B,">1")', 'COUNTIFS(src_Master!A:A,"a",src_Master!B:B,">1")'],
    ['SUMIFS(Master!C:C,Master!A:A,C5)', 'SUMIFS(src_Master!C:C,src_Master!A:A,E5)'],
    ['_xlfn.XLOOKUP(C5,Master!A:A,Master!B:B)', '_xlfn.XLOOKUP(E5,src_Master!A:A,src_Master!B:B)'],
    ['LOOKUP(C5,レート!A2:A9,レート!B2:B9)', 'LOOKUP(E5,src_レート!A2:A9,src_レート!B2:B9)'],
    ['RANK(C5,Master!B2:B9)', 'RANK(E5,src_Master!B2:B9)'],
    ['LARGE(Master!B2:B9,2)', 'LARGE(src_Master!B2:B9,2)'],
    ['_xlfn.TEXTJOIN(",",TRUE,Master!A2:A5)', '_xlfn.TEXTJOIN(",",TRUE,src_Master!A2:A5)'],
    ['MAX(Master!B2:B9)-C5', 'MAX(src_Master!B2:B9)-E5'],
    ['AND(C5>1,Master!B1)', 'AND(E5>1,src_Master!B1)'],
    ['SUMPRODUCT(--(レート!A2:A10>5),レート!B2:B10)', 'SUMPRODUCT(--(src_レート!A2:A10>5),src_レート!B2:B10)'],
    ['LOOKUP(2,1/(レート!A2:A10=C5),レート!B2:B10)', 'LOOKUP(2,1/(src_レート!A2:A10=E5),src_レート!B2:B10)'],
    ['INDEX(レート!A2:A10,3)', 'INDEX(src_レート!A2:A10,3)'],
    ['MATCH(C5,レート!A2:A10,0)', 'MATCH(E5,src_レート!A2:A10,0)'],
    ['SUBTOTAL(9,レート!A2:A10)', 'SUBTOTAL(9,src_レート!A2:A10)'],
    ['_xlfn.AGGREGATE(9,6,レート!A2:A10)', '_xlfn.AGGREGATE(9,6,src_レート!A2:A10)'],
    ['ROWS(レート!A2:A10)', 'ROWS(src_レート!A2:A10)'],
    ['COLUMNS(レート!A2:C2)', 'COLUMNS(src_レート!A2:C2)'],
    ['SUM(レート!2:2)', 'SUM(src_レート!2:2)'],
    ['SUM(レート!B:B)', 'SUM(src_レート!B:B)'],
    ['"レート!A1"&C5', '"レート!A1"&E5'],
  ])('%s -> %s', (f, want) => {
    expect(okC(f)).toBe(want);
  });

  it('a reference to another sheet always reads its copy, in every row and for every merged row', () => {
    expect(okC('レート!$B$1', 5, 8, { rowPairs: reversed() })).toBe('src_レート!$B$1');
    expect(okC('C5*レート!$B$1', 5, 8, { rowPairs: reversed() })).toBe('E8*src_レート!$B$1');
    expect(okC('レート!B5', 5, 12)).toBe('src_レート!B5');
  });

  it('a sheet name is found whatever its case; a sheet that was not copied (or a 3-D / external reference) stays a value', () => {
    expect(whyC('Unknown!A1')).toBe('otherSheet');
    expect(whyC('Unknown!A1+C5')).toBe('otherSheet');
    expect(whyC('SUM(レート:Master!B2)')).toBe('otherSheet');
    expect(whyC('SUM(Sheet1:Sheet3!A1)')).toBe('otherSheet');
    expect(whyC('[1]Sheet1!A1')).toBe('unsafe');
    expect(whyC("'[Book1.xlsx]Sheet1'!A1")).toBe('unsafe');
    expect(whyC("'C:\\x\\[B.xlsx]S'!A1")).toBe('unsafe');
    expect(whyC('[1]!MyName')).toBe('unsafe');
  });

  it('only the sheets that are in the output can be read: a map without the sheet changes nothing', () => {
    const only = new Map([['master', 'src_Master']]);
    const r = (f: string) => new FormulaRemapper(spec({ sheetCopies: only })).remap(f, 5, 5);
    expect(r('Master!A1')).toMatchObject({ ok: true, formula: 'src_Master!A1' });
    expect(r('レート!A1')).toMatchObject({ ok: false, reason: 'otherSheet' });
    // the data sheet itself is not in the map: its references that the merged table cannot keep stay values
    expect(r('C5*$H$1')).toMatchObject({ ok: false, reason: 'externalCell' });
    expect(r('C5+Master!B1')).toMatchObject({ ok: true, formula: 'E5+src_Master!B1' });
  });
});

describe('copies: references to the DATA sheet that the merged table cannot represent read its copy at the original coordinates', () => {
  it.each<[string, string]>([
    // non-imported columns
    ['E5*2', 'src_評価!E5*2'],
    ['C5+E5', 'E5+src_評価!E5'],
    ['E5+F5', 'src_評価!E5+src_評価!F5'],
    ['$E$5', 'src_評価!$E$5'],
    ['E$5', 'src_評価!E$5'],
    ['$E5', 'src_評価!$E5'],
    ['E4', 'src_評価!E4'],
    ['E6-E5', 'src_評価!E6-src_評価!E5'],
    ['SUM(E5:F5)', 'SUM(src_評価!E5:F5)'],
    ['SUM(C5:F5)', 'SUM(src_評価!C5:F5)'],
    ['A5&B5', 'src_評価!A5&src_評価!B5'],
    ['SUM(E2:E11)', 'SUM(src_評価!E2:E11)'],
    ['SUM(E:E)', 'SUM(src_評価!E:E)'],
    ['COUNTIF(E:E,"x")', 'COUNTIF(src_評価!E:E,"x")'],
    ['AVERAGE(E$2:E5)', 'AVERAGE(src_評価!E$2:E5)'],
    ["'評価'!E5*2", 'src_評価!E5*2'],
    ['評価!E5*2', 'src_評価!E5*2'],
    ['IF(E5>0,C5/E5,0)', 'IF(src_評価!E5>0,E5/src_評価!E5,0)'],
    // the merged table first: C and D are imported
    ['C5*D5+E5', 'E5*F5+src_評価!E5'],
    ['SUM(C2:C11)*E5', 'SUM(E2:E11)*src_評価!E5'],
    ['E5/SUM(C2:C11)', 'src_評価!E5/SUM(E2:E11)'],
  ])('PARTIAL: %s -> %s', (f, want) => {
    expect(okC(f, 5, 5, PARTIAL)).toBe(want);
  });

  it.each<[string, string]>([
    // cells outside the table
    ['C5*$H$1', 'E5*src_評価!$H$1'],
    ['C5/C14', 'E5/src_評価!C14'],
    ['C5*$H5', 'E5*src_評価!$H5'],
    ['$P$1', 'src_評価!$P$1'],
    ['C$13', 'src_評価!C$13'],
    ['C5/C$34', 'E5/src_評価!C$34'],
    ['SUM(5:5)', 'SUM(src_評価!5:5)'],
    ['MATCH("x",$1:$1,0)', 'MATCH("x",src_評価!$1:$1,0)'],
    ['C5-$H$1*$H$2', 'E5-src_評価!$H$1*src_評価!$H$2'],
    ['IF(C5>$H$1,"高","低")', 'IF(E5>src_評価!$H$1,"高","低")'],
    ['ROUND(C5*$H$1,0)', 'ROUND(E5*src_評価!$H$1,0)'],
  ])('externalCell: %s -> %s', (f, want) => {
    expect(okC(f)).toBe(want);
  });

  it('cells above the table (the header row is 3) are outside it', () => {
    const over = { region: { headerRow: 3, firstDataRow: 4, lastDataRow: 11, firstCol: 1, lastCol: 6 }, target: { headerRow: 2, firstDataRow: 3, lastDataRow: 10 }, rowPairs: [4, 5, 6, 7, 8, 9, 10, 11].map((r): [number, number] => [r, r - 1]) };
    expect(okC('C1', 5, 4, over)).toBe('src_評価!C1');
    expect(okC('SUM(C2:E2)', 5, 4, over)).toBe('SUM(src_評価!C2:E2)');
    expect(okC('C$3', 5, 4, over)).toBe('E$2');
    expect(whyC('C2:E2', 5, 4, over)).toBe('externalCell');
  });

  it('whole-column references that also cover cells above / below the table read the copy (the footer is part of that column in the source)', () => {
    const over = { hasContentOutside: (col: number, _r1: number, r2: number) => col === 3 && r2 >= 14 };
    expect(whyC('SUM(C:C)', 5, 5, over)).toBe('ok');
    expect(okC('SUM(C:C)', 5, 5, over)).toBe('SUM(src_評価!C:C)');
    expect(okC('SUM(D:D)', 5, 5, over)).toBe('SUM(F:F)'); // column D has nothing outside the table: the merged column
    expect(okC('SUM(C2:C20)', 5, 5, over)).toBe('SUM(src_評価!C2:C20)');
  });

  it.each<[string, string, Array<[number, number]>, number, number]>([
    ['C5-C4', 'E5-src_評価!C4', without(4), 5, 5],
    ['C6-C7', 'E6-src_評価!C7', without(7), 6, 6],
    ['SUM(C$2:C5)', 'SUM(src_評価!C$2:C5)', reversed(), 5, 8],
    ['SUM(C3:C6)', 'SUM(src_評価!C3:C6)', without(4), 5, 5],
    ['COUNTA($C$2:C5)', 'COUNTA(src_評価!$C$2:C5)', reversed(), 5, 8],
    ['C5/C2', 'E5/src_評価!C2', without(2), 5, 5],
    ['(C5+C6)/2', '(E5+src_評価!C6)/2', without(6), 5, 5],
    ['C4-C5', 'src_評価!C4-E8', without(4), 5, 8],
  ])('rowRelation: %s -> %s', (f, want, rowPairs, sr, tr) => {
    expect(okC(f, sr, tr, { rowPairs })).toBe(want);
  });

  it('rows that the merged table CAN follow stay in the merged table (round 8), also when the formula sits on another merged row', () => {
    expect(okC('C5-C6', 5, 8, { rowPairs: identity().map(([r]): [number, number] => [r, r + 3]) })).toBe('E8-E9');
    expect(okC('C4', 5, 8, { rowPairs: reversed() })).toBe('E9');
    expect(viaC('C4', 5, 8, { rowPairs: reversed() })).toBe(false);
    expect(okC('SUM(C$2:C$11)', 5, 8, { rowPairs: reversed() })).toBe('SUM(E$2:E$11)');
  });

  it('a source row placed twice has no unique merged row: reading it falls back to the copy, the own row never does', () => {
    const dup: Array<[number, number]> = [...identity(), [4, 12]];
    expect(okC('C4', 5, 5, { rowPairs: dup })).toBe('src_評価!C4');
    expect(okC('C4+D4', 4, 12, { rowPairs: dup })).toBe('E12+F12');
    expect(viaC('C4+D4', 4, 12, { rowPairs: dup })).toBe(false);
  });

  it('rows of the source that matched nothing: a formula that reads them reads the copy (which keeps the people that did not match)', () => {
    expect(okC('C7*2', 5, 5, { rowPairs: without(7) })).toBe('src_評価!C7*2');
    expect(okC('SUM(C2:C11)', 5, 5, { rowPairs: without(7) })).toBe('SUM(E2:E11)'); // the whole table is the merged table (round 8, widened)
    expect(runC('SUM(C2:C11)', 5, 5, { rowPairs: without(7) })).toMatchObject({ widened: true, exact: false });
  });

  it('key columns the base also has are read from the base (merged table first); without that they read the copy', () => {
    expect(okC('A5&B5', 5, 5, KEYS)).toBe('A5&B5');
    expect(viaC('A5&B5', 5, 5, KEYS)).toBe(false);
    expect(runC('A5&B5', 5, 5, KEYS)).toMatchObject({ substituted: true });
    expect(okC('A5&B5')).toBe('src_評価!A5&src_評価!B5');
    expect(okC('VLOOKUP(A5,Master!A:C,3,0)', 5, 5, KEYS)).toBe('VLOOKUP(A5,src_Master!A:C,3,0)');
    expect(okC('VLOOKUP(A5,Master!A:C,3,0)')).toBe('VLOOKUP(src_評価!A5,src_Master!A:C,3,0)');
    expect(okC('E5&A5', 5, 5, { ...PARTIAL, ...KEYS })).toBe('src_評価!E5&A5');
  });

  it('the formula\'s own row: the copy keeps the source row of the person whatever the merged row is', () => {
    expect(okC('E5*2', 5, 9, { ...PARTIAL, rowPairs: reversed() })).toBe('src_評価!E5*2');
    expect(okC('C5*E5', 5, 9, { ...PARTIAL, rowPairs: reversed() })).toBe('E9*src_評価!E5');
    expect(okC('E5*2', 5, 40, PARTIAL)).toBe('src_評価!E5*2');
  });
});

describe('copies: defined names', () => {
  it.each<[string, Array<[string, string | null, string]>, string]>([
    ['C5*Rate', [['Rate', null, '評価!$H$1']], 'E5*src_評価!$H$1'],
    ['C5*Rate', [['Rate', null, "'評価'!$H$1"]], 'E5*src_評価!$H$1'],
    ['C5*TaxRate', [['TaxRate', null, 'レート!$B$1']], 'E5*src_レート!$B$1'],
    ['C5*Tax', [['Tax', null, 'レート!$B$1*1.1']], 'E5*(src_レート!$B$1*1.1)'],
    ['C5>Lim', [['Lim', null, 'IF(レート!$B$1>0,レート!$B$1,5)']], 'E5>(IF(src_レート!$B$1>0,src_レート!$B$1,5))'],
    ['SUM(Scores)', [['Scores', null, '評価!$C$2:$C$11']], 'SUM($E$2:$E$11)'],
    ['C5*Other', [['Other', '評価', 'Master!$A$1']], 'E5*src_Master!$A$1'],
    ['C5*K', [['K', null, '5']], 'E5*5'],
    ['C5*Pct', [['Pct', null, '8%']], 'E5*(8%)'],
    ['SUM(Master_B)', [['Master_B', null, 'Master!$B$2:$B$9']], 'SUM(src_Master!$B$2:$B$9)'],
    ['COUNTIF(Names,C5)', [['Names', null, 'Master!$A:$A']], 'COUNTIF(src_Master!$A:$A,E5)'],
    ['C5*(1+Tax)', [['Tax', null, 'レート!$B$1*0.1']], 'E5*(1+(src_レート!$B$1*0.1))'],
    ['Lo+Hi', [['Lo', null, 'レート!$B$1'], ['Hi', null, 'レート!$B$2']], 'src_レート!$B$1+src_レート!$B$2'],
  ])('%s %j -> %s', (f, list, want) => {
    expect(okC(f, 5, 5, names(list))).toBe(want);
  });

  it('a name of the data sheet that the merged table cannot read reads the copy (priority: merged table first)', () => {
    expect(okC('SUM(Extra)', 5, 5, { ...PARTIAL, ...names([['Extra', null, '評価!$E$2:$E$11']]) })).toBe('SUM(src_評価!$E$2:$E$11)');
    expect(okC('Cell*2', 5, 5, { ...PARTIAL, ...names([['Cell', null, '評価!$E$5']]) })).toBe('src_評価!$E$5*2');
    expect(okC('C5*Cell', 5, 5, names([['Cell', null, '評価!$D$2']]))).toBe('E5*$F$2'); // D2 is a placed row of an imported column
  });

  it.each<[string, Array<[string, string | null, string]>, string]>([
    ['C5+Rel', [['Rel', null, '評価!A1']], 'unresolvedName'],
    ['C5+Fn', [['Fn', null, 'OFFSET(評価!$C$1,1,0)']], 'unresolvedName'],
    ['C5+Bad', [['Bad', null, 'レート!$B$1:$B$3*2']], 'unresolvedName'],
    ['C5+Other', [['Other', null, 'Unknown!$A$1']], 'otherSheet'],
    ['C5+Three', [['Three', null, 'レート:Master!$A$1']], 'otherSheet'],
    ['C5+Ext', [['Ext', null, '[1]Sheet1!$A$1']], 'unsafe'],
    ['C5+Nope', [], 'unresolvedName'],
    ['C5+Both', [['Both', null, 'レート!$B$1+Both']], 'unresolvedName'],
    ['C5+Ind', [['Ind', null, 'INDIRECT("A1")']], 'unresolvedName'],
    ['C5+Rg', [['Rg', null, 'レート!$A$2:$A$9']], 'otherSheet'], // a bare range operand is intersected implicitly
  ])('stays a value: %s %j -> %s', (f, list, reason) => {
    expect(whyC(f, 5, 5, names(list))).toBe(reason);
  });
});

describe('copies: structured references of the source', () => {
  const T = (over: Record<string, unknown> = {}) => ({
    tables: [
      { name: 'Tbl', sheet: '評価', ref: 'A1:F11', headerRow: true, totalsRow: false, columns: ['学生ID', '氏名', '点1', '点2', '点3', '点4'] },
      { name: 'Tbl2', sheet: '評価', ref: 'A1:F12', headerRow: true, totalsRow: true, columns: ['学生ID', '氏名', '点1', '点2', '点3', '点4'] },
      { name: 'Tbl3', sheet: 'レート', ref: 'A1:B5', headerRow: true, totalsRow: false, columns: ['点', '評価'] },
      { name: 'Tbl4', sheet: 'Zzz', ref: 'A1:B5', headerRow: true, totalsRow: false, columns: ['x', 'y'] },
    ],
    ...over,
  });
  it.each<[string, string]>([
    ['Tbl[@点3]*2', 'src_評価!E5*2'],
    ['Tbl[[#This Row],[点3]]*2', 'src_評価!E5*2'],
    ['SUM(Tbl[点3])', 'SUM(src_評価!$E$2:$E$11)'],
    ['SUM(Tbl[点1])', 'SUM($E$2:$E$11)'],
    ['SUM(Tbl[[点1]:[点3]])', 'SUM(src_評価!$C$2:$E$11)'],
    ['SUM(Tbl[[#Data],[点3]])', 'SUM(src_評価!$E$2:$E$11)'],
    ['ROWS(Tbl[#Data])', 'ROWS(src_評価!$A$2:$F$11)'],
    ['Tbl[[#Headers],[点3]]', 'src_評価!$E$1'],
    ['Tbl[@点1]+Tbl[@点3]', 'E5+src_評価!E5'],
    ['Tbl2[[#Totals],[点1]]', 'src_評価!$C$12'],
    ['SUM(Tbl2[[#All],[点3]])', 'SUM(src_評価!$E$1:$E$12)'],
    ['SUM(Tbl3[評価])', 'SUM(src_レート!$B$2:$B$5)'],
    ['VLOOKUP(C5,Tbl3[#Data],2,FALSE)', 'VLOOKUP(E5,src_レート!$A$2:$B$5,2,FALSE)'],
    ['INDEX(Tbl3[評価],2)', 'INDEX(src_レート!$B$2:$B$5,2)'],
  ])('PARTIAL: %s -> %s', (f, want) => {
    expect(okC(f, 5, 5, { ...PARTIAL, ...T() })).toBe(want);
  });

  it('an unqualified [@col] (XlsxWriter) is the table of the formula\'s own cell: it needs the source column of the formula', () => {
    const remapper = new FormulaRemapper(spec({ ...PARTIAL, ...T(), sheetCopies: COPIES }));
    expect(remapper.remap('[@点3]*2', 5, 5, 5, false)).toMatchObject({ ok: true, formula: 'src_評価!E5*2', viaCopy: true });
    expect(remapper.remap('[@点3]*2', 5, 5, undefined, false)).toMatchObject({ ok: false, reason: 'unresolvedTable' });
    // a column outside the table the formula is in
    expect(remapper.remap('[@点3]*2', 5, 5, 12, false)).toMatchObject({ ok: false, reason: 'unresolvedTable' });
  });

  it.each<[string, string]>([
    ['Tbl3[@評価]', 'otherSheet'], // [@col] is the row of the table the formula is in (round 8 reports the other sheet first)
    ['Nope[x]', 'unresolvedTable'],
    ['Tbl4[x]', 'otherSheet'],
    ['SUM(Tbl4[x])', 'otherSheet'],
    ['Tbl[点9]', 'unresolvedTable'],
    ['Tbl[[#Totals],[点1]]', 'unresolvedTable'],
    ['Tbl3[評価]', 'otherSheet'], // a whole column as a bare operand is intersected implicitly
  ])('stays a value: %s -> %s', (f, reason) => {
    expect(whyC(f, 5, 5, T())).toBe(reason);
  });

  it('the totals row of a table whose merged table cannot represent it', () => {
    // without copies: unresolvedTable (round 8); with the copy of the sheet the cell is readable
    expect(why('Tbl2[[#Totals],[点1]]', 5, 5, T())).toBe('unresolvedTable');
    expect(okC('C5/Tbl2[[#Totals],[点1]]', 5, 5, T())).toBe('E5/src_評価!$C$12');
  });
});

describe('copies: functions that depend on where a reference is', () => {
  it.each<[string, string]>([
    ['ROW(C5)', 'ROW(src_評価!C5)'],
    ['ROW(C5)-1', 'ROW(src_評価!C5)-1'],
    ['COLUMN(C5)', 'COLUMN(src_評価!C5)'],
    ['COLUMN(E5)', 'COLUMN(src_評価!E5)'],
    ['ROW(C5:C9)', 'ROW(src_評価!C5:C9)'],
    ['SUM(ROW(C5:C9))', 'SUM(ROW(src_評価!C5:C9))'],
    ['OFFSET(C5,0,1)', 'OFFSET(src_評価!C5,0,1)'],
    ['OFFSET(C5,1,0)+C5', 'OFFSET(src_評価!C5,1,0)+src_評価!C5'],
    ['SUM(OFFSET(C5,0,0,1,3))', 'SUM(OFFSET(src_評価!C5,0,0,1,3))'],
    ['OFFSET(レート!A1,2,1)', 'OFFSET(src_レート!A1,2,1)'],
    ['ROW(INDEX(C2:C11,3))', 'ROW(INDEX(src_評価!C2:C11,3))'],
    ['SUMPRODUCT((ROW(C2:C11)>3)*C2:C11)', 'SUMPRODUCT((ROW(src_評価!C2:C11)>3)*src_評価!C2:C11)'],
    ['C5*ROW(C5)', 'src_評価!C5*ROW(src_評価!C5)'],
  ])('every reference reads the copies: %s -> %s', (f, want) => {
    expect(okC(f)).toBe(want);
    expect(viaC(f)).toBe(true);
    expect((runC(f) as { exact: boolean }).exact).toBe(true);
  });

  it.each([
    'ROW()', 'ROW()-1', 'COLUMN()', 'INDEX(C2:C11,ROW()-1)', 'INDEX(C2:C11,ROW(C5)-ROW())', 'INDIRECT("C5")', 'INDIRECT("C"&ROW(C5))', 'ADDRESS(1,1)', 'CELL("address",C5)',
    'CELL("row")', 'FORMULATEXT(C5)', 'SHEET()', 'SHEETS()', 'INFO("osversion")', 'ISFORMULA(C5)', 'GETPIVOTDATA("x",A1)', 'OFFSET()', 'OFFSET(,1,1)', 'ROW(,)',
    'HYPERLINK("#\'評価\'!C"&ROW(),"移動")',
  ])('stays a value (positional): %s', (f) => {
    expect(whyC(f)).toBe('positional');
  });
});

describe('copies: reference operators and rows that are not aligned', () => {
  it.each<[string, string]>([
    ['SUM(C5:INDEX(E2:E11,3))', 'SUM(src_評価!C5:INDEX(src_評価!E2:E11,3))'],
    ['SUM(C2:C11 E5:F5)', 'SUM(src_評価!C2:C11 src_評価!E5:F5)'],
    ['SUM((C2:C3,E5:E6))', 'SUM((src_評価!C2:C3,src_評価!E5:E6))'],
    ['SUM(C5:E5 D5:F5)', 'SUM(src_評価!C5:E5 src_評価!D5:F5)'],
    ['SUM(INDEX(E2:E11,2):INDEX(E2:E11,4))', 'SUM(INDEX(src_評価!E2:E11,2):INDEX(src_評価!E2:E11,4))'],
  ])('PARTIAL, one sheet for every operand of a reference operator: %s -> %s', (f, want) => {
    expect(okC(f, 5, 5, PARTIAL)).toBe(want);
  });

  it.each<[string, string, boolean]>([
    // [formula, result, every reference reads the copies]
    ['SUMIF(E2:E11,">5",C2:C11)', 'SUMIF(src_評価!E2:E11,">5",src_評価!C2:C11)', true],
    ['SUMPRODUCT(C2:C11,E2:E11)', 'SUMPRODUCT(src_評価!C2:C11,src_評価!E2:E11)', true],
    ['SUMPRODUCT(C2:C11,D2:D11)', 'SUMPRODUCT(E2:E11,F2:F11)', false],
    ['INDEX(C2:C11,MATCH(E5,E2:E11,0))', 'INDEX(src_評価!C2:C11,MATCH(src_評価!E5,src_評価!E2:E11,0))', true],
    ['INDEX(C2:C11,MATCH(D5,D2:D11,0))', 'INDEX(E2:E11,MATCH(F5,F2:F11,0))', false],
    ['SUM(E2:E11)/SUM(C2:C11)', 'SUM(src_評価!E2:E11)/SUM(src_評価!C2:C11)', true],
    ['SUMPRODUCT(C2:C11,レート!B2:B11)', 'SUMPRODUCT(src_評価!C2:C11,src_レート!B2:B11)', true],
    ['SUM(C2:C11)/SUM(レート!B2:B11)', 'SUM(src_評価!C2:C11)/SUM(src_レート!B2:B11)', true],
    ['SUM(C2:C11)*レート!B1', 'SUM(E2:E11)*src_レート!B1', false],
    ['SUM(C2:C11)*E5', 'SUM(E2:E11)*src_評価!E5', false],
    ['SUMIF(C2:C11,">"&E5,D2:D11)', 'SUMIF(E2:E11,">"&src_評価!E5,F2:F11)', false],
    ['SUMPRODUCT((E2:E11>5)*(C2:C11))', 'SUMPRODUCT((src_評価!E2:E11>5)*(src_評価!C2:C11))', true],
    ['SUMPRODUCT((C2:C11>5)*(D2:D11))', 'SUMPRODUCT((E2:E11>5)*(F2:F11))', false],
    ['COUNTIFS(E2:E11,">1",C2:C11,">1")', 'COUNTIFS(src_評価!E2:E11,">1",src_評価!C2:C11,">1")', true],
    ['SUMIFS(C2:C11,E2:E11,">1")', 'SUMIFS(src_評価!C2:C11,src_評価!E2:E11,">1")', true],
    ['_xlfn.XLOOKUP(E5,E2:E11,C2:C11)', '_xlfn.XLOOKUP(src_評価!E5,src_評価!E2:E11,src_評価!C2:C11)', true],
    ['LOOKUP(E5,E2:E11,C2:C11)', 'LOOKUP(src_評価!E5,src_評価!E2:E11,src_評価!C2:C11)', true],
    ['CORREL(C2:C11,E2:E11)', 'CORREL(src_評価!C2:C11,src_評価!E2:E11)', true],
    ['RANK(C5,E2:E11)', 'RANK(E5,src_評価!E2:E11)', false],
  ])('PARTIAL, a range of the merged table and a range of a copy: %s', (f, want, allCopies) => {
    expect(okC(f, 5, 5, PARTIAL)).toBe(want);
    // "all copies" = no unprefixed reference is left
    const left = /(^|[^!A-Za-z$])\$?[A-F]\$?\d/.test(want.replace(/src_[^!]*![$A-Z0-9:]+/g, '').replace(/"[^"]*"/g, ''));
    expect(!left).toBe(allCopies);
  });

  it.each<[string, string]>([
    // INDEX(range of the merged table, position read from the copy): the position counts rows of the SOURCE layout
    ['INDEX(C2:C11,E5)', 'INDEX(src_評価!C2:C11,src_評価!E5)'],
    ['INDEX(C2:C11,ABS(E5)+1)', 'INDEX(src_評価!C2:C11,ABS(src_評価!E5)+1)'],
    ['INDEX(C2:D11,E5,2)', 'INDEX(src_評価!C2:D11,src_評価!E5,2)'],
    ['_xlfn.TAKE(C2:C11,E5)', '_xlfn.TAKE(src_評価!C2:C11,src_評価!E5)'],
    // the position read from the merged table, or a range of the copy indexed by a number: nothing to align
    ['INDEX(C2:C11,D5)', 'INDEX(E2:E11,F5)'],
    ['INDEX(E2:E11,C5)', 'INDEX(src_評価!E2:E11,E5)'],
    ['INDEX(E2:E11,3)', 'INDEX(src_評価!E2:E11,3)'],
    // a value (not a position) from the copy is fine
    ['SMALL(C2:C11,E5)', 'SMALL(E2:E11,src_評価!E5)'],
    ['LARGE(C2:C11,E5)+INDEX(C2:C11,1)', 'LARGE(E2:E11,src_評価!E5)+INDEX(E2:E11,1)'],
    ['COUNTIF(C2:C11,E5)', 'COUNTIF(E2:E11,src_評価!E5)'],
    ['VLOOKUP(E5,C2:D11,2,FALSE)', 'VLOOKUP(src_評価!E5,E2:F11,2,FALSE)'],
  ])('PARTIAL, positions: %s -> %s', (f, want) => {
    expect(okC(f, 5, 5, PARTIAL)).toBe(want);
  });

  it('the rows of the merged table and the rows of the copy are NOT aligned (rows reversed): that is why one formula never mixes two ranges', () => {
    const over = { ...PARTIAL, rowPairs: reversed() };
    expect(okC('SUMIF(E2:E11,">5",C2:C11)', 5, 8, over)).toBe('SUMIF(src_評価!E2:E11,">5",src_評価!C2:C11)');
    expect(okC('SUM(C2:C11)*E5', 5, 8, over)).toBe('SUM(E2:E11)*src_評価!E5');
  });
});

describe('copies: ranges on another sheet are not intersected implicitly (Excel evaluates them for the formula\'s own row)', () => {
  it.each([
    'レート!A2:A10', 'レート!A2:A10*2', 'ABS(レート!A2:A10)', 'IF(C5>0,レート!A2:A10,0)', 'SUM(ABS(レート!A2:A10))', '(レート!A2:A10)',
    'VLOOKUP(レート!A2:A10,Master!A:B,2,0)', 'ROUND(レート!A2:A10,1)', 'LEN(レート!A2:A3)', 'C5*レート!B:B', 'レート!2:2', 'IF(レート!A2:C2>1,1,0)', 'SUMIF(レート!A2:A9,">1")*レート!B2:B9',
    '-レート!A2:A3', 'INDEX(レート!A2:A10,3)+レート!A2:A10', 'SUM(1,レート!A2:A10)+レート!A2:A10', 'MATCH(レート!A2:A3,レート!A2:A10,0)',
  ])('a bare / scalar-position range stays a value: %s', (f) => {
    expect(whyC(f)).toBe('otherSheet');
  });

  it.each<[string, string]>([
    ['レート!A2:A10*2', 'src_レート!A2:A10*2'],
    ['SUM(レート!A2:A10*2)', 'SUM(src_レート!A2:A10*2)'],
    ['レート!A2:A10', 'src_レート!A2:A10'],
    ['IF(レート!A2:A10>1,1,0)', 'IF(src_レート!A2:A10>1,1,0)'],
  ])('... but an array formula has no implicit intersection: %s -> %s', (f, want) => {
    expect(whyC(f)).toBe('otherSheet');
    expect(okC(f, 5, 5, {}, true)).toBe(want);
  });

  it('a single cell is never intersected, nor is A1:A1', () => {
    expect(okC('レート!A2*2')).toBe('src_レート!A2*2');
    expect(okC('レート!A2:A2*2')).toBe('src_レート!A2:A2*2');
    expect(okC('ABS(レート!$B$1)')).toBe('ABS(src_レート!$B$1)');
  });

  it('the same rule for the data sheet\'s own fallback: a bare half-imported range is a value, a range inside SUM is kept', () => {
    expect(whyC('E2:E11*2', 5, 5, PARTIAL)).toBe('notImported');
    expect(whyC('ABS(E2:E11)', 5, 5, PARTIAL)).toBe('notImported');
    expect(okC('E2:E11*2', 5, 5, PARTIAL, true)).toBe('src_評価!E2:E11*2');
    expect(okC('SUM(E2:E11)', 5, 5, PARTIAL)).toBe('SUM(src_評価!E2:E11)');
    expect(okC('SUMPRODUCT(--(E2:E11>1))', 5, 5, PARTIAL)).toBe('SUMPRODUCT(--(src_評価!E2:E11>1))');
  });
});

describe('copies: the cached result, widened ranges and what the formula reports', () => {
  it('reading only copies: the very same inputs, the cached result of the source stays valid', () => {
    for (const f of ['C5*レート!$B$1', 'SUMIF(E2:E11,">5",C2:C11)', 'C5/C14', 'SUM(C$2:C5)']) {
      const r = runC(f, 5, 5, f === 'SUM(C$2:C5)' ? { rowPairs: reversed() } : f.includes('E2') ? PARTIAL : {});
      expect(r).toMatchObject({ ok: true, viaCopy: true, exact: true, widened: false });
    }
  });

  it('copySheets: the copies read, once each, in order of first use', () => {
    expect(runC('レート!A1+Master!A1+レート!B1')).toMatchObject({ viaCopy: true, copySheets: ['src_レート', 'src_Master'] });
    expect(runC('C5+E5', 5, 5, PARTIAL)).toMatchObject({ viaCopy: true, copySheets: ['src_評価'] });
    expect(runC('C5+D5')).not.toHaveProperty('viaCopy');
    expect(runC('C5+D5')).not.toHaveProperty('copySheets');
  });

  it('a hybrid formula keeps the flags of its merged part (a whole-table range over rows that were not placed is widened and not exact)', () => {
    const r = runC('SUM(C2:C11)*レート!$B$1', 5, 5, { rowPairs: without(7) });
    expect(r).toMatchObject({ ok: true, formula: 'SUM(E2:E11)*src_レート!$B$1', viaCopy: true, widened: true, exact: false });
    // but when the formula falls back to copy mode nothing of the merged table is read: exact again
    const c = runC('SUM(C2:C11)/SUM(レート!B2:B11)', 5, 5, { rowPairs: without(7) });
    expect(c).toMatchObject({ ok: true, formula: 'SUM(src_評価!C2:C11)/SUM(src_レート!B2:B11)', viaCopy: true, widened: false, exact: true });
  });

  it('a renamed header read through the merged table is inexact; through the copy it is exact', () => {
    const over = { headerSame: () => false, ...PARTIAL };
    expect(runC('C$1', 5, 5, over)).toMatchObject({ ok: true, formula: 'E$1', exact: false });
    expect(runC('E$1', 5, 5, over)).toMatchObject({ ok: true, formula: 'src_評価!E$1', exact: true, viaCopy: true });
  });
});

describe('copies: sheet names the planner makes (Japanese, spaces, apostrophes, (1) suffix, 31 characters)', () => {
  // [sheet of the source, name of its copy, the reference as written in the source, the reference to the copy]
  const LONG = 'abcdefghij_abcdefghij_abcdefghi';
  const CASES: Array<[string, string, string, string]> = [
    ['評価', 'a_評価', '評価!B2', 'a_評価!B2'],
    ['面接 評価', '02_面接評価_面接 評価', "'面接 評価'!B2", "'02_面接評価_面接 評価'!B2"],
    ['長い', LONG, '長い!A1', `${LONG}!A1`],
    ["it's", "x_it's", "'it''s'!A1", "'x_it''s'!A1"],
    ['sheet', 'a_Sheet(1)', 'Sheet!A1', "'a_Sheet(1)'!A1"],
    ['a-b', 'a_a-b', "'a-b'!A1", "'a_a-b'!A1"],
    ['1月', '02_1月', "'1月'!A1", "'02_1月'!A1"],
    ['tax2025', 'q_TAX2025', "'TAX2025'!A1", 'q_TAX2025!A1'],
    ['a.b', 'q_a.b', "'a.b'!A1", "'q_a.b'!A1"],
    ['ＡＢ'.toLowerCase(), 'q_ＡＢ', "'ＡＢ'!A1", "'q_ＡＢ'!A1"],
    ['売上（税込）', '03_売上（税込）', "'売上（税込）'!C3", "'03_売上（税込）'!C3"],
    ['x', 'あ_x', 'x!A1:B2', 'あ_x!A1:B2'],
  ];
  const copies = new Map(CASES.map(([sheet, copy]) => [sheet.toLowerCase(), copy] as [string, string]));
  it.each(CASES)('%s: %s -> %s', (_sheet, copy, f, want) => {
    const r = new FormulaRemapper(spec({ sheetCopies: copies })).remap(`SUM(${f})`, 5, 5);
    expect(r).toMatchObject({ ok: true, formula: `SUM(${want})`, viaCopy: true, copySheets: [copy] });
  });

  it('the longest name Excel accepts (31 characters) is used whole', () => {
    expect(LONG.length).toBe(31);
    expect(new FormulaRemapper(spec({ sheetCopies: copies })).remap('長い!A1', 5, 5)).toMatchObject({ ok: true, formula: `${LONG}!A1` });
  });

  it('two sheets of the same name from different files: the one added later has the "(1)" suffix and is the one that is read', () => {
    const two = new Map([['master', 'a_Master(1)'], ['評価', 'a_評価(1)']]);
    const r = new FormulaRemapper(spec({ ...PARTIAL, sheetCopies: two })).remap('VLOOKUP(C5,Master!A:B,2,0)+E5', 5, 5);
    expect(r).toMatchObject({ ok: true, formula: "VLOOKUP(E5,'a_Master(1)'!A:B,2,0)+'a_評価(1)'!E5" });
  });
});

describe('copies: hostile input', () => {
  it('a formula that would be longer than Excel allows once the sheet names are in it stays a value', () => {
    const f = Array.from({ length: 900 }, () => 'レート!$B$1').join('+');
    expect(f.length).toBeLessThan(8192);
    expect(whyC(f)).toBe('otherSheet'); // the failure of the merged-table rules is reported
    expect(runC(f)).toMatchObject({ ok: false });
  });

  it('an unterminated string / unknown token / #REF! never reaches the copies', () => {
    for (const f of ['レート!A1+"abc', 'レート!A1+#REF!', 'レート!#REF!', 'レート!A1+[', 'レート!A1 ¥ 2', 'レート!A1+\u0007']) {
      const r = runC(f);
      expect(r.ok, f).toBe(false);
    }
  });

  it('spill references and spill functions stay values', () => {
    expect(whyC('C5#')).toBe('array');
    expect(whyC('SUM(C5#)')).toBe('array');
    expect(whyC('_xlfn.ANCHORARRAY(C5)')).toBe('array');
    // the copies cannot help: the spill master is imported as a value (the reason of the first failing reference is reported)
    expect(runC('レート!A1#').ok).toBe(false);
    expect(runC('SUM(レート!A1#)').ok).toBe(false);
  });

  it('a named LAMBDA / unknown function of the workbook stays a value', () => {
    expect(whyC('MyFn(レート!A1)', 5, 5, names([['MyFn', null, 'LAMBDA(x,x*2)']]))).toBe('unresolvedName');
    expect(whyC('_xludf.Custom(レート!A1)')).toBe('unsupported');
    expect(whyC('レート!Fn(A1)')).toBe('unsupported');
  });

  it('thousands of references remap in well under a second, deep nesting does not overflow', () => {
    const many = Array.from({ length: 1500 }, (_x, i) => `レート!$B$${(i % 90) + 1}`).join('+');
    const t0 = Date.now();
    const r = runC(many);
    expect(Date.now() - t0).toBeLessThan(2000);
    expect(r.ok).toBe(false); // > 8192 characters with the names: stays a value, but without throwing
    const deep = '('.repeat(500) + 'レート!A1' + ')'.repeat(500);
    expect(okC(deep)).toBe('('.repeat(500) + 'src_レート!A1' + ')'.repeat(500));
    const calls = 'IF(C5>1,'.repeat(100) + 'レート!A1' + ',0)'.repeat(100);
    expect(runC(calls).ok).toBe(true);
  });

  it('a sheet name that looks like a function / a boolean / a cell is only a sheet when it is followed by "!"', () => {
    const map = new Map([['sum', 'x_sum'], ['true', 'x_true'], ['a1', 'x_a1'], ['評価', 'x_評価']]);
    const r = (f: string) => new FormulaRemapper(spec({ sheetCopies: map })).remap(f, 5, 5);
    expect(r('SUM(C5:D5)')).toMatchObject({ ok: true, formula: 'SUM(E5:F5)' });
    expect(r("'sum'!A1")).toMatchObject({ ok: true, formula: "x_sum!A1" });
    expect(r("'a1'!B2")).toMatchObject({ ok: true, formula: "'x_a1'!B2".replace(/'/g, '') });
    expect(r('TRUE')).toMatchObject({ ok: true, formula: 'TRUE' });
  });

  it('non-string / odd input never throws', () => {
    for (const f of ['', ' ', '=', '()', ',,,', '{}', '{1,2;3,4}', '1:1', 'A:A', '$A$1:$B$2:$C$3', '!', 'レート!', "'レート'!", "''!A1", 'レート!A1:Master!B2', '\n']) {
      expect(() => runC(f)).not.toThrow();
    }
  });
});
