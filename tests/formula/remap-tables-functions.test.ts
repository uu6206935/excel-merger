/**
 * FormulaRemapper: Excel Table structured references (rule d) and functions (rule e).
 * Table1 = A1:F11 on 評価 with the columns 学生ID 氏名 評価1 評価2 評価3 合計 (A, B not imported; C..F -> E..H).
 */
import { describe, expect, it } from 'vitest';
import { FormulaRemapper, parseSref, type RemapReason, type RemapTable } from '../../src/core/formula/remap';
import { ok, run, spec, why, type Over } from './remapHelpers';

const COLS = ['学生ID', '氏名', '評価1', '評価2', '評価3', '合計'];
const table = (over: Partial<RemapTable> = {}): RemapTable => ({ name: 'Table1', sheet: '評価', ref: 'A1:F11', headerRow: true, totalsRow: false, columns: COLS, ...over });
const T = (over: Partial<RemapTable> = {}): Over => ({ tables: [table(over)] });
/** every column imported (A..F -> C..H) */
const ALL: Over = { colMap: new Map([[1, 3], [2, 4], [3, 5], [4, 6], [5, 7], [6, 8]]) };

describe('structured references: this row', () => {
  it.each<[string, string]>([
    ['Table1[[#This Row],[評価1]]', 'E5'],
    ['Table1[[#This Row],[評価1]]+Table1[[#This Row],[評価2]]', 'E5+F5'],
    ['Table1[[#This Row],[評価1]]*2', 'E5*2'],
    ['Table1[@評価1]', 'E5'],
    ['Table1[@[評価1]]', 'E5'],
    ['Table1[#This Row]', 'notImported'],
    ['SUM(Table1[[#This Row],[評価1]:[評価3]])', 'SUM(E5:G5)'],
    ['SUM(Table1[@[評価1]:[評価3]])', 'SUM(E5:G5)'],
    ['AVERAGE(Table1[[#This Row],[評価1]:[評価2]])', 'AVERAGE(E5:F5)'],
    ['IF(Table1[[#This Row],[評価1]]>=80,"優","可")', 'IF(E5>=80,"優","可")'],
    ['table1[[#this row],[評価1]]', 'E5'],
    ['TABLE1[[#This Row],[評価3]]', 'G5'],
  ])('%s -> %s', (f, want) => {
    if (want === 'notImported') expect(why(f, 5, 5, T())).toBe('notImported');
    else expect(ok(f, 5, 5, T())).toBe(want);
  });

  it('follows the row the formula was moved to', () => {
    expect(ok('Table1[[#This Row],[評価1]]', 5, 8, T())).toBe('E8');
  });

  it('all columns imported: the whole row is one area', () => {
    expect(ok('SUM(Table1[#This Row])', 5, 5, { ...T(), ...ALL })).toBe('SUM(C5:H5)');
  });

  it('a formula outside the table\'s data rows has no "this row"', () => {
    expect(why('Table1[[#This Row],[評価1]]', 14, 14, T())).toBe('unresolvedTable');
    expect(why('Table1[[#This Row],[評価1]]', 1, 1, T())).toBe('unresolvedTable');
  });
});

describe('structured references: columns, headers, the whole table', () => {
  it.each<[string, string]>([
    ['SUM(Table1[評価3])', 'SUM($G$2:$G$11)'],
    ['SUM(Table1[[#Data],[評価3]])', 'SUM($G$2:$G$11)'],
    ['AVERAGE(Table1[評価1])', 'AVERAGE($E$2:$E$11)'],
    ['Table1[[#Headers],[評価1]]', '$E$1'],
    ['SUM(Table1[[評価1]:[評価3]])', 'SUM($E$2:$G$11)'],
    ['SUM(Table1[[#Data],[評価1]:[評価3]])', 'SUM($E$2:$G$11)'],
    ['RANK(Table1[[#This Row],[評価1]],Table1[評価1])', 'RANK(E5,$E$2:$E$11)'],
    ['INDEX(Table1[評価2],MATCH(Table1[[#This Row],[評価1]],Table1[評価1],0))', 'INDEX($F$2:$F$11,MATCH(E5,$E$2:$E$11,0))'],
    ['COUNTA(Table1[[#Data],[評価1]])', 'COUNTA($E$2:$E$11)'],
    ['Table1[[#Headers],[評価1]]&":"&Table1[[#This Row],[評価1]]', '$E$1&":"&E5'],
    ['SUMIF(Table1[評価1],">"&Table1[[#This Row],[評価1]],Table1[評価2])', 'SUMIF($E$2:$E$11,">"&E5,$F$2:$F$11)'],
    ['Table1[評価1]', '$E$2:$E$11'],
  ])('%s -> %s', (f, want) => {
    expect(ok(f, 5, 5, T())).toBe(want);
  });

  it('the whole Table (data body) inside a function that takes ranges; all columns must be imported', () => {
    expect(why('ROWS(Table1[])', 5, 5, T())).toBe('notImported');
    expect(ok('ROWS(Table1[])', 5, 5, { ...T(), ...ALL })).toBe('ROWS($C$2:$H$11)');
    expect(ok('SUM(Table1[#Data])', 5, 5, { ...T(), ...ALL })).toBe('SUM($C$2:$H$11)');
    expect(ok('COUNTA(Table1[#All])', 5, 5, { ...T(), ...ALL })).toBe('COUNTA($C$1:$H$11)');
    expect(ok('COUNTA(Table1[[#Headers],[#Data]])', 5, 5, { ...T(), ...ALL })).toBe('COUNTA($C$1:$H$11)');
    expect(ok('COLUMNS(Table1[[#Headers]])', 5, 5, { ...T(), ...ALL })).toBe('COLUMNS($C$1:$H$1)');
  });

  it('the whole Table outside a function that takes a range is not a value the formula could use', () => {
    expect(why('Table1[]', 5, 5, { ...T(), ...ALL })).toBe('unresolvedTable');
    expect(why('Table1[#Data]', 5, 5, { ...T(), ...ALL })).toBe('unresolvedTable');
    expect(why('IF(Table1[#Data],1,2)', 5, 5, { ...T(), ...ALL })).toBe('unresolvedTable');
  });

  it('the table\'s data rows follow the rules of the A1 range (rows not placed / appended)', () => {
    const miss = Array.from({ length: 10 }, (_x, i) => [i + 2, i + 2] as [number, number]).filter(([r]) => r !== 7);
    const r = run('SUM(Table1[評価1])', 5, 5, { ...T(), rowPairs: miss });
    expect(r.ok && r.formula).toBe('SUM($E$2:$E$11)');
    expect(r.ok && r.widened && !r.exact).toBe(true);
  });

  it('a table with a totals row: the data body stops above it; the totals row itself cannot be read', () => {
    const t = T({ ref: 'A1:F12', totalsRow: true });
    expect(ok('SUM(Table1[評価3])', 5, 5, t)).toBe('SUM($G$2:$G$11)');
    expect(why('Table1[[#Totals],[評価1]]', 5, 5, t)).toBe('unresolvedTable');
    expect(why('SUM(Table1[[#All],[評価1]])', 5, 5, t)).toBe('unresolvedTable');
    expect(why('Table1[#Totals]', 5, 5, t)).toBe('unresolvedTable');
    expect(why('Table1[[#Data],[#Totals],[評価1]]', 5, 5, t)).toBe('unresolvedTable');
  });

  it('a table without a header row', () => {
    const t = T({ ref: 'A2:F11', headerRow: false });
    expect(ok('SUM(Table1[評価3])', 5, 5, t)).toBe('SUM($G$2:$G$11)');
    expect(why('Table1[[#Headers],[評価1]]', 5, 5, t)).toBe('unresolvedTable');
  });

  it.each<[string, RemapReason, Over, string]>([
    ['SUM(Nope[評価1])', 'unresolvedTable', T(), 'unknown table'],
    ['SUM(Table1[ない])', 'unresolvedTable', T(), 'unknown column'],
    ['SUM(Table1[[評価1]:[ない]])', 'unresolvedTable', T(), 'unknown end column'],
    ['SUM(Table1[評価1])', 'otherSheet', T({ sheet: 'Other' }), 'a table on another sheet'],
    ['SUM(Sheet2!Table1[評価1])', 'otherSheet', T(), 'qualified with another sheet'],
    ['SUM([1]!Table1[評価1])', 'unsafe', T(), 'an external workbook'],
    ['SUM(Table1[[#Unknown],[評価1]])', 'unresolvedTable', T(), 'unknown special item'],
    ['SUM(Table1[[評価1],[評価2]])', 'unresolvedTable', T(), 'two columns without a range'],
  ])('%s -> %s (%s)', (f, want, over) => {
    expect(why(f, 5, 5, over)).toBe(want);
  });

  it('column names with special characters (escaped with an apostrophe in the file form)', () => {
    const t = T({ columns: ['学生ID', '氏名', "点数[1]", "率#", "a'b", '合計'] });
    expect(ok("SUM(Table1[点数'[1']])", 5, 5, t)).toBe('SUM($E$2:$E$11)');
    expect(ok("SUM(Table1[[#Data],[点数'[1']]])", 5, 5, t)).toBe('SUM($E$2:$E$11)');
    expect(ok("SUM(Table1[率'#])", 5, 5, t)).toBe('SUM($F$2:$F$11)');
    expect(ok("Table1[[#This Row],[a''b]]", 5, 5, t)).toBe('G5');
  });

  it('structured references reached through a defined name', () => {
    const names = [{ name: '評価列', scope: null, refersTo: 'Table1[評価1]' }];
    expect(ok('SUM(評価列)', 5, 5, { ...T(), names })).toBe('SUM($E$2:$E$11)');
    expect(why('SUM(評価列)', 5, 5, { names })).toBe('unresolvedTable');
  });
});

describe('structured references without a table name (XlsxWriter writes them inside a table)', () => {
  it('belong to the table that contains the formula cell', () => {
    const rem = new FormulaRemapper(spec(T()));
    const f = '[[#This Row],評価1]+10';
    expect(rem.remap(f, 5, 5, 6)).toMatchObject({ ok: true, formula: 'E5+10' });
    expect(rem.remap('SUM([[#This Row],[評価1]:[評価3]])', 5, 5, 6)).toMatchObject({ ok: true, formula: 'SUM(E5:G5)' });
    expect(rem.remap(f, 5, 5, 9)).toMatchObject({ ok: false, reason: 'unresolvedTable' }); // the cell is outside the table
    expect(rem.remap(f, 5, 5)).toMatchObject({ ok: false, reason: 'unresolvedTable' }); // the column of the cell is not known
  });
});

describe('parseSref', () => {
  it('bare items between the brackets and no table name', () => {
    expect(parseSref('[[#This Row],点数]')).toEqual({ table: '', specials: ['this row'], colFrom: '点数', colTo: '点数' });
    expect(parseSref('[[#This Row],点数:合計]')).toEqual({ table: '', specials: ['this row'], colFrom: '点数', colTo: '合計' });
  });

  it.each<[string, { table: string; specials: string[]; colFrom: string | null; colTo: string | null } | null]>([
    ['Table1[評価1]', { table: 'Table1', specials: [], colFrom: '評価1', colTo: '評価1' }],
    ['Table1[[#This Row],[評価1]]', { table: 'Table1', specials: ['this row'], colFrom: '評価1', colTo: '評価1' }],
    ['Table1[@評価1]', { table: 'Table1', specials: ['this row'], colFrom: '評価1', colTo: '評価1' }],
    ['Table1[@[評価1]:[評価3]]', { table: 'Table1', specials: ['this row'], colFrom: '評価1', colTo: '評価3' }],
    ['Table1[[#Headers],[評価1]]', { table: 'Table1', specials: ['headers'], colFrom: '評価1', colTo: '評価1' }],
    ['Table1[#Data]', { table: 'Table1', specials: ['data'], colFrom: null, colTo: null }],
    ['Table1[]', { table: 'Table1', specials: [], colFrom: null, colTo: null }],
    ['Table1[[#All]]', { table: 'Table1', specials: ['all'], colFrom: null, colTo: null }],
    ['[@評価1]', { table: '', specials: ['this row'], colFrom: '評価1', colTo: '評価1' }],
    ["Table1[点数'[1']]", { table: 'Table1', specials: [], colFrom: '点数[1]', colTo: '点数[1]' }],
    ['Table1[[#Foo]]', null],
    ['Table1[[a],[b],[c]]', null],
    ['Table1', null],
    ['Table1[a', null],
  ])('%s', (text, want) => {
    expect(parseSref(text)).toEqual(want);
  });
});

describe('functions (rule e)', () => {
  it.each<[string]>([
    ['ROW()'], ['ROW(C5)'], ['ROW()-1'], ['ROW(C5)-1'], ['MOD(ROW(),2)'], ['row()'], ['COLUMN()'], ['COLUMN(C5)'], ['OFFSET(C5,1,0)'], ['SUM(OFFSET(C5,0,0,3,1))'],
    ['INDIRECT("C"&ROW())'], ['INDIRECT("C5")'], ['ADDRESS(1,1)'], ['CELL("address",C5)'], ['FORMULATEXT(C5)'], ['_xlfn.FORMULATEXT(C5)'], ['SHEET()'], ['_xlfn.SHEET()'],
    ['_xlfn.SHEETS()'], ['INFO("recalc")'], ['ISFORMULA(C5)'], ['_xlfn.ISFORMULA(C5)'], ['GETPIVOTDATA("x",C5)'], ['IF(1,ROW())'], ['HYPERLINK("#\'評価\'!C"&ROW(),"x")'],
    ['TEXT(ROW(),"0")'], ['_xlfn.IFS(C5>1,ROW(),TRUE,0)'],
  ])('%s is positional', (f) => {
    expect(why(f)).toBe('positional');
  });

  it.each<[string, string]>([
    ['ROWS(C2:C11)', 'ROWS(E2:E11)'],
    ['COLUMNS(C2:F2)', 'COLUMNS(E2:H2)'],
    ['SUBTOTAL(9,C2:C11)', 'SUBTOTAL(9,E2:E11)'],
    ['AGGREGATE(9,6,C2:C11)', 'AGGREGATE(9,6,E2:E11)'],
    ['SUMPRODUCT(C2:C11,D2:D11)', 'SUMPRODUCT(E2:E11,F2:F11)'],
    ['CHOOSE(2,C5,D5)', 'CHOOSE(2,E5,F5)'],
    ['_xlfn.XMATCH(C5,$C$2:$C$11)', '_xlfn.XMATCH(E5,$E$2:$E$11)'],
    ['SUMIFS($E$2:$E$11,$C$2:$C$11,">"&C5,$D$2:$D$11,"<"&D5)', 'SUMIFS($G$2:$G$11,$E$2:$E$11,">"&E5,$F$2:$F$11,"<"&F5)'],
    ['AVERAGEIFS($D$2:$D$11,$C$2:$C$11,C5)', 'AVERAGEIFS($F$2:$F$11,$E$2:$E$11,E5)'],
    ['HYPERLINK("https://example.com/"&C5,D5)', 'HYPERLINK("https://example.com/"&E5,F5)'],
    ['NOW()', 'NOW()'],
    ['TODAY()-C5', 'TODAY()-E5'],
    ['RAND()', 'RAND()'],
    ['RANDBETWEEN(1,6)', 'RANDBETWEEN(1,6)'],
    ['TRANSPOSE(C2:E2)', 'TRANSPOSE(E2:G2)'],
    ['N(C5)', 'N(E5)'],
    ['T(C5)', 'T(E5)'],
    ['ISNUMBER(C5)', 'ISNUMBER(E5)'],
    ['IFNA(VLOOKUP(C5,$C$2:$F$11,2,0),"-")', 'IFNA(VLOOKUP(E5,$E$2:$H$11,2,0),"-")'],
    ['_xlfn.SWITCH(C5,1,"a",2,"b","c")', '_xlfn.SWITCH(E5,1,"a",2,"b","c")'],
    ['_xlfn.TEXTJOIN(",",TRUE,C5:E5)', '_xlfn.TEXTJOIN(",",TRUE,E5:G5)'],
    ['_xlfn.MAXIFS($C$2:$C$11,$D$2:$D$11,">"&D5)', '_xlfn.MAXIFS($E$2:$E$11,$F$2:$F$11,">"&F5)'],
    ['NOSUCHFN(C5)', 'NOSUCHFN(E5)'],
    ['_xlfn.SINGLE(C5)', '_xlfn.SINGLE(E5)'],
    ['_xlfn._xlws.FILTER($D$2:$D$11,$C$2:$C$11>C5)', '_xlfn._xlws.FILTER($F$2:$F$11,$E$2:$E$11>E5)'],
    ['SUM(_xlfn._TRO_TRAILING(C2:C11))', 'SUM(_xlfn._TRO_TRAILING(E2:E11))'],
    ['DATE(2026,4,1)+E5', 'DATE(2026,4,1)+G5'],
  ])('%s -> %s', (f, want) => {
    expect(ok(f)).toBe(want);
  });

  it('spill references and functions the module cannot know', () => {
    expect(why('_xlfn.ANCHORARRAY(F2)')).toBe('array');
    expect(why('SUM(_xlfn.ANCHORARRAY(C2))')).toBe('array');
    expect(why('SUM(F2#)')).toBe('array');
    expect(why('_xludf.MyFunc(C5)')).toBe('unsupported');
    expect(why('[1]!Func(C5)')).toBe('unsupported');
    expect(why('Sheet1!Fn(C5)')).toBe('unsupported');
  });
});
