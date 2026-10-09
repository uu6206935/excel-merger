// isDynamicArrayFormula (dynamic.ts): predicts Excel's dynamic-array flag (cm="1") of a formula typed in Excel 365.
// Checked against every formula typed through Range.Formula2 in the forms fixtures (the stored cell's cm attribute is
// Excel's answer): never "false" where Excel says dynamic; the few extra "true"s are listed (harmless).
import { describe, expect, it } from 'vitest';
import { isDynamicArrayFormula } from '../../src/core/formula/dynamic';
import { toFileForm, type FormOptions } from '../../src/core/formula/forms';
import { readFixture } from './helpers';

interface ExcelCell { sheet: string; cell: string; input: string | null; file: { f: string | null; cm?: string } | null }
const NAMES: Record<string, string> = { nmspill: 'X!$Z$1', nmrange: 'X!$B$1:$B$3', nmrate: '0.1' };
const resolveName = (n: string) => NAMES[n.toLowerCase()] ?? null;

describe('isDynamicArrayFormula vs real Excel (cm flag of formulas typed through Formula2)', () => {
  const cells: Array<ExcelCell & { fixture: string }> = [];
  for (const fixture of ['forms-excel-results.json', 'forms-extra-excel-results.json', 'forms-new-excel-results.json', 'forms-tro-excel-results.json',
    'forms-dyn-excel-results.json']) {
    for (const c of readFixture<{ cells: ExcelCell[] }>(fixture).cells) if (c.input?.startsWith('=') && c.file?.f != null) cells.push({ ...c, fixture });
  }
  const opts = (c: ExcelCell): FormOptions => ({
    hostTable: (c.sheet === 'T' && /^E\d/.test(c.cell)) || (c.sheet === 'X' && /^[D-G]\d$/.test(c.cell)) ? (c.sheet === 'T' ? 'Tbl' : 'TblSales') : undefined,
    tables: ['Tbl', 'TblSales'], names: ['nmSpill', 'nmRange', 'nmRate'],
  });
  it(`never misses a dynamic formula (${cells.length} typed formulas)`, () => {
    expect(cells.length).toBeGreaterThan(450);
    const missed: string[] = [];
    const extra: string[] = [];
    let dyn = 0;
    for (const c of cells) {
      const ours = isDynamicArrayFormula(toFileForm(c.input!, opts(c)), { resolveName });
      const excel = !!c.file!.cm;
      if (excel) dyn++;
      if (excel && !ours) missed.push(`${c.fixture} ${c.sheet}!${c.cell} ${c.input} (${c.file!.f})`);
      if (!excel && ours) extra.push(`${c.sheet}!${c.cell} ${c.input}`);
    }
    expect(missed).toEqual([]);
    expect(dyn).toBeGreaterThan(90);
    // conservative extras (Excel does not flag them; flagging a formula that returns one value is harmless):
    // INDEX with one position over a 2-D or spill range (returns a whole row; Excel does not flag it)
    expect(extra).toEqual(['S!A16 =INDEX(Z1#,2)', 'X!J276 =INDEX(B1:C4,2)']);
  });
});

describe('isDynamicArrayFormula unit cases', () => {
  const plain = [
    'SUM(A1:A10)', "VLOOKUP(A2,'マスタ'!A:D,3,FALSE)", 'IF(A2>0,B2*C2,"")', 'SUMIFS(C:C,A:A,A2,B:B,">0")',
    'INDEX(B:B,MATCH(A2,A:A,0))', 'INDEX($B$2:$D$9,MATCH(A2,$A$2:$A$9,0),2)', 'COUNTIF(A:A,A2)>1', 'ROUND(A1*1.1,0)',
    'TEXT(A1,"yyyy/mm/dd")&"様"', 'IFERROR(VLOOKUP(A2,B:D,2,0),"")', '_xlfn.XLOOKUP(A2,B:B,C:C,"")', 'DATE(2024,1,1)',
    'Tbl[[#This Row],[a]]*2', 'SUM(Tbl[a])', '_xlfn.LET(_xlpm.x,A1,_xlpm.x*2)', 'SUMPRODUCT(B1:B3,C1:C3)', 'AVERAGE(1,2,3)',
    'MAX(A:A)-MIN(A:A)', 'COUNTIFS(A:A,">0",B:B,"x")', '_xlfn.SINGLE(A1:A3)+1', 'SUM({1,2;3,4})', 'ROWS(A1:A9)*2',
    'SUBTOTAL(109,B2:B9)', '_xlfn.TEXTJOIN(",",TRUE,A1:A3)', 'LARGE(A1:A9,2)', 'MATCH(A1,B1:B9,0)', 'SUM(_xlfn.ANCHORARRAY(A1))',
    'SUM(_xlfn._TRO_TRAILING(A:A))', 'NOW()', '"a"&"b"', '#N/A', 'TRUE', '1', 'SUM((A1:B2,C3))', 'SUM(A1:B3 B2:C4)',
    'Tbl[[#Totals],[a]]', 'IF(nmRate>0,1,0)', 'OFFSET(A1,1,1)', 'SUM(OFFSET(A1,0,0,3,1))', 'OFFSET(A1,0,0,1,1)',
  ];
  const dynamic = [
    'A1:A3', 'Sheet2!B:B', '_xlfn.SEQUENCE(3)', '_xlfn._xlws.FILTER(A1:A9,B1:B9>0)', '_xlfn._xlws.SORT(A1:A9)', '_xlfn.UNIQUE(A:A)',
    'SUM(A1:A3*B1:B3)', 'IF(A1:A3>0,1,0)', 'VLOOKUP(A1:A3,B:C,2,0)', '_xlfn.XLOOKUP(A1,B:B,C:D)', 'INDEX(A1:C3,2)',
    'INDEX(A1:A9,0)', 'ROW(A1:A3)', 'A1:A3&"x"', 'TRANSPOSE(A1:C1)', 'FOOBAR(1)', '_xludf.MYFN(A1)', 'nmUnknown',
    '_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(1)', '_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.n,_xlpm.n),_xlpm.f(1))', '{1,2,3}',
    'Tbl[a]', 'Tbl[[#This Row],[a]:[b]]', '_xlfn.ANCHORARRAY(A1)', '-A1:A3', 'SUMIF(A:A,B1:B3,C:C)',
    'INDIRECT("A1")', 'MMULT(A1:B2,C1:D2)', 'OFFSET(A1:A2,1,1)', 'OFFSET(A1,0,0,3)', '_xlfn._TRO_TRAILING(A:A)', '_xlfn.IFS(A1>0,1,TRUE,0)', 'SUM(A1:A3', '[1]Sheet1!A1:A3',
    '_xlfn.SINGLE((A1:A3+1))', 'CHOOSE(1,A1:A3,B1:B3)', '_xlfn.LET(_xlpm.r,A1:A3,_xlpm.r)', 'SUM(nmRange*2)',
  ];
  it('ordinary business formulas are not dynamic', () => {
    expect(plain.filter((f) => isDynamicArrayFormula(f, { resolveName }))).toEqual([]);
  });
  it('spilling / array formulas are', () => {
    expect(dynamic.filter((f) => !isDynamicArrayFormula(f, { resolveName }))).toEqual([]);
  });
  it('defined names: without a resolver every name counts as a range', () => {
    expect(isDynamicArrayFormula('B1*nmRate')).toBe(true);
    expect(isDynamicArrayFormula('B1*nmRate', { resolveName })).toBe(false);
    expect(isDynamicArrayFormula('SUM(nmRange)', { resolveName })).toBe(false);
    expect(isDynamicArrayFormula('nmRange', { resolveName })).toBe(true);
    // a self-referencing name never loops
    expect(isDynamicArrayFormula('a+1', { resolveName: () => 'a' })).toBe(true);
  });
  it('never throws on junk', () => {
    const parts = ['SUM(', 'A1:B2', ')', '(', ',', '{', '}', '@', '#', 'nm', '_xlfn.LET(', '_xlpm.x', '"', '+', 'Tbl[a]', ' ', '1', ':'];
    let seed = 3;
    const rnd = (n: number) => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed % n; };
    for (let k = 0; k < 5000; k++) {
      let f = '';
      for (let j = 0, n = 1 + rnd(9); j < n; j++) f += parts[rnd(parts.length)];
      expect(typeof isDynamicArrayFormula(f)).toBe('boolean');
    }
    expect(isDynamicArrayFormula('('.repeat(20000) + 'A1' + ')'.repeat(20000))).toBe(true);
  });
});
