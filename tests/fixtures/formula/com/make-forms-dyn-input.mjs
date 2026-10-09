// Fourth forms probe: which formulas typed through Range.Formula2 does Excel store as dynamic arrays (cm="1")?
// Range-parameter functions of dynamic.ts RANGE_PARAMS with ranges in their range positions (must NOT be dynamic
// for isDynamicArrayFormula to be allowed to say false) and ranges in value positions (dynamic).
// Writes forms-dyn-input.json for excel-forms-extra.ps1; extract-forms.mjs <xlsx> <com.json> forms-dyn-excel-results.json.
import fs from 'node:fs';
const f = [
  // range parameters -> expected NOT dynamic
  '=SUMIF(B1:B4,">1",C1:C4)', '=SUMIFS(C1:C4,B1:B4,">1",B1:B4,"<9")', '=COUNTIF(B1:B4,">1")', '=COUNTIFS(B1:B4,">1",C1:C4,">1")',
  '=AVERAGEIF(B1:B4,">1",C1:C4)', '=AVERAGEIFS(C1:C4,B1:B4,">1")', '=MAXIFS(C1:C4,B1:B4,">1")', '=MINIFS(C1:C4,B1:B4,">1")',
  '=VLOOKUP(2,B1:C4,2,FALSE)', '=HLOOKUP(1,B1:C4,2,FALSE)', '=MATCH(2,B1:B4,0)', '=XMATCH(2,B1:B4)', '=LOOKUP(2,B1:B4,C1:C4)',
  '=INDEX(B1:C4,2,2)', '=INDEX(B1:B4,MATCH(2,B1:B4,0))', '=LARGE(B1:B4,2)', '=SMALL(B1:B4,2)', '=RANK.EQ(2,B1:B4)',
  '=PERCENTILE.INC(B1:B4,0.5)', '=QUARTILE.INC(B1:B4,1)', '=PERCENTRANK.INC(B1:B4,2)', '=TRIMMEAN(B1:B4,0.2)',
  '=SUBTOTAL(9,B1:B4)', '=NPV(0.1,B1:B4)', '=IRR(B1:C4)', '=MIRR(B1:B4,0.1,0.1)', '=CORREL(B1:B3,C1:C3)',
  '=FORECAST.LINEAR(4,B1:B3,C1:C3)', '=SLOPE(B1:B3,C1:C3)', '=SUMX2MY2(B1:B3,C1:C3)', '=MEDIAN(B1:B4)', '=MODE.SNGL(B1:B4)',
  '=STDEV.S(B1:B4)', '=AND(B1:B4)', '=OR(B1:B4)', '=CONCAT(B1:B4)', '=COUNTBLANK(B1:C9)', '=AREAS((B1:B2,C1:C2))', '=SUMSQ(B1:B4)',
  '=PRODUCT(B1:B4)', '=XNPV(0.1,B1:B3,C1:C3)', '=DSUM(D1:F4,"a",D1:D2)', '=SUMPRODUCT(B1:B3,C1:C3)', '=GEOMEAN(B1:B4)',
  '=COUNT(B1:C4)+COUNTA(B:B)', '=MAX(B:B)-MIN(B:B)', '=AVERAGE(B1:B4)', '=ROWS(B1:C4)*COLUMNS(B1:C4)', '=XLOOKUP(2,B1:B4,C1:C4,"")',
  '=SUMIF(B:B,2,C:C)', '=INDEX(B:B,3)', '=VAR.S(B1:B4)', '=AVEDEV(B1:B4)',
  // ranges in value positions -> expected dynamic
  '=VLOOKUP(B1:B3,B1:C4,2,FALSE)', '=SUMIF(B1:B4,B1:B2,C1:C4)', '=INDEX(B1:C4,2)', '=ROW(B1:B3)', '=IF(B1:B3>1,1,0)',
  '=CHOOSE(1,B1:B3,C1:C3)', '=MMULT(B1:C2,B1:C2)', '=IFERROR(B1:B3,0)', '=XLOOKUP(2,B1:B4,B1:C4)', '=OFFSET(B1,0,0)',
  '=SUM(OFFSET(B1,0,0,3,1))', '=N(B1:B3)', '=SUM(B1:B3*2)', '=MAX(IF(B1:B3>1,B1:B3))', '=LEN(B1:B3)', '=ABS(B1:B3)',
  '=B1:B3=2', '=COLUMN(B1:C1)', '=INDEX(B1:B4,0)', '=MDETERM(B1:C2)', '=ISNUMBER(B1:B3)', '=TEXT(B1:B3,"0")',
];
const cells = f.map((x, i) => ({ sheet: 'X', cell: 'J' + (i * 5 + 1), candidates: [x] }));
fs.writeFileSync(new URL('./forms-dyn-input.json', import.meta.url), JSON.stringify({ cells }, null, 1));
console.log(`${cells.length} cells`);
