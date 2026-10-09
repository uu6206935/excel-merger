// Generates forms-extra-input.json for excel-forms-extra.ps1: a second batch of formulas typed into real Excel
// (Range.Formula2) — unknown / user-defined functions, defined names with '#' and '@', table specials incl. a totals
// row, '@' on scalars, whitespace, array constants. Sheet X: data B1:B3, spill Z1 (=SEQUENCE(3)); table TblSales on
// X!D1:G5 (header a | b | 金額 | calc, 3 data rows, totals row with SUM on 金額); names nmSpill = X!$Z$1, nmRange = X!$B$1:$B$3, nmRate = 0.1.
// Usage: node make-forms-extra-input.mjs   (writes ./forms-extra-input.json next to this file)
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const X = [
  ['A1', '=FOOBAR(1)'],
  ['A2', '=foobar(B1)'],
  ['A3', '=_xludf.FOO(1)'],
  ['A4', '=SUM(nmSpill#)'],
  ['A5', '=@nmRange'],
  ['A6', '=SUM(@nmRange)'],
  ['A7', '=B1*nmRate'],
  ['A8', '=@TODAY()'],
  ['A9', '=@B1'],
  ['A10', '=@1'],
  ['A11', '=-@B1:B3'],
  ['A12', '=SUM( B1:B3 )'],
  ['A13', '={1,2,3}'],
  ['A14', '=SUM({1,2;3,4})'],
  ['A15', '=Z1#+1'],
  ['A16', '=ROWS(Z1#)+@Z1#'],
  ['A17', '=SUM(TblSales[[#Totals],[a]])'],
  ['A18', '=SUM(TblSales[[#Data],[#Totals],[a]])'],
  ['A19', '=COUNTA(TblSales[#Totals])'],
  ['A20', '=SUM(TblSales[[#Headers],[#Data],[b]])'],
  ['A21', '=SUM(TblSales[[a]:[b]])'],
  ['A22', '=XLOOKUP(1,TblSales[a],TblSales[b])'],
  ['A23', '=LET(x,TblSales[a],SUM(x))'],
  ['A24', '=SUM(TblSales[金額])'],
  ['A25', '=LAMBDA(x,x*2)(B1:B3)'],
  ['A26', '=IFS(B1>0,1,TRUE,0)'],
  ['A27', '=SUMPRODUCT((B1:B3>1)*B1:B3)'],
  ['A28', '=TEXTJOIN(",",TRUE,B1:B3)&CONCAT("a","b")'],
  ['A29', '=UNIQUE(B1:B3)'],
  ['A30', '=SORTBY(B1:B3,B1:B3,-1)'],
  ['A31', '=FILTER(TblSales[a],TblSales[b]>0,"")'],
  ['A32', '=COUNTA(TblSales)'],
  ['A33', '=X!B1:B3'],
  ['A34', '=SUM(X!B1:B3 X!B2:B4)'],
  ['A35', '=IFERROR(VLOOKUP(1,TblSales,2,FALSE),"")'],
  ['A36', '=MAP(B1:B3,LAMBDA(v,IF(v>1,v,"")))'],
  ['A37', '=LET(a,1,b,2,LAMBDA(x,x+a+b)(3))'],
  ['A38', '=@INDEX(B1:B3,2)'],
  ['A39', '=INDEX(B1:B3,2)'],
  ['A40', '=CHOOSECOLS(B1:C3,1)'],
  ['A41', '=GROUPBY(B1:B3,B1:B3,SUM)'],
  ['A42', '=REGEXTEST("abc","b")'],
  ['A43', '=BYCOL(B1:C3,LAMBDA(c,MAX(c)))'],
  ['A44', '=ISOMITTED(1)'],
  // case normalisation (Excel upper-cases function names, refs, booleans, errors; names/tables get their defined case)
  ['A45', '=sum(b1:b3)'],
  ['A46', '=true'],
  ['A47', '=iferror(#n/a,1)'],
  ['A48', '=TblSales[[#this row],[a]]'],
  ['A49', '=SUM(tblsales[A])'],
  ['A57', '=xlookup(1,b1:b3,c1:c3)'],
  ['A58', '=LET(Xy,1,xy+1)'],
  ['A59', '=SUM(b:b)+SUM($b$1:c$3)'],
  ['A60', '=SUM(x!b1:b3)'],
  ['A61', '=nmrange'],
  ['A62', '=LET(z99,1,z99+1)'],
  ['A63', '=_XLFN.XLOOKUP(1,B1:B3,C1:C3)'],
  // eta-reduced lambdas (a function name passed as a value)
  ['A50', '=GROUPBY(B1:B3,B1:B3,PERCENTOF)'],
  ['A51', '=BYROW(B1:C3,SUM)'],
  ['A52', '=MAP(B1:B3,ABS)'],
  ['A53', '=PIVOTBY(B1:B3,C1:C3,B1:B3,SUM)'],
  ['A54', '=GROUPBY(B1:B3,B1:B3,HSTACK(SUM,AVERAGE))'],
  ['A55', '=LET(f,SUM,f(B1:B3))'],
  ['A56', '=BYCOL(B1:C3,counta)'],
  ['A64', '=BYROW(B1:C3,TEXTJOIN)'],
  ['A65', '=MAP(B1:B3,SUM)+SUM(B1)'],
  ['A66', '=REDUCE(0,B1:B3,SUM)'],
  ['A67', '=SCAN(0,B1:B3,SUM)'],
  ['A68', '=BYROW(B1:C3,STDEV.S)'],
  ['A69', '=GROUPBY(B1:B3,B1:B3,ARRAYTOTEXT)'],
];
const cells = X.map(([cell, f]) => ({ sheet: 'X', cell, candidates: [f] }));
// calculated column inside the table (UI form without table name)
cells.push({ sheet: 'X', cell: 'G2', candidates: ['=[@a]+[@金額]'] });
fs.writeFileSync(path.join(here, 'forms-extra-input.json'), JSON.stringify({ cells }, null, 1));
console.log(`wrote forms-extra-input.json: ${cells.length} cells`);
