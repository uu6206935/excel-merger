// Generates forms-input.json for excel-forms.ps1: formulas typed into real Excel (Range.Formula2) to learn how Excel
// stores them in the .xlsx (<f> text): _xlfn./_xlws. prefixes of every candidate "future function", LET/LAMBDA
// _xlpm. parameters, spill refs (ANCHORARRAY), the @ operator (SINGLE) and structured references.
// Usage: node make-forms-input.mjs   (writes ./forms-input.json next to this file)
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));

// Candidates: MS-XLSX 2.2.2 "Functions" table + future-function-list grammar + newer functions + controls
const FUTURE_CANDIDATES = `ACOT ACOTH AGGREGATE ARABIC BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND BITLSHIFT
BITOR BITRSHIFT BITXOR BYCOL BYROW CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT
CHISQ.TEST CHOOSECOLS CHOOSEROWS COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC
CSCH DAYS DECIMAL DROP ECMA.CEILING ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST
FIELDVALUE FILTERXML FLOOR.MATH FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY
FORECAST.ETS.STAT FORECAST.LINEAR FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS HSTACK HYPGEOM.DIST
IFNA IFS IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ISFORMULA ISO.CEILING ISOMITTED ISOWEEKNUM LAMBDA LET
LOGNORM.DIST LOGNORM.INV MAKEARRAY MAP MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST NETWORKDAYS.INTL
NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC
PERCENTRANK.INC PERMUTATIONA PHI POISSON.DIST QUARTILE.EXC QUARTILE.INC RANDARRAY RANK.AVG RANK.EQ REDUCE RRI SCAN
SEC SECH SEQUENCE SHEET SHEETS SKEW.P SORTBY STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST
TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW UNICHAR UNICODE UNIQUE VAR.P VAR.S VSTACK WEBSERVICE
WEIBULL.DIST WORKDAY.INTL WRAPCOLS WRAPROWS XLOOKUP XOR Z.TEST FILTER SORT
XMATCH ANCHORARRAY SINGLE ARRAYTOTEXT VALUETOTEXT IMAGE STOCKHISTORY GROUPBY PIVOTBY PERCENTOF REGEXTEST REGEXEXTRACT
REGEXREPLACE TRIMRANGE TRANSLATE DETECTLANGUAGE ENCODEURL COPILOT PY QUERYSTRING PQSOURCE LONGTEXT PYTHON_STR
PYTHON_TYPE PYTHON_TYPENAME`.split(/\s+/).filter(Boolean);

// controls: long-standing functions that must NOT get a prefix
const CONTROLS = `SUM VLOOKUP HLOOKUP INDEX MATCH OFFSET CHOOSE IFERROR SUMIFS COUNTIFS AVERAGEIFS SUMIF COUNTIF NETWORKDAYS
WORKDAY EDATE EOMONTH CEILING FLOOR RANK STDEV PERCENTILE QUARTILE TEXT DATEDIF INDIRECT ADDRESS ROUND ROUNDUP
ROUNDDOWN IF AND OR NOT LEFT RIGHT MID LEN FIND SEARCH SUBSTITUTE TRIM JIS ASC PHONETIC DBCS BAHTTEXT YEN
SUMPRODUCT TRANSPOSE FREQUENCY MODE NORMDIST CONCATENATE VALUE TODAY NOW RAND ISERROR ISNUMBER ISBLANK
HYPERLINK GETPIVOTDATA CUBEVALUE ROWS COLUMNS ROW COLUMN AVERAGE MAX MIN COUNT COUNTA`.split(/\s+/).filter(Boolean);

const argLists = ['', '1', '1,1', '1,1,1', '1,1,1,1', '1,1,1,1,1', '1,1,1,1,1,1', 'Z99', 'Z99:Z100', '1,Z99:Z100', 'Z99:Z100,1', 'Z99:Z100,1,1', 'Z99:Z100,Z99:Z100,1', 'Z99:Z100,Z99:Z100,Z99:Z100', 'Z99:Z100,1,Z99:Z100', 'Z99,1,1', 'x,1,x', '1,1,Z99:Z100', 'Z99:Z100,Z99:Z100'];
const cells = [];
let row = 1;
for (const fn of [...FUTURE_CANDIDATES, ...CONTROLS]) {
  cells.push({ sheet: 'F', cell: `A${row}`, fn, candidates: argLists.map((a) => `=IFERROR(${fn}(${a}),0)`) });
  row++;
}

// modern syntax on sheet S (spill source at Z1:Z3, data at B1:B3)
const S = [
  ['Z1', '=SEQUENCE(3)'],
  ['B1', '1'], ['B2', '2'], ['B3', '3'],
  ['A1', '=LET(x,1,y,x+1,x*y)'],
  ['A2', '=LET(x,1,LET(y,x+1,x+y))'],
  ['A3', '=LAMBDA(a,b,a+b)(1,2)'],
  ['A4', '=SUM(MAP(SEQUENCE(3),LAMBDA(v,v*2)))'],
  ['A5', '=LET(f,LAMBDA(n,n+1),f(2))'],
  ['A6', '=LET(total,SUM(B1:B3),total/2)'],
  ['A7', '=REDUCE(0,SEQUENCE(3),LAMBDA(acc,v,acc+v))'],
  ['A8', '=LET(x,1,x)+LET(x,2,x)'],
  ['A9', '=LAMBDA(a,[b],IF(ISOMITTED(b),a,a+b))(1)'],
  ['A10', '=SUM(Z1#)'],
  ['A11', '=SUM(S!Z1#)'],
  ['A12', '=@Z1:Z3'],
  ['A13', '=@INDEX(Z1:Z3,0)'],
  ['A14', '=@(Z1:Z3+1)'],
  ['A15', '=SUM(@Z1:Z3)'],
  ['A16', '=INDEX(Z1#,2)'],
  ['A17', '=ROWS(Z1#)'],
  ['A18', "=SUM('集計 シート'!B1:B3)"],
  ['A19', "=ROWS('集計 シート'!Z1#)"],
  ['A20', '=LET(x,B1:B3,SUM(@x))'],
  ['A21', '=LET(rate,0.1,B1*rate)'],
  ['A22', '=LET(a,1,b,LAMBDA(x,x+a),b(1))'],
  ['A23', '=BYROW(B1:B3,LAMBDA(r,SUM(r)))'],
  ['A24', '=Z1#*2'],
  ['A25', '=Z1:Z3'],
  ['A26', '=MAX(@B1:B3,1)'],
  ['A27', '=@B1:B3+1'],
  ['A28', '=IF(@B1:B3>1,1,0)'],
  ['A29', '=XLOOKUP(2,B1:B3,Z1:Z3)'],
  ['A30', '=SORT(B1:B3,1,-1)'],
  ['A31', '=FILTER(B1:B3,B1:B3>1)'],
  ['A32', '=_xlfn.XLOOKUP(2,B1:B3,Z1:Z3)'],
  ['A33', '=LET(_x,1,_x+1)'],
  ['A34', '=LET(x.y,1,x.y+1)'],
  ['A35', '=LET(値,5,値*2)'],
];
for (const [cell, f] of S) cells.push({ sheet: 'S', cell, candidates: [f] });
cells.push({ sheet: '集計 シート', cell: 'Z1', candidates: ['=SEQUENCE(2)'] });
cells.push({ sheet: '集計 シート', cell: 'B1', candidates: ['5'] });

// structured references (table Tbl on T!A1:E4, column E = calculated column)
const T = [
  ['G2', '=Tbl[@a]'],
  ['G3', '=Tbl[@[b c]]'],
  ['G4', '=SUM(Tbl[@[a]:[b c]])'],
  ['H2', '=SUM(Tbl[a])'],
  ['H3', '=Tbl[[#This Row],[a]]'],
  ['H4', '=Tbl[@単価]'],
  ['I2', "=Tbl[@[d'#e]]"],
  ['I3', '=SUM(Tbl[@])'],
  ['I4', '=ROWS(Tbl[#Data])'],
  ['J2', '=SUM(Tbl[[#Headers],[a]])'],
  ['J3', '=SUM(Tbl[[a]:[b c]])'],
  ['J4', '=COUNTA(Tbl[#All])'],
  ['K2', '=SUM(Tbl[[#This Row],[a]:[単価]])'],
  ['K3', '=Tbl[@[単価]]'],
  ['K4', '=SUM(Tbl)'],
];
for (const [cell, f] of T) cells.push({ sheet: 'T', cell, candidates: [f] });

fs.writeFileSync(path.join(here, 'forms-input.json'), JSON.stringify({ cells }, null, 1));
console.log(`wrote forms-input.json: ${cells.length} cells`);
