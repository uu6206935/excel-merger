// Third forms probe (2024-2026 functions, IMPORTTEXT / IMPORTCSV, eta lambdas): writes forms-new-input.json for
// excel-forms-extra.ps1; extract-forms.mjs <xlsx> <com.json> forms-new-excel-results.json builds the fixture.
import fs from 'node:fs';
const f = [
  '=IMPORTTEXT("c:/x.txt")', '=IMPORTCSV("c:/x.csv")', '=PY(1,0)', '=TEXTSPLIT("a,b",",")', '=CHOOSECOLS(B1:C3,1)',
  '=PERCENTOF(B1:B3,B1:B4)', '=TRANSLATE("a")', '=DETECTLANGUAGE("a")', '=REGEXTEST("a","a")', '=GROUPBY(B1:B3,C1:C3,SUM)',
  '=ECMA.CEILING(1,1)', '=AGGREGATE(9,6,B1:B3)', '=SHEETS()', '=LAMBDA(x,x+1)(2)', '=MAKEARRAY(2,2,LAMBDA(r,c,r*c))',
  '=SUM(B1:B3)+@B1:B3', '=ISOMITTED(1)', '=TOCOL(B1:C3)', '=VALUETOTEXT(1)', '=EXPAND(B1,2,2)',
];
const cells = f.map((x, i) => ({ sheet: 'X', cell: 'J' + (i * 4 + 1), candidates: [x] }));
fs.writeFileSync(new URL('./forms-new-input.json', import.meta.url), JSON.stringify({ cells }, null, 1));
