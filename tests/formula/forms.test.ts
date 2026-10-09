// UI form <-> FILE form conversion (forms.ts), checked against real Excel 365 (fixtures/formula/forms-excel-results.json:
// formulas typed via Range.Formula2, saved, <f> text read from the .xlsx XML — see fixtures/formula/com/*), plus
// round trips over every Excel-written formula of all fixtures and many unit cases. rewriteRefs / isFormulaInput.
import { describe, expect, it } from 'vitest';
import {
  toFileForm, toUiForm, isFormulaInput, rewriteRefs, FUTURE_FUNCTIONS, XLWS_FUNCTIONS, fileFunctionName, type FormOptions,
} from '../../src/core/formula/forms';
import { parseAreaText, tokenize } from '../../src/core/formula/excelRef';
import { readFixture } from './helpers';

interface ExcelCell {
  sheet: string; cell: string; fn: string | null; input: string | null; formula2: string | null; formula: string | null;
  err: string | null; file: { f: string | null; t?: string; ref?: string; cm?: string } | null;
}
const X = readFixture<{ cells: ExcelCell[] }>('forms-excel-results.json').cells;
const optsFor = (c: ExcelCell): FormOptions => ({ hostTable: c.sheet === 'T' && /^E\d/.test(c.cell) ? 'Tbl' : undefined, tables: ['Tbl'] });

// Excel drops a typed '@' where legacy implicit intersection applies anyway (stored without _xlfn.SINGLE, shown
// again with '@'); we keep _xlfn.SINGLE — same meaning, same display (DEVIATIONS.md "Formula forms").
const LEGACY_AT = new Set(['S!A12', 'S!A13', 'S!A27', 'S!A28']);
// functions this Excel build (16.0.20430) does not know: stored as plain names; the MS-XLSX list says _xlfn.
const UNKNOWN_TO_BUILD = new Set(['COPILOT', 'QUERYSTRING', 'LONGTEXT', 'PYTHON_STR', 'PYTHON_TYPE', 'PYTHON_TYPENAME']);

describe('forms vs real Excel (Range.Formula2 -> saved <f>)', () => {
  const typed = X.filter((c) => c.input && c.input.startsWith('=') && c.file?.f != null);
  it('fixture is complete', () => {
    expect(typed.length).toBeGreaterThan(280);
  });
  it('toFileForm(typed) equals the stored <f> text', () => {
    const fails: string[] = [];
    let n = 0;
    for (const c of typed) {
      if (c.fn && UNKNOWN_TO_BUILD.has(c.fn)) continue;
      const got = toFileForm(c.input!, optsFor(c));
      if (LEGACY_AT.has(`${c.sheet}!${c.cell}`)) {
        // ours keeps _xlfn.SINGLE(x) where Excel wrote x
        if (got.replace(/_xlfn\.SINGLE\(((?:[^()]|\([^()]*\))*)\)/g, '$1') !== c.file!.f) fails.push(`${c.cell}: ${got} vs ${c.file!.f}`);
        continue;
      }
      n++;
      if (got !== c.file!.f) fails.push(`${c.sheet}!${c.cell} ${c.input} -> ours ${got} excel ${c.file!.f}`);
    }
    expect(fails).toEqual([]);
    expect(n).toBeGreaterThan(270);
  });
  it('toUiForm(stored) equals what Excel shows (Formula2)', () => {
    const fails: string[] = [];
    for (const c of typed) {
      if (LEGACY_AT.has(`${c.sheet}!${c.cell}`)) {
        expect(toUiForm(c.file!.f!, optsFor(c))).toBe(c.formula2!.replace('=@', '=').replace('(@', '('));
        continue;
      }
      const got = toUiForm(c.file!.f!, optsFor(c));
      if (got !== c.formula2) fails.push(`${c.sheet}!${c.cell} ${c.file!.f} -> ours ${got} excel ${c.formula2}`);
    }
    expect(fails).toEqual([]);
  });
  it('every function Excel prefixed is in the lists, and no control function is', () => {
    const prefixed = new Map<string, string>();
    for (const c of X) {
      if (c.sheet !== 'F' || !c.fn || !c.file?.f) continue;
      const m = /^IFERROR\(((?:_xlfn\.)?(?:_xlws\.)?)/.exec(c.file.f);
      prefixed.set(c.fn, m ? m[1] : '');
    }
    const fails: string[] = [];
    for (const [fn, pre] of prefixed) {
      if (UNKNOWN_TO_BUILD.has(fn)) continue;
      const want = pre === '_xlfn._xlws.' ? 'xlws' : pre === '_xlfn.' ? 'xlfn' : 'none';
      const have = XLWS_FUNCTIONS.has(fn) ? 'xlws' : FUTURE_FUNCTIONS.has(fn) ? 'xlfn' : 'none';
      if (want !== have) fails.push(`${fn}: excel ${want} ours ${have}`);
    }
    expect(fails).toEqual([]);
    expect(prefixed.size).toBeGreaterThan(240);
  });
});

// Second probe (fixtures/formula/com/{make-forms-extra-input.mjs,excel-forms-extra.ps1}): sheet X with table TblSales
// (D1:G5, totals row, calculated column G), names nmSpill / nmRange / nmRate: unknown functions, eta-reduced lambdas
// (_xleta.), '@' on scalars and names, table specials, Excel's case normalisation.
const XX = readFixture<{ cells: ExcelCell[] }>('forms-extra-excel-results.json').cells;
const X_OPTS: FormOptions = { tables: ['TblSales'], names: ['nmSpill', 'nmRange', 'nmRate'], sheets: ['X'] };
const optsX = (c: ExcelCell): FormOptions => ({ ...X_OPTS, hostTable: /^(G\d+|F5)$/.test(c.cell) ? 'TblSales' : undefined });
// '@' dropped by Excel where legacy implicit intersection applies anyway (DEVIATIONS.md "Formula forms")
const LEGACY_AT_X = new Set(['X!A5', 'X!A11', 'X!A16']);

describe('forms vs real Excel, second probe (eta lambdas, case, names, table specials)', () => {
  it('fixture is complete', () => {
    expect(XX.filter((c) => c.file?.f != null).length).toBeGreaterThanOrEqual(70);
    expect(XX.filter((c) => c.err).map((c) => c.cell)).toEqual([]);
  });
  it('toFileForm(typed) equals the stored <f> text', () => {
    const fails: string[] = [];
    for (const c of XX) {
      if (!c.input || c.file?.f == null) continue;
      const id = `${c.sheet}!${c.cell}`;
      let got = toFileForm(c.input, optsX(c));
      if (LEGACY_AT_X.has(id)) got = got.replace(/_xlfn\.SINGLE\(((?:[^()]|\([^()]*\))*)\)/g, '$1');
      // table COLUMN names keep the typed case (Excel matches them case-insensitively and re-renders them on load)
      const same = id === 'X!A49' ? got.toLowerCase() === c.file.f.toLowerCase() : got === c.file.f;
      if (!same) fails.push(`${id} ${c.input} -> ours ${got} excel ${c.file.f}`);
    }
    expect(fails).toEqual([]);
  });
  it('toUiForm(stored) equals what Excel shows (Formula2)', () => {
    const fails: string[] = [];
    for (const c of XX) {
      if (c.file?.f == null) continue;
      const id = `${c.sheet}!${c.cell}`;
      const want = LEGACY_AT_X.has(id) ? c.formula2!.replace(/@/g, '') : c.formula2;
      const got = toUiForm(c.file.f, optsX(c));
      if (got !== want) fails.push(`${id} ${c.file.f} -> ours ${got} excel ${c.formula2}`);
    }
    expect(fails).toEqual([]);
  });
});

describe('round trip on every Excel-written formula', () => {
  const files: string[] = [];
  for (const c of X) if (c.file?.f && !(c.fn && UNKNOWN_TO_BUILD.has(c.fn))) files.push(c.file.f);
  for (const c of XX) if (c.file?.f) files.push(c.file.f);
  const r2b = readFixture<{ cases: Array<{ kind: string; fileBefore?: unknown; fileAfter?: unknown }> }>('r2b-fixtures.json');
  for (const c of r2b.cases) for (const f of [c.fileBefore, c.fileAfter]) {
    const s = typeof f === 'string' ? f : (f as { f?: string } | null)?.f;
    if (s) files.push(s);
  }
  it(`toFileForm(toUiForm(f)) === f for ${files.length} formulas`, () => {
    const fails: string[] = [];
    const seen = new Set<string>();
    for (const f of files) {
      if (seen.has(f)) continue;
      seen.add(f);
      const o: FormOptions = { tables: ['Tbl', 'tblMain', 'Table1', 'TblSales'], names: ['nmSpill', 'nmRange', 'nmRate'] };
      const ui = toUiForm(f, o);
      const back = toFileForm(ui, o);
      if (back !== f) fails.push(`${f} -> ${ui} -> ${back}`);
    }
    expect(fails.slice(0, 20)).toEqual([]);
    expect(seen.size).toBeGreaterThan(1000);
  });
});

describe('toFileForm / toUiForm unit cases', () => {
  const cases: Array<[string, string, FormOptions?]> = [
    // [UI, FILE, opts]
    ['=SUM(A1:B2)', 'SUM(A1:B2)'],
    ['=xlookup(1,A:A,B:B)', '_xlfn.XLOOKUP(1,A:A,B:B)'],
    ['=CONCAT("_xlfn.X","@A1","A1#")', '_xlfn.CONCAT("_xlfn.X","@A1","A1#")'],
    ['=SORT(FILTER(A1:A9,B1:B9>0))', '_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A1:A9,B1:B9>0))'],
    ['=SUM(A1#)+Sheet2!B3#', 'SUM(_xlfn.ANCHORARRAY(A1))+_xlfn.ANCHORARRAY(Sheet2!B3)'],
    ["=ROWS('集計 シート'!$C$5#)", "ROWS(_xlfn.ANCHORARRAY('集計 シート'!$C$5))"],
    ['=SUM(Data!#REF!#)', 'SUM(_xlfn.ANCHORARRAY(Data!#REF!))'],
    ['=SUM(@A1:A3)', 'SUM(_xlfn.SINGLE(A1:A3))'],
    ['=@INDEX(A:A,1)+1', '_xlfn.SINGLE(INDEX(A:A,1))+1'],
    ['=@(A1:A3+1)', '_xlfn.SINGLE((A1:A3+1))'],
    ['=@A1#', '_xlfn.SINGLE(_xlfn.ANCHORARRAY(A1))'],
    ['=LET(x,1,y,x+1,x*y)', '_xlfn.LET(_xlpm.x,1,_xlpm.y,_xlpm.x+1,_xlpm.x*_xlpm.y)'],
    ['=LET(x, 1, x + 1)', '_xlfn.LET(_xlpm.x, 1, _xlpm.x + 1)'],
    ['=LET(x,1,LET(x,x+1,x))+x', '_xlfn.LET(_xlpm.x,1,_xlfn.LET(_xlpm.x,_xlpm.x+1,_xlpm.x))+x'],
    ['=LET(f,LAMBDA(n,n+1),f(2))', '_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.n,_xlpm.n+1),_xlpm.f(2))'],
    ['=LAMBDA(a,[b],IF(ISOMITTED(b),a,a+b))(1)', '_xlfn.LAMBDA(_xlpm.a,_xlop.b,IF(_xlfn.ISOMITTED(_xlpm.b),_xlpm.a,_xlpm.a+_xlpm.b))(1)'],
    ['=MAP(A1:A3,LAMBDA(v,v*TAX))', '_xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*TAX))'],
    ['=LET(値,5,値*2)', '_xlfn.LET(_xlpm.値,5,_xlpm.値*2)'],
    ['=LET(x,A1:A3,SUM(@x))', '_xlfn.LET(_xlpm.x,A1:A3,SUM(_xlfn.SINGLE(_xlpm.x)))'],
    ['=LET(x,"x",x&"x")', '_xlfn.LET(_xlpm.x,"x",_xlpm.x&"x")'],
    ['=Tbl[@a]', 'Tbl[[#This Row],[a]]'],
    ['=Tbl[@[b c]]', 'Tbl[[#This Row],[b c]]'],
    ["=Tbl[@[d'#e]]", "Tbl[[#This Row],[d'#e]]"],
    ['=SUM(Tbl[@[a]:[b c]])', 'SUM(Tbl[[#This Row],[a]:[b c]])'],
    ['=SUM(Tbl[@])', 'SUM(Tbl[#This Row])'],
    ['=SUM(Tbl[a])+SUM(Tbl[[#Headers],[a]])', 'SUM(Tbl[a])+SUM(Tbl[[#Headers],[a]])'],
    ['=[@単価]*2', 'Tbl[[#This Row],[単価]]*2', { hostTable: 'Tbl' }],
    ['=SUM([単価])', 'SUM(Tbl[単価])', { hostTable: 'Tbl' }],
    ['=SUM(Tbl)', 'SUM(Tbl[])', { tables: ['Tbl'] }],
    ['=Other[@a]', 'Other[[#This Row],[a]]', { hostTable: 'Tbl' }],
    ["='[Book.xlsx]S'!A1+[1]S!B2", "'[Book.xlsx]S'!A1+[1]S!B2"],
    ['=IFERROR(1/0,#N/A)', 'IFERROR(1/0,#N/A)'],
    ['={1,2;3,4}', '{1,2;3,4}'],
    ['=ECMA.CEILING(1,1)+WORKDAY.INTL(1,1)', 'ECMA.CEILING(1,1)+WORKDAY.INTL(1,1)'],
    ['=_xludf.MYFN(1)', '_xludf.MYFN(1)'],
    // eta-reduced lambdas
    ['=GROUPBY(A2:A9,C2:C9,SUM)', '_xlfn.GROUPBY(A2:A9,C2:C9,_xleta.SUM)'],
    ['=GROUPBY(A2:A9,C2:C9,HSTACK(SUM,PERCENTOF))', '_xlfn.GROUPBY(A2:A9,C2:C9,_xlfn.HSTACK(_xleta.SUM,_xleta.PERCENTOF))'],
    ['=PIVOTBY(A2:A9,B2:B9,C2:C9,AVERAGE)', '_xlfn.PIVOTBY(A2:A9,B2:B9,C2:C9,_xleta.AVERAGE)'],
    ['=BYROW(A1:C3,STDEV.S)', '_xlfn.BYROW(A1:C3,_xleta.STDEV.S)'],
    ['=MAP(A1:A3,B1:B3,MAX)', '_xlfn.MAP(A1:A3,B1:B3,_xleta.MAX)'],
    ['=MAKEARRAY(2,2,SUM)', '_xlfn.MAKEARRAY(2,2,_xleta.SUM)'],
    ['=LET(f,SUM,f(A1:A3))', '_xlfn.LET(_xlpm.f,_xleta.SUM,_xlpm.f(A1:A3))'],
    ['=LET(p,PRICE,p*2)', '_xlfn.LET(_xlpm.p,PRICE,_xlpm.p*2)'],          // not called: PRICE stays a (defined) name
    ['=MAP(A1:A3,Price)', '_xlfn.MAP(A1:A3,Price)', { names: ['Price'] }], // a defined name is never an eta lambda
    ['=INDEX(PRICE,2)', 'INDEX(PRICE,2)'],                                  // not a LAMBDA position
    ['=BYROW(A1:C3,LAMBDA(r,SUM(r)))', '_xlfn.BYROW(A1:C3,_xlfn.LAMBDA(_xlpm.r,SUM(_xlpm.r)))'],
    // defined-case names, tables, sheets; LET parameters take their declared case
    ['=SUM(nmRange)*nmRate', 'SUM(nmRange)*nmRate', { names: ['nmRange', 'nmRate'] }],
    ['=SUM(TblSales[[#Totals],[a]])', 'SUM(TblSales[[#Totals],[a]])', { tables: ['TblSales'] }],
    ["=SUM('集計 シート'!B1:B3)", "SUM('集計 シート'!B1:B3)", { sheets: ['集計 シート'] }],
    ['=LET(Xy,1,Xy+1)', '_xlfn.LET(_xlpm.Xy,1,_xlpm.Xy+1)'],
    ['=LET(z99,1,z99+1)', '_xlfn.LET(_xlpm.z99,1,_xlpm.z99+1)'],
  ];
  it('Excel case normalisation (verified: X!A45..A63)', () => {
    const o: FormOptions = { tables: ['TblSales'], names: ['nmRange'], sheets: ['X', 'データ 表'] };
    expect(toFileForm('=sum(b1:b3)+sum($b$1:c$3)+sum(b:b)', o)).toBe('SUM(B1:B3)+SUM($B$1:C$3)+SUM(B:B)');
    expect(toFileForm('=iferror(#n/a,true)*false', o)).toBe('IFERROR(#N/A,TRUE)*FALSE');
    expect(toFileForm('=SUM(x!b1:b3)+x!#ref!', o)).toBe('SUM(X!B1:B3)+X!#REF!');
    expect(toFileForm("=sum('データ 表'!a1)", o)).toBe("SUM('データ 表'!A1)");
    expect(toFileForm('=nmrange+COUNTA(tblsales)', o)).toBe('nmRange+COUNTA(TblSales[])');
    expect(toFileForm('=tblsales[[#this row],[a]]+TBLSALES[#totals]', o)).toBe('TblSales[[#This Row],[a]]+TblSales[#Totals]');
    expect(toFileForm('=xlookup(1,b1:b3,c1:c3)', o)).toBe('_xlfn.XLOOKUP(1,B1:B3,C1:C3)');
    expect(toFileForm('=LET(Xy,1,xy+XY)', o)).toBe('_xlfn.LET(_xlpm.Xy,1,_xlpm.Xy+_xlpm.Xy)');
    expect(toFileForm('=BYCOL(B1:C3,counta)', o)).toBe('_xlfn.BYCOL(B1:C3,_xleta.COUNTA)');
    expect(toFileForm('=[1]Sheet1!a1+[1]!myFn(1)', o)).toBe('[1]Sheet1!A1+[1]!myFn(1)');
    expect(toFileForm('="sum(a1)"&"#n/a"', o)).toBe('"sum(a1)"&"#n/a"');
  });
  it('UI -> FILE', () => {
    for (const [ui, file, o] of cases) expect(toFileForm(ui, o), ui).toBe(file);
  });
  it('FILE -> UI (exact inverse)', () => {
    const inverse = cases.filter(([ui]) => !/LET\(x, 1|xlookup/.test(ui));
    for (const [ui, file, o] of inverse) expect(toUiForm(file, o), file).toBe(ui);
    expect(toUiForm('_xlfn.XLOOKUP(1,A:A,B:B)')).toBe('=XLOOKUP(1,A:A,B:B)');
  });
  it('is idempotent on file form and never touches strings', () => {
    for (const [, file, o] of cases) expect(toFileForm(file, o)).toBe(file);
    expect(toFileForm('="LET(x,1,x)"&"@A1"')).toBe('"LET(x,1,x)"&"@A1"');
  });
  it('keeps unbalanced / junk input lossless', () => {
    for (const s of ['=SUM(A1', '=LET(x,1', '=@', '=A1#+', '=Tbl[@', '="unterminated', '=)(', '=LAMBDA(']) {
      const f = toFileForm(s);
      expect(tokenize(f).map((t) => t.text).join('')).toBe(f);
      expect(() => toUiForm(f)).not.toThrow();
    }
  });
  it('fileFunctionName', () => {
    expect(fileFunctionName('textjoin')).toBe('_xlfn.TEXTJOIN');
    expect(fileFunctionName('filter')).toBe('_xlfn._xlws.FILTER');
    expect(fileFunctionName('Sum')).toBe('SUM'); // Excel upper-cases every function name it stores (X!A45)
    expect(fileFunctionName('foobar')).toBe('FOOBAR'); // unknown functions too (X!A2)
    expect(fileFunctionName('_XLFN.xlookup')).toBe('_xlfn.XLOOKUP');
    expect(fileFunctionName('_xlfn._xlws.sort')).toBe('_xlfn._xlws.SORT');
    expect(fileFunctionName('_xludf.MyFn')).toBe('_xludf.MyFn');
    expect(fileFunctionName('_xlfn.NEWFN')).toBe('_xlfn.NEWFN'); // a typed _xlfn. of an unknown function is kept
  });
});

describe('robustness (random junk never throws, output stays tokenizable)', () => {
  const parts = ['=', 'LET(', 'LAMBDA(', 'x', 'y', ',', ')', '(', '@', '#', 'A1', '$B$2:C9', '[', ']', '"', "'", 'Tbl[@a]',
    '[@[b c]]', '_xlfn.', '_xlpm.x', '_xlop.b', '[b]', '{', '}', ' ', '集計!', "'集計 シート'!", 'SUM(', 'XLOOKUP(', 'Z1#',
    '_xlfn.SINGLE(', '_xlfn.ANCHORARRAY(', 'Tbl[]', '#REF!', '1', '+', ':', 'Tbl', '[1]S!A1', 'INDIRECT("A1")', '　'];
  let seed = 12345;
  const rnd = (n: number) => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed % n; };
  it('20,000 random formulas', () => {
    for (let k = 0; k < 20000; k++) {
      let f = '';
      const len = 1 + rnd(12);
      for (let j = 0; j < len; j++) f += parts[rnd(parts.length)];
      const opts = { hostTable: rnd(2) ? 'Tbl' : undefined, tables: ['Tbl'] };
      const file = toFileForm(f, opts);
      expect(tokenize(file).map((t) => t.text).join('')).toBe(file);
      const ui = toUiForm(file, opts);
      expect(ui.startsWith('=')).toBe(true);
      expect(() => rewriteRefs(file, 'S', (r) => ({ area: r.area }), { keepNames: true, keepStructured: true, allowExternal: true, allowIndirect: true })).not.toThrow();
    }
  });
});

describe('isFormulaInput', () => {
  it('Excel-like', () => {
    expect(isFormulaInput('=A1')).toBe(true);
    expect(isFormulaInput('= 1')).toBe(true);
    expect(isFormulaInput('=')).toBe(false);
    expect(isFormulaInput('=  ')).toBe(false);
    expect(isFormulaInput('A1')).toBe(false);
    expect(isFormulaInput('＝A1')).toBe(false);
    expect(isFormulaInput("'=A1")).toBe(false);
  });
});

describe('rewriteRefs', () => {
  // premerge-style map: same-row refs of source sheet "Src" columns B..D -> merged columns K..M, row 7 -> 12
  const colMap: Record<number, number> = { 2: 11, 3: 12, 4: 13 };
  const premerge = (f: string, host = 'Src') => rewriteRefs(f, host, (r) => {
    if (r.sheet !== 'Src' || r.area.kind !== 'cell' || r.area.r1 !== 7) return null;
    const c = colMap[r.area.c1];
    if (!c) return null;
    return { area: { ...r.area, c1: c, c2: c, r1: 12, r2: 12 }, prefixText: '' };
  });
  const premergeRows = (f: string) => rewriteRefs(f, 'Src', (r) => {
    const c1 = colMap[r.area.c1], c2 = colMap[r.area.c2];
    if (r.sheet !== 'Src' || r.area.r1 !== 7 || r.area.r2 !== 7 || !c1 || !c2) return null;
    return { area: { ...r.area, c1, c2, r1: 12, r2: 12 }, prefixText: '' };
  });
  it('rewrites same-row refs and keeps everything else verbatim', () => {
    expect(premerge('B7*C7+"B7"')).toBe('K12*L12+"B7"');
    expect(premerge('ROUND($B7*1.1,0)')).toBe('ROUND($K12*1.1,0)');
    expect(premerge('Src!D7&"円"')).toBe('M13&"円"'.replace('M13', 'M12'));
    expect(premerge('_xlfn.LET(_xlpm.x,B7,_xlpm.x*2)')).toBe('_xlfn.LET(_xlpm.x,K12,_xlpm.x*2)');
  });
  it('returns null for anything unmappable', () => {
    expect(premerge('B7+E7')).toBeNull();          // column not imported
    expect(premerge('B8')).toBeNull();             // other row
    expect(premerge('SUM(B7:C7)')).toBeNull();     // area (map refuses)
    expect(premerge('B7*TAX_RATE')).toBeNull();    // defined name of the source book
    expect(premerge('Tbl[@a]*B7')).toBeNull();     // structured ref
    expect(premerge('[1]Src!B7')).toBeNull();      // external
    expect(premerge('INDIRECT("B7")')).toBeNull(); // text reference
    expect(premerge('Src!#REF!+B7')).toBeNull();
    expect(premerge('"unterminated')).toBeNull();
    // position-dependent functions would silently change value in another row / column
    expect(premerge('ROW()-1')).toBeNull();
    expect(premerge('COLUMN(B7)')).toBeNull();
    expect(premerge('OFFSET(B7,0,1)')).toBeNull();
    expect(premerge('CELL("address",B7)')).toBeNull();
    expect(premerge('_xlfn.SHEET()')).toBeNull();
    expect(premerge('_xlfn.BYROW(B7:D7,_xleta.ROW)')).toBeNull();
  });
  it('options relax the defaults; unchanged refs keep their text', () => {
    expect(rewriteRefs('a1+TAX', 'S', (r) => ({ area: r.area }), { keepNames: true })).toBe('a1+TAX');
    expect(rewriteRefs('SUM(Tbl[a])+A1', 'S', (r) => ({ area: parseAreaText('B2')!, prefixText: r.prefix?.text }), { keepStructured: true })).toBe('SUM(Tbl[a])+B2');
    expect(rewriteRefs('#REF!+1', 'S', () => null)).toBe('#REF!+1');
    expect(rewriteRefs('ROW(A1)+B2', 'S', (r) => ({ area: r.area }), { allowPositional: true })).toBe('ROW(A1)+B2');
    expect(rewriteRefs('_xlfn.BYROW(A1:B2,_xleta.SUM)', 'S', (r) => ({ area: r.area }))).toBe('_xlfn.BYROW(A1:B2,_xleta.SUM)');
    expect(premergeRows('ROWS(B7:D7)*C7')).toBe('ROWS(K12:M12)*L12');
  });
});

describe('forms vs the R2b Excel ground truth (Formula2 <-> saved <f> of ~17,500 formulas, before and after each op)', () => {
  interface R2bFormula { id: string; kind: string; where: string; cellBefore?: string; cellAfter?: string | null;
    before?: string; after?: string | null; before2?: string; after2?: string; fileBefore?: unknown; fileAfter?: unknown; opId: string }
  const fx = readFixture<{ suites: Record<string, { tables: Array<{ name: string; sheet: string; ref: string }> }>;
    ops: Array<{ opId: string; suite: string; error: string | null; noTruth: boolean }>; cases: R2bFormula[] }>('r2b-fixtures.json');
  const ops = new Map(fx.ops.map((o) => [o.opId, o]));
  const fileText = (f: unknown) => (typeof f === 'string' ? f : (f as { f?: string | null } | null | undefined)?.f ?? null);
  // table ranges after each op (the 'table' cases: Excel's <table ref>)
  const tableAfter = new Map<string, string>();
  for (const c of fx.cases as Array<R2bFormula & { before?: any; after?: any }>) {
    if (c.kind === 'table' && c.after?.ref) tableAfter.set(`${c.opId}|${c.where.replace(/^table:/, '')}`, c.after.ref);
  }
  /** the table that contains the host cell (UI form omits its name in [@col]) */
  const hostTableOf = (opId: string, suite: string, where: string, cell: string | null | undefined, after: boolean): string | undefined => {
    if (!cell) return undefined;
    const a = parseAreaText(cell.split(':')[0]);
    for (const t of fx.suites[suite]?.tables ?? []) {
      const r = parseAreaText(after ? tableAfter.get(`${opId}|${t.name}`) ?? t.ref : t.ref);
      if (a && r && t.sheet === where && a.r1 >= r.r1 && a.r1 <= r.r2 && a.c1 >= r.c1 && a.c1 <= r.c2) return t.name;
    }
    return undefined;
  };
  const pairs: Array<{ id: string; ui: string; file: string; o: FormOptions }> = [];
  for (const c of fx.cases) {
    if (c.kind !== 'formula') continue;
    const op = ops.get(c.opId)!;
    if (op.noTruth) continue;
    const tables = (fx.suites[op.suite]?.tables ?? []).map((t) => t.name);
    for (const [ui, file, cell, after] of [[c.before2 ?? c.before, c.fileBefore, c.cellBefore, false], [c.after2 ?? c.after, c.fileAfter, c.cellAfter, true]] as const) {
      const f = fileText(file);
      if (typeof ui !== 'string' || f == null) continue;
      pairs.push({ id: c.id, ui, file: f, o: { tables, hostTable: hostTableOf(c.opId, op.suite, c.where, cell, after && !op.error) } });
    }
  }
  // Legacy implicit intersection: Excel shows '@' for a stored bare range where legacy evaluation intersects
  // (DEVIATIONS.md "Formula forms"); toUiForm gives the bare range, toFileForm(=@x) gives _xlfn.SINGLE(x).
  const legacyAt = (p: { ui: string; file: string }) =>
    tokenize(p.ui.slice(1)).some((t) => t.type === 'op' && t.text === '@') && !p.file.includes('_xlfn.SINGLE');
  it(`toUiForm(<f>) === Formula2 and toFileForm(Formula2) === <f> (${pairs.length} pairs)`, () => {
    expect(pairs.length).toBeGreaterThan(30000);
    const fails: string[] = [];
    let legacy = 0;
    for (const p of pairs) {
      const ui = toUiForm(p.file, p.o), file = toFileForm(p.ui, p.o);
      if (legacyAt(p)) {
        legacy++;
        if (ui !== p.ui.replace(/(^=|[(,])@/g, '$1')) fails.push(`${p.id} legacy-@ UI ${p.file} -> ours ${ui} excel ${p.ui}`);
        continue;
      }
      if (ui !== p.ui) fails.push(`${p.id} UI ${p.file} -> ours ${ui} excel ${p.ui}`);
      if (file !== p.file) fails.push(`${p.id} FILE ${p.ui} -> ours ${file} excel ${p.file}`);
    }
    expect(fails.slice(0, 30)).toEqual([]);
    if (process.env.FORMS_REPORT) console.log(`pairs ${pairs.length}, legacy-@ ${legacy}`);
    expect(legacy).toBeLessThan(pairs.length / 50);
  });
});

describe('forms vs real Excel, third probe (2024-2026 functions: IMPORTTEXT / IMPORTCSV, REGEX*, GROUPBY, ...)', () => {
  const cells = readFixture<{ cells: ExcelCell[] }>('forms-new-excel-results.json').cells.filter((c) => c.input && c.file?.f != null);
  it('toFileForm(typed) equals the stored <f> text and toUiForm(<f>) equals Formula2', () => {
    expect(cells.length).toBeGreaterThanOrEqual(18);
    const fails: string[] = [];
    for (const c of cells) {
      const f = toFileForm(c.input!), u = toUiForm(c.file!.f!);
      if (c.cell === 'J61') { // legacy '@' (DEVIATIONS.md "Formula forms"): Excel stores SUM(B1:B3)+B1:B3
        expect(f).toBe('SUM(B1:B3)+_xlfn.SINGLE(B1:B3)');
        continue;
      }
      if (f !== c.file!.f) fails.push(`${c.cell} ${c.input} -> ours ${f} excel ${c.file!.f}`);
      if (u !== c.formula2) fails.push(`${c.cell} ${c.file!.f} -> ours ${u} excel ${c.formula2}`);
    }
    expect(fails).toEqual([]);
  });
});

describe('COM write-back rule: toUiForm(<f>) is Range.Formula of a normal cell, Range.Formula2 of an array / dynamic cell', () => {
  it('on every Excel-written cell of the forms fixtures and every R2b formula', () => {
    const fails: string[] = [];
    let n = 0;
    for (const fixture of ['forms-excel-results.json', 'forms-extra-excel-results.json', 'forms-new-excel-results.json', 'forms-tro-excel-results.json', 'forms-dyn-excel-results.json']) {
      for (const c of readFixture<{ cells: ExcelCell[] }>(fixture).cells) {
        if (!c.file?.f) continue;
        const inTable = (c.sheet === 'T' && /^E\d/.test(c.cell)) || (c.sheet === 'X' && /^[D-G]\d$/.test(c.cell));
        const ui = toUiForm(c.file.f, { hostTable: inTable ? (c.sheet === 'T' ? 'Tbl' : 'TblSales') : undefined });
        const want = c.file.t === 'array' || c.file.cm ? c.formula2 : c.formula;
        n++;
        if (ui !== want) fails.push(`${fixture} ${c.sheet}!${c.cell} ${c.file.f} -> ${ui} want ${want}`);
      }
    }
    const r2b = readFixture<{ cases: Array<{ id: string; kind: string; where: string; before?: unknown; after?: unknown; before2?: string; after2?: string; fileBefore?: unknown; fileAfter?: unknown }> }>('r2b-fixtures.json');
    for (const c of r2b.cases) {
      if (c.kind !== 'formula') continue;
      for (const [f, F, F2] of [[c.fileBefore, c.before, c.before2], [c.fileAfter, c.after, c.after2]] as const) {
        if (f == null || typeof F !== 'string') continue;
        const arr = typeof f === 'object';
        const text = arr ? (f as { f: string | null }).f : (f as string);
        if (text == null) continue;
        const want = arr ? F2 ?? F : F;
        if (c.where === 'Data' && /\[@/.test(want)) continue; // calculated-column cells: UI form omits the table name
        n++;
        const ui = toUiForm(text, { tables: ['tblMain'] });
        if (ui !== want) fails.push(`${c.id} ${text} -> ${ui} want ${want}`);
      }
    }
    expect(fails.slice(0, 20)).toEqual([]);
    expect(n).toBeGreaterThan(35000);
  });
});
