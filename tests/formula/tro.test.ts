// Trim-range references (Excel 365 "trim refs": A1:.B9 trailing, A1.:B9 leading, A1.:.B9 all), checked against real
// Excel: forms (fixtures/formula/forms-tro-excel-results.json: typed via Range.Formula2, stored <f> read from the
// saved .xlsx; the file form is _xlfn._TRO_TRAILING(A1:B9) / _TRO_LEADING / _TRO_ALL) and structural ops
// (fixtures/formula/tro-ops-excel-results.json: insert / delete / clip-free move of columns and rows; generators in
// fixtures/formula/com/excel-tro-ops.ps1 + extract-tro-ops.mjs).
import { describe, expect, it } from 'vitest';
import { adjustFormula, mapIndex, parseAreaText, renderArea, tokenize, type StructOp } from '../../src/core/formula/excelRef';
import { toFileForm, toUiForm } from '../../src/core/formula/forms';
import { readFixture } from './helpers';

interface FormsCell { cell: string; input: string | null; formula2: string | null; file: { f: string | null } | null }
const FORMS = readFixture<{ cells: FormsCell[] }>('forms-tro-excel-results.json').cells.filter((c) => c.input && c.file?.f);

describe('trim refs: UI form <-> file form (real Excel)', () => {
  it('toFileForm(typed) equals the stored <f>, toUiForm(<f>) equals Formula2', () => {
    expect(FORMS.length).toBeGreaterThanOrEqual(16);
    const fails: string[] = [];
    for (const c of FORMS) {
      const f = toFileForm(c.input!), u = toUiForm(c.file!.f!);
      if (f !== c.file!.f) fails.push(`${c.cell} FILE ${c.input} -> ours ${f} excel ${c.file!.f}`);
      if (u !== c.formula2) fails.push(`${c.cell} UI ${c.file!.f} -> ours ${u} excel ${c.formula2}`);
    }
    expect(fails).toEqual([]);
  });
  it('tokenizer: a trim range is one ref token (Area.trim), lossless, rendered back unchanged', () => {
    for (const [text, trim] of [['B1:.B10', 'trailing'], ['b1.:b10', 'leading'], ['$B$1.:.$C$3', 'all'], ['B:.D', 'trailing'], ['$1.:$3', 'leading'], ["'a b'!C.:.D", 'all']] as const) {
      const toks = tokenize(`SUM(${text})`);
      expect(toks.map((t) => t.type)).toEqual(['func', 'op', 'ref', 'op']);
      expect(toks[2].area!.trim).toBe(trim);
      expect(toks.map((t) => t.text).join('')).toBe(`SUM(${text})`);
      const body = text.slice(text.indexOf('!') + 1);
      expect(renderArea(parseAreaText(body)!)).toBe(body.toUpperCase());
    }
    // a number / name / decimal is never mistaken for a trim range
    expect(tokenize('1.5+B1.x+A1:B2').map((t) => t.type)).toEqual(['number', 'op', 'name', 'op', 'ref']);
    // #REF! keeps its marker: one error token
    for (const [text, trim] of [['Data!#REF!:.', 'trailing'], ['#REF!.:', 'leading'], ['#REF!.:.', 'all']] as const) {
      const toks = tokenize(`SUM(${text})`);
      expect(toks.map((t) => [t.type, t.text, t.trim])).toEqual([['func', 'SUM', undefined], ['op', '(', undefined], ['error', text, trim], ['op', ')', undefined]]);
    }
    expect(tokenize('#REF!:.5').map((t) => t.type)).toEqual(['error', 'op', 'number']);
  });
  it('#REF! trim markers convert both ways; odd TRO arguments round-trip', () => {
    for (const [ui, file] of [
      ['=SUM(Data!#REF!:.)', 'SUM(_xlfn._TRO_TRAILING(Data!#REF!))'],
      ['=SUM(Data!#REF!.:)', 'SUM(_xlfn._TRO_LEADING(Data!#REF!))'],
      ['=SUM(Data!#REF!.:.)', 'SUM(_xlfn._TRO_ALL(Data!#REF!))'],
      ['=SUM(#REF!:.)', 'SUM(_xlfn._TRO_TRAILING(#REF!))'],
      ['=SUM(_TRO_ALL(INDEX(B:B,1):B9))', 'SUM(_xlfn._TRO_ALL(INDEX(B:B,1):B9))'],
      ['=_TRO_TRAILING(B1)', '_xlfn._TRO_TRAILING(B1)'],
    ] as const) {
      expect(toFileForm(ui)).toBe(file);
      expect(toUiForm(file)).toBe(ui);
    }
  });
});

// ------------------------------------------------------------------------------------------------ structural ops
// the formulas typed before every op (same list as excel-tro-ops.ps1)
const BEFORE: Array<[sheet: string, cell: string, ui: string]> = [
  ['Other', 'A1', '=SUM(Data!B1:.D10)'], ['Other', 'A2', '=SUM(Data!C:.C)'], ['Other', 'A3', '=SUM(Data!B1.:.E5)'],
  ['Other', 'A4', '=ROWS(Data!$C$2:.$C$9)'], ['Other', 'A5', '=SUM(Data!B2.:F2)'], ['Other', 'A6', '=SUM(Data!C1:.C3)'],
  ['Other', 'A7', '=SUM(Data!$2:.$4)'], ['Other', 'A8', '=SUM(Data!B.:D)'], ['Data', 'K11', '=SUM(B1:.D3)'], ['Data', 'K12', '=SUM(C:.E)'],
];
const OPS: Record<string, StructOp> = {
  insC: { type: 'insert', axis: 'col', at: 3, count: 1 },
  delC: { type: 'delete', axis: 'col', at: 3, count: 1 },
  delCD: { type: 'delete', axis: 'col', at: 3, count: 2 },
  delBF: { type: 'delete', axis: 'col', at: 2, count: 5 },
  movDtoB: { type: 'move', axis: 'col', from: 4, count: 1, to: 2 },
  movBtoF: { type: 'move', axis: 'col', from: 2, count: 1, to: 6 },
  insRow3: { type: 'insert', axis: 'row', at: 3, count: 1 },
  delRow2: { type: 'delete', axis: 'row', at: 2, count: 1 },
  delRows1to10: { type: 'delete', axis: 'row', at: 1, count: 10 },
};
interface OpCell { op: string; sheet: string; cell: string; formula2: string; file: string | null }
const AFTER = readFixture<{ cells: OpCell[] }>('tro-ops-excel-results.json').cells;

/** Excel renders an UNPREFIXED trimmed range that became #REF! as "=#REF! SUM()" (Formula2 display quirk; the stored
 * <f> is SUM(_xlfn._TRO_TRAILING(#REF!))). Our UI form is "=SUM(#REF!:.)", the exact inverse of that <f>. */
const displayQuirk = (c: OpCell) => /^=#REF! /.test(c.formula2);

describe('trim refs under structural ops (real Excel)', () => {
  it('fixture covers every op', () => {
    expect(new Set(AFTER.map((c) => c.op))).toEqual(new Set(Object.keys(OPS)));
  });
  for (const [name, op] of Object.entries(OPS)) {
    it(`${name}: file form, UI form and form conversion equal Excel`, () => {
      const fails: string[] = [];
      for (const [sheet, cell, ui] of BEFORE) {
        const a = parseAreaText(cell)!;
        let at = cell;
        if (sheet === 'Data') {
          const v = mapIndex(op.axis === 'col' ? a.c1 : a.r1, op);
          at = renderArea(op.axis === 'col' ? { ...a, c1: v!, c2: v! } : { ...a, r1: v!, r2: v! });
        }
        const x = AFTER.find((c) => c.op === name && c.sheet === sheet && c.cell === at);
        if (!x) { fails.push(`${sheet}!${cell}: no Excel cell at ${at}`); continue; }
        const ctx = { hostSheet: sheet, targetSheet: 'Data' };
        const file = adjustFormula(toFileForm(ui), op, ctx).formula;
        if (file !== x.file) fails.push(`${sheet}!${cell} FILE ${toFileForm(ui)} -> ours ${file} excel ${x.file}`);
        if (toFileForm(x.formula2) !== x.file && !displayQuirk(x)) fails.push(`${sheet}!${cell} toFileForm(${x.formula2}) -> ${toFileForm(x.formula2)} excel ${x.file}`);
        const uiAfter = '=' + adjustFormula(ui.slice(1), op, ctx).formula;
        const want = displayQuirk(x) ? toUiForm(x.file!) : x.formula2;
        if (uiAfter !== want) fails.push(`${sheet}!${cell} UI ${ui} -> ours ${uiAfter} excel ${x.formula2}`);
        if (toUiForm(x.file!) !== want) fails.push(`${sheet}!${cell} toUiForm(${x.file}) -> ${toUiForm(x.file!)} excel ${x.formula2}`);
      }
      expect(fails).toEqual([]);
    });
  }
  it('createdRefError is reported for a trimmed range that is deleted', () => {
    const r = adjustFormula('SUM(Data!C:.C)', OPS.delC, { hostSheet: 'Other', targetSheet: 'Data' });
    expect(r).toMatchObject({ formula: 'SUM(Data!#REF!:.)', createdRefError: true, changed: true, unsafe: false });
    // the #REF! token with its marker is stable under further ops
    expect(adjustFormula(r.formula, OPS.insC, { hostSheet: 'Other', targetSheet: 'Data' })).toMatchObject({ formula: 'SUM(Data!#REF!:.)', changed: false, unsafe: false });
  });
});
