// quoteSheetName vs the quoting Excel 365 produced (fixtures/formula/storage-results.json), renameSheetInFormula,
// "delete would create #REF!" detection, listRefs, adjustSqref, sheet-edge rules.
import { describe, expect, it } from 'vitest';
import {
  quoteSheetName, renameSheetInFormula, adjustFormula, tokenize, listRefs, adjustSqref, adjustArea, parseAreaText,
  renderArea, mapIndex, colToNum, numToCol, parsePrefix,
} from '../../src/core/formula/excelRef';
import { readFixture } from './helpers';

const S = readFixture<{ quoting: Record<string, string> }>('storage-results.json');

describe('quoteSheetName vs Excel', () => {
  it('quotes everything Excel quotes and round-trips through the tokenizer', () => {
    const fails: string[] = [];
    for (const [name, excelFormula] of Object.entries(S.quoting)) {
      const excelQuoted = excelFormula.startsWith("='");
      const q = quoteSheetName(name);
      if (excelQuoted && !q.startsWith("'")) fails.push(`must quote ${name} (Excel does): got ${q}`);
      const t = tokenize(`${q}!B2`);
      if (!(t.length === 1 && t[0].type === 'ref' && t[0].prefix!.sheet === name)) fails.push(`round trip ${q}!B2 -> ${JSON.stringify(t.map((x) => [x.type, x.text]))}`);
    }
    expect(fails).toEqual([]);
  });
});

describe('renameSheetInFormula', () => {
  it('renames prefixes only', () => {
    expect(renameSheetInFormula("SUM(集計!A1:B2)+'集計'!C3+Other!D4+\"集計!A1\"", '集計', '集計 2024')).toBe("SUM('集計 2024'!A1:B2)+'集計 2024'!C3+Other!D4+\"集計!A1\"");
    expect(renameSheetInFormula('SUM(Sheet1:Sheet3!A1)', 'Sheet3', 'Last')).toBe("SUM('Sheet1:Last'!A1)");
    expect(renameSheetInFormula('[1]Sheet1!A1+Sheet1!A1', 'Sheet1', 'S')).toBe('[1]Sheet1!A1+S!A1');
  });
});

describe('#REF! dry run', () => {
  const del = { type: 'delete', axis: 'col', at: 3, count: 1 } as const;
  it('detects refs that a delete would break', () => {
    expect(adjustFormula('Data!C1+1', del, { hostSheet: 'F', targetSheet: 'Data' }).createdRefError).toBe(true);
    expect(adjustFormula('SUM(Data!B1:D1)', del, { hostSheet: 'F', targetSheet: 'Data' }).createdRefError).toBe(false);
    expect(adjustFormula('SUM(Data!C1:C9)', del, { hostSheet: 'F', targetSheet: 'Data' }).formula).toBe('SUM(Data!#REF!)');
    expect(adjustFormula("'データ 表'!C1", del, { hostSheet: 'F', targetSheet: 'データ 表' }).formula).toBe("'データ 表'!#REF!");
    expect(adjustFormula('C1*2', del, { hostSheet: 'Data', targetSheet: 'Data' }).formula).toBe('#REF!*2');
  });
  it('never adjusts other sheets, 3D or external refs', () => {
    expect(adjustFormula('C1+Other!C1+Data:Mid!C1+[1]Data!C1', del, { hostSheet: 'Other', targetSheet: 'Data' }).changed).toBe(false);
  });
  it('flags unknown tokens as unsafe', () => {
    expect(adjustFormula('SUM(A1', del, { hostSheet: 'Data', targetSheet: 'Data' }).unsafe).toBe(false);
    expect(adjustFormula('"unterminated', del, { hostSheet: 'Data', targetSheet: 'Data' }).unsafe).toBe(true);
  });
});

describe('listRefs / areas', () => {
  it('lists refs with resolved sheets', () => {
    const r = listRefs("A1+'集計 シート'!B2:C3+[1]X!A1+S1:S2!B1", 'Host');
    expect(r.map((x) => [x.text, x.sheet, x.external, x.threeD])).toEqual([
      ['A1', 'Host', false, false],
      ["'集計 シート'!B2:C3", '集計 シート', false, false],
      ['[1]X!A1', 'X', true, false],
      ['S1:S2!B1', null, false, true],
    ]);
  });
  it('parses and renders areas', () => {
    for (const s of ['A1', '$A$1', 'A1:B2', '$C:$E', '3:5', 'XFD1048576', 'A1:XFD1']) expect(renderArea(parseAreaText(s)!)).toBe(s);
    expect(parseAreaText('XFE1')).toBeNull();
    expect(parseAreaText('A0')).toBeNull();
    expect(colToNum('XFD')).toBe(16384);
    expect(numToCol(16384)).toBe('XFD');
    expect(parsePrefix("'O''Brien'!")).toEqual({ text: "'O''Brien'!", sheet: "O'Brien" });
  });
  it('insert at the sheet edge: a cell pushed past XFD is #REF!, a range is clamped and normalised', () => {
    const ins = { type: 'insert', axis: 'col', at: 3, count: 1 } as const;
    expect(adjustArea(parseAreaText('XFD1')!, ins)).toBeNull();
    expect(renderArea(adjustArea(parseAreaText('A1:XFC1')!, ins)!)).toBe('1:1');
    expect(renderArea(adjustArea(parseAreaText('K5:XFD5')!, ins)!)).toBe('L5:XFD5');
    expect(mapIndex(5, { type: 'delete', axis: 'col', at: 5, count: 1 })).toBeNull();
  });
  it('adjustSqref drops fully deleted parts', () => {
    expect(adjustSqref('B2:E10 G2', { type: 'delete', axis: 'col', at: 7, count: 1 })).toBe('B2:E10');
    expect(adjustSqref('B2:E10,G2', { type: 'insert', axis: 'col', at: 3, count: 2 }, ',')).toBe('B2:G10,I2');
  });
});
