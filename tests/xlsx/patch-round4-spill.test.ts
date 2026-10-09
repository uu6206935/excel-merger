/**
 * A value / formula written INTO the spill range of a dynamic array (not at its master) must not leave an array range over a
 * cell that holds something else: Excel repairs such a file ("cell information removed", found by the random op-log fuzz,
 * tests/golden/fuzz.ts fuzzOps seed 0 on torture.xlsx). The spill collapses like when a structural op cuts through it:
 * the master's ref shrinks to the master cell, the cached spilled values go, Excel re-spills on load (#SPILL! while the
 * cell is occupied, like in Excel).
 * Fixture r2-arrays.xlsx: A!D2 = SEQUENCE(3,2) (D2:E4), A!G2 = A2:A4&"x" (G2:G4), A!B2:B4 legacy CSE array, B!A2 =
 * ANCHORARRAY(A!G2) (A2:A4 of sheet B).
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import type { CellWrite, LoggedOp } from '../../src/core/types';
import { replayLog } from '../../src/core/xlsx/replay';
import { packageProblems, text, unpack } from '../golden/package-tools';

const here = path.dirname(fileURLToPath(import.meta.url));
const fixture = () => new Uint8Array(fs.readFileSync(path.join(here, '../fixtures/golden/r2-arrays.xlsx')));
const cells = (sheet: string, writes: CellWrite[], reason: 'edit' | 'toValues' = 'edit'): LoggedOp => ({ kind: 'cells', sheet, writes, reason });
const w = (row: number, col: number, cell: CellWrite['cell']): CellWrite => ({ row, col, cell });
const insC = (sheet: string, at: number): LoggedOp => ({ kind: 'struct', sheet, op: { type: 'insert', axis: 'col', at, count: 1 } });

/** the <c> element of `ref` in sheet `n` (1-based) of a package */
async function cellXml(bytes: Uint8Array, n: number, ref: string): Promise<string> {
  const xml = text(await unpack(bytes), `xl/worksheets/sheet${n}.xml`);
  return new RegExp(`<c r="${ref}"[^>]*?(?:/>|>[\\s\\S]*?</c>)`).exec(xml)?.[0] ?? '(none)';
}

describe('a write into the spill range of a dynamic array collapses the spill', () => {
  it('a number inside D2:E4 (E3): the master keeps only its own cell, the cached spill values are gone, other spills stay', async () => {
    const out = await replayLog(fixture(), [cells('A', [w(3, 5, { t: 'n', v: 7 })])]);
    expect(await packageProblems(out)).toEqual([]);
    expect(await cellXml(out, 1, 'D2')).toMatch(/cm="1"><f t="array" ref="D2">_xlfn\.SEQUENCE\(3,2\)<\/f>/);
    expect(await cellXml(out, 1, 'E3')).toBe('<c r="E3"><v>7</v></c>');
    for (const ref of ['D3', 'D4', 'E2', 'E4']) expect(await cellXml(out, 1, ref), ref).toMatch(/^<c r="[DE]\d"\/>$/);
    // the other dynamic array (G2:G4) and the legacy array (B2:B4) are untouched
    expect(await cellXml(out, 1, 'G2')).toContain('ref="G2:G4"');
    expect(await cellXml(out, 1, 'G3')).toContain('<v>2x</v>');
    expect(await cellXml(out, 1, 'B2')).toContain('ref="B2:B4"');
  });

  it('a formula inside G2:G4 (G3): same, and the sheet that refers to the spill is not touched', async () => {
    const out = await replayLog(fixture(), [cells('A', [w(3, 7, { t: 'z', v: null, f: '1+1' })])]);
    expect(await packageProblems(out)).toEqual([]);
    expect(await cellXml(out, 1, 'G2')).toContain('ref="G2"');
    expect(await cellXml(out, 1, 'G3')).toBe('<c r="G3"><f>1+1</f></c>');
    expect(await cellXml(out, 1, 'G4')).toBe('<c r="G4"/>');
    expect(await cellXml(out, 1, 'D2')).toContain('ref="D2:E4"');
    expect(await cellXml(out, 2, 'A2')).toContain('ref="A2:A4"');
  });

  it('a write outside every spill range and a clear of a spilled cell leave the spill alone', async () => {
    const out = await replayLog(fixture(), [cells('A', [w(6, 5, { t: 'n', v: 7 }), w(3, 5, null)])]);
    expect(await cellXml(out, 1, 'D2')).toContain('ref="D2:E4"');
    expect(await cellXml(out, 1, 'D3')).toContain('<v>3</v>');
    expect(await cellXml(out, 1, 'E3')).toBe('<c r="E3"/>');
  });

  it('the range is followed through earlier ops: a column inserted in front, then a write into the shifted spill', async () => {
    // D2:E4 is now E2:F4; F3 lies inside it
    const out = await replayLog(fixture(), [insC('A', 1), cells('A', [w(3, 6, { t: 's', v: '文字' })])]);
    expect(await packageProblems(out)).toEqual([]);
    expect(await cellXml(out, 1, 'E2')).toContain('ref="E2"');
    expect(await cellXml(out, 1, 'F3')).toMatch(/^<c r="F3" t="s"><v>\d+<\/v><\/c>$/);
    expect(await cellXml(out, 1, 'E3')).toBe('<c r="E3"/>');
    // the formulas that refer to the spill follow the insert as usual
    expect(await cellXml(out, 1, 'L3')).toContain('ANCHORARRAY(E2)');
  });

  it('several writes into the same spill in one batch, and into two spills: each spill collapses once', async () => {
    const out = await replayLog(fixture(), [cells('A', [w(3, 5, { t: 'n', v: 1 }), w(4, 4, { t: 'n', v: 2 }), w(4, 7, { t: 'n', v: 3 })])]);
    expect(await packageProblems(out)).toEqual([]);
    expect(await cellXml(out, 1, 'D2')).toContain('ref="D2"');
    expect(await cellXml(out, 1, 'G2')).toContain('ref="G2"');
    expect(await cellXml(out, 1, 'E3')).toBe('<c r="E3"><v>1</v></c>');
    expect(await cellXml(out, 1, 'D4')).toBe('<c r="D4"><v>2</v></c>');
    expect(await cellXml(out, 1, 'G4')).toBe('<c r="G4"><v>3</v></c>');
  });

  it('the master itself: a value replaces the whole spill (unchanged behaviour); toValues freezes it', async () => {
    const over = await replayLog(fixture(), [cells('A', [w(2, 4, { t: 'n', v: 5 })])]);
    expect(await cellXml(over, 1, 'D2')).toBe('<c r="D2"><v>5</v></c>');
    expect(await cellXml(over, 1, 'D3')).toBe('<c r="D3"/>');
    const frozen = await replayLog(fixture(), [cells('A', [w(2, 4, { t: 'n', v: 1 })], 'toValues')]);
    expect(await cellXml(frozen, 1, 'D2')).toBe('<c r="D2"><v>1</v></c>');
    expect(await cellXml(frozen, 1, 'E4')).toContain('<v>6</v>'); // the other cells keep their values as constants
  });
});
