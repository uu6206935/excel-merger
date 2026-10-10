/**
 * Speed and size of the sheet copier. Always run: 3 sheets x 5,000 rows x 30 columns (a few seconds) with the same checks as the big run.
 * ADDSHEETS_PERF=1 runs the task's workload: 10 sheets x 20,000 rows x 30 columns (6,000,000 cells, 250 MB of sheet XML) — buildPremerged must
 * stay under 20 s (measured about 16 s on the development machine: native deflate, one-match well-formedness check, fast cell parser).
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { buildPremerged } from '../../src/core/xlsx';
import { KITCHEN_STYLES } from './addsheets-kitchen';
import { buildPackage, cn, row, type SheetSpec } from './addsheets-build';
import { keySource } from './addsheets-helpers';

async function source(sheets: number, rows: number, cols: number): Promise<Uint8Array> {
  const specs: SheetSpec[] = [];
  for (let s = 0; s < sheets; s++) {
    const parts: string[] = [];
    for (let r = 1; r <= rows; r++) {
      let cells = '';
      for (let c = 1; c <= cols; c++) {
        const ref = cn(c) + r;
        const k = (r * 31 + c * 7 + s) % 10;
        if (k < 4) cells += `<c r="${ref}" s="${(c % 5) + 1}" t="s"><v>${(r * c) % 5000}</v></c>`;
        else if (k < 8) cells += `<c r="${ref}" s="${(c % 5) + 1}"><v>${(r * c) / 7}</v></c>`;
        else if (k === 8) cells += `<c r="${ref}" s="2"><f>SUM(A${r}:B${r})</f><v>${r * 2}</v></c>`;
        else cells += `<c r="${ref}" s="3"/>`;
      }
      parts.push(row(r, cells));
    }
    specs.push({ name: `S${s + 1}`, data: parts.join('') });
  }
  return buildPackage({ sheets: specs, styles: KITCHEN_STYLES, sst: Array.from({ length: 5000 }, (_v, i) => `text ${i} あいう`) });
}

async function measure(sheets: number, rows: number, cols: number) {
  const src = await source(sheets, rows, cols);
  const base = await buildPackage({ sheets: [{ name: '名簿', data: row(1, '<c r="A1" t="inlineStr"><is><t>ID</t></is></c>') + row(2, '<c r="A2" t="inlineStr"><is><t>x</t></is></c>') }] });
  const b = await keySource({ bytes: base, fileName: '名簿.xlsx', region: { sheet: '名簿', headerRow: 1, firstDataRow: 2, lastDataRow: 2, firstCol: 1, lastCol: 1 } }, 'b');
  const o = await keySource({ bytes: src, fileName: 'big.xlsx', sheet: 'S1' }, 's2');
  const t = Date.now();
  const res = await buildPremerged({
    base: b, sources: [o], rows: [{ baseRow: 2, matches: {} }], imports: [], keyFill: {},
    options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix', includeSourceSheets: true, dropOtherSheets: true },
  });
  const ms = Date.now() - t;
  // every part reads back (CRC checked) and every copied sheet has all its rows
  const zip = await JSZip.loadAsync(res.bytes, { checkCRC32: true });
  const names = Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
  expect(names.length).toBe(1 + sheets);
  for (const n of names.slice(1)) {
    const x = await zip.file(n)!.async('string');
    expect((x.match(/<row /g) ?? []).length).toBe(rows);
    expect(x).toContain(`<dimension ref="A1:AD${rows}"/>`);
  }
  expect(res.layout.addedSheets!.length).toBe(sheets);
  return ms;
}

describe('sheet copier: speed', () => {
  it('3 sheets x 5,000 rows x 30 columns: copied completely, well under the budget', async () => {
    const ms = await measure(3, 5000, 30);
    expect(ms).toBeLessThan(15_000);
  }, 120_000);

  it.skipIf(process.env.ADDSHEETS_PERF !== '1')('10 sheets x 20,000 rows x 30 columns (6,000,000 cells) in under 20 s', async () => {
    const ms = await measure(10, 20000, 30);
    console.log(`buildPremerged with 10 copied sheets of 20000x30: ${ms} ms`);
    expect(ms).toBeLessThan(20_000);
  }, 900_000);
});
