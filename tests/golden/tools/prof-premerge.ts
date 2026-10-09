// Dev profiler for buildPremerged at 10,000 x 40.
import { buildPremerged } from '../../../src/core/xlsx/premerge';
import { buildXlsx } from '../build';
import { source } from '../premerge-scenario';
const colL = (n: number) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
async function main() {
  const rows = 10000, cols = 40;
  const hdr = Array.from({ length: cols + 1 }, (_x, c) => `<c r="${colL(c + 1)}1" t="inlineStr"><is><t>H${c}</t></is></c>`).join('');
  const data = (n: number, withId: boolean) => {
    let x = `<row r="1">${hdr}</row>`;
    for (let r = 2; r <= n + 1; r++) {
      x += `<row r="${r}"><c r="A${r}"><v>${r}</v></c>`;
      if (withId) for (let c = 2; c <= cols + 1; c++) x += `<c r="${colL(c)}${r}" s="1"><v>${r * c}</v></c>`;
      x += '</row>';
    }
    return x;
  };
  let t = Date.now();
  const baseBytes = await buildXlsx([{ name: 'B', data: data(rows, false) }]);
  const srcBytes = await buildXlsx([{ name: 'S', data: data(rows, true) }]);
  console.log('build fixtures', Date.now() - t); t = Date.now();
  const r = await buildPremerged({
    base: source('b', 'b.xlsx', { sheet: 'B', headerRow: 1, firstDataRow: 2, lastDataRow: rows + 1, firstCol: 1, lastCol: 1 }, 'b', 'full', baseBytes),
    sources: [source('s', 's.xlsx', { sheet: 'S', headerRow: 1, firstDataRow: 2, lastDataRow: rows + 1, firstCol: 1, lastCol: cols + 1 }, 's', 'full', srcBytes)],
    rows: Array.from({ length: rows }, (_x, i) => ({ baseRow: i + 2, matches: { s: { row: rows + 1 - i, via: 'ID', flags: [] } } })),
    imports: Array.from({ length: cols }, (_x, i) => ({ sourceId: 's', col: i + 2 })),
    keyFill: {},
    options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix' },
  }, (p) => console.log(Date.now() - t, p.message));
  console.log('premerge', Date.now() - t, r.bytes.length);
}
main();
