import fs from 'node:fs'; import { createRequire } from 'node:module';
const XLSX = createRequire('D:/ダウンロード/ExcelMerger/excel-merger/package.json')('xlsx');
const dir = 'D:/ダウンロード/ExcelMerger/honda-mock/';
const dump = JSON.parse(fs.readFileSync('recalc-src.json','utf8').replace(/^\uFEFF/,''));
let bad = 0; const cache = {};
for (const d of dump) {
  const wb = (cache[d.file] ??= XLSX.readFile(dir + d.file));
  const c = wb.Sheets[d.sheet][XLSX.utils.encode_cell({ r: d.row - 1, c: d.col - 1 })];
  const mine = c?.v ?? '';
  const x = d.v ?? '';
  const same = typeof mine === 'number' && typeof x === 'number' ? Math.abs(mine - x) < 1e-9 : String(mine) === String(x);
  if (!same) { bad++; if (bad <= 0) console.log('DIFF', d.file, d.sheet, XLSX.utils.encode_cell({ r: d.row - 1, c: d.col - 1 }), d.f, 'cached=', mine, 'excel=', x); }
}
console.log(`${dump.length} formula cells checked, ${bad} differences`);
const by = {}; for (const d of dump) { const wb = cache[d.file]; const c = wb.Sheets[d.sheet][XLSX.utils.encode_cell({ r: d.row - 1, c: d.col - 1 })]; const mine = c?.v ?? ''; const x = d.v ?? ''; const same = typeof mine === 'number' && typeof x === 'number' ? Math.abs(mine - x) < 1e-9 : String(mine) === String(x); if (!same) { const k = d.file + '|' + d.sheet + '|' + d.f.replace(/\d+/g,'#'); (by[k] ??= []).push(`${XLSX.utils.encode_cell({ r: d.row - 1, c: d.col - 1 })} cached=${mine} excel=${x}`); } }
for (const [k, v] of Object.entries(by)) console.log(k, v.length, v[0]);
const wb = cache['01_部品マスタ.xlsx']; console.log(JSON.stringify(wb.Sheets['設定']), JSON.stringify(wb.Sheets['部品マスタ']['F2']), JSON.stringify(wb.Sheets['部品マスタ']['G2']));
