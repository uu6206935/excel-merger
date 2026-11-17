import fs from 'node:fs';
const rd = (f) => JSON.parse(fs.readFileSync(f, 'utf8').replace(/^\uFEFF/, ''));
const B = rd('dumpB.json');
const L = (i) => { let s = ''; for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
const hdr = B.values[0]; const ix = Object.fromEntries(hdr.map((h, i) => [h, i]));
// 経費・品質・発注の 3 つが全部入っている行を探す
const r = B.values.findIndex((row, i) => i > 0 && row[ix['申請数量']] != null && row[ix['検査数']] != null && row[ix['要望個数']] != null && row[ix['不良数']] > 0);
const a = (h) => L(ix[h]) + (r + 1);
const first = B.sheets[0];
const edits = [
  { sheet: first, addr: a('単価(税抜)'), value: 50000 },
  { sheet: first, addr: a('申請数量'), value: 10 },
  { sheet: first, addr: a('要望個数'), value: 100 },
  { sheet: first, addr: a('実際個数'), value: 120 },
  { sheet: first, addr: a('社内在庫'), value: 7 },
  { sheet: first, addr: a('サプライヤー現在在庫'), value: 3 },
  { sheet: first, addr: a('安全在庫'), value: 50 },
  { sheet: '設定', addr: 'B2', value: 0.08 },
  { sheet: '03_品質チェック_判定基準', addr: 'B2', value: 0 },
];
fs.writeFileSync('edits.json', JSON.stringify(edits), 'utf8');
fs.writeFileSync('edit-meta.json', JSON.stringify({ r, ix, before: Object.fromEntries(['部品ID','単価(税抜)','消費税額','税込単価','申請数量','申請単価','申請金額(税抜)','申請消費税','不良数','検査数','合否','差異','充足率','充足判定','合計在庫','補充判定'].map((h) => [h, B.values[r][ix[h]]])) }), 'utf8');
console.log('row', r + 1, JSON.stringify(JSON.parse(fs.readFileSync('edit-meta.json', 'utf8')).before));
