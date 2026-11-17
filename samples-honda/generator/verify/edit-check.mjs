import fs from 'node:fs';
const rd = (f) => JSON.parse(fs.readFileSync(f, 'utf8').replace(/^\uFEFF/, ''));
const D = rd('dumpB2.json'); const { r, ix, before } = rd('edit-meta.json');
const row = D.values[r]; const g = (h) => row[ix[h]];
const exp = { '単価(税抜)': 50000, '消費税額': 4000, '税込単価': 54000, '申請数量': 10, '申請金額(税抜)': 10 * before['申請単価'], '申請消費税': Math.round(10 * before['申請単価'] * 0.1), '合否': before['不良数'] > 0 ? '不合格' : '合格', '差異': 20, '充足率': 1.2, '充足判定': '充足', '合計在庫': 10, '補充判定': '要補充' };
let bad = 0;
for (const [h, e] of Object.entries(exp)) { const a = g(h); const ok = typeof e === 'number' ? Math.abs(a - e) < 1e-9 : a === e; if (!ok) bad++; console.log(ok ? 'OK  ' : 'NG  ', h, 'before=', before[h], 'after=', a, 'expected=', e); }
console.log('mismatches:', bad);
