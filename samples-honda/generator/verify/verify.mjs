import fs from 'node:fs';
const rd = (f) => JSON.parse(fs.readFileSync(f, 'utf8').replace(/^\uFEFF/, ''));
const truth = rd('D:/ダウンロード/ExcelMerger/honda-mock/truth.json');
const A = rd('dumpA.json'), B = rd('dumpB.json');
let fails = 0; const fail = (m) => { fails++; if (fails <= 40) console.log('FAIL', m); };
const blank = (v) => v === null || v === undefined || v === '';
const same = (a, b) => (typeof a === 'number' && typeof b === 'number') ? Math.abs(a - b) < 1e-9 : String(a ?? '') === String(b ?? '');
function table(D) {
  const hdr = D.values[0]; const idx = {}; hdr.forEach((h, i) => { if (h != null) idx[h] = i; });
  return { hdr, idx, rows: D.values.slice(1) };
}
function col(T, h) {
  if (h in T.idx) return T.idx[h];
  const k = Object.keys(T.idx).filter((x) => x.startsWith(h + '（'));
  return k.length === 1 ? T.idx[k[0]] : undefined;
}
function check(D, label) {
  const T = table(D); let okCells = 0, blankCells = 0, rejectChecked = 0;
  const idc = T.idx['部品ID'], nmc = T.idx['部品名'];
  const rowOf = new Map(T.rows.map((r, i) => [r[idc] + '|' + r[nmc], i]));
  if (T.rows.length !== truth.master.length) fail(`${label}: rows ${T.rows.length} != 44`);
  for (const f of truth.files.slice(1)) {
    const sh = f.sheets.find((s) => s.isData);
    const keyH = f.keyHeaders; const dataH = sh.headers.filter((h) => !keyH.includes(h));
    const present = new Set();
    for (const tr of sh.rows) {
      const ri = rowOf.get(tr[keyH[0]] + '|' + tr[keyH[1]]);
      if (ri === undefined) { fail(`${label} ${f.file}: master row for ${tr[keyH[0]]} not found`); continue; }
      present.add(ri);
      for (const h of dataH) {
        const c = col(T, h); if (c === undefined) { fail(`${label} ${f.file}: header ${h} missing`); continue; }
        const exp = tr[h]; const act = T.rows[ri][c];
        if (!(blank(exp) && blank(act)) && !same(exp, act)) fail(`${label} ${f.file} ${tr[keyH[0]]} ${h}: expected ${exp} got ${act}`); else okCells++;
      }
    }
    // 一致しない行（rejects）: 取り込まれていない = その部品の欄が空、かつ ID/名前の列に出てこない
    T.rows.forEach((r, ri) => { if (present.has(ri)) return; for (const h of dataH) { const c = col(T, h); if (!blank(r[c])) fail(`${label} ${f.file} row ${ri + 2} ${h}: should be blank but is ${r[c]}`); else blankCells++; } });
    for (const rj of sh.rejects) {
      rejectChecked++;
      if (rowOf.has(rj[keyH[0]] + '|' + rj[keyH[1]])) fail(`${label} ${f.file}: reject ${rj[keyH[0]]} ${rj[keyH[1]]} became a merged row`);
    }
  }
  console.log(`${label}: values ok=${okCells}, blank-as-expected=${blankCells}, rejects excluded=${rejectChecked}`);
  return T;
}
const TA = check(A, '並び替えなし'); const TB = check(B, '列入れ替え済み');
// シート
const srcSheets = truth.files.flatMap((f) => f.sheets.map((s) => ({ file: f.file.replace(/\.xlsx$/, ''), sheet: s.name })));
console.log('sheets in output:', A.sheets.length, 'expected', srcSheets.length);
for (const s of srcSheets.slice(2)) { if (!A.sheets.some((n) => n.includes(s.sheet))) fail(`sheet ${s.file}/${s.sheet} missing`); }
console.log(A.sheets.join(' , '));
// B vs A
const ha = Object.keys(TA.idx), hb = Object.keys(TB.idx);
console.log('same header set:', ha.length === hb.length && ha.every((h) => h in TB.idx), '| column order differs:', ha.some((h, i) => TB.idx[h] !== i));
let diff = 0; for (let r = 0; r < TA.rows.length; r++) for (const h of ha) { if (!same(TA.rows[r][TA.idx[h]], TB.rows[r][TB.idx[h]])) { diff++; if (diff < 5) console.log('DIFF', r, h, TA.rows[r][TA.idx[h]], TB.rows[r][TB.idx[h]]); } }
console.log('values differing between A and B (by header):', diff);
console.log('formula counts A', JSON.stringify(A.formulaCounts).slice(0, 200));
const fa = Object.values(A.formulaCounts).reduce((a, b) => a + b, 0), fb = Object.values(B.formulaCounts).reduce((a, b) => a + b, 0);
console.log('formula cells total A', fa, 'B', fb);
const errs = [...A.values.flat(), ...B.values.flat()].filter((v) => typeof v === 'number' && v < -2146826000);
console.log('error cells in merged sheet:', errs.length, ' FAILS:', fails);
