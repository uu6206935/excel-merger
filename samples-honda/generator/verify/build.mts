// ID + 部品名 の複合キーで 01 に 02〜16 を結合 → 列入れ替え → 出力（アプリ本体と同じ処理を Node で）
import fs from 'node:fs';
import path from 'node:path';
import { loadSpec } from 'file:///D:/ダウンロード/ExcelMerger/excel-merger/tests/oracle/copied-sheets/pipeline.ts';
import { analyzeSources, defaultPlan, joinSources, addManualCandidate } from 'file:///D:/ダウンロード/ExcelMerger/excel-merger/src/core/keys/index.ts';
import { createEditor, moveColumns } from 'file:///D:/ダウンロード/ExcelMerger/excel-merger/src/core/doc/index.ts';
import { readBook } from 'file:///D:/ダウンロード/ExcelMerger/excel-merger/src/core/xlsx/index.ts';
import { buildPremerged } from 'file:///D:/ダウンロード/ExcelMerger/excel-merger/src/core/xlsx/premerge.ts';
import { exportWithJs } from 'file:///D:/ダウンロード/ExcelMerger/excel-merger/src/core/export/index.ts';

const dir = 'D:/ダウンロード/ExcelMerger/honda-mock';
const outDir = path.join(dir, 'merged');
fs.mkdirSync(outDir, { recursive: true });
const mode = process.argv[2] ?? 'default'; // default = アプリの既定のまま
const names = fs.readdirSync(dir).filter((f) => /^\d\d_.*\.xlsx$/.test(f)).sort();
const base = await loadSpec({ path: path.join(dir, names[0]) }, 's1');
const others: any[] = [];
for (let i = 1; i < names.length; i++) others.push(await loadSpec({ path: path.join(dir, names[i]) }, `s${i + 1}`));

let analysis = analyzeSources(base, others);
const plan: any = { ...defaultPlan(analysis), joinType: 'left' };
const baseId = base.book.sheets.find((s: any) => s.name === base.region.sheet);
const idCol = 2, nameCol = 3; // 01 の 部品ID / 部品名（1 始まり）
for (const o of others) {
  const cands = analysis.pairs[o.id].candidates;
  const idc = cands.find((c: any) => c.kind === 'single' && c.baseColumns[0]?.col === idCol);
  const nmc = cands.find((c: any) => c.baseColumns[0]?.col === nameCol);
  if (!idc || !nmc) throw new Error(`${o.label}: ID/名前の候補が見つかりません`);
  if (false) {
    analysis = addManualCandidate(analysis, o.id, [idCol, nameCol], [idc.otherColumns[0].col, nmc.otherColumns[0].col]);
    const added = analysis.pairs[o.id].candidates.find((c: any) => c.baseColumns.length === 2);
    if (!added) throw new Error(`${o.label}: 複合キー候補を作れませんでした`);
    plan.tiers[o.id] = [added.id];
  }
}
const join = joinSources(analysis, plan);
console.log(mode, 'summary', JSON.stringify(join.summary));
const matched: Record<string, number> = {};
for (const [id, r] of Object.entries<any>(join.reports)) matched[id] = r.matched;
console.log('matched per file', JSON.stringify(matched));
console.log('conflict/ambiguous items', join.items.filter((i: any) => ['conflict', 'ambiguous', 'fallback', 'vetoed'].includes(i.code)).length);

const imports = join.defaultImports.map((d: any) => (d.header ? { sourceId: d.sourceId, col: d.col, header: d.header } : { sourceId: d.sourceId, col: d.col }));
const input: any = { base, sources: others, rows: join.rows, imports, keyFill: join.keyFill, ...(join.keyPairs ? { keyPairs: join.keyPairs } : {}),
  options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix', dropOtherSheets: false, includeSourceSheets: true } };
const result = await buildPremerged(input);
console.log('report', JSON.stringify({ merged: result.report.mergedRows, imported: result.report.importedColumns, kept: result.report.formulasKept, toValues: result.report.formulasToValues, added: result.report.addedSheets?.length }));
const tag = '';
fs.writeFileSync(path.join(outDir, `統合結果_並び替えなし${tag}.xlsx`), result.bytes);


// ---- 列の入れ替え（アプリの「列を移動」と同じ操作）----
const book = await readBook(result.bytes, base.book.fileName, { id: 'merged' });
const sourceInfo: any = {};
[base, ...others].forEach((k: any, i: number) => (sourceInfo[k.id] = { label: k.label, color: '#2563eb' }));
let last = base.region.lastDataRow;
for (const a of result.layout.appended) if (a.row > last) last = a.row;
let snap: any = createEditor(book, result.layout, sourceInfo, { headerRow: base.region.headerRow, firstDataRow: base.region.firstDataRow, lastDataRow: last });
const n0 = snap.doc.columns.length;
const hdr = (s: any) => s.doc.columns.map((c: any) => c.header ?? c.title ?? c.id);
console.log('columns before', n0, hdr(snap).slice(0, 12).join(' | '));
const move = (idx: number[], gap: number) => {
  const r = moveColumns(snap, idx.map((i) => snap.doc.columns[i].id), gap, { unmergeTitles: true });
  if (!r.ok) throw new Error('move failed: ' + r.reason);
  snap = r.snapshot;
};
move([n0 - 8, n0 - 7, n0 - 6, n0 - 5, n0 - 4, n0 - 3, n0 - 2, n0 - 1], 2); // 末尾 8 列を 部品名 の右へ
move([10, 11, 12], snap.doc.columns.length); // 中ほどの 3 列を最後へ
move([4], 20); // 搭載車種 を 20 列目あたりへ
move([30, 31], 3); // 離れた 2 列を前方へ
console.log('columns after ', snap.doc.columns.length, hdr(snap).slice(0, 14).join(' | '));
const out = await exportWithJs(result.bytes, snap.log);
fs.writeFileSync(path.join(outDir, '統合結果_列入れ替え済み.xlsx'), out);
fs.writeFileSync(path.join(outDir, '_ops.json'), JSON.stringify(snap.log, null, 1));
console.log('ops', snap.log.length, 'wrote outputs');
