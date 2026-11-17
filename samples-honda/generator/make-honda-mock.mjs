#!/usr/bin/env node
// Honda 向けモックデータ生成: 部品ID + 部品名 をキーにして Excel 統合ツールで結合できる 16 ファイルを作る。
//   node make-honda-mock.mjs [出力フォルダ]     （既定: このフォルダの親 = honda-mock）
// - 部品名は実在する「○○ASSY」形式の名称（部品ID・金額・在庫などの数値、サプライヤー名、URL は全て架空）
// - 数式セルは「式 + 計算済みの値」の両方を書く（Excel で開き直して再計算しても同じ値になることを verify で確認する）
// - 各ファイルに「ID または部品名が一致しない行」(rejects) を混ぜる。正しい行（rows）だけが結合される想定。
// - truth.json に期待値を書き出す（verify-merge.ts が突き合わせる）
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const outDir = path.resolve(process.argv[2] ?? path.join(here, '..'));
const require = createRequire('D:/ダウンロード/ExcelMerger/excel-merger/package.json');
const XLSX = require('xlsx');

// ---------- 乱数・小物 ----------
function mulberry32(a) {
  return () => {
    a |= 0; a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const rnd = mulberry32(20261117);
const ri = (a, b) => a + Math.floor(rnd() * (b - a + 1));
const pick = (arr) => arr[Math.floor(rnd() * arr.length)];
const shuffle = (arr) => { const a = [...arr]; for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(rnd() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; };
const serial = (y, m, d) => Math.round(Date.UTC(y, m - 1, d) / 86400000) + 25569;
const colLetter = (i) => { let s = ''; for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
const round = (x, d = 0) => { const k = 10 ** d; return Math.round(Number((x * k).toPrecision(15))) / k; };

// ---------- 部品マスタ（実在する ASSY 名称 / ID は Honda 風の架空採番） ----------
const VEHICLES = [
  ['N-BOX', 'TTA'], ['N-ONE', 'THA'], ['N-WGN', 'TFA'], ['フィット', 'T5A'], ['ヴェゼル', 'T7A'], ['フリード', 'T9A'],
  ['ステップワゴン', 'TAA'], ['シビック', 'TBA'], ['ZR-V', 'TCA'], ['CR-V', 'TDA'], ['アコード', 'TEA'], ['WR-V', 'TGA'],
];
// [グループ番号, 部品名, 部品区分, 税抜単価の目安]
const PART_DEFS = [
  ['19010', 'ラジエーターASSY', '冷却', 38500], ['19200', 'ウォーターポンプASSY', '冷却', 21800], ['31100', 'オルタネーターASSY', '電装', 52400],
  ['31200', 'スターターモーターASSY', '電装', 41700], ['30500', 'イグニッションコイルASSY', '電装', 9800], ['16400', 'スロットルボディASSY', 'エンジン', 33600],
  ['17100', 'インテークマニホールドASSY', 'エンジン', 47200], ['18100', 'エキゾーストマニホールドASSY', 'エンジン', 58900], ['18900', 'ターボチャージャーASSY', 'エンジン', 148000],
  ['17045', '燃料ポンプASSY', '燃料', 27300], ['17500', '燃料タンクASSY', '燃料', 36800], ['20000', 'CVTトランスミッションASSY', '駆動', 412000],
  ['44306', 'ドライブシャフトASSY', '駆動', 46500], ['44300', 'ハブベアリングASSY', '足回り', 14200], ['51600', 'フロントストラットASSY', '足回り', 39600],
  ['52610', 'リアダンパーASSY', '足回り', 18700], ['51350', 'フロントロアアームASSY', '足回り', 22400], ['53620', 'ステアリングギヤボックスASSY', 'ステアリング', 96800],
  ['53200', 'ステアリングコラムASSY', 'ステアリング', 41200], ['78501', 'ステアリングホイールASSY', 'ステアリング', 28600], ['46100', 'ブレーキマスターシリンダーASSY', 'ブレーキ', 31900],
  ['45019', 'フロントブレーキキャリパーASSY', 'ブレーキ', 24800], ['57110', 'ABSモジュレーターASSY', 'ブレーキ', 118000], ['47200', 'パーキングブレーキレバーASSY', 'ブレーキ', 8900],
  ['38810', 'エアコンコンプレッサーASSY', '空調', 68900], ['80100', 'コンデンサーASSY', '空調', 19400], ['79105', 'ヒーターユニットASSY', '空調', 54300],
  ['78100', 'メーターASSY', '電装', 43800], ['77100', 'インストルメントパネルASSY', '内装', 88700], ['81100', 'フロントシートASSY', '内装', 79500],
  ['82100', 'リアシートASSY', '内装', 63400], ['04810', 'フロントシートベルトASSY', '安全', 16800], ['77800', '運転席エアバッグASSY', '安全', 72300],
  ['33100', 'ヘッドライトASSY', '外装', 61800], ['33500', 'テールランプASSY', '外装', 22900], ['71100', 'フロントバンパーASSY', '外装', 34700],
  ['71501', 'リアバンパーASSY', '外装', 29800], ['67050', 'フロントドアASSY', '外装', 84600], ['76200', 'ドアミラーASSY', '外装', 26400],
  ['60100', 'ボンネットASSY', '外装', 49200], ['76505', 'ワイパーモーターASSY', '電装', 11800], ['72210', 'ウインドウレギュレーターASSY', '内装', 13600],
  ['70100', 'サンルーフASSY', '外装', 128000], ['68100', 'バックドアASSY', '外装', 92400],
];
const SEQ = 'ABCDEFGHJK';
const PARTS = PART_DEFS.map(([grp, name, cat, price], i) => {
  const [veh, vcode] = VEHICLES[i % VEHICLES.length];
  const id = `${grp}-${vcode}-${SEQ[i % SEQ.length]}${String(10 + ((i * 7) % 89)).padStart(2, '0')}`;
  return { id, name, cat, veh, price: Math.round(price * (0.95 + rnd() * 0.1) / 100) * 100 };
});
// マスタにない部品（廃番・他社・誤入力の想定）
const UNKNOWN = [
  { id: '46400-TTA-Z01', name: 'ブレーキブースターASSY' }, { id: '19101-T5A-Z02', name: 'オイルクーラーASSY' },
  { id: '39100-TBA-Z03', name: 'ナビゲーションユニットASSY' }, { id: '74100-THA-Z04', name: 'フロアカーペットASSY' },
  { id: '36000-TCA-Z05', name: 'バックカメラASSY' }, { id: '12000-TEA-Z06', name: 'シリンダーヘッドASSY' },
];
const byId = new Map(PARTS.map((p) => [p.id, p]));

// ---------- シート組み立て ----------
const dateCell = (y, m, d) => serial(y, m, d);
function build(fileDef) {
  const wb = XLSX.utils.book_new();
  const truthSheets = [];
  for (const sh of fileDef.sheets) {
    const startRow = sh.title ? 1 : 0; // 見出し行 (0-based row index)
    const cols = sh.cols;
    const colIndex = Object.fromEntries(cols.map((c, i) => [c.h, i]));
    const ws = {};
    const put = (r, c, cell) => { ws[XLSX.utils.encode_cell({ r, c })] = cell; };
    if (sh.title) put(0, 0, { t: 's', v: sh.title });
    cols.forEach((c, i) => put(startRow, i, { t: 's', v: c.h }));
    const rowsAll = [...sh.rows.map((r) => ({ ...r, __rej: false })), ...(sh.rejects ?? []).map((r) => ({ ...r, __rej: true }))];
    // 並び順は混ぜる（先頭と末尾に偏らないように）
    const ordered = sh.keepOrder ? rowsAll : shuffleStable(rowsAll);
    ordered.forEach((row, k) => {
      const r = startRow + 1 + k;
      for (const c of cols) {
        const i = colIndex[c.h];
        if (c.calc) row[c.h] = c.calc(row);
        const v = row[c.h];
        if ((v === undefined || v === null || v === '') && !c.f) continue;
        const cell = typeof v === 'number' ? { t: 'n', v } : { t: 's', v: String(v ?? '') };
        if (c.f) cell.f = c.f.replace(/\{([^}]+)\}/g, (_, h) => `${colLetter(colIndex[h])}${r + 1}`);
        if (c.z) cell.z = c.z;
        put(r, i, cell);
      }
    });
    const lastRow = startRow + ordered.length;
    for (const x of sh.extra ?? []) put(x.r, x.c, x.cell);
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(lastRow, ...(sh.extra ?? []).map((x) => x.r)), c: cols.length - 1 } });
    ws['!cols'] = cols.map((c) => ({ wch: c.w ?? Math.max(10, c.h.length * 2.2) }));
    if (sh.title) ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: cols.length - 1 } }];
    XLSX.utils.book_append_sheet(wb, ws, sh.name);
    truthSheets.push({
      name: sh.name, isData: !!sh.data, headerRow: startRow + 1, headers: cols.map((c) => c.h),
      rows: ordered.filter((r) => !r.__rej).map((r) => Object.fromEntries(cols.map((c) => [c.h, r[c.h] ?? null]))),
      rejects: ordered.filter((r) => r.__rej).map((r) => ({ ...Object.fromEntries(cols.map((c) => [c.h, r[c.h] ?? null])), __reason: r.__reason, __intended: r.__intended ?? null })),
    });
  }
  XLSX.writeFile(wb, path.join(outDir, fileDef.file), { bookType: 'xlsx', cellStyles: true });
  return { file: fileDef.file, keyHeaders: fileDef.keyHeaders, sheets: truthSheets, note: fileDef.note };
}
function shuffleStable(rows) { return shuffle(rows); }

// ---------- 行の作り方 ----------
/** マスタの一部を選び（coverage）、ID/名前が一致しない行を作る。 */
function selectParts(coverage, nRejects) {
  const picked = shuffle(PARTS).slice(0, Math.round(PARTS.length * coverage));
  const rest = PARTS.filter((p) => !picked.includes(p));
  const pickedSet = new Set(picked.map((p) => p.id));
  const rejects = [];
  const kinds = ['typo-id', 'wrong-name', 'unknown', 'typo-id', 'wrong-name', 'unknown'];
  const used = [...picked];
  for (let k = 0; k < nRejects; k++) {
    const kind = kinds[k % kinds.length];
    if (kind === 'unknown') { rejects.push({ id: UNKNOWN[(k + used.length) % UNKNOWN.length].id, name: UNKNOWN[(k + used.length) % UNKNOWN.length].name, kind, intended: null }); continue; }
    // 正しい行がない部品を 1 つ選んでキーを壊す（その部品の欄は結合後に空欄になるはず）
    const target = rest.shift();
    if (!target) continue;
    if (kind === 'typo-id') {
      const ch = target.id.slice(-1);
      const bad = target.id.slice(0, -1) + (ch === '9' ? '8' : String(Number(ch) + 1 > 9 ? 0 : Number(ch) + 1));
      rejects.push({ id: bad, name: target.name, kind, intended: target });
    } else {
      const other = PARTS.find((p) => p.cat !== target.cat && !pickedSet.has(p.id) && p !== target) ?? PARTS[0];
      rejects.push({ id: target.id, name: other.name, kind, intended: target });
    }
  }
  return { picked, rejects };
}
const REASON = { 'typo-id': 'ID が 1 桁違う（名前は一致）', 'wrong-name': 'ID は一致・部品名が別物', unknown: 'マスタに無い部品' };
function mk(def) {
  const { picked, rejects } = selectParts(def.coverage, def.rejects ?? 3);
  const rows = picked.map((p, i) => ({ ...def.row(p, i, false), [def.idH]: p.id, [def.nameH]: p.name }));
  const rj = rejects.map((r, i) => ({
    ...def.row(r.intended ?? { id: r.id, name: r.name, price: 10000, cat: '-', veh: '-' }, 100 + i, true),
    [def.idH]: r.id, [def.nameH]: r.name, __reason: REASON[r.kind], __intended: r.intended ? r.intended.id : null,
  }));
  return { rows, rejects: rj };
}

const IDN = (idH, nameH) => [{ h: idH, w: 18 }, { h: nameH, w: 34 }];
const files = [];

// 01 部品マスタ（基準・2 シート。税額は「設定」シートの税率を参照）
{
  const rows = PARTS.map((p, i) => ({ No: i + 1, 部品ID: p.id, 部品名: p.name, 部品区分: p.cat, 搭載車種: p.veh, '単価(税抜)': p.price }));
  files.push({
    file: '01_部品マスタ.xlsx', keyHeaders: ['部品ID', '部品名'], note: '基準ファイル。税額は「設定」シートの税率を参照する数式（2 シート）',
    sheets: [
      { name: '部品マスタ', data: true, rows, keepOrder: true, cols: [
        { h: 'No', w: 6 }, { h: '部品ID', w: 18 }, { h: '部品名', w: 34 }, { h: '部品区分', w: 12 }, { h: '搭載車種', w: 14 },
        { h: '単価(税抜)', z: '#,##0', w: 14 },
        { h: '消費税額', f: 'ROUND({単価(税抜)}*設定!$B$2,0)', calc: (r) => Math.round(r['単価(税抜)'] * 0.1), z: '#,##0', w: 12 },
        { h: '税込単価', f: '{単価(税抜)}+{消費税額}', calc: (r) => r['単価(税抜)'] + Math.round(r['単価(税抜)'] * 0.1), z: '#,##0', w: 14 },
      ] },
      { name: '設定', rows: [{ 項目: '消費税率', 値: 0.1 }], keepOrder: true, cols: [{ h: '項目' }, { h: '値', z: '0%' }] },
    ],
  });
  // 設定シート: 見出し行のすぐ下ではなく B2 が税率になるように（見出し=行1, データ=行2）
}

// 02 経費申請（2 シート・先頭にタイトル行）
{
  const KINDS = [['E01', '試作費'], ['E02', '金型償却費'], ['E03', '輸送費'], ['E04', '試験評価費'], ['E05', '補修・手直し費'], ['E06', '出張・立会費']];
  const STATES = ['申請中', '承認済', '承認済', '承認済', '差戻し', '却下', '支払済'];
  const { rows, rejects } = mk({
    coverage: 0.82, rejects: 4, idH: '部品ID', nameH: '部品名',
    row: (p) => { const k = pick(KINDS); const qty = ri(1, 40); const price = Math.round(p.price * (0.2 + rnd() * 0.5)); return { 経費種別コード: k[0], 申請数量: qty, 申請単価: price, 申請状況: pick(STATES), 申請日: dateCell(2026, ri(8, 11), ri(1, 28)) }; },
  });
  const kindsRows = KINDS.map(([c, n], i) => ({ 種別コード: c, 種別名: n, 月次上限額: [800000, 1500000, 400000, 600000, 300000, 200000][i] }));
  files.push({
    file: '02_経費申請.xlsx', keyHeaders: ['部品ID', '部品名'], note: 'タイトル行あり（見出しは 2 行目）。VLOOKUP で 2 枚目のシートを引く。2 シート',
    sheets: [
      { name: '経費申請', data: true, title: '経費申請一覧（2026年度 下期）', rows, rejects, cols: [
        ...IDN('部品ID', '部品名'), { h: '経費種別コード', w: 14 },
        { h: '経費の種類', f: 'VLOOKUP({経費種別コード},経費種別!$A$2:$B$7,2,FALSE)', calc: (r) => (KINDS.find((k) => k[0] === r['経費種別コード']) ?? [, '不明'])[1], w: 16 },
        { h: '申請数量', w: 10 }, { h: '申請単価', z: '#,##0', w: 12 },
        { h: '申請金額(税抜)', f: '{申請数量}*{申請単価}', calc: (r) => r['申請数量'] * r['申請単価'], z: '#,##0', w: 16 },
        { h: '申請消費税', f: 'ROUND({申請金額(税抜)}*0.1,0)', calc: (r) => Math.round(r['申請数量'] * r['申請単価'] * 0.1), z: '#,##0', w: 12 },
        { h: '申請状況', w: 10 }, { h: '申請日', z: 'yyyy/mm/dd', w: 12 },
      ] },
      { name: '経費種別', rows: kindsRows, keepOrder: true, cols: [{ h: '種別コード' }, { h: '種別名', w: 18 }, { h: '月次上限額', z: '#,##0', w: 14 }] },
    ],
  });
}

// 03 品質チェック（2 シート。合否は「判定基準」シートの許容不良率を参照）
{
  const { rows, rejects } = mk({
    coverage: 0.9, rejects: 3, idH: '部品ID', nameH: '部品名',
    row: () => { const n = pick([50, 100, 200, 300, 500]); const bad = rnd() < 0.2 ? ri(Math.ceil(n * 0.02), Math.ceil(n * 0.05)) : ri(0, Math.max(1, Math.floor(n * 0.008))); return { 検査日: dateCell(2026, ri(9, 11), ri(1, 28)), 検査数: n, 不良数: bad, 検査員: pick(['佐藤', '鈴木', '高橋', '田中', '伊藤', '渡辺']) }; },
  });
  const TH = 0.01;
  files.push({
    file: '03_品質チェック.xlsx', keyHeaders: ['部品ID', '部品名'], note: '合否は 2 枚目「判定基準」の許容不良率(B2)を参照。2 シート',
    sheets: [
      { name: '検査結果', data: true, rows, rejects, cols: [
        ...IDN('部品ID', '部品名'), { h: '検査日', z: 'yyyy/mm/dd', w: 12 }, { h: '検査数', w: 10 }, { h: '不良数', w: 10 },
        { h: '不良率', f: '{不良数}/{検査数}', calc: (r) => r['不良数'] / r['検査数'], z: '0.0%', w: 10 },
        { h: '合否', f: 'IF({不良率}<=判定基準!$B$2,"合格","不合格")', calc: (r) => (r['不良数'] / r['検査数'] <= TH ? '合格' : '不合格'), w: 8 },
        { h: '検査員', w: 10 },
      ] },
      { name: '判定基準', rows: [{ 項目: '許容不良率', 値: TH }], keepOrder: true, cols: [{ h: '項目', w: 14 }, { h: '値', z: '0.0%' }] },
    ],
  });
}

// 04 発注数量（要望個数と実際の個数）
{
  const { rows, rejects } = mk({
    coverage: 0.95, rejects: 3, idH: '品番', nameH: '品名',
    row: () => { const want = ri(2, 60) * 10; const got = Math.max(0, Math.round(want * (0.7 + rnd() * 0.4))); return { 要望個数: want, 実際個数: got }; },
  });
  files.push({
    file: '04_発注数量.xlsx', keyHeaders: ['品番', '品名'], note: '見出しが「品番／品名」（表記ゆれ）。要望個数と実際の個数、差異・充足率の数式',
    sheets: [{ name: '発注数量', data: true, rows, rejects, cols: [
      ...IDN('品番', '品名'), { h: '要望個数', w: 10 }, { h: '実際個数', w: 10 },
      { h: '差異', f: '{実際個数}-{要望個数}', calc: (r) => r['実際個数'] - r['要望個数'], w: 8 },
      { h: '充足率', f: '{実際個数}/{要望個数}', calc: (r) => r['実際個数'] / r['要望個数'], z: '0%', w: 8 },
      { h: '充足判定', f: 'IF({差異}>=0,"充足","不足")', calc: (r) => (r['実際個数'] - r['要望個数'] >= 0 ? '充足' : '不足'), w: 10 },
    ] }],
  });
}

// 05 在庫状況（社内在庫・サプライヤー在庫）
{
  const { rows, rejects } = mk({
    coverage: 0.88, rejects: 4, idH: '部品ID', nameH: '部品名',
    row: () => ({ 社内在庫: ri(0, 600), サプライヤー現在在庫: ri(0, 1500), 安全在庫: ri(2, 20) * 10 }),
  });
  files.push({
    file: '05_在庫状況.xlsx', keyHeaders: ['部品ID', '部品名'], note: '社内在庫＋サプライヤー現在在庫の合計と補充判定が数式',
    sheets: [{ name: '在庫', data: true, rows, rejects, cols: [
      ...IDN('部品ID', '部品名'), { h: '社内在庫', w: 10 }, { h: 'サプライヤー現在在庫', w: 20 },
      { h: '合計在庫', f: '{社内在庫}+{サプライヤー現在在庫}', calc: (r) => r['社内在庫'] + r['サプライヤー現在在庫'], w: 10 },
      { h: '安全在庫', w: 10 },
      { h: '補充判定', f: 'IF({合計在庫}<{安全在庫},"要補充","")', calc: (r) => (r['社内在庫'] + r['サプライヤー現在在庫'] < r['安全在庫'] ? '要補充' : ''), w: 10 },
    ] }],
  });
}

// 06 取付位置・工数
{
  const POS = { 冷却: 'エンジンルーム前方', 電装: 'エンジンルーム側面', エンジン: 'エンジン本体', 燃料: '床下後部', 駆動: 'フロア下', 足回り: 'サスペンション周辺', ステアリング: 'ダッシュ下', ブレーキ: 'ホイール周辺', 空調: 'ダッシュ内', 内装: '室内', 安全: '室内（乗員側）', 外装: '車体外板' };
  const LINES = ['第1組立ライン', '第2組立ライン', '艤装ライン', 'シャシーライン', '最終ライン'];
  const { rows, rejects } = mk({
    coverage: 0.86, rejects: 3, idH: '部品番号', nameH: '部品名称',
    row: (p) => ({ 取付位置: POS[p.cat] ?? '車体', 組付ライン: pick(LINES), 工程番号: `P${ri(10, 99)}`, ボルト本数: ri(2, 14), 基本工数分: round(0.5 + rnd() * 2.5, 1) }),
  });
  files.push({
    file: '06_取付位置.xlsx', keyHeaders: ['部品番号', '部品名称'], note: '見出しが「部品番号／部品名称」。取付位置・ライン・工数（数式）',
    sheets: [{ name: '取付位置', data: true, rows, rejects, cols: [
      ...IDN('部品番号', '部品名称'), { h: '取付位置', w: 20 }, { h: '組付ライン', w: 16 }, { h: '工程番号', w: 10 }, { h: 'ボルト本数', w: 10 },
      { h: '基本工数分', z: '0.0', w: 10 },
      { h: '締付工数分', f: 'ROUND({ボルト本数}*0.4,1)', calc: (r) => round(r['ボルト本数'] * 0.4, 1), z: '0.0', w: 10 },
      { h: '総工数分', f: '{基本工数分}+{締付工数分}', calc: (r) => round(r['基本工数分'] + round(r['ボルト本数'] * 0.4, 1), 1), z: '0.0', w: 10 },
    ] }],
  });
}

// 07 マニュアル（PDF 格納 URL）
{
  const { rows, rejects } = mk({
    coverage: 0.9, rejects: 3, idH: '部品ID', nameH: '部品名',
    row: (p) => ({ マニュアルPDF格納URL: `https://manual.example.com/honda/service/${p.id.replace(/-/g, '_')}.pdf`, 版数: `Rev.${ri(1, 6)}`, 更新日: dateCell(2026, ri(1, 10), ri(1, 28)), ページ数: ri(4, 48) }),
  });
  files.push({
    file: '07_整備マニュアル.xlsx', keyHeaders: ['部品ID', '部品名'], note: 'マニュアル PDF の格納 URL（架空ドメイン）。HYPERLINK 数式つき',
    sheets: [{ name: 'マニュアル', data: true, rows, rejects, cols: [
      ...IDN('部品ID', '部品名'), { h: 'マニュアルPDF格納URL', w: 60 },
      { h: '開くリンク', f: 'HYPERLINK({マニュアルPDF格納URL},"PDFを開く")', calc: () => 'PDFを開く', w: 12 },
      { h: '版数', w: 8 }, { h: '更新日', z: 'yyyy/mm/dd', w: 12 }, { h: 'ページ数', w: 8 },
    ] }],
  });
}

// 08 サプライヤー（架空の社名）
{
  const SUP = [['東海精機工業', '愛知県刈谷市'], ['関東プレス製作所', '埼玉県狭山市'], ['西日本電装', '福岡県北九州市'], ['北陸ダイカスト', '富山県高岡市'], ['中部ラバー化工', '静岡県浜松市'], ['九州シートシステム', '熊本県菊陽町']];
  const { rows, rejects } = mk({
    coverage: 0.92, rejects: 3, idH: '部品ID', nameH: '部品名',
    row: (p) => { const s = pick(SUP); return { サプライヤー名: s[0], 所在地: s[1], 'リードタイム(日)': ri(3, 45), 発注ロット: pick([10, 20, 50, 100]), 発注単価: Math.round(p.price * (0.55 + rnd() * 0.15) / 10) * 10 }; },
  });
  files.push({
    file: '08_サプライヤー.xlsx', keyHeaders: ['部品ID', '部品名'], note: 'サプライヤー名は架空。ロット金額が数式',
    sheets: [{ name: 'サプライヤー', data: true, rows, rejects, cols: [
      ...IDN('部品ID', '部品名'), { h: 'サプライヤー名', w: 20 }, { h: '所在地', w: 16 }, { h: 'リードタイム(日)', w: 14 }, { h: '発注ロット', w: 10 },
      { h: '発注単価', z: '#,##0', w: 10 },
      { h: 'ロット金額', f: '{発注ロット}*{発注単価}', calc: (r) => r['発注ロット'] * r['発注単価'], z: '#,##0', w: 14 },
    ] }],
  });
}

// 09 原価（2 シート。為替前提シートを参照）
{
  const { rows, rejects } = mk({
    coverage: 0.84, rejects: 4, idH: '部品ID', nameH: '部品名',
    row: (p) => { const mat = Math.round(p.price * (0.3 + rnd() * 0.1)); const proc = Math.round(p.price * (0.12 + rnd() * 0.08)); const lg = Math.round(p.price * (0.02 + rnd() * 0.03)); return { 材料費: mat, 加工費: proc, 物流費: lg, 販売単価: p.price }; },
  });
  files.push({
    file: '09_原価.xlsx', keyHeaders: ['部品ID', '部品名'], note: '原価合計・粗利が数式。ドル建て参考額は「為替前提」シートの 1USD レートを参照（2 シート）',
    sheets: [
      { name: '原価明細', data: true, rows, rejects, cols: [
        ...IDN('部品ID', '部品名'), { h: '材料費', z: '#,##0', w: 10 }, { h: '加工費', z: '#,##0', w: 10 }, { h: '物流費', z: '#,##0', w: 10 },
        { h: '原価合計', f: 'SUM({材料費}:{物流費})', calc: (r) => r['材料費'] + r['加工費'] + r['物流費'], z: '#,##0', w: 12 },
        { h: '販売単価', z: '#,##0', w: 12 },
        { h: '粗利', f: '{販売単価}-{原価合計}', calc: (r) => r['販売単価'] - (r['材料費'] + r['加工費'] + r['物流費']), z: '#,##0', w: 12 },
        { h: '粗利率', f: '{粗利}/{販売単価}', calc: (r) => (r['販売単価'] - (r['材料費'] + r['加工費'] + r['物流費'])) / r['販売単価'], z: '0.0%', w: 10 },
        { h: '原価USD参考', f: 'ROUND({原価合計}/為替前提!$B$2,2)', calc: (r) => round((r['材料費'] + r['加工費'] + r['物流費']) / 150, 2), z: '#,##0.00', w: 14 },
      ] },
      { name: '為替前提', rows: [{ 項目: '1USD(円)', 値: 150 }], keepOrder: true, cols: [{ h: '項目', w: 14 }, { h: '値' }] },
    ],
  });
}

// 10 設計変更
{
  const CH = ['材質変更', '形状変更', '締結方法変更', 'コネクタ変更', '塗装仕様変更', '取付ブラケット追加'];
  const { rows, rejects } = mk({
    coverage: 0.5, rejects: 3, idH: '部品ID', nameH: '部品名',
    row: () => ({ 設計変更番号: `ECN-${ri(2026001, 2026999)}`, 変更区分: pick(CH), 変更日: dateCell(2026, ri(4, 11), ri(1, 28)), 影響台数: ri(2, 80) * 100, 切替予定日: dateCell(2027, ri(1, 6), ri(1, 28)) }),
  });
  files.push({
    file: '10_設計変更.xlsx', keyHeaders: ['部品ID', '部品名'], note: '一部の部品だけ（約半数）。切替までの日数は数式',
    sheets: [{ name: '設計変更', data: true, rows, rejects, cols: [
      ...IDN('部品ID', '部品名'), { h: '設計変更番号', w: 16 }, { h: '変更区分', w: 18 }, { h: '変更日', z: 'yyyy/mm/dd', w: 12 }, { h: '影響台数', z: '#,##0', w: 10 },
      { h: '切替予定日', z: 'yyyy/mm/dd', w: 12 },
      { h: '切替までの日数', f: '{切替予定日}-{変更日}', calc: (r) => r['切替予定日'] - r['変更日'], w: 14 },
    ] }],
  });
}

// 11 保管・物流
{
  const WH = ['刈谷第1倉庫', '狭山部品センター', '鈴鹿物流センター', '寄居保管庫', '熊本パーツDC'];
  const { rows, rejects } = mk({
    coverage: 0.9, rejects: 3, idH: '部品ID', nameH: '部品名',
    row: () => ({ 入数: pick([1, 2, 4, 5, 10, 20]), 箱サイズ: pick(['S(30x20x15)', 'M(45x30x25)', 'L(60x40x35)', 'LL(80x60x50)']), 重量kg: round(0.4 + rnd() * 24, 1), 保管倉庫: pick(WH), 棚番: `${pick('ABCDE')}-${ri(1, 20)}-${ri(1, 6)}`, 保管数量: ri(5, 800) }),
  });
  files.push({
    file: '11_保管物流.xlsx', keyHeaders: ['部品ID', '部品名'], note: '必要箱数（ROUNDUP）・総重量が数式',
    sheets: [{ name: '保管', data: true, rows, rejects, cols: [
      ...IDN('部品ID', '部品名'), { h: '入数', w: 8 }, { h: '箱サイズ', w: 16 }, { h: '重量kg', z: '0.0', w: 10 }, { h: '保管倉庫', w: 18 }, { h: '棚番', w: 10 }, { h: '保管数量', w: 10 },
      { h: '必要箱数', f: 'ROUNDUP({保管数量}/{入数},0)', calc: (r) => Math.ceil(r['保管数量'] / r['入数']), w: 10 },
      { h: '総重量kg', f: 'ROUND({重量kg}*{保管数量},1)', calc: (r) => round(r['重量kg'] * r['保管数量'], 1), z: '#,##0.0', w: 12 },
    ] }],
  });
}

// 12 保証・クレーム
{
  const { rows, rejects } = mk({
    coverage: 0.87, rejects: 3, idH: '部品ID', nameH: '部品名',
    row: () => { const ship = ri(8, 120) * 1000; return { '保証期間(月)': pick([12, 24, 36, 60]), 出荷台数: ship, 不具合件数: ri(0, Math.floor(ship / 2000)) }; },
  });
  files.push({
    file: '12_保証クレーム.xlsx', keyHeaders: ['部品ID', '部品名'], note: '1000台あたり不具合（IPTV 相当）が数式',
    sheets: [{ name: '保証', data: true, rows, rejects, cols: [
      ...IDN('部品ID', '部品名'), { h: '保証期間(月)', w: 12 }, { h: '出荷台数', z: '#,##0', w: 10 }, { h: '不具合件数', w: 10 },
      { h: '1000台あたり不具合', f: 'ROUND({不具合件数}/{出荷台数}*1000,2)', calc: (r) => round((r['不具合件数'] / r['出荷台数']) * 1000, 2), z: '0.00', w: 18 },
      { h: '要対策', f: 'IF({1000台あたり不具合}>=1,"要対策","")', calc: (r) => (round((r['不具合件数'] / r['出荷台数']) * 1000, 2) >= 1 ? '要対策' : ''), w: 10 },
    ] }],
  });
}

// 13 環境・規制
{
  const MAT = { 冷却: 'アルミ合金', 電装: '銅・樹脂', エンジン: '鋳鉄・アルミ', 燃料: '樹脂・鋼板', 駆動: '鋼・アルミ', 足回り: '鋼', ステアリング: '鋼・樹脂', ブレーキ: '鋳鉄・鋼', 空調: 'アルミ・樹脂', 内装: '樹脂・ウレタン', 安全: 'ナイロン・鋼', 外装: '鋼板・樹脂' };
  const { rows, rejects } = mk({
    coverage: 0.83, rejects: 3, idH: '部品ID', nameH: '部品名',
    row: (p) => ({ 主材質: MAT[p.cat] ?? '鋼', 質量kg: round(0.3 + rnd() * 18, 2), リサイクル率: round(0.5 + rnd() * 0.45, 2), 規制物質含有: pick(['なし', 'なし', 'なし', '届出済']) }),
  });
  files.push({
    file: '13_環境規制.xlsx', keyHeaders: ['部品ID', '部品名'], note: 'リサイクル質量が数式',
    sheets: [{ name: '環境規制', data: true, rows, rejects, cols: [
      ...IDN('部品ID', '部品名'), { h: '主材質', w: 16 }, { h: '質量kg', z: '0.00', w: 10 }, { h: 'リサイクル率', z: '0%', w: 12 },
      { h: 'リサイクル質量kg', f: 'ROUND({質量kg}*{リサイクル率},2)', calc: (r) => round(r['質量kg'] * r['リサイクル率'], 2), z: '0.00', w: 16 },
      { h: '規制物質含有', w: 12 },
    ] }],
  });
}

// 14 月次出荷（2 シート。2 枚目は 1 枚目を VLOOKUP で参照）
{
  const base = mk({
    coverage: 0.9, rejects: 3, idH: '部品ID', nameH: '部品名',
    row: (p) => ({ 出荷数: ri(20, 900), 単価: p.price }),
  });
  const qty10 = new Map(base.rows.map((r) => [r['部品ID'], r['出荷数']]));
  const nov = [...base.rows.slice().reverse().filter((_, i) => i % 7 !== 3).map((r) => ({ 部品ID: r['部品ID'], 部品名: r['部品名'], 出荷数: Math.max(0, Math.round(r['出荷数'] * (0.7 + rnd() * 0.6))), 単価: r['単価'] }))];
  files.push({
    file: '14_月次出荷.xlsx', keyHeaders: ['部品ID', '部品名'], note: '2 シート（10月出荷 / 11月出荷）。11月は 10月を VLOOKUP した前月差の数式',
    sheets: [
      { name: '10月出荷', data: true, rows: base.rows, rejects: base.rejects, cols: [
        ...IDN('部品ID', '部品名'), { h: '出荷数', z: '#,##0', w: 10 }, { h: '単価', z: '#,##0', w: 10 },
        { h: '出荷金額', f: '{出荷数}*{単価}', calc: (r) => r['出荷数'] * r['単価'], z: '#,##0', w: 14 },
      ] },
      { name: '11月出荷', rows: nov, cols: [
        ...IDN('部品ID', '部品名'), { h: '出荷数', z: '#,##0', w: 10 }, { h: '単価', z: '#,##0', w: 10 },
        { h: '出荷金額', f: '{出荷数}*{単価}', calc: (r) => r['出荷数'] * r['単価'], z: '#,##0', w: 14 },
        { h: '前月差', f: "{出荷数}-IFERROR(VLOOKUP({部品ID},'10月出荷'!$A:$C,3,FALSE),0)", calc: (r) => r['出荷数'] - (qty10.get(r['部品ID']) ?? 0), w: 10 },
      ] },
    ],
  });
}

// 15 試験評価
{
  const TESTS = [['耐久試験(h)', 500, 1200], ['気密試験(kPa)', 80, 140], ['引張強度(MPa)', 300, 640], ['耐振動(G)', 10, 22], ['耐熱(℃)', 120, 190]];
  const { rows, rejects } = mk({
    coverage: 0.7, rejects: 3, idH: '部品ID', nameH: '部品名',
    row: () => { const t = pick(TESTS); return { 試験項目: t[0], 規格値: t[1], 測定値: Math.round(t[1] * (0.9 + rnd() * 0.45)), 試験日: dateCell(2026, ri(7, 11), ri(1, 28)), 評価者: pick(['山田', '中村', '小林', '加藤']) }; },
  });
  files.push({
    file: '15_試験評価.xlsx', keyHeaders: ['部品ID', '部品名'], note: '規格値と測定値から OK/NG（数式）',
    sheets: [{ name: '試験評価', data: true, rows, rejects, cols: [
      ...IDN('部品ID', '部品名'), { h: '試験項目', w: 16 }, { h: '規格値', w: 10 }, { h: '測定値', w: 10 },
      { h: '余裕率', f: '({測定値}-{規格値})/{規格値}', calc: (r) => (r['測定値'] - r['規格値']) / r['規格値'], z: '0.0%', w: 10 },
      { h: '試験判定', f: 'IF({測定値}>={規格値},"OK","NG")', calc: (r) => (r['測定値'] >= r['規格値'] ? 'OK' : 'NG'), w: 10 },
      { h: '試験日', z: 'yyyy/mm/dd', w: 12 }, { h: '評価者', w: 10 },
    ] }],
  });
}

// 16 仕向地・生産計画
{
  const { rows, rejects } = mk({
    coverage: 0.93, rejects: 3, idH: '部品ID', nameH: '部品名',
    row: () => ({ 仕向地: pick(['国内', '北米', '欧州', 'アジア', '中国']), 月産計画: ri(2, 90) * 100, 生産開始: dateCell(ri(2020, 2025), ri(1, 12), 1), 生産終了予定: dateCell(ri(2028, 2032), ri(1, 12), 1) }),
  });
  files.push({
    file: '16_生産計画.xlsx', keyHeaders: ['部品ID', '部品名'], note: '年間計画・生産期間が数式',
    sheets: [{ name: '生産計画', data: true, rows, rejects, cols: [
      ...IDN('部品ID', '部品名'), { h: '仕向地', w: 10 }, { h: '月産計画', z: '#,##0', w: 10 },
      { h: '年間計画', f: '{月産計画}*12', calc: (r) => r['月産計画'] * 12, z: '#,##0', w: 12 },
      { h: '生産開始', z: 'yyyy/mm', w: 10 }, { h: '生産終了予定', z: 'yyyy/mm', w: 12 },
      { h: '生産期間(年)', f: 'ROUND(({生産終了予定}-{生産開始})/365.25,1)', calc: (r) => round((r['生産終了予定'] - r['生産開始']) / 365.25, 1), z: '0.0', w: 12 },
    ] }],
  });
}

// ---------- 「設定」「判定基準」「為替前提」: 見出し行の直下 B2 に値を置く（上の build は見出し=1行目・データ=2行目で作る） ----------
fs.mkdirSync(outDir, { recursive: true });
const truth = { generatedAt: '2026-11-17', master: PARTS, files: [] };
for (const f of files) truth.files.push(build(f));
fs.writeFileSync(path.join(outDir, 'truth.json'), JSON.stringify(truth, null, 1) + '\n', 'utf8');
console.log(`wrote ${files.length} files + truth.json to ${outDir}`);
for (const t of truth.files) console.log(`  ${t.file}: ${t.sheets.map((s) => `${s.name}(${s.rows.length}+${s.rejects.length})`).join(' / ')}`);
