// Generates the hostile / messy input files used by the round-5 UI checks (tests/ui/hardening.spec.ts and the manual
// stress runs). Everything is fake data. Usage:
//   node tests/ui/hostile-fixtures.mjs <outDir> [--big] [--huge]   (--big: the 20,000 x 30 base + two others; --huge: the 60 MB files)
// or import { makeHostileFixtures } from './hostile-fixtures.mjs'.
import { mkdirSync, writeFileSync, existsSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { randomBytes } from 'node:crypto';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';

const SURNAMES = ['佐藤', '鈴木', '高橋', '田中', '伊藤', '渡辺', '山本', '中村', '小林', '加藤', '吉田', '山田', '佐々木', '山口', '松本', '井上', '木村', '林', '斎藤', '清水'];
const GIVEN = ['太郎', '花子', '一郎', '美咲', '健太', '結衣', '大輔', '陽菜', '翔太', '愛', '蓮', '彩', '拓海', '葵', '直樹', '優', '誠', '恵', '亮', '奈々'];

export function nameOf(i) {
  return `${SURNAMES[i % SURNAMES.length]} ${GIVEN[Math.floor(i / SURNAMES.length) % GIVEN.length]}${i >= 400 ? ' ' + Math.floor(i / 400) : ''}`;
}
export const idOf = (i) => `S${String(i + 1).padStart(6, '0')}`;

function book(sheets) {
  const wb = XLSX.utils.book_new();
  for (const [name, aoa, opts] of sheets) {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    if (opts?.merges) ws['!merges'] = opts.merges;
    XLSX.utils.book_append_sheet(wb, ws, name);
  }
  return wb;
}
const xlsxBytes = (wb, bookType = 'xlsx') => Buffer.from(XLSX.write(wb, { type: 'buffer', bookType }));

/** small, normal-ish file: ID / 氏名 / n extra columns */
function simple(n, cols, headerExtra = 'スコア', idHeader = '学生ID', nameHeader = '氏名', idFmt = idOf) {
  const head = [idHeader, nameHeader, ...Array.from({ length: cols }, (_, c) => `${headerExtra}${c + 1}`)];
  const rows = Array.from({ length: n }, (_, i) => [idFmt(i), nameOf(i), ...Array.from({ length: cols }, (_, c) => ((i * 7 + c * 13) % 100))]);
  return [head, ...rows];
}

/**
 * @param {string} dir output folder (created)
 * @param {{ big?: boolean, huge?: boolean }} [opts]
 * @returns {Record<string, string>} name -> absolute path
 */
export async function makeHostileFixtures(dir, opts = {}) {
  dir = resolve(dir);
  mkdirSync(dir, { recursive: true });
  const out = {};
  const put = (name, bytes) => {
    const p = join(dir, name);
    writeFileSync(p, bytes);
    out[name] = p;
    return p;
  };

  // --- plain good files
  put('a_名簿.xlsx', xlsxBytes(book([['名簿', simple(40, 3, '名簿項目')]])));
  put('b_成績.xlsx', xlsxBytes(book([['成績', simple(38, 2, '成績', '学生ID', '氏名')]])));
  put('c_アンケート.xlsx', xlsxBytes(book([['回答', simple(35, 4, 'Q', '学生ID', '氏名')]])));

  // --- broken inputs
  put('zero.xlsx', Buffer.alloc(0));
  put('notexcel.xlsx', Buffer.from('これはExcelではありません。ただのテキストです。\n'.repeat(20), 'utf8'));
  put('メモ.txt', Buffer.from('memo', 'utf8'));
  put('photo.png', Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0]));
  put('~$a_名簿.xlsx', Buffer.alloc(165, 0x41));
  const good = xlsxBytes(book([['名簿', simple(40, 3)]]));
  put('corrupt.xlsx', good.subarray(0, Math.floor(good.length / 2)));
  // password-protected workbooks are OLE2 containers with EncryptionInfo / EncryptedPackage streams
  const cfb = Buffer.alloc(4096, 0);
  Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).copy(cfb, 0);
  Buffer.from('EncryptionInfo', 'utf16le').copy(cfb, 1024);
  Buffer.from('EncryptedPackage', 'utf16le').copy(cfb, 1152);
  put('password.xlsx', cfb);

  // --- layouts
  // data sheet is not the first one (a cover sheet comes first)
  put(
    'sheet2_データが2枚目.xlsx',
    xlsxBytes(
      book([
        ['表紙', [['2027年卒 インターン応募者データ'], ['作成: 人事部'], [''], ['※ 次のシートにデータがあります']]],
        ['データ', simple(40, 3, '項目')],
      ]),
    ),
  );
  // header not in row 1: title + blank + note, header in row 4
  const titled = [['応募者一覧（2027年卒）'], [''], ['作成日: 2026/10/01 / 人事部'], ...simple(36, 3, '評価')];
  put('header4_見出しが4行目.xlsx', xlsxBytes(book([['一覧', titled]])));
  // two-row header (group row + sub row)
  const two = [
    ['', '', '面接', '', '', '適性', ''],
    ['学生ID', '氏名', '1次', '2次', '最終', '言語', '非言語'],
    ...Array.from({ length: 34 }, (_, i) => [idOf(i), nameOf(i), i % 5, (i * 3) % 5, (i * 7) % 5, (i * 11) % 100, (i * 13) % 100]),
  ];
  put('twohdr_2段見出し.xlsx', xlsxBytes(book([['評価', two, { merges: [{ s: { r: 0, c: 2 }, e: { r: 0, c: 4 } }, { s: { r: 0, c: 5 }, e: { r: 0, c: 6 } }] }]])));
  // English-only headers
  const en = [['ID', 'Name', 'Score', 'Grade'], ...Array.from({ length: 30 }, (_, i) => [idOf(i), nameOf(i), (i * 9) % 100, 'ABC'[i % 3]])];
  put('english_headers.xlsx', xlsxBytes(book([['Sheet1', en]])));
  // numbers only (no text at all)
  const nums = Array.from({ length: 50 }, (_, i) => [i + 1, (i * 37) % 1000, (i * 11) % 500, (i * 5) % 97]);
  put('numbers_only.xlsx', xlsxBytes(book([['Sheet1', nums]])));
  put('numbers_only2.xlsx', xlsxBytes(book([['Sheet1', nums.map((r) => [r[0], ...r.slice(1).map((x) => x + 1)])]])));
  // a sheet with a header only (no data rows) and an entirely empty book
  put('header_only.xlsx', xlsxBytes(book([['空', [['学生ID', '氏名', '点数']]]])));
  put('empty_sheet.xlsx', xlsxBytes(book([['空', [[]]]])));
  // very long file name
  const long = '非常に長いファイル名の例_'.repeat(9) + '終.xlsx';
  put(long, xlsxBytes(book([['名簿', simple(30, 2, '長')]])));
  // csv (UTF-8) with a BOM and a csv with only a header
  put('utf8.csv', Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from('学生ID,氏名,メモ\n' + Array.from({ length: 20 }, (_, i) => `${idOf(i)},${nameOf(i)},メモ${i}`).join('\n'), 'utf8')]));

  // very wide sheets (300 columns each)
  const wide = (prefix, n, rows) => [
    ['学生ID', '氏名', ...Array.from({ length: n }, (_, c) => `${prefix}${c + 1}`)],
    ...Array.from({ length: rows }, (_, i) => [idOf(i), nameOf(i), ...Array.from({ length: n }, (_, c) => (i * 3 + c) % 1000)]),
  ];
  put('wide1_300列.xlsx', xlsxBytes(book([['広い', wide('甲', 300, 60)]])));
  put('wide2_300列.xlsx', xlsxBytes(book([['広い', wide('乙', 300, 55)]])));
  // sheet names with an apostrophe / brackets-like characters and a formula column that refers to the sheet
  const fsheet = [['学生ID', '氏名', '国語', '数学', '合計'], ...Array.from({ length: 20 }, (_, i) => [idOf(i), nameOf(i), 50 + i, 40 + i, { t: 'n', f: `SUM(C${i + 2}:D${i + 2})`, v: 90 + 2 * i }])];
  const wbf = XLSX.utils.book_new();
  const wsf = XLSX.utils.aoa_to_sheet(fsheet.map((r) => r.map((c) => (c && typeof c === 'object' ? c.v : c))));
  for (let i = 0; i < 20; i++) wsf[`E${i + 2}`] = { t: 'n', f: `SUM(C${i + 2}:D${i + 2})`, v: 90 + 2 * i };
  XLSX.utils.book_append_sheet(wbf, wsf, "it's data（2024）");
  put('sheetname_アポストロフィ.xlsx', xlsxBytes(wbf));

  if (opts.big) {
    // 20,000 rows x 30 columns base (+ two smaller files that share the IDs / names)
    const cols = 28;
    put('big_base_20000x30.xlsx', xlsxBytes(book([['大', simple(20000, cols, '項目')]])));
    put('big_other1.xlsx', xlsxBytes(book([['評価', simple(19000, 3, '評価', '学生ID', '氏名')]])));
    const o2 = [['氏名', '所属', '備考'], ...Array.from({ length: 15000 }, (_, i) => [nameOf(i), `部署${i % 40}`, `備考${i}`])];
    put('big_other2.xlsx', xlsxBytes(book([['所属', o2]])));
  }
  if (opts.huge) {
    // 60 MB package: a normal small workbook plus an incompressible 58 MB image part
    const z = await JSZip.loadAsync(xlsxBytes(book([['名簿', simple(40, 3)]])));
    z.file('xl/media/image1.png', randomBytes(58 * 1024 * 1024), { compression: 'STORE' });
    const big = await z.generateAsync({ type: 'nodebuffer', compression: 'STORE' });
    put('big60_画像入り.xlsx', big);
    // 60 MB csv
    const lines = ['学生ID,氏名,メモ,値1,値2,値3,値4,値5'];
    for (let i = 0; i < 700000; i++) lines.push(`${idOf(i)},${nameOf(i)},メモ${i},${i % 97},${i % 89},${i % 83},${i % 79},${i % 73}`);
    put('big60.csv', Buffer.from(lines.join('\n'), 'utf8'));
  }
  return out;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
  const dir = process.argv[2];
  if (!dir) {
    console.error('usage: node tests/ui/hostile-fixtures.mjs <outDir> [--big] [--huge]');
    process.exit(2);
  }
  const res = await makeHostileFixtures(dir, { big: process.argv.includes('--big'), huge: process.argv.includes('--huge') });
  for (const [k, v] of Object.entries(res)) console.log(existsSync(v) ? 'ok' : 'MISSING', k);
}
