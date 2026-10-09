// Generates the CSV / TSV reader fixtures (owner: reader R):
//   sjis.csv      Shift_JIS (CP932) with leading-zero codes, quoted commas / line breaks, thousands separators,
//                 CP932-only characters (髙 ① ～ －), full-width spaces
//   utf8bom.csv   the same content in UTF-8 with BOM
//   plain.tsv     tab separated, UTF-8 without BOM
// Run: node tests/fixtures/xlsx/gen/make-csv.mjs
import Encoding from 'encoding-japanese';
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const dir = join(here, '..');

export const CSV_ROWS = [
  ['学生ID', '氏名', '大学', '金額', 'メモ', '点数'],
  ['00123', '髙橋　一郎', '慶應義塾大学', '1,234', '①要確認', '85'],
  ['00456', '山田～太郎', '早稲田大学', '-5', '"引用"を含む', '72.5'],
  ['1001', '佐藤 花子', '東京大学', '1234567890123456', '改行を\n含む', '0'],
  ['', '鈴木－次郎', '', '3.14', 'a,b', '2026/10/20'],
];

function csvLine(cells, sep) {
  return cells.map((c) => (/[",\n\t]/.test(c) ? '"' + c.replace(/"/g, '""') + '"' : c)).join(sep);
}

const csvText = CSV_ROWS.map((r) => csvLine(r, ',')).join('\r\n') + '\r\n';
const sjis = Encoding.convert(Encoding.stringToCode(csvText), { to: 'SJIS', from: 'UNICODE' });
writeFileSync(join(dir, 'sjis.csv'), Buffer.from(sjis));
writeFileSync(join(dir, 'utf8bom.csv'), Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(csvText, 'utf8')]));
writeFileSync(join(dir, 'plain.tsv'), Buffer.from(CSV_ROWS.map((r) => csvLine(r, '\t')).join('\n') + '\n', 'utf8'));
console.log('wrote sjis.csv utf8bom.csv plain.tsv');
