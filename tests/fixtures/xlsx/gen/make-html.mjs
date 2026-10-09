// Generates html-sjis.xls (owner: reader R): an HTML table saved as "*.xls" in Shift_JIS, the way many business
// systems export "Excel" files — meta charset, <th> header, leading-zero codes, a colspan, <br> line breaks,
// entities, CP932-only characters (髙 ① ～ －), full-width spaces, an mso-number-format text style.
// Run: node tests/fixtures/xlsx/gen/make-html.mjs
import Encoding from 'encoding-japanese';
import { writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const dir = join(here, '..');

const html = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<style>.txt { mso-number-format:"\\@"; }</style>
</head>
<body>
<table border="1">
<tr><th>学生ID</th><th>氏名</th><th>大学</th><th>金額</th><th>メモ</th><th>日付</th></tr>
<tr><td class="txt">00123</td><td>髙橋　一郎</td><td>慶應義塾大学</td><td>1,234</td><td>①要確認</td><td>2026/10/20</td></tr>
<tr><td class="txt">00456</td><td>山田～太郎</td><td>早稲田大学</td><td>-5</td><td>A&amp;B &lt;注&gt;</td><td>2026/10/21</td></tr>
<tr><td>1001</td><td>佐藤 花子</td><td>東京大学</td><td>72.5</td><td>改行を<br>含む</td><td></td></tr>
<tr><td colspan="3">鈴木－次郎（合計行）</td><td>1302.5</td><td></td><td></td></tr>
</table>
</body>
</html>
`;
const sjis = Encoding.convert(Encoding.stringToCode(html), { to: 'SJIS', from: 'UNICODE' });
writeFileSync(join(dir, 'html-sjis.xls'), Buffer.from(sjis));
console.log('wrote html-sjis.xls');
