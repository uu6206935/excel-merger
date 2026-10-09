/**
 * Generated workbooks of the consistency oracle (task V1), built in Node (no Excel, no Python):
 *  - the "rich base": a student list whose columns carry formulas of every reference kind (relative / absolute / mixed / whole-column /
 *    cross-sheet / names / structured references of an Excel Table / lookups by column number), a total row below the data, a report
 *    sheet that reads the list, CF and DV rules with formulas;
 *  - the "row source": a second file with the same people in another order (reversed / shuffled / extra / missing / duplicated) and
 *    formulas of the kinds that depend on rows (same row, whole-table ranges, lookups by ID, running totals, key-column reads, names,
 *    structured references).
 * Formula cells carry NO stored result (like openpyxl / XlsxWriter output); the cached twins are authored by Excel (excel-batch.ps1 resave).
 */
import { buildXlsx, STYLES } from '../../golden/build';
import { colLetter } from './lib';
import { sheetData, type G } from '../../xlsx/formulaimport-helpers';
import { rng } from '../../doc/fuzzcore';

export const PEOPLE: Array<[string, string, string, number]> = [
  ['S001', '鈴木美咲', '京都大学', 2], ['S002', '高橋大輔', '慶應義塾大学', 3], ['S003', '田中愛', '早稲田大学', 4], ['S004', '伊藤一郎', '大阪大学', 1],
  ['S005', '渡辺結衣', '東京大学', 2], ['S006', '山本翔太', '京都大学', 3], ['S007', '中村花子', '慶應義塾大学', 4], ['S008', '小林健太', '早稲田大学', 1],
  ['S009', '加藤陽菜', '大阪大学', 2], ['S010', '佐藤花子', '東京大学', 3], ['S011', '鈴木健太', '京都大学', 4], ['S012', '高橋陽菜', '慶應義塾大学', 1],
  ['S013', '田中太郎', '早稲田大学', 2], ['S014', '伊藤美咲', '大阪大学', 3], ['S015', '渡辺大輔', '東京大学', 4], ['S016', '山本愛', '京都大学', 1],
  ['S017', '中村一郎', '慶應義塾大学', 2], ['S018', '小林結衣', '早稲田大学', 3], ['S019', '加藤翔太', '大阪大学', 4], ['S020', '佐藤一郎', '東京大学', 1],
  ['S021', '鈴木結衣', '京都大学', 2], ['S022', '高橋翔太', '慶應義塾大学', 3], ['S023', '田中花子', '早稲田大学', 4], ['S024', '伊藤健太', '大阪大学', 1],
  ['S025', '渡辺陽菜', '東京大学', 2], ['S026', '山本太郎', '京都大学', 3], ['S027', '中村美咲', '慶應義塾大学', 4], ['S028', '小林大輔', '早稲田大学', 1],
  ['S029', '加藤愛', '大阪大学', 2], ['S030', '佐藤美咲', '東京大学', 3],
];
const EXTRA_PEOPLE: Array<[string, string, string, number]> = [
  ['S901', '新井一', '京都大学', 2], ['S902', '森田二', '東京大学', 3], ['S903', '石井三', '大阪大学', 1], ['S904', '岡田四', '早稲田大学', 4],
];

const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** STYLES with one differential format for conditional formatting (dxfId 0) */
const STYLES_DXF = STYLES.replace('<dxfs count="0"/>', '<dxfs count="1"><dxf><font><b/><color rgb="FFC00000"/></font></dxf></dxfs>');

/** 'quoted': sheet names that need quoting in a formula (a space, an apostrophe, a hyphen) */
export type BaseVariant = 'plain' | 'table' | 'title' | 'quoted';

/** a sheet name as a formula writes it */
export const quoteSheet = (n: string): string => (/^[A-Za-z_\u0080-￿][A-Za-z0-9_.\u0080-￿]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);

export interface RichBase {
  bytes: Uint8Array;
  /** sheet of the data */
  sheet: string;
  headerRow: number;
  firstDataRow: number;
  lastDataRow: number;
}

const HEAD = ['学生ID', '氏名', '大学', '学年', '学年2倍', '学年累計', '大学人数', '全体比', '列全体', '地域', '基準以上', '連結', '検索', '3倍'];

/** columns 1..14 of the rich base; the last (3倍) reads the table by structured reference in the table variant */
function richFormulas(r: number, f: number, l: number, t: number, table: boolean, M: string): G[] {
  return [
    { f: `D${r}*2` }, // E  relative
    { f: `SUM($D$${f}:D${r})` }, // F  mixed (running)
    { f: `COUNTIF($C$${f}:$C$${l},C${r})` }, // G  absolute range
    { f: `D${r}/$D$${t}` }, // H  absolute single cell (total row)
    { f: `COUNTIF(C:C,C${r})` }, // I  whole column
    { f: `INDEX(${M}!$B$2:$B$9,MATCH(C${r},${M}!$A$2:$A$9,0))` }, // J  other sheet
    { f: `D${r}>=基準学年` }, // K  defined name
    { f: `A${r}&"-"&B${r}` }, // L  text
    { f: `VLOOKUP(A${r},$A$${f}:$D$${l},4,FALSE)` }, // M  lookup by column number (the documented exception)
    table ? { f: `tblS[[#This Row],[学年]]*3` } : { f: `D${r}*3` }, // N  structured reference / relative
  ];
}

/** The rich base workbook (30 students S001..S030, see HEAD for the columns). */
export async function buildRichBase(variant: BaseVariant): Promise<RichBase> {
  const title = variant === 'title';
  const h = title ? 3 : 1;
  const f = h + 1;
  const l = f + PEOPLE.length - 1;
  const t = l + 2;
  const table = variant === 'table';
  const NM = variant === 'quoted' ? { main: "Student's List", rep: 'Report 1', mst: 'Master-1' } : { main: '名簿', rep: '集計', mst: 'マスタ' };
  const MAIN = quoteSheet(NM.main);
  const M = quoteSheet(NM.mst);
  const rows: G[][] = [];
  if (title) {
    rows.push(['学生名簿 2027', null, null, null, null, null, null, null]);
    rows.push([]);
  }
  rows.push(HEAD.map((x) => x));
  PEOPLE.forEach(([id, name, uni, year], i) => rows.push([id, name, uni, year, ...richFormulas(f + i, f, l, t, table, M)]));
  rows.push([]);
  const total: G[] = ['合計', null, null, { f: `SUM(D${f}:D${l})` }, { f: `SUM(E${f}:E${l})` }, null, null, { f: `SUM(H${f}:H${l})` }];
  rows.push(total);
  // the title cell carries the style (fill + borders) of an openpyxl-style merged band: s=1
  let main = sheetData(rows);
  if (title) main = main.replace('<c r="A1" t="inlineStr">', '<c r="A1" s="1" t="inlineStr">');
  const merges = title ? `<mergeCells count="1"><mergeCell ref="A1:H1"/></mergeCells>` : '';
  const cf =
    `<conditionalFormatting sqref="A${f}:D${l}"><cfRule type="expression" dxfId="0" priority="1"><formula>$D${f}&gt;=基準学年</formula></cfRule></conditionalFormatting>` +
    `<conditionalFormatting sqref="E${f}:E${l}"><cfRule type="expression" dxfId="0" priority="2"><formula>$E${f}&gt;AVERAGE($E$${f}:$E$${l})</formula></cfRule></conditionalFormatting>`;
  const dv =
    `<dataValidations count="2"><dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="C${f}:C${l}"><formula1>${M}!$A$2:$A$9</formula1></dataValidation>` +
    `<dataValidation type="whole" operator="between" allowBlank="1" sqref="D${f}:D${l}"><formula1>1</formula1><formula2>$D$${t}</formula2></dataValidation></dataValidations>`;
  const tableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="tblS" displayName="tblS" ref="A${h}:N${l}"><autoFilter ref="A${h}:N${l}"/><tableColumns count="${HEAD.length}">${HEAD.map((c, k) => `<tableColumn id="${k + 1}" name="${esc(c)}"/>`).join('')}</tableColumns><tableStyleInfo name="TableStyleLight1" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`;
  const mainSheet = {
    name: NM.main,
    data: main,
    tail: `${merges}${cf}${dv}${table ? `<tableParts count="1"><tablePart r:id="rId1"/></tableParts>` : ''}`,
    ...(table ? { rels: [{ type: 'table', id: 'rId1', path: 'xl/tables/table1.xml', contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml', xml: tableXml }] } : {}),
  };
  const report: G[][] = [
    [{ f: `SUM(${MAIN}!E${f}:E${l})` }, '列範囲の合計'],
    [{ f: `AVERAGE(${MAIN}!$D:$D)` }, '学年の平均(列全体)'],
    [{ f: `COUNTIF(${MAIN}!C:C,"京都大学")` }, '京都大学の人数'],
    [{ f: `SUMPRODUCT((${MAIN}!D${f}:D${l}>2)*${MAIN}!E${f}:E${l})` }, 'SUMPRODUCT'],
    [{ f: `INDEX(${MAIN}!B${f}:B${l},MATCH(MAX(${MAIN}!E${f}:E${l}),${MAIN}!E${f}:E${l},0))` }, 'INDEX/MATCH'],
    [{ f: 'SUM(学年列)' }, '名前'],
    [{ f: `${MAIN}!B${f}&"/"&${MAIN}!A${f + 1}` }, '単一セル'],
    [{ f: `SUM(${MAIN}!D${f}:F${l})` }, '複数列の範囲'],
    [{ f: `VLOOKUP("S003",${MAIN}!A${f}:D${l},4,FALSE)` }, 'VLOOKUP(列番号)'],
    [{ f: `COUNTA(${MAIN}!A:A)` }, 'COUNTA(列全体)'],
    [{ f: `SUMIF(${MAIN}!C${f}:C${l},"早稲田大学",${MAIN}!D${f}:D${l})` }, 'SUMIF'],
    [{ f: `${MAIN}!$D$${t}` }, '合計行'],
    [{ f: `SUM(${MAIN}!G:G)` }, '列全体の合計(G)'],
  ];
  if (table) {
    report.push([{ f: 'SUM(tblS[学年])' }, '構造化参照'], [{ f: 'COUNTA(tblS[氏名])' }, '構造化参照(氏名)'], [{ f: 'SUM(tblS[[学年]:[学年2倍]])' }, '構造化参照(複数列)']);
  }
  const master: G[][] = [
    ['大学', '地域', null, '基準', '基準学年'],
    ['京都大学', '関西', null, null, null],
    ['慶應義塾大学', '関東', null, null, null],
    ['早稲田大学', '関東', null, null, null],
    ['大阪大学', '関西', null, null, null],
    ['東京大学', '関東', null, null, null],
    ['名古屋大学', '中部', null, null, null],
    ['北海道大学', '北海道', null, null, null],
    ['九州大学', '九州', null, null, null],
  ];
  master[0][4] = '基準学年';
  master[1][3] = '学年';
  master[1][4] = 2;
  // the name 基準学年 points at マスタ!$E$2
  const names = `<definedName name="基準学年">${esc(M)}!$E$2</definedName><definedName name="学年列">${esc(MAIN)}!$D$${f}:$D$${l}</definedName>`;
  const bytes = await buildXlsx(
    [mainSheet, { name: NM.rep, data: sheetData(report) }, { name: NM.mst, data: sheetData(master) }],
    { names, styles: STYLES_DXF },
  );
  return { bytes, sheet: NM.main, headerRow: h, firstDataRow: f, lastDataRow: l };
}

// ---------------------------------------------------------------------------------------------------------------------
// the row source
// ---------------------------------------------------------------------------------------------------------------------

export type RowOrder = 'same' | 'rev' | 'shuf' | 'extra' | 'missing' | 'dup' | 'mix';
export const ROW_ORDERS: RowOrder[] = ['same', 'rev', 'shuf', 'extra', 'missing', 'dup', 'mix'];

export interface RowSource {
  bytes: Uint8Array;
  sheet: string;
  /** person key of every data row (1-based sheet row -> id) */
  persons: string[];
  firstDataRow: number;
  lastDataRow: number;
  /** formula column headers in order, so a test can name them */
  headers: string[];
}

const SRC_HEAD = ['学生ID', '氏名', '点1', '点2', '点3', '合計', '平均', '判定', '順位', '偏差', '累計', '様', '検索', '割合', '最高との差', '人数'];

/**
 * The second file: people in `order` and formulas that read the same row, the whole table, other rows (running total), the key column, a defined name.
 * `table`: the list is an Excel Table with structured references.
 */
export async function buildRowSource(order: RowOrder, seed: number, opts: { table?: boolean } = {}): Promise<RowSource> {
  const r = rng(seed * 7919 + 13);
  const list = PEOPLE.map((p, i) => ({ p, i }));
  const shuffle = <T,>(a: T[]): T[] => {
    const x = a.slice();
    for (let i = x.length - 1; i > 0; i--) {
      const j = Math.floor(r() * (i + 1));
      [x[i], x[j]] = [x[j], x[i]];
    }
    return x;
  };
  type Row = { p: [string, string, string, number]; i: number; dup?: boolean };
  const extras: Row[] = EXTRA_PEOPLE.map((p, k) => ({ p, i: 100 + k }));
  /** a second row of the same person AFTER its first row (policy 'first' keeps the first one) */
  const withDuplicates = (rs: Row[], picks: number[]): Row[] => {
    const out = rs.slice();
    for (const k of picks) {
      const orig = rs[k % rs.length];
      const pos = out.indexOf(orig);
      const at = pos + 1 + Math.floor(r() * (out.length - pos));
      out.splice(at, 0, { ...orig, dup: true });
    }
    return out;
  };
  let rows: Row[] = list;
  if (order === 'rev') rows = list.slice().reverse();
  else if (order === 'shuf') rows = shuffle(list);
  else if (order === 'extra') rows = shuffle([...list, ...extras]);
  else if (order === 'missing') rows = shuffle(list).filter((_x, k) => k % 6 !== 2);
  else if (order === 'dup') rows = withDuplicates(shuffle(list), [3, 9]);
  else if (order === 'mix') rows = withDuplicates(shuffle([...list, ...extras]).filter((_x, k) => k % 7 !== 3), [2, 8]);
  const n = rows.length;
  const f = 2;
  const l = f + n - 1;
  const tbl = !!opts.table;
  const grid: G[][] = [SRC_HEAD.map((x) => x)];
  const persons: string[] = [''];
  rows.forEach((row, k) => {
    const rr = f + k;
    // the scores come from the person index (a different value for the duplicate copy of a person)
    const idx = row.i % 100;
    const dupShift = row.dup ? 3 : 0;
    const s1 = 50 + ((idx * 7 + dupShift) % 41);
    const s2 = 45 + ((idx * 11) % 37);
    const s3 = 30 + idx * 2 + (row.dup ? 1 : 0) + (row.i >= 100 ? 40 : 0);
    const fm: G[] = tbl
      ? [
          { f: `tblR[[#This Row],[点1]]+tblR[[#This Row],[点2]]+tblR[[#This Row],[点3]]` },
          { f: `AVERAGE(tblR[[#This Row],[点1]:[点3]])` },
          { f: `IF(tblR[[#This Row],[平均]]>=70,"優","可")` },
          { f: `RANK(tblR[[#This Row],[合計]],tblR[合計])` },
          { f: `tblR[[#This Row],[合計]]-AVERAGE(tblR[合計])` },
          { f: `SUM($F$${f}:F${rr})` },
          { f: `tblR[[#This Row],[学生ID]]&"様"` },
          { f: `INDEX(tblR[合計],MATCH(tblR[[#This Row],[学生ID]],tblR[学生ID],0))` },
          { f: `tblR[[#This Row],[合計]]/SUM(tblR[合計])` },
          { f: `MAX(tblR[合計])-tblR[[#This Row],[合計]]` },
          { f: `COUNTA(tblR[学生ID])` },
        ]
      : [
          { f: `C${rr}+D${rr}+E${rr}` },
          { f: `AVERAGE(C${rr}:E${rr})` },
          { f: `IF(G${rr}>=70,"優","可")` },
          { f: `RANK(F${rr},$F$${f}:$F$${l})` },
          { f: `F${rr}-AVERAGE($F$${f}:$F$${l})` },
          { f: `SUM($F$${f}:F${rr})` },
          { f: `A${rr}&"様"` },
          { f: `INDEX($F$${f}:$F$${l},MATCH(A${rr},$A$${f}:$A$${l},0))` },
          { f: `F${rr}/SUM(合計列)` },
          { f: `MAX($F$${f}:$F$${l})-F${rr}` },
          { f: `COUNTA($A$${f}:$A$${l})` },
        ];
    grid.push([row.p[0], row.p[1], s1, s2, s3, ...fm]);
    persons.push(row.p[0]);
  });
  const names = tbl ? '' : `<definedName name="合計列">評価!$F$${f}:$F$${l}</definedName>`;
  const tableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="tblR" displayName="tblR" ref="A1:${colLetter(SRC_HEAD.length)}${l}"><autoFilter ref="A1:${colLetter(SRC_HEAD.length)}${l}"/><tableColumns count="${SRC_HEAD.length}">${SRC_HEAD.map((c, k) => `<tableColumn id="${k + 1}" name="${esc(c)}"/>`).join('')}</tableColumns><tableStyleInfo name="TableStyleLight1" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`;
  const sheet = {
    name: '評価',
    data: sheetData(grid),
    ...(tbl
      ? { tail: `<tableParts count="1"><tablePart r:id="rId1"/></tableParts>`, rels: [{ type: 'table', id: 'rId1', path: 'xl/tables/table1.xml', contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml', xml: tableXml }] }
      : {}),
  };
  const bytes = await buildXlsx([sheet], { names, styles: STYLES_DXF });
  return { bytes, sheet: '評価', persons, firstDataRow: f, lastDataRow: l, headers: SRC_HEAD };
}
