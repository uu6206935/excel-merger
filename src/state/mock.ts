/**
 * ?mock=1 fake services (DEV only: main.tsx imports this module dynamically behind import.meta.env.DEV, so it
 * is not part of the production bundle). Realistic fake SourceBooks / KeyAnalysis / JoinResult / EditorSnapshot
 * so that every screen of the UI can be exercised before the real modules land.
 *
 * URL variants: ?mock=1 (server + Excel), ?mock=noexcel (server without Excel), ?mock=standalone (no server),
 * ?mock=excelfail (Excel engine fails -> JS fallback), ?mock=excelwarn (Excel engine returns runner warnings).
 * Add &demo=files or &demo=merged to preload the demo files; &titlemerge=1 merges the title row A1:K1 of the
 * student list (moves across it are refused with the 'unmergeTitles' fix, like the real doc module).
 *
 * Doc ops use the real formula engine (excelRef.adjustFormula / planMove) so formulas follow moves / deletes.
 */
import type { ServerStatus } from '../api';
import type {
  CandidateDTO,
  ColumnDTO,
  DeleteImpact,
  ImportColumnSpec,
  JoinPlan,
  JoinResult,
  KeyAnalysis,
  KeyChip,
  KeySource,
  NormalizeOptions,
  OpResult,
  PairAnalysis,
  PairJoinReport,
  PremergeInput,
  PremergeLayout,
  PremergeResult,
  ProgressFn,
  RowRef,
  RunnerJob,
  SheetDTO,
} from '../core/contracts';
import type {
  Cell,
  CellWrite,
  CheckItem,
  DocColumn,
  EditorSnapshot,
  ExternalFormulaCell,
  LoggedOp,
  MergedDoc,
  MergedRowPlan,
  RangeBox,
  RowMatch,
  SheetGrid,
  SourceBook,
  SourceFormat,
  StructOp,
  TableRegion,
  XStyle,
} from '../core/types';
import { adjustFormula, movePerm, planMove } from '../core/formula/excelRef';
import { displayValue } from '../core/display/numfmt';
import { planSheetNames } from '../core/xlsx/sheetnames';
import { DEFAULT_OPTIONS, colLetter, extOf, stripExt } from './helpers';
import type { Services } from './services';

// =============================================================================================
// small utils

/** simulated latency; ?delay=0 (tests) makes the fakes instant */
let delayScale = 1;
const sleep = (ms: number) => new Promise<void>((r) => setTimeout(r, Math.round(ms * delayScale)));

function rng(seed: number) {
  let s = seed >>> 0;
  return () => {
    s = (s * 1664525 + 1013904223) >>> 0;
    return s / 4294967296;
  };
}

function abortError(): Error {
  const e = new Error('中止しました');
  e.name = 'AbortError';
  return e;
}

// =============================================================================================
// books

const BORDER = { style: 'thin' as const, color: '#BFBFBF' };
const ALL_BORDERS = { top: BORDER, right: BORDER, bottom: BORDER, left: BORDER };

/** style ids used by the generated books */
const ST = { def: 0, head: 1, text: 2, num: 3, date: 4, title: 5, id: 6, score: 7, formula: 8, note: 9 } as const;

function bookStyles(accent: string): XStyle[] {
  return [
    { font: '游ゴシック', size: 11 },
    { font: '游ゴシック', size: 11, bold: true, fill: accent, hAlign: 'center', vAlign: 'middle', border: ALL_BORDERS },
    { font: '游ゴシック', size: 11, border: ALL_BORDERS },
    { font: '游ゴシック', size: 11, border: ALL_BORDERS, numFmt: '#,##0' },
    { font: '游ゴシック', size: 11, border: ALL_BORDERS, numFmt: 'yyyy/m/d', hAlign: 'center' },
    { font: '游ゴシック', size: 14, bold: true },
    { font: '游ゴシック', size: 11, border: ALL_BORDERS, numFmt: '@' },
    { font: '游ゴシック', size: 11, border: ALL_BORDERS, numFmt: '0', hAlign: 'right' },
    { font: '游ゴシック', size: 11, bold: true, border: ALL_BORDERS, fill: '#FFF2CC', numFmt: '0' },
    { font: '游ゴシック', size: 10, color: '#595959', border: ALL_BORDERS, wrap: true },
  ];
}

type CellInit = string | number | boolean | null | undefined | { v?: string | number | boolean | null; f?: string; s?: number; t?: Cell['t'] };

function mkCell(x: CellInit, s: number): Cell | undefined {
  if (x === undefined) return s ? { t: 'z', v: null, s } : undefined;
  if (x === null) return { t: 'z', v: null, s };
  if (typeof x === 'string') return { t: 's', v: x, s };
  if (typeof x === 'number') return { t: s === ST.date ? 'd' : 'n', v: x, s };
  if (typeof x === 'boolean') return { t: 'b', v: x, s };
  const v = x.v ?? null;
  const t: Cell['t'] = x.t ?? (v === null ? 'z' : typeof v === 'number' ? 'n' : typeof v === 'boolean' ? 'b' : 's');
  const c: Cell = { t, v, s: x.s ?? s };
  if (x.f) c.f = x.f;
  return c;
}

interface SheetSpec {
  name: string;
  rows: CellInit[][];
  /** style per column for data rows */
  colStyles?: number[];
  /** style per row (overrides colStyles), e.g. header */
  rowStyles?: Record<number, number>;
  widths?: number[];
  freeze?: { cols: number; rows: number };
  region?: TableRegion;
  merges?: RangeBox[];
}

function mkSheet(spec: SheetSpec): SheetGrid {
  const rows: SheetGrid['rows'] = [];
  let maxCol = 0;
  spec.rows.forEach((r, ri) => {
    const out: Array<Cell | undefined> = [];
    r.forEach((x, ci) => {
      const s = spec.rowStyles?.[ri + 1] ?? spec.colStyles?.[ci] ?? 0;
      out[ci] = mkCell(x, s);
    });
    rows[ri] = out;
    maxCol = Math.max(maxCol, r.length);
  });
  return {
    name: spec.name,
    state: 'visible',
    maxRow: spec.rows.length,
    maxCol,
    rows,
    colProps: (spec.widths ?? []).map((w) => ({ width: w })),
    rowProps: [],
    features: { merges: spec.merges ?? [], rules: [], tables: [], arrays: [], hyperlinks: [], freeze: spec.freeze },
    defaultColWidth: undefined,
    defaultRowHeight: 18.75,
  };
}

interface BookMeta {
  dataSheet: string;
  regions: Record<string, TableRegion>;
}

const bookMeta = new WeakMap<SourceBook, BookMeta>();
const mergedStore = new WeakMap<Uint8Array, SourceBook>();

const FAMILY = ['青木', '石川', '上田', '遠藤', '大野', '加藤', '菊地', '工藤', '小林', '斎藤', '佐々木', '清水', '鈴木', '高橋', '田中', '中村', '西田', '野口', '橋本', '林', '藤田', '前田', '松本', '三浦', '村上', '森', '山口', '山田', '吉田', '渡辺'];
const FAMILY_KANA = ['アオキ', 'イシカワ', 'ウエダ', 'エンドウ', 'オオノ', 'カトウ', 'キクチ', 'クドウ', 'コバヤシ', 'サイトウ', 'ササキ', 'シミズ', 'スズキ', 'タカハシ', 'タナカ', 'ナカムラ', 'ニシダ', 'ノグチ', 'ハシモト', 'ハヤシ', 'フジタ', 'マエダ', 'マツモト', 'ミウラ', 'ムラカミ', 'モリ', 'ヤマグチ', 'ヤマダ', 'ヨシダ', 'ワタナベ'];
const GIVEN = ['翔太', '美咲', '大輝', '結衣', '拓海', '陽菜', '健太', 'さくら', '悠斗', '彩花', '蓮', '七海', '颯太', '葵', '湊', '凛', '大和', '杏奈', '樹', '結菜', '陸', '心春', '海斗', '芽衣', '蒼', '莉子', '陽向', '花音', '律', '真央'];
const GIVEN_KANA = ['ショウタ', 'ミサキ', 'ダイキ', 'ユイ', 'タクミ', 'ヒナ', 'ケンタ', 'サクラ', 'ユウト', 'アヤカ', 'レン', 'ナナミ', 'ソウタ', 'アオイ', 'ミナト', 'リン', 'ヤマト', 'アンナ', 'イツキ', 'ユイナ', 'リク', 'コハル', 'カイト', 'メイ', 'アオ', 'リコ', 'ヒナタ', 'カノン', 'リツ', 'マオ'];
const UNIV = ['早稲田大学', '慶應義塾大学', '東京大学', '京都大学', '大阪大学', '一橋大学', '東北大学', '名古屋大学', '九州大学', '北海道大学'];
const UNIV_ABBR: Record<string, string> = { 早稲田大学: '早大', 慶應義塾大学: '慶大', 東京大学: '東大', 京都大学: '京大', 大阪大学: '阪大' };
const FACULTY = ['経済学部', '法学部', '商学部', '文学部', '理工学部', '工学部', '情報学部', '経営学部'];
const WORKPLACE = ['営業企画部', '人事部', '経理部', 'システム開発部', 'マーケティング部', '総務部'];

/** number of generated people (?rows=10000 for performance checks) and extra numeric base columns (?cols=40) */
let nStudents = 30;
let extraCols = 0;
/** &titlemerge=1: the title row of the student list is merged (A1:K1) */
let titleMerge = false;

interface Student {
  id: string;
  name: string;
  kana: string;
  univ: string;
  faculty: string;
  grade: number;
  mail: string;
  written: number;
  interview: number;
}

function students(): Student[] {
  const r = rng(20260930);
  const out: Student[] = [];
  for (let i = 0; i < nStudents; i++) {
    const f = i % 30;
    const g = i < 30 ? i : (i * 7 + Math.floor(i / 30)) % 30;
    const dup = i >= 900 ? String(Math.floor(i / 900)) : '';
    out.push({
      id: String(10001 + i),
      name: `${FAMILY[f]} ${GIVEN[g]}${dup}`,
      kana: `${FAMILY_KANA[f]} ${GIVEN_KANA[g]}`,
      univ: UNIV[i % UNIV.length],
      faculty: FACULTY[(i * 3) % FACULTY.length],
      grade: 3 + (i % 2),
      mail: `s${10001 + i}@example.ac.jp`,
      written: 50 + Math.round(r() * 50),
      interview: 50 + Math.round(r() * 50),
    });
  }
  return out;
}

function makeBookShell(id: string, fileName: string, size: number, sheets: SheetGrid[], accent: string): SourceBook {
  const ext = extOf(fileName).slice(1);
  const format: SourceFormat = (['xlsx', 'xlsm', 'xls', 'csv', 'ods'] as const).includes(ext as 'xlsx') ? (ext as SourceFormat) : 'other';
  const values = format === 'xls' || format === 'csv' || format === 'ods';
  return {
    id,
    fileName,
    size,
    format,
    bytes: new Uint8Array([0x50, 0x4b, 0x03, 0x04]),
    fidelity: values ? 'values' : 'full',
    styles: values ? [{ font: '游ゴシック', size: 11 }] : bookStyles(accent),
    sheets,
    names: [],
    features: { charts: 0, pivotTables: 0, vba: format === 'xlsm', drawings: 0, slicers: 0, sparklineGroups: 0, externalLinks: 0, threadedComments: 0, formControls: 0, tables: 0, images: 0 },
    date1904: false,
    warnings: [],
  };
}

function studentBook(id: string, fileName: string, size: number): SourceBook {
  const st = students();
  const extra = Array.from({ length: extraCols }, (_, k) => `項目${k + 1}`);
  const header = ['No', '学生ID', '氏名', 'フリガナ', '大学', '学部', '学年', 'メール', '筆記', '面接', '合計', ...extra];
  const rows: CellInit[][] = [['2026年度 夏季インターンシップ 応募者一覧'], header];
  st.forEach((s, i) => {
    const r = i + 3;
    rows.push([
      i + 1, s.id, s.name, s.kana, s.univ, s.faculty, s.grade, s.mail, s.written, s.interview,
      { f: `I${r}+J${r}`, v: s.written + s.interview, s: ST.formula },
      ...extra.map((_, k) => (i * 31 + k * 17) % 1000),
    ]);
  });
  const list = mkSheet({
    name: '応募者一覧',
    rows,
    colStyles: [ST.score, ST.id, ST.text, ST.text, ST.text, ST.text, ST.score, ST.text, ST.score, ST.score, ST.formula, ...extra.map(() => ST.num)],
    rowStyles: { 1: ST.title, 2: ST.head },
    widths: [5.63, 9.63, 13.63, 18.63, 15.63, 12.63, 5.63, 26.63, 6.63, 6.63, 7.63, ...extra.map(() => 8.63)],
    freeze: { cols: 3, rows: 2 },
    merges: titleMerge ? [{ r1: 1, c1: 1, r2: 1, c2: 11 }] : [],
  });
  const sum = (fn: string, col: string) => `${fn}(応募者一覧!${col}3:${col}${nStudents + 2})`;
  const avg = st.reduce((a, s) => a + s.written, 0) / st.length;
  const max = Math.max(...st.map((s) => s.written + s.interview));
  const agg = mkSheet({
    name: '集計',
    rows: [
      ['項目', '値'],
      ['応募者数', { f: sum('COUNTA', 'B'), v: nStudents }],
      ['筆記の平均', { f: sum('AVERAGE', 'I'), v: Math.round(avg * 10) / 10 }],
      ['合計の最高点', { f: sum('MAX', 'K'), v: max }],
      ['面接の合格ライン', { f: `${sum('AVERAGE', 'J')}+5`, v: 80 }],
    ],
    colStyles: [ST.text, ST.num],
    rowStyles: { 1: ST.head },
    widths: [18, 10],
  });
  const book = makeBookShell(id, fileName, size, [list, agg], '#DDEBF7');
  book.features.charts = 1;
  book.features.drawings = 1;
  bookMeta.set(book, {
    dataSheet: '応募者一覧',
    regions: {
      応募者一覧: { sheet: '応募者一覧', headerRow: 2, firstDataRow: 3, lastDataRow: nStudents + 2, firstCol: 1, lastCol: 11 + extraCols },
      集計: { sheet: '集計', headerRow: 1, firstDataRow: 2, lastDataRow: 5, firstCol: 1, lastCol: 2 },
    },
  });
  return book;
}

function workplaceBook(id: string, fileName: string, size: number): SourceBook {
  const st = students();
  const r = rng(77);
  const header = ['学生ID', '氏名', '配属先', '評価', '面接日', 'コメント'];
  const rows: CellInit[][] = [header];
  const grades = ['A', 'B', 'B', 'C', 'A'];
  const comments = ['積極的に質問していた', '論理的に説明できる', '', 'チームでの経験が豊富', '志望動機が明確', ''];
  st.forEach((s, i) => {
    if (i === 6 || i === 19) return; // not interviewed (row 22 is in no other file at all)
    const name = i === 13 ? s.name.replace('高', '髙') : s.name; // 異体字
    rows.push([s.id, name, WORKPLACE[i % WORKPLACE.length], grades[Math.floor(r() * grades.length)], 46300 + (i % 5), comments[i % comments.length]]);
    if (i === 11) rows.push([s.id, name, WORKPLACE[(i + 1) % WORKPLACE.length], 'B', 46302, '2回目の面接']); // duplicate
  });
  rows.push(['10031', '岡田 美月', '人事部', 'A', 46303, '追加の応募者']);
  const sheet = mkSheet({
    name: '回答',
    rows,
    colStyles: [ST.id, ST.text, ST.text, ST.text, ST.date, ST.note],
    rowStyles: { 1: ST.head },
    widths: [9.63, 13.63, 16.63, 6.63, 11.63, 24.63],
    freeze: { cols: 0, rows: 1 },
  });
  const book = makeBookShell(id, fileName, size, [sheet], '#E2EFDA');
  bookMeta.set(book, { dataSheet: '回答', regions: { 回答: { sheet: '回答', headerRow: 1, firstDataRow: 2, lastDataRow: rows.length, firstCol: 1, lastCol: 6 } } });
  return book;
}

function hrBook(id: string, fileName: string, size: number): SourceBook {
  const st = students();
  const header = ['学籍番号', '名前', '大学名', '判定', '備考'];
  const rows: CellInit[][] = [header];
  const verdicts = ['合格', '保留', '合格', '不合格', '合格'];
  st.forEach((s, i) => {
    if (i === 19 || i === 20) return; // no answer yet
    let name = i % 7 === 3 ? `${s.name} 様` : s.name;
    if (i === 25) name = `${FAMILY[i]} 葵`; // typo -> ID matches, name differs
    const univ = UNIV_ABBR[s.univ] && i % 3 === 0 ? UNIV_ABBR[s.univ] : s.univ;
    rows.push([Number(s.id), name, univ, verdicts[i % verdicts.length], i % 6 === 0 ? '最終面接へ' : '']);
  });
  const sheet = mkSheet({
    name: 'Sheet1',
    rows,
    colStyles: [ST.score, ST.text, ST.text, ST.text, ST.note],
    rowStyles: { 1: ST.head },
    widths: [9.63, 15.63, 14.63, 7.63, 18.63],
  });
  const book = makeBookShell(id, fileName, size, [sheet], '#FCE4D6');
  bookMeta.set(book, { dataSheet: 'Sheet1', regions: { Sheet1: { sheet: 'Sheet1', headerRow: 1, firstDataRow: 2, lastDataRow: rows.length, firstCol: 1, lastCol: 5 } } });
  return book;
}

function genericBook(id: string, fileName: string, size: number): SourceBook {
  const st = students();
  const header = ['ID', '氏名', '部署', '点数', '備考'];
  const rows: CellInit[][] = [header];
  st.slice(0, 18).forEach((s, i) => rows.push([s.id, s.name, WORKPLACE[(i * 5) % WORKPLACE.length], 60 + ((i * 7) % 40), '']));
  const sheet = mkSheet({ name: 'データ', rows, colStyles: [ST.id, ST.text, ST.text, ST.score, ST.note], rowStyles: { 1: ST.head }, widths: [9, 13, 16, 7, 16] });
  const book = makeBookShell(id, fileName, size, [sheet], '#EDEDED');
  bookMeta.set(book, { dataSheet: 'データ', regions: { データ: { sheet: 'データ', headerRow: 1, firstDataRow: 2, lastDataRow: rows.length, firstCol: 1, lastCol: 5 } } });
  return book;
}

function matchingBook(id: string, fileName: string, size: number): SourceBook {
  const st = students();
  const rows: CellInit[][] = [['No', '学生ID', '学生名', '職場ID', '職場名']];
  st.forEach((s, i) => rows.push([i + 1, s.id, s.name, String(8000000001 + (i % 6)), `職場担当${'ABCDEF'[i % 6]}`]));
  const sheet = mkSheet({ name: 'マッチング', rows, colStyles: [ST.score, ST.id, ST.text, ST.id, ST.text], rowStyles: { 1: ST.head }, widths: [5, 10, 13, 12, 12] });
  const book = makeBookShell(id, fileName, size, [sheet], '#EDEDED');
  bookMeta.set(book, { dataSheet: 'マッチング', regions: { マッチング: { sheet: 'マッチング', headerRow: 1, firstDataRow: 2, lastDataRow: rows.length, firstCol: 1, lastCol: 5 } } });
  return book;
}

/** A list whose columns the key detection does not recognise (「参加者」 holds the names, but is no 氏名 header). */
function noKeyBook(id: string, fileName: string, size: number): SourceBook {
  const st = students();
  const header = ['番号', '参加者', '班', '点数'];
  const rows: CellInit[][] = [header];
  st.slice(0, 20).forEach((s, i) => rows.push([i + 1, s.name, `第${(i % 4) + 1}班`, 60 + ((i * 11) % 40)]));
  const sheet = mkSheet({ name: '参加状況', rows, colStyles: [ST.score, ST.text, ST.text, ST.score], rowStyles: { 1: ST.head }, widths: [6, 14, 8, 7] });
  const book = makeBookShell(id, fileName, size, [sheet], '#EDEDED');
  bookMeta.set(book, { dataSheet: '参加状況', regions: { 参加状況: { sheet: '参加状況', headerRow: 1, firstDataRow: 2, lastDataRow: rows.length, firstCol: 1, lastCol: 4 } } });
  return book;
}

function bookFor(id: string, fileName: string, size: number): SourceBook {
  const n = fileName.normalize('NFKC');
  if (/broken|壊れ/i.test(n)) throw new Error('Corrupted zip: end of central directory record signature not found');
  if (/password|パスワード/i.test(n)) throw new Error('File is password-protected (encrypted package)');
  if (/キーなし|nokey/i.test(n)) return noKeyBook(id, fileName, size);
  if (/マッチング/.test(n)) return matchingBook(id, fileName, size);
  if (/学生|応募|student/i.test(n)) return studentBook(id, fileName, size);
  if (/職場|workplace/i.test(n)) return workplaceBook(id, fileName, size);
  if (/人事|hr/i.test(n)) return hrBook(id, fileName, size);
  return genericBook(id, fileName, size);
}

function sheetOf(book: SourceBook, name: string): SheetGrid | undefined {
  return book.sheets.find((s) => s.name === name);
}

function mockCellText(book: SourceBook, cell: Cell | undefined): string {
  if (!cell || cell.v === null || cell.v === undefined) return '';
  try {
    return displayValue(cell, book.styles[cell.s ?? 0]?.numFmt).text;
  } catch {
    return String(cell.v);
  }
}

function rawText(cell: Cell | undefined): string {
  if (!cell || cell.v === null || cell.v === undefined) return '';
  if (typeof cell.v === 'number') return Number.isInteger(cell.v) ? String(cell.v) : String(cell.v);
  return String(cell.v);
}

function detectRegionGeneric(book: SourceBook, sheetName: string, opts?: { headerRow?: number; lastDataRow?: number }): TableRegion {
  const sh = sheetOf(book, sheetName);
  let auto: TableRegion;
  try {
    auto = detectAuto(book, sheetName);
  } catch (e) {
    if (opts?.headerRow === undefined || !sh) throw e;
    auto = { sheet: sheetName, headerRow: 0, firstDataRow: 1, lastDataRow: 0, firstCol: 1, lastCol: Math.max(1, sh.maxCol) };
  }
  let region = auto;
  if (opts?.headerRow !== undefined && opts.headerRow !== auto.headerRow && sh) {
    // columns from the chosen row, data below it until the sheet ends (or the auto table end when below)
    const hr = Math.max(1, Math.round(opts.headerRow));
    const row = sh.rows[hr - 1] ?? [];
    let first = 0;
    let last = 0;
    row.forEach((c, i) => {
      if (c && c.v !== null && c.v !== '') {
        if (!first) first = i + 1;
        last = i + 1;
      }
    });
    const lastData = auto.lastDataRow > hr ? auto.lastDataRow : Math.max(hr, sh.maxRow);
    region = { sheet: sheetName, headerRow: hr, firstDataRow: hr + 1, lastDataRow: lastData, firstCol: first || auto.firstCol, lastCol: last || auto.lastCol };
  }
  if (opts?.lastDataRow !== undefined && sh) {
    const want = Math.max(region.firstDataRow, Math.min(Math.max(sh.maxRow, region.firstDataRow), Math.round(opts.lastDataRow)));
    region = { ...region, lastDataRow: want };
  }
  return region;
}

function detectAuto(book: SourceBook, sheetName: string): TableRegion {
  const meta = bookMeta.get(book)?.regions[sheetName];
  if (meta) return { ...meta };
  const sh = sheetOf(book, sheetName);
  if (!sh || sh.maxRow === 0) throw new Error('データが見つかりません');
  for (let r = 1; r <= Math.min(sh.maxRow, 30); r++) {
    const row = sh.rows[r - 1] ?? [];
    const texts = row.filter((c) => c && c.t === 's' && String(c.v).trim() !== '').length;
    if (texts >= 2) {
      let first = 0;
      let last = 0;
      row.forEach((c, i) => {
        if (c && c.v !== null && c.v !== '') {
          if (!first) first = i + 1;
          last = i + 1;
        }
      });
      return { sheet: sheetName, headerRow: r, firstDataRow: r + 1, lastDataRow: Math.max(r + 1, sh.maxRow), firstCol: first, lastCol: last };
    }
  }
  throw new Error('見出し行が見つかりません');
}

// =============================================================================================
// keys (a tiny but honest implementation: real value comparison with the normalisation options)

type Kind = 'id' | 'name' | 'univ' | 'mail';

const ITAIJI: Record<string, string> = { 髙: '高', 﨑: '崎', 邊: '辺', 邉: '辺', 齋: '斎', 齊: '斉', 濵: '浜', 德: '徳', 櫻: '桜', 澤: '沢', 廣: '広', 國: '国' };

function classify(header: string): Kind | null {
  const h = header.normalize('NFKC').replace(/\s/g, '');
  if (/^(学生ID|学籍番号|社員番号|受験番号|ID|会員番号|応募番号)$/i.test(h)) return 'id';
  if (/^(氏名|名前|学生名|お名前|氏名漢字)$/.test(h)) return 'name';
  if (/大学/.test(h)) return 'univ';
  if (/メール|mail/i.test(h)) return 'mail';
  return null;
}

const KIND_LABEL: Record<Kind, string> = { id: 'ID', name: '氏名', univ: '大学', mail: 'メール' };

function normalize(raw: string, kind: Kind, o: NormalizeOptions): string {
  let s = raw.normalize('NFKC').trim();
  if (!s) return '';
  if (kind === 'id') {
    s = s.replace(/\s/g, '').toUpperCase();
    if (o.leadingZeros) s = s.replace(/^0+(?=\d)/, '');
    return s;
  }
  if (kind === 'mail') return s.toLowerCase();
  if (kind === 'name') {
    if (o.maidenName) s = s.replace(/[（(][^）)]*[）)]/g, '');
    if (o.honorifics) s = s.replace(/\s*(様|さん|殿|君|くん|ちゃん)$/, '');
    s = s.replace(/[\s\u3000]/g, '');
    if (o.itaiji) s = [...s].map((ch) => ITAIJI[ch] ?? ch).join('');
    return s;
  }
  // univ
  s = s.replace(/\s/g, '');
  if (o.universityAbbrev) {
    for (const [full, abbr] of Object.entries(UNIV_ABBR)) if (s === abbr || s === full.replace('大学', '大')) s = full;
  }
  return s;
}

interface Profile {
  src: KeySource;
  sheet: SheetGrid;
  cols: Array<{ col: number; header: string; kind: Kind | null }>;
  byKind: Partial<Record<Kind, number>>;
  rows: number[];
}

function profile(src: KeySource): Profile {
  const sheet = sheetOf(src.book, src.region.sheet) ?? src.book.sheets[0];
  const hdr = sheet.rows[src.region.headerRow - 1] ?? [];
  const cols: Profile['cols'] = [];
  const byKind: Profile['byKind'] = {};
  for (let c = src.region.firstCol; c <= src.region.lastCol; c++) {
    const header = rawText(hdr[c - 1]);
    const kind = header ? classify(header) : null;
    cols.push({ col: c, header, kind });
    if (kind && byKind[kind] === undefined) byKind[kind] = c;
  }
  const rows: number[] = [];
  for (let r = src.region.firstDataRow; r <= src.region.lastDataRow; r++) rows.push(r);
  return { src, sheet, cols, byKind, rows };
}

function keyOf(p: Profile, row: number, kinds: Kind[], o: NormalizeOptions): string {
  const parts: string[] = [];
  for (const k of kinds) {
    const c = p.byKind[k];
    if (c === undefined) return '';
    const v = normalize(rawText(p.sheet.rows[row - 1]?.[c - 1]), k, o);
    if (!v) return '';
    parts.push(v);
  }
  return parts.join('\u0001');
}

function columnDTOs(p: Profile, o: NormalizeOptions): ColumnDTO[] {
  return p.cols.map(({ col, header, kind }) => {
    const vals = p.rows.map((r) => rawText(p.sheet.rows[r - 1]?.[col - 1]));
    const filled = vals.filter((v) => v !== '');
    const norm = kind ? filled.map((v) => normalize(v, kind, o)) : filled;
    const uniq = new Set(norm).size;
    const counts = new Map<string, number>();
    for (const v of norm) counts.set(v, (counts.get(v) ?? 0) + 1);
    let dupRows = 0;
    for (const n of counts.values()) if (n > 1) dupRows += n;
    const typeLabel = kind ? KIND_LABEL[kind] : filled.every((v) => /^-?\d+(\.\d+)?$/.test(v)) && filled.length ? '数値' : 'テキスト';
    return {
      col,
      header,
      type: kind ?? (typeLabel === '数値' ? 'number' : 'text'),
      typeLabel,
      confidence: kind ? '高' : '中',
      typeScore: kind ? 0.95 : 0.5,
      keyScore: kind === 'id' ? 0.95 : kind === 'name' ? 0.8 : kind ? 0.4 : 0.1,
      uniqueness: filled.length ? uniq / filled.length : 0,
      fill: vals.length ? filled.length / vals.length : 0,
      blank: vals.length - filled.length,
      dupRows,
      notes: [],
    };
  });
}

function sheetDTO(p: Profile, o: NormalizeOptions): SheetDTO {
  return { sourceId: p.src.id, file: p.src.book.fileName, sheet: p.src.region.sheet, headerRow: p.src.region.headerRow, rows: p.rows.length, columns: columnDTOs(p, o) };
}

interface CandDef {
  id: string;
  kinds: Kind[];
  kind: CandidateDTO['kind'];
  family: CandidateDTO['family'];
  label: string;
  /** a key the user built from columns (addManualCandidate): sheet column numbers of the base / of the other file */
  manual?: { baseCols: number[]; otherCols: number[] };
}

/** normalisation of a manually chosen column without a detected kind */
function genericNormalize(raw: string): string {
  return raw.normalize('NFKC').trim().replace(/[\s　]+/g, '').toLowerCase();
}

function keyOfCols(p: Profile, row: number, cols: number[], o: NormalizeOptions): string {
  const parts: string[] = [];
  for (const col of cols) {
    const raw = rawText(p.sheet.rows[row - 1]?.[col - 1]);
    const kind = p.cols.find((c) => c.col === col)?.kind ?? null;
    const v = kind ? normalize(raw, kind, o) : genericNormalize(raw);
    if (!v) return '';
    parts.push(v);
  }
  return parts.join('\u0001');
}

const defCols = (d: CandDef, side: 'base' | 'other', p: Profile): number[] =>
  d.manual ? (side === 'base' ? d.manual.baseCols : d.manual.otherCols) : d.kinds.map((k) => p.byKind[k]!);

function defKeyOf(p: Profile, row: number, d: CandDef, side: 'base' | 'other', o: NormalizeOptions): string {
  return d.manual ? keyOfCols(p, row, defCols(d, side, p), o) : keyOf(p, row, d.kinds, o);
}

function defKeyMap(p: Profile, d: CandDef, side: 'base' | 'other', o: NormalizeOptions): Map<string, number[]> {
  const m = new Map<string, number[]>();
  for (const r of p.rows) {
    const k = defKeyOf(p, r, d, side, o);
    if (!k) continue;
    const a = m.get(k);
    if (a) a.push(r);
    else m.set(k, [r]);
  }
  return m;
}

function candidateDefs(base: Profile, other: Profile): CandDef[] {
  const out: CandDef[] = [];
  const has = (k: Kind) => base.byKind[k] !== undefined && other.byKind[k] !== undefined;
  const hdr = (p: Profile, k: Kind) => p.cols.find((c) => c.col === p.byKind[k])?.header ?? KIND_LABEL[k];
  const pairLabel = (k: Kind) => {
    const a = hdr(base, k);
    const b = hdr(other, k);
    return a === b ? a : `${a} ⟷ ${b}`;
  };
  if (has('id')) out.push({ id: 'id', kinds: ['id'], kind: 'single', family: 'id', label: `ID（${pairLabel('id')}）` });
  if (has('name')) out.push({ id: 'name', kinds: ['name'], kind: 'single', family: 'name', label: `氏名（${pairLabel('name')}）` });
  if (has('name') && has('univ')) out.push({ id: 'name+univ', kinds: ['name', 'univ'], kind: 'composite', family: 'name', label: '氏名＋大学' });
  if (has('mail')) out.push({ id: 'mail', kinds: ['mail'], kind: 'discovered', family: 'other', label: `メール（${pairLabel('mail')}）` });
  return out;
}

function evalCandidate(def: CandDef, base: Profile, other: Profile, o: NormalizeOptions, tier: number | undefined): CandidateDTO {
  const bm = defKeyMap(base, def, 'base', o);
  const om = defKeyMap(other, def, 'other', o);
  let matchedOther = 0;
  let nonBlankOther = 0;
  for (const [k, rows] of om) {
    nonBlankOther += rows.length;
    if (bm.has(k)) matchedOther += rows.length;
  }
  let matchedBase = 0;
  let nonBlankBase = 0;
  for (const [k, rows] of bm) {
    nonBlankBase += rows.length;
    if (om.has(k)) matchedBase += rows.length;
  }
  const dup = (m: Map<string, number[]>) => [...m.values()].reduce((n, a) => n + (a.length > 1 ? a.length : 0), 0);
  const matchRate = nonBlankOther ? matchedOther / nonBlankOther : 0;
  const coverage = nonBlankBase ? matchedBase / nonBlankBase : 0;
  const uniqBase = nonBlankBase ? bm.size / nonBlankBase : 0;
  const uniqOther = nonBlankOther ? om.size / nonBlankOther : 0;
  const verdict = matchRate >= 0.9 ? '良好' : matchRate >= 0.5 ? '一部一致' : matchRate > 0.1 ? '不一致多数' : '使用不可';
  const colsOf = (p: Profile, side: 'base' | 'other') => defCols(def, side, p).map((col) => ({ col, header: p.cols.find((c) => c.col === col)?.header ?? '' }));
  const dupB = dup(bm);
  const dupO = dup(om);
  return {
    id: def.id,
    kind: def.kind,
    label: def.label,
    family: def.family,
    baseColumns: colsOf(base, 'base'),
    otherColumns: colsOf(other, 'other'),
    matchRate,
    coverage,
    matchedOther,
    nonBlankOther,
    matchedBase,
    nonBlankBase,
    uniqBase,
    uniqOther,
    dupRowsBase: dupB,
    dupRowsOther: dupO,
    blankBase: base.rows.length - nonBlankBase,
    blankOther: other.rows.length - nonBlankOther,
    ambiguous: 0,
    cardinality: dupB === 0 && dupO === 0 ? '1:1' : dupB === 0 ? '1:N' : 'N:1',
    score: Math.round((matchRate * 60 + uniqBase * 20 + uniqOther * 20) * 10) / 10,
    verdict,
    selected: tier !== undefined,
    tier,
  };
}

interface MockInternal {
  base: Profile;
  others: Record<string, Profile>;
  defs: Record<string, CandDef[]>;
}

function analyzeSources(base: KeySource, others: KeySource[], options?: Partial<NormalizeOptions>): KeyAnalysis {
  const o: NormalizeOptions = { ...DEFAULT_OPTIONS, ...options };
  const bp = profile(base);
  const pairs: Record<string, PairAnalysis> = {};
  const internal: MockInternal = { base: bp, others: {}, defs: {} };
  for (const src of others) {
    const op = profile(src);
    const defs = candidateDefs(bp, op);
    internal.others[src.id] = op;
    internal.defs[src.id] = defs;
    // auto plan: ID then 氏名 when they are usable
    const prelim = defs.map((d) => evalCandidate(d, bp, op, o, undefined));
    let tier = 1;
    const tiers = new Map<string, number>();
    for (const id of ['id', 'name']) {
      const c = prelim.find((x) => x.id === id);
      if (c && c.matchRate >= 0.5) tiers.set(id, tier++);
    }
    if (tiers.size === 0) {
      const best = [...prelim].sort((a, b) => b.score - a.score)[0];
      if (best && best.matchRate > 0.1) tiers.set(best.id, 1);
    }
    const candidates = defs.map((d) => evalCandidate(d, bp, op, o, tiers.get(d.id)));
    const first = candidates.find((c) => c.tier === 1);
    pairs[src.id] = { sourceId: src.id, other: sheetDTO(op, o), mode: first && first.dupRowsBase > 0 ? 'lookup' : 'oneToOne', candidates };
  }
  const chips: KeyChip[] = (['id', 'name'] as const).map((family) => {
    const perSource: KeyChip['perSource'] = {};
    for (const s of Object.keys(pairs)) {
      const c = pairs[s].candidates.filter((x) => x.family === family && x.kind === 'single')[0];
      perSource[s] = c ? { candidateId: c.id, label: c.label, matchRate: c.matchRate } : null;
    }
    return { family, label: family === 'id' ? 'ID' : '氏名', perSource };
  });
  return { baseId: base.id, base: sheetDTO(bp, o), pairs, chips, options: o, internal };
}

function defaultPlan(analysis: KeyAnalysis): JoinPlan {
  const tiers: JoinPlan['tiers'] = {};
  for (const [s, pair] of Object.entries(analysis.pairs)) {
    tiers[s] = pair.candidates.filter((c) => c.tier !== undefined).sort((a, b) => (a.tier ?? 0) - (b.tier ?? 0)).map((c) => c.id);
  }
  return { tiers, joinType: 'left', duplicatePolicy: 'first', manualLinks: {} };
}

/** keys.addManualCandidate for the mock: a candidate built from chosen columns, with real match statistics. */
function addManualCandidate(analysis: KeyAnalysis, sourceId: string, baseCols: number[], otherCols: number[]): KeyAnalysis {
  const internal = analysis.internal as MockInternal | undefined;
  const pair = analysis.pairs[sourceId];
  const other = internal?.others[sourceId];
  if (!internal || !pair || !other) throw new Error('このファイルの分析結果がありません。ファイルを読み込み直してください');
  if (!baseCols.length || !otherCols.length) throw new Error('基準ファイルとこのファイルの列を選んでください');
  if (baseCols.length > 3 || otherCols.length > 3) throw new Error('組み合わせに使える列は 3 列までです');
  if (baseCols.length !== otherCols.length) throw new Error('基準ファイルとこのファイルで同じ数の列を選んでください');
  const baseCol = (c: number) => internal.base.cols.find((x) => x.col === c);
  const otherCol = (c: number) => other.cols.find((x) => x.col === c);
  if (baseCols.some((c) => !baseCol(c))) throw new Error('基準ファイルの表にない列が選ばれています');
  if (otherCols.some((c) => !otherCol(c))) throw new Error('このファイルの表にない列が選ばれています');
  if (new Set(baseCols).size < baseCols.length || new Set(otherCols).size < otherCols.length) throw new Error('同じ列が重複して選ばれています');
  const id = `manual:${baseCols.join('+')}|${otherCols.join('+')}`;
  if (pair.candidates.some((c) => c.id === id)) throw new Error('同じ組み合わせがすでに候補にあります');
  const headers = (cols: number[], get: (c: number) => { header: string } | undefined) => cols.map((c) => get(c)?.header || `${colLetter(c)}列`).join('＋');
  const hb = headers(baseCols, baseCol);
  const ho = headers(otherCols, otherCol);
  const kindOf = baseCol(baseCols[0])?.kind;
  const def: CandDef = {
    id,
    kinds: [],
    kind: baseCols.length > 1 ? 'composite' : 'single',
    family: kindOf === 'id' ? 'id' : kindOf === 'name' ? 'name' : 'other',
    label: hb === ho ? hb : `${hb} ⟷ ${ho}`,
    manual: { baseCols: baseCols.slice(), otherCols: otherCols.slice() },
  };
  const cand = evalCandidate(def, internal.base, other, analysis.options, undefined);
  return {
    ...analysis,
    pairs: { ...analysis.pairs, [sourceId]: { ...pair, candidates: [...pair.candidates, cand] } },
    internal: { ...internal, defs: { ...internal.defs, [sourceId]: [...(internal.defs[sourceId] ?? []), def] } } satisfies MockInternal,
  };
}

function rowRef(p: Profile, row: number): RowRef {
  const preview: string[] = [];
  for (const k of ['id', 'name', 'univ'] as Kind[]) {
    const c = p.byKind[k];
    if (c !== undefined) preview.push(rawText(p.sheet.rows[row - 1]?.[c - 1]));
  }
  if (!preview.some(Boolean)) for (const c of p.cols.slice(0, 3)) preview.push(rawText(p.sheet.rows[row - 1]?.[c.col - 1]));
  return { excelRow: row, preview };
}

function joinSources(analysis: KeyAnalysis, plan: JoinPlan): JoinResult {
  const internal = analysis.internal as MockInternal | undefined;
  if (!internal) throw new Error('分析結果が古くなっています。もう一度読み込んでください');
  const o = analysis.options;
  const base = internal.base;
  const reports: Record<string, PairJoinReport> = {};
  const matchesByRow = new Map<number, Record<string, RowMatch | null>>();
  for (const r of base.rows) matchesByRow.set(r, {});
  const appended: MergedRowPlan[] = [];
  const keyFill: JoinResult['keyFill'] = {};
  const keyPairs: NonNullable<JoinResult['keyPairs']> = {};
  const defaultImports: ImportColumnSpec[] = [];
  const items: JoinResult['items'] = [];

  for (const [s, other] of Object.entries(internal.others)) {
    const defs = internal.defs[s] ?? [];
    const tierDefs = (plan.tiers[s] ?? []).map((id) => defs.find((d) => d.id === id)).filter((d): d is CandDef => !!d);
    const maps = tierDefs.map((d) => ({ d, om: defKeyMap(other, d, 'other', o), bm: defKeyMap(base, d, 'base', o) }));
    const used = new Set<number>();
    const perTier: Record<string, number> = {};
    const conflicts: PairJoinReport['conflicts'] = [];
    const ambiguous: PairJoinReport['ambiguous'] = [];
    let variants = 0;
    const variantExamples: string[] = [];
    let matched = 0;
    const links = new Map((plan.manualLinks[s] ?? []).map((l) => [l.baseRow, l.otherRow]));
    for (const br of base.rows) {
      let m: RowMatch | null = null;
      const linked = links.get(br);
      if (linked !== undefined && !used.has(linked)) {
        m = { row: linked, via: '手動', flags: ['manual'] };
      } else {
        for (let t = 0; t < maps.length && !m; t++) {
          const { d, om } = maps[t];
          const k = defKeyOf(base, br, d, 'base', o);
          if (!k) continue;
          const cands = (om.get(k) ?? []).filter((r) => !used.has(r) || plan.duplicatePolicy !== 'skip');
          if (!cands.length) continue;
          if (cands.length > 1 && plan.duplicatePolicy === 'skip') {
            ambiguous.push({ base: rowRef(base, br), candidates: cands.map((r) => rowRef(other, r)), tier: t + 1 });
            continue;
          }
          const row = plan.duplicatePolicy === 'last' ? cands[cands.length - 1] : cands[0];
          const flags: RowMatch['flags'] = [];
          if (t > 0) flags.push('fallback');
          if (cands.length > 1) flags.push('duplicateInSource');
          const rawB = defCols(d, 'base', base).map((c) => rawText(base.sheet.rows[br - 1]?.[c - 1]).normalize('NFKC').trim()).join(' ');
          const rawO = defCols(d, 'other', other).map((c) => rawText(other.sheet.rows[row - 1]?.[c - 1]).normalize('NFKC').trim()).join(' ');
          if (rawB !== rawO && (d.kinds.includes('name') || (d.manual && d.family === 'name'))) {
            flags.push('variant');
            variants++;
            if (variantExamples.length < 3) variantExamples.push(`${br}行目 ⟷ ${row}行目: ${rawB} / ${rawO}`);
          }
          // ID matched but the names disagree
          if (d.family === 'id' && base.byKind.name !== undefined && other.byKind.name !== undefined) {
            const nb = keyOf(base, br, ['name'], o);
            const no = keyOf(other, row, ['name'], o);
            if (nb && no && nb !== no) {
              flags.push('conflict');
              conflicts.push({
                base: rowRef(base, br),
                other: rowRef(other, row),
                message: `ID は一致・氏名が異なる（${rawText(base.sheet.rows[br - 1]?.[base.byKind.name - 1])} / ${rawText(other.sheet.rows[row - 1]?.[other.byKind.name - 1])}）`,
              });
            }
          }
          m = { row, via: d.kind === 'composite' ? d.label : d.family === 'id' ? 'ID' : d.family === 'name' ? '氏名' : d.label, flags };
          perTier[m.via] = (perTier[m.via] ?? 0) + 1;
        }
      }
      if (m) {
        used.add(m.row);
        matched++;
      }
      matchesByRow.get(br)![s] = m;
    }
    const unmatchedBase = base.rows.filter((r) => !matchesByRow.get(r)![s]).map((r) => rowRef(base, r));
    const unmatchedOtherRows = other.rows.filter((r) => !used.has(r));
    const unmatchedOther = unmatchedOtherRows.map((r) => rowRef(other, r));
    const dups = (m: Map<string, number[]>, p: Profile) =>
      [...m.values()].filter((a) => a.length > 1).map((a) => ({ key: rowRef(p, a[0]).preview.filter(Boolean).slice(0, 2).join(' '), count: a.length }));
    const first = maps[0];
    const suggestions: PairJoinReport['suggestions'] = [];
    for (const ub of unmatchedBase) {
      const bn = keyOf(base, ub.excelRow, ['name'], o);
      if (!bn) continue;
      for (const uo of unmatchedOther) {
        const on = keyOf(other, uo.excelRow, ['name'], o);
        if (on && on.slice(0, 2) === bn.slice(0, 2)) suggestions.push({ base: ub, other: uo, score: 0.72, reasons: ['姓が一致', '学年・大学が近い'] });
      }
    }
    const flags: PairJoinReport['flags'] = {};
    if (variants) flags['表記ゆれを補正して一致'] = { level: 'info', count: variants, examples: variantExamples };
    const fallback = base.rows.filter((r) => matchesByRow.get(r)![s]?.flags.includes('fallback')).length;
    if (fallback) flags['優先度の低いキーで一致（要確認）'] = { level: 'warning', count: fallback, examples: base.rows.filter((r) => matchesByRow.get(r)![s]?.flags.includes('fallback')).slice(0, 3).map((r) => `${r}行目`) };
    reports[s] = {
      sourceId: s,
      mode: analysis.pairs[s]?.mode ?? 'oneToOne',
      plan: tierDefs.map((d, i) => ({ order: i + 1, name: d.label })),
      matched,
      perTier,
      unmatchedBase,
      unmatchedOther,
      duplicatesBase: first ? dups(first.bm, base) : [],
      duplicatesOther: first ? dups(first.om, other) : [],
      ambiguous,
      conflicts,
      flags,
      suggestions: suggestions.slice(0, 10),
    };
    if (plan.joinType === 'full') {
      for (const u of unmatchedOtherRows) {
        const matches: Record<string, RowMatch | null> = {};
        for (const s2 of Object.keys(internal.others)) matches[s2] = null;
        matches[s] = { row: u, via: '追加', flags: [] };
        appended.push({ baseRow: null, appendedFrom: { sourceId: s, row: u }, matches });
      }
    }
    const fill: Array<{ baseCol: number; sourceCol: number }> = [];
    for (const k of ['id', 'name'] as Kind[]) {
      if (base.byKind[k] !== undefined && other.byKind[k] !== undefined) fill.push({ baseCol: base.byKind[k]!, sourceCol: other.byKind[k]! });
    }
    keyFill[s] = fill;
    // the direct (base column, source column) pairs of the enabled tiers: here the id / name columns that exist on both sides
    if (tierDefs.length && fill.length) keyPairs[s] = fill.map((p) => ({ ...p }));
    const keyCols = new Set<number>();
    for (const d of tierDefs) for (const c of defCols(d, 'other', other)) if (c !== undefined) keyCols.add(c);
    for (const c of other.cols) if (!keyCols.has(c.col) && c.header) defaultImports.push({ sourceId: s, col: c.col });
    if (unmatchedBase.length) items.push({ severity: 'warning', code: 'unmatched', message: `${other.src.label}: 基準の ${unmatchedBase.length} 行に該当する行がありません`, sourceId: s, rows: unmatchedBase.slice(0, 5).map((r) => r.excelRow) });
    if (unmatchedOtherRows.length) {
      items.push({
        severity: plan.joinType === 'full' ? 'info' : 'warning',
        code: 'unmatchedOther',
        message: `「${other.src.label}」にだけある行: ${unmatchedOtherRows.length}件`,
        sourceId: s,
        otherRows: unmatchedOtherRows.slice(0, 5),
      });
    }
    const dupOtherRows = first ? [...first.om.values()].filter((a) => a.length > 1).flat() : [];
    if (dupOtherRows.length) {
      const dupBase = base.rows.filter((r) => matchesByRow.get(r)![s]?.flags.includes('duplicateInSource'));
      items.push({
        severity: 'warning',
        code: 'duplicate',
        message: `「${other.src.label}」でキーが重複している行: ${dupBase.length || dupOtherRows.length}件（先頭の行を使う）`,
        sourceId: s,
        ...(dupBase.length ? { rows: dupBase.slice(0, 5) } : {}),
        otherRows: dupOtherRows.slice(0, 5),
      });
    }
  }

  const rows: MergedRowPlan[] = base.rows.map((r) => ({ baseRow: r, matches: matchesByRow.get(r)! }));
  const srcIds = Object.keys(internal.others);
  let matchedAny = 0;
  let matchedAll = 0;
  for (const r of base.rows) {
    const m = matchesByRow.get(r)!;
    const n = srcIds.filter((s) => m[s]).length;
    if (n > 0) matchedAny++;
    if (n === srcIds.length) matchedAll++;
  }
  return {
    rows: [...rows, ...appended],
    reports,
    summary: { baseRows: base.rows.length, matchedAny, matchedAll, appended: appended.length },
    items,
    keyFill,
    defaultImports,
    keyPairs,
  };
}

// =============================================================================================
// premerge (builds the merged SourceBook in memory; the "bytes" are a token for readBook)

/** A copy of a sheet of another file for the merged book (plan A): own name, styles shifted into the merged style table, sheet references renamed. */
function copySheetGrid(sh: SheetGrid, name: string, styleOffset: number, renames: ReadonlyMap<string, string>): SheetGrid {
  const rewrite = (f: string): string => {
    let out = f;
    for (const [from, to] of renames) {
      const esc = from.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      out = out.replace(new RegExp(`(^|[^\\w'])'?${esc}'?!`, 'g'), (_m, pre: string) => `${pre}'${to.replace(/'/g, "''")}'!`);
    }
    return out;
  };
  return {
    ...sh,
    name,
    rows: sh.rows.map((row) => row?.map((c) => (c ? { ...c, s: (c.s ?? 0) + styleOffset, ...(c.f ? { f: rewrite(c.f) } : {}) } : c))),
  };
}

async function buildPremerged(input: PremergeInput, onProgress?: ProgressFn): Promise<PremergeResult> {
  onProgress?.({ ratio: 0.1, message: 'スタイルを統合中…' });
  await sleep(250);
  const baseBook = input.base.book;
  const reg = input.base.region;
  const baseSheet = sheetOf(baseBook, reg.sheet)!;
  const styles: XStyle[] = baseBook.styles.slice();
  const styleOffset: Record<string, number> = {};
  const srcById = new Map(input.sources.map((s) => [s.id, s]));
  for (const s of input.sources) {
    styleOffset[s.id] = styles.length;
    styles.push(...s.book.styles);
  }
  const insertAt = reg.lastCol + 1;
  const n = input.imports.length;
  const headerStyle = baseSheet.rows[reg.headerRow - 1]?.[reg.lastCol - 1]?.s ?? 0;
  const existing = new Set<string>();
  for (let c = reg.firstCol; c <= reg.lastCol; c++) existing.add(rawText(baseSheet.rows[reg.headerRow - 1]?.[c - 1]));
  const imported: PremergeLayout['imported'] = [];
  const importCols: Array<{ spec: ImportColumnSpec; src: KeySource; sheet: SheetGrid; header: string }> = [];
  input.imports.forEach((spec, i) => {
    const src = srcById.get(spec.sourceId)!;
    const sheet = sheetOf(src.book, src.region.sheet)!;
    let header = spec.header ?? rawText(sheet.rows[src.region.headerRow - 1]?.[spec.col - 1]);
    if (existing.has(header)) header = `${header}（${src.label}）`;
    existing.add(header);
    imported.push({ col: insertAt + i, sourceId: spec.sourceId, sourceCol: spec.col, header });
    importCols.push({ spec, src, sheet, header });
  });
  onProgress?.({ ratio: 0.4, message: '列を挿入中…' });
  await sleep(250);
  // rows
  const rows: SheetGrid['rows'] = [];
  let toValues = 0;
  /** imported formulas per source (the mock turns every one into a value: it has no formula engine) */
  const formulasBySource: Record<string, number> = {};
  const importedCells =(r: number, plan: MergedRowPlan | undefined): Array<Cell | undefined> =>
    importCols.map(({ spec, src, sheet }) => {
      if (r === reg.headerRow) return { t: 's', v: importCols.find((x) => x.spec === spec)!.header, s: headerStyle };
      const m = plan?.matches[spec.sourceId];
      if (!plan || !m) {
        const typical = sheet.rows[src.region.firstDataRow - 1]?.[spec.col - 1]?.s;
        return typical !== undefined ? { t: 'z', v: null, s: typical + styleOffset[spec.sourceId] } : undefined;
      }
      const c = sheet.rows[m.row - 1]?.[spec.col - 1];
      if (!c) return undefined;
      const out: Cell = { t: c.t, v: c.v, s: (c.s ?? 0) + styleOffset[spec.sourceId] };
      if (c.f) {
        toValues++;
        formulasBySource[spec.sourceId] = (formulasBySource[spec.sourceId] ?? 0) + 1;
      }
      return out;
    });
  const planByBaseRow = new Map<number, MergedRowPlan>();
  for (const p of input.rows) if (p.baseRow !== null) planByBaseRow.set(p.baseRow, p);
  for (let r = 1; r <= baseSheet.maxRow; r++) {
    const src = baseSheet.rows[r - 1] ?? [];
    const left = src.slice(0, insertAt - 1);
    while (left.length < insertAt - 1) left.push(undefined);
    const inRegion = r === reg.headerRow || (r >= reg.firstDataRow && r <= reg.lastDataRow);
    const mid = inRegion ? importedCells(r, planByBaseRow.get(r)) : new Array<Cell | undefined>(n).fill(undefined);
    rows[r - 1] = [...left, ...mid, ...src.slice(insertAt - 1)];
  }
  // appended rows (join 'full'): right after the last data row
  const appended: PremergeLayout['appended'] = [];
  const appendedPlans = input.rows.filter((p) => p.baseRow === null);
  if (appendedPlans.length) {
    const tail = rows.splice(reg.lastDataRow);
    const tmpl = rows[reg.lastDataRow - 1] ?? [];
    appendedPlans.forEach((p, i) => {
      const r = reg.lastDataRow + 1 + i;
      const row: Array<Cell | undefined> = tmpl.map((c) => (c ? { t: 'z', v: null, s: c.s } : undefined));
      const from = p.appendedFrom!;
      const src = srcById.get(from.sourceId)!;
      const sheet = sheetOf(src.book, src.region.sheet)!;
      for (const kf of input.keyFill[from.sourceId] ?? []) {
        const srcRow = sheet.rows[from.row - 1] ?? [];
        let val: { t: Cell['t']; v: Cell['v'] } | null = null;
        if (kf.sourceCols?.length) {
          // base 氏名 <- source 姓 + 名
          const parts = kf.sourceCols.map((c) => rawText(srcRow[c - 1]).trim()).filter(Boolean);
          if (parts.length) val = { t: 's', v: parts.join(kf.joiner ?? '　') };
        } else if (kf.sourceCol !== undefined) {
          const c = srcRow[kf.sourceCol - 1];
          if (c && kf.splitPart !== undefined) {
            // base 姓 / 名 <- one source cell split on whitespace
            const part = rawText(c).trim().split(/[\s　]+/)[kf.splitPart];
            if (part) val = { t: 's', v: part };
          } else if (c) {
            val = { t: c.t, v: c.v };
          }
        }
        if (val) row[kf.baseCol - 1] = { ...val, s: row[kf.baseCol - 1]?.s };
      }
      const imp = importedCells(r, p);
      for (let j = 0; j < n; j++) row[insertAt - 1 + j] = imp[j];
      rows.push(row);
      appended.push({ row: r, sourceId: from.sourceId, sourceRow: from.row });
    });
    rows.push(...tail);
  }
  const colProps = baseSheet.colProps.slice(0, insertAt - 1);
  while (colProps.length < insertAt - 1) colProps.push(undefined);
  const importedProps = importCols.map(({ spec, sheet }) => ({ width: sheet.colProps[spec.col - 1]?.width ?? 10 }));
  const merged: SheetGrid = {
    ...baseSheet,
    rows,
    maxRow: rows.length,
    maxCol: Math.max(baseSheet.maxCol + n, insertAt - 1 + n),
    colProps: [...colProps, ...importedProps, ...baseSheet.colProps.slice(insertAt - 1)],
    features: { ...baseSheet.features },
  };
  onProgress?.({ ratio: 0.8, message: 'パッケージを書き出し中…' });
  await sleep(200);
  // single-sheet output (PremergeOptions.dropOtherSheets): every other sheet goes; merged-sheet formulas that read one become values
  const dropOther = input.options.dropOtherSheets === true;
  const droppedSheets = dropOther ? baseBook.sheets.filter((s) => s.name !== reg.sheet).map((s) => s.name) : [];
  let detached = 0;
  if (droppedSheets.length) {
    const escaped = droppedSheets.map((n) => n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
    const reads = new RegExp(`(^|[^\\w.])'?(${escaped.join('|')})'?!`);
    for (const row of merged.rows) {
      row?.forEach((c, i) => {
        if (c?.f && reads.test(c.f)) {
          const { f: _f, ...rest } = c;
          row[i] = rest;
          detached++;
        }
      });
    }
  }
  // plan A (PremergeOptions.includeSourceSheets): every sheet of every other file is copied after the merged sheet and the base's
  // kept sheets, named <file>_<sheet> by planSheetNames (the real rule); the mock copies cells + styles and renames sheet references
  const includeSource = input.options.includeSourceSheets === true;
  const baseKept = dropOther ? [] : baseBook.sheets.filter((s) => s.name !== reg.sheet);
  const copies: SheetGrid[] = [];
  const addedSheets: NonNullable<PremergeLayout['addedSheets']> = [];
  const copyItems: CheckItem[] = [];
  if (includeSource) {
    const planned = planSheetNames(
      [reg.sheet, ...baseKept.map((s) => s.name)],
      input.sources.map((s) => ({ file: s.book.fileName, sheets: s.book.sheets.map((x) => x.name) })),
    );
    let at = 0;
    for (const src of input.sources) {
      const mine = src.book.sheets.map((sh) => ({ sh, p: planned[at++] }));
      const renames = new Map(mine.map(({ sh, p }) => [sh.name, p.name] as const));
      mine.forEach(({ sh, p }, k) => {
        copies.push(copySheetGrid(sh, p.name, styleOffset[src.id] ?? 0, renames));
        // the mock has no chart copier: a book with charts leaves them out of its last sheet (shows the 警告 entry of the real thing)
        const limited = src.book.features.charts > 0 && k === mine.length - 1;
        addedSheets.push({ sourceId: src.id, file: src.book.fileName, original: sh.name, name: p.name, suffixed: p.suffixed, complete: !limited });
        if (p.suffixed) copyItems.push({ severity: 'info', code: 'sheetRenamed', message: `シート名が重なったため「${p.name}」にしました`, sourceId: src.id });
        if (limited) {
          copyItems.push({
            severity: 'warning',
            code: 'sheetCopyLimits',
            message: `「${p.name}」: グラフ ${src.book.features.charts} 個はコピーできませんでした（セルの値・数式・書式はコピーしました）`,
            sourceId: src.id,
          });
        }
      });
    }
    if (addedSheets.length) {
      const first = addedSheets[0].name;
      copyItems.unshift({ severity: 'info', code: 'sheetsAdded', message: `他のファイルのシート ${addedSheets.length} 枚をコピーしました（${first}${addedSheets.length > 1 ? ' ほか' : ''}）` });
    }
  }
  const book: SourceBook = {
    ...baseBook,
    id: 'merged',
    styles,
    sheets: includeSource ? [merged, ...baseKept, ...copies] : dropOther ? [merged] : baseBook.sheets.map((s) => (s.name === reg.sheet ? merged : s)),
    warnings: [],
  };
  const bytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04, Math.floor(Math.random() * 255)]);
  mergedStore.set(bytes, book);
  const matchedRows = input.rows.filter((p) => p.baseRow !== null && Object.values(p.matches).some(Boolean)).length;
  const baseRows = input.rows.filter((p) => p.baseRow !== null).length;
  onProgress?.({ ratio: 1, message: '完了' });
  return {
    bytes,
    layout: { sheet: reg.sheet, insertAt, imported, appendAt: appended.length ? appended[0].row : null, appended, ...(includeSource ? { addedSheets } : {}) },
    report: {
      mergedRows: input.rows.length,
      matchedRows,
      baseOnlyRows: baseRows - matchedRows,
      appendedRows: appended.length,
      importedColumns: n,
      formulasKept: 0,
      formulasToValues: toValues,
      formulaStats: Object.entries(formulasBySource).map(([sourceId, n]) => ({
        sourceId,
        label: srcById.get(sourceId)?.label ?? sourceId,
        kept: 0,
        toValues: n,
        empty: 0,
        reasons: { unsupported: n },
      })),
      ...(droppedSheets.length ? { droppedSheets, formulasDetached: detached } : {}),
      ...(includeSource ? { addedSheets: addedSheets.map((a) => a.name) } : {}),
      items: copyItems,
    },
  };
}

// =============================================================================================
// doc (editor)

let newColSeq = 1;

function createEditor(
  premerged: SourceBook,
  layout: PremergeLayout,
  sourceInfo: MergedDoc['sourceInfo'],
  region: { headerRow: number; firstDataRow: number; lastDataRow: number },
): EditorSnapshot {
  const sheet = sheetOf(premerged, layout.sheet);
  if (!sheet) throw new Error(`シート「${layout.sheet}」が見つかりません`);
  const n = layout.imported.length;
  const columns: DocColumn[] = [];
  for (let c = 1; c <= sheet.maxCol; c++) {
    const imp = layout.imported.find((x) => x.col === c);
    const origin: DocColumn['origin'] = imp ? { kind: 'import', sourceId: imp.sourceId, col: imp.sourceCol } : { kind: 'base', col: c < layout.insertAt ? c : c - n };
    const id = imp ? `i:${imp.sourceId}:${imp.sourceCol}` : `b${origin.kind === 'base' ? origin.col : c}`;
    const cells: Array<Cell | undefined> = [];
    for (let r = 1; r <= sheet.maxRow; r++) cells[r - 1] = sheet.rows[r - 1]?.[c - 1];
    columns.push({ id, origin, props: { ...(sheet.colProps[c - 1] ?? {}) }, cells });
  }
  const externalFormulas: ExternalFormulaCell[] = [];
  for (const sh of premerged.sheets) {
    if (sh.name === sheet.name) continue;
    sh.rows.forEach((row, ri) => row?.forEach((cell, ci) => cell?.f && externalFormulas.push({ sheet: sh.name, row: ri + 1, col: ci + 1, f: cell.f })));
  }
  const doc: MergedDoc = {
    sheetName: sheet.name,
    sheetNames: premerged.sheets.map((s) => s.name),
    rowCount: sheet.maxRow,
    headerRow: region.headerRow,
    firstDataRow: region.firstDataRow,
    lastDataRow: region.lastDataRow,
    columns,
    rowProps: sheet.rowProps,
    styles: premerged.styles,
    features: sheet.features,
    externalFormulas,
    externalRules: [],
    names: premerged.names,
    sourceInfo,
    version: 1,
  };
  return { doc, log: [], label: '結合' };
}

function adjustAll(doc: MergedDoc, op: StructOp): { columns: DocColumn[]; externalFormulas: ExternalFormulaCell[] } {
  const ctx = { hostSheet: doc.sheetName, targetSheet: doc.sheetName };
  const columns = doc.columns.map((col) => {
    let changed = false;
    const cells = col.cells.map((c) => {
      if (!c?.f) return c;
      const r = adjustFormula(c.f, op, ctx);
      if (!r.changed) return c;
      changed = true;
      return { ...c, f: r.formula, ...(r.createdRefError ? { t: 'e' as const, v: '#REF!' } : null) };
    });
    return changed ? { ...col, cells } : col;
  });
  const externalFormulas = doc.externalFormulas.map((ef) => {
    const r = adjustFormula(ef.f, op, { hostSheet: ef.sheet, targetSheet: doc.sheetName });
    return r.changed ? { ...ef, f: r.formula } : ef;
  });
  return { columns, externalFormulas };
}

function applyStruct(doc: MergedDoc, op: StructOp): MergedDoc {
  const adj = adjustAll(doc, op);
  let cols = adj.columns;
  if (op.type === 'move') {
    const perm = movePerm(op.from, op.count, op.to);
    const out = new Array<DocColumn>(cols.length);
    cols.forEach((c, i) => (out[perm(i + 1) - 1] = c));
    cols = out;
  } else if (op.type === 'insert') {
    const left = cols[op.at - 2];
    const added: DocColumn[] = [];
    for (let k = 0; k < op.count; k++) {
      const cells: Array<Cell | undefined> = new Array(doc.rowCount).fill(undefined);
      if (left) {
        for (let r = doc.firstDataRow; r <= doc.lastDataRow; r++) {
          const s = left.cells[r - 1]?.s;
          if (s !== undefined) cells[r - 1] = { t: 'z', v: null, s };
        }
      }
      added.push({ id: `n${newColSeq++}`, origin: { kind: 'new' }, props: { width: left?.props.width }, cells });
    }
    cols = [...cols.slice(0, op.at - 1), ...added, ...cols.slice(op.at - 1)];
  } else {
    cols = [...cols.slice(0, op.at - 1), ...cols.slice(op.at - 1 + op.count)];
  }
  const merges = adjustMerges(doc.features.merges, op);
  const features = merges === doc.features.merges ? doc.features : { ...doc.features, merges };
  return { ...doc, columns: cols, features, externalFormulas: adj.externalFormulas, version: doc.version + 1 };
}

/** Merged areas after a column op (the op never splits a merge: moveColumns refuses that first). */
function adjustMerges(merges: RangeBox[], op: StructOp): RangeBox[] {
  if (!merges.length || op.axis !== 'col') return merges;
  const out: RangeBox[] = [];
  for (const m of merges) {
    if (op.type === 'move') {
      const perm = movePerm(op.from, op.count, op.to);
      let c1 = Infinity;
      let c2 = 0;
      for (let c = m.c1; c <= m.c2; c++) {
        const n = perm(c);
        c1 = Math.min(c1, n);
        c2 = Math.max(c2, n);
      }
      out.push({ ...m, c1, c2 });
    } else if (op.type === 'insert') {
      if (op.at <= m.c1) out.push({ ...m, c1: m.c1 + op.count, c2: m.c2 + op.count });
      else if (op.at <= m.c2) out.push({ ...m, c2: m.c2 + op.count });
      else out.push(m);
    } else {
      const d1 = op.at;
      const d2 = op.at + op.count - 1;
      const keep = (m.c2 - m.c1 + 1) - Math.max(0, Math.min(m.c2, d2) - Math.max(m.c1, d1) + 1);
      if (keep <= 0) continue;
      const c1 = m.c1 > d2 ? m.c1 - op.count : Math.min(m.c1, d1);
      const next = { ...m, c1, c2: c1 + keep - 1 };
      if (next.c1 === next.c2 && next.r1 === next.r2) continue;
      out.push(next);
    }
  }
  return out;
}

const boxRef = (m: RangeBox) => `${colLetter(m.c1)}${m.r1}:${colLetter(m.c2)}${m.r2}`;

/** Merged areas a move would split (moved partly) or land inside of (like Excel's refusal). */
function mergeObstacles(doc: MergedDoc, sel: number[], to: number): RangeBox[] {
  const moved = new Set(sel);
  const out: RangeBox[] = [];
  for (const m of doc.features.merges) {
    if (m.c1 === m.c2) continue;
    let inSel = 0;
    for (let c = m.c1; c <= m.c2; c++) if (moved.has(c)) inSel++;
    const partial = inSel > 0 && inSel < m.c2 - m.c1 + 1;
    const into = inSel === 0 && to > m.c1 && to <= m.c2;
    if (partial || into) out.push(m);
  }
  return out;
}

const ok = (snapshot: EditorSnapshot, warnings: string[] = []): OpResult => ({ ok: true, snapshot, warnings });

function moveColumns(s: EditorSnapshot, ids: string[], gap: number, opts?: { unmergeTitles?: boolean }): OpResult {
  const order = s.doc.columns.map((c) => c.id);
  const sel = ids.map((id) => order.indexOf(id) + 1).filter((i) => i > 0);
  if (!sel.length) return { ok: false, reason: '移動する列が見つかりません' };
  const to = Math.max(1, Math.min(order.length + 1, gap + 1));
  const ops = planMove(sel, to, 'col');
  if (!ops.length) return ok(s);
  let doc = s.doc;
  const log = s.log.slice();
  const obstacles = mergeObstacles(doc, sel, to);
  if (obstacles.length) {
    const refs = obstacles.map(boxRef);
    const titlesOnly = obstacles.every((m) => m.r2 < doc.headerRow);
    const reason = `結合されたセル（${refs.join('、')}）の一部の列だけを移動することはできません。結合を解除するか、結合範囲の列をまとめて移動してください。`;
    if (!titlesOnly) return { ok: false, reason };
    if (!opts?.unmergeTitles) {
      return {
        ok: false,
        reason,
        fix: {
          kind: 'unmergeTitles',
          refs,
          message: `見出しの上のタイトル（${refs.join('、')}）がセル結合されているため、このままでは移動できません。結合を解除して「選択範囲内で中央」にすると、見た目を保ったまま移動できます。`,
        },
      };
    }
    const drop = new Set(obstacles);
    doc = { ...doc, features: { ...doc.features, merges: doc.features.merges.filter((m) => !drop.has(m)) }, version: doc.version + 1 };
    log.push({ kind: 'unmerge', sheet: doc.sheetName, refs, centerAcross: true });
  }
  for (const op of ops) {
    doc = applyStruct(doc, op);
    log.push({ kind: 'struct', sheet: doc.sheetName, op });
  }
  // like the real doc: the unmerge is reported as a warning of the (successful) move
  const warnings = obstacles.length ? [`タイトル行の結合（${obstacles.map(boxRef).join('、')}）を解除し、「選択範囲内で中央」にしました。`] : [];
  return ok({ doc, log, label: obstacles.length ? '結合を解除して列を移動' : '列を移動' }, warnings);
}

function insertColumns(s: EditorSnapshot, index: number, count: number): OpResult {
  const op: StructOp = { type: 'insert', axis: 'col', at: Math.max(1, Math.min(s.doc.columns.length + 1, index + 1)), count: Math.max(1, count) };
  const doc = applyStruct(s.doc, op);
  return ok({ doc, log: [...s.log, { kind: 'struct', sheet: doc.sheetName, op }], label: '列を挿入' });
}

function deleteOps(s: EditorSnapshot, ids: string[]): StructOp[] {
  const order = s.doc.columns.map((c) => c.id);
  const idx = ids.map((id) => order.indexOf(id) + 1).filter((i) => i > 0).sort((a, b) => a - b);
  const runs: Array<[number, number]> = [];
  for (const i of idx) {
    const last = runs[runs.length - 1];
    if (last && i === last[1] + 1) last[1] = i;
    else runs.push([i, i]);
  }
  return runs.reverse().map(([a, b]) => ({ type: 'delete', axis: 'col', at: a, count: b - a + 1 }));
}

function brokenFormulas(s: EditorSnapshot, ids: string[]): Array<{ where: string; text: string; colIdx?: number; row?: number; sheet?: string }> {
  const ops = deleteOps(s, ids);
  const del = new Set(ids);
  const out: Array<{ where: string; text: string; colIdx?: number; row?: number; sheet?: string }> = [];
  const breaks = (f: string, host: string) => {
    let cur = f;
    for (const op of ops) {
      const r = adjustFormula(cur, op, { hostSheet: host, targetSheet: s.doc.sheetName });
      if (r.createdRefError) return true;
      cur = r.formula;
    }
    return false;
  };
  s.doc.columns.forEach((col, ci) => {
    if (del.has(col.id)) return;
    col.cells.forEach((c, ri) => {
      if (c?.f && breaks(c.f, s.doc.sheetName)) out.push({ where: `${colLetter(ci + 1)}${ri + 1}`, text: '=' + c.f, colIdx: ci, row: ri + 1 });
    });
  });
  for (const ef of s.doc.externalFormulas) {
    if (breaks(ef.f, ef.sheet)) out.push({ where: `${ef.sheet}!${colLetter(ef.col)}${ef.row}`, text: '=' + ef.f, sheet: ef.sheet });
  }
  return out;
}

function analyzeDelete(s: EditorSnapshot, ids: string[]): DeleteImpact {
  const b = brokenFormulas(s, ids);
  const lookupWarnings: string[] = [];
  for (const col of s.doc.columns) for (const c of col.cells) if (c?.f && /VLOOKUP|HLOOKUP/i.test(c.f) && lookupWarnings.length < 1) lookupWarnings.push('VLOOKUP の列番号は自動では調整されません（Excel と同じ）');
  return { breaks: b.slice(0, 20).map(({ where, text }) => ({ where, text })), breakCount: b.length, lookupWarnings };
}

function deleteColumns(s: EditorSnapshot, ids: string[], opts: { convertDependents: boolean }): OpResult {
  let doc = s.doc;
  const log: LoggedOp[] = s.log.slice();
  if (opts.convertDependents) {
    const b = brokenFormulas(s, ids).filter((x) => x.colIdx !== undefined);
    if (b.length) {
      const writes: CellWrite[] = [];
      const columns = doc.columns.map((col, ci) => {
        const hits = b.filter((x) => x.colIdx === ci);
        if (!hits.length) return col;
        const cells = col.cells.slice();
        for (const h of hits) {
          const c = cells[h.row! - 1]!;
          const t = c.t === 'z' ? 'z' : c.t;
          cells[h.row! - 1] = { t, v: c.v, s: c.s };
          writes.push({ row: h.row!, col: ci + 1, cell: { t, v: c.v } });
        }
        return { ...col, cells };
      });
      doc = { ...doc, columns, version: doc.version + 1 };
      log.push({ kind: 'cells', sheet: doc.sheetName, writes, reason: 'toValues' });
    }
  }
  const snap: EditorSnapshot = { doc, log, label: '列を削除' };
  for (const op of deleteOps(snap, ids)) {
    doc = applyStruct(doc, op);
    log.push({ kind: 'struct', sheet: doc.sheetName, op });
  }
  return ok({ doc, log, label: '列を削除' });
}

function parseInput(text: string, textColumn: boolean): { t: Cell['t']; v: Cell['v']; f?: string } | null {
  if (text === '') return null;
  if (text.startsWith('=') && text.length > 1) return { t: 'z', v: null, f: text.slice(1) };
  if (textColumn) return { t: 's', v: text };
  const t = text.trim();
  if (/^[+-]?\d+(\.\d+)?$/.test(t) && !/^0\d/.test(t)) return { t: 'n', v: Number(t) };
  if (/^(TRUE|FALSE)$/i.test(t)) return { t: 'b', v: t.toUpperCase() === 'TRUE' };
  const m = /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/.exec(t);
  if (m) {
    const ms = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return { t: 'd', v: Math.round(ms / 86400000) + 25569 };
  }
  return { t: 's', v: text };
}

function setCellInput(s: EditorSnapshot, colId: string, row: number, text: string): OpResult {
  const ci = s.doc.columns.findIndex((c) => c.id === colId);
  if (ci < 0) return { ok: false, reason: '列が見つかりません' };
  const col = s.doc.columns[ci];
  const old = col.cells[row - 1];
  const numFmt = s.doc.styles[old?.s ?? col.props.s ?? 0]?.numFmt;
  const parsed = parseInput(text, numFmt === '@' || row === s.doc.headerRow);
  const cells = col.cells.slice();
  while (cells.length < row) cells.push(undefined);
  cells[row - 1] = parsed ? { ...parsed, s: old?.s } : old?.s !== undefined ? { t: 'z', v: null, s: old.s } : undefined;
  const columns = s.doc.columns.slice();
  columns[ci] = { ...col, cells };
  const doc: MergedDoc = { ...s.doc, columns, version: s.doc.version + 1, rowCount: Math.max(s.doc.rowCount, row) };
  const write: CellWrite = { row, col: ci + 1, cell: parsed ? { t: parsed.t, v: parsed.v, ...(parsed.f ? { f: parsed.f } : null) } : null };
  return ok({ doc, log: [...s.log, { kind: 'cells', sheet: doc.sheetName, writes: [write], reason: 'edit' }], label: 'セルを編集' });
}

function setColumnWidth(s: EditorSnapshot, colId: string, width: number): OpResult {
  const ci = s.doc.columns.findIndex((c) => c.id === colId);
  if (ci < 0) return { ok: false, reason: '列が見つかりません' };
  const columns = s.doc.columns.slice();
  columns[ci] = { ...columns[ci], props: { ...columns[ci].props, width } };
  const doc = { ...s.doc, columns, version: s.doc.version + 1 };
  return ok({ doc, log: [...s.log, { kind: 'colWidth', sheet: doc.sheetName, col: ci + 1, width }], label: '列幅を変更' });
}

// =============================================================================================
// export / server

function logToRunnerJob(log: LoggedOp[], opts: { format: 'xlsx' | 'xlsm'; sheet: string }): RunnerJob {
  return {
    version: 1,
    format: opts.format,
    ops: log.map((op) => ({ ...op, sheet: opts.sheet }) as Record<string, unknown>),
    options: { recalc: true, clipboard: 'restoreAtEnd' },
  };
}

async function fakeWork(steps: string[], onProgress?: ProgressFn, signal?: AbortSignal, stepMs = 350): Promise<void> {
  for (let i = 0; i < steps.length; i++) {
    if (signal?.aborted) throw abortError();
    onProgress?.({ ratio: i / steps.length, message: steps[i] });
    await sleep(stepMs);
  }
  if (signal?.aborted) throw abortError();
  onProgress?.({ ratio: 1, message: '完了' });
}

const MOCK_BYTES = () => new TextEncoder().encode('MOCK EXPORT — this file is produced by the ?mock=1 demo and is not a real workbook');

export type MockVariant = 'excel' | 'noexcel' | 'standalone' | 'excelfail' | 'excelwarn';

export function mockVariant(search: string): MockVariant {
  const v = new URLSearchParams(search).get('mock') ?? '';
  if (v === 'noexcel' || v === 'standalone' || v === 'excelfail' || v === 'excelwarn') return v;
  return 'excel';
}

export function createMockServices(search: string): Services {
  const variant = mockVariant(search);
  const params = new URLSearchParams(search);
  const delay = Number(params.get('delay') ?? '1');
  delayScale = Number.isFinite(delay) && delay >= 0 ? delay : 1;
  const rowsParam = Number(params.get('rows') ?? '30');
  nStudents = Number.isInteger(rowsParam) && rowsParam >= 30 && rowsParam <= 200_000 ? rowsParam : 30;
  const colsParam = Number(params.get('cols') ?? '11');
  extraCols = Number.isInteger(colsParam) && colsParam > 11 && colsParam <= 1000 ? colsParam - 11 : 0;
  const tm = params.get('titlemerge');
  titleMerge = tm !== null && tm !== '0' && tm !== 'false';
  const status: ServerStatus | null =
    variant === 'standalone'
      ? null
      : { app: 'excel-merger', version: '1.0.0-mock', excel: { installed: variant !== 'noexcel', version: variant === 'noexcel' ? undefined : '16.0' }, outputDir: 'C:\\Users\\user\\Documents\\Excel統合ツール\\出力' };
  return {
    mock: true,
    async readBook(bytes, fileName, opts) {
      const stored = mergedStore.get(bytes);
      if (stored) {
        await sleep(150);
        return { ...stored, fileName };
      }
      opts?.onProgress?.({ ratio: 0.2, message: 'ブックを開いています…' });
      await sleep(220);
      opts?.onProgress?.({ ratio: 0.6, message: 'シートを読み込み中…' });
      await sleep(220);
      return bookFor(opts?.id ?? stripExt(fileName), fileName, bytes.length);
    },
    cellText: mockCellText,
    buildPremerged,
    pickDataSheet: (book) => bookMeta.get(book)?.dataSheet ?? book.sheets[0]?.name ?? '',
    detectRegion: detectRegionGeneric,
    analyzeSources,
    defaultPlan,
    joinSources,
    addManualCandidate,
    createEditor,
    moveColumns,
    insertColumns,
    analyzeDelete,
    deleteColumns,
    setCellInput,
    setColumnWidth,
    formulaForDisplay: (doc, colIndex, row) => {
      const f = doc.columns[colIndex]?.cells[row - 1]?.f;
      return f === undefined ? null : '=' + f;
    },
    async exportWithJs(_premerged, log, onProgress) {
      await fakeWork(['パッケージを展開しています…', `列の操作を適用しています（${log.length} 件）…`, '数式・書式を調整しています…', '圧縮しています…'], onProgress);
      return MOCK_BYTES();
    },
    logToRunnerJob,
    async getServerStatus() {
      await sleep(80);
      return status;
    },
    async runExcelJob(_premerged, job, onProgress, signal) {
      await fakeWork(['Excel を起動しています…', 'ブックを開いています…', `操作を再生しています（${job.ops.length} 件）…`], onProgress, signal, 450);
      if (variant === 'excelfail') throw new Error('Excel が応答しませんでした（タイムアウト）');
      await fakeWork(['保存しています…'], onProgress, signal, 300);
      const warnings =
        variant === 'excelwarn'
          ? ['マクロ（VBA）は保存形式の都合で保持されませんでした。', '#REF! エラーのセルが 2 個あります（集計!B7 など）。出力したファイルで確認してください。']
          : [];
      return { bytes: MOCK_BYTES(), warnings };
    },
    async saveToOutput(_bytes, fileName) {
      if (!status) throw new Error('サーバーに接続していません');
      await sleep(150);
      return { path: `${status.outputDir}\\${fileName}`, name: fileName };
    },
    async openFile() {
      await sleep(50);
    },
    async openOutputFolder() {
      await sleep(50);
    },
    startHeartbeat() {
      /* no server */
    },
  };
}

/** Demo input files for ?demo=files|merged (content is ignored by the mock reader). */
export function demoFiles(search = ''): File[] {
  const mk = (name: string, size: number) => new File([new Uint8Array(size)], name, { lastModified: 1_700_000_000_000 });
  const files = [mk('学生回答.xlsx', 32_600), mk('職場回答.xlsx', 15_600), mk('人事回答.xlsx', 17_000)];
  // &nokey=1: a fourth file without any recognisable key column (「結合対象外」 / manual key picker demo)
  const nk = new URLSearchParams(search).get('nokey');
  if (nk !== null && nk !== '0' && nk !== 'false') files.push(mk('参加状況（キーなし）.xlsx', 9_800));
  return files;
}

export function demoStage(search: string): 'files' | 'merged' | null {
  const d = new URLSearchParams(search).get('demo');
  return d === 'files' || d === 'merged' ? d : null;
}
