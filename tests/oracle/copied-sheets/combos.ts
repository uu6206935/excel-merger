/**
 * The combinations of the copied-sheets oracle (task V9): which base + other files are merged (card order = the order of `others`), how
 * (join type, the base's other sheets kept / dropped) and with which random edit logs. Every combination has includeSourceSheets = true.
 *
 *  - smp-*  : the app's samples 01..05 as base and in several card orders (xlsx / csv / xls, merged title, hidden helper column, charts, pivot,
 *             a table sheet, CF / DV), joined for real (imported columns, rows by key)
 *  - cor-*  : one foreign file next to the sample 01 (Google Sheets, LibreOffice, Mac 1904, Strict, openpyxl / XlsxWriter / pandas, hidden sheets,
 *             many sheets, tables, external links, macro-enabled ...) — the foreign file only gets its sheets copied ('copy' mode)
 *  - ob-*   : the Excel-authored object fixtures of A2 (charts, shapes, notes, tables, pivots, slicers / controls, Strict), next to a joined
 *             sample, and the golden fixtures that own objects
 *  - nam-*  : three files with the same file name, a long file name, forbidden characters in the file name, tricky sheet names
 *  - rich-* : the generated rich base with F1's formula sources (consistency/sets.ts): imported formulas that read the source's other sheets /
 *             side cells / footer / running totals stay formulas by reading the copies
 */
import fs from 'node:fs';
import path from 'node:path';
import type { KeySource } from '../../../src/core/contracts';
import { generatedWorkbooks, buildSets } from '../consistency/sets';
import { REPO, SAMPLES } from '../consistency/lib';
import type { FileSpec, PipelineParams } from './pipeline';

export interface Participant {
  ks: KeySource;
  mode: 'join' | 'copy';
}

export interface Combo {
  id: string;
  group: 'smp' | 'cor' | 'ob' | 'nam' | 'rich';
  base?: FileSpec;
  others?: FileSpec[];
  /** a builder instead of files (the rich sets) */
  build?: (ctx: { genDir: string }) => Promise<{ base: KeySource; others: Participant[] }>;
  params?: Partial<PipelineParams>;
  /** random edit logs to make (one scenario each); default [1] */
  seeds?: number[];
  note?: string;
}

const S = (n: string): FileSpec => ({ path: path.join(SAMPLES, n) });
const C = (n: string): FileSpec => ({ path: path.join(REPO, 'tests', 'fixtures', 'foreign', 'corpus', n), mode: 'copy' });
const X = (n: string, mode: 'join' | 'copy' = 'copy'): FileSpec => ({ path: path.join(REPO, 'tests', 'fixtures', 'addsheets', n), mode });
const O = (n: string): FileSpec => ({ path: path.join(REPO, 'tests', 'fixtures', 'addsheets-objects', n), mode: 'copy' });
const G = (n: string): FileSpec => ({ path: path.join(REPO, 'tests', 'fixtures', 'golden', n), mode: 'copy' });
const withName = (f: FileSpec, fileName: string, mode?: 'join' | 'copy'): FileSpec => ({ ...f, fileName, ...(mode ? { mode } : {}) });

const KEEP: Partial<PipelineParams> = { dropOtherSheets: false };
const DROP: Partial<PipelineParams> = { dropOtherSheets: true };
const FULL = { joinType: 'full' } as const;

export function buildCombos(): Combo[] {
  const out: Combo[] = [];
  // ---- the samples: every file as base, several card orders, both join types, base's other sheets kept / dropped
  const smp = (id: string, base: string, others: string[], params: Partial<PipelineParams>, seeds: number[] = [1]): void => {
    out.push({ id: `smp-${id}`, group: 'smp', base: S(base), others: others.map((o) => S(o)), params, seeds });
  };
  smp('01-all-keep', '01_学生名簿.xlsx', ['02_面接評価.xlsx', '03_連絡先.csv', '04_適性検査.xls', '05_研修出欠.xlsx'], KEEP, [1, 2]);
  smp('01-all-drop', '01_学生名簿.xlsx', ['02_面接評価.xlsx', '03_連絡先.csv', '04_適性検査.xls', '05_研修出欠.xlsx'], DROP);
  smp('01-rev-keep', '01_学生名簿.xlsx', ['05_研修出欠.xlsx', '04_適性検査.xls', '03_連絡先.csv', '02_面接評価.xlsx'], KEEP);
  smp('01-shuf-full-drop', '01_学生名簿.xlsx', ['03_連絡先.csv', '05_研修出欠.xlsx', '02_面接評価.xlsx', '04_適性検査.xls'], { ...DROP, ...FULL });
  smp('01-full-keep', '01_学生名簿.xlsx', ['02_面接評価.xlsx', '05_研修出欠.xlsx'], { ...KEEP, ...FULL }, [1, 2]);
  smp('02-base-keep', '02_面接評価.xlsx', ['01_学生名簿.xlsx', '05_研修出欠.xlsx'], KEEP, [1, 2]);
  smp('02-base-full-drop', '02_面接評価.xlsx', ['01_学生名簿.xlsx', '03_連絡先.csv', '05_研修出欠.xlsx'], { ...DROP, ...FULL });
  smp('02-rev-keep', '02_面接評価.xlsx', ['05_研修出欠.xlsx', '04_適性検査.xls', '03_連絡先.csv', '01_学生名簿.xlsx'], KEEP);
  smp('03csv-base-keep', '03_連絡先.csv', ['01_学生名簿.xlsx', '02_面接評価.xlsx'], KEEP);
  smp('04xls-base-drop', '04_適性検査.xls', ['01_学生名簿.xlsx', '02_面接評価.xlsx', '05_研修出欠.xlsx'], DROP);
  smp('05-base-keep', '05_研修出欠.xlsx', ['01_学生名簿.xlsx', '02_面接評価.xlsx', '03_連絡先.csv', '04_適性検査.xls'], KEEP, [1, 2]);
  smp('05-base-full-drop', '05_研修出欠.xlsx', ['04_適性検査.xls', '03_連絡先.csv', '02_面接評価.xlsx', '01_学生名簿.xlsx'], { ...DROP, ...FULL });

  // ---- foreign files next to sample 01 (sheets copied only)
  const cor = (name: string, params: Partial<PipelineParams>, others: FileSpec[] = []): void => {
    out.push({ id: `cor-${name}`, group: 'cor', base: S('01_学生名簿.xlsx'), others: [S('02_面接評価.xlsx'), ...others, C(`${name}.xlsx`)], params });
  };
  const COR_KEEP = ['hc-01-google-sheets', 'hc-05-mac-excel-2011-1904', 'hc-08-strict-ooxml', 'hc-17-sheetnames-hidden-first', 'hc-30-libreoffice-hidden-filter', 'hc-35-names-cf-dv', 'hc-37-external-link', 'js-02-sheetjs-date1904', 'js-05-exceljs-table', 'py-04-openpyxl-table', 'py-05-openpyxl-hidden', 'py-06-openpyxl-date1904', 'py-14-openpyxl-formats-errors', 'py-25-pandas-many-sheets', 'py-18-xlsxwriter-chart-image'];
  const COR_DROP = ['hc-02-libreoffice', 'hc-06-excel2007-ja-numfmts', 'hc-10-shared-formulas', 'hc-19-merged-two-row-header', 'hc-29-formula-variety', 'hc-32-sheetnames-tricky', 'hc-36-theme2013-colours', 'js-04-exceljs-basic', 'js-10-write-excel-file', 'py-01-openpyxl-styled', 'py-12-openpyxl-protected', 'py-17-xlsxwriter-arrays', 'py-19-xlsxwriter-table', 'py-27-openpyxl-chart-on-data-sheet'];
  for (const n of COR_KEEP) cor(n, KEEP);
  for (const n of COR_DROP) cor(n, DROP);
  // Excel-authored: another theme, formats, two sheets, 1904, macro-enabled (as other file and as base)
  for (const n of ['ex-formats.xlsx', 'ex-theme-other.xlsx', 'ex-two-sheets.xlsx', 'ex-1904.xlsx', 'ex-macro.xlsm']) {
    out.push({ id: `cor-${n.replace(/\.[^.]+$/, '')}`, group: 'cor', base: S('01_学生名簿.xlsx'), others: [S('02_面接評価.xlsx'), X(n)], params: n === 'ex-1904.xlsx' ? KEEP : DROP });
  }
  out.push({ id: 'cor-ex-all-four-keep', group: 'cor', base: S('01_学生名簿.xlsx'), others: [X('ex-formats.xlsx'), X('ex-theme-other.xlsx'), X('ex-two-sheets.xlsx'), X('ex-1904.xlsx')], params: KEEP });
  out.push({ id: 'cor-ex-macro-as-base', group: 'cor', base: X('ex-macro.xlsm'), others: [S('02_面接評価.xlsx'), X('ex-formats.xlsx')], params: KEEP, note: 'a macro-enabled base: .xlsm output, vbaProject kept' });
  out.push({ id: 'cor-ex-1904-as-base', group: 'cor', base: X('ex-1904.xlsx'), others: [S('01_学生名簿.xlsx'), X('ex-formats.xlsx')], params: KEEP, note: 'a 1904 base: the 1900 sources are shifted into it' });
  out.push({ id: 'cor-hc-17-as-base', group: 'cor', base: { ...C('hc-17-sheetnames-hidden-first.xlsx'), mode: 'join' }, others: [{ ...S('02_面接評価.xlsx'), mode: 'copy' }], params: KEEP, note: 'the data sheet of the base is not its first sheet' });

  // ---- the object fixtures of A2
  for (const n of ['ob-charts', 'ob-shapes', 'ob-comments', 'ob-table', 'ob-pivot', 'ob-slicer-control', 'ob-strict']) {
    out.push({ id: `ob-${n.replace(/^ob-/, '')}`, group: 'ob', base: S('01_学生名簿.xlsx'), others: [S('02_面接評価.xlsx'), O(`${n}.xlsx`)], params: n === 'ob-table' || n === 'ob-pivot' ? KEEP : DROP });
  }
  out.push({ id: 'ob-all-keep', group: 'ob', base: S('02_面接評価.xlsx'), others: [S('01_学生名簿.xlsx'), O('ob-charts.xlsx'), O('ob-shapes.xlsx'), O('ob-comments.xlsx'), O('ob-table.xlsx'), O('ob-pivot.xlsx')], params: KEEP, seeds: [1, 2] });
  out.push({ id: 'ob-twice', group: 'ob', base: S('02_面接評価.xlsx'), others: [O('ob-table.xlsx'), withName(O('ob-table.xlsx'), 'ob-table-2.xlsx'), O('ob-pivot.xlsx'), withName(O('ob-pivot.xlsx'), 'ob-pivot-2.xlsx')], params: KEEP, note: 'table / pivot names that collide' });
  out.push({ id: 'ob-golden-table', group: 'ob', base: S('01_学生名簿.xlsx'), others: [S('02_面接評価.xlsx'), G('g-table.xlsx'), G('torture.xlsx')], params: KEEP });
  out.push({ id: 'ob-table-as-base-keep', group: 'ob', base: { ...O('ob-table.xlsx'), mode: 'join' }, others: [withName(O('ob-table.xlsx'), 'ob-table-2.xlsx'), S('01_学生名簿.xlsx')], params: KEEP, note: 'a base that owns a table, an other file with the same table name' });

  // ---- names
  out.push({ id: 'nam-same-name-x3', group: 'nam', base: S('02_面接評価.xlsx'), others: [S('01_学生名簿.xlsx'), withName(S('01_学生名簿.xlsx'), '01_学生名簿.xlsx', 'copy'), withName(S('01_学生名簿.xlsx'), 'C:\\Users\\x\\01_学生名簿.xlsx', 'copy')], params: KEEP, note: 'the same file dropped three times: (1) and (2) go to the later ones' });
  out.push({ id: 'nam-long-file-name', group: 'nam', base: S('01_学生名簿.xlsx'), others: [withName(S('02_面接評価.xlsx'), '営業部門の月次実績報告書（2026年度第3四半期確定版）.xlsx')], params: DROP, note: 'file part is shortened to fit 31 characters' });
  out.push({ id: 'nam-forbidden-chars', group: 'nam', base: S('01_学生名簿.xlsx'), others: [withName(S('02_面接評価.xlsx'), '売上[確定]:v2*?.xlsx'), withName(S('05_研修出欠.xlsx'), "田中's 出欠.xlsx")], params: KEEP, note: '[ ] : * ? become "_"; an apostrophe in the name needs quoting in formulas' });
  out.push({ id: 'nam-tricky-sheet-names', group: 'nam', base: S('01_学生名簿.xlsx'), others: [C('hc-32-sheetnames-tricky.xlsx'), C('hc-25-sheet-order-ids.xlsx')], params: KEEP });

  // ---- the generated rich base + F1's formula sources: formulas that read the copies
  const richVariants: Array<{ set: string; params: Partial<PipelineParams>; seeds?: number[] }> = [
    { set: 'rich-f07_othersheet-c', params: KEEP, seeds: [1, 2] },
    { set: 'rich-f04_lookups-c', params: KEEP },
    { set: 'rich-f12_footer-c', params: { ...DROP, ...FULL } },
    { set: 'rich-f06_running-c', params: KEEP },
    { set: 'rich-f08_names-n', params: KEEP },
    { set: 'rich-f09_tables-c', params: DROP },
    { set: 'rich-f05_ranges-n', params: { ...KEEP, ...FULL } },
    { set: 'rich-f03_keycol-c', params: KEEP },
  ];
  for (const rv of richVariants) {
    out.push({
      id: rv.set, group: 'rich', params: rv.params, ...(rv.seeds ? { seeds: rv.seeds } : {}),
      build: async (ctx) => {
        const set = buildSets().find((s) => s.id === rv.set);
        if (!set) throw new Error(`no such set ${rv.set}`);
        const built = await set.build(ctx);
        return { base: built.base, others: built.others.map((ks) => ({ ks, mode: 'join' as const })) };
      },
    });
  }
  return out;
}

/** writes the generated workbooks of the rich sets (raw: no Excel needed) into <genDir>/raw */
export async function writeGenerated(genDir: string): Promise<void> {
  const raw = path.join(genDir, 'raw');
  fs.mkdirSync(raw, { recursive: true });
  for (const [name, bytes] of await generatedWorkbooks()) fs.writeFileSync(path.join(raw, `${name}.xlsx`), bytes);
}
