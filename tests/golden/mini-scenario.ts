/**
 * A small premerge scenario built from hand-written packages (used by tests/xlsx/patch-fuzz.test.ts and the tool
 * tools/make-fuzz-outputs.ts that writes representative outputs for the "opens in Excel without repair" check).
 *   base  sheet <baseName>: ID / 氏名 (S1 山田, S2 佐藤, S3 鈴木 in rows 2..4), 合計 =COUNTA(B2:B4) in row 6;
 *         sheet Report: a formula that refers to the 合計 cell (sheet name quoted when needed)
 *   source sheet <srcName>: ID 点数 備考 計算 (計算 = 点数 * 2) in rows 2..5 (S3, S1, S2, S9)
 *   join  S1..S3 match, S9 is appended as a new row (join 'full'); 点数 備考 計算 are imported.
 */
import type { PremergeInput } from '../../src/core/contracts';
import { buildXlsx } from './build';
import { text, type Parts } from './package-tools';
import { source } from './premerge-scenario';

export const esc = (s: string) => s.replace(/&/g, '&amp;');
export const BASE_SST = ['ID', '氏名', 'S1', '山田', 'S2', '佐藤', 'S3', '鈴木', '合計'];
export const SRC_SST = ['ID', '点数', '備考', '計算', 'S3', 'c', 'S1', 'a', 'S2', 'b', 'S9', 'z'];
export const s = (ref: string, i: number, st = 0) => `<c r="${ref}"${st ? ` s="${st}"` : ''} t="s"><v>${i}</v></c>`;
export const n = (ref: string, v: number, st = 0) => `<c r="${ref}"${st ? ` s="${st}"` : ''}><v>${v}</v></c>`;

export interface Mini {
  baseName: string;
  srcName: string;
  base: Uint8Array;
  src: Uint8Array;
}

/** base: ID / 氏名, S1..S3 in rows 2..4, 合計 (COUNTA) in row 6, a Report sheet that refers to the 合計; source: ID 点数 備考 計算 */
export async function mini(baseName = '名簿', srcName = 'Src'): Promise<Mini> {
  const q = `'${baseName.replace(/'/g, "''")}'`;
  const base = await buildXlsx([
    {
      name: esc(baseName),
      data: `<row r="1">${s('A1', 0, 1)}${s('B1', 1, 1)}</row><row r="2">${s('A2', 2)}${s('B2', 3)}</row><row r="3">${s('A3', 4)}${s('B3', 5)}</row><row r="4">${s('A4', 6)}${s('B4', 7)}</row>`
        + `<row r="6">${s('A6', 8)}<c r="B6"><f>COUNTA(B2:B4)</f><v>3</v></c></row>`,
    },
    { name: 'Report', data: `<row r="1"><c r="A1"><f>${esc(q)}!B6*2</f><v>6</v></c></row>` },
  ], { sst: BASE_SST });
  const src = await buildXlsx([
    {
      name: esc(srcName),
      data: `<row r="1">${s('A1', 0, 1)}${s('B1', 1, 1)}${s('C1', 2, 1)}${s('D1', 3, 1)}</row>`
        + [[2, 4, 70, 5], [3, 6, 90, 7], [4, 8, 80, 9], [5, 10, 50, 11]].map(([r, id, pt, memo]) => `<row r="${r}">${s('A' + r, id)}${n('B' + r, pt, 2)}${s('C' + r, memo)}<c r="D${r}"><f>B${r}*2</f><v>${pt * 2}</v></c></row>`).join(''),
    },
  ], { sst: SRC_SST });
  return { baseName, srcName, base, src };
}

export function miniInput(m: Mini, base = m.base, src = m.src): PremergeInput {
  const match = (row: number) => ({ src: { row, via: 'ID', flags: [] as never[] } });
  return {
    base: source('base', '', { sheet: m.baseName, headerRow: 1, firstDataRow: 2, lastDataRow: 4, firstCol: 1, lastCol: 2 }, 'base', 'full', base),
    sources: [source('src', '', { sheet: m.srcName, headerRow: 1, firstDataRow: 2, lastDataRow: 5, firstCol: 1, lastCol: 4 }, 'src', 'full', src)],
    rows: [
      { baseRow: 2, matches: match(3) }, { baseRow: 3, matches: match(4) }, { baseRow: 4, matches: match(2) },
      { baseRow: null, appendedFrom: { sourceId: 'src', row: 5 }, matches: match(5) },
    ],
    imports: [2, 3, 4].map((col) => ({ sourceId: 'src', col })),
    keyFill: { src: [{ baseCol: 1, sourceCol: 1 }] },
    options: { formulaPolicy: 'keep-if-possible', extendTable: true, headerConflict: 'suffix' },
  };
}

export const inlineStrings = (p: Parts, sst: string[], sheets: string[]) => {
  for (const sh of sheets) {
    p.set(sh, (p.get(sh) as string).replace(/<c ([^>]*?)t="s"([^>]*)><v>(\d+)<\/v><\/c>/g, (_m, a: string, b: string, i: string) => `<c ${a}t="inlineStr"${b}><is><t>${esc(sst[Number(i)])}</t></is></c>`));
  }
  p.delete('xl/sharedStrings.xml');
  p.set('xl/_rels/workbook.xml.rels', text(p, 'xl/_rels/workbook.xml.rels').replace(/<Relationship [^>]*\/sharedStrings"[^>]*\/>/, ''));
  p.set('[Content_Types].xml', text(p, '[Content_Types].xml').replace(/<Override PartName="\/xl\/sharedStrings.xml"[^>]*\/>/, ''));
};
export const noStyles = (p: Parts, sheets: string[]) => {
  for (const sh of sheets) p.set(sh, (p.get(sh) as string).replace(/(<c r="[A-Z]+\d+") s="\d+"/g, '$1'));
  p.delete('xl/styles.xml');
  p.set('xl/_rels/workbook.xml.rels', text(p, 'xl/_rels/workbook.xml.rels').replace(/<Relationship [^>]*\/styles"[^>]*\/>/, ''));
  p.set('[Content_Types].xml', text(p, '[Content_Types].xml').replace(/<Override PartName="\/xl\/styles.xml"[^>]*\/>/, ''));
};
export const SH1 = 'xl/worksheets/sheet1.xml';
export const SH2 = 'xl/worksheets/sheet2.xml';
