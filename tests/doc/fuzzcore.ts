// Shared by fuzz.test.ts (golden fixtures) and fuzz-samples.test.ts (a premerged build of samples/*): the seeded random
// op generator. Every sequence of moves / inserts / deletes / typed cells / widths the doc accepts must read back from
// replayLog exactly as the doc shows it (compare.ts).
import JSZip from 'jszip';
import type { EditorSnapshot } from '../../src/core/types';
import type { OpResult } from '../../src/core/contracts';
import { deleteColumns, insertColumns, moveColumns, setCellInput, setColumnWidth } from '../../src/core/doc';

/** mulberry32 */
export function rng(seed: number) {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

export const TEXTS = [
  '12', '１，２３４', '2026/10/20', '10/20', '50%', '¥1,000', '1,000', 'テキスト', '=SUM(A1:B2)*2', '=SEQUENCE(2)', '=A1:A3', "'007", '',
  'TRUE', '#N/A', 'R8.10.20', '令和8年10月20日', '1e3', '3/4', '9:30', '2026年10月20日', '0.5', '-7', '(100)',
  "'abc", "'2026/10/20", '03-1234-5678', "'TRUE", '00123',
  // round 4: more shapes the parser / quote flag / number formats treat differently
  '12345678901234567890', '1234567890123456', "a\nb", '=', "'", ' 12 ', '=SUM(', '１２３', '2026-10-20', '2026/2/30', '12:30:15', '1/2/2026',
  '100%', '-50%', '$1,234.50', '1.5E+3', '0', '=1/0', '=NA()', "'=A1", "'1/2", '+81 3 1234', '2026年10月', '=TODAY()', '=ROW()*2',
];

export function step(
  s: EditorSnapshot,
  r: () => number,
  noInsert: boolean,
  totalsRow?: number,
  texts: string[] = TEXTS,
): { res: OpResult; what: string } {
  const n = s.doc.columns.length;
  const pick = () => Math.floor(r() * n);
  const distinct = (k: number) => [...new Set(Array.from({ length: k }, pick))];
  const ids = (idx: number[]) => idx.map((i) => s.doc.columns[i].id);
  const x = r();
  if (x < 0.3) {
    const idx = distinct(1 + Math.floor(r() * 3));
    const gap = Math.floor(r() * (n + 1));
    const unmergeTitles = r() < 0.6;
    return { res: moveColumns(s, ids(idx), gap, { unmergeTitles }), what: `move ${idx} -> ${gap}${unmergeTitles ? ' (unmerge)' : ''}` };
  }
  if (x < 0.4 && !noInsert) {
    const at = Math.floor(r() * (n + 1));
    const count = 1 + Math.floor(r() * 2);
    return { res: insertColumns(s, at, count), what: `insert ${count} @${at}` };
  }
  if (x < 0.55) {
    const idx = distinct(1 + Math.floor(r() * 2));
    const convertDependents = r() < 0.7;
    return { res: deleteColumns(s, ids(idx), { convertDependents }), what: `delete ${idx}${convertDependents ? ' (convert)' : ''}` };
  }
  if (x < 0.9) {
    const col = pick();
    let row = r() < 0.8 ? s.doc.firstDataRow + Math.floor(r() * Math.max(1, s.doc.lastDataRow - s.doc.firstDataRow + 1)) : 1 + Math.floor(r() * s.doc.headerRow);
    if (totalsRow !== undefined && r() < 0.2) row = totalsRow;
    const text = texts[Math.floor(r() * texts.length)];
    return { res: setCellInput(s, s.doc.columns[col].id, Math.min(row, s.doc.rowCount), text), what: `type ${JSON.stringify(text)} @col${col} row${row}` };
  }
  const col = pick();
  const w = 5 + Math.floor(r() * 35) + (r() < 0.5 ? 0.625 : 0);
  return { res: setColumnWidth(s, s.doc.columns[col].id, w), what: `width col${col} ${w}` };
}

/**
 * Insert CF / DV XML (`fragment`) into worksheet part `path`, right after <phoneticPr> / before <pageMargins> (where
 * conditionalFormatting / dataValidations belong). Used to give the OTHER sheets of a fixture rules that read the merged
 * sheet (doc.externalRules), which no golden fixture has.
 */
export async function withRulesOn(bytes: Uint8Array, path: string, fragment: string, replace = false): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  let xml = await zip.file(path)!.async('string');
  if (!xml.includes('<pageMargins')) throw new Error(`${path}: no <pageMargins> to insert before`);
  // replace: the sheet's own classic CF blocks and data validations go (x14 ones stay)
  if (replace) xml = xml.replace(/<conditionalFormatting\b[\s\S]*?<\/conditionalFormatting>/g, '').replace(/<dataValidations\b[\s\S]*?<\/dataValidations>/g, '');
  zip.file(path, xml.replace('<pageMargins', `${fragment}<pageMargins`));
  return zip.generateAsync({ type: 'uint8array' });
}

/** <cfRule> elements for the CF stress fixture (priority = running number) */
export const cfRules = {
  expr: (f: string) => `<cfRule type="expression" dxfId="0" priority="P"><formula>${f.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</formula></cfRule>`,
  cellIs: (op: string, ...f: string[]) => `<cfRule type="cellIs" dxfId="0" priority="P" operator="${op}">${f.map((x) => `<formula>${x}</formula>`).join('')}</cfRule>`,
  scale: (maxVal?: string) => `<cfRule type="colorScale" priority="P"><colorScale><cfvo type="min"/>${maxVal ? `<cfvo type="formula" val="${maxVal}"/>` : '<cfvo type="max"/>'}<color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule>`,
  bar: () => '<cfRule type="dataBar" priority="P"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule>',
  top: (rank: number) => `<cfRule type="top10" dxfId="0" priority="P" rank="${rank}"/>`,
  dup: () => '<cfRule type="duplicateValues" dxfId="0" priority="P"/>',
};
/** <conditionalFormatting> blocks; the priorities are numbered in document order */
export function cfBlocks(blocks: Array<[string, string[]]>): string {
  let p = 0;
  return blocks.map(([sqref, rules]) => `<conditionalFormatting sqref="${sqref}">${rules.map((r) => r.replace('priority="P"', `priority="${++p}"`)).join('')}</conditionalFormatting>`).join('');
}

/** Rewrite the <col> elements of every worksheet part (see Fixture.cols in fuzz.test.ts). */
export async function recol(bytes: Uint8Array, mode: 'auto' | 'bestFit'): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  for (const name of Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))) {
    const xml = await zip.file(name)!.async('string');
    const out = xml.replace(/<col\s([^>]*?)\/>/g, (_m, attrs: string) => {
      const base = attrs.replace(/\s*(customWidth|bestFit)="[^"]*"/g, '').trim();
      return `<col ${base}${mode === 'bestFit' ? ' bestFit="1" customWidth="1"' : ''}/>`;
    });
    zip.file(name, out);
  }
  return zip.generateAsync({ type: 'uint8array' });
}
