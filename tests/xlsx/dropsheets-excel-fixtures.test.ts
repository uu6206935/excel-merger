/**
 * Single-sheet output on workbooks that Excel itself wrote with every feature the option has to clean up
 * (tests/golden/com/make-fixtures-dropsheets.ps1 -> tests/fixtures/dropsheets/ds-features.xlsx, ds-table.xlsx):
 *  - ds-features: merged sheet 名簿 + 設定 マスタ 売上 (Table) 集計 (chart) ピボット (pivot table); formulas that read the other sheets
 *    (direct, whole column, 3D, names, a table, spills, a CSE array, a spill reference), validations (list on another sheet, via a
 *    name), conditional formats that read another sheet (x14), links, 3 charts, sparklines, text boxes with a cell link, a note + a
 *    threaded comment, two pivot tables on the merged sheet and two slicers bound to things of other sheets;
 *  - ds-table: merged sheet 名簿 is an Excel Table (calculated column through a name, totals row) with a table slicer, a pivot table
 *    and its slicer; another table slicer is bound to the table of 売上.
 * Excel opened the outputs of these tests without repair, with the same values after a full recalculation (tests/golden/RESULTS.md).
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { createEditor } from '../../src/core/doc';
import { buildPremerged, readBook, replayLog } from '../../src/core/xlsx';
import type { PremergeInput } from '../../src/core/contracts';
import type { LoggedOp, MergedDoc } from '../../src/core/types';
import { compare } from '../doc/compare';
import { rng, step } from '../doc/fuzzcore';
import { source } from '../golden/premerge-scenario';
import { baseOptions, baseSource, files, lintSingleSheet, lintText, premergeOne, readFixture } from './dropsheets-helpers';

const here = path.dirname(fileURLToPath(import.meta.url));
const FIX = path.resolve(here, '../fixtures/dropsheets');
const have = (f: string) => fs.existsSync(path.join(FIX, f));
const hasAll = have('ds-features.xlsx') && have('ds-table.xlsx');

/** a `hyperlinks doc … file …` difference that is only about the text of a link's location (see the replay test below) */
const notOnlyLinkLocation = (d: string): boolean => {
  const m = /^hyperlinks doc (.*) file (.*)$/.exec(d);
  const refs = (x: string) => x.split(' ').map((e) => e.split('|').slice(0, 2).join('|')).sort().join(' ');
  return !(m && refs(m[1]) === refs(m[2]));
};

const cellAt = (sheet: { rows: Array<Array<{ f?: string; v: unknown; t: string } | undefined> | undefined> }, ref: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!;
  return sheet.rows[Number(m[2]) - 1]?.[m[1].split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1];
};

describe.skipIf(!hasAll)('ds-features.xlsx (written by Excel), single-sheet output', () => {
  const REMOVED = ['設定', 'マスタ', '売上', '集計', 'ピボット'];
  const run = async () => {
    const bytes = readFixture(FIX, 'ds-features.xlsx');
    const base = await baseSource({ bytes, fileName: 'ds-features.xlsx', sheet: '名簿' });
    return { bytes, base, res: await premergeOne(base) };
  };

  it('one sheet, a clean package, and the report tells everything that was left out', async () => {
    const { res } = await run();
    expect(lintText(await lintSingleSheet(res.bytes, REMOVED))).toBe('');
    expect(res.report.droppedSheets).toEqual(REMOVED);
    expect(res.report.formulasDetached).toBe(45);
    const by = (code: string) => res.report.items.find((i) => i.code === code)?.message;
    expect(by('detachedFormulas')).toBe('他のシートを参照していた数式 45 個を値にしました');
    expect(by('dvToList')).toBe('入力規則 2 個を固定のリストにしました');
    expect(by('dvDropped')).toContain('1 個');
    expect(by('cfDropped')).toBe('条件付き書式 4 個を外しました');
    expect(by('linksDropped')).toBe('他のシートへのリンク 1 個を外しました');
    expect(by('chartsDropped')).toBe('他のシートのデータを使うグラフ 2 個は含めません');
    expect(by('sparklinesDropped')).toBe('他のシートのデータを使うスパークライン 1 個を外しました');
    expect(by('pivotsToValues')).toBe('他のシートのデータを使うピボットテーブル 1 個を値にしました');
    expect(by('slicersDropped')).toContain('2 個');
    expect(res.report.items.every((i) => i.severity === 'info' && /[ぁ-んァ-ン一-龥]/.test(i.message))).toBe(true);
  });

  it('formulas: the 45 that read another sheet are values (cached results), the own-sheet ones stay live', async () => {
    const { res, bytes } = await run();
    const out = (await readBook(res.bytes, 'o.xlsx')).sheets[0];
    const inp = (await readBook(bytes, 'i.xlsx')).sheets[0];
    const detached: string[] = [];
    for (let r = 1; r <= inp.maxRow; r++) {
      for (let c = 1; c <= inp.maxCol; c++) {
        const a = inp.rows[r - 1]?.[c - 1];
        const b = out.rows[r - 1]?.[c - 1];
        if (a?.f !== undefined && b?.f === undefined) detached.push(`${String.fromCharCode(64 + c)}${r}`);
        if (a?.f !== undefined && b?.f !== undefined) expect(b.f, `${r},${c}`).toBe(a.f);
        if (a && b) expect(b.v, `${r},${c}`).toEqual(a.v);
      }
    }
    expect(detached).toHaveLength(45);
    // E税込 (name), F VLOOKUP, G IF(設定), H table, I names chain, K/L spill + spill reference, N CSE array, O 3D, D11 whole column
    for (const ref of ['E2', 'E9', 'F2', 'G3', 'H4', 'I5', 'K2', 'L2', 'N2', 'O2', 'D11']) expect(cellAt(out, ref)?.f, ref).toBeUndefined();
    // live: J = C*2, M = SORT(UNIQUE(D2:D9)) (own spill), O3 = SUM(名簿!C2:C9), C11 = SUM(C2:C9)
    expect(cellAt(out, 'J2')?.f).toBe('C2*2');
    expect(cellAt(out, 'O3')?.f).toBe('SUM(名簿!C2:C9)');
    expect(cellAt(out, 'C11')?.f).toBe('SUM(C2:C9)');
    expect(cellAt(out, 'M2')).toMatchObject({ dyn: true });
    expect(out.features.arrays).toHaveLength(1);
  });

  it('names, validations, conditional formats, links: only what does not read another sheet is left', async () => {
    const { res } = await run();
    const book = await readBook(res.bytes, 'o.xlsx');
    const s = book.sheets[0];
    expect(book.names.map((n) => `${n.scope ?? ''}!${n.name}`).sort()).toEqual(['名簿!_xlnm._FilterDatabase', '名簿!_xlnm.Print_Area', '名簿!_xlnm.Print_Titles', '名簿!局所'].sort());
    expect(s.features.rules.filter((r) => r.kind === 'dv').map((r) => `${r.sqref}|${r.type}|${r.ext ? 'x14' : ''}|${r.formulas.join(',')}`).sort()).toEqual([
      'B2:B9|list||"東大,京大,阪大,東北大,九大"', 'C2:C9|whole||0,100', 'D2:D9|list||"東大,京大,阪大,東北大,九大"',
    ]);
    expect(s.features.rules.filter((r) => r.kind === 'cf').map((r) => `${r.sqref}|${r.type}|${r.ext ? 'x14' : ''}`).sort()).toEqual(['E2:E9|cellIs|', 'F2:F9|dataBar|', 'F2:F9|dataBar|x14']);
    expect(s.features.hyperlinks.map((h) => h.ref).sort()).toEqual(['A13', 'A14']);
    expect(s.features.hyperlinks.find((h) => h.ref === 'A13')).toMatchObject({ location: '名簿!B2' });
  });

  it('drawings, pivot tables, slicers, comments: own-sheet chart / pivot / notes stay, everything bound to another sheet is gone', async () => {
    const { res } = await run();
    const f = await files(res.bytes);
    const keys = [...f.keys()];
    expect(keys.filter((k) => /charts\/chart\d+\.xml$/.test(k))).toHaveLength(1);
    expect(keys.filter((k) => /pivotTables\/pivotTable\d+\.xml$/.test(k))).toHaveLength(1);
    expect(keys.filter((k) => /slicer|tables\/table|worksheets\/sheet[2-9]|calcChain/i.test(k))).toEqual([]);
    for (const k of ['xl/comments1.xml', 'xl/threadedComments/threadedComment1.xml', 'xl/persons/person.xml', 'xl/metadata.xml', 'xl/printerSettings/printerSettings1.bin']) expect(keys, k).toContain(k);
    const pivot = f.get('xl/pivotTables/pivotTable1.xml')!;
    expect(pivot).toContain('name="PT自"');
    const dr = f.get('xl/drawings/drawing1.xml')!;
    expect(dr).not.toMatch(/slicer|設定|マスタ/);
    expect(dr).toContain('textlink="名簿!$A$1"');
    expect(dr).toContain('textlink=""');
    expect((dr.match(/<xdr:twoCellAnchor\b/g) ?? []).length).toBe(3); // the own chart and the two text boxes
    const wb = f.get('xl/workbook.xml')!;
    expect(wb).not.toMatch(/slicer|スライサー/i);
    expect(wb).toContain('<pivotCaches><pivotCache cacheId="11" r:id=');
    const sheet = f.get('xl/worksheets/sheet1.xml')!;
    expect(sheet).not.toMatch(/slicer|設定|マスタ|売上|集計/);
  });

  it('the doc built on the output has no other sheets; random op sequences replay like the doc says', async () => {
    const { res, base } = await run();
    const book = await readBook(res.bytes, 'merged.xlsx', { id: 'merged' });
    const sourceInfo: MergedDoc['sourceInfo'] = { b: { label: 'ds-features', color: '#2563eb' } };
    const snapshot = createEditor(book, res.layout, sourceInfo, { headerRow: base.region.headerRow, firstDataRow: base.region.firstDataRow, lastDataRow: base.region.lastDataRow });
    expect(snapshot.doc.sheetNames).toEqual(['名簿']);
    expect(snapshot.doc.externalFormulas).toEqual([]);
    expect(compare(snapshot.doc, book)).toEqual([]);
    const failures: string[] = [];
    for (const seed of Array.from({ length: 16 }, (_x, i) => 4400 + i)) {
      const r = rng(seed);
      let s = snapshot;
      const trail: string[] = [];
      for (let i = 0; i < 8; i++) {
        const { res: o, what } = step(s, r, false, s.doc.lastDataRow + 2);
        trail.push(o.ok ? what : `${what} [refused]`);
        if (o.ok) s = o.snapshot;
      }
      if (!s.log.length) continue;
      let diffs: string[];
      try {
        const replayed = await replayLog(res.bytes, s.log);
        expect(lintText(await lintSingleSheet(replayed, REMOVED))).toBe('');
        // KNOWN, independent of dropOtherSheets: replayLog follows a hyperlink's `location` (名簿!B2, a link to a cell of the merged
        // sheet itself) through column moves, the doc keeps the text (it only moves the link's `ref`); tests/golden/RESULTS.md
        diffs = compare(s.doc, await readBook(replayed, 'after.xlsx')).filter(notOnlyLinkLocation);
      } catch (e) {
        diffs = [`replay threw: ${(e as Error).message}`];
      }
      if (diffs.length) failures.push(`seed ${seed}: ${diffs.slice(0, 3).join(' | ')}\n  ops: ${trail.join('; ')}`);
    }
    expect(failures).toEqual([]);
  }, 300_000);
});

describe.skipIf(!hasAll)('ds-table.xlsx (written by Excel), single-sheet output', () => {
  it('the table of the merged sheet survives with its totals row, table slicer and pivot; the calculated column became values', async () => {
    const bytes = readFixture(FIX, 'ds-table.xlsx');
    const res = await premergeOne(await baseSource({ bytes, fileName: 'ds-table.xlsx', sheet: '名簿' }));
    expect(lintText(await lintSingleSheet(res.bytes, ['設定', '売上']))).toBe('');
    expect(res.report).toMatchObject({ droppedSheets: ['設定', '売上'], formulasDetached: 12 });
    const f = await files(res.bytes);
    const keys = [...f.keys()];
    expect(keys.filter((k) => /tables\/table\d+\.xml$/.test(k))).toHaveLength(1);
    expect(keys.filter((k) => /slicerCaches\/slicerCache\d+\.xml$/.test(k))).toHaveLength(2);
    expect(keys.filter((k) => /slicers\/slicer\d+\.xml$/.test(k))).toHaveLength(2);
    expect(keys.filter((k) => /pivotTables\/pivotTable\d+\.xml$/.test(k))).toHaveLength(1);
    const table = f.get(keys.find((k) => /tables\/table\d+\.xml$/.test(k))!)!;
    expect(table).toContain('displayName="tbl名簿"');
    expect(table).not.toContain('calculatedColumnFormula');
    expect(table).toContain('totalsRowCount="1"');
    const s = (await readBook(res.bytes, 'o.xlsx')).sheets[0];
    expect(s.features.tables.map((t) => t.name)).toEqual(['tbl名簿']);
    expect(cellAt(s, 'E2')?.f).toBeUndefined();
    expect(cellAt(s, 'E2')?.t).toBe('n');
    // the slicer that was bound to the table 売上 is gone, the ones of tbl名簿 and its pivot table are still there
    const names = [...(f.get('xl/workbook.xml')!.matchAll(/<definedName name="([^"]*)"/g))].map((m) => m[1]);
    expect(names.sort()).toEqual(['スライサー_大学', 'スライサー_大学1'].sort());
    expect(f.get('xl/drawings/drawing1.xml')!.match(/<sle:slicer\b[^>]*name="([^"]*)"/g)?.length).toBe(2);
  });

  it('with imported columns and an appended row: the table grows over them, slicers and pivot stay', async () => {
    const bytes = readFixture(FIX, 'ds-table.xlsx');
    const base = await baseSource({ bytes, fileName: 'ds-table.xlsx', sheet: '名簿' });
    const src = source('src', 'pm-src.xlsx', { sheet: '回答', headerRow: 1, firstDataRow: 2, lastDataRow: 6, firstCol: 1, lastCol: 9 }, 'pm-src');
    const rows: PremergeInput['rows'] = [];
    for (let r = base.region.firstDataRow; r <= base.region.lastDataRow; r++) rows.push({ baseRow: r, matches: { src: r <= 6 ? { row: r, via: 'ID', flags: [] } : null } });
    rows.push({ baseRow: null, appendedFrom: { sourceId: 'src', row: 4 }, matches: { src: { row: 4, via: 'ID', flags: [] } } });
    const res = await buildPremerged({
      base, sources: [src], rows, imports: [3, 4, 5].map((col) => ({ sourceId: 'src', col })), keyFill: { src: [{ baseCol: 1, sourceCol: 1 }, { baseCol: 2, sourceCol: 9 }] },
      options: baseOptions({ dropOtherSheets: true }),
    });
    expect(lintText(await lintSingleSheet(res.bytes, ['設定', '売上']))).toBe('');
    expect(res.report).toMatchObject({ droppedSheets: ['設定', '売上'], appendedRows: 1, importedColumns: 3 });
    const f = await files(res.bytes);
    const table = f.get([...f.keys()].find((k) => /tables\/table\d+\.xml$/.test(k))!)!;
    expect(table).toMatch(/ref="A1:I9"/);
    expect([...f.keys()].filter((k) => /slicers\/slicer\d+\.xml$/.test(k))).toHaveLength(2);
    // an op log on top of it (insert, move, delete, width) still gives a clean one-sheet package
    const log: LoggedOp[] = [
      { kind: 'struct', sheet: '名簿', op: { type: 'insert', axis: 'col', at: 2, count: 1 } },
      { kind: 'struct', sheet: '名簿', op: { type: 'move', axis: 'col', from: 8, count: 1, to: 3 } },
      { kind: 'colWidth', sheet: '名簿', col: 1, width: 12 },
    ];
    const after = await replayLog(res.bytes, log);
    expect(lintText(await lintSingleSheet(after, ['設定', '売上']))).toBe('');
  });
});
