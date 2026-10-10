/**
 * The object phase of the sheet copier (task A2) on workbooks authored by Excel itself (tests/fixtures/addsheets-objects, made by
 * tests/golden/com/make-fixtures-addsheets-objects.ps1) and on samples/01_学生名簿.xlsx (chart + pivot table): the copy of every sheet owns the
 * objects its source sheet owns (counts of charts / pictures / shapes / notes / tables / pivot tables), the package passes the linter of
 * A1 plus the rules of the object parts, chart formulas read the renamed copies, the cached chart values are the source's, tables and pivot
 * tables keep their ranges, and what is not copied (slicers, form controls) is named in the sheet's one warning. Real Excel opens the same
 * outputs without repair: tests/golden/RESULTS.md ("Round 9, objects").
 */
import { describe, expect, it } from 'vitest';
import { readBook } from '../../src/core/xlsx';
import { premergeCopies, readFix, renameRefs, type FileIn } from './addsheets-helpers';
import { lintObjects, lintObjectsText, objectsOf, openPkg, type SheetObjects } from './addsheets-objects-helpers';

const FX = 'tests/fixtures/addsheets-objects';
const src = (name: string): FileIn => ({ bytes: readFix(FX, name), fileName: name });
const base = (): FileIn => ({ bytes: readFix('samples', '01_学生名簿.xlsx'), fileName: '01_学生名簿.xlsx' });

interface Pair {
  from: SheetObjects;
  to: SheetObjects;
}

async function copyOf(file: FileIn, opts: { base?: FileIn; keep?: boolean } = {}) {
  const res = await premergeCopies(opts.base ?? base(), [file], { dropOtherSheets: !opts.keep });
  const added = res.layout.addedSheets!;
  const from = await objectsOf(file.bytes);
  const to = await objectsOf(res.bytes);
  const pairs: Pair[] = added.map((a) => ({ from: from.find((s) => s.name === a.original)!, to: to.find((s) => s.name === a.name)! }));
  const rename = new Map(added.map((a) => [a.original.toLowerCase(), a.name]));
  return { res, pairs, rename, added };
}

const warnings = (res: Awaited<ReturnType<typeof copyOf>>['res'], sheet: string) => res.report.items.filter((i) => i.code === 'sheetCopyLimits' && i.message.includes(`「${sheet}」`) );

describe('Excel-authored fixtures: charts', () => {
  it('ob-charts: embedded + own-sheet + chart sheet keep their charts, formulas follow the copies, cached values are the source\'s', async () => {
    const { res, pairs, rename } = await copyOf(src('ob-charts.xlsx'));
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    expect(pairs.map((p) => [p.from.kind, p.from.charts])).toEqual([['worksheet', 1], ['chartsheet', 1], ['worksheet', 4]]);
    for (const { from, to } of pairs) {
      expect(to.kind).toBe(from.kind);
      expect(to.charts).toBe(from.charts);
      expect(to.chartFormulas).toEqual(from.chartFormulas.map((f) => renameRefs(f, rename)));
      expect(to.chartValues).toEqual(from.chartValues);
    }
    expect(res.report.items.filter((i) => i.code === 'sheetCopyLimits')).toEqual([]);
    // the source's chart cache = what the copy's cells hold
    const out = await readBook(res.bytes, 'o.xlsx');
    const g = out.sheets.find((s) => s.name === 'ob-charts_データ')!;
    expect(g.rows[1]![1]!.v).toBe(120);
  });
});

describe('Excel-authored fixtures: pictures, shapes, text boxes, groups, connectors', () => {
  it('ob-shapes: every shape / picture is copied, images once per source image, links follow the renamed sheets', async () => {
    const { res, pairs } = await copyOf(src('ob-shapes.xlsx'));
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const fig = pairs.find((p) => p.from.name === '図形')!;
    expect(fig.to.pictures).toBe(fig.from.pictures);
    expect(fig.to.shapes).toBe(fig.from.shapes);
    expect(fig.to.groups).toBe(fig.from.groups);
    expect(fig.to.connectors).toBe(fig.from.connectors);
    expect(fig.from.pictures).toBe(4);
    const bg = pairs.find((p) => p.from.name === '背景')!;
    expect(bg.from.background).toBe(true);
    expect(bg.to.background).toBe(true);
    // the images: the logo is used by two pictures and copied once (+ banner + background)
    const names = Object.keys((await (await import('jszip')).default.loadAsync(res.bytes)).files).filter((n) => /^xl\/media\/.+/.test(n));
    expect(names.length).toBe(3);
    // the shape that shows a cell and the hyperlinks of shapes
    const { openPkg } = await import('./addsheets-objects-helpers');
    const p = await openPkg(res.bytes);
    const drawing = [...p.files.entries()].find(([n]) => /^xl\/drawings\/drawing\d+\.xml$/.test(n))![1];
    expect(drawing).toMatch(/textlink="'?ob-shapes_図形'?!\$A\$3"/);
    const rels = [...p.files.entries()].find(([n]) => /^xl\/drawings\/_rels\/drawing\d+\.xml\.rels$/.test(n))![1];
    expect(rels).toContain("Target=\"#'ob-shapes_飛び先'!A1\"");
    expect(rels).toContain("Target=\"#'ob-shapes_飛び先'!B2\"");
    expect(res.report.items.filter((i) => i.code === 'sheetCopyLimits')).toEqual([]);
  });
});

describe('Excel-authored fixtures: notes and threaded comments', () => {
  it('ob-comments: notes are copied (visible / hidden, authors kept), threaded comments become notes with their text, VML ids are workbook-unique', async () => {
    const { res, pairs, added } = await copyOf(src('ob-comments.xlsx'));
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const main = pairs.find((p) => p.from.name === 'コメント')!;
    expect(main.from.notes).toBe(5); // 3 notes + the fallback notes of 2 threaded comments
    expect(main.from.threaded).toBe(2);
    expect(main.to.notes).toBe(5);
    expect(main.to.threaded).toBe(0);
    const only = pairs.find((p) => p.from.name === 'メモのみ')!;
    expect(only.to.notes).toBe(only.from.notes);
    const { openPkg } = await import('./addsheets-objects-helpers');
    const p = await openPkg(res.bytes);
    const c = [...p.files.entries()].filter(([n]) => /^xl\/comments\d+\.xml$/.test(n)).map(([, x]) => x).join('\n');
    expect(c).toContain('確認済み（メモ）');
    expect(c).toContain('スレッドコメントの本文\n返信（');
    expect(c).not.toContain('tc=');
    expect(c).not.toContain('スレッド化されたコメント]');
    // visible note stays visible
    const vml = [...p.files.entries()].filter(([n]) => /\.vml$/.test(n)).map(([, x]) => x).join('\n');
    expect(vml.match(/visibility:visible/g)?.length).toBe(1);
    const w = warnings(res, added[0].name);
    expect(w).toHaveLength(1);
    expect(w[0].message).toMatch(/スレッドコメント 2 件は、通常のコメント（メモ）にしました/);
  });
});

describe('Excel-authored fixtures: tables', () => {
  it('ob-table: tables keep ranges / totals / columns / style, structured references stay formulas, the custom table style is copied', async () => {
    const { res, pairs } = await copyOf(src('ob-table.xlsx'));
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const t = pairs.find((p) => p.from.name === 'テーブル')!;
    const byName = (a: { name: string }, b: { name: string }) => a.name.localeCompare(b.name);
    expect([...t.to.tables].sort(byName)).toEqual([...t.from.tables].sort(byName));
    expect(t.to.tables.map((x) => x.name)).toEqual(['T売上', 'T小計']); // in the order of the sheet's tableParts
    const s = pairs.find((p) => p.from.name === '集計')!;
    expect(s.to.charts).toBe(1);
    const out = await readBook(res.bytes, 'o.xlsx');
    const g = out.sheets.find((x) => x.name === 'ob-table_集計')!;
    expect(g.rows.slice(0, 6).map((r) => r?.[0]?.f)).toEqual(['SUM(T売上[金額])', 'T売上[[#Totals],[金額]]', 'INDEX(T売上[商品],2)', 'SUM(金額列)', 'ROWS(T売上[])', 'SUM(T小計[値])']);
    expect(g.rows[6]![0]!.f).toBe("COUNTA('ob-table_テーブル'!A:A)");
    expect(out.names.find((n) => n.name === '金額列')).toMatchObject({ scope: 'ob-table_集計', refersTo: 'T売上[金額]' });
    // the custom style and its dxfs
    const { openPkg } = await import('./addsheets-objects-helpers');
    const p = await openPkg(res.bytes);
    expect(p.files.get('xl/styles.xml')).toMatch(/<tableStyle name="マイ表スタイル"[^>]*count="2">/);
    expect(res.report.items.filter((i) => i.code === 'sheetCopyLimits')).toEqual([]);
  });

  it('a second copy of the same file: table names get _2, the structured references of the second copy read the second tables', async () => {
    const f = src('ob-table.xlsx');
    const res = await premergeCopies(base(), [f, { ...f, fileName: 'ob-table2.xlsx' }], { dropOtherSheets: true });
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const to = await objectsOf(res.bytes);
    expect(to.flatMap((s) => s.tables.map((t) => t.name)).sort()).toEqual(['T売上', 'T売上_2', 'T小計', 'T小計_2']);
    const out = await readBook(res.bytes, 'o.xlsx');
    expect(out.sheets.find((x) => x.name === 'ob-table2_集計')!.rows[0]![0]!.f).toBe('SUM(T売上_2[金額])');
    expect(out.sheets.find((x) => x.name === 'ob-table_集計')!.rows[0]![0]!.f).toBe('SUM(T売上[金額])');
    // an info item, not a limit of the copy: the sheet stays "complete"
    const w = res.report.items.filter((i) => i.code === 'tableRenamed');
    expect(w.map((i) => i.message)).toEqual([
      'シート「ob-table2_テーブル」のテーブル「T売上」は、名前が重なったため「T売上_2」にしました（数式の参照も付け替えました）。',
      'シート「ob-table2_テーブル」のテーブル「T小計」は、名前が重なったため「T小計_2」にしました（数式の参照も付け替えました）。',
    ]);
    expect(w[0].severity).toBe('info');
    expect(res.report.items.filter((i) => i.code === 'sheetCopyLimits')).toEqual([]);
    expect(res.layout.addedSheets!.every((s) => s.complete)).toBe(true);
  });
});

describe('Excel-authored fixtures: pivot tables', () => {
  it('ob-pivot: range and Table sources, a shared cache, number formats, a pivot chart', async () => {
    const { res, pairs } = await copyOf(src('ob-pivot.xlsx'));
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const pv1 = pairs.find((p) => p.from.name === '集計1')!;
    expect(pv1.to.pivots.map((x) => x.name)).toEqual(pv1.from.pivots.map((x) => x.name));
    expect(pv1.to.pivots.map((x) => x.source)).toEqual(['ob-pivot_元データ!A1:E25', 'ob-pivot_元データ!A1:E25']);
    expect(pv1.to.pivots[0].cacheId).toBe(pv1.to.pivots[1].cacheId); // shared cache
    const pv2 = pairs.find((p) => p.from.name === '集計2')!;
    expect(pv2.to.pivots.map((x) => x.source)).toEqual(['!T元']);
    expect(pv2.to.pivots[0].cacheId).not.toBe(pv1.to.pivots[0].cacheId);
    expect(pv1.to.charts).toBe(1);
    expect(pv1.to.chartFormulas.every((f) => f.startsWith("'ob-pivot_集計1'!"))).toBe(true);
    expect(res.report.items.filter((i) => i.code === 'sheetCopyLimits')).toEqual([]);
    // the pivot chart follows its pivot table
    const { openPkg } = await import('./addsheets-objects-helpers');
    const p = await openPkg(res.bytes);
    const chart = [...p.files.entries()].find(([n]) => /^xl\/charts\/chart\d+\.xml$/.test(n))![1];
    expect(chart).toContain('<c:name>[ob-pivot.xlsx]ob-pivot_集計1!ピボット範囲</c:name>');
  });

  it('samples/01_学生名簿.xlsx (chart + pivot) as the other file', async () => {
    const file: FileIn = { bytes: readFix('samples', '01_学生名簿.xlsx'), fileName: '01_学生名簿.xlsx' };
    const b: FileIn = { bytes: readFix('samples', '02_面接評価.xlsx'), fileName: '02_面接評価.xlsx' };
    const { res, pairs } = await copyOf(file, { base: b });
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const total = pairs.reduce((n, p) => n + p.to.charts, 0);
    expect(total).toBe(pairs.reduce((n, p) => n + p.from.charts, 0));
    expect(total).toBeGreaterThanOrEqual(1);
    expect(pairs.reduce((n, p) => n + p.to.pivots.length, 0)).toBe(pairs.reduce((n, p) => n + p.from.pivots.length, 0));
    expect(res.report.items.filter((i) => i.code === 'sheetCopyLimits')).toEqual([]);
  });
});

describe('Excel-authored fixtures: what is not copied', () => {
  it('ob-slicer-control: slicers and form controls are left out and named; the table, the pivot table and the notes are copied', async () => {
    const { res, pairs, added } = await copyOf(src('ob-slicer-control.xlsx'));
    expect(lintObjectsText(await lintObjects(res.bytes))).toBe('');
    const main = pairs.find((p) => p.from.name === 'スライサー')!;
    expect(main.from.slicers).toBe(1);
    expect(main.from.vmlControls).toBe(2);
    expect(main.to.slicers).toBe(0);
    expect(main.to.vmlControls).toBe(0);
    expect(main.to.controlPlaceholders).toBe(0);
    expect(main.to.notes).toBe(main.from.notes);
    expect(main.to.tables).toEqual(main.from.tables);
    expect(pairs.find((p) => p.from.name === 'ピボット')!.to.pivots).toHaveLength(1);
    // one warning per sheet, with what is missing and how many
    const w1 = warnings(res, added[0].name);
    expect(w1).toHaveLength(1);
    expect(w1[0].message).toMatch(/スライサー・タイムライン 1 個/);
    expect(w1[0].message).toMatch(/フォームコントロール 2 個/);
    expect(w1[0].message).not.toMatch(/テーブル|コメント/);
    const w2 = warnings(res, added[1].name);
    expect(w2).toHaveLength(1);
    expect(w2[0].message).toMatch(/スライサー・タイムライン 1 個/);
    // no slicer part or control property part anywhere
    const { openPkg } = await import('./addsheets-objects-helpers');
    const p = await openPkg(res.bytes);
    expect(p.names.filter((n) => /slicer|ctrlProp|activeX/i.test(n))).toEqual([]);
  });
});

describe('Excel-authored fixtures: Strict Open XML Spreadsheet', () => {
  const strictOf = async (bytes: Uint8Array) => (await openPkg(bytes)).files.get('xl/workbook.xml')!.includes('purl.oclc.org');
  const noLint = async (bytes: Uint8Array) => lintObjectsText((await lintObjects(bytes)).filter((l) => !/^(APP_|VML_ID_DUPLICATE)/.test(l.code)));
  // (docProps/app.xml of a Strict base is not updated with the new sheet names: addsheets.ts updateAppSheets looks for the transitional relationship type /
  // namespace; Excel rebuilds the part on save and does not read the sheet list from it — reported to the lead, not part of the object phase).
  // VML_ID_DUPLICATE: the notes of a Strict source copied into a Strict base keep their VML numbers (see copyComments): the same ids in two sheets, as Excel's own Strict reader wants them.

  it('a Strict source into a transitional base: notes, tables and pivot tables are copied in the namespaces of the base; charts, pictures and shapes are left out and named (their values differ in Strict)', async () => {
    expect(await strictOf(src('ob-strict.xlsx').bytes)).toBe(true);
    const { res, pairs, added } = await copyOf(src('ob-strict.xlsx'));
    expect(await strictOf(res.bytes)).toBe(false);
    expect(await noLint(res.bytes)).toBe('');
    const data = pairs.find((p) => p.from.name === 'データ')!;
    expect([data.to.charts, data.to.pictures, data.to.shapes]).toEqual([0, 0, 0]);
    expect(data.from.charts).toBe(2);
    expect([data.to.notes, data.to.tables.length]).toEqual([data.from.notes, data.from.tables.length]);
    expect(pairs.find((p) => p.from.name === '集計')!.to.pivots).toHaveLength(1);
    const p = await openPkg(res.bytes);
    for (const [n, x] of p.files) if (/^xl\/(tables|pivotTables|pivotCache|comments)/.test(n) && !/\.(rels|vml)$/.test(n)) expect(x, n).not.toContain('purl.oclc.org');
    for (const [n, x] of p.files) if (/\.rels$/.test(n)) expect(x, n).not.toContain('purl.oclc.org');
    const comments = [...p.files.entries()].find(([n]) => /^xl\/comments\d+\.xml$/.test(n))![1];
    expect(comments).not.toContain('commentPr');
    // a transitional package links no comment to its VML shape by number (Excel removes comments that carry one)
    expect(comments).toMatch(/<comment ref="B2" authorId="0" shapeId="0"/);
    const w = warnings(res, added[0].name);
    expect(w).toHaveLength(1);
    expect(w[0].message).toMatch(/次のものはコピーできませんでした: グラフ・画像・図形（Strict 形式のファイルと通常の形式のファイルの間では写せません）/);
    expect(w[0].message).not.toMatch(/コントロール|コメント|テーブル/);
  });

  it('transitional sources into a Strict base: the copies of the notes / tables / pivot tables are written in the Strict namespaces; their drawings are left out', async () => {
    const strictBase: FileIn = { ...src('ob-strict.xlsx'), sheet: 'データ' };
    const res = await premergeCopies(strictBase, [src('ob-comments.xlsx'), src('ob-table.xlsx'), src('ob-pivot.xlsx')], { dropOtherSheets: false });
    expect(await strictOf(res.bytes)).toBe(true);
    expect(await noLint(res.bytes)).toBe('');
    const p = await openPkg(res.bytes);
    for (const [n, x] of p.files) {
      if (/^xl\/(tables|pivotTables|pivotCache|comments\d)/.test(n) && !/\.(rels|vml)$/.test(n)) expect(x, n).not.toContain('schemas.openxmlformats.org/spreadsheetml');
      if (/\.rels$/.test(n) && /^xl\/(worksheets|pivot)/.test(n)) expect(x.replace(/<Relationship [^>]*vmlDrawing"[^>]*>/g, ''), n).not.toMatch(/Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships/);
    }
    const objs = await objectsOf(res.bytes);
    expect(objs.find((s) => s.name === 'ob-table_テーブル')!.tables).toHaveLength(2);
    expect(objs.find((s) => s.name === 'ob-table_集計')!.charts).toBe(0);
    expect(objs.find((s) => s.name === 'ob-pivot_集計1')!.pivots).toHaveLength(2);
    expect(objs.find((s) => s.name === 'ob-comments_コメント')!.notes).toBe(5);
    expect(res.report.items.filter((i) => i.code === 'sheetCopyLimits').map((i) => i.message).join('\n')).toMatch(/グラフ・画像・図形（Strict 形式/);
  });

  it('a Strict source into a Strict base: charts, pictures, shapes are copied as they are (same namespaces)', async () => {
    const strictBase: FileIn = { ...src('ob-strict.xlsx'), sheet: 'データ' };
    const res = await premergeCopies(strictBase, [{ ...src('ob-strict.xlsx'), fileName: 'ob-strict-2.xlsx' }], { dropOtherSheets: false });
    expect(await strictOf(res.bytes)).toBe(true);
    expect(await noLint(res.bytes)).toBe('');
    const objs = await objectsOf(res.bytes);
    const orig = objs.find((s) => s.name === 'データ')!;
    const copy = objs.find((s) => s.name === 'ob-strict-2_データ')!;
    expect([copy.charts, copy.pictures, copy.shapes, copy.notes, copy.tables.length]).toEqual([orig.charts, orig.pictures, orig.shapes, orig.notes, orig.tables.length]);
    expect(orig.charts).toBe(2);
    expect(copy.chartFormulas).toEqual(orig.chartFormulas.map((f) => f.replace(/^データ!/, "'ob-strict-2_データ'!")));
    expect(copy.tables[0].name).toBe('T厳密_2');
    // a Strict package links every comment to its VML shape by number
    const px = await openPkg(res.bytes);
    for (const [n, x] of px.files) {
      if (!/^xl\/comments\d+\.xml$/.test(n)) continue;
      const vml = [...px.files.entries()].filter(([m]) => /\.vml$/.test(m)).map(([, v]) => v).join('');
      for (const m of x.matchAll(/shapeId="(\d+)"/g)) expect(m[1] !== '0' && vml.includes(`_x0000_s${m[1]}"`), `${n}: shapeId ${m[1]}`).toBe(true);
    }
    const p = await openPkg(res.bytes);
    for (const [n, x] of p.files) if (/^xl\/(drawings|charts)\/[^/]+\.xml$/.test(n)) expect(x, n).not.toContain('schemas.openxmlformats.org/drawingml');
    expect(res.report.items.filter((i) => i.code === 'sheetCopyLimits')).toEqual([]);
    expect(res.report.items.filter((i) => i.code === 'tableRenamed')).toHaveLength(1);
  });
});
