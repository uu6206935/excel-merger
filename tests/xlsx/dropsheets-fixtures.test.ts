/**
 * Single-sheet output (PremergeOptions.dropOtherSheets) on Excel-authored workbooks: samples/01_学生名簿.xlsx merged with the other
 * samples, tests/fixtures/golden/torture.xlsx (every feature on the Data sheet, formulas / names / tables / pivot on the others),
 * the golden premerge bases and the multi-sheet files of the foreign corpus (tests/fixtures/foreign/corpus).
 * Every output: one worksheet, a clean package (tests/xlsx/dropsheets-helpers.ts lintSingleSheet + the structural linter of the
 * foreign-file tests), no formula that names a removed sheet or reads #REF!, consistent names, and the doc / editor built on it
 * works: a random op sequence replays to exactly what the doc shows.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { createEditor } from '../../src/core/doc';
import { analyzeSources, defaultPlan, detectRegion, joinSources, pickDataSheet } from '../../src/core/keys';
import { buildPremerged, readBook, replayLog } from '../../src/core/xlsx';
import type { KeySource, PremergeInput, PremergeResult } from '../../src/core/contracts';
import type { JoinType, MergedDoc, SourceBook } from '../../src/core/types';
import { compare } from '../doc/compare';
import { rng, step } from '../doc/fuzzcore';
import { SAMPLES_DIR, samplesAvailable } from '../doc/sampleMerge';
import { formatProblems, lintPackage, newProblems } from '../fixtures/foreign/ooxml-lint';
import { GOLDEN_FIX, insideTableScenario, scenario } from '../golden/premerge-scenario';
import { baseOptions, baseSource, files, lintSingleSheet, lintText, premergeOne, readFixture, readsAnySheet, sheetParts } from './dropsheets-helpers';

const here = path.dirname(fileURLToPath(import.meta.url));
const CORPUS = path.resolve(here, '../fixtures/foreign/corpus');

/** every formula text of a book (cells, rules, names, links) */
function formulaTexts(book: SourceBook): string[] {
  const texts: string[] = [];
  for (const sh of book.sheets) {
    for (const row of sh.rows) for (const c of row ?? []) if (c?.f !== undefined) texts.push(c.f);
    for (const r of sh.features.rules) texts.push(...r.formulas);
    for (const h of sh.features.hyperlinks) if (h.location) texts.push(h.location);
  }
  for (const n of book.names) texts.push(n.refersTo);
  return texts;
}

/** what every single-sheet output has to satisfy */
async function expectClean(bytes: Uint8Array, removed: string[], inputBytes?: Uint8Array): Promise<SourceBook> {
  expect(lintText(await lintSingleSheet(bytes, removed))).toBe('');
  const lint = await lintPackage(bytes);
  const before = inputBytes ? await lintPackage(inputBytes) : [];
  expect(formatProblems(newProblems(before, lint))).toBe('');
  const book = await readBook(bytes, 'out.xlsx', { id: 'out' });
  expect(book.sheets).toHaveLength(1);
  const texts = formulaTexts(book);
  expect(texts.filter((t) => readsAnySheet(t, removed))).toEqual([]);
  // no #REF! that the input did not have
  const had = new Set(inputBytes ? formulaTexts(await readBook(inputBytes, 'in.xlsx', { id: 'in' })) : []);
  expect(texts.filter((t) => t.includes('#REF!') && !had.has(t))).toEqual([]);
  // names: every scope is the merged sheet or the workbook
  for (const n of book.names) expect(n.scope === null || n.scope === book.sheets[0].name).toBe(true);
  return book;
}

// ---------------------------------------------------------------------------------------------------
// samples/01 merged with the other samples
// ---------------------------------------------------------------------------------------------------

const SAMPLE_LABELS = ['02_面接評価.xlsx', '03_連絡先.csv', '05_研修出欠.xlsx'];

async function sampleInput(joinType: JoinType, options: Partial<PremergeInput['options']>, baseName = '01_学生名簿.xlsx', otherNames = SAMPLE_LABELS): Promise<PremergeInput> {
  const load = async (name: string, id: string): Promise<KeySource> => {
    const book = await readBook(new Uint8Array(fs.readFileSync(path.join(SAMPLES_DIR, name))), name, { id });
    return { id, book, region: detectRegion(book, pickDataSheet(book)), label: name.replace(/\.[^.]+$/, '') };
  };
  const base = await load(baseName, 's1');
  const others = await Promise.all(otherNames.map((n, i) => load(n, `s${i + 2}`)));
  const analysis = analyzeSources(base, others);
  const join = joinSources(analysis, { ...defaultPlan(analysis), joinType });
  return { base, sources: others, rows: join.rows, imports: join.defaultImports, keyFill: join.keyFill, options: baseOptions(options) };
}

async function editorOf(input: PremergeInput, res: PremergeResult) {
  const book = await readBook(res.bytes, 'merged.xlsx', { id: 'merged' });
  const sourceInfo: MergedDoc['sourceInfo'] = {};
  [input.base, ...input.sources].forEach((k, i) => (sourceInfo[k.id] = { label: k.label, color: ['#2563eb', '#15803d', '#c2410c', '#7c3aed'][i % 4] }));
  let last = input.base.region.lastDataRow;
  for (const a of res.layout.appended) if (a.row > last) last = a.row;
  const region = { headerRow: input.base.region.headerRow, firstDataRow: input.base.region.firstDataRow, lastDataRow: last };
  return { book, snapshot: createEditor(book, res.layout, sourceInfo, region) };
}

describe.skipIf(!samplesAvailable())('samples/01_学生名簿.xlsx merged with the other samples, single-sheet output', () => {
  it('keeps only 名簿: formulas that read マスタ / 基準日 became values, the list validation is a literal list', async () => {
    const input = await sampleInput('left', { dropOtherSheets: true });
    const res = await buildPremerged(input);
    const inBook = input.base.book;
    expect(inBook.sheets.map((s) => s.name)).toEqual(['名簿', '集計', 'ピボット', 'マスタ']);
    const book = await expectClean(res.bytes, ['集計', 'ピボット', 'マスタ'], input.base.book.bytes);
    expect(book.sheets[0].name).toBe('名簿');
    expect(res.report.droppedSheets).toEqual(['集計', 'ピボット', 'マスタ']);
    expect(res.report.formulasDetached).toBe(60);
    const det = res.report.items.find((i) => i.code === 'detachedFormulas')!;
    expect(det).toMatchObject({ severity: 'info', message: '他のシートを参照していた数式 60 個を値にしました' });
    expect(det.rows).toEqual(Array.from({ length: 20 }, (_x, i) => 4 + i));
    expect(res.report.items.find((i) => i.code === 'dvToList')).toMatchObject({ severity: 'info', message: '入力規則 1 個を固定のリストにしました' });
    expect(res.report.items.find((i) => i.code === 'sheetsDropped')).toBeTruthy();
    expect(res.report.items.every((i) => /[ぁ-んァ-ン一-龥]/.test(i.message))).toBe(true);
    // DATEDIF(I4,基準日,"Y") in column J: the cached ages stay as plain values
    const s = book.sheets[0];
    expect(s.rows[3]?.[9]).toMatchObject({ t: 'n', v: 20 });
    expect(s.rows[3]?.[9]?.f).toBeUndefined();
    // the other formulas of the sheet (COUNTIF in L, the imported same-row formulas) are still live
    expect(s.rows[3]?.[11]?.f).toBe('COUNTIF($B$4:$B$63,B4)');
    // names: the one that points at 名簿 stays, the ones that point at マスタ are gone; _FilterDatabase follows the autofilter
    expect(book.names.map((n) => `${n.scope ?? ''}!${n.name}`).sort()).toEqual(['!志望度列', '名簿!_xlnm._FilterDatabase']);
    expect(book.names.find((n) => n.name === '志望度列')!.refersTo).toBe('名簿!$K:$K');
    // the university list validation reads the values of マスタ!A2:A20 and holds them literally
    const master = inBook.sheets.find((x) => x.name === 'マスタ')!;
    const univ: string[] = [];
    for (let r = 2; r <= 20; r++) { const c = master.rows[r - 1]?.[0]; if (c && c.v !== null && c.v !== '') univ.push(String(c.v)); }
    const dv = s.features.rules.filter((r) => r.kind === 'dv' && r.type === 'list');
    expect(dv.map((r) => r.formulas[0]).sort()).toEqual([`"${univ.join(',')}"`, '"1,2,3,4,5"'].sort());
    expect(dv.every((r) => !r.ext)).toBe(true);
    // CF / merge / freeze / autofilter of the sheet are untouched (autofilter grew over the imported columns)
    expect(s.features.rules.filter((r) => r.kind === 'cf').map((r) => r.type).sort()).toEqual(['cellIs', 'expression']);
    expect(s.features.merges).toHaveLength(1);
    expect(s.features.freeze).toEqual({ cols: 2, rows: 3 });
    expect(s.features.autoFilter).toBe('A3:X63');
  });

  it('the package has only the parts of one worksheet and shows the sheet names it keeps in docProps/app.xml', async () => {
    const res = await buildPremerged(await sampleInput('left', { dropOtherSheets: true }));
    const f = await files(res.bytes);
    expect([...f.keys()].sort()).toEqual([
      '[Content_Types].xml', '_rels/.rels', 'docProps/app.xml', 'docProps/core.xml', 'xl/_rels/workbook.xml.rels', 'xl/sharedStrings.xml', 'xl/styles.xml', 'xl/theme/theme1.xml',
      'xl/workbook.xml', 'xl/worksheets/sheet1.xml',
    ]);
    const wb = f.get('xl/workbook.xml')!;
    expect(wb).not.toMatch(/pivotCaches|activeTab|firstSheet/);
    expect(wb).toContain('<sheets><sheet name="名簿" sheetId="4" r:id="rId1"/></sheets>');
    const app = f.get('docProps/app.xml')!;
    expect(app).toContain('<vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>名前付き一覧</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4>');
    expect(app).toContain('<TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>名簿</vt:lpstr><vt:lpstr>志望度列</vt:lpstr></vt:vector></TitlesOfParts>');
    expect(wb).toContain('<calcPr calcId="191029" fullCalcOnLoad="1"/>');
  });

  it('join full: rows below the table move with the appended rows, the rows named in the report are the final ones', async () => {
    const input = await sampleInput('full', { dropOtherSheets: true });
    const res = await buildPremerged(input);
    const book = await expectClean(res.bytes, ['集計', 'ピボット', 'マスタ'], input.base.book.bytes);
    expect(res.report.droppedSheets).toEqual(['集計', 'ピボット', 'マスタ']);
    const n = res.report.appendedRows;
    expect(n).toBeGreaterThan(0);
    // no detached formula sits below the table in this sample: the rows are the same as without appended rows
    expect(res.report.items.find((i) => i.code === 'detachedFormulas')!.rows![0]).toBe(4);
    expect(book.sheets[0].maxRow).toBeGreaterThanOrEqual(66 + n);
  });

  it('default behaviour is unchanged: every sheet, pivot, chart and name stays (dropOtherSheets false or absent)', async () => {
    const a = await buildPremerged(await sampleInput('left', {}));
    const b = await buildPremerged(await sampleInput('left', { dropOtherSheets: false }));
    expect([...(await files(a.bytes))]).toEqual([...(await files(b.bytes))]);
    expect(a.report.droppedSheets).toBeUndefined();
    expect(a.report.formulasDetached).toBeUndefined();
    const book = await readBook(a.bytes, 'x.xlsx');
    expect(book.sheets.map((s) => s.name)).toEqual(['名簿', '集計', 'ピボット', 'マスタ']);
    expect(book.names.map((n) => n.name)).toEqual(expect.arrayContaining(['基準日', '大学一覧', '志望度列']));
    expect(book.features).toMatchObject({ charts: 1, pivotTables: 1 });
    expect(a.report.items.some((i) => i.code === 'detachedFormulas' || i.code === 'sheetsDropped')).toBe(false);
  });

  it('the doc / editor built on the output has no other sheets and replays like the doc says (random op sequences)', async () => {
    const input = await sampleInput('full', { dropOtherSheets: true });
    const res = await buildPremerged(input);
    const { book, snapshot } = await editorOf(input, res);
    expect(snapshot.doc.sheetNames).toEqual(['名簿']);
    expect(snapshot.doc.externalFormulas).toEqual([]);
    expect(snapshot.doc.externalRules).toEqual([]);
    expect(snapshot.doc.names.map((n) => n.name).sort()).toEqual(['_xlnm._FilterDatabase', '志望度列'].sort());
    expect(compare(snapshot.doc, book)).toEqual([]);
    const failures: string[] = [];
    for (const seed of Array.from({ length: 12 }, (_x, i) => 9100 + i)) {
      const r = rng(seed);
      let s = snapshot;
      const trail: string[] = [];
      for (let i = 0; i < 10; i++) {
        const { res: o, what } = step(s, r, false, s.doc.lastDataRow + 2);
        trail.push(o.ok ? what : `${what} [refused]`);
        if (o.ok) s = o.snapshot;
      }
      if (!s.log.length) continue;
      let diffs: string[];
      try {
        diffs = compare(s.doc, await readBook(await replayLog(res.bytes, s.log), 'after.xlsx'));
      } catch (e) {
        diffs = [`replay threw: ${(e as Error).message}`];
      }
      if (diffs.length) failures.push(`seed ${seed}: ${diffs.slice(0, 3).join(' | ')}\n  ops: ${trail.join('; ')}`);
    }
    expect(failures).toEqual([]);
  }, 300_000);

  it('works as a base with a table on another sheet (02_面接評価.xlsx: 面接評価 + 評価テーブル)', async () => {
    const input = await sampleInput('left', { dropOtherSheets: true }, '02_面接評価.xlsx', ['01_学生名簿.xlsx']);
    const res = await buildPremerged(input);
    const dropped = input.base.book.sheets.map((s) => s.name).filter((n) => n !== res.layout.sheet);
    expect(res.report.droppedSheets).toEqual(dropped);
    await expectClean(res.bytes, dropped, input.base.book.bytes);
  });
});

// ---------------------------------------------------------------------------------------------------
// torture.xlsx
// ---------------------------------------------------------------------------------------------------

describe('torture.xlsx (Data + 集計 シート + Tbl + Pivot + hidden Lists), single-sheet output', () => {
  const bytes = readFixture(GOLDEN_FIX, 'torture.xlsx');
  const REMOVED = ['集計 シート', 'Tbl', 'Pivot', 'Lists'];

  it('drops the four other sheets with the pivot, the table and their caches; keeps the chart, picture, notes and threaded comment of Data', async () => {
    const base = await baseSource({ bytes, fileName: 'torture.xlsx', sheet: 'Data' });
    const res = await premergeOne(base);
    expect(res.report.droppedSheets).toEqual(REMOVED);
    const book = await expectClean(res.bytes, REMOVED, bytes);
    const f = await files(res.bytes);
    const keys = [...f.keys()];
    expect(keys.filter((k) => /pivot|tables\/|worksheets\/sheet[2-9]|calcChain/i.test(k))).toEqual([]);
    for (const k of ['xl/charts/chart1.xml', 'xl/drawings/drawing1.xml', 'xl/media/image1.png', 'xl/comments1.xml', 'xl/threadedComments/threadedComment1.xml', 'xl/persons/person.xml', 'xl/metadata.xml', 'xl/drawings/vmlDrawing1.vml']) expect(keys).toContain(k);
    expect(book.features).toMatchObject({ charts: 1, pivotTables: 0, threadedComments: 1, images: 1 });
  });

  it('detaches the 42 formulas that read a removed sheet (direct, through TaxRate, a whole column, a table) and keeps the others live', async () => {
    const base = await baseSource({ bytes, fileName: 'torture.xlsx', sheet: 'Data' });
    const res = await premergeOne(base);
    expect(res.report.formulasDetached).toBe(42);
    expect(res.report.items.find((i) => i.code === 'detachedFormulas')!.message).toBe('他のシートを参照していた数式 42 個を値にしました');
    const s = (await readBook(res.bytes, 'o.xlsx')).sheets[0];
    const inp = (await readBook(bytes, 'i.xlsx')).sheets[0];
    const at = (sh: typeof s, ref: string) => { const m = /^([A-Z]+)(\d+)$/.exec(ref)!; return sh.rows[Number(m[2]) - 1]?.[m[1].split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1]; };
    // I = E*(1+TaxRate): TaxRate -> '集計 シート'!B2 ; O = '集計 シート'!B2*E ; R6 = SUM(SalesTbl[金額]) ; R7 = COUNTA(Lists!A:A)
    for (const ref of ['I4', 'I23', 'O4', 'O23', 'R6', 'R7']) {
      expect(at(s, ref)?.f, ref).toBeUndefined();
      expect(at(s, ref)?.v, ref).toEqual(at(inp, ref)?.v);
    }
    expect(at(s, 'R7')).toMatchObject({ t: 'n', v: 5 });
    // not reading a removed sheet: stay formulas (named range in the sheet, whole-column name of Data, errors, CSE and dynamic arrays)
    for (const ref of ['R4', 'R5', 'R8', 'S4', 'T4', 'Q4', 'Q5', 'N4', 'P4']) expect(at(s, ref)?.f, ref).toBe(at(inp, ref)?.f);
    expect(at(s, 'N4')).toMatchObject({ arr: 'N4' });
    expect(at(s, 'P4')).toMatchObject({ dyn: true, arr: 'P4:P8' });
    expect(at(s, 'Q4')).toMatchObject({ t: 'e', v: '#DIV/0!' });
  });

  it('names: the ones that read a removed sheet are gone, the scoped / built-in ones of Data are kept with sheet index 0', async () => {
    const base = await baseSource({ bytes, fileName: 'torture.xlsx', sheet: 'Data' });
    const res = await premergeOne(base);
    const f = await files(res.bytes);
    const wb = f.get('xl/workbook.xml')!;
    const names = [...wb.matchAll(/<definedName\b([^>]*)>([^<]*)</g)].map((m) => `${/name="([^"]*)"/.exec(m[1])![1]}|${/localSheetId="(\d+)"/.exec(m[1])?.[1] ?? ''}|${m[2]}`).sort();
    expect(names).toEqual([
      'AmountCol||Data!$E:$E',
      'LocalRange|0|Data!$A$4:$A$23',
      '_xlnm.Print_Area|0|Data!$A$1:$J$23',
      '_xlnm.Print_Titles|0|Data!$3:$3',
      '_xlnm._FilterDatabase|0|Data!$A$3:$J$23',
    ].sort());
    expect(wb).not.toContain('TaxRate');
    const app = f.get('docProps/app.xml')!;
    expect(app).toContain('<vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>名前付き一覧</vt:lpstr></vt:variant><vt:variant><vt:i4>4</vt:i4>');
    expect(app).toContain('<vt:lpstr>Data</vt:lpstr><vt:lpstr>AmountCol</vt:lpstr><vt:lpstr>Data!LocalRange</vt:lpstr><vt:lpstr>Data!Print_Area</vt:lpstr><vt:lpstr>Data!Print_Titles</vt:lpstr>');
  });

  it('validations / links: the hidden Lists sheet list is literal, the link to 集計 シート is a plain cell, the external link stays', async () => {
    const base = await baseSource({ bytes, fileName: 'torture.xlsx', sheet: 'Data' });
    const res = await premergeOne(base);
    expect(res.report.items.find((i) => i.code === 'dvToList')!.message).toBe('入力規則 1 個を固定のリストにしました');
    expect(res.report.items.find((i) => i.code === 'linksDropped')!.message).toBe('他のシートへのリンク 1 個を外しました');
    const s = (await readBook(res.bytes, 'o.xlsx')).sheets[0];
    const dv = s.features.rules.filter((r) => r.kind === 'dv').map((r) => `${r.sqref}|${r.type}|${r.ext ? 'x14' : ''}|${r.formulas.join(',')}`).sort();
    expect(dv).toEqual(['C4:C23|list||"東大,京大,慶應,早稲田,SFC"', 'G4:G23|whole||1,100']);
    expect(s.features.hyperlinks).toEqual([{ ref: 'J4', target: 'https://example.com/ja' }]);
    // CF (incl. the x14 data bar) is untouched
    expect(s.features.rules.filter((r) => r.kind === 'cf').map((r) => r.type).sort()).toEqual(['cellIs', 'colorScale', 'dataBar', 'dataBar', 'expression', 'iconSet']);
  });

  it('with appended rows and imported columns the premerge of the single-sheet package still works', async () => {
    const input = scenario('plain', { dropOtherSheets: true });
    const res = await buildPremerged(input);
    expect(res.report.droppedSheets).toEqual(['Report']);
    expect(res.report.formulasDetached).toBe(0);
    const book = await expectClean(res.bytes, ['Report']);
    const s = book.sheets[0];
    expect(s.features.autoFilter).toBe('A1:K7');
    // the 合計 formulas still grew over the appended row
    expect(s.rows[8]?.[3]?.f).toBe('SUM(D2:D7)');
    // the chart that read 学生!D2:D6 lives on Report: gone with it
    expect(book.features.charts).toBe(0);
  });
});

// ---------------------------------------------------------------------------------------------------
// golden premerge bases and the other Excel-authored fixtures
// ---------------------------------------------------------------------------------------------------

describe('Excel-authored golden fixtures, single-sheet output', () => {
  const cases: Array<{ file: string; sheet: string; removed: string[] }> = [
    { file: 'g-table.xlsx', sheet: '名簿', removed: ['Summary', 'Lists'] },
    { file: 'pm-base-plain.xlsx', sheet: '学生', removed: ['Report'] },
    { file: 'pm-base-table.xlsx', sheet: '学生', removed: ['Report'] },
    { file: 'r2-arrays.xlsx', sheet: 'A', removed: ['B'] },
    { file: 'r2-table.xlsx', sheet: 'T', removed: ['U'] },
    { file: 'r3-misc.xlsx', sheet: 'Band', removed: ['Q', 'Tb', 'Outline'] },
  ];
  for (const c of cases) {
    it(`${c.file}: ${c.sheet} stays, ${c.removed.join(' / ')} go`, async () => {
      const bytes = readFixture(GOLDEN_FIX, c.file);
      const base = await baseSource({ bytes, fileName: c.file, sheet: c.sheet });
      const res = await premergeOne(base);
      expect(res.report.droppedSheets).toEqual(c.removed);
      const book = await expectClean(res.bytes, c.removed, bytes);
      expect(book.sheets[0].name).toBe(c.sheet);
    });
  }

  it('g-table: the table of the merged sheet survives, a formula that uses it from another sheet is gone with that sheet', async () => {
    const bytes = readFixture(GOLDEN_FIX, 'g-table.xlsx');
    const base = await baseSource({ bytes, fileName: 'g-table.xlsx', sheet: '名簿' });
    const res = await premergeOne(base);
    const s = (await readBook(res.bytes, 'o.xlsx')).sheets[0];
    expect(s.features.tables.map((t) => t.name)).toEqual(['tblMain']);
    expect(res.report.formulasDetached).toBe(0);
  });

  it('r2-table: a name that points at the table of the merged sheet (ScoreCol = Tbl[Score]) is kept', async () => {
    const bytes = readFixture(GOLDEN_FIX, 'r2-table.xlsx');
    const base = await baseSource({ bytes, fileName: 'r2-table.xlsx', sheet: 'T' });
    const res = await premergeOne(base);
    expect((await readBook(res.bytes, 'o.xlsx')).names.map((n) => `${n.name}=${n.refersTo}`)).toEqual(['ScoreCol=Tbl[Score]']);
  });

  it('a base that is the second of several sheets: the sheet in front of it goes too, activeTab / firstSheet are reset', async () => {
    const bytes = readFixture(GOLDEN_FIX, 'r3-misc.xlsx');
    const base = await baseSource({ bytes, fileName: 'r3-misc.xlsx', sheet: 'Tb' });
    const res = await premergeOne(base);
    expect(res.report.droppedSheets).toEqual(['Band', 'Q', 'Outline']);
    const book = await expectClean(res.bytes, ['Band', 'Q', 'Outline'], bytes);
    expect(book.sheets[0].name).toBe('Tb');
    expect(book.sheets[0].features.tables).toHaveLength(1);
  });

  it('a premerge that was already a single-sheet workbook is a no-op: droppedSheets is empty', async () => {
    const first = await premergeOne(await baseSource({ bytes: readFixture(GOLDEN_FIX, 'pm-base-plain.xlsx'), fileName: 'p.xlsx', sheet: '学生' }));
    const second = await premergeOne(await baseSource({ bytes: first.bytes, fileName: 'p1.xlsx', sheet: '学生' }));
    expect(second.report.droppedSheets).toEqual([]);
    expect(second.report.formulasDetached).toBe(0);
    expect(second.report.items.some((i) => ['sheetsDropped', 'detachedFormulas'].includes(i.code))).toBe(false);
    await expectClean(second.bytes, []);
  });

  it('the golden insideTable scenario with dropOtherSheets: imported columns inside the table, Summary / Lists gone', async () => {
    const input = insideTableScenario();
    input.options = baseOptions({ dropOtherSheets: true });
    const res = await buildPremerged(input);
    expect(res.report.droppedSheets).toEqual(['Summary', 'Lists']);
    const book = await expectClean(res.bytes, ['Summary', 'Lists']);
    expect(book.sheets[0].features.tables[0].columns).toEqual(['ID', '氏名', '大学', '点数1', '点数2', '志望', 'Calc', '合計', '列1']);
  });
});

// ---------------------------------------------------------------------------------------------------
// the doc / editor on every single-sheet output
// ---------------------------------------------------------------------------------------------------

describe('doc / editor on the single-sheet outputs: random op sequences replay like the doc says', () => {
  const dsDir = path.resolve(here, '../fixtures/dropsheets');
  const subjects: Array<{ label: string; file: string; fileName: string; sheet: string }> = [
    { label: 'torture.xlsx', file: path.join(GOLDEN_FIX, 'torture.xlsx'), fileName: 'torture.xlsx', sheet: 'Data' },
    { label: 'g-table.xlsx', file: path.join(GOLDEN_FIX, 'g-table.xlsx'), fileName: 'g-table.xlsx', sheet: '名簿' },
    { label: 'pm-base-plain.xlsx', file: path.join(GOLDEN_FIX, 'pm-base-plain.xlsx'), fileName: 'pm-base-plain.xlsx', sheet: '学生' },
    { label: 'pm-base-table.xlsx', file: path.join(GOLDEN_FIX, 'pm-base-table.xlsx'), fileName: 'pm-base-table.xlsx', sheet: '学生' },
    { label: 'r2-arrays.xlsx', file: path.join(GOLDEN_FIX, 'r2-arrays.xlsx'), fileName: 'r2-arrays.xlsx', sheet: 'A' },
    { label: 'r2-table.xlsx', file: path.join(GOLDEN_FIX, 'r2-table.xlsx'), fileName: 'r2-table.xlsx', sheet: 'T' },
    { label: 'r3-misc.xlsx (Tb)', file: path.join(GOLDEN_FIX, 'r3-misc.xlsx'), fileName: 'r3-misc.xlsx', sheet: 'Tb' },
    { label: 'ds-table.xlsx', file: path.join(dsDir, 'ds-table.xlsx'), fileName: 'ds-table.xlsx', sheet: '名簿' },
    { label: 'hc-29-formula-variety.xlsx', file: path.join(CORPUS, 'hc-29-formula-variety.xlsx'), fileName: 'hc-29-formula-variety.xlsx', sheet: '名簿' },
    { label: 'py-04-openpyxl-table.xlsx', file: path.join(CORPUS, 'py-04-openpyxl-table.xlsx'), fileName: 'py-04-openpyxl-table.xlsx', sheet: '名簿' },
  ].filter((x) => fs.existsSync(x.file));
  for (const sub of subjects) {
    it(sub.label, async () => {
      const base = await baseSource({ bytes: readFixture(sub.file), fileName: sub.fileName, sheet: sub.sheet });
      const res = await premergeOne(base);
      const book = await readBook(res.bytes, 'merged.xlsx', { id: 'merged' });
      const snapshot = createEditor(book, res.layout, { b: { label: sub.label, color: '#2563eb' } }, { headerRow: base.region.headerRow, firstDataRow: base.region.firstDataRow, lastDataRow: base.region.lastDataRow });
      expect(snapshot.doc.sheetNames).toEqual([sub.sheet]);
      expect(snapshot.doc.externalFormulas).toEqual([]);
      expect(snapshot.doc.externalRules).toEqual([]);
      expect(compare(snapshot.doc, book)).toEqual([]);
      const failures: string[] = [];
      for (const seed of Array.from({ length: 6 }, (_x, i) => 7700 + i)) {
        const r = rng(seed * 31 + sub.label.length);
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
          expect(lintText(await lintSingleSheet(replayed, res.report.droppedSheets ?? []))).toBe('');
          diffs = compare(s.doc, await readBook(replayed, 'after.xlsx'));
        } catch (e) {
          diffs = [`replay threw: ${(e as Error).message}`];
        }
        if (diffs.length) failures.push(`seed ${seed}: ${diffs.slice(0, 3).join(' | ')}
  ops: ${trail.join('; ')}`);
      }
      expect(failures).toEqual([]);
    }, 300_000);
  }
});

// ---------------------------------------------------------------------------------------------------
// the foreign-file corpus: multi-sheet files from other producers
// ---------------------------------------------------------------------------------------------------

describe('foreign corpus files with several sheets, single-sheet output', () => {
  const corpus = fs.existsSync(CORPUS)
    ? ['hc-17-sheetnames-hidden-first.xlsx', 'hc-25-sheet-order-ids.xlsx', 'hc-29-formula-variety.xlsx', 'hc-32-sheetnames-tricky.xlsx', 'hc-34-formulas-without-cache.xlsx', 'hc-35-names-cf-dv.xlsx',
      'js-01-sheetjs-basic.xlsx', 'js-04-exceljs-basic.xlsx', 'py-04-openpyxl-table.xlsx', 'py-05-openpyxl-hidden.xlsx', 'py-11-openpyxl-chart-names-links.xlsx', 'py-15-xlsxwriter-basic.xlsx',
      'py-25-pandas-many-sheets.xlsx'].filter((f) => fs.existsSync(path.join(CORPUS, f)))
    : [];
  for (const file of corpus) {
    it(file, async () => {
      const bytes = readFixture(CORPUS, file);
      const book0 = await readBook(bytes, file, { id: 'b' });
      const base = await baseSource({ bytes, fileName: file });
      const res = await premergeOne(base);
      const removed = book0.sheets.map((s) => s.name).filter((n) => n !== res.layout.sheet);
      expect(res.report.droppedSheets).toEqual(removed);
      expect(removed.length).toBeGreaterThan(0);
      const out = await expectClean(res.bytes, removed, bytes);
      expect(out.sheets[0].name).toBe(res.layout.sheet);
      expect(sheetParts(await files(res.bytes))).toHaveLength(1);
    });
  }
});
