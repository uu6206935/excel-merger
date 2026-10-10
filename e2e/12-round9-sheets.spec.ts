// Round 9 ("plan A"): the output can ALSO hold every sheet of every file (owner: T).
//   「他のシートも出力に含める」 (the base file's other sheets) and 「他のファイルのシートも含める」 (the original sheets of the 2nd, 3rd … files, as they are)
//   (a) the default: the merged sheet only — both engines, check-output --single-sheet PASS, exactly one sheet
//   (b) both boxes: the merged sheet, the base's sheets, then <file>_<sheet> of every other file in card order (02_面接評価_面接評価 …) — preview list
//       = side panel list = the sheets of both engines' files; check-output --with-copies --source … PASS (every copy equals its source sheet)
//   (c) only the first box: the base's other sheets, no copies (what rounds 1-6 produced)
//   (d) the same file name twice: the later one gets "(1)" — preview, merge result and files agree
//   (e) imported formulas that read other sheets of their own file stay formulas and read the copies (g01_samples__shuf: 58 of them)
//   (f) edits of the merged sheet (move / insert / delete / type) never change a copied sheet: replayed by both engines, copies still equal their sources
//   (g) tools/check-output.mjs itself (the CLI): --single-sheet / --with-copies --source, and the old strictness without the flags
import fs from 'node:fs';
import path from 'node:path';
import { expect, test, type Page } from '@playwright/test';
import { excelAvailable } from '../tools/lib/inspect.mjs';
import {
  cardSources,
  checkOutput,
  columnIndex,
  columnTitles,
  describeVerdict,
  editCount,
  exportBoth,
  expectPass,
  expectSameFormulas,
  insertColumn,
  loadFiles,
  mergeNow,
  moveColumns,
  openApp,
  ROOT,
  SAMPLE_FILES,
  SAMPLES_DIR,
  samplePaths,
  selectColumns,
  setSheetMode,
  sheetModeOf,
  step,
  typeInCell,
  verifyAndProbe,
  waitMergeReady,
  watchErrors,
} from './helpers';
import { waitIdle } from './round5';

/** the sheets of the five samples with both boxes ticked (base 01: 名簿 / 集計 / ピボット / マスタ), in card order */
const ALL_NAMES = ['名簿', '集計', 'ピボット', 'マスタ', '02_面接評価_面接評価', '02_面接評価_評価テーブル', '03_連絡先_03_連絡先', '04_適性検査_適性検査結果', '05_研修出欠_出欠'];
const BASE_ONLY_NAMES = ['名簿', '集計', 'ピボット', 'マスタ'];
const FIXTURES = path.join(ROOT, 'tests', 'fixtures', 'formulaimport', 'cached');
const SCRATCH = path.join(ROOT, 'test-results', `e2e-round9-scratch-${process.pid}`);

test.afterAll(() => {
  fs.rmSync(SCRATCH, { recursive: true, force: true });
});

/** a copy of a sample in its own folder with its own modified time (the app treats the same name + size + time as the same file) */
function scratchCopy(folder: string, name: string, mtimeSec: number, source = path.join(SAMPLES_DIR, name)): string {
  const dir = path.join(SCRATCH, folder);
  fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, name);
  fs.copyFileSync(source, file);
  fs.utimesSync(file, mtimeSec, mtimeSec);
  return file;
}

async function plannedNames(page: Page): Promise<string[]> {
  return page.locator('[data-testid^="planned-sheet-"][data-name]').evaluateAll((els) => els.map((e) => e.getAttribute('data-name') ?? ''));
}
async function sheetListNames(page: Page): Promise<string[]> {
  const list = page.getByTestId('sheet-list');
  if (!(await list.count())) await page.getByTestId('sheet-list-details').evaluate((el) => ((el as HTMLDetailsElement).open = true)).catch(() => undefined);
  return page.locator('[data-testid^="sheet-list-item-"][data-name]').evaluateAll((els) => els.map((e) => e.getAttribute('data-name') ?? ''));
}

/** the sheet names of every output (from the Excel dumps of verifyAndProbe: needs Excel) */
function sheetNamesOfDumps(v: Awaited<ReturnType<typeof verifyAndProbe>>, out: { files: unknown[] }): string[][] {
  return v.dumps.slice(1, 1 + out.files.length).map((d) => [...(d.sheets ?? [])].sort((a, b) => a.index - b.index).map((s) => s.name));
}

test.describe('ラウンド9: 出力に入れるシート（他のシート・他のファイルのシート）', () => {
  test('(a) 既定は「結合した 1 シートだけ」: 両エンジンで 1 シート、check-output --single-sheet で PASS', async ({ page }) => {
    test.setTimeout(20 * 60_000);
    const errors = watchErrors(page);
    await openApp(page);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, samplePaths()));
    await waitMergeReady(page, errors);
    expect(await sheetModeOf(page), '既定は 2 つともオフ').toBe('single');
    await expect(page.getByTestId('include-other-sheets')).not.toBeChecked();
    await expect(page.getByTestId('include-source-sheets')).not.toBeChecked();
    await expect(page.getByTestId('preview-output-sheets'), '「出力されるシート」の一覧は、他のファイルのシートを入れるときだけ').toHaveCount(0);
    await mergeNow(page, errors);
    await expect(page.getByTestId('side-dropped-sheets'), '出力は 1 シートだけと伝える').toBeVisible();
    await expect(page.getByTestId('side-output-sheets')).toHaveCount(0);

    const out = await exportBoth(page, errors, 'e2e_r9a_1シート');
    expect(out.mode).toBe('single');
    test.skip(!excelAvailable(), 'Excel で開いての確認には Excel が必要です');
    const v = await verifyAndProbe(out);
    expectPass(v.verdicts, '既定（1 シート）');
    for (const [engine, verdict] of v.verdicts) {
      expect(verdict.items.map((i) => i.code), `${engine}: 基準のほかのシートがないことは情報として出る`).toContain('sheets-dropped');
      expect(verdict.items.map((i) => i.code)).toContain('no-dangling-ref');
    }
    for (const names of sheetNamesOfDumps(v, out)) expect(names, '出力のシートは結合した「名簿」だけ').toEqual(['名簿']);
    expectSameFormulas(out);
    // the old strictness still exists: asked WITHOUT --single-sheet the same file is a failure (a sheet is "lost")
    const strict = await checkOutput(out.files[0].file);
    expect(strict.items.filter((i) => i.level === 'fail').map((i) => i.code)).toContain('sheets-lost');
    expect(strict.items.find((i) => i.code === 'sheets-lost')!.message).toContain('--single-sheet');
  });

  test('(b) 「他のシートも出力に含める」を入れる → 結合した表 + 基準のシート + <ファイル名>_<シート名>。両エンジンで同じシート名、元のシートと同じ内容', async ({ page }) => {
    test.setTimeout(30 * 60_000);
    const errors = watchErrors(page);
    await openApp(page);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, samplePaths()));
    await waitMergeReady(page, errors);
    await setSheetMode(page, 'all');
    expect(await sheetModeOf(page)).toBe('all');
    // the preview lists the sheets of the output before the merge, in the order: merged, base, then the cards
    await expect(page.getByTestId('preview-output-sheets')).toBeVisible();
    expect(await plannedNames(page), '結合前の一覧').toEqual(ALL_NAMES);
    await mergeNow(page, errors);
    await expect(page.getByTestId('side-output-sheets')).toHaveText('出力は 9 シート（結合 1 + 基準のほか 3 + 他ファイル 5）');
    expect(await sheetListNames(page), '結合後の一覧は結合前と同じ').toEqual(ALL_NAMES);
    // the two boxes cannot be changed after the merge
    await expect(page.getByTestId('include-other-sheets')).toBeDisabled();
    await expect(page.getByTestId('include-source-sheets')).toBeDisabled();

    const out = await exportBoth(page, errors, 'e2e_r9b_全シート');
    expect(out.mode).toBe('all');
    expect(out.sources?.map((s) => s.name), 'コピー元（カードの順）').toEqual(['02_面接評価.xlsx', '03_連絡先.csv', '04_適性検査.xls', '05_研修出欠.xlsx']);
    test.skip(!excelAvailable(), 'Excel で開いての確認には Excel が必要です');
    const v = await verifyAndProbe(out);
    expectPass(v.verdicts, '全シート');
    for (const names of sheetNamesOfDumps(v, out)) expect(names, 'Excel で開いたシートの名前と順番').toEqual(ALL_NAMES);
    for (const [engine, verdict] of v.verdicts) {
      expect(verdict.items.filter((i) => i.level === 'pass').map((i) => i.code), `${engine}: 名前と中身`).toEqual(expect.arrayContaining(['copy-names', 'copy-equal']));
      expect(verdict.stats.copies, `${engine}: 比べたコピー`).toEqual({ compared: 5 });
    }
    expectSameFormulas(out);
  });

  test('(c) 最初のチェックだけ（他のファイルのシートは外す）: 基準のほかのシートは残り、コピーは入らない。従来どおり --single-sheet なしで PASS', async ({ page }) => {
    test.setTimeout(20 * 60_000);
    const errors = watchErrors(page);
    await openApp(page);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, samplePaths()));
    await waitMergeReady(page, errors);
    await setSheetMode(page, 'base');
    await expect(page.getByTestId('preview-output-sheets'), '他のファイルのシートを入れないときは一覧を出さない').toHaveCount(0);
    await mergeNow(page, errors);
    await expect(page.getByTestId('side-output-sheets')).toHaveCount(0);
    const out = await exportBoth(page, errors, 'e2e_r9c_基準のシートだけ');
    expect(out.mode).toBe('base');
    test.skip(!excelAvailable(), 'Excel で開いての確認には Excel が必要です');
    const v = await verifyAndProbe(out);
    expectPass(v.verdicts, '基準のシートを残す');
    for (const names of sheetNamesOfDumps(v, out)) expect(names).toEqual(BASE_ONLY_NAMES);
  });

  test('(d) 同じファイル名が重なる: あとから入れたほうに (1)。結合前の一覧 = 結合後の一覧 = 出力ファイル', async ({ page }) => {
    test.setTimeout(30 * 60_000);
    const errors = watchErrors(page);
    const a = scratchCopy('a', '02_面接評価.xlsx', 1_760_000_000);
    const b = scratchCopy('b', '02_面接評価.xlsx', 1_760_000_100);
    await openApp(page);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, [path.join(SAMPLES_DIR, SAMPLE_FILES[0]), a, b]));
    await waitMergeReady(page, errors);
    await waitIdle(page);
    await setSheetMode(page, 'all');
    const planned = await plannedNames(page);
    expect(planned.slice(4), '2 つ目の同名ファイルのシートに (1)').toEqual(['02_面接評価_面接評価', '02_面接評価_評価テーブル', '02_面接評価_面接評価(1)', '02_面接評価_評価テーブル(1)']);
    await expect(page.locator('[data-testid^="planned-sheet-suffixed-"]')).toHaveCount(2);
    await mergeNow(page, errors);
    expect(await sheetListNames(page), '結合後の一覧').toEqual(planned);
    const out = await exportBoth(page, errors, 'e2e_r9d_同名', { sources: [{ path: a, name: '02_面接評価.xlsx' }, { path: b, name: '02_面接評価.xlsx' }] });
    test.skip(!excelAvailable(), 'Excel で開いての確認には Excel が必要です');
    const v = await verifyAndProbe(out);
    expectPass(v.verdicts, '同名のファイル');
    for (const names of sheetNamesOfDumps(v, out)) expect(names, 'Excel で開いたシート名').toEqual(planned);
  });

  test('(e) 他のシートを参照する数式は、コピーしたシートを指して数式のまま残る（結合した表の値ではなく、元のファイルの値）', async ({ page }) => {
    test.setTimeout(30 * 60_000);
    const errors = watchErrors(page);
    const g01 = path.join(FIXTURES, 'g01_samples__shuf.xlsx');
    test.skip(!fs.existsSync(g01), 'tests/fixtures/formulaimport がありません');
    await openApp(page);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, [path.join(SAMPLES_DIR, SAMPLE_FILES[0]), g01]));
    await waitMergeReady(page, errors);
    await waitIdle(page);
    await setSheetMode(page, 'all');
    await mergeNow(page, errors);
    const planned = await sheetListNames(page);
    // g01_samples__shuf.xlsx has one sheet (評価): its copy is the one the formulas of the merged sheet read
    expect(planned.slice(4), 'コピーされるシート').toEqual(['g01_samples__shuf_評価']);
    const out = await exportBoth(page, errors, 'e2e_r9e_コピーを参照', { sources: [{ path: g01, name: 'g01_samples__shuf.xlsx' }] });
    test.skip(!excelAvailable(), 'Excel で開いての確認には Excel が必要です');
    const v = await verifyAndProbe(out);
    expectPass(v.verdicts, 'コピーを参照する数式');
    // the merged sheet of every output holds formulas that read the copies (the expected-copies.json of F2: 58 of this kind)
    v.dumps.slice(1, 1 + out.files.length).forEach((d, i) => {
      const merged = d.sheets!.find((s) => s.name === '名簿')!;
      const viaCopy = Object.entries(merged.formulas).filter(([, f]) => /g01_samples__shuf_[^!]*!|'g01_samples__shuf_[^']*'!/.test(f));
      test.info().annotations.push({ type: `コピーを読む数式 ${out.files[i].engine}`, description: `${viaCopy.length} 個 例: ${viaCopy.slice(0, 3).map(([a, f]) => `${a}: ${f}`).join(' / ')}` });
      expect(viaCopy.length, `${out.files[i].engine}: コピーを参照する数式が結合した表に残っている`).toBeGreaterThanOrEqual(20);
      expect(Object.values(merged.errors ?? {}).filter((e) => e === '#REF!' || e === '#NAME?'), `${out.files[i].engine}: #REF! / #NAME? がない`).toEqual([]);
    });
    expectSameFormulas(out);
  });

  test('(f) 結合した表の列を動かす・挿入する・値を入れる → 出力しても、コピーしたシートは元のシートのまま（両エンジン）', async ({ page }) => {
    test.setTimeout(30 * 60_000);
    const errors = watchErrors(page);
    await openApp(page);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, samplePaths([SAMPLE_FILES[0], SAMPLE_FILES[1], SAMPLE_FILES[4]])));
    await waitMergeReady(page, errors);
    await setSheetMode(page, 'all');
    await mergeNow(page, errors);
    const titles = await columnTitles(page);
    // a non-contiguous move (inside the title merge: the app asks to unmerge first), an insert and a typed value
    await moveColumns(page, ['学年', '判定'], 2);
    const dlg = page.getByTestId('unmerge-confirm-dialog');
    if (await dlg.isVisible().catch(() => false)) await page.getByTestId('confirm-ok').click();
    await expect.poll(async () => (await columnTitles(page)).indexOf('判定'), { timeout: 20_000 }).toBeLessThan(5);
    await insertColumn(page, await columnIndex(page, 'メール'), 'right');
    const at = (await columnIndex(page, 'メール')) + 1;
    await typeInCell(page, 4, at, 'メモ');
    await selectColumns(page, ['年齢']);
    await page.keyboard.press('Delete');
    // 集計 (kept: both boxes are ticked) reads the age column: the app asks before deleting it
    const del = page.getByTestId('delete-confirm-dialog');
    await expect(del).toBeVisible();
    await page.getByTestId('delete-convert-button').click();
    await expect.poll(async () => editCount(page), { timeout: 20_000 }).toBeGreaterThanOrEqual(4);
    expect((await columnTitles(page)).length).toBe(titles.length); // +1 inserted, -1 deleted
    const out = await exportBoth(page, errors, 'e2e_r9f_編集後');
    test.skip(!excelAvailable(), 'Excel で開いての確認には Excel が必要です');
    const v = await verifyAndProbe(out);
    expectPass(v.verdicts, '編集後の全シート');
    for (const [engine, verdict] of v.verdicts) expect(verdict.items.map((i) => i.code), `${engine}: コピーは元のシートと同じ`).toContain('copy-equal');
    expectSameFormulas(out);
  });

  test('(g) check-output（コマンド）: --single-sheet / --with-copies --source … と、フラグなしの従来の厳しさ', async ({ page }) => {
    test.setTimeout(30 * 60_000);
    test.skip(!excelAvailable(), 'check-output には Excel が必要です');
    const errors = watchErrors(page);
    await openApp(page);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, samplePaths([SAMPLE_FILES[0], SAMPLE_FILES[1]])));
    await waitMergeReady(page, errors);
    await setSheetMode(page, 'all');
    const { others } = await cardSources(page); // (the cards are gone once the merge is done)
    await mergeNow(page, errors);
    const out = await exportBoth(page, errors, 'e2e_r9g_コマンド');
    const file = out.files[0].file;
    const base = path.join(SAMPLES_DIR, SAMPLE_FILES[0]);
    // 1. the right flags: PASS
    const ok = await checkOutput(file, base, { withCopies: true, sources: others });
    expect(ok.verdict, describeVerdict(ok)).toBe('PASS');
    expect(ok.items.find((i) => i.code === 'copied-sheets')?.message).toContain('02_面接評価_面接評価');
    // 2. without the flags the extra sheets are not looked at and the file is judged like a file of rounds 1-6: still PASS
    const plain = await checkOutput(file, base);
    expect(plain.verdict, describeVerdict(plain)).toBe('PASS');
    // 3. --single-sheet on an output that holds more sheets: FAIL (it must hold only the merged sheet)
    const wrong = await checkOutput(file, base, { singleSheet: true });
    expect(wrong.items.filter((i) => i.level === 'fail').map((i) => i.code)).toContain('not-single-sheet');
    // 4. a wrong source: the copies cannot equal it
    const bad = await checkOutput(file, base, { withCopies: true, sources: [{ path: path.join(SAMPLES_DIR, SAMPLE_FILES[4]), name: SAMPLE_FILES[1] }] });
    expect(bad.verdict).toBe('FAIL');
    expect(bad.items.filter((i) => i.level === 'fail').map((i) => i.code)).toEqual(expect.arrayContaining(['copy-missing']));
  });
});

test.describe('ラウンド9: 設定の連動と結合前の一覧', () => {
  test('2 つのチェックの連動（片方を入れるともう片方も入る）、ツールチップ、結合後は押せない、やり直すと選び直せる', async ({ page }) => {
    const errors = watchErrors(page);
    await openApp(page);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, samplePaths([SAMPLE_FILES[0], SAMPLE_FILES[1]])));
    await waitMergeReady(page, errors);
    const other = page.getByTestId('include-other-sheets');
    const source = page.getByTestId('include-source-sheets');
    await expect(source.locator('xpath=ancestor::label')).toHaveAttribute('title', /ファイル名_シート名/);
    await other.check();
    expect(await sheetModeOf(page)).toBe('all');
    await source.uncheck();
    expect(await sheetModeOf(page)).toBe('base');
    await other.uncheck();
    expect(await sheetModeOf(page)).toBe('single');
    await source.check();
    expect(await sheetModeOf(page), '2 つ目だけを入れると 1 つ目も入る').toBe('all');
    await mergeNow(page, errors);
    await expect(other).toBeDisabled();
    await expect(source).toBeDisabled();
    // 結合をやり直す: the choice survives and can be changed again
    await page.getByTestId('remerge-button').click();
    await expect(page.getByTestId('confirm-ok').or(page.getByTestId('merge-button'))).toBeVisible();
    if (await page.getByTestId('confirm-ok').isVisible()) await page.getByTestId('confirm-ok').click();
    await expect(page.getByTestId('merge-button')).toBeVisible();
    expect(await sheetModeOf(page), '設定は残る').toBe('all');
    await expect(other).toBeEnabled();
    expect(errors.messages.filter((m) => !/favicon|DevTools|Failed to load resource/i.test(m))).toEqual([]);
  });
});

