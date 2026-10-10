/**
 * Round 9 (plan A) in a browser: 「他のファイルのシートも含める」, the planned sheet list of the merge preview, and — once the copy
 * engine (A1: PremergeOptions.includeSourceSheets) has landed — the merge result: 「出力は N シート」, 「シート一覧」, the saved file.
 *
 *   npx playwright test -c tests/ui/playwright.config.ts sheets-option       (builds the UI, starts the REAL server on :5192)
 *   VITE_ALLOW_MOCK=1 npx playwright test -c tests/ui/playwright.config.ts sheets-option
 *       the same, with the ?mock=1 demo data in the build (the "mock" tests below need it; without it they skip themselves)
 *   UI_BASE_URL=http://localhost:5192 npx playwright test -c tests/ui/playwright.config.ts sheets-option   (a server that is already running)
 *
 * Three groups:
 *   real modules  the samples (samples/) through the real reader / keys / premerge and the real local server. Everything that does not
 *                 need the copy engine runs now (checkboxes, preview list = planSheetNames of the real sheet names, card order,
 *                 conflicts, long names, locked after the merge, export hint).
 *   real engine   needs the copy engine: test.fixme until src/core/xlsx/premerge.ts mentions includeSourceSheets (A1 landed).
 *   mock          ?mock=1 demo data: the whole after-merge UI (side line, シート一覧, 警告) with the mock's copies.
 * Only the built-in engine is used (no Excel start). Screenshots: tests/ui/shots/p9-*.png.
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { expect, test, type Locator, type Page } from '@playwright/test';
import JSZip from 'jszip';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, '..', '..');
const SAMPLES = path.join(ROOT, 'samples');
const SHOTS = path.join(HERE, 'shots');
fs.mkdirSync(SHOTS, { recursive: true });
const FILES = ['01_学生名簿.xlsx', '02_面接評価.xlsx', '03_連絡先.csv', '04_適性検査.xls', '05_研修出欠.xlsx'].map((f) => path.join(SAMPLES, f));
const shot = (page: Page, name: string) => page.screenshot({ path: path.join(SHOTS, `p9-${name}.png`) });

/** the copy engine has landed when premerge.ts knows the option (until then buildPremerged ignores includeSourceSheets) */
const ENGINE_LANDED = /includeSourceSheets/.test(fs.readFileSync(path.join(ROOT, 'src/core/xlsx/premerge.ts'), 'utf8'));
const ENGINE_REASON = 'A1 (the sheet copier) has not landed: src/core/xlsx/premerge.ts does not know PremergeOptions.includeSourceSheets yet, so no sheet is copied';

const TIP =
  '2つ目以降のファイルの元のシートを、すべてそのまま入れます。シート名は「ファイル名_シート名」になり、名前が重なるときは(1)(2)…が付きます。元のシートを参照している数式もエラーにならずに動きます';

/** what planSheetNames gives for the five samples (base 01: 名簿 / 集計 / ピボット / マスタ), both options on */
const SAMPLE_PLAN = ['名簿', '集計', 'ピボット', 'マスタ', '02_面接評価_面接評価', '02_面接評価_評価テーブル', '03_連絡先_03_連絡先', '04_適性検査_適性検査結果', '05_研修出欠_出欠'];

// ------------------------------------------------------------------------------------------------ helpers

async function openApp(page: Page, url = '/'): Promise<string[]> {
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(`pageerror: ${e.message}`));
  page.on('console', (m) => {
    if (m.type() === 'error' && !/Failed to load resource|ERR_/.test(m.text())) errors.push(m.text());
  });
  await page.goto(url);
  await page.getByTestId('server-status').waitFor();
  return errors;
}

async function openReal(page: Page): Promise<string[]> {
  const errors = await openApp(page);
  await expect(page.getByTestId('server-status')).toHaveAttribute('data-online', 'true');
  return errors;
}

async function settle(page: Page, timeout = 120_000) {
  await expect(page.locator('[data-testid=busy-overlay], .card-loading')).toHaveCount(0, { timeout });
}

async function addFiles(page: Page, paths: string[]) {
  await page.getByTestId('file-input').setInputFiles(paths);
  await settle(page);
}

async function mergeNow(page: Page) {
  await page.getByTestId('merge-button').click();
  await page.getByTestId('grid').waitFor({ timeout: 120_000 });
  await settle(page);
}

const other = (page: Page) => page.getByTestId('include-other-sheets');
const source = (page: Page) => page.getByTestId('include-source-sheets');
const card = (page: Page, id: string): Locator => page.locator(`[data-file-id="${id}"]`);
const cardIds = (page: Page) => page.$$eval('[data-file-id]', (els) => els.map((e) => e.getAttribute('data-file-id')));

async function state(page: Page): Promise<{ other: boolean; source: boolean }> {
  return { other: await other(page).isChecked(), source: await source(page).isChecked() };
}

/** the rows of 「出力されるシート」 (merge preview) */
async function plannedRows(page: Page) {
  return page.$$eval('li[data-testid^="planned-sheet-"]', (els) =>
    els.map((e) => ({
      kind: e.getAttribute('data-kind') ?? '',
      name: e.getAttribute('data-name') ?? '',
      suffixed: !!e.querySelector('[data-testid^="planned-sheet-suffixed-"]'),
      shortened: !!e.querySelector('[data-testid^="planned-sheet-shortened-"]'),
      text: (e.textContent ?? '').replace(/\s+/g, ' ').trim(),
    })),
  );
}
const plannedNames = async (page: Page) => (await plannedRows(page)).map((r) => r.name);

/** names of the rows of 「シート一覧」 (side panel, after the merge); the list is opened first */
async function sheetListNames(page: Page): Promise<string[]> {
  const d = page.getByTestId('sheet-list-details');
  if (!(await d.evaluate((e) => e.hasAttribute('open')))) await d.locator('summary').click();
  return page.$$eval('[data-testid^="sheet-list-item-"]', (els) => els.map((e) => e.getAttribute('data-name') ?? ''));
}

async function openWarnings(page: Page) {
  const d = page.getByTestId('warning-details');
  if (!(await d.evaluate((e) => e.hasAttribute('open')))) await d.locator('summary').click();
}

/** the sheet names of an .xlsx, in workbook order */
async function sheetNamesOf(file: string): Promise<string[]> {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const wb = await zip.file('xl/workbook.xml')!.async('string');
  return [...wb.matchAll(/<sheet\b[^>]*\bname="([^"]*)"/g)].map((m) => m[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'"));
}

// scratch copies of a sample under other names / folders (same name twice, a very long name) — removed afterwards
const SCRATCH = path.join(os.tmpdir(), 'excel-merger-ui-tests', 'p9-inputs');
const LONG_NAME = 'とても長いファイル名のインターンシップ応募者面接評価一覧2026年度版.xlsx';
function scratchCopy(sub: string, name: string, mtimeSec: number): string {
  const dir = path.join(SCRATCH, sub);
  fs.mkdirSync(dir, { recursive: true });
  const dest = path.join(dir, name);
  fs.copyFileSync(FILES[1], dest);
  fs.utimesSync(dest, mtimeSec, mtimeSec); // another timestamp: the app refuses the very same name + size + date twice
  return dest;
}
test.afterAll(() => {
  fs.rmSync(SCRATCH, { recursive: true, force: true });
});

// ================================================================================================ real modules

test.describe('real modules: the two checkboxes', () => {
  test('他のシートも出力に含める の下に 他のファイルのシートも含める がある。連動のしかた・ツールチップ', async ({ page }) => {
    const errors = await openReal(page);
    await addFiles(page, FILES);
    await expect(other(page)).toBeVisible();
    await expect(source(page)).toBeVisible();
    expect(await state(page)).toEqual({ other: false, source: false });

    // the second one is under the first one and indented
    const a = (await other(page).boundingBox())!;
    const b = (await source(page).boundingBox())!;
    expect(b.y).toBeGreaterThan(a.y + 8);
    expect(b.x).toBeGreaterThan(a.x + 8);
    await expect(page.locator('label', { has: source(page) })).toHaveAttribute('title', TIP);
    await expect(page.locator('label', { has: source(page) })).toContainText('他のファイルのシートも含める');
    await expect(page.locator('label', { has: other(page) })).toContainText('他のシートも出力に含める');

    // 1st on -> 2nd on
    await other(page).check();
    expect(await state(page)).toEqual({ other: true, source: true });
    await shot(page, '01-options-both');
    // the user may switch the 2nd off: the base's sheets stay
    await source(page).uncheck();
    expect(await state(page)).toEqual({ other: true, source: false });
    // 1st off -> 2nd off (already off), then 2nd on while the 1st is off -> the 1st goes on too
    await other(page).uncheck();
    expect(await state(page)).toEqual({ other: false, source: false });
    await source(page).check();
    expect(await state(page)).toEqual({ other: true, source: true });
    // 1st off while both on -> both off
    await other(page).uncheck();
    expect(await state(page)).toEqual({ other: false, source: false });
    // switching the 1st on again brings the 2nd back (default on)
    await other(page).check();
    expect(await state(page)).toEqual({ other: true, source: true });
    expect(errors).toEqual([]);
  });

  test('1280px の幅でも設定バーが横にはみ出さない', async ({ page }) => {
    await page.setViewportSize({ width: 1280, height: 800 });
    const errors = await openReal(page);
    await addFiles(page, FILES);
    await other(page).check();
    const bar = await page.getByTestId('settings-bar').evaluate((e) => ({ sw: e.scrollWidth, cw: e.clientWidth }));
    expect(bar.sw).toBeLessThanOrEqual(bar.cw + 1);
    const doc = await page.evaluate(() => ({ sw: document.documentElement.scrollWidth, cw: document.documentElement.clientWidth }));
    expect(doc.sw).toBeLessThanOrEqual(doc.cw + 1);
    await page.getByTestId('settings-bar').screenshot({ path: path.join(SHOTS, 'p9-02-settings-bar-1280.png') });
    expect(errors).toEqual([]);
  });
});

test.describe('real modules: 出力されるシート (before the merge)', () => {
  test('5 つのサンプル: 結合した表 → 基準のシート → <ファイル名>_<シート名>（カードの順）', async ({ page }) => {
    const errors = await openReal(page);
    await addFiles(page, FILES);
    await expect(page.getByTestId('preview-output-sheets')).toHaveCount(0);
    // only the first one: the old line, no list
    await other(page).check();
    await source(page).uncheck();
    await expect(page.getByTestId('preview-output-sheets')).toHaveCount(0);
    await expect(page.getByTestId('preview-sheets')).toContainText('基準ファイルの他のシート 3 枚');
    await source(page).check();
    await expect(page.getByTestId('preview-output-sheets')).toBeVisible();
    await expect(page.getByTestId('preview-sheets')).toHaveCount(0);

    const rows = await plannedRows(page);
    expect(rows.map((r) => r.name)).toEqual(SAMPLE_PLAN);
    expect(rows.map((r) => r.kind)).toEqual(['merged', 'base', 'base', 'base', 'copy', 'copy', 'copy', 'copy', 'copy']);
    expect(rows[0].text).toContain('結合した表');
    expect(rows[1].text).toContain('基準ファイル');
    expect(rows[4].text).toContain('← 「02_面接評価.xlsx」の「面接評価」');
    expect(rows[6].text).toContain('← 「03_連絡先.csv」の「03_連絡先」');
    expect(rows.some((r) => r.suffixed || r.shortened)).toBe(false);
    await expect(page.getByTestId('preview-output-sheets')).toContainText('出力されるシート');
    await page.getByTestId('preview-output-sheets').scrollIntoViewIfNeeded();
    await shot(page, '03-preview-list');
    expect(errors).toEqual([]);
  });

  test('カードを並べ替えると一覧も入れ替わる（基準のシートは先頭のまま）', async ({ page }) => {
    const errors = await openReal(page);
    await addFiles(page, FILES);
    await source(page).check();
    expect(await plannedNames(page)).toEqual(SAMPLE_PLAN);
    // 02 one step to the right: 03 now comes before 02 (card order = order of the copies)
    await card(page, 's2').focus();
    await page.keyboard.press('Alt+ArrowRight');
    await settle(page);
    await expect.poll(() => plannedNames(page)).toEqual(['名簿', '集計', 'ピボット', 'マスタ', '03_連絡先_03_連絡先', '02_面接評価_面接評価', '02_面接評価_評価テーブル', '04_適性検査_適性検査結果', '05_研修出欠_出欠']);
    // 05 to the front of the others
    await card(page, 's5').focus();
    for (let i = 0; i < 4; i++) {
      await page.keyboard.press('Alt+ArrowLeft');
      // wait for the move (and the re-analysis it starts) before the next key: a card cannot be moved while the app is busy
      await expect.poll(async () => (await cardIds(page)).indexOf('s5')).toBe(3 - i);
      await expect(page.getByTestId('merge-button')).toBeEnabled(); // the key analysis follows the new order and is ready again
      await settle(page);
      await expect(card(page, 's5')).toBeFocused();
    }
    expect(await cardIds(page)).toEqual(['s5', 's1', 's3', 's2', 's4']);
    await expect.poll(async () => (await plannedNames(page)).slice(4)).toEqual(['05_研修出欠_出欠', '03_連絡先_03_連絡先', '02_面接評価_面接評価', '02_面接評価_評価テーブル', '04_適性検査_適性検査結果']);
    expect(errors).toEqual([]);
  });

  test('同じファイル名が重なる: あとから入れるほうに (1)(2)… と「名前が重なったため」。長い名前は 31 文字までに短縮', async ({ page }) => {
    const a = scratchCopy('a', '02_面接評価.xlsx', 1_760_000_000);
    const b = scratchCopy('b', '02_面接評価.xlsx', 1_760_000_100);
    const long = scratchCopy('c', LONG_NAME, 1_760_000_200);
    const errors = await openReal(page);
    await addFiles(page, [FILES[0], a, b, long]);
    await source(page).check();
    const rows = await plannedRows(page);
    const names = rows.map((r) => r.name);
    expect(names.slice(0, 4)).toEqual(['名簿', '集計', 'ピボット', 'マスタ']);
    expect(names.slice(4, 8)).toEqual(['02_面接評価_面接評価', '02_面接評価_評価テーブル', '02_面接評価_面接評価(1)', '02_面接評価_評価テーブル(1)']);
    expect(rows.slice(4, 8).map((r) => r.suffixed)).toEqual([false, false, true, true]);
    for (const r of rows.slice(6, 8)) expect(r.text).toContain('名前が重なったため(1)を付けました');
    // the long file name: shortened to at most 31 characters, flagged, the sheet part is kept
    for (const r of rows.slice(8)) {
      expect(r.name.length).toBeLessThanOrEqual(31);
      expect(r.shortened).toBe(true);
      expect(r.text).toContain('31文字までに短縮');
    }
    expect(rows.slice(8).map((r) => r.name.endsWith('面接評価') || r.name.endsWith('評価テーブル'))).toEqual([true, true]);
    // each name once (Excel compares them case-insensitively)
    expect(new Set(names.map((n) => n.toLowerCase())).size).toBe(names.length);
    await page.getByTestId('preview-output-sheets').scrollIntoViewIfNeeded();
    await shot(page, '04-preview-conflicts');
    expect(errors).toEqual([]);
  });
});

test.describe('real modules: after the merge', () => {
  test('結合後は 2 つのチェックが固定される（元の設定のまま・押せない）。やり直すと戻る', async ({ page }) => {
    const errors = await openReal(page);
    await addFiles(page, FILES);
    await other(page).check();
    await mergeNow(page);
    await expect(other(page)).toBeDisabled();
    await expect(source(page)).toBeDisabled();
    expect(await state(page)).toEqual({ other: true, source: true });
    await expect(page.locator('label', { has: source(page) })).toHaveAttribute('title', /結合をやり直す/);
    await shot(page, '05-merged-locked');
    await page.getByTestId('remerge-button').click();
    await expect(page.getByTestId('merge-button')).toBeVisible();
    // the choice survives, and can be changed again
    expect(await state(page)).toEqual({ other: true, source: true });
    await other(page).uncheck();
    expect(await state(page)).toEqual({ other: false, source: false });
    expect(errors).toEqual([]);
  });

  test('出力ダイアログ: コピーしたシートがあるときだけ、Excel本体ではグラフ・ピボットが確実という説明が出る', async ({ page }) => {
    const errors = await openReal(page);
    await addFiles(page, FILES);
    await other(page).check();
    await mergeNow(page);
    await page.getByTestId('export-button').click();
    const hint = page.getByTestId('export-engine-sheets-hint');
    await expect(hint).toBeVisible();
    await expect(hint).toContainText('グラフ・ピボットテーブルは、この方法がもっとも確実に保たれます');
    await expect(page.getByTestId('export-filename')).toHaveValue(/^01_学生名簿_統合_\d{8}\.xlsx$/);
    await shot(page, '06-export-hint');
    await page.keyboard.press('Escape');
    await expect(page.getByTestId('export-dialog')).toHaveCount(0);
    // without the copies: no hint
    await page.getByTestId('remerge-button').click();
    await source(page).uncheck();
    await mergeNow(page);
    await page.getByTestId('export-button').click();
    await expect(page.getByTestId('export-dialog')).toBeVisible();
    await expect(page.getByTestId('export-engine-sheets-hint')).toHaveCount(0);
    await page.keyboard.press('Escape');
    expect(errors).toEqual([]);
  });
});

// ================================================================================================ real engine (A1)

test.describe('real engine: the merge result with the copied sheets', () => {
  test('出力は N シート（結合 1 + 基準のほか M + 他ファイル K）。シート一覧 = 結合前の一覧。警告にコピーの制限', async ({ page }) => {
    test.fixme(!ENGINE_LANDED, ENGINE_REASON);
    const errors = await openReal(page);
    await addFiles(page, FILES);
    await other(page).check();
    const planned = await plannedNames(page);
    expect(planned).toEqual(SAMPLE_PLAN);
    await mergeNow(page);
    const line = page.getByTestId('side-output-sheets');
    await expect(line).toHaveText('出力は 9 シート（結合 1 + 基準のほか 3 + 他ファイル 5）');
    // the 'only one sheet' note is gone
    await expect(page.getByTestId('side-dropped-sheets')).toHaveCount(0);
    expect(await sheetListNames(page)).toEqual(planned);
    await expect(page.getByTestId('sheet-list-item-0')).toHaveAttribute('data-kind', 'merged');
    await expect(page.getByTestId('sheet-list-item-8')).toHaveAttribute('data-kind', 'copy');
    // the grid shows only the merged sheet
    await expect(page.getByTestId('grid')).toBeVisible();
    await openWarnings(page);
    await expect(page.getByTestId('warning-list')).not.toContainText('他のファイルのシートは、出力に追加されませんでした');
    await shot(page, '07-merged-sheet-list');
    expect(errors).toEqual([]);
  });

  test('名前が重なるファイル: 結合結果の名前 = 結合前の一覧（(1)(2)… 付き）', async ({ page }) => {
    test.fixme(!ENGINE_LANDED, ENGINE_REASON);
    const a = scratchCopy('a', '02_面接評価.xlsx', 1_760_000_000);
    const b = scratchCopy('b', '02_面接評価.xlsx', 1_760_000_100);
    const errors = await openReal(page);
    await addFiles(page, [FILES[0], a, b]);
    await source(page).check();
    const planned = await plannedNames(page);
    expect(planned).toContain('02_面接評価_面接評価(1)');
    await mergeNow(page);
    expect(await sheetListNames(page)).toEqual(planned);
    expect(errors).toEqual([]);
  });

  test('保存した .xlsx のシート名と順番 = 結合前の一覧（内蔵エンジン）', async ({ page }) => {
    test.fixme(!ENGINE_LANDED, ENGINE_REASON);
    const errors = await openReal(page);
    await addFiles(page, FILES);
    await other(page).check();
    const planned = await plannedNames(page);
    await mergeNow(page);
    await page.getByTestId('export-button').click();
    await page.getByTestId('export-engine-js').check();
    await page.getByTestId('export-confirm').click();
    await expect(page.getByTestId('export-result')).toBeVisible({ timeout: 120_000 });
    const saved = (await page.getByTestId('export-result').locator('.result-path').innerText()).trim();
    expect(await sheetNamesOf(saved)).toEqual(planned);
    expect(errors).toEqual([]);
  });
});

// ================================================================================================ mock (?mock=1)

test.describe('mock: the after-merge UI with the demo data', () => {
  async function openMock(page: Page, query = 'mock=1&demo=files'): Promise<string[]> {
    const errors = await openApp(page, `/?${query}`);
    const mocked = await page.getByText('デモデータ（?mock）').isVisible().catch(() => false);
    test.skip(!mocked, 'the UI was built without VITE_ALLOW_MOCK=1 (the ?mock=1 demo data is not part of it): run with VITE_ALLOW_MOCK=1 or against the dev server');
    await settle(page);
    return errors;
  }

  test('結合前の一覧 → 結合 → 出力は 4 シート、シート一覧、トーストの案内', async ({ page }) => {
    const errors = await openMock(page);
    await source(page).check();
    expect(await state(page)).toEqual({ other: true, source: true });
    expect(await plannedNames(page)).toEqual(['応募者一覧', '集計', '職場回答_回答', '人事回答_Sheet1']);
    await page.getByTestId('preview-output-sheets').scrollIntoViewIfNeeded();
    await shot(page, '08-mock-preview-list');
    await mergeNow(page);
    await expect(page.getByTestId('side-output-sheets')).toHaveText('出力は 4 シート（結合 1 + 基準のほか 1 + 他ファイル 2）');
    await expect(page.getByTestId('toast')).toContainText('他のファイルのシート 2 枚を追加');
    expect(await sheetListNames(page)).toEqual(['応募者一覧', '集計', '職場回答_回答', '人事回答_Sheet1']);
    await openWarnings(page);
    await expect(page.getByTestId('warning-list')).toContainText('他のファイルのシート 2 枚をコピーしました');
    // names only: the grid shows the merged sheet
    await expect(page.getByTestId('formula-bar-address')).toBeVisible();
    await shot(page, '09-mock-merged');
    // the export dialog hint
    await page.getByTestId('export-button').click();
    await expect(page.getByTestId('export-engine-sheets-hint')).toBeVisible();
    await page.keyboard.press('Escape');
    expect(errors).toEqual([]);
  });

  test('初期状態（どちらもオフ）の結合: 1 シートのみ。一覧もコピーも出ない', async ({ page }) => {
    const errors = await openMock(page);
    expect(await state(page)).toEqual({ other: false, source: false });
    await expect(page.getByTestId('preview-output-sheets')).toHaveCount(0);
    await expect(page.getByTestId('preview-sheets')).toContainText('出力は結合した 1 シートのみです');
    await mergeNow(page);
    await expect(page.getByTestId('side-output-sheets')).toHaveCount(0);
    await expect(page.getByTestId('side-dropped-sheets')).toContainText('出力は結合した 1 シートのみです');
    await expect(page.getByTestId('sheet-list-details')).toHaveCount(0);
    expect(errors).toEqual([]);
  });

  test('コピーできなかったもの（sheetCopyLimits）は警告に平易な日本語で出て、シート一覧に「一部のみ」が付く', async ({ page }) => {
    // 学生回答 (it has a chart) is the SECOND file: its sheets are copied, and the mock's copier leaves the chart out
    const errors = await openMock(page, 'mock=1');
    const files = ['職場回答.xlsx', '学生回答.xlsx'].map((n) => ({ name: n, mimeType: 'application/octet-stream', buffer: Buffer.alloc(1000) }));
    await page.getByTestId('file-input').setInputFiles(files);
    await settle(page);
    await source(page).check();
    await expect.poll(() => plannedNames(page)).toEqual(['回答', '学生回答_応募者一覧', '学生回答_集計']);
    await mergeNow(page);
    await openWarnings(page);
    const list = page.getByTestId('warning-list');
    await expect(list).toContainText('「学生回答_集計」: グラフ 1 個はコピーできませんでした');
    await expect(list.locator('li.sev-warning', { hasText: 'コピーできませんでした' })).toHaveCount(1);
    expect(await sheetListNames(page)).toEqual(['回答', '学生回答_応募者一覧', '学生回答_集計']);
    await expect(page.locator('[data-testid^="sheet-list-item-"]', { hasText: '一部のみ' })).toHaveCount(1);
    await list.locator('li.sev-warning', { hasText: 'コピーできませんでした' }).scrollIntoViewIfNeeded(); // the entry itself in the picture
    await shot(page, '10-mock-copy-limits');
    expect(errors).toEqual([]);
  });
});
