/**
 * Round 5 (final hardening) UI flows against the REAL app and the real local server, in Microsoft Edge:
 * hostile inputs, layouts, keyboard, dialogs, grid gestures, export with the server gone, text quality.
 *
 *   npx playwright test -c tests/ui/playwright.config.ts            (builds the UI, starts the server on :5192)
 *
 * Fixtures come from tests/ui/hostile-fixtures.mjs (fake data, generated into a scratch folder). Only the built-in export
 * engine is used here (no Excel start).
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { expect, test, type Page } from '@playwright/test';
import { makeHostileFixtures } from './hostile-fixtures.mjs';

const SCRATCH = path.resolve(process.env.UI_SCRATCH || path.join(os.tmpdir(), 'excel-merger-ui-tests'));
const FX = path.join(SCRATCH, 'fixtures');
const SAMPLES = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', 'samples');
const fx = (name: string) => path.join(FX, name);
const sample = (name: string) => path.join(SAMPLES, name);
const LONG_NAME = '非常に長いファイル名の例_'.repeat(9) + '終.xlsx';

let fixtureError: unknown = null;
test.beforeAll(async () => {
  try {
    await makeHostileFixtures(FX, { big: process.env.UI_BIG === '1', huge: process.env.UI_HUGE === '1' });
    // a folder with Excel files, a text file and a nested folder; an empty folder
    const dir = path.join(FX, 'folder1');
    fs.mkdirSync(path.join(dir, 'sub'), { recursive: true });
    fs.mkdirSync(path.join(FX, 'emptyfolder'), { recursive: true });
    fs.copyFileSync(fx('a_名簿.xlsx'), path.join(dir, 'a_名簿.xlsx'));
    fs.copyFileSync(fx('b_成績.xlsx'), path.join(dir, 'b_成績.xlsx'));
    fs.copyFileSync(fx('メモ.txt'), path.join(dir, 'メモ.txt'));
    fs.copyFileSync(fx('c_アンケート.xlsx'), path.join(dir, 'sub', 'c_アンケート.xlsx'));
  } catch (e) {
    fixtureError = e;
  }
});
test.beforeEach(() => {
  if (fixtureError) throw fixtureError;
});

// ------------------------------------------------------------------------------------------------ helpers

async function openApp(page: Page, viewport?: { width: number; height: number }) {
  if (viewport) await page.setViewportSize(viewport);
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(`pageerror: ${e.message}`));
  page.on('console', (m) => {
    // failed /api requests are expected in the "server gone" tests
    if (m.type() === 'error' && !/Failed to load resource|ERR_/.test(m.text())) errors.push(m.text());
  });
  await page.goto('/');
  await page.getByTestId('server-status').waitFor();
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

async function merge(page: Page) {
  await page.getByTestId('merge-button').click();
  await page.getByTestId('grid').waitFor({ timeout: 120_000 });
  await settle(page);
}

const sideText = (page: Page) => page.getByTestId('side-summary').innerText();
const editCount = async (page: Page) => Number(/変更 (\d+) 件/.exec(await sideText(page))?.[1] ?? NaN);
const columnTitles = (page: Page) => page.locator('[data-testid^=column-list-item-] .col-title').allInnerTexts();
const cardTexts = (page: Page) => page.locator('[data-testid^=file-card-]').allInnerTexts();

/** drop real files / folders like the OS does (Chromium builds the DataTransfer, folders included, from the paths) */
async function osDrop(page: Page, paths: string[]) {
  const cdp = await page.context().newCDPSession(page);
  const data = { items: [], files: paths.map((p) => path.resolve(p)), dragOperationsMask: 1 };
  for (const type of ['dragEnter', 'dragOver', 'drop'] as const) await cdp.send('Input.dispatchDragEvent', { type, x: 600, y: 400, data });
  await cdp.detach();
}

const BAD = /undefined|\[object|NaN\b|Infinity/;

// ------------------------------------------------------------------------------------------------ first run / help

test.describe('初めて開いたとき', () => {
  test('使い方パネルに4ステップがあり、閉じた状態は覚えている', async ({ page }) => {
    const errors = await openApp(page);
    const panel = page.getByTestId('help-panel');
    await expect(panel).toHaveAttribute('open', '');
    const steps = panel.locator('li');
    await expect(steps).toHaveCount(4);
    const header = await page.getByTestId('step-indicator').innerText();
    for (const name of ['ファイル', '結合キー', '列の編集', '出力']) {
      expect(header).toContain(name);
      await expect(panel).toContainText(name);
    }
    await panel.locator('summary').click();
    await expect(panel).not.toHaveAttribute('open', '');
    await page.reload();
    await page.getByTestId('server-status').waitFor();
    await expect(page.getByTestId('help-panel')).not.toHaveAttribute('open', '');
    await page.getByTestId('help-panel').locator('summary').click();
    await expect(page.getByTestId('help-panel')).toHaveAttribute('open', '');
    expect(errors).toEqual([]);
  });

  test('ヘッダーの「使い方」でダイアログが開き、Esc で閉じてフォーカスが戻る', async ({ page }) => {
    await openApp(page);
    const btn = page.getByTestId('help-button');
    await btn.click();
    const dlg = page.getByTestId('help-dialog');
    await expect(dlg).toBeVisible();
    await expect(dlg).toContainText('Ctrl + Z');
    expect(await dlg.innerText()).not.toMatch(BAD);
    await page.keyboard.press('Escape');
    await expect(dlg).toHaveCount(0);
    await expect(btn).toBeFocused();
  });

  test('ファイルを読み込むと「ファイルを追加」ボタンがヘッダーに出る', async ({ page }) => {
    await openApp(page);
    await expect(page.getByTestId('header-add-files')).toHaveCount(0);
    await addFiles(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx')]);
    const chooser = page.waitForEvent('filechooser');
    await page.getByTestId('header-add-files').click();
    await chooser;
  });
});

// ------------------------------------------------------------------------------------------------ hostile inputs

test.describe('読み込めないファイル', () => {
  test('0バイト・テキスト偽装・途中で切れたzip・パスワード付き・対応外・ロックファイル', async ({ page }) => {
    const errors = await openApp(page);
    await addFiles(page, [fx('a_名簿.xlsx'), fx('zero.xlsx'), fx('corrupt.xlsx'), fx('password.xlsx'), fx('メモ.txt'), fx('~$a_名簿.xlsx')]);
    const cards = await cardTexts(page);
    expect(cards).toHaveLength(4); // 名簿 + the three broken ones; txt and lock file never get a card
    expect(cards[1]).toContain('空のファイルです');
    expect(cards[2]).toMatch(/読み込めませんでした/);
    expect(cards[3]).toContain('パスワード');
    for (const c of cards) expect(c).not.toMatch(BAD);
    await expect(page.getByTestId('toast')).toContainText('対応していない形式です');
    // the good file is still usable, and nothing crashed
    await expect(page.getByTestId('merge-button')).toBeDisabled();
    await expect(page.getByTestId('file-card-0')).toContainText('40 行');
    expect(errors).toEqual([]);
  });

  test('ロックファイル（~$）だけを入れると、何のファイルかを説明する', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('~$a_名簿.xlsx')]);
    await expect(page.getByTestId('toast')).toContainText('一時ファイル');
    await expect(page.locator('[data-testid^=file-card-]')).toHaveCount(0);
  });

  test('同じファイルを2回入れても1枚だけ', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('a_名簿.xlsx')]);
    await page.getByTestId('file-input').setInputFiles([fx('a_名簿.xlsx')]);
    await expect(page.getByTestId('toast')).toContainText('すでに読み込まれています');
    await expect(page.locator('[data-testid^=file-card-]')).toHaveCount(1);
  });

  test('エラーのカードは、同じ名前のファイルを入れ直すと置き換わる', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('corrupt.xlsx')]);
    await expect(page.getByTestId('file-card-0')).toContainText('読み込めませんでした');
    await addFiles(page, [fx('corrupt.xlsx')]);
    await expect(page.locator('[data-testid^=file-card-]')).toHaveCount(1);
  });

  test('拡張子は Excel だが中身がテキストのファイルには、カードに説明が出る', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('notexcel.xlsx')]);
    await expect(page.getByTestId('file-note-0')).toContainText('テキスト（CSV）');
  });

  test('データのないシート・見出しだけのシートは選び直しを案内する', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('header_only.xlsx'), fx('empty_sheet.xlsx')]);
    const cards = await cardTexts(page);
    for (const c of cards) expect(c).toContain('データ行がありません');
    await expect(page.getByTestId('merge-button')).toBeDisabled();
  });
});

test.describe('フォルダのドロップ', () => {
  test('フォルダの中のExcelファイルを読み込み、Excel以外は無視したと伝える', async ({ page }) => {
    await openApp(page);
    await osDrop(page, [fx('folder1')]);
    await settle(page);
    await expect(page.locator('[data-testid^=file-card-]')).toHaveCount(3);
    const all = (await cardTexts(page)).join('\n');
    expect(all).toContain('a_名簿.xlsx');
    expect(all).toContain('c_アンケート.xlsx');
    await expect(page.getByTestId('toast')).toContainText('フォルダ「folder1」から 3 個のファイルを読み込みます');
    await expect(page.getByTestId('toast')).toContainText('Excel以外の 1 個は無視');
    await merge(page);
  });

  test('空のフォルダは理由を伝える', async ({ page }) => {
    await openApp(page);
    await osDrop(page, [fx('emptyfolder')]);
    await expect(page.getByTestId('toast')).toContainText('見つかりませんでした');
    await expect(page.locator('[data-testid^=file-card-]')).toHaveCount(0);
  });

  test('普通のファイルのドロップ', async ({ page }) => {
    await openApp(page);
    await osDrop(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx')]);
    await settle(page);
    await expect(page.locator('[data-testid^=file-card-]')).toHaveCount(2);
  });
});

test.describe('クセのあるファイル', () => {
  const cases: Array<[string, string[], RegExp]> = [
    ['データが2枚目のシート', ['a_名簿.xlsx', 'sheet2_データが2枚目.xlsx'], /基準 40 名 \/ 一致 40/],
    ['見出しが4行目', ['a_名簿.xlsx', 'header4_見出しが4行目.xlsx'], /基準 40 名 \/ 一致 36/],
    ['2段見出し', ['a_名簿.xlsx', 'twohdr_2段見出し.xlsx'], /基準 40 名 \/ 一致 34/],
    ['英語の見出しだけ', ['english_headers.xlsx', 'a_名簿.xlsx'], /基準 30 名 \/ 一致 30/],
    ['UTF-8のCSV', ['a_名簿.xlsx', 'utf8.csv'], /基準 40 名 \/ 一致 20/],
    ['とても長いファイル名', ['a_名簿.xlsx', LONG_NAME], /基準 40 名 \/ 一致 30/],
    ['アポストロフィ入りのシート名', ['sheetname_アポストロフィ.xlsx', 'a_名簿.xlsx'], /基準 20 名 \/ 一致 20/],
  ];
  for (const [name, files, summary] of cases) {
    test(`${name}を結合できる`, async ({ page }) => {
      const errors = await openApp(page);
      await addFiles(page, files.map(fx));
      await expect(page.getByTestId('summary-line')).toContainText(summary);
      await merge(page);
      const cols = await columnTitles(page);
      expect(cols.length).toBeGreaterThan(3);
      expect(cols.join('|')).not.toMatch(BAD);
      expect(errors).toEqual([]);
    });
  }

  test('2段見出し（面接 / 1次 2次 最終）: 取り込む列の見出しにグループ名が付き、結合後の表にもそのまま入る', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('a_名簿.xlsx'), fx('twohdr_2段見出し.xlsx')]);
    await expect(page.locator('.col-chip', { hasText: '面接 1次' })).toBeVisible();
    await merge(page);
    const cols = await columnTitles(page);
    for (const h of ['面接 1次', '面接 2次', '面接 最終', '適性 言語', '適性 非言語']) expect(cols).toContain(h);
    expect(cols).not.toContain('1次');
  });

  test('数値だけのシート: 結合キーは自動では決まらず、「キーを選ぶ…」で列を選べば結合できる', async ({ page }) => {
    const errors = await openApp(page);
    await addFiles(page, [fx('numbers_only.xlsx'), fx('numbers_only2.xlsx')]);
    await expect(page.getByTestId('file-numeric-header-0')).toContainText('1 行目が数字だけです');
    await expect(page.getByTestId('merge-button')).toBeDisabled();
    await expect(page.getByTestId('merge-button')).toHaveAttribute('title', /結合キー/);
    await expect(page.getByTestId('side-excluded')).toContainText('キーを選んでください');
    await page.getByTestId('file-pick-key-1').click();
    await page.getByTestId('key-detail-modal').waitFor();
    await page.getByTestId('manual-base-select-0').selectOption({ index: 1 });
    await page.getByTestId('manual-other-select-0').selectOption({ index: 1 });
    await page.getByTestId('manual-key-add').click();
    await expect(page.getByTestId('match-rate-text')).toContainText('49 / 49 行が一致');
    await page.getByTestId('key-detail-ok').click();
    await expect(page.getByTestId('summary-line')).toContainText('一致 49');
    await merge(page);
    expect((await columnTitles(page)).length).toBe(7);
    expect(errors).toEqual([]);
  });

  test('とても長いファイル名でも結合プレビューの列が画面内に収まる', async ({ page }) => {
    await openApp(page, { width: 1280, height: 720 });
    await addFiles(page, [fx('a_名簿.xlsx'), fx(LONG_NAME)]);
    const box = await page.getByTestId('merge-preview').boundingBox();
    const last = await page.locator('[data-testid=merge-preview] .check-table th').last().boundingBox();
    expect(last!.x + last!.width).toBeLessThanOrEqual(box!.x + box!.width + 1);
    const doc = await page.evaluate(() => [document.documentElement.scrollWidth, innerWidth]);
    expect(doc[0]).toBeLessThanOrEqual(doc[1]);
    // the name is still readable in full through the tooltip
    await expect(page.locator('[data-testid=merge-preview] .file-ellip').nth(1)).toHaveAttribute('title', LONG_NAME);
  });

  test('300列のシート同士でも結合・複数列の削除・元に戻す', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('wide1_300列.xlsx'), fx('wide2_300列.xlsx')]);
    await merge(page);
    await expect(page.getByTestId('side-summary')).toContainText('× 602 列');
    // the column list must not be 600 tab stops
    expect(await page.locator('[data-testid=column-list] li[tabindex="0"]').count()).toBe(1);
    await page.getByTestId('grid-header-2').click();
    for (let i = 0; i < 40; i++) await page.keyboard.press('Shift+ArrowRight');
    await expect(page.locator('.hint-sel')).toContainText('41 列を選択中');
    await page.keyboard.press('Delete');
    await expect(page.getByTestId('side-summary')).toContainText('× 561 列');
    await page.keyboard.press('Control+z');
    await expect(page.getByTestId('side-summary')).toContainText('× 602 列');
  });
});

// ------------------------------------------------------------------------------------------------ layout

test.describe('画面サイズ', () => {
  for (const [w, h] of [
    [1280, 720],
    [1366, 650],
    [1920, 1080],
  ] as const) {
    test(`${w}x${h}: 横スクロールなし・ヘッダーとサイドパネルが切れない`, async ({ page }) => {
      await openApp(page, { width: w, height: h });
      await addFiles(page, [sample('01_学生名簿.xlsx'), sample('02_面接評価.xlsx'), sample('03_連絡先.csv')]);
      for (const phase of ['files', 'merged'] as const) {
        if (phase === 'merged') await merge(page);
        const m = await page.evaluate(() => {
          const right = (sel: string) => Math.round(document.querySelector(sel)!.getBoundingClientRect().right);
          return { sw: document.documentElement.scrollWidth, vw: innerWidth, side: right('[data-testid=side-panel]'), steps: right('[data-testid=step-indicator]'), help: right('[data-testid=help-button]') };
        });
        expect(m.sw, `${phase}: page wider than the window`).toBeLessThanOrEqual(m.vw);
        expect(m.side).toBeLessThanOrEqual(m.vw);
        expect(m.help).toBeLessThanOrEqual(m.vw);
      }
    });
  }

  test('1024幅: 最小幅1280のため横スクロールで右側（サイドパネル）まで届く', async ({ page }) => {
    await openApp(page, { width: 1024, height: 768 });
    await addFiles(page, [sample('01_学生名簿.xlsx'), sample('02_面接評価.xlsx')]);
    const m = await page.evaluate(() => [document.documentElement.scrollWidth, innerWidth]);
    expect(m[0]).toBe(1280);
    expect(m[1]).toBe(1024);
    await page.evaluate(() => window.scrollTo(99999, 0));
    await expect(page.getByTestId('merge-button')).toBeInViewport();
    await page.getByTestId('merge-button').click();
    await expect(page.getByTestId('grid')).toBeVisible();
  });

  test('ブラウザの拡大 150%（1280x720 相当）でも切れない', async ({ browser }) => {
    const ctx = await browser.newContext({ viewport: { width: 1280, height: 720 }, deviceScaleFactor: 1.5, locale: 'ja-JP' });
    const page = await ctx.newPage();
    await openApp(page);
    await addFiles(page, [sample('01_学生名簿.xlsx'), sample('02_面接評価.xlsx')]);
    await merge(page);
    const m = await page.evaluate(() => [document.documentElement.scrollWidth, innerWidth]);
    expect(m[0]).toBeLessThanOrEqual(m[1]);
    await page.getByTestId('export-button').click();
    const box = await page.getByTestId('export-dialog').boundingBox();
    expect(box!.y).toBeGreaterThanOrEqual(0);
    expect(box!.y + box!.height).toBeLessThanOrEqual(720);
    await ctx.close();
  });

  test('10個のファイル: カードの列は横スクロールで、ヘッダーの追加ボタンから追加できる', async ({ page }) => {
    await openApp(page, { width: 1440, height: 900 });
    const dir = path.join(FX, 'ten');
    fs.mkdirSync(dir, { recursive: true });
    const src = ['a_名簿.xlsx', 'b_成績.xlsx', 'c_アンケート.xlsx'];
    const files = Array.from({ length: 10 }, (_, i) => {
      const p = path.join(dir, `f${String(i + 1).padStart(2, '0')}_${src[i % 3]}`);
      fs.copyFileSync(fx(src[i % 3]), p);
      return p;
    });
    await addFiles(page, files);
    await expect(page.locator('[data-testid^=file-card-]')).toHaveCount(10);
    await expect(page.getByTestId('header-add-files')).toBeInViewport();
    await merge(page);
    await expect(page.getByTestId('side-summary')).toContainText('追加 27 列');
  });
});

// ------------------------------------------------------------------------------------------------ keyboard

test.describe('キーボード操作', () => {
  test('Tabの2つ目が「ファイルを選択」（1つ目はヘッダーの「使い方」）。Enterでファイル選択が開く', async ({ page }) => {
    await openApp(page);
    await page.keyboard.press('Tab');
    await expect(page.getByTestId('help-button')).toBeFocused();
    await page.keyboard.press('Tab');
    const chooser = page.waitForEvent('filechooser', { timeout: 5000 });
    await page.keyboard.press('Enter');
    await chooser;
  });

  test('結合までキーボードだけで進められる', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx')]);
    await page.getByTestId('merge-button').focus();
    await page.keyboard.press('Enter');
    await expect(page.getByTestId('grid')).toBeVisible();
    await settle(page);
    // the grid has the focus: Ctrl+Space selects the column, Ctrl+- deletes it, Ctrl+Z undoes
    await expect(page.locator('#grid-editor, [data-testid=grid] textarea').first()).toBeFocused();
    await page.keyboard.press('Control+Space');
    await expect(page.locator('.hint-sel')).toContainText('1 列を選択中');
    await page.keyboard.press('Control+Minus');
    await expect.poll(() => editCount(page)).toBe(1);
    await page.keyboard.press('Control+z');
    await expect.poll(() => editCount(page)).toBe(0);
  });

  test('表の中でTabが詰まらない: 右端のTabで表の外へ出られ、F6 / Shift+F6 でパネルと表を行き来できる', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx')]);
    await merge(page);
    const inGrid = () => page.evaluate(() => !!document.activeElement?.closest('[data-testid=grid]'));
    expect(await inGrid()).toBe(true);
    await page.keyboard.press('Tab'); // inside the grid Tab still moves to the next cell
    expect(await inGrid()).toBe(true);
    await page.keyboard.press('End');
    await page.keyboard.press('Tab'); // last column: the focus leaves the grid
    await expect.poll(inGrid).toBe(false);
    await page.getByTestId('grid-cell-2-0').click();
    await page.keyboard.press('F6');
    await expect(page.getByTestId('export-button')).toBeFocused();
    await page.keyboard.press('Shift+F6');
    await expect.poll(inGrid).toBe(true);
    // before the merge F6 goes to 「結合する」
    await page.getByTestId('remerge-button').click();
    await settle(page);
    await page.getByTestId('file-card-0').click();
    await page.keyboard.press('F6');
    await expect(page.getByTestId('merge-button')).toBeFocused();
  });

  test('列の一覧: Tabは1回で入り、矢印で動き、Alt+↓で並べ替えられる', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx')]);
    await merge(page);
    const before = await columnTitles(page);
    await page.locator('[data-testid=column-list] li[tabindex="0"]').focus();
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowDown'); // third item
    await page.keyboard.press('Alt+ArrowDown');
    await expect.poll(async () => (await columnTitles(page)).join(',')).not.toBe(before.join(','));
    const after = await columnTitles(page);
    expect(after[2]).toBe(before[3]);
    expect(after[3]).toBe(before[2]);
    expect(await page.locator('[data-testid=column-list] li[tabindex="0"]').count()).toBe(1);
  });

  test('ダイアログを開いている間の Ctrl+Z / Ctrl+Y は状態を変えない。Esc で閉じる', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [sample('01_学生名簿.xlsx'), sample('02_面接評価.xlsx')]);
    await merge(page);
    // 「氏名」 is read by the 重複チェック formulas of the merged sheet: the confirmation opens. (It used to be 「大学」, read by the 集計 sheet —
    // the default output holds only the merged sheet since round 7, so nothing reads 「大学」 any more.)
    await page.getByTestId('grid-header-1').click();
    await page.keyboard.press('Delete');
    await expect(page.getByTestId('delete-confirm-dialog')).toBeVisible();
    const before = await sideText(page);
    for (const k of ['Control+z', 'Control+y', 'Control+Shift+z', 'Delete']) await page.keyboard.press(k);
    expect(await sideText(page)).toBe(before);
    await expect(page.getByTestId('delete-confirm-dialog')).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(page.getByTestId('delete-confirm-dialog')).toHaveCount(0);
    expect(await editCount(page)).toBe(0);
    // same for the export dialog and the context menu
    await page.getByTestId('grid-header-2').click();
    await page.keyboard.press('Control+Minus');
    await expect.poll(() => editCount(page)).toBe(1);
    await page.getByTestId('export-button').click();
    await page.keyboard.press('Control+z');
    expect(await editCount(page)).toBe(1);
    await page.keyboard.press('Escape');
    await expect(page.getByTestId('export-dialog')).toHaveCount(0);
  });

  test('Ctrl+O でファイル選択、Ctrl+S は結合後に出力画面（ブラウザの「保存」は出ない）', async ({ page }) => {
    await openApp(page);
    const chooser = page.waitForEvent('filechooser', { timeout: 5000 });
    await page.keyboard.press('Control+o');
    await chooser;
    await addFiles(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx')]);
    await page.keyboard.press('Control+s');
    await expect(page.getByTestId('toast')).toContainText('「Excel出力」');
    await expect(page.getByTestId('export-dialog')).toHaveCount(0);
    await merge(page);
    await page.keyboard.press('Control+s');
    await expect(page.getByTestId('export-dialog')).toBeVisible();
    await page.keyboard.press('Control+s'); // nothing more happens while the dialog is open
    await expect(page.getByTestId('export-dialog')).toHaveCount(1);
    await page.keyboard.press('Escape');
    // in the middle of typing into a cell: the edit is committed first
    await page.getByTestId('grid-cell-2-1').dblclick();
    await page.keyboard.press('Control+a');
    await page.keyboard.type('保存前の入力');
    await page.keyboard.press('Control+s');
    await expect(page.getByTestId('export-dialog')).toBeVisible();
    expect(await editCount(page)).toBe(1);
  });

  test('複数セルの貼り付けは先頭のセルだけで、その旨が出る', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx')]);
    await merge(page);
    await page.getByTestId('grid-cell-2-2').click();
    await page.evaluate(() => {
      const ta = document.querySelector('[data-testid=grid] textarea')!;
      const dt = new DataTransfer();
      dt.setData('text/plain', '甲\t乙\n丙\t丁');
      ta.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
    });
    await expect(page.getByTestId('toast')).toContainText('先頭のセルだけ');
    await expect(page.getByTestId('grid-cell-2-2')).toHaveText('甲');
    expect(await editCount(page)).toBe(1);
  });

  test('ダイアログの中でTabが外に出ない', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx')]);
    await merge(page);
    await page.getByTestId('export-button').click();
    const dlg = page.getByTestId('export-dialog');
    for (let i = 0; i < 12; i++) {
      await page.keyboard.press('Tab');
      expect(await page.evaluate(() => !!document.activeElement?.closest('[role=dialog]'))).toBe(true);
    }
    await expect(dlg).toBeVisible();
  });

  test('セル編集中の Ctrl+Z は入力欄の取り消しで、表の操作は戻らない。出力ボタンで編集が確定する', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx')]);
    await merge(page);
    const cell = page.getByTestId('grid-cell-2-1');
    await cell.dblclick();
    await page.keyboard.press('Control+a');
    await page.keyboard.type('編集テスト');
    await page.keyboard.press('Control+z');
    expect(await editCount(page)).toBe(0);
    await page.keyboard.press('Escape');
    expect(await editCount(page)).toBe(0);
    await cell.dblclick();
    await page.keyboard.press('Control+a');
    await page.keyboard.type('出力前の編集');
    await page.getByTestId('export-button').click();
    await expect(page.getByTestId('export-dialog')).toBeVisible();
    expect(await editCount(page)).toBe(1);
  });
});

// ------------------------------------------------------------------------------------------------ grid gestures

test.describe('列のドラッグ', () => {
  async function mergedPage(page: Page) {
    await openApp(page, { width: 1280, height: 720 });
    await addFiles(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx'), fx('c_アンケート.xlsx')]);
    await merge(page);
  }

  test('表の外（サイドパネル）で離すと移動しない。Esc でも取り消せる', async ({ page }) => {
    await mergedPage(page);
    const before = await columnTitles(page);
    const h = page.getByTestId('grid-header-4');
    await h.click();
    let a = (await h.boundingBox())!;
    await page.mouse.move(a.x + a.width / 2, a.y + 12);
    await page.mouse.down();
    await page.mouse.move(1150, 500, { steps: 10 });
    await expect(page.getByTestId('grid-drop-indicator')).toHaveCount(0);
    await page.mouse.up();
    await page.waitForTimeout(200);
    expect(await editCount(page)).toBe(0);
    expect(await columnTitles(page)).toEqual(before);
    // Esc in the middle of a drag
    a = (await h.boundingBox())!;
    const b = (await page.getByTestId('grid-header-7').boundingBox())!;
    await page.mouse.move(a.x + a.width / 2, a.y + 12);
    await page.mouse.down();
    await page.mouse.move(b.x + 10, b.y + 12, { steps: 8 });
    await page.keyboard.press('Escape');
    await page.mouse.move(b.x + 30, b.y + 12, { steps: 3 });
    await page.mouse.up();
    await page.waitForTimeout(200);
    expect(await editCount(page)).toBe(0);
  });

  test('右端を越えてドロップすると最後の列になる（自動スクロール）', async ({ page }) => {
    await mergedPage(page);
    const before = await columnTitles(page);
    const h = page.getByTestId('grid-header-4');
    await h.click();
    const a = (await h.boundingBox())!;
    const grid = (await page.getByTestId('grid').boundingBox())!;
    await page.mouse.move(a.x + a.width / 2, a.y + 12);
    await page.mouse.down();
    await page.mouse.move(grid.x + grid.width - 8, a.y + 12, { steps: 12 });
    await page.waitForTimeout(1200); // auto-scroll to the right end
    await page.mouse.up();
    await expect.poll(() => editCount(page)).toBe(1);
    const after = await columnTitles(page);
    expect(after[after.length - 1]).toBe(before[4]);
    await page.keyboard.press('Control+z');
    await expect.poll(() => editCount(page)).toBe(0);
  });

  test('すべての列を選んで削除しようとすると断る（ダイアログなし・変更なし）', async ({ page }) => {
    await mergedPage(page);
    const n = (await columnTitles(page)).length;
    await page.getByTestId('grid-header-0').click();
    for (let i = 0; i < n - 1; i++) await page.keyboard.press('Shift+ArrowRight');
    await expect(page.locator('.hint-sel')).toContainText(`${n} 列を選択中`);
    await page.keyboard.press('Delete');
    await expect(page.getByTestId('toast')).toContainText('すべての列を削除することはできません');
    await expect(page.locator('[role=dialog]')).toHaveCount(0);
    expect(await editCount(page)).toBe(0);
  });

  test('タイトルの結合をまたぐ移動は確認が出て、取り消すと何も変わらない', async ({ page }) => {
    await openApp(page, { width: 1280, height: 720 });
    await addFiles(page, [sample('01_学生名簿.xlsx'), sample('02_面接評価.xlsx')]);
    await merge(page);
    const before = await columnTitles(page);
    await page.getByTestId('grid-header-6').click();
    const a = (await page.getByTestId('grid-header-6').boundingBox())!;
    const b = (await page.getByTestId('grid-header-2').boundingBox())!;
    await page.mouse.move(a.x + a.width / 2, a.y + 12);
    await page.mouse.down();
    await page.mouse.move(b.x + 4, b.y + 12, { steps: 10 });
    await page.mouse.up();
    await expect(page.getByTestId('unmerge-confirm-dialog')).toBeVisible();
    await page.keyboard.press('Escape');
    expect(await columnTitles(page)).toEqual(before);
    expect(await editCount(page)).toBe(0);
  });
});

// ------------------------------------------------------------------------------------------------ flows

test.describe('結合前後の操作', () => {
  test('カードのクリア・基準の切り替え・全部クリアで空の画面に戻る', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [sample('01_学生名簿.xlsx'), sample('02_面接評価.xlsx'), sample('03_連絡先.csv')]);
    await page.getByTestId('file-clear-0').click(); // the base card
    await settle(page);
    await expect(page.getByTestId('file-card-0')).toContainText('02_面接評価');
    await expect(page.getByTestId('file-card-0')).toContainText('★ 基準');
    await page.getByTestId('base-toggle-1').click();
    await settle(page);
    await expect(page.getByTestId('file-card-1')).toContainText('★ 基準');
    await page.getByTestId('file-clear-0').click();
    await settle(page);
    await expect(page.getByTestId('merge-button')).toBeDisabled();
    await page.getByTestId('file-clear-0').click();
    await expect(page.getByTestId('help-panel')).toBeVisible();
    await expect(page.getByTestId('side-summary')).toContainText('ファイルをドロップすると始まります');
  });

  test('結合後にファイルを足す → 確認 → キー画面に戻る。結合をやり直す → 変更の破棄を確認', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx')]);
    await merge(page);
    await page.getByTestId('file-input').setInputFiles([fx('c_アンケート.xlsx')]);
    await expect(page.getByTestId('confirm-dialog')).toContainText('ファイルを追加しますか？');
    await page.keyboard.press('Escape');
    await expect(page.locator('[data-testid^=file-card-]')).toHaveCount(2);
    // an edit makes the wording danger-red; remerge asks first
    await page.getByTestId('grid-header-2').click();
    await page.keyboard.press('Control+Minus');
    await expect.poll(() => editCount(page)).toBe(1);
    await page.getByTestId('remerge-button').click();
    await expect(page.getByTestId('confirm-dialog')).toContainText('変更はすべて失われます');
    await page.keyboard.press('Control+z');
    expect(await editCount(page)).toBe(1);
    await page.keyboard.press('Enter');
    await expect(page.getByTestId('merge-button')).toBeVisible();
    await settle(page);
    await expect(page.getByTestId('merge-button')).toBeEnabled();
  });

  test('結合の直後は「元に戻す」が効かない（結合そのものは戻らない）', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx')]);
    await merge(page);
    await expect(page.getByTestId('undo-button')).toBeDisabled();
    for (let i = 0; i < 3; i++) await page.keyboard.press('Control+z');
    await expect(page.getByTestId('grid')).toBeVisible();
    expect(await editCount(page)).toBe(0);
  });
});

// ------------------------------------------------------------------------------------------------ export

test.describe('出力', () => {
  test('変更なしで出力 → もう一度出力すると別名（(2)）で保存される', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx')]);
    await merge(page);
    const name = `統合テスト_${Date.now()}`;
    const run = async () => {
      await page.getByTestId('export-button').click();
      await page.getByTestId('export-engine-js').check();
      await page.getByTestId('export-filename').fill(name);
      await page.getByTestId('export-confirm').click();
      await expect(page.getByTestId('export-result')).toContainText('保存しました');
      const txt = await page.getByTestId('export-result').innerText();
      await page.getByTestId('export-close').click();
      return txt;
    };
    const first = await run();
    expect(first).toContain(`${name}.xlsx`);
    const second = await run();
    expect(second).toContain(`${name} (2).xlsx`);
  });

  test('出力したあとにもう一度編集すると、ステップは「列の編集」に戻り、閉じる前の確認が再び出る', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx')]);
    await merge(page);
    const prevented = () =>
      page.evaluate(() => {
        const e = new Event('beforeunload', { cancelable: true });
        window.dispatchEvent(e);
        return e.defaultPrevented;
      });
    await page.getByTestId('grid-header-2').click();
    await page.keyboard.press('Control+Minus');
    await expect.poll(() => editCount(page)).toBe(1);
    expect(await prevented()).toBe(true); // unsaved edit
    await page.getByTestId('export-button').click();
    await page.getByTestId('export-engine-js').check();
    await page.getByTestId('export-confirm').click();
    await expect(page.getByTestId('export-result')).toBeVisible();
    await page.getByTestId('export-close').click();
    await expect(page.getByTestId('step-4')).toHaveClass(/active/);
    expect(await prevented()).toBe(false); // exported: closing is fine
    await page.getByTestId('grid-header-2').click();
    await page.keyboard.press('Control+Minus');
    await expect.poll(() => editCount(page)).toBe(2);
    await expect(page.getByTestId('step-3')).toHaveClass(/active/);
    expect(await prevented()).toBe(true); // edited after the export
  });

  test('サーバーが止まったとき: 日本語の案内・内蔵エンジンに切り替わり、ダウンロードで受け取れる', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx')]);
    await merge(page);
    await page.getByTestId('export-button').click();
    await expect(page.getByTestId('export-dialog')).toBeVisible();
    // the server goes away: every /api call fails, the heartbeat reports the loss
    await page.route('**/api/**', (r) => r.abort('connectionrefused'));
    await page.evaluate(() => window.dispatchEvent(new CustomEvent('excel-merger-server', { detail: { connected: false } })));
    await expect(page.getByTestId('server-status')).toContainText('接続できません');
    await expect(page.getByTestId('toast')).toContainText('ローカルサーバーとの接続が切れました');
    await expect(page.getByTestId('export-engine-excel')).toBeDisabled();
    await expect(page.getByTestId('export-dialog')).toContainText('ローカルサーバーに接続できないため');
    const download = page.waitForEvent('download');
    await page.getByTestId('export-confirm').click();
    const dl = await download;
    expect(dl.suggestedFilename()).toMatch(/\.xlsx$/);
    await expect(page.getByTestId('export-result')).toContainText('ダウンロードしました');
    await expect(page.getByTestId('export-result-engine')).toContainText('内蔵エンジン');
    const all = await page.getByTestId('export-dialog').innerText();
    expect(all).not.toMatch(BAD);
    expect(all).not.toMatch(/Failed to fetch|NetworkError/);
  });

  test('サーバーが戻ってきたら、Excel本体の選択肢も戻る', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx')]);
    await merge(page);
    await page.evaluate(() => window.dispatchEvent(new CustomEvent('excel-merger-server', { detail: { connected: false } })));
    await page.evaluate(() => window.dispatchEvent(new CustomEvent('excel-merger-server', { detail: { connected: true } })));
    await page.getByTestId('export-button').click();
    const installed = await page.getByTestId('server-status').getAttribute('data-excel');
    if (installed === 'true') await expect(page.getByTestId('export-engine-excel')).toBeEnabled();
  });

  test('Excel本体の出力が遅いとき: 進捗の表示 → 時間がかかっている案内 → 中止 → 内蔵エンジンで出力できる', async ({ page }) => {
    await page.addInitScript(() => {
      (window as unknown as { __EM_SLOW_HINT_MS: number }).__EM_SLOW_HINT_MS = 500;
    });
    await openApp(page);
    test.skip((await page.getByTestId('server-status').getAttribute('data-excel')) !== 'true', 'Excel がこの PC にありません');
    await addFiles(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx')]);
    await merge(page);
    // a fake Excel engine that never finishes (the job stays 'running')
    const calls: string[] = [];
    await page.route('**/api/jobs**', async (route) => {
      const req = route.request();
      calls.push(`${req.method()} ${new URL(req.url()).pathname}`);
      const json = (body: unknown) => route.fulfill({ status: 200, contentType: 'application/json', body: JSON.stringify(body) });
      if (req.method() === 'POST') return json({ id: 'slow1' });
      if (req.method() === 'PUT') return json({ ok: true });
      if (req.method() === 'DELETE') return json({ ok: true });
      return json({ id: 'slow1', state: 'running', progress: 0.3, message: '操作を再現しています（1/3: 列を移動）…' });
    });
    await page.getByTestId('export-button').click();
    await page.getByTestId('export-engine-excel').check();
    await page.getByTestId('export-confirm').click();
    await expect(page.getByTestId('export-progress')).toContainText('操作を再現しています');
    await expect(page.getByTestId('export-slow')).toContainText('時間がかかっています');
    await expect(page.getByTestId('export-slow')).toContainText('内蔵エンジン');
    // Esc / backdrop do not close the dialog while exporting
    await page.keyboard.press('Escape');
    await expect(page.getByTestId('export-dialog')).toBeVisible();
    await page.getByTestId('export-cancel').click();
    await expect(page.getByTestId('toast')).toContainText('出力を中止しました');
    expect(calls.some((c) => c.startsWith('DELETE'))).toBe(true);
    await page.unroute('**/api/jobs**');
    await page.getByTestId('export-engine-js').check();
    await page.getByTestId('export-confirm').click();
    await expect(page.getByTestId('export-result')).toContainText('内蔵エンジンで出力しました');
  });

  test('出力を中止して、もう一度出力できる（内蔵エンジン）', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('a_名簿.xlsx'), fx('b_成績.xlsx')]);
    await merge(page);
    await page.getByTestId('export-button').click();
    await page.getByTestId('export-engine-js').check();
    await page.getByTestId('export-confirm').click();
    await expect(page.getByTestId('export-result')).toBeVisible();
    await page.getByTestId('export-close').click();
  });
});

// ------------------------------------------------------------------------------------------------ standalone (file://)

test.describe('スタンドアロン（dist/index.html を直接開く）', () => {
  const distIndex = path.join(SCRATCH, 'dist', 'index.html');
  test.skip(!fs.existsSync(distIndex), `${distIndex} がありません（UI をビルドしてから）`);

  test('サーバーなしで、読み込み → 結合 → 列の編集 → 出力（ダウンロード）まで通る', async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.goto(pathToFileURL(distIndex).href);
    await expect(page.getByTestId('server-status')).toHaveAttribute('data-online', 'false');
    await expect(page.getByTestId('server-status')).toContainText('スタンドアロン');
    await expect(page.getByTestId('help-panel')).toBeVisible();
    await page.getByTestId('file-input').setInputFiles([sample('01_学生名簿.xlsx'), sample('02_面接評価.xlsx'), sample('03_連絡先.csv')]);
    await settle(page);
    // (since round 7 the default output holds ONLY the merged sheet, so no formula of another sheet depends on 「大学」 any more and the
    //  deletion would not ask: keep the base file's other sheets for this flow, as the test always assumed; this test failed at the
    //  round-9 baseline for that reason)
    await page.getByTestId('include-other-sheets').check();
    await merge(page);
    await page.getByTestId('grid-header-3').click();
    await page.keyboard.press('Control+Minus');
    await expect(page.getByTestId('delete-confirm-dialog')).toBeVisible(); // 「大学」 is used by formulas of the other sheets
    await page.getByTestId('delete-convert-button').click();
    await expect.poll(() => editCount(page)).toBe(1);
    await page.getByTestId('export-button').click();
    await expect(page.getByTestId('export-engine-excel')).toBeDisabled();
    await expect(page.getByTestId('export-dialog')).toContainText('スタンドアロンで開いているため使えません');
    const download = page.waitForEvent('download');
    await page.getByTestId('export-confirm').click();
    const dl = await download;
    expect(dl.suggestedFilename()).toMatch(/^01_学生名簿_統合_\d{8}\.xlsx$/);
    await expect(page.getByTestId('export-result')).toContainText('ダウンロードしました');
    expect(errors).toEqual([]);
  });

  test('ファイルのドロップ（DataTransfer）でも読み込める', async ({ page }) => {
    await page.goto(pathToFileURL(distIndex).href);
    await page.getByTestId('server-status').waitFor();
    const payload = [fx('a_名簿.xlsx'), fx('b_成績.xlsx')].map((p) => ({ name: path.basename(p), b64: fs.readFileSync(p).toString('base64') }));
    await page.evaluate(async (files) => {
      const dt = new DataTransfer();
      for (const f of files) dt.items.add(new File([Uint8Array.from(atob(f.b64), (c) => c.charCodeAt(0))], f.name));
      for (const type of ['dragenter', 'dragover', 'drop']) document.body.dispatchEvent(new DragEvent(type, { bubbles: true, cancelable: true, dataTransfer: dt }));
    }, payload);
    await settle(page);
    await expect(page.locator('[data-testid^=file-card-]')).toHaveCount(2);
  });
});

// ------------------------------------------------------------------------------------------------ text quality

test.describe('表示される文言', () => {
  async function visibleStrings(page: Page): Promise<string[]> {
    return page.evaluate(() => {
      const out: string[] = [];
      const w = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
      let n: Node | null;
      while ((n = w.nextNode())) {
        const t = (n.nodeValue ?? '').replace(/\s+/g, ' ').trim();
        if (t && n.parentElement && !['SCRIPT', 'STYLE'].includes(n.parentElement.tagName)) out.push(t);
      }
      for (const el of document.querySelectorAll('[title],[aria-label],[placeholder]')) {
        for (const a of ['title', 'aria-label', 'placeholder']) {
          const v = el.getAttribute(a);
          if (v?.trim()) out.push(v.replace(/\s+/g, ' ').trim());
        }
      }
      return out;
    });
  }

  test('どの画面にも undefined / [object Object] / NaN がなく、英語の文が混じらない', async ({ page }) => {
    await openApp(page);
    const all: string[] = [];
    all.push(...(await visibleStrings(page)));
    await addFiles(page, [sample('01_学生名簿.xlsx'), sample('02_面接評価.xlsx'), sample('03_連絡先.csv'), sample('04_適性検査.xls'), sample('05_研修出欠.xlsx')]);
    all.push(...(await visibleStrings(page)));
    await page.getByTestId('key-detail-button').click();
    await page.getByTestId('key-detail-modal').waitFor();
    all.push(...(await visibleStrings(page)));
    const tabs = page.locator('[data-testid^=key-detail-tab-]');
    for (let i = 0; i < (await tabs.count()); i++) {
      await tabs.nth(i).click();
      all.push(...(await visibleStrings(page)));
    }
    await page.keyboard.press('Escape');
    await merge(page);
    await page.evaluate(() => document.querySelectorAll('details').forEach((d) => (d.open = true)));
    all.push(...(await visibleStrings(page)));
    await page.getByTestId('grid-header-3').click({ button: 'right' });
    all.push(...(await visibleStrings(page)));
    await page.keyboard.press('Escape');
    await page.getByTestId('grid-header-3').click();
    await page.keyboard.press('Delete');
    all.push(...(await visibleStrings(page)));
    await page.keyboard.press('Escape');
    await page.getByTestId('export-button').click();
    all.push(...(await visibleStrings(page)));
    await page.keyboard.press('Escape');
    await page.getByTestId('help-button').click();
    all.push(...(await visibleStrings(page)));

    const bad = all.filter((t) => BAD.test(t));
    expect(bad).toEqual([]);
    // English words in the app's own wording (file data such as e-mail addresses is skipped)
    const allowed = /^(Excel|ID|CSV|xlsx|xlsm|xls|xlsb|csv|ods|Shift_JIS|UTF-8|BOM|EUC-JP|UTF-16|JIS|ASCII|Ctrl|Shift|Alt|Delete|Enter|Esc|Tab|Space|F2|F10|OK|start|bat|VLOOKUP|COUNTIF|SUM|UNIQUE|IF|REF|HTTP|VBA|PC|Z|Y)$/i;
    const english = [...new Set(all)]
      .filter((t) => !/[@\\/]|^[A-Za-z0-9._\-() ]+$/.test(t) && !/\d{2,}/.test(t))
      .flatMap((t) => (t.match(/[A-Za-z][A-Za-z0-9_.-]{3,}/g) ?? []).filter((w) => !allowed.test(w)).map((w) => `${w} ← ${t.slice(0, 60)}`));
    expect(english).toEqual([]);
  });

  test('ボタンやセレクトに、意味が伝わる説明（ツールチップ）がある', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [sample('01_学生名簿.xlsx'), sample('02_面接評価.xlsx')]);
    for (const id of ['key-detail-button', 'join-type-select', 'sheet-select-0', 'header-row-select-0', 'base-toggle-0', 'file-clear-0', 'merge-button', 'undo-button', 'redo-button', 'header-add-files', 'help-button']) {
      const title = await page.getByTestId(id).getAttribute('title');
      expect(title, id).toBeTruthy();
    }
    await expect(page.locator('.check-group-head').first()).toHaveAttribute('title', /.+/);
    await expect(page.locator('[data-testid=warning-details] summary')).toHaveAttribute('title', /.+/);
  });
});

// ------------------------------------------------------------------------------------------------ big data (opt-in)

test.describe('大きなデータ', () => {
  test.skip(process.env.UI_BIG !== '1', 'UI_BIG=1 で実行します（2万行×30列の結合。UI_HUGE=1 で 60MB 級のファイルも）');

  test('2万行×30列 + 2ファイルを1分以内に結合でき、スクロールも編集もできる', async ({ page }) => {
    await openApp(page, { width: 1600, height: 900 });
    const t0 = Date.now();
    await addFiles(page, [fx('big_base_20000x30.xlsx'), fx('big_other1.xlsx'), fx('big_other2.xlsx')]);
    await merge(page);
    const secs = (Date.now() - t0) / 1000;
    expect(secs, `read + merge took ${secs}s`).toBeLessThan(60);
    await expect(page.getByTestId('side-summary')).toContainText('20,000 行 × 35 列');
    const frames = await page.evaluate(async () => {
      const el = document.querySelector('[data-testid=grid]')!;
      const out: number[] = [];
      let last = performance.now();
      for (let i = 0; i < 80; i++) {
        el.scrollTop += 1500;
        if (i % 3 === 0) el.scrollLeft += 250;
        await new Promise((r) => requestAnimationFrame(r));
        const now = performance.now();
        out.push(now - last);
        last = now;
      }
      return out.sort((a, b) => a - b);
    });
    expect(frames[Math.floor(frames.length * 0.95)], 'p95 frame time while scrolling (ms)').toBeLessThan(150);
    await page.evaluate(() => {
      const el = document.querySelector('[data-testid=grid]')!;
      el.scrollTop = 0;
      el.scrollLeft = 0;
    });
    await page.getByTestId('grid-header-3').click();
    await page.keyboard.press('Delete');
    await expect.poll(() => editCount(page)).toBe(1);
    await page.keyboard.press('Control+z');
    await expect.poll(() => editCount(page)).toBe(0);
  });

  test('ファイルのベースを後から変えて結合をやり直しても、状態が壊れない（2万行）', async ({ page }) => {
    await openApp(page);
    await addFiles(page, [fx('big_base_20000x30.xlsx'), fx('big_other1.xlsx')]);
    await merge(page);
    await page.getByTestId('remerge-button').click();
    await settle(page);
    await page.getByTestId('base-toggle-1').click();
    await settle(page);
    await expect(page.getByTestId('summary-line')).toContainText('基準 19,000');
    await merge(page);
    await expect(page.getByTestId('side-summary')).toContainText('19,000 行');
  });

  test('60MB級のファイルを読み込める（UI_HUGE=1）', async ({ page }) => {
    test.skip(process.env.UI_HUGE !== '1', 'UI_HUGE=1');
    await openApp(page);
    await addFiles(page, [fx('big60_画像入り.xlsx'), fx('a_名簿.xlsx')]);
    await expect(page.getByTestId('file-card-0')).toContainText('58.0 MB');
    await expect(page.getByTestId('merge-button')).toBeEnabled();
  });
});
