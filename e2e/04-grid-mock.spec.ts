// Grid / editor interactions on the UI's own demo data (?mock=standalone&demo=merged, vite dev server only), so the
// data-testid contract can be exercised before the core modules (premerge / replay) are finished (owner: T).
//   npx vite --config e2e/vite.e2e.config.ts
//   E2E_BASE_URL=http://127.0.0.1:5176 npx playwright test e2e/04-grid-mock.spec.ts
import fs from 'node:fs';
import path from 'node:path';
import { expect, test, type Page } from '@playwright/test';
import { columnTitles, dragHeader, openApp, requireTestIds, revealColumn, ROOT, watchErrors } from './helpers';

const FIRST_DATA_ROW = 3; // the demo base sheet 応募者一覧 has its header on row 2

test.describe.serial('グリッド操作（?mock のデモデータ）', () => {
  let page: Page;
  let ready = false;
  const need = () => test.fixme(!ready, 'デモデータの表を開けなかったため実行できません');

  test.beforeAll(async ({ browser }) => {
    page = await browser.newPage({ acceptDownloads: true, viewport: { width: 1440, height: 900 }, locale: 'ja-JP' });
    watchErrors(page);
  });
  test.afterAll(async () => {
    await page?.close();
  });

  test('デモデータで結合済みの表が開く', async () => {
    test.skip(!process.env.E2E_BASE_URL, '?mock は vite の開発サーバーでだけ使えます（E2E_BASE_URL を指定）');
    await openApp(page, '?mock=standalone&demo=merged&delay=0');
    await expect(page.getByTestId('server-status')).toBeVisible();
    const isMock = await page.getByText('デモデータ（?mock）').isVisible().catch(() => false);
    test.skip(!isMock, 'この画面はデモデータ（?mock）に対応していません（本番ビルド）');
    await requireTestIds(page, ['grid', 'column-list', 'undo-button', 'redo-button'], '結合後の画面', 30_000);
    await expect(page.getByTestId('grid-header-0')).toBeVisible();
    expect((await columnTitles(page)).length).toBeGreaterThan(11);
    ready = true;
  });

  test('クリック / Ctrl+クリック / Ctrl+Shift+クリックで列を選択できる', async () => {
    need();
    await (await revealColumn(page, 3)).click();
    await (await revealColumn(page, 5)).click({ modifiers: ['Control'] });
    await (await revealColumn(page, 7)).click({ modifiers: ['Control', 'Shift'] });
    for (const i of [3, 5, 6, 7]) await expect(page.getByTestId(`grid-header-${i}`)).toHaveAttribute('aria-selected', 'true');
    for (const i of [2, 4, 8]) await expect(page.getByTestId(`grid-header-${i}`)).toHaveAttribute('aria-selected', 'false');
    await (await revealColumn(page, 4)).click();
    await expect(page.getByTestId('grid-header-4')).toHaveAttribute('aria-selected', 'true');
    await expect(page.getByTestId('grid-header-3')).toHaveAttribute('aria-selected', 'false');
  });

  test('離れた2列をドラッグで先頭へ移動し、元に戻す / やり直すができる', async () => {
    need();
    const before = await columnTitles(page);
    await (await revealColumn(page, 4)).click();
    await (await revealColumn(page, 6)).click({ modifiers: ['Control'] });
    await dragHeader(page, 6, 0);
    const expected = [before[4], before[6]];
    await expect.poll(async () => (await columnTitles(page)).slice(0, 2)).toEqual(expected);
    await expect(page.getByTestId('grid-drop-indicator')).toHaveCount(0);

    await page.getByTestId(`grid-cell-${FIRST_DATA_ROW + 1}-0`).click();
    await page.keyboard.press('Control+z');
    await expect.poll(async () => columnTitles(page)).toEqual(before);
    await page.keyboard.press('Control+y');
    await expect.poll(async () => (await columnTitles(page)).slice(0, 2)).toEqual(expected);
    await page.getByTestId('undo-button').click();
    await expect.poll(async () => columnTitles(page)).toEqual(before);
    await page.getByTestId('redo-button').click();
    await expect.poll(async () => (await columnTitles(page)).slice(0, 2)).toEqual(expected);
    await page.getByTestId('undo-button').click();
    await expect.poll(async () => columnTitles(page)).toEqual(before);
  });

  test('右クリックメニューで列を挿入できる', async () => {
    need();
    const before = await columnTitles(page);
    const h = await revealColumn(page, 2);
    await h.click({ button: 'right' });
    await requireTestIds(page, ['context-menu', 'context-menu-item-insert-left'], '列のメニュー');
    await page.getByTestId('context-menu-item-insert-left').click();
    await expect.poll(async () => (await columnTitles(page)).length).toBe(before.length + 1);
    const after = await columnTitles(page);
    expect(after[3]).toBe(before[2]);
    await page.getByTestId('undo-button').click();
    await expect.poll(async () => columnTitles(page)).toEqual(before);
  });

  test('参照されていない列は Delete キーですぐ削除され、元に戻せる', async () => {
    need();
    const before = await columnTitles(page);
    const i = before.length - 1;
    await (await revealColumn(page, i)).click();
    await page.keyboard.press('Delete');
    await expect.poll(async () => (await columnTitles(page)).length).toBe(before.length - 1);
    await expect(page.getByTestId('delete-confirm-dialog')).toHaveCount(0);
    await page.getByTestId('undo-button').click();
    await expect.poll(async () => columnTitles(page)).toEqual(before);
  });

  test('数式が参照している列（筆記: 合計・集計が参照）は確認ダイアログ → キャンセル / 値に変換して削除 → 元に戻す', async () => {
    need();
    const before = await columnTitles(page);
    const i = before.indexOf('筆記');
    expect(i, 'デモデータの「筆記」列').toBeGreaterThanOrEqual(0);
    await (await revealColumn(page, i)).click();
    await page.keyboard.press('Delete');
    await requireTestIds(page, ['delete-confirm-dialog', 'delete-convert-button', 'delete-cancel-button'], '削除の確認');
    const dlg = page.getByTestId('delete-confirm-dialog');
    await expect(dlg).toBeVisible();
    await expect(dlg).toContainText('参照');
    await expect(page.getByTestId('delete-convert-button')).toContainText('値に変換');
    await page.getByTestId('delete-cancel-button').click();
    await expect(dlg).toBeHidden();
    expect(await columnTitles(page)).toEqual(before);
    await (await revealColumn(page, i)).click();
    await page.keyboard.press('Delete');
    await page.getByTestId('delete-convert-button').click();
    await expect.poll(async () => (await columnTitles(page)).length).toBe(before.length - 1);
    expect(await columnTitles(page)).not.toContain('筆記');
    await page.getByTestId('undo-button').click();
    await expect.poll(async () => columnTitles(page)).toEqual(before);
  });

  test('セルに入力して Enter で確定、数式バーに表示、元に戻せる', async () => {
    need();
    const cell = page.getByTestId(`grid-cell-${FIRST_DATA_ROW}-1`);
    await cell.scrollIntoViewIfNeeded();
    const original = (await cell.innerText()).trim();
    await cell.click();
    await page.keyboard.type('テスト 太郎');
    await page.keyboard.press('Enter');
    await expect(cell).toHaveText('テスト 太郎');
    await cell.click();
    await requireTestIds(page, ['formula-bar-input'], '数式バー');
    await expect(page.getByTestId('formula-bar-input')).toHaveValue('テスト 太郎');
    await page.keyboard.press('Control+z');
    await expect(cell).toHaveText(original);
  });

  test('スタンドアロンでは内蔵エンジンで出力してダウンロードする', async () => {
    need();
    await page.getByTestId('export-button').click();
    await requireTestIds(page, ['export-dialog', 'export-engine-js', 'export-engine-excel', 'export-confirm', 'export-filename'], '出力ダイアログ');
    await expect(page.getByTestId('export-engine-excel')).toBeDisabled();
    await expect(page.getByTestId('export-engine-js')).toBeChecked();
    await expect(page.getByTestId('export-filename')).toHaveValue(/_統合_\d{8}\.xlsx$/);
    const download = page.waitForEvent('download');
    await page.getByTestId('export-confirm').click();
    const d = await download;
    const dir = path.join(ROOT, 'test-results', 'e2e-downloads');
    fs.mkdirSync(dir, { recursive: true });
    const file = path.join(dir, `mock-${d.suggestedFilename()}`);
    await d.saveAs(file);
    expect(fs.statSync(file).size).toBeGreaterThan(0);
    await expect(page.getByTestId('export-result')).toContainText('ダウンロードしました');
  });
});
