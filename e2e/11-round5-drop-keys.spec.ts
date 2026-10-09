// Round 5: dropping files ON THE WINDOW (not through the hidden file input) and the Excel habits on the keyboard (owner: T).
// The drops are synthesised DragEvents with a real DataTransfer holding real File objects (what the browser builds for a drop
// of files from Explorer); folders cannot be synthesised (webkitGetAsEntry needs a real OS drag).
import fs from 'node:fs';
import path from 'node:path';
import { expect, test, type Page } from '@playwright/test';
import { columnTitles, loadFiles, mergeNow, openApp, SAMPLE_FILES, SAMPLES_DIR, samplePaths, step, toastLog, waitMergeReady, watchErrors } from './helpers';
import { passwordProtectedXlsx, truncatedXlsx, waitIdle } from './round5';

const TWO = [SAMPLE_FILES[0], SAMPLE_FILES[1]];
const appErrors = (errors: { messages: string[] }) => errors.messages.filter((m) => !/favicon|DevTools|Failed to load resource/i.test(m));

interface DropItem {
  name: string;
  bytes: Buffer;
}
const sample = (name: string): DropItem => ({ name, bytes: fs.readFileSync(path.join(SAMPLES_DIR, name)) });

/** dragenter → dragover → drop on `target` with a DataTransfer carrying the files (and, optionally, plain text only). */
async function dropOn(page: Page, items: DropItem[], target = 'body', textOnly?: string): Promise<void> {
  const dt = await page.evaluateHandle(
    ({ list, text }) => {
      const d = new DataTransfer();
      if (text != null) d.setData('text/plain', text);
      for (const f of list) d.items.add(new File([new Uint8Array(f.bytes)], f.name, { lastModified: 1_700_000_000_000 }));
      return d;
    },
    { list: textOnly != null ? [] : items.map((i) => ({ name: i.name, bytes: Array.from(i.bytes) })), text: textOnly ?? null },
  );
  for (const type of ['dragenter', 'dragover', 'drop']) await page.dispatchEvent(target, type, { dataTransfer: dt });
}

test.describe('ラウンド5: ウィンドウへのドロップ', () => {
  test('(a) ウィンドウのどこにドロップしても読み込まれる（空の画面・カードの上・結合後）。ファイルでないドロップは何も起こさない', async ({ page }) => {
    test.setTimeout(5 * 60_000);
    const errors = watchErrors(page);
    await openApp(page);
    const url = page.url();

    // text dropped on the window: nothing is added, nothing breaks
    await dropOn(page, [], 'body', 'ただの文字');
    await expect(page.locator('[data-testid^="file-card-"]')).toHaveCount(0);
    expect(page.url()).toBe(url);

    // two files on the empty screen
    await dropOn(page, [sample(SAMPLE_FILES[0]), sample(SAMPLE_FILES[1])]);
    await expect(page.locator('[data-testid^="file-card-"]')).toHaveCount(2);
    await waitMergeReady(page, errors);
    await waitIdle(page);
    await expect(page.getByTestId('base-toggle-0')).toHaveAttribute('aria-pressed', 'true');
    await expect(page.getByTestId('summary-line')).toContainText(/基準\s*60\s*名/);

    // a third one on top of a card; the drag overlay is gone afterwards
    await dropOn(page, [sample(SAMPLE_FILES[2])], '[data-testid="file-card-0"]');
    await expect(page.locator('[data-testid^="file-card-"]')).toHaveCount(3);
    await waitMergeReady(page, errors);
    await waitIdle(page);
    await expect(page.locator('.dragging'), 'ドラッグ中の表示が残らない').toHaveCount(0);
    await expect(page.getByTestId('file-card-2')).toContainText('03_連絡先.csv');

    // a broken and a protected file dropped together: cards with reasons, the others are untouched
    await dropOn(page, [
      { name: '切れている.xlsx', bytes: truncatedXlsx() },
      { name: '保護.xlsx', bytes: passwordProtectedXlsx() },
    ]);
    await expect(page.locator('[data-testid^="file-card-"]')).toHaveCount(5);
    await expect(page.getByTestId('file-card-3').locator('.card-error')).toContainText('読み込めませんでした');
    await expect(page.getByTestId('file-card-4').locator('.card-error')).toContainText('パスワード');
    await waitMergeReady(page, errors);

    // after the merge a drop asks first (edits would be lost), 「キャンセル」 leaves everything as it is
    await mergeNow(page, errors);
    const before = await columnTitles(page);
    await dropOn(page, [sample(SAMPLE_FILES[4])]);
    await expect(page.getByTestId('confirm-dialog')).toBeVisible();
    await page.getByTestId('confirm-cancel').click();
    await expect(page.getByTestId('confirm-dialog')).toBeHidden();
    await expect(page.getByTestId('grid')).toBeVisible();
    expect(await columnTitles(page)).toEqual(before);
    await expect(page.locator('[data-testid^="file-card-"]')).toHaveCount(5);
    expect(appErrors(errors)).toEqual([]);
  });

  test('(b) ダイアログが開いている間のドロップは取り込まず、理由を知らせる。ダイアログを閉じればまたドロップできる', async ({ page }) => {
    test.setTimeout(5 * 60_000);
    const errors = watchErrors(page);
    await openApp(page);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, samplePaths(TWO)));
    await waitMergeReady(page, errors);
    await waitIdle(page);
    await page.getByTestId('key-detail-button').click();
    await expect(page.getByTestId('key-detail-modal')).toBeVisible();
    const n = (await toastLog(page)).length;
    await dropOn(page, [sample(SAMPLE_FILES[2])]);
    await expect.poll(async () => (await toastLog(page)).slice(n).some((t) => t.includes('開いている画面を閉じてから、もう一度ドロップしてください'))).toBe(true);
    await expect(page.locator('[data-testid^="file-card-"]')).toHaveCount(2);
    await page.getByTestId('key-detail-ok').click();
    await expect(page.getByTestId('key-detail-modal')).toBeHidden();
    await dropOn(page, [sample(SAMPLE_FILES[2])]);
    await expect(page.locator('[data-testid^="file-card-"]')).toHaveCount(3);
    expect(appErrors(errors)).toEqual([]);
  });
});

test.describe('ラウンド5: Excel の習慣のキー（Ctrl+S / Ctrl+O / F6）', () => {
  test('(c) Ctrl+S は結合前は理由を知らせ、結合後は出力ダイアログを開く。Ctrl+O はファイル選択。ダイアログ中・入力中は何も壊さない', async ({ page }) => {
    test.setTimeout(5 * 60_000);
    const errors = watchErrors(page);
    await openApp(page);
    // before the files / the merge: Ctrl+S explains, Ctrl+O opens the chooser
    const n0 = (await toastLog(page)).length;
    await page.keyboard.press('Control+s');
    await expect.poll(async () => (await toastLog(page)).slice(n0).some((t) => t.includes('保存は「結合する」のあと、「Excel出力」から行います'))).toBe(true);
    await expect(page.getByTestId('export-dialog')).toHaveCount(0);
    const chooser = page.waitForEvent('filechooser', { timeout: 10_000 });
    await page.keyboard.press('Control+o');
    const fc = await chooser;
    expect(fc.isMultiple(), 'ファイル選択は複数選べる').toBe(true);
    await fc.setFiles(samplePaths(TWO));
    await expect(page.locator('[data-testid^="file-card-"]')).toHaveCount(2);
    await waitMergeReady(page, errors);
    await mergeNow(page, errors);

    // after the merge: Ctrl+S opens the export dialog (and commits a cell that is being edited first)
    await page.getByTestId('grid-cell-4-3').click();
    await page.keyboard.type('編集中');
    await page.keyboard.press('Control+s');
    await expect(page.getByTestId('export-dialog')).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(page.getByTestId('export-dialog')).toBeHidden();
    await expect(page.getByTestId('grid-cell-4-3'), 'Ctrl+S で入力中の文字が確定している').toHaveText('編集中');
    await expect(page.getByTestId('side-summary')).toContainText('結合後の変更 1 件');
    // with a dialog open Ctrl+O / Ctrl+S do nothing more
    await page.getByTestId('export-button').click();
    await expect(page.getByTestId('export-dialog')).toBeVisible();
    await page.keyboard.press('Control+s');
    await page.keyboard.press('Control+o');
    await expect(page.getByTestId('export-dialog')).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(page.getByTestId('export-dialog')).toBeHidden();

    // F6 jumps between the grid and the main button of the side panel
    await page.getByTestId('grid-cell-5-1').click();
    await page.keyboard.press('F6');
    await expect(page.getByTestId('export-button')).toBeFocused();
    await page.keyboard.press('Shift+F6');
    await expect.poll(() => page.evaluate(() => !!document.activeElement?.closest('[data-testid="grid"]') || document.activeElement?.getAttribute('data-testid') === 'grid-editor')).toBe(true);
    expect(appErrors(errors)).toEqual([]);
  });
});
