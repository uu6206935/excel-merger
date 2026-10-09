// Loading the 5 samples, automatic keys (chips), the base toggle and the key detail modal (owner: T).
import { expect, test } from '@playwright/test';
import { loadFiles, openApp, requireTestIds, samplePaths, step, waitMergeReady, watchErrors } from './helpers';

test.describe('ファイル読み込みと結合キー', () => {
  test('5つのサンプルを読み込むと、基準・キー・見込みが自動で決まる', async ({ page }) => {
    const errors = watchErrors(page);
    const files = samplePaths();
    await openApp(page);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, files));

    // file cards: names, base (first file), row counts, formats
    await requireTestIds(page, ['file-card-0', 'base-toggle-0'], 'ファイルカード');
    await expect(page.getByTestId('file-card-0')).toContainText('01_学生名簿.xlsx');
    await expect(page.getByTestId('file-card-2')).toContainText('03_連絡先.csv');
    await expect(page.getByTestId('file-card-3')).toContainText('04_適性検査.xls');
    await expect(page.getByTestId('base-toggle-0')).toHaveAttribute('aria-pressed', 'true');
    for (let i = 1; i < files.length; i++) await expect(page.getByTestId(`base-toggle-${i}`)).toHaveAttribute('aria-pressed', 'false');
    await step(page, errors, 'ファイル読み込み', async () => {
      await expect(page.getByTestId('file-card-0')).toContainText(/\d+\s*行/, { timeout: 60_000 });
    });
    // row counts = data rows only (the 名簿 footer 人数/平均 below a blank row is NOT data)
    await expect(page.getByTestId('file-card-0'), '名簿は 60 行（1 行あけた下の集計行 人数・平均は含めない）').toContainText(/(^|[^0-9])60\s*行/);
    await expect.soft(page.getByTestId('file-card-1')).toContainText(/(^|[^0-9])63\s*行/);
    await expect.soft(page.getByTestId('file-card-2')).toContainText(/(^|[^0-9])47\s*行/);
    await expect.soft(page.getByTestId('file-card-3')).toContainText(/(^|[^0-9])42\s*行/);
    await expect.soft(page.getByTestId('file-card-4')).toContainText(/(^|[^0-9])32\s*行/);

    // header rows detected: 名簿 = row 3 (title + note above), 研修出欠 = row 2
    await requireTestIds(page, ['header-row-select-0', 'sheet-select-0'], 'シート・見出し行の選択');
    await expect(page.getByTestId('header-row-select-0')).toHaveValue('3');
    await expect(page.getByTestId('header-row-select-4')).toHaveValue('2');
    await expect(page.getByTestId('sheet-select-0')).toHaveValue(/名簿/);

    // keys
    await waitMergeReady(page, errors);
    await requireTestIds(page, ['key-chip-id', 'key-chip-name', 'summary-line'], '結合キーのチップ');
    await expect(page.getByTestId('key-chip-id')).toHaveAttribute('aria-pressed', 'true');
    await expect(page.getByTestId('key-chip-name')).toHaveAttribute('aria-pressed', 'true');
    await expect(page.getByTestId('summary-line'), '基準は 60 名（集計行を含めない）').toContainText(/基準\s*60\s*名/);
    // detected key columns per file
    await expect(page.getByTestId('file-keys-0')).toContainText('学生ID');
    await expect(page.getByTestId('file-keys-1')).toContainText('受験番号');
    await expect.soft(page.getByTestId('file-keys-2')).toContainText('氏名');
    await expect.soft(page.getByTestId('file-keys-3')).toContainText('氏名カナ');
    await expect.soft(page.getByTestId('file-keys-4')).toContainText('大学');

    // check panel: 未一致 / 重複 are counted before the merge. 未一致 = base rows found in NO other file: with all
    // five samples every student appears somewhere (truth.json), the rows missing from single files are listed
    // under 「一部のファイルに該当なし」 instead (e.g. 31・56 行目 have no 面接評価)
    await requireTestIds(page, ['check-panel'], 'チェック結果');
    await expect.soft(page.getByTestId('check-unmatched-count'), '5 ファイルでは、どのファイルにもいない学生はいない').toHaveText('0');
    await expect.soft(page.getByTestId('check-panel')).toContainText('一部のファイルに該当なし');
    await expect.soft(page.getByTestId('check-duplicate-count')).toHaveText(/[1-9]/);
  });

  test('キーのチップを切り替えられ、基準ファイルを変更して戻せる', async ({ page }) => {
    const errors = watchErrors(page);
    await openApp(page);
    await loadFiles(page, samplePaths());
    await waitMergeReady(page, errors);
    await requireTestIds(page, ['key-chip-name', 'summary-line', 'base-toggle-1'], '結合キーのチップ');

    const chip = page.getByTestId('key-chip-name');
    await chip.click();
    await expect(chip).toHaveAttribute('aria-pressed', 'false');
    await chip.click();
    await expect(chip).toHaveAttribute('aria-pressed', 'true');

    await page.getByTestId('base-toggle-1').click();
    await expect(page.getByTestId('base-toggle-1')).toHaveAttribute('aria-pressed', 'true');
    await expect(page.getByTestId('base-toggle-0')).toHaveAttribute('aria-pressed', 'false');
    await page.getByTestId('base-toggle-0').click();
    await expect(page.getByTestId('base-toggle-0')).toHaveAttribute('aria-pressed', 'true');
    await waitMergeReady(page, errors);
    await expect(page.getByTestId('summary-line'), '基準は 60 名').toContainText(/基準\s*60\s*名/);
  });

  test('詳細設定でファイルごとのキー候補（氏名＋大学など）を確認できる', async ({ page }) => {
    const errors = watchErrors(page);
    await openApp(page);
    await loadFiles(page, samplePaths());
    await waitMergeReady(page, errors);
    await requireTestIds(page, ['key-detail-button'], '詳細設定ボタン');
    await page.getByTestId('key-detail-button').click();
    await requireTestIds(page, ['key-detail-modal', 'key-detail-ok'], '詳細画面');
    const modal = page.getByTestId('key-detail-modal');
    await expect(modal).toBeVisible();
    const tabs = modal.locator('[data-testid^="key-detail-tab-"]');
    await expect(tabs).toHaveCount(4);
    await expect(modal).toContainText('一致率');

    // 02: ID key with a good match rate
    await tabs.filter({ hasText: '02_面接評価' }).click();
    await expect(modal).toContainText('受験番号');
    await expect(modal).toContainText('良好');
    // 05: 同姓同名 -> a composite key 氏名＋大学 must be offered
    await tabs.filter({ hasText: '05_研修出欠' }).click();
    await expect.soft(modal).toContainText(/氏名\s*[＋+]\s*大学/);
    // 04: 学籍番号 looks like an ID but does not match 学生ID
    await tabs.filter({ hasText: '04_適性検査' }).click();
    await expect.soft(modal).toContainText(/氏名カナ|フリガナ/);

    await page.getByTestId('key-detail-ok').click();
    await expect(modal).toBeHidden();
  });
});
