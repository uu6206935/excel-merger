// Round 5: the clumsy things real users do (owner: T) — double-clicking buttons, odd export file names, removing the base file and
// dropping it again. None of it may double a merge / an export, write a bad file name, or leave the app in a state that needs a reload.
import fs from 'node:fs';
import path from 'node:path';
import { expect, test } from '@playwright/test';
import {
  columnTitles,
  editCount,
  isServerMode,
  loadFiles,
  mergeNow,
  mergeSamples,
  openApp,
  OUT_DIR,
  requireTestIds,
  runExport,
  SAMPLE_FILES,
  samplePaths,
  step,
  waitMergeReady,
  watchErrors,
} from './helpers';
import { readSheet, waitIdle } from './round5';

const TWO = [SAMPLE_FILES[0], SAMPLE_FILES[1]];
const appErrors = (errors: { messages: string[] }) => errors.messages.filter((m) => !/favicon|DevTools|Failed to load resource/i.test(m));

test.describe('ラウンド5: 連打・変なファイル名・基準ファイルの取り外し', () => {
  test('(a) 「結合する」を素早く 2 回押しても 1 回だけ結合され、「Excel出力」の「出力」を 2 回押しても 1 ファイルしかできない', async ({ page }) => {
    test.setTimeout(10 * 60_000);
    const errors = watchErrors(page);
    await openApp(page);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, samplePaths(TWO)));
    await waitMergeReady(page, errors);
    await waitIdle(page);
    await page.getByTestId('merge-button').dblclick({ delay: 10 });
    await expect(page.getByTestId('grid')).toBeVisible({ timeout: 120_000 });
    await waitIdle(page);
    expect(await editCount(page), '結合直後は変更 0 件').toBe(0);
    const titles = await columnTitles(page);
    expect(titles.filter((t, i) => titles.indexOf(t) !== i), '見出しが重複していない（2 回結合されていない）').toEqual([]);
    expect(titles.length, '01 の 12 列 + 02 の取り込み 7 列').toBe(19);

    // the export button twice in a row
    const downloads: string[] = [];
    page.on('download', (d) => downloads.push(d.suggestedFilename()));
    const stem = `e2e_r5j_連打_${Date.now()}`;
    await page.getByTestId('export-button').click();
    await page.getByTestId('export-engine-js').check();
    await page.getByTestId('export-filename').fill(`${stem}.xlsx`);
    await page.getByTestId('export-confirm').dblclick({ delay: 10 });
    const result = page.getByTestId('export-result');
    await expect(result.or(page.getByTestId('export-error'))).toBeVisible({ timeout: 2 * 60_000 });
    await expect(page.getByTestId('export-error')).toHaveCount(0);
    await expect(result).toBeVisible();
    await waitIdle(page);
    if (await isServerMode(page)) {
      const dir = ((await (await page.request.get('/api/status')).json()) as { outputDir: string }).outputDir;
      const made = fs.readdirSync(dir).filter((f) => f.startsWith(stem));
      expect(made, '出力フォルダに 1 ファイルだけ').toEqual([`${stem}.xlsx`]);
    } else {
      expect(downloads, 'ダウンロードは 1 回だけ').toEqual([`${stem}.xlsx`]);
    }
    await page.getByTestId('export-close').click();
    expect(appErrors(errors)).toEqual([]);
  });

  test('(b) 出力のファイル名: 使えない文字・拡張子なし・別の拡張子・空・長すぎる名前でも、Excel で開ける正しい名前のファイルになる', async ({ page }) => {
    test.setTimeout(10 * 60_000);
    const errors = watchErrors(page);
    await mergeSamples(page, errors, TWO);
    const stamp = Date.now();
    const cases: Array<{ typed: string; expectName: RegExp }> = [
      { typed: `a:b*c?d_${stamp}`, expectName: new RegExp(`^a_b_c_d_${stamp}\\.xlsx$`) }, // illegal characters, no extension
      { typed: `名簿 統合 ${stamp}.XLSX`, expectName: new RegExp(`^名簿 統合 ${stamp}\\.xlsx$`) }, // upper-case extension
      { typed: `旧形式_${stamp}.xls`, expectName: new RegExp(`^旧形式_${stamp}(\\.xls)?\\.xlsx$`) }, // the wrong extension never stays alone
      { typed: `${'とても長いファイル名_'.repeat(12)}${stamp}`, expectName: /\.xlsx$/ },
    ];
    for (const c of cases) {
      const file = await runExport(page, errors, 'js', c.typed, path.join(OUT_DIR, 'names'));
      const base = path.basename(file);
      test.info().annotations.push({ type: `ファイル名「${c.typed.slice(0, 30)}…」`, description: base });
      expect(base, `「${c.typed.slice(0, 40)}」で出力したファイル名`).toMatch(c.expectName);
      expect(base, 'Windows で使えない文字がない').not.toMatch(/[\\/:*?"<>|]/);
      expect(readSheet(file, '名簿').rows.length, '開けるブック').toBeGreaterThan(60);
    }

    // an empty name cannot be exported; the button says why is not needed but must be disabled
    await page.getByTestId('export-button').click();
    await requireTestIds(page, ['export-filename', 'export-confirm'], '出力ダイアログ');
    await page.getByTestId('export-filename').fill('   ');
    await expect(page.getByTestId('export-confirm')).toBeDisabled();
    await expect(page.getByTestId('export-dialog')).toContainText('ファイル名を入力してください');
    await page.getByTestId('export-filename').fill(`e2e_r5j_再入力_${stamp}`);
    await expect(page.getByTestId('export-confirm')).toBeEnabled();
    await page.keyboard.press('Escape');
    await expect(page.getByTestId('export-dialog')).toBeHidden();
    expect(appErrors(errors)).toEqual([]);
  });

  test('(c) 基準ファイルを「クリア」して、もう一度入れる → 基準は最初に読めたファイルに移り、また入れた方は後ろに付く。結合はやり直せる', async ({ page }) => {
    test.setTimeout(10 * 60_000);
    const errors = watchErrors(page);
    await openApp(page);
    const [f01, f02] = samplePaths(TWO);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, [f01, f02]));
    await waitMergeReady(page, errors);
    await waitIdle(page);
    await expect(page.getByTestId('summary-line')).toContainText(/基準\s*60\s*名/);

    // remove the base (card 0): 02 is alone and becomes the base; nothing to merge with
    await page.getByTestId('file-clear-0').click();
    await expect(page.locator('[data-testid^="file-card-"]')).toHaveCount(1);
    await expect(page.getByTestId('file-card-0')).toContainText('02_面接評価.xlsx');
    await expect(page.getByTestId('merge-button')).toBeDisabled();
    await expect(page.getByTestId('summary-line')).toHaveCount(0);
    await waitIdle(page);

    // 01 comes back: it is the SECOND file now, so 02 stays the base (「基準にする」 switches)
    await page.getByTestId('file-input').setInputFiles([f01]);
    await expect(page.locator('[data-testid^="file-card-"]')).toHaveCount(2);
    await waitMergeReady(page, errors);
    await waitIdle(page);
    await expect(page.getByTestId('base-toggle-0'), '残っていた 02 が基準').toHaveAttribute('aria-pressed', 'true');
    await expect(page.getByTestId('summary-line')).toContainText(/基準\s*63\s*名/);
    await page.getByTestId('base-toggle-1').click();
    await expect(page.getByTestId('base-toggle-1')).toHaveAttribute('aria-pressed', 'true');
    await waitMergeReady(page, errors);
    await expect(page.getByTestId('summary-line')).toContainText(/基準\s*60\s*名/);

    // merge, go back with 「結合をやり直す」 and merge again: the same result
    await mergeNow(page, errors);
    const first = await columnTitles(page);
    await page.getByTestId('remerge-button').click();
    await expect(page.getByTestId('confirm-ok').or(page.getByTestId('merge-button'))).toBeVisible();
    if (await page.getByTestId('confirm-ok').isVisible()) await page.getByTestId('confirm-ok').click();
    await expect(page.getByTestId('merge-button')).toBeVisible();
    await waitMergeReady(page, errors);
    await mergeNow(page, errors);
    expect(await columnTitles(page), '結合をやり直しても同じ結果').toEqual(first);
    expect(appErrors(errors)).toEqual([]);
  });
});
