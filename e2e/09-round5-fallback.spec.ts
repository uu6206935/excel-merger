// Round 5: the Excel本体 engine FAILS — the app must fall back to the built-in engine by itself (owner: T).
//
// The real UI and the real local server are driven against a second server (playwright.config.ts, :5176 / E2E_FAKE_PORT) whose
// "Excel" is e2e/fixtures/fake-excel-runner.mjs (EXCEL_DETECT=installed, EXCEL_RUNNER=<fake>); the spec tells the fake how to fail
// through the JSON file named by E2E_FAKE_CONTROL (setFakeExcel). Nothing starts Excel. Server mode only: in standalone mode
// (E2E_BASE_URL, vite on :5176) there is no Excel engine at all — that case is checked by the last test.
import fs from 'node:fs';
import path from 'node:path';
import { expect, test, type Page } from '@playwright/test';
import { columnTitles, editCount, mergeSamples, moveColumns, OUT_DIR, runExport, SAMPLE_FILES, toastLog, watchErrors } from './helpers';
import { readSheet, samePackage, setFakeExcel, waitIdle } from './round5';

const FAKE_URL = process.env.E2E_FAKE_URL;
const standalone = !!process.env.E2E_BASE_URL;
const TWO = [SAMPLE_FILES[0], SAMPLE_FILES[1]];

test.use({ baseURL: standalone ? process.env.E2E_BASE_URL : FAKE_URL });

/** Opens the export dialog, leaves the Excel engine selected (the default when the server offers it) and starts the export. */
async function startExcelExport(page: Page, fileName: string): Promise<void> {
  await page.getByTestId('export-button').click();
  await expect(page.getByTestId('export-dialog')).toBeVisible();
  await expect(page.getByTestId('export-engine-excel'), 'サーバーが Excel ありと答えるので「Excel本体で出力」を選べる').toBeEnabled();
  await page.getByTestId('export-engine-excel').check();
  await page.getByTestId('export-filename').fill(fileName);
  await page.getByTestId('export-confirm').click();
}

/** Waits for the result of an export and returns it: the saved path, the engine line, the notes. */
async function exportResult(page: Page) {
  const result = page.getByTestId('export-result');
  const error = page.getByTestId('export-error');
  await expect(result.or(error)).toBeVisible({ timeout: 2 * 60_000 });
  if (await error.isVisible()) throw new Error(`出力に失敗しました: ${await error.innerText()}`);
  const pathEl = result.locator('.result-path');
  return {
    file: (await pathEl.count()) ? (await pathEl.innerText()).trim() : '',
    engine: (await page.getByTestId('export-result-engine').innerText()).trim(),
    text: (await result.innerText()).replace(/\s+/g, ' ').trim(),
    fellBack: (await result.getByText('Excel本体での出力に失敗したため、内蔵エンジンで出力しました。').count()) > 0,
  };
}

test.describe('ラウンド5: Excel本体エンジンが失敗したら、内蔵エンジンに自動で切り替わる', () => {
  test.skip(standalone || !FAKE_URL, 'サーバーモードのみ（E2E_BASE_URL の vite 単体には Excel エンジンがない）。第 2 のサーバーは playwright.config.ts が :5176 で起動する');

  const CODES: Array<[string, string]> = [
    ['MOVE_VERIFY_FAILED', 'テーブルの中の列の移動を Excel本体では正しく行えなかったため中止しました（保存していません）。「内蔵エンジン」で出力してください。'],
    ['OPEN_FAILED', 'Excel でブックを開けませんでした'],
    ['PASSWORD_PROTECTED', 'パスワードで保護されているため Excel で開けません'],
    ['EXCEL_UNAVAILABLE', 'Excel を起動できませんでした'],
    ['RUNNER_MISSING', 'Excel 用のスクリプトが見つかりません'],
    ['EXCEL_REFUSED', 'Excel が「列の移動」を実行できませんでした'], // not one of the server's "fall back" codes: the UI falls back for any failure
  ];

  for (const [code, message] of CODES) {
    test(`${code}: 自動で内蔵エンジンに切り替わり、出力は内蔵エンジンで直接出したものと同じ。アプリはそのまま使える`, async ({ page }) => {
      test.setTimeout(6 * 60_000);
      const errors = watchErrors(page);
      setFakeExcel({ mode: 'fail', code, message });
      await mergeSamples(page, errors, TWO);
      await moveColumns(page, ['判定'], 0); // a real edit to replay (an imported column lies outside the title merge: no confirmation)
      await expect.poll(async () => (await columnTitles(page))[0]).toBe('判定');
      const edits = await editCount(page);
      const n = (await toastLog(page)).length;

      await startExcelExport(page, `e2e_r5fb_${code}_${Date.now()}.xlsx`);
      const r = await exportResult(page);
      test.info().annotations.push({ type: `出力の結果 ${code}`, description: r.text });
      expect(r.engine, '内蔵エンジンで出力された').toBe('内蔵エンジンで出力しました');
      expect(r.fellBack, 'ダイアログが「Excel本体での出力に失敗したため、内蔵エンジンで出力しました。」と知らせる').toBe(true);
      expect(r.file, '出力フォルダに保存された').not.toBe('');
      expect(fs.existsSync(r.file)).toBe(true);
      const toasts = (await toastLog(page)).slice(n);
      expect(toasts.some((t) => t.includes('Excel本体で出力できなかったため、内蔵エンジンで出力します')), `トーストで知らせる（${toasts.join(' / ')}）`).toBe(true);

      // the fallback file is the built-in engine's own output of the same state
      await page.getByTestId('export-close').click();
      await expect(page.getByTestId('export-dialog')).toBeHidden();
      const direct = await runExport(page, errors, 'js', `e2e_r5fb_direct_${Date.now()}.xlsx`, path.join(OUT_DIR, 'fallback'));
      const same = samePackage(r.file, direct);
      expect(same.identical, `切り替わった出力は内蔵エンジンの出力と同じ\n${same.text}`).toBe(true);
      const grid = readSheet(r.file, '名簿');
      expect(grid.rows[2][0], '移動した列（判定）が先頭').toBe('判定');

      // nothing was lost on the screen: the edits are still there, undo works, the next export works
      expect(await editCount(page)).toBe(edits);
      await page.getByTestId('undo-button').click();
      await expect.poll(async () => (await columnTitles(page))[0]).toBe('学生ID');
      await page.getByTestId('redo-button').click();
      await expect.poll(async () => (await columnTitles(page))[0]).toBe('判定');
      expect(errors.messages.filter((m) => !/favicon|DevTools|Failed to load resource/i.test(m))).toEqual([]);
    });
  }

  test('Excel が突然終了した（最終行なしで終了コード 3）・成功と言いながら出力がない → どちらも内蔵エンジンで出力できる', async ({ page }) => {
    test.setTimeout(6 * 60_000);
    const errors = watchErrors(page);
    await mergeSamples(page, errors, TWO);
    for (const mode of ['crash', 'noout'] as const) {
      setFakeExcel({ mode });
      await startExcelExport(page, `e2e_r5fb_${mode}_${Date.now()}.xlsx`);
      const r = await exportResult(page);
      test.info().annotations.push({ type: `出力の結果 ${mode}`, description: r.text });
      expect(r.engine).toBe('内蔵エンジンで出力しました');
      expect(r.fellBack).toBe(true);
      expect(fs.existsSync(r.file)).toBe(true);
      expect(readSheet(r.file, '名簿').rows.length, '出力は読めるブック').toBeGreaterThan(60);
      await page.getByTestId('export-close').click();
      await expect(page.getByTestId('export-dialog')).toBeHidden();
    }
  });

  test('Excel本体の出力を「中止」する → ダイアログは元に戻り、そのあと内蔵エンジンで出力できる（残らない）', async ({ page }) => {
    test.setTimeout(6 * 60_000);
    const errors = watchErrors(page);
    await mergeSamples(page, errors, TWO);
    setFakeExcel({ mode: 'slow', ms: 120_000 });
    const n = (await toastLog(page)).length;
    await startExcelExport(page, `e2e_r5fb_cancel_${Date.now()}.xlsx`);
    await expect(page.getByTestId('export-progress'), 'Excel本体が動いている間は進行表示').toBeVisible();
    await expect(page.getByTestId('export-cancel')).toBeVisible();
    await page.getByTestId('export-cancel').click();
    await expect(page.getByTestId('export-progress')).toHaveCount(0);
    await expect(page.getByTestId('export-confirm'), '中止したらもう一度出力できる').toBeVisible();
    await expect(page.getByTestId('export-result')).toHaveCount(0);
    await expect(page.getByTestId('export-error'), '中止はエラーではない').toHaveCount(0);
    await expect.poll(async () => (await toastLog(page)).slice(n).some((t) => t.includes('出力を中止しました'))).toBe(true);
    // the same dialog: now with the built-in engine
    await page.getByTestId('export-engine-js').check();
    await page.getByTestId('export-filename').fill(`e2e_r5fb_after_cancel_${Date.now()}.xlsx`);
    await page.getByTestId('export-confirm').click();
    const r = await exportResult(page);
    expect(r.engine).toBe('内蔵エンジンで出力しました');
    expect(r.fellBack, '自分で選んだ内蔵エンジンは「失敗して切り替わった」扱いにならない').toBe(false);
    expect(fs.existsSync(r.file)).toBe(true);
    await page.getByTestId('export-close').click();
    await waitIdle(page);
  });
});

test.describe('ラウンド5: スタンドアロン（サーバーなし）', () => {
  test.skip(!standalone, 'E2E_BASE_URL（vite 単体）のときだけ: サーバーモードには Excel 本体エンジンがある');

  test('Excel本体エンジンは選べず、内蔵エンジンで出力でき、理由が出る', async ({ page }) => {
    test.setTimeout(5 * 60_000);
    const errors = watchErrors(page);
    await mergeSamples(page, errors, TWO);
    await page.getByTestId('export-button').click();
    await expect(page.getByTestId('export-dialog')).toBeVisible();
    await expect(page.getByTestId('export-engine-excel')).toBeDisabled();
    await expect(page.getByTestId('export-engine-js')).toBeChecked();
    await expect(page.getByTestId('export-dialog')).toContainText('スタンドアロンで開いているため使えません');
    await page.keyboard.press('Escape');
    await expect(page.getByTestId('export-dialog')).toBeHidden();
    const file = await runExport(page, errors, 'js', `e2e_r5fb_standalone_${Date.now()}.xlsx`, OUT_DIR);
    expect(readSheet(file, '名簿').rows.length).toBeGreaterThan(60);
  });
});
