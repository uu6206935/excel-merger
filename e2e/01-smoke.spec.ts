// Startup: the app renders, shows the empty drop zone and the server / Excel status (owner: T).
import { expect, test } from '@playwright/test';
import { isServerMode, openApp, requireTestIds, watchErrors } from './helpers';

test('起動すると空のドロップ領域と接続状態を表示する', async ({ page }) => {
  const errors = watchErrors(page);
  await openApp(page);
  await expect(page).toHaveTitle(/Excel 統合ツール/);
  await expect(page.getByRole('heading', { name: 'Excel 統合ツール' })).toBeVisible();
  await requireTestIds(page, ['drop-zone', 'server-status', 'file-input'], '初期画面');
  await expect(page.getByTestId('drop-zone')).toBeVisible();
  const server = await isServerMode(page);
  const status = page.getByTestId('server-status');
  await expect(status).toContainText(server ? 'ローカル実行中' : 'スタンドアロン');
  await expect(status).toContainText('Excel本体');
  // nothing blew up while rendering
  expect(errors.messages.filter((m) => !/favicon|DevTools|Failed to load resource/i.test(m))).toEqual([]);
});

test('ローカルサーバーの API に応答がある（サーバーモードのみ）', async ({ page }) => {
  test.skip(!!process.env.E2E_BASE_URL, 'E2E_BASE_URL（vite 単体）ではサーバーを使わない');
  const ping = await page.request.get('/api/ping');
  expect(ping.ok()).toBe(true);
  expect((await ping.json()).app).toBe('excel-merger');
  const st = await page.request.get('/api/status');
  expect(st.ok()).toBe(true);
  const j = await st.json();
  expect(typeof j.excel?.installed).toBe('boolean');
  expect(typeof j.outputDir).toBe('string');
});
