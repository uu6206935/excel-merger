/**
 * Playwright config for the grid's browser tests (owner: grid G). Runs against the dev-only demo page
 * http://127.0.0.1:5175/grid-demo.html in the installed Microsoft Edge (channel 'msedge'; no browser download).
 *
 *   npx playwright test -c tests/grid/playwright.config.ts
 */
import { defineConfig } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../..', import.meta.url));

export default defineConfig({
  testDir: './e2e',
  timeout: 90_000,
  expect: { timeout: 10_000 },
  workers: 1,
  fullyParallel: false,
  retries: 0,
  reporter: [['list']],
  outputDir: fileURLToPath(new URL('./.pw-results', import.meta.url)),
  use: {
    baseURL: 'http://127.0.0.1:5175',
    channel: 'msedge',
    headless: true,
    viewport: { width: 1400, height: 820 },
    deviceScaleFactor: 1,
    locale: 'ja-JP',
    trace: 'off',
  },
  webServer: {
    command: 'npx vite --config tests/grid/vite.demo.config.ts',
    cwd: root,
    url: 'http://127.0.0.1:5175/grid-demo.html',
    reuseExistingServer: true,
    timeout: 60_000,
  },
});
