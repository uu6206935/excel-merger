/**
 * Playwright config of the round-5 UI hardening flows (tests/ui/hardening.spec.ts), owner U. Installed Microsoft Edge only
 * (channel 'msedge'; never downloads browsers).
 *
 *   npx playwright test -c tests/ui/playwright.config.ts
 *
 * It builds the UI into a scratch folder and starts the REAL server (server/server.mjs, port 5192) with its output, data and job
 * folders in the scratch folder too, so the repo's dist/ and output/ stay untouched:
 *   UI_SCRATCH=<dir>   scratch folder (default: <TEMP>/excel-merger-ui-tests)
 *   UI_PORT=5192       server port
 *   UI_BASE_URL=http://localhost:5192   use a server that is already running (no build, no start)
 *   UI_SKIP_BUILD=1    reuse <scratch>/dist
 * Set TEMP / TMP to a big drive first (every Excel start writes ~80 MB of diagnostics into %TEMP%). These specs use the built-in
 * engine only (no Excel start) unless a test says otherwise.
 */
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig } from '@playwright/test';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const PORT = Number(process.env.UI_PORT || 5192);
const external = process.env.UI_BASE_URL?.replace(/\/+$/, '');
const scratch = path.resolve(process.env.UI_SCRATCH || path.join(os.tmpdir(), 'excel-merger-ui-tests'));
const dist = path.join(scratch, 'dist');
const build = process.env.UI_SKIP_BUILD === '1' ? 'node -e "0"' : `npx vite build --outDir "${dist}" --emptyOutDir`;

export default defineConfig({
  testDir: '.',
  testMatch: /.*\.spec\.ts$/,
  timeout: 3 * 60 * 1000,
  expect: { timeout: 20 * 1000 },
  fullyParallel: false,
  workers: 1,
  retries: 0,
  reporter: [['list']],
  outputDir: path.join(scratch, 'pw-results'),
  use: {
    baseURL: external ?? `http://localhost:${PORT}`,
    channel: 'msedge',
    headless: process.env.UI_HEADED !== '1',
    viewport: { width: 1440, height: 900 },
    locale: 'ja-JP',
    timezoneId: 'Asia/Tokyo',
    acceptDownloads: true,
    actionTimeout: 30 * 1000,
  },
  projects: [{ name: 'msedge', use: { channel: 'msedge' } }],
  webServer: external
    ? undefined
    : {
        command: `${build} && node server/server.mjs`,
        cwd: root,
        url: `http://localhost:${PORT}/api/ping`,
        env: {
          NO_OPEN: '1',
          PORT: String(PORT),
          DIST_DIR: dist,
          OUTPUT_DIR: path.join(scratch, 'output'),
          EM_DATA_DIR: path.join(scratch, 'data'),
          EM_TEMP_DIR: path.join(scratch, 'jobs'),
          EXIT_GRACE_MS: '0',
          IDLE_EXIT_MS: '0',
        },
        reuseExistingServer: false,
        timeout: 5 * 60 * 1000,
        stdout: 'pipe',
        stderr: 'pipe',
      },
});
