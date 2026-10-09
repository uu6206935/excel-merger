/**
 * Playwright E2E config (owner: T). Uses the INSTALLED Microsoft Edge (channel 'msedge'); never downloads browsers.
 *
 * Server mode (default): starts `node server/server.mjs` on :5199 (NO_OPEN, never exits on window close,
 * outputs in test-results/e2e-output). The UI it serves:
 *   E2E_BUILD=1 npx playwright test        builds the UI into test-results/e2e-dist-<run id> first (does not touch dist/)
 *   npx playwright test                    uses the newest test-results/e2e-dist* build if there is one, else dist/ (npm run build)
 *   E2E_PORT=5176 npx playwright test      if :5199 is busy
 * Standalone mode (no server: the JS engine downloads, no Excel engine):
 *   npx vite --config e2e/vite.e2e.config.ts   (port 5176)
 *   E2E_BASE_URL=http://127.0.0.1:5176 npx playwright test
 * Fallback server (server mode only): a SECOND server on :5176 (E2E_FAKE_PORT) whose "Excel" is e2e/fixtures/fake-excel-runner.mjs
 * (EXCEL_DETECT=installed, EXCEL_RUNNER=<fake>), switched per test through the JSON file named by E2E_FAKE_CONTROL. It is what
 * e2e/09-round5-fallback.spec.ts drives to test the automatic fall-back to the built-in engine without starting Excel. In
 * standalone mode the spec skips (vite owns :5176 there).
 *
 * Disk / Excel: every Excel start (the Excel engine, check-output, cell probes) writes ~80 MB into %TEMP%; run with TEMP/TMP on a
 * big drive. QA_NO_EXCEL=1 skips the steps that need Excel (check-output, probes); the server's Excel engine is unaffected by it.
 *
 * Every run has its own artifact folders (test-results/e2e-run-<id>, e2e-dist-<id>, e2e-fake-control-<id>.json; folders older than
 * 6 h are removed at the start): another `playwright test` started in this repo by someone sharing the machine clears its outputDir
 * at startup, which used to delete the trace files of a test that was running here ("browserContext.close: ENOENT … .trace").
 *
 * Specs follow the data-testid contract in docs/ARCHITECTURE.md §5; steps the UI / core modules do not support
 * yet mark themselves test.fixme at runtime (see e2e/helpers.ts requireTestIds / fixmeIfNotImplemented).
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig } from '@playwright/test';

// E2E_PORT: override when 5199 is taken (the assigned E2E ports are 5199 and 5176)
const PORT = Number(process.env.E2E_PORT || 5199);
const external = process.env.E2E_BASE_URL?.replace(/\/+$/, '');
const root = path.dirname(fileURLToPath(import.meta.url));
const results = path.join(root, 'test-results');
const outputDir = path.join(results, 'e2e-output');
const build = process.env.E2E_BUILD === '1';

// one id per run (the main process creates it, the workers inherit it through the environment)
const freshRun = !process.env.E2E_RUN_ID;
if (freshRun) process.env.E2E_RUN_ID = String(Date.now());
const RUN_ID = process.env.E2E_RUN_ID as string;
if (freshRun) {
  // artifacts of earlier runs (older than 6 h; the folders of a run that may still be running are younger)
  try {
    for (const name of fs.readdirSync(results)) {
      if (!/^e2e-(run|dist)-\d{10,}$|^e2e-fake-control-\d{10,}\.json$/.test(name)) continue;
      const full = path.join(results, name);
      if (Date.now() - fs.statSync(full).mtimeMs > 6 * 3600_000) fs.rmSync(full, { recursive: true, force: true });
    }
  } catch {
    /* nothing to clean / not allowed: harmless */
  }
}
/** the UI build to serve when E2E_BUILD is not set: the newest test-results/e2e-dist* with an index.html */
function newestBuild(): string | null {
  try {
    const found = fs
      .readdirSync(results)
      .filter((n) => /^e2e-dist(-\d+)?$/.test(n) && fs.existsSync(path.join(results, n, 'index.html')))
      .map((n) => ({ dir: path.join(results, n), t: fs.statSync(path.join(results, n, 'index.html')).mtimeMs }))
      .sort((a, b) => b.t - a.t);
    return found[0]?.dir ?? null;
  } catch {
    return null;
  }
}
const e2eDist = path.join(results, `e2e-dist-${RUN_ID}`);
// the second (fake-Excel) server of the fallback spec; the spec finds it through these two variables (workers inherit them)
const FAKE_PORT = Number(process.env.E2E_FAKE_PORT || 5176);
const fakeControl = path.join(results, `e2e-fake-control-${RUN_ID}.json`);
process.env.E2E_FAKE_CONTROL = fakeControl;
process.env.E2E_FAKE_URL = `http://localhost:${FAKE_PORT}`;
const distDir = process.env.E2E_DIST ?? (build ? e2eDist : newestBuild() ?? path.join(root, 'dist'));

export default defineConfig({
  testDir: './e2e',
  testMatch: /.*\.spec\.ts$/,
  timeout: 4 * 60 * 1000,
  expect: { timeout: 30 * 1000 },
  fullyParallel: false,
  workers: 1,
  retries: 0,
  reporter: [['list'], ['html', { open: 'never', outputFolder: 'test-results/e2e-report' }]],
  outputDir: path.join('test-results', `e2e-run-${RUN_ID}`),
  use: {
    baseURL: external ?? `http://localhost:${PORT}`,
    channel: 'msedge',
    headless: process.env.E2E_HEADED !== '1',
    viewport: { width: 1440, height: 900 },
    locale: 'ja-JP',
    timezoneId: 'Asia/Tokyo',
    acceptDownloads: true,
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
    actionTimeout: 30 * 1000,
  },
  projects: [{ name: 'msedge', use: { channel: 'msedge' } }],
  webServer: external
    ? undefined
    : [
        {
          command: build
            ? `npx vite build --outDir "${e2eDist}" --emptyOutDir && node server/server.mjs`
            : 'node server/server.mjs',
          url: `http://localhost:${PORT}/api/ping`,
          env: {
            NO_OPEN: '1',
            PORT: String(PORT),
            DIST_DIR: distDir,
            // the server normally exits when its window closes; E2E opens/closes pages all the time
            EXIT_GRACE_MS: '0',
            IDLE_EXIT_MS: '0',
            OUTPUT_DIR: outputDir,
          },
          // never reuse: a foreign server on the port (e.g. a stale vite preview) would answer /api/ping with HTML
          reuseExistingServer: false,
          timeout: 5 * 60 * 1000,
          stdout: 'pipe',
          stderr: 'pipe',
        },
        {
          // same UI, but the "Excel" is a fake runner that fails on demand (started after the first server: dist exists by then)
          command: 'node server/server.mjs',
          url: `http://localhost:${FAKE_PORT}/api/ping`,
          env: {
            NO_OPEN: '1',
            PORT: String(FAKE_PORT),
            DIST_DIR: distDir,
            EXIT_GRACE_MS: '0',
            IDLE_EXIT_MS: '0',
            OUTPUT_DIR: path.join(root, 'test-results', 'e2e-output-fake'),
            EXCEL_DETECT: 'installed',
            EXCEL_RUNNER: path.join(root, 'e2e', 'fixtures', 'fake-excel-runner.mjs'),
            E2E_FAKE_CONTROL: fakeControl,
          },
          reuseExistingServer: false,
          timeout: 60 * 1000,
          stdout: 'pipe',
          stderr: 'pipe',
        },
      ],
});
