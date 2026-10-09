// Round 5: screenshots of the main screens of the REAL app (real modules, real server) -> tests/ui/shots/final-*.png
//   node tests/ui/final-shots.mjs [baseUrl=http://localhost:5192/] [fixtureDir]
// The fixtures (hostile files) are generated with tests/ui/hostile-fixtures.mjs into fixtureDir (default: a folder in %TEMP%).
import { chromium } from '@playwright/test';
import { mkdirSync } from 'node:fs';
import os from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { makeHostileFixtures } from './hostile-fixtures.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const outDir = join(here, 'shots');
mkdirSync(outDir, { recursive: true });
const base = process.argv[2] ?? 'http://localhost:5192/';
const fxDir = resolve(process.argv[3] ?? join(os.tmpdir(), 'excel-merger-final-shots'));
const samples = join(here, '..', '..', 'samples');
const sample = (n) => join(samples, n);
const fx = await makeHostileFixtures(fxDir);

const browser = await chromium.launch({ channel: 'msedge', headless: true });
const failures = [];
const errors = [];

async function session(viewport, dsf, fn) {
  const ctx = await browser.newContext({ viewport, deviceScaleFactor: dsf, locale: 'ja-JP', acceptDownloads: true });
  const page = await ctx.newPage();
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (m) => m.type() === 'error' && !/Failed to load resource|ERR_/.test(m.text()) && errors.push(m.text()));
  await page.goto(base);
  await page.getByTestId('server-status').waitFor();
  await page.waitForTimeout(500);
  try {
    await fn(page, ctx);
  } catch (e) {
    failures.push(String(e));
    console.log('FAILED:', e.message);
  }
  await ctx.close();
}
const settle = async (page) => {
  await page.waitForFunction(() => !document.querySelector('[data-testid=busy-overlay]') && !document.querySelector('.card-loading'), null, { timeout: 120000 });
  await page.waitForTimeout(300);
};
const shot = async (page, name) => {
  await page.waitForTimeout(250);
  await page.screenshot({ path: join(outDir, `final-${name}.png`) });
  console.log('saved', name);
};
const load = async (page, files) => {
  await page.getByTestId('file-input').setInputFiles(files);
  await settle(page);
};
const merge = async (page) => {
  await page.getByTestId('merge-button').click();
  await page.getByTestId('grid').waitFor({ timeout: 120000 });
  await settle(page);
  await page.waitForTimeout(400);
};
const five = ['01_学生名簿.xlsx', '02_面接評価.xlsx', '03_連絡先.csv', '04_適性検査.xls', '05_研修出欠.xlsx'].map(sample);

// 1440x900: the main screens of the flow
await session({ width: 1440, height: 900 }, 1, async (page) => {
  await shot(page, '01-empty-help');
  await page.getByTestId('help-button').click();
  await shot(page, '02-help-dialog');
  await page.keyboard.press('Escape');
  await load(page, five);
  await shot(page, '03-files-five');
  await page.getByTestId('key-detail-button').click();
  await page.getByTestId('key-detail-modal').waitFor();
  await shot(page, '04-key-detail');
  await page.keyboard.press('Escape');
  await merge(page);
  await shot(page, '05-merged');
  await page.getByTestId('grid-header-3').click({ button: 'right' });
  await shot(page, '06-context-menu');
  await page.keyboard.press('Escape');
  await page.getByTestId('grid-header-3').click();
  await page.keyboard.press('Delete');
  await page.getByTestId('delete-confirm-dialog').waitFor();
  await shot(page, '07-delete-confirm');
  await page.getByTestId('delete-cancel-button').click();
  await page.getByTestId('export-button').click();
  await page.getByTestId('export-dialog').waitFor();
  await shot(page, '08-export-dialog');
  await page.getByTestId('export-engine-js').check();
  await page.getByTestId('export-confirm').click();
  await page.getByTestId('export-result').waitFor({ timeout: 120000 });
  await shot(page, '09-export-result');
});

// hostile inputs: error cards, a long file name, the folder drop message
await session({ width: 1440, height: 900 }, 1, async (page) => {
  const long = Object.keys(fx).find((k) => k.startsWith('非常に長い'));
  await load(page, [fx['a_名簿.xlsx'], fx['zero.xlsx'], fx['corrupt.xlsx'], fx['password.xlsx'], fx['notexcel.xlsx'], fx[long], fx['メモ.txt']]);
  await shot(page, '10-error-cards');
});

// 1280x720 (small laptop), 1920x1080, 1024 wide (horizontal scroll), browser zoom 150%
for (const [label, vp, dsf] of [
  ['1280x720', { width: 1280, height: 720 }, 1],
  ['1920x1080', { width: 1920, height: 1080 }, 1],
  ['1024-narrow', { width: 1024, height: 768 }, 1],
  ['zoom150', { width: 1280, height: 720 }, 1.5],
]) {
  await session(vp, dsf, async (page) => {
    await load(page, five.slice(0, 3));
    await shot(page, `11-${label}-files`);
    await merge(page);
    await shot(page, `12-${label}-merged`);
    if (label === '1024-narrow') {
      await page.evaluate(() => window.scrollTo(99999, 0));
      await shot(page, `13-${label}-merged-scrolled-right`);
    }
  });
}

// 20,000 rows x 30 columns
await session({ width: 1440, height: 900 }, 1, async (page) => {
  const big = await makeHostileFixtures(join(fxDir, 'big'), { big: true });
  const t0 = Date.now();
  await load(page, [big['big_base_20000x30.xlsx'], big['big_other1.xlsx'], big['big_other2.xlsx']]);
  await merge(page);
  console.log(`20,000 x 30 merged in ${Date.now() - t0} ms`);
  await shot(page, '14-big-20000x30');
});

// the server goes away while the export dialog is open
await session({ width: 1440, height: 900 }, 1, async (page) => {
  await load(page, five.slice(0, 2));
  await merge(page);
  await page.getByTestId('export-button').click();
  await page.route('**/api/**', (r) => r.abort('connectionrefused'));
  await page.evaluate(() => window.dispatchEvent(new CustomEvent('excel-merger-server', { detail: { connected: false } })));
  await page.waitForTimeout(300);
  await shot(page, '15-server-lost-export');
  await page.getByTestId('export-confirm').click();
  await page.getByTestId('export-result').waitFor({ timeout: 60000 });
  await shot(page, '16-server-lost-result');
});

await browser.close();
console.log(errors.length ? `console errors: ${JSON.stringify(errors)}` : 'no console errors');
if (failures.length || errors.length) process.exit(1);
