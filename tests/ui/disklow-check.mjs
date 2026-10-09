// Final round: browser check of the low-disk notice -> tests/ui/shots/final-disklow-*.png
//
//   Fake server (default): a tiny node http server serves the built UI plus a scripted GET /api/status (and just enough of the other
//   endpoints for a built-in-engine export), so every combination of diskLow / outputDiskLow can be shown without filling a drive.
//     npx vite build --outDir <scratch>/dist --emptyOutDir
//     UI_DIST=<scratch>/dist UI_PORT=5192 node tests/ui/disklow-check.mjs
//   Real server: the real server/server.mjs started with a huge EM_DISK_LOW_MB (so ITS OWN status says diskLow / outputDiskLow):
//     EM_DISK_LOW_MB=99999999 PORT=5192 DIST_DIR=<scratch>/dist OUTPUT_DIR=<scratch>/output EM_DATA_DIR=<scratch>/data NO_OPEN=1 node server/server.mjs
//     UI_REAL=http://localhost:5192/ UI_REAL_OUT=<scratch>/output node tests/ui/disklow-check.mjs
//
// Installed Microsoft Edge only (channel msedge), built-in engine only (no Excel is started). Exit code 1 when a check fails.
import { chromium } from '@playwright/test';
import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '..', '..');
const outDir = join(here, 'shots');
fs.mkdirSync(outDir, { recursive: true });
const samples = join(root, 'samples');
const files3 = ['01_学生名簿.xlsx', '02_面接評価.xlsx', '03_連絡先.csv'].map((n) => join(samples, n));
const PORT = Number(process.env.UI_PORT || 5192);
const DIST = resolve(process.env.UI_DIST || join(root, 'dist'));
const REAL = process.env.UI_REAL || '';

const failures = [];
const jsErrors = [];
let passed = 0;
function check(name, cond, extra = '') {
  if (cond) {
    passed++;
    console.log(`  ok   ${name}`);
  } else {
    failures.push(name);
    console.log(`  FAIL ${name} ${extra}`);
  }
}

// ------------------------------------------------------------------------------------------------ fake server

const MSG_C = (mb) => `C: ドライブの空き容量が少なくなっています（残り ${mb >= 1024 ? (mb / 1024).toFixed(1) + ' GB' : mb + ' MB'}）。Excel本体での出力は一時ファイルを作るため、空きが足りないと失敗することがあります。不要なファイルを削除するか、「内蔵エンジン」で出力してください。`;

async function startFake() {
  const out = resolve(process.env.UI_FAKE_OUT || join(path.dirname(DIST), 'fake-output'));
  fs.mkdirSync(out, { recursive: true });
  const st = { diskLow: false, outputDiskLow: false, delayMs: 0, hits: 0, freeMB: 1200, saved: [] };
  const statusBody = () => {
    const j = { app: 'excel-merger', version: '1.0.0-fake', excel: { installed: true, version: '16.0' }, outputDir: out, diskFreeMB: st.diskLow ? st.freeMB : 20000, diskLow: st.diskLow, outputDiskLow: st.outputDiskLow };
    if (st.diskLow) j.diskMessage = MSG_C(st.freeMB);
    return JSON.stringify(j);
  };
  const server = http.createServer((req, res) => {
    const url = new URL(req.url, 'http://localhost');
    const json = (obj, code = 200) => {
      res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' });
      res.end(typeof obj === 'string' ? obj : JSON.stringify(obj));
    };
    if (url.pathname === '/api/ping') return json({ app: 'excel-merger', version: '1.0.0-fake' });
    if (url.pathname === '/api/status') {
      st.hits++;
      const body = statusBody(); // the answer reflects the state at request time, like the real server
      return void setTimeout(() => json(body), st.delayMs);
    }
    if (url.pathname === '/api/events') {
      res.writeHead(200, { 'Content-Type': 'text/event-stream', 'Cache-Control': 'no-store', Connection: 'keep-alive' });
      res.write('retry: 1000\n\n');
      const t = setInterval(() => res.write(': keepalive\n\n'), 5000);
      req.on('close', () => clearInterval(t));
      return;
    }
    if (url.pathname === '/api/bye') return void res.writeHead(204).end();
    if (url.pathname === '/api/save' && req.method === 'POST') {
      const chunks = [];
      req.on('data', (c) => chunks.push(c));
      req.on('end', () => {
        const name = decodeURIComponent(String(req.headers['x-file-name'] || 'out.xlsx'));
        const p = join(out, name);
        fs.writeFileSync(p, Buffer.concat(chunks));
        st.saved.push(p);
        json({ path: p, name });
      });
      return;
    }
    if (url.pathname === '/api/open' || url.pathname === '/api/open-folder') return json({ ok: true });
    if (url.pathname === '/__set') {
      // test control: /__set?diskLow=1&outputDiskLow=0&delayMs=700&freeMB=800
      for (const [k, v] of url.searchParams) {
        if (k === 'diskLow' || k === 'outputDiskLow') st[k] = v === '1';
        else if (k === 'delayMs' || k === 'freeMB') st[k] = Number(v);
      }
      return json({ ok: true });
    }
    if (url.pathname === '/' || url.pathname === '/index.html') {
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' });
      return void res.end(fs.readFileSync(join(DIST, 'index.html')));
    }
    res.writeHead(404).end();
  });
  await new Promise((r, j) => server.once('error', j).listen(PORT, '127.0.0.1', r));
  const set = async (q) => {
    const r = await fetch(`http://localhost:${PORT}/__set?${new URLSearchParams(q)}`);
    await r.text();
  };
  return { st, set, out, close: () => new Promise((r) => (server.closeAllConnections?.(), server.close(r))), base: `http://localhost:${PORT}/` };
}

// ------------------------------------------------------------------------------------------------ browser helpers

const settle = async (page) => {
  await page.waitForFunction(() => !document.querySelector('[data-testid=busy-overlay]') && !document.querySelector('.card-loading'), null, { timeout: 120000 });
  await page.waitForTimeout(300);
};
const shot = async (page, name, opts = {}) => {
  await page.waitForTimeout(250);
  await page.screenshot({ path: join(outDir, `final-disklow-${name}.png`), ...opts });
  console.log(`  saved final-disklow-${name}.png`);
};
async function newPage(browser, base, viewport = { width: 1440, height: 900 }) {
  const ctx = await browser.newContext({ viewport, deviceScaleFactor: 1, locale: 'ja-JP', acceptDownloads: true });
  const page = await ctx.newPage();
  page.on('pageerror', (e) => jsErrors.push(`pageerror: ${e.message}`));
  page.on('console', (m) => m.type() === 'error' && !/Failed to load resource|ERR_/.test(m.text()) && jsErrors.push(m.text()));
  await page.goto(base);
  await page.getByTestId('server-status').waitFor();
  await page.waitForFunction(() => document.querySelector('[data-testid=server-status]')?.getAttribute('data-online') === 'true');
  return { ctx, page };
}
async function mergeFiles(page) {
  await page.getByTestId('file-input').setInputFiles(files3);
  await settle(page);
  await page.getByTestId('merge-button').click();
  await page.getByTestId('grid').waitFor({ timeout: 120000 });
  await settle(page);
}
const openDialog = async (page) => {
  await page.getByTestId('export-button').click();
  await page.getByTestId('export-dialog').waitFor();
};
const closeDialog = async (page) => {
  await page.keyboard.press('Escape');
  await page.getByTestId('export-dialog').waitFor({ state: 'detached' });
};
const checked = (page, id) => page.getByTestId(id).isChecked();
const chip = (page) => page.getByTestId('disk-low-chip');
const notice = (page) => page.getByTestId('export-disk-notice');
/** the dialog has asked /api/status again (fake server hit counter went up) and the UI had time to apply it */
async function waitRefreshed(fake, page, before) {
  const t0 = Date.now();
  while (fake.st.hits <= before && Date.now() - t0 < 10000) await page.waitForTimeout(50);
  await page.waitForTimeout(Math.max(250, fake.st.delayMs + 250));
}

// ------------------------------------------------------------------------------------------------ fake-server scenarios

async function runFake(browser) {
  const fake = await startFake();
  console.log(`fake server on ${fake.base} (dist ${DIST}, output ${fake.out})`);
  try {
    const { ctx, page } = await newPage(browser, fake.base);
    await mergeFiles(page);

    console.log('1. disk fine');
    check('no chip in the header', (await chip(page).count()) === 0);
    let hits = fake.st.hits;
    await openDialog(page);
    await waitRefreshed(fake, page, hits);
    check('opening the dialog asked /api/status again', fake.st.hits > hits);
    check('no notice', (await notice(page).count()) === 0);
    check('default engine = Excel本体', (await checked(page, 'export-engine-excel')) && !(await checked(page, 'export-engine-js')));
    await shot(page, '1-ok-dialog');
    await closeDialog(page);

    console.log('2. diskLow (without reloading the page)');
    await fake.set({ diskLow: '1', freeMB: '1200' });
    hits = fake.st.hits;
    await openDialog(page);
    await waitRefreshed(fake, page, hits);
    check('notice appears', await notice(page).isVisible());
    const msg = MSG_C(1200);
    check('notice carries the server message', ((await page.getByTestId('export-disk-temp').innerText()) || '').trim() === msg);
    check('notice says Excel本体 may not work and recommends the built-in engine', /Excel本体エンジンは.*使えないことがあります.*内蔵エンジン.*おすすめ/.test(await page.getByTestId('export-disk-advice').innerText()));
    check('no output-drive sentence', (await page.getByTestId('export-disk-output').count()) === 0);
    check('default engine = 内蔵エンジン', (await checked(page, 'export-engine-js')) && !(await checked(page, 'export-engine-excel')));
    check('Excel本体 can still be picked (enabled)', await page.getByTestId('export-engine-excel').isEnabled());
    const nBox = await notice(page).boundingBox();
    const eBox = await page.locator('[role=radiogroup]').boundingBox();
    const fBox = await page.getByTestId('export-filename').boundingBox();
    check('notice sits above the engine choice (and below the file name)', !!nBox && !!eBox && !!fBox && nBox.y + nBox.height <= eBox.y + 1 && nBox.y >= fBox.y + fBox.height);
    check('notice is amber', await notice(page).evaluate((el) => getComputedStyle(el).backgroundColor === 'rgb(255, 243, 205)'));
    check('chip appears in the header with the message as tooltip and aria-label', (await chip(page).count()) === 1 && (await chip(page).getAttribute('title')) === msg && (await chip(page).getAttribute('aria-label')) === msg && (await chip(page).innerText()) === '空き容量が少ない');
    await shot(page, '2-low-dialog');
    await page.getByTestId('export-engine-excel').check();
    check('the user can still pick Excel本体', await checked(page, 'export-engine-excel'));
    await shot(page, '2b-low-dialog-excel-picked');
    await closeDialog(page);
    check('chip stays after closing the dialog', await chip(page).isVisible());
    check('chip is not clipped', await chip(page).evaluate((el) => el.scrollWidth <= el.clientWidth + 1));
    await shot(page, '3-low-header', { clip: { x: 0, y: 0, width: 1440, height: 130 } });
    const statusBox = (await page.getByTestId('server-status').boundingBox()) ?? { x: 0, y: 0, width: 0, height: 0 };
    check('header status line stays on one row', statusBox.height < 40);

    console.log('3. outputDiskLow only');
    await fake.set({ diskLow: '0', outputDiskLow: '1' });
    hits = fake.st.hits;
    await openDialog(page);
    await waitRefreshed(fake, page, hits);
    check('notice shows the output drive sentence', /出力フォルダのドライブ.*空き容量が少なくなっています/.test(await page.getByTestId('export-disk-output').innerText()));
    check('no temp-drive sentence / advice', (await page.getByTestId('export-disk-temp').count()) === 0 && (await page.getByTestId('export-disk-advice').count()) === 0);
    check('default engine stays Excel本体', await checked(page, 'export-engine-excel'));
    check('chip still shown (output drive low)', (await chip(page).count()) === 1 && /出力フォルダのドライブ/.test((await chip(page).getAttribute('aria-label')) ?? ''));
    await shot(page, '4-output-low-dialog');
    await closeDialog(page);

    console.log('4. both low');
    await fake.set({ diskLow: '1', outputDiskLow: '1', freeMB: '800' });
    hits = fake.st.hits;
    await openDialog(page);
    await waitRefreshed(fake, page, hits);
    check('both sentences', (await page.getByTestId('export-disk-temp').count()) === 1 && (await page.getByTestId('export-disk-output').count()) === 1 && (await page.getByTestId('export-disk-temp').innerText()).includes('残り 800 MB'));
    check('built-in engine default', await checked(page, 'export-engine-js'));
    await shot(page, '5-both-dialog');
    await closeDialog(page);

    console.log('5. back to fine (notice and chip disappear without reloading)');
    await fake.set({ diskLow: '0', outputDiskLow: '0' });
    hits = fake.st.hits;
    await openDialog(page);
    await waitRefreshed(fake, page, hits);
    check('no notice', (await notice(page).count()) === 0);
    check('no chip', (await chip(page).count()) === 0);
    check('default engine = Excel本体 again', await checked(page, 'export-engine-excel'));
    await shot(page, '6-recovered-dialog');
    await closeDialog(page);

    console.log('6. a slow status answer: the default follows it, the user\'s own pick does not');
    await fake.set({ diskLow: '1', outputDiskLow: '0', freeMB: '1200', delayMs: '900' });
    hits = fake.st.hits;
    await openDialog(page);
    check('right after opening: old status (no notice, Excel本体 default)', (await notice(page).count()) === 0 && (await checked(page, 'export-engine-excel')));
    await waitRefreshed(fake, page, hits);
    check('after the answer: notice + built-in default (nobody picked)', (await notice(page).isVisible()) && (await checked(page, 'export-engine-js')));
    await closeDialog(page);
    await fake.set({ diskLow: '0', delayMs: '0' }); // header chip goes with the next refresh
    hits = fake.st.hits;
    await openDialog(page);
    await waitRefreshed(fake, page, hits);
    await closeDialog(page);
    await fake.set({ diskLow: '1', delayMs: '900' });
    hits = fake.st.hits;
    await openDialog(page);
    await page.getByTestId('export-engine-excel').click(); // the user clicks Excel本体 (already the shown default) before the answer arrives
    await waitRefreshed(fake, page, hits);
    check('user\'s pick (Excel本体) survives the late low-disk answer', (await notice(page).isVisible()) && (await checked(page, 'export-engine-excel')));
    await closeDialog(page);
    await fake.set({ delayMs: '0' });

    console.log('7. export in the low state: the default (内蔵エンジン) works end to end');
    hits = fake.st.hits;
    await openDialog(page);
    await waitRefreshed(fake, page, hits);
    check('built-in default again in a new dialog', await checked(page, 'export-engine-js'));
    await page.getByTestId('export-confirm').click();
    await page.getByTestId('export-result').waitFor({ timeout: 120000 });
    check('result says built-in engine', (await page.getByTestId('export-result-engine').innerText()).includes('内蔵エンジンで出力しました'));
    check('file saved by the (fake) server', fake.st.saved.length === 1 && fs.existsSync(fake.st.saved[0]) && fs.statSync(fake.st.saved[0]).size > 1000);
    await shot(page, '7-low-export-result');
    await page.getByTestId('export-close').click();
    await ctx.close();

    console.log('8. 1280x720');
    await fake.set({ diskLow: '1', outputDiskLow: '1', freeMB: '1200' });
    const small = await newPage(browser, fake.base, { width: 1280, height: 720 });
    await mergeFiles(small.page);
    await small.page.waitForTimeout(300);
    check('1280: no horizontal page scroll', await small.page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth));
    check('1280: chip fully visible', await chip(small.page).evaluate((el) => el.scrollWidth <= el.clientWidth + 1 && el.getBoundingClientRect().right <= window.innerWidth));
    await openDialog(small.page);
    await waitRefreshed(fake, small.page, 0);
    const dlg = await small.page.getByTestId('export-dialog').boundingBox();
    check('1280x720: dialog fits the window', !!dlg && dlg.y >= 0 && dlg.y + dlg.height <= 720);
    check('1280x720: 出力 button visible', await small.page.getByTestId('export-confirm').isVisible());
    await shot(small.page, '8-1280x720-both-dialog');
    await small.ctx.close();
  } finally {
    await fake.close();
  }
}

// ------------------------------------------------------------------------------------------------ real-server scenario

async function runReal(browser) {
  const base = REAL;
  const realOut = process.env.UI_REAL_OUT ? resolve(process.env.UI_REAL_OUT) : '';
  console.log(`real server ${base}`);
  const st = await (await fetch(new URL('api/status', base))).json();
  console.log(`  /api/status: diskLow=${st.diskLow} diskFreeMB=${st.diskFreeMB} outputDiskLow=${st.outputDiskLow}`);
  console.log(`  diskMessage: ${st.diskMessage}`);
  check('the real server reports diskLow (started with a huge EM_DISK_LOW_MB)', st.diskLow === true && typeof st.diskMessage === 'string' && st.diskMessage.length > 10);
  // the free space in the message moves between two requests (a real drive): compare without the number
  const noNum = (s) => s.trim().replace(/残り [\d.]+ [GM]B/, '残り _');
  const { ctx, page } = await newPage(browser, base);
  check('chip in the header', (await chip(page).count()) === 1);
  check(
    'chip tooltip / aria-label start with the real diskMessage',
    noNum((await chip(page).getAttribute('title')) ?? '').startsWith(noNum(st.diskMessage)) && noNum((await chip(page).getAttribute('aria-label')) ?? '').startsWith(noNum(st.diskMessage)),
  );
  await shot(page, 'real-1-header', { clip: { x: 0, y: 0, width: 1440, height: 130 } });
  await mergeFiles(page);
  await openDialog(page);
  await page.waitForTimeout(800);
  check('notice with the real message', noNum(await page.getByTestId('export-disk-temp').innerText()) === noNum(st.diskMessage));
  check('output-drive sentence iff the real server says so', (await page.getByTestId('export-disk-output').count()) === (st.outputDiskLow ? 1 : 0));
  check('default engine = 内蔵エンジン; Excel本体 enabled', (await checked(page, 'export-engine-js')) && (await page.getByTestId('export-engine-excel').isEnabled()) === !!st.excel?.installed);
  await shot(page, 'real-2-dialog');
  await page.getByTestId('export-confirm').click();
  await page.getByTestId('export-result').waitFor({ timeout: 120000 });
  check('built-in engine export saved by the real server', (await page.getByTestId('export-result-engine').innerText()).includes('内蔵エンジンで出力しました'));
  const saved = await page.locator('.result-path').innerText();
  console.log(`  saved: ${saved}`);
  check('saved file exists in the output folder', fs.existsSync(saved) && (!realOut || path.resolve(saved).toLowerCase().startsWith(realOut.toLowerCase())));
  await shot(page, 'real-3-result');
  await ctx.close();
}

// ------------------------------------------------------------------------------------------------

const browser = await chromium.launch({ channel: 'msedge', headless: true });
try {
  if (REAL) await runReal(browser);
  else await runFake(browser);
} catch (e) {
  failures.push(`exception: ${e?.stack ?? e}`);
  console.log('EXCEPTION', e?.stack ?? e);
}
await browser.close();
check('no JS errors in the page', jsErrors.length === 0, JSON.stringify(jsErrors.slice(0, 3)));
console.log(`\n${passed} checks passed, ${failures.length} failed`);
if (failures.length) {
  console.log(failures.map((f) => ` - ${f}`).join('\n'));
  process.exit(1);
}
