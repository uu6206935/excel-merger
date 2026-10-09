// UI performance check with 10,000 rows x 40 columns (?mock=1 fakes; real modules are measured by their owners).
//   node tests/ui/perf.mjs [baseUrl]
import { chromium } from '@playwright/test';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const base = process.argv[2] ?? 'http://127.0.0.1:5174/';
const browser = await chromium.launch({ channel: 'msedge', headless: true });
const page = await (await browser.newContext({ viewport: { width: 1440, height: 900 } })).newPage();
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
page.on('console', (m) => m.type() === 'error' && errors.push(m.text()));

const t0 = Date.now();
await page.goto(base + '?mock=1&demo=merged&rows=10000&cols=40&delay=0.2');
await page.waitForSelector('[data-testid=grid]', { timeout: 120000 });
await page.waitForFunction(() => !document.querySelector('[data-testid=busy-overlay]'), null, { timeout: 120000 });
console.log(`load 3 files + analyse + merge (10000 x 40): ${Date.now() - t0} ms`);

const h = (i) => page.locator(`[data-testid=grid-header-${i}]`);
const titleOf = (i) => h(i).innerText();

// scroll: long-task time while scrolling the grid down and right
const scroll = await page.evaluate(async () => {
  const el = document.querySelector('[data-testid=grid]');
  const frames = [];
  let last = performance.now();
  for (let i = 0; i < 60; i++) {
    el.scrollTop += 900;
    el.scrollLeft += i % 3 === 0 ? 300 : 0;
    await new Promise((r) => requestAnimationFrame(r));
    const now = performance.now();
    frames.push(now - last);
    last = now;
  }
  frames.sort((a, b) => a - b);
  return { median: frames[Math.floor(frames.length / 2)], p95: frames[Math.floor(frames.length * 0.95)], max: frames[frames.length - 1] };
});
console.log(`scroll frames: median ${scroll.median.toFixed(1)} ms, p95 ${scroll.p95.toFixed(1)} ms, max ${scroll.max.toFixed(1)} ms`);
await page.evaluate(() => {
  const el = document.querySelector('[data-testid=grid]');
  el.scrollTop = 0;
  el.scrollLeft = 0;
});
await page.waitForTimeout(200);

async function timed(label, action, done, arg = null) {
  const t = Date.now();
  await action();
  await page.waitForFunction(done, arg, { timeout: 60000 });
  console.log(`${label}: ${Date.now() - t} ms`);
}

// move column 5 before column 1 via drag
const before = await titleOf(5);
const from = await h(5).boundingBox();
const to = await h(1).boundingBox();
await page.mouse.move(from.x + from.width / 2, from.y + 10);
await page.mouse.down();
await page.mouse.move(to.x + 4, to.y + 10, { steps: 8 });
await timed(
  'move 1 column (drop -> redraw)',
  () => page.mouse.up(),
  (t) => document.querySelector('[data-testid=grid-header-1]')?.textContent?.includes(t),
  before.split(/\n/).pop().trim(),
);

const count = async () => (await page.locator('[data-testid=side-summary]').innerText()).match(/変更 (\d+) 件/)?.[1];
await timed('undo', () => page.click('[data-testid=undo-button]'), () => document.querySelector('[data-testid=side-summary]')?.textContent?.includes('変更 0 件'));
await timed('redo', () => page.click('[data-testid=redo-button]'), () => document.querySelector('[data-testid=side-summary]')?.textContent?.includes('変更 1 件'));
console.log('edits:', await count());

// insert via context menu
await h(2).click({ button: 'right' });
await timed(
  'insert column (menu -> redraw)',
  () => page.click('[data-testid=context-menu-item-insert-left]'),
  () => document.querySelector('[data-testid=grid-header-2]')?.textContent?.includes('（C列）'),
);

// delete a column that 10,000 formulas depend on -> dialog
const idx = await page.evaluate(() => {
  const hs = [...document.querySelectorAll('[data-testid^=grid-header-]')];
  const hit = hs.find((e) => e.textContent?.includes('筆記'));
  return hit ? Number(hit.getAttribute('data-testid').slice('grid-header-'.length)) : -1;
});
if (idx >= 0) {
  await h(idx).click();
  await timed('analyse delete (10,000 dependents) -> dialog', () => page.keyboard.press('Delete'), () => !!document.querySelector('[data-testid=delete-confirm-dialog]'));
  await timed('convert + delete', () => page.click('[data-testid=delete-convert-button]'), () => !document.querySelector('[data-testid=delete-confirm-dialog]'));
}
await page.screenshot({ path: join(here, 'shots', 'perf-10000x40.png') });
console.log(errors.length ? 'errors:\n' + errors.join('\n') : 'no errors');
await browser.close();
