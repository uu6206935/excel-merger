/**
 * Scroll performance of the PRODUCTION build (single file, file://) with 10,000 rows × 60 columns.
 * Headless Edge rasterises in software, so the numbers are pessimistic (R3 §3.6: the idle rAF cadence is 16.7 ms;
 * the prototype measured median 16.7 / 20-23 / 20-22 / 30-36 ms for the four scenarios below).
 */
import { expect, test } from '@playwright/test';
import { execSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import fs from 'node:fs';

const root = fileURLToPath(new URL('../../..', import.meta.url));
const dist = fileURLToPath(new URL('../.demo-dist/grid-demo.html', import.meta.url));

test.beforeAll(() => {
  execSync('npx vite build --config tests/grid/vite.demo.config.ts', { cwd: root, stdio: 'pipe', timeout: 180_000 });
  expect(fs.existsSync(dist)).toBe(true);
});

test('production build: scrolling 10,000 × 60 stays smooth', async ({ page }) => {
  test.setTimeout(120_000);
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  await page.goto(pathToFileURL(dist).href);
  await page.waitForSelector('[data-testid="grid-cell-2-0"]');
  const run = (dy: number, dx: number, n = 120) =>
    page.evaluate(
      async ({ dy, dx, n }) => {
        const sc = document.querySelector('[data-testid=grid]') as HTMLElement;
        sc.scrollTop = 0;
        sc.scrollLeft = 0;
        await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(() => r(null))));
        const frames: number[] = [];
        let last = performance.now();
        for (let i = 0; i < n; i++) {
          sc.scrollTop += dy;
          sc.scrollLeft += dx;
          await new Promise((r) => requestAnimationFrame(() => r(null)));
          const now = performance.now();
          frames.push(now - last);
          last = now;
        }
        frames.sort((a, b) => a - b);
        const q = (p: number) => Math.round(frames[Math.min(frames.length - 1, Math.floor(frames.length * p))] * 10) / 10;
        return { median: q(0.5), p95: q(0.95), cells: document.querySelectorAll('.g-cell').length };
      },
      { dy, dx, n },
    );
  const v1 = await run(22, 0);
  const v8 = await run(180, 0);
  const hz = await run(0, 120, 40);
  const diag = await run(180, 40);
  console.log('perf (prod, headless sw raster):', JSON.stringify({ v1, v8, hz, diag }));
  expect(v1.median).toBeLessThan(25);
  expect(v8.median).toBeLessThan(35);
  expect(hz.median).toBeLessThan(40);
  expect(diag.median).toBeLessThan(45);
  expect(diag.cells).toBeLessThan(3000);
  expect(errors).toEqual([]);
});
