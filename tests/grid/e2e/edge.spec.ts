/**
 * Grid browser tests, part 3: edge cases — hidden rows (non-contiguous `rows`), file row heights (rowHeight),
 * an empty grid, no frozen columns, a very wide sheet.
 */
import { expect, test } from '@playwright/test';
import { address, calls, clearCalls, expectNoErrors, frames, openDemo } from './helpers';

test('hidden rows: `rows` may skip Excel rows; navigation and testids use Excel row numbers', async ({ page }) => {
  const errors = await openDemo(page, '?gaps=1');
  await expect(page.getByTestId('grid-rowhdr-9')).toBeVisible();
  await expect(page.getByTestId('grid-rowhdr-10')).toHaveCount(0);
  await expect(page.getByTestId('grid-rowhdr-11')).toBeVisible();
  await page.getByTestId('grid-cell-9-0').click();
  await page.keyboard.press('ArrowDown');
  expect(await address(page)).toBe('A11'); // row 10 is hidden
  await page.keyboard.press('ArrowUp');
  expect(await address(page)).toBe('A9');
  expect(await page.evaluate(() => window.__gridDemo!.grid()!.scrollToRow(20))).toBe(false);
  expect(await page.evaluate(() => window.__gridDemo!.grid()!.scrollToRow(21, { select: true }))).toBe(true);
  await frames(page, 3);
  expect(await address(page)).toBe('A21');
  await expectNoErrors(errors);
});

test('row heights from the file: tall rows are drawn taller and the active frame / editor follow', async ({ page }) => {
  const errors = await openDemo(page, '?tall=1');
  const c7 = (await page.getByTestId('grid-cell-7-3').boundingBox())!;
  const c8 = (await page.getByTestId('grid-cell-8-3').boundingBox())!;
  expect(Math.round(c7.height)).toBe(40);
  expect(Math.round(c8.height)).toBe(22);
  expect(Math.round(c8.y - c7.y)).toBe(40);
  await page.getByTestId('grid-cell-7-3').click();
  const frame = (await page.locator('.g-active').boundingBox())!;
  expect(Math.abs(frame.y - c7.y)).toBeLessThan(1);
  expect(Math.abs(frame.height - c7.height)).toBeLessThan(1);
  await page.keyboard.type('x');
  const ed = (await page.getByTestId('grid-editor').boundingBox())!;
  expect(Math.abs(ed.y - c7.y)).toBeLessThan(1);
  await page.keyboard.press('Escape');
  // jump far down: the virtualiser uses the variable heights
  await page.evaluate(() => window.__gridDemo!.grid()!.scrollToRow(9997, { select: true }));
  await frames(page, 4);
  await expect(page.getByTestId('grid-rowhdr-9997')).toBeVisible();
  const t = (await page.getByTestId('grid-rowhdr-9997').boundingBox())!;
  const fr = (await page.locator('.g-active').boundingBox())!;
  expect(Math.abs(fr.y - t.y)).toBeLessThan(1);
  await expectNoErrors(errors);
});

test('empty grid: message, no crash on keys and clicks', async ({ page }) => {
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  page.on('console', (m) => {
    if (m.type() === 'error' || m.type() === 'warning') errors.push(m.text());
  });
  await page.goto('/grid-demo.html?rows=0');
  await expect(page.locator('.g-empty')).toHaveText('表示する行がありません');
  await expect(page.getByTestId('grid-header-0')).toBeVisible();
  await page.getByTestId('grid-header-2').click();
  await page.keyboard.press('ArrowDown');
  await page.keyboard.type('abc');
  await page.keyboard.press('Enter');
  await page.keyboard.press('F2');
  await page.keyboard.press('Delete'); // deletes the selected column: allowed
  await page.keyboard.press('Control+z');
  expect((await calls(page, 'onCommitEdit')).length).toBe(0);
  await expect(page.getByTestId('grid-editor')).not.toHaveClass(/is-editing/);
  await expectNoErrors(errors);
});

test('no frozen columns: dragging to the left edge auto-scrolls back; drop at gap 0', async ({ page }) => {
  const errors = await openDemo(page, '?frozen=0');
  const sc = page.getByTestId('grid');
  await sc.evaluate((e) => (e.scrollLeft = 1500));
  await frames(page, 3);
  const box = (await sc.boundingBox())!;
  // a header fully right of the row-number gutter
  const title = await page.$$eval(
    '.g-header .g-hcell',
    (els, minX) => els.map((e) => ({ x: e.getBoundingClientRect().left, t: e.querySelector('.g-htitle')!.textContent ?? '' })).filter((o) => o.x > minX).sort((a, b) => a.x - b.x)[0].t,
    box.x + 150,
  );
  const hb = (await page.locator('.g-header .g-hcell', { has: page.locator('.g-htitle', { hasText: new RegExp(`^${title}$`) }) }).boundingBox())!;
  await clearCalls(page);
  await page.mouse.move(hb.x + 10, hb.y + 10);
  await page.mouse.down();
  await page.mouse.move(hb.x - 30, hb.y + 10, { steps: 3 });
  await page.mouse.move(box.x + 60, hb.y + 10, { steps: 8 }); // just right of the row-number gutter
  // auto-scroll runs at up to 32 px per frame: wait until it reaches the left edge (slow shared machine)
  await expect.poll(() => sc.evaluate((e) => e.scrollLeft), { timeout: 10_000 }).toBe(0);
  await page.mouse.move(box.x + 60, hb.y + 10);
  await page.waitForTimeout(100);
  await page.mouse.up();
  const mv = await calls(page, 'onMoveColumns');
  expect(mv).toHaveLength(1);
  expect(mv[0].args[1]).toBe(0);
  expect((await page.evaluate(() => window.__gridDemo!.titles()))[0]).toBe(title);
  await expectNoErrors(errors);
});

test('very wide sheet (400 columns) renders and scrolls to the last column', async ({ page }) => {
  const errors = await openDemo(page, '?cols=400&rows=2000');
  expect(await page.evaluate(() => window.__gridDemo!.titles().length)).toBe(400);
  await page.getByTestId('grid-cell-2-0').click();
  await page.keyboard.press('Control+ArrowRight');
  await frames(page, 4);
  expect(await address(page)).toBe('OJ2'); // column 400
  await expect(page.getByTestId('grid-header-399')).toBeVisible();
  expect(await page.locator('.g-cell').count()).toBeLessThan(3000);
  await expectNoErrors(errors);
});
