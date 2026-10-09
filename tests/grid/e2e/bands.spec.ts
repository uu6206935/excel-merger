/**
 * Grid browser tests, round 7: the source bands above the headers (one band = a RUN of contiguous columns of the same file).
 * True counts in the labels, click / Ctrl+click / Shift+click selection, drag of a whole run (same gesture as a header drag),
 * and the GridHandle method scrollToColumnsCentered.
 *
 * Demo data (src/grid-demo.tsx): 60 columns, frozen 2. Run 0 = 学生名簿.xlsx（基準）8 columns (ID … 区分), run 1 = 成績一覧.xlsx 28
 * columns (科目1 … 順位（未計算）), run 2 = 連絡先.xlsx 24 columns.
 */
import { expect, test, type Page } from '@playwright/test';
import path from 'node:path';
import { SHOTS, calls, clearCalls, expectNoErrors, frames, header, history, idOf, openDemo, selectedTitles, titles } from './helpers';

const status = (page: Page) => page.getByTestId('demo-status').innerText();
const selCount = async (page: Page) => Number(/選択 (\d+)列/.exec(await status(page))?.[1] ?? NaN);
const gridBox = async (page: Page) => (await page.getByTestId('grid').boundingBox())!;
const scrollLeft = (page: Page) => page.getByTestId('grid').evaluate((e) => e.scrollLeft);

/** screen x of the middle of the columns with these titles (all must be rendered) */
async function centreOf(page: Page, names: string[]): Promise<number> {
  const first = (await header(page, names[0]).boundingBox())!;
  const last = (await header(page, names[names.length - 1]).boundingBox())!;
  return (first.x + last.x + last.width) / 2;
}

/** the middle of the visible, scrolling area of the grid (row-number gutter 56 + frozen pane excluded) in screen coordinates */
async function visibleCentre(page: Page, frozenW = 180): Promise<number> {
  const g = await gridBox(page);
  const cw = await page.getByTestId('grid').evaluate((e) => e.clientWidth);
  return g.x + (56 + frozenW + cw) / 2;
}

async function scrollTo(page: Page, left: number, top = 0) {
  await page.evaluate(([l, t]) => {
    const sc = document.querySelector('[data-testid=grid]') as HTMLElement;
    sc.scrollLeft = l;
    sc.scrollTop = t;
  }, [left, top]);
  await frames(page, 3);
}

test('band labels count ALL columns of the run (also the virtualised ones); a run cut by the frozen pane keeps its true count in both pieces', async ({ page }) => {
  const errors = await openDemo(page);
  // run 0 has 8 columns, 2 of them frozen: the pinned piece used to say 「2列」
  await expect(page.getByTestId('grid-band-0-pin')).toContainText('8列');
  await expect(page.getByTestId('grid-band-0')).toContainText('8列');
  await expect(page.getByTestId('grid-band-0-pin')).toHaveAttribute('title', /学生名簿\.xlsx（基準）（8列）\nドラッグで移動／クリックで選択/);
  // run 1: 28 columns, most of them far outside the 1400 px viewport
  await expect(page.getByTestId('grid-band-1')).toContainText('28列');
  await expect(page.getByTestId('grid-band-1')).toHaveAttribute('title', /成績一覧\.xlsx（28列）/);
  expect(await page.locator('.g-hcell').count()).toBeLessThan(40); // virtualised: the 28 are not all rendered
  await scrollTo(page, 3000);
  await expect(page.getByTestId('grid-band-2')).toContainText('24列');
  await expect(page.getByTestId('grid-band-1')).toContainText('28列'); // the label sticks / still true while scrolled
  await expectNoErrors(errors);
});

test('a band is a handle: grab cursor and the tooltip 「ドラッグで移動／クリックで選択」', async ({ page }) => {
  const errors = await openDemo(page);
  const band = page.getByTestId('grid-band-1');
  expect(await band.evaluate((e) => getComputedStyle(e).cursor)).toBe('grab');
  await expect(band).toHaveAttribute('title', /ドラッグで移動／クリックで選択$/);
  await expectNoErrors(errors);
});

test('click on a band selects ALL columns of the run (anchor = the first), the header selection style, the first column becomes active', async ({ page }) => {
  const errors = await openDemo(page);
  await page.getByTestId('grid-band-1').click({ position: { x: 40, y: 8 } });
  expect(await selCount(page)).toBe(28);
  await expect(page.getByTestId('grid-band-1')).toHaveAttribute('aria-pressed', 'true');
  await expect(page.getByTestId('grid-band-1')).toHaveClass(/is-sel/);
  // the rendered header cells of the run carry the very same class as a header selection
  const picked = await selectedTitles(page);
  expect(picked.length).toBeGreaterThan(3);
  expect(picked.every((t) => /^科目|^平均点|^奨学金|^出席率|^順位/.test(t ?? ''))).toBe(true);
  await expect(page.getByTestId('formula-bar-address')).toHaveText('I2'); // column 8 (I), first data row 2
  // another band replaces the selection
  await page.getByTestId('grid-band-0-pin').click({ position: { x: 40, y: 8 } });
  expect(await selCount(page)).toBe(8);
  await expect(page.getByTestId('grid-band-1')).toHaveAttribute('aria-pressed', 'false');
  await expect(page.getByTestId('grid-band-0-pin')).toHaveAttribute('aria-pressed', 'true');
  await expect(page.getByTestId('grid-band-0')).toHaveAttribute('aria-pressed', 'true'); // the second piece of the same run
  // a plain click on a header afterwards behaves as ever
  await header(page, '大学').click();
  expect(await selectedTitles(page)).toEqual(['大学']);
  await page.screenshot({ path: path.join(SHOTS, '08-band-selected.png') });
  await expectNoErrors(errors);
});

test('Ctrl+click toggles a whole run; Shift+click extends from the anchor to the run like a header Shift+click', async ({ page }) => {
  const errors = await openDemo(page);
  await page.getByTestId('grid-band-0-pin').click({ position: { x: 40, y: 8 } });
  await page.getByTestId('grid-band-1').click({ position: { x: 40, y: 8 }, modifiers: ['Control'] });
  expect(await selCount(page)).toBe(36);
  await page.getByTestId('grid-band-1').click({ position: { x: 40, y: 8 }, modifiers: ['Control'] });
  expect(await selCount(page)).toBe(8);
  // Shift: anchor = 大学 (header click), then the run 1 band -> 大学 .. 順位（未計算）
  await header(page, '大学').click();
  await page.getByTestId('grid-band-1').click({ position: { x: 40, y: 8 }, modifiers: ['Shift'] });
  expect(await selCount(page)).toBe(33); // 大学 (index 3) .. 順位（未計算） (index 35)
  // Ctrl+Shift adds a range
  await header(page, '氏名カナ').click();
  await page.getByTestId('grid-band-1').click({ position: { x: 40, y: 8 }, modifiers: ['Control', 'Shift'] });
  expect(await selCount(page)).toBe(34); // 氏名カナ (2) .. 35
  await expectNoErrors(errors);
});

test('drag a whole run: ghost with the file name, insertion indicator, ONE move callback with all its columns, undo', async ({ page }) => {
  const errors = await openDemo(page);
  const initial = await titles(page);
  const runIds: string[] = [];
  for (const t of initial.slice(8, 36)) runIds.push(await idOf(page, t));
  await clearCalls(page);
  const band = (await page.getByTestId('grid-band-1').boundingBox())!;
  const uni = (await header(page, '大学').boundingBox())!;
  await page.mouse.move(band.x + 60, band.y + 8);
  await page.mouse.down();
  await page.mouse.move(band.x + 40, band.y + 14, { steps: 3 });
  await page.mouse.move(uni.x + 15, uni.y + 150, { steps: 14 }); // left half of 大学 -> gap 3
  await page.waitForTimeout(80);
  await expect(page.locator('.g-ghost')).toHaveText('成績一覧.xlsx（28列）を移動');
  const ind = (await page.getByTestId('grid-drop-indicator').boundingBox())!;
  expect(Math.abs(ind.x + 2 - uni.x)).toBeLessThan(4);
  expect(await page.locator('.g-hcell.is-dragging').count()).toBeGreaterThan(0);
  await page.screenshot({ path: path.join(SHOTS, '09-band-drag.png') });
  await page.mouse.up();
  const mv = await calls(page, 'onMoveColumns');
  expect(mv).toHaveLength(1);
  expect(mv[0].args[0]).toEqual(runIds);
  expect(mv[0].args[1]).toBe(3);
  const after = await titles(page);
  expect(after.slice(0, 3)).toEqual(initial.slice(0, 3));
  expect(after.slice(3, 31)).toEqual(initial.slice(8, 36)); // the run moved as one block, in its own order
  expect(after.slice(31, 36)).toEqual(initial.slice(3, 8));
  expect((await history(page)).past).toBe(1);
  // the label now sits at the new place
  // the base columns are two runs now (ID 氏名 氏名カナ | 大学 …): the pinned band counts the TRUE size of its run, 3
  await expect(page.getByTestId('grid-band-0-pin')).toContainText('3列');
  await page.keyboard.press('Control+z');
  expect(await titles(page)).toEqual(initial);
  await expectNoErrors(errors);
});

test('dropping a run inside itself (or releasing far outside the grid) is a no-op; Esc cancels a running band drag', async ({ page }) => {
  const errors = await openDemo(page);
  const grid = await gridBox(page);
  await clearCalls(page);
  const band = (await page.getByTestId('grid-band-1').boundingBox())!;
  // inside the run
  const k5 = (await header(page, '科目5').boundingBox())!;
  await page.mouse.move(band.x + 60, band.y + 8);
  await page.mouse.down();
  await page.mouse.move(band.x + 40, band.y + 14, { steps: 3 });
  await page.mouse.move(k5.x + 30, k5.y + 20, { steps: 8 });
  await page.waitForTimeout(50);
  await page.mouse.up();
  expect(await calls(page, 'onMoveColumns')).toHaveLength(0);
  expect((await history(page)).past).toBe(0);
  // far outside
  await page.mouse.move(band.x + 60, band.y + 8);
  await page.mouse.down();
  await page.mouse.move(grid.x + grid.width + 250, band.y + 8, { steps: 8 });
  await page.waitForTimeout(80);
  await expect(page.getByTestId('grid-drop-indicator')).toHaveCount(0);
  await page.mouse.up();
  expect(await calls(page, 'onMoveColumns')).toHaveLength(0);
  // Esc (the held button out there auto-scrolled the grid: back to the start, measure again)
  await scrollTo(page, 0);
  const band3 = (await page.getByTestId('grid-band-1').boundingBox())!;
  const uni = (await header(page, '大学').boundingBox())!;
  await page.mouse.move(band3.x + 60, band3.y + 8);
  await page.mouse.down();
  await page.mouse.move(band3.x + 40, band3.y + 14, { steps: 3 });
  await page.mouse.move(uni.x + 15, uni.y + 100, { steps: 8 });
  await expect(page.getByTestId('grid-drop-indicator')).toHaveCount(1);
  await page.keyboard.press('Escape');
  await expect(page.getByTestId('grid-drop-indicator')).toHaveCount(0);
  await page.mouse.up();
  expect(await calls(page, 'onMoveColumns')).toHaveLength(0);
  expect((await history(page)).past).toBe(0);
  // the click part of that gesture (selection) happened, nothing else
  expect(await selCount(page)).toBe(28);
  await expectNoErrors(errors);
});

test('the pinned piece of a run cut by the frozen pane drags the WHOLE run (all 8 columns), dropped after 科目3', async ({ page }) => {
  const errors = await openDemo(page);
  const initial = await titles(page);
  await clearCalls(page);
  const pin = (await page.getByTestId('grid-band-0-pin').boundingBox())!;
  const k3 = (await header(page, '科目3').boundingBox())!;
  const k4 = (await header(page, '科目4').boundingBox())!;
  await page.mouse.move(pin.x + 80, pin.y + 8);
  await page.mouse.down();
  await page.mouse.move(pin.x + 100, pin.y + 14, { steps: 3 });
  await page.mouse.move((k3.x + k3.width + k4.x) / 2 + 6, k3.y + 120, { steps: 14 }); // boundary 科目3 | 科目4 -> gap 11
  await page.waitForTimeout(80);
  await page.mouse.up();
  const mv = await calls(page, 'onMoveColumns');
  expect(mv).toHaveLength(1);
  expect((mv[0].args[0] as string[]).length).toBe(8);
  expect(mv[0].args[1]).toBe(11);
  const after = await titles(page);
  expect(after.slice(0, 3)).toEqual(['科目1', '科目2', '科目3']);
  expect(after.slice(3, 11)).toEqual(initial.slice(0, 8));
  await expectNoErrors(errors);
});

test('right-click on a band selects the run and opens the column menu for it', async ({ page }) => {
  const errors = await openDemo(page);
  await page.getByTestId('grid-band-1').click({ button: 'right', position: { x: 40, y: 8 } });
  const menu = page.getByTestId('context-menu');
  await expect(menu).toBeVisible();
  expect(await selCount(page)).toBe(28);
  await expect(page.locator('[data-testid=context-menu-item-delete] .g-menu-hint')).toHaveAttribute('data-hint', '28列');
  await page.keyboard.press('Escape');
  await expectNoErrors(errors);
});

// ---------------------------------------------------------------------------------------------------- scrollToColumnsCentered

test.describe('GridHandle.scrollToColumnsCentered', () => {
  const call = (page: Page, names: string[], opts?: { select?: boolean; behavior?: 'smooth' | 'instant' }) =>
    page.evaluate(
      ([ids, o]) => window.__gridDemo!.grid()!.scrollToColumnsCentered(ids as string[], o as { select?: boolean; behavior?: 'smooth' | 'instant' } | undefined),
      [names, opts] as const,
    );
  const ids = async (page: Page, names: string[]) => Promise.all(names.map((n) => idOf(page, n)));

  test('puts a group of columns into the middle of the visible area and scrolls to the top', async ({ page }) => {
    const errors = await openDemo(page);
    const names = ['科目10', '科目11', '科目12'];
    await scrollTo(page, 0, 4000); // far down
    expect(await call(page, await ids(page, names), { behavior: 'instant' })).toBe(true);
    await frames(page, 4);
    expect(await page.getByTestId('grid').evaluate((e) => e.scrollTop)).toBe(0);
    const c = await centreOf(page, names);
    const v = await visibleCentre(page);
    expect(Math.abs(c - v)).toBeLessThanOrEqual(2);
    const g = await gridBox(page);
    expect(Math.abs(c - (g.x + g.width / 2))).toBeLessThan(g.width * 0.15); // and near the middle of the whole viewport
    await expectNoErrors(errors);
  });

  test('the LAST columns can be centred too (blank room behind the last column)', async ({ page }) => {
    const errors = await openDemo(page);
    const all = await titles(page);
    const names = all.slice(-3);
    const w0 = await page.getByTestId('grid').evaluate((e) => e.scrollWidth);
    await call(page, await ids(page, names), { behavior: 'instant' });
    await frames(page, 4);
    expect(await page.getByTestId('grid').evaluate((e) => e.scrollWidth)).toBeGreaterThan(w0);
    const c = await centreOf(page, names);
    expect(Math.abs(c - (await visibleCentre(page)))).toBeLessThanOrEqual(2);
    await expectNoErrors(errors);
  });

  test('a group inside the frozen pane needs no scrolling; a group wider than the viewport starts at the left edge of the scrolling area', async ({ page }) => {
    const errors = await openDemo(page);
    await scrollTo(page, 500);
    const before = await scrollLeft(page);
    expect(await call(page, await ids(page, ['ID', '氏名']), { behavior: 'instant' })).toBe(true);
    await frames(page, 3);
    expect(await scrollLeft(page)).toBe(before);
    // the 28 columns of run 1 are ~1,900 px wide
    const run1 = (await titles(page)).slice(8, 36);
    await call(page, await ids(page, run1), { behavior: 'instant' });
    await frames(page, 4);
    const first = (await header(page, '科目1').boundingBox())!;
    const g = await gridBox(page);
    const left = first.x - g.x;
    expect(left).toBeGreaterThanOrEqual(236 - 1);
    expect(left).toBeLessThan(236 + 25);
    await expectNoErrors(errors);
  });

  test('columns split over several places: the first contiguous run is centred; select:true selects all of them and makes the first one active', async ({ page }) => {
    const errors = await openDemo(page);
    const names = ['科目20', '科目21', '追加項目5'];
    expect(await call(page, await ids(page, names), { behavior: 'instant', select: true })).toBe(true);
    await frames(page, 4);
    const c = await centreOf(page, ['科目20', '科目21']);
    expect(Math.abs(c - (await visibleCentre(page)))).toBeLessThanOrEqual(2);
    expect(await selCount(page)).toBe(3);
    await expect(page.getByTestId('formula-bar-address')).toHaveText('AB2'); // 科目20 = index 27 -> AB
    expect(await call(page, ['no-such-column'], { behavior: 'instant' })).toBe(false);
    await expectNoErrors(errors);
  });

  test('smooth by default, instant with prefers-reduced-motion', async ({ page }) => {
    const errors = await openDemo(page);
    const target = await ids(page, ['科目14', '科目15']);
    // default behaviour: one frame after the call the scroll is under way, it has not arrived yet
    const r = await page.evaluate(async (t) => {
      const sc = document.querySelector('[data-testid=grid]') as HTMLElement;
      window.__gridDemo!.grid()!.scrollToColumnsCentered(t);
      await new Promise((res) => requestAnimationFrame(() => res(null)));
      const early = sc.scrollLeft;
      await new Promise((res) => setTimeout(res, 1800));
      return { early, late: sc.scrollLeft };
    }, target);
    expect(r.late).toBeGreaterThan(r.early + 100);
    expect(Math.abs((await centreOf(page, ['科目14', '科目15'])) - (await visibleCentre(page)))).toBeLessThanOrEqual(2);
    // reduced motion: there at once
    await scrollTo(page, 0);
    await page.emulateMedia({ reducedMotion: 'reduce' });
    await call(page, await ids(page, ['科目14', '科目15']));
    await frames(page, 2);
    expect(Math.abs((await centreOf(page, ['科目14', '科目15'])) - (await visibleCentre(page)))).toBeLessThanOrEqual(2);
    await expectNoErrors(errors);
  });
});
