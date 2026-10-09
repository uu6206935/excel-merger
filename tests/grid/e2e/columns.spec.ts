/**
 * Grid browser tests, part 1: rendering, header selection, column drag (non-contiguous), auto-scroll, resize,
 * context menu, Delete key, undo/redo callbacks.
 */
import { expect, test } from '@playwright/test';
import path from 'node:path';
import {
  SHOTS,
  calls,
  clearCalls,
  expectNoErrors,
  frames,
  header,
  history,
  idOf,
  indexOfTitle,
  openDemo,
  reveal,
  selectedTitles,
  titles,
  widthOf,
} from './helpers';

test('renders virtualised 10,000 × 60 with Excel formats, row numbers, bands, key badges, highlight, formula text', async ({ page }) => {
  const errors = await openDemo(page);
  const cellCount = await page.locator('.g-cell').count();
  expect(cellCount).toBeGreaterThan(100);
  expect(cellCount).toBeLessThan(3000); // 600,000 cells exist
  await expect(page.getByTestId('grid')).toHaveAttribute('role', 'grid');
  await expect(page.getByTestId('grid-header-0')).toContainText('ID');
  await expect(page.getByTestId('grid-header-0')).toContainText('A');
  await expect(page.getByTestId('grid-header-0').locator('.g-keybadge')).toHaveText('キー:ID');
  await expect(page.getByTestId('grid-header-1').locator('.g-keybadge')).toHaveText('キー:氏名');
  await expect(page.getByTestId('grid-header-0')).toHaveAttribute('title', /ファイル: 学生名簿\.xlsx（基準）/);
  // first data row is Excel row 2 (header row 1 is the column title)
  await expect(page.getByTestId('grid-rowhdr-2')).toHaveText('2');
  await expect(page.locator('.g-corner-row')).toHaveText('1');
  await expect(page.getByTestId('grid-cell-2-0')).toHaveText('100001');
  // number formats: 和暦 / yyyy/m/d
  const row2 = await page.locator('.g-row[aria-rowindex="2"]').textContent();
  expect(row2).toMatch(/平成\d+年\d+月\d+日/);
  expect(row2).toMatch(/\d{4}\/\d+\/\d+/);
  // source band with the file name and tooltip
  await expect(page.locator('.g-band').first()).toHaveAttribute('title', /学生名簿\.xlsx/);
  // highlighted (unmatched) rows: Excel rows divisible by 97
  await page.evaluate(() => window.__gridDemo!.grid()!.scrollToRow(97));
  await frames(page, 3);
  await expect(page.locator('.g-row.is-hl').first()).toBeVisible();
  await expect(page.getByTestId('grid-rowhdr-97')).toHaveAttribute('title', /未一致/);
  // a formula without a cached value is shown as grey italic formula text
  const i = await indexOfTitle(page, '順位（未計算）');
  await reveal(page, '順位（未計算）');
  const f = page.getByTestId(`grid-cell-97-${i}`);
  await expect(f).toHaveClass(/is-ftext/);
  await expect(f).toHaveText(/^=RANK\([A-Z]+97,/);
  expect(await f.evaluate((e) => getComputedStyle(e).fontStyle)).toBe('italic');
  // hidden column drawn hatched
  await reveal(page, '内部メモ');
  await expect(header(page, '内部メモ')).toHaveClass(/is-hiddencol/);
  await page.evaluate(() => {
    const sc = document.querySelector('[data-testid=grid]')!;
    sc.scrollLeft = 0;
    sc.scrollTop = 0;
  });
  await frames(page, 3);
  await page.screenshot({ path: path.join(SHOTS, '01-initial.png') });
  await expectNoErrors(errors);
});

test('scrolling 10,000 × 60 stays smooth; jump to the last row', async ({ page }) => {
  const errors = await openDemo(page);
  const perf = await page.evaluate(async () => {
    const sc = document.querySelector('[data-testid=grid]') as HTMLElement;
    const frames: number[] = [];
    let last = performance.now();
    for (let i = 0; i < 120; i++) {
      sc.scrollTop += 180; // ~8 rows per frame
      if (i % 3 === 0) sc.scrollLeft += 120;
      await new Promise((r) => requestAnimationFrame(() => r(null)));
      const now = performance.now();
      frames.push(now - last);
      last = now;
    }
    frames.sort((a, b) => a - b);
    return { median: frames[60], p95: frames[114], cells: document.querySelectorAll('.g-cell').length };
  });
  // dev build (React development mode + StrictMode) on a shared machine: only a sanity bound here;
  // the real frame-time budget is asserted on the production build in perf.spec.ts
  console.log('scroll frames (dev build, headless):', JSON.stringify(perf));
  expect(perf.median).toBeLessThan(150);
  expect(perf.cells).toBeLessThan(3000);
  const jump = await page.evaluate(async () => {
    const sc = document.querySelector('[data-testid=grid]') as HTMLElement;
    sc.scrollTop = sc.scrollHeight;
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(() => r(null))));
    return Math.max(...[...document.querySelectorAll('.g-rowhdr')].map((e) => Number(e.textContent)));
  });
  expect(jump).toBe(10001);
  await expectNoErrors(errors);
});

test('header click / Ctrl+click toggle / Shift+click range / Ctrl+Shift+click add range', async ({ page }) => {
  const errors = await openDemo(page);
  await header(page, '氏名カナ').click();
  await header(page, '学籍番号').click({ modifiers: ['Control'] });
  await header(page, '区分').click({ modifiers: ['Control'] });
  expect(await selectedTitles(page)).toEqual(['氏名カナ', '学籍番号', '区分']);
  await expect(page.getByTestId('demo-status')).toContainText('選択 3列');
  await header(page, '学籍番号').click({ modifiers: ['Control'] });
  expect(await selectedTitles(page)).toEqual(['氏名カナ', '区分']);
  await header(page, '大学').click();
  await header(page, '入学日').click({ modifiers: ['Shift'] });
  expect(await selectedTitles(page)).toEqual(['大学', '学籍番号', '生年月日', '入学日']);
  await header(page, '科目2').click({ modifiers: ['Control', 'Shift'] });
  expect((await selectedTitles(page)).length).toBe(7); // 大学..科目2 (anchor 大学)
  // header click makes the column active (formula bar address follows)
  await header(page, '氏名カナ').click();
  await expect(page.getByTestId('formula-bar-address')).toHaveText('C2');
  // Esc clears the column selection; a cell click too
  await page.keyboard.press('Escape');
  expect(await selectedTitles(page)).toEqual([]);
  await header(page, '大学').click();
  await page.getByTestId('grid-cell-3-5').click();
  expect(await selectedTitles(page)).toEqual([]);
  // keyboard: Ctrl+Space selects the active column, Shift+Right extends
  await page.keyboard.press('Control+Space');
  expect(await selectedTitles(page)).toEqual(['生年月日']);
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowRight');
  expect(await selectedTitles(page)).toEqual(['生年月日', '入学日', '区分']);
  await expectNoErrors(errors);
});

test('drag non-contiguous columns: indicator, ghost, exact order, callbacks, undo/redo', async ({ page }) => {
  const errors = await openDemo(page);
  const initial = await titles(page);
  await header(page, '氏名カナ').click();
  await header(page, '学籍番号').click({ modifiers: ['Control'] });
  await header(page, '区分').click({ modifiers: ['Control'] });
  await clearCalls(page);
  const src = (await header(page, '学籍番号').boundingBox())!;
  const k3 = (await header(page, '科目3').boundingBox())!;
  const k4 = (await header(page, '科目4').boundingBox())!;
  const dropX = (k3.x + k3.width + k4.x) / 2; // boundary 科目3 | 科目4
  await page.mouse.move(src.x + src.width / 2, src.y + src.height / 2);
  await page.mouse.down();
  await page.mouse.move(src.x + src.width / 2 + 30, src.y + 20, { steps: 3 });
  await page.mouse.move(dropX + 5, src.y + 200, { steps: 12 });
  await page.waitForTimeout(80);
  const ind = await page.getByTestId('grid-drop-indicator').boundingBox();
  expect(ind).not.toBeNull();
  expect(Math.abs(ind!.x + 2 - k4.x)).toBeLessThan(4);
  await expect(page.locator('.g-ghost')).toHaveText(/^3列を移動: 氏名カナ、学籍番号、区分/);
  expect(await page.locator('.g-hcell.is-dragging').count()).toBe(3);
  await page.screenshot({ path: path.join(SHOTS, '02-drag-indicator.png') });
  await page.mouse.up();
  const moved = ['氏名カナ', '学籍番号', '区分'];
  const rest = initial.filter((t) => !moved.includes(t));
  const at = rest.indexOf('科目4');
  expect(await titles(page)).toEqual([...rest.slice(0, at), ...moved, ...rest.slice(at)]);
  const mv = await calls(page, 'onMoveColumns');
  expect(mv).toHaveLength(1);
  expect(mv[0].args[0]).toEqual([await idOf(page, '氏名カナ'), await idOf(page, '学籍番号'), await idOf(page, '区分')]);
  expect(mv[0].args[1]).toBe(initial.indexOf('科目4')); // gap in pre-move coordinates
  expect(await selectedTitles(page)).toEqual(moved); // selection follows the ids
  // header letters follow the new positions
  await expect(page.getByTestId('grid-header-2')).toContainText('大学');
  // undo / redo through the grid's keyboard handling
  await page.keyboard.press('Control+z');
  expect(await titles(page)).toEqual(initial);
  await page.keyboard.press('Control+y');
  expect((await titles(page)).indexOf('氏名カナ')).toBe(at);
  await page.keyboard.press('Control+Shift+Z'); // nothing left to redo, but the callback fires
  expect((await calls(page, 'onUndo')).length).toBe(1);
  expect((await calls(page, 'onRedo')).length).toBe(2);
  await expectNoErrors(errors);
});

test('dropping inside the selection is a no-op (no callback), plain click on a multi-selection collapses it', async ({ page }) => {
  const errors = await openDemo(page);
  await header(page, '大学').click();
  await header(page, '学籍番号').click({ modifiers: ['Shift'] });
  await clearCalls(page);
  const a = (await header(page, '大学').boundingBox())!;
  const b = (await header(page, '学籍番号').boundingBox())!;
  await page.mouse.move(a.x + 20, a.y + 20);
  await page.mouse.down();
  await page.mouse.move(b.x + b.width / 2 - 5, a.y + 20, { steps: 6 });
  await page.waitForTimeout(50);
  await page.mouse.up();
  expect(await calls(page, 'onMoveColumns')).toHaveLength(0);
  expect((await history(page)).past).toBe(0);
  expect(await selectedTitles(page)).toEqual(['大学', '学籍番号']);
  await header(page, '学籍番号').click();
  expect(await selectedTitles(page)).toEqual(['学籍番号']);
  await expectNoErrors(errors);
});

test('releasing a dragged column far outside the grid cancels the move (no callback); a small overshoot at the edge still drops', async ({ page }) => {
  const errors = await openDemo(page);
  const grid = (await page.getByTestId('grid').boundingBox())!;
  await header(page, '大学').click();
  await clearCalls(page);
  const a = (await header(page, '大学').boundingBox())!;
  // far to the right of the grid (another panel of the app): cancel, the drop indicator disappears while the pointer is out there
  await page.mouse.move(a.x + 20, a.y + 20);
  await page.mouse.down();
  await page.mouse.move(grid.x + grid.width + 200, a.y + 20, { steps: 8 });
  await page.waitForTimeout(80);
  await expect(page.getByTestId('grid-drop-indicator')).toHaveCount(0);
  await page.mouse.up();
  expect(await calls(page, 'onMoveColumns')).toHaveLength(0);
  expect((await history(page)).past).toBe(0);
  // far above the grid: cancel as well
  await page.mouse.move(a.x + 20, a.y + 20);
  await page.mouse.down();
  await page.mouse.move(a.x + 200, grid.y - 150, { steps: 8 });
  await page.mouse.up();
  expect(await calls(page, 'onMoveColumns')).toHaveLength(0);
  // a drop a few pixels beyond the right edge (overshoot while auto-scrolling) is a normal drop
  await clearCalls(page);
  await page.mouse.move(a.x + 20, a.y + 20);
  await page.mouse.down();
  await page.mouse.move(grid.x + grid.width + 10, a.y + 20, { steps: 8 });
  await page.waitForTimeout(600);
  await page.mouse.up();
  expect((await calls(page, 'onMoveColumns')).length).toBe(1);
  await expectNoErrors(errors);
});

test('auto-scroll while dragging, drag survives virtualisation of the source header, Esc cancels', async ({ page }) => {
  const errors = await openDemo(page);
  const sc = (await page.getByTestId('grid').boundingBox())!;
  const src = (await header(page, '大学').boundingBox())!;
  await page.mouse.click(src.x + 10, src.y + 10);
  await page.mouse.move(src.x + 10, src.y + 10);
  await page.mouse.down();
  await page.mouse.move(src.x + 40, src.y + 10, { steps: 3 });
  await page.mouse.move(sc.x + sc.width - 20, src.y + 10, { steps: 10 });
  const sl0 = await page.getByTestId('grid').evaluate((e) => e.scrollLeft);
  // up to 32 px per frame near the edge; poll instead of a fixed wait (slow shared machine)
  await expect.poll(() => page.getByTestId('grid').evaluate((e) => e.scrollLeft), { timeout: 10_000 }).toBeGreaterThan(sl0 + 1200);
  expect(await header(page, '大学').count()).toBe(0); // virtualised away
  await expect(page.getByTestId('grid-drop-indicator')).toHaveCount(1);
  await page.screenshot({ path: path.join(SHOTS, '03-autoscroll.png') });
  await page.mouse.move(sc.x + sc.width / 2, src.y + 10, { steps: 4 });
  await page.mouse.up();
  const i = (await titles(page)).indexOf('大学');
  expect(i).toBeGreaterThan(20);
  await page.keyboard.press('Control+z');
  expect((await titles(page)).indexOf('大学')).toBe(3);
  // Esc cancels a running drag
  await page.evaluate(() => {
    (document.querySelector('[data-testid=grid]') as HTMLElement).scrollLeft = 0;
  });
  await frames(page, 3);
  const b = (await header(page, '大学').boundingBox())!;
  await clearCalls(page);
  await page.mouse.move(b.x + 10, b.y + 10);
  await page.mouse.down();
  await page.mouse.move(b.x + 300, b.y + 10, { steps: 8 });
  await expect(page.getByTestId('grid-drop-indicator')).toHaveCount(1);
  await page.keyboard.press('Escape');
  await expect(page.getByTestId('grid-drop-indicator')).toHaveCount(0);
  await page.mouse.up();
  expect(await calls(page, 'onMoveColumns')).toHaveLength(0);
  expect((await titles(page)).indexOf('大学')).toBe(3);
  await expectNoErrors(errors);
});

test('drop into the frozen pane (gap 1) while scrolled; frozen columns and header stay in place', async ({ page }) => {
  const errors = await openDemo(page);
  const id0 = (await header(page, 'ID').boundingBox())!;
  await page.evaluate(() => {
    const sc = document.querySelector('[data-testid=grid]') as HTMLElement;
    sc.scrollLeft = 2500;
    sc.scrollTop = 50000;
  });
  await frames(page, 3);
  const id1 = (await header(page, 'ID').boundingBox())!;
  expect(Math.abs(id1.x - id0.x)).toBeLessThan(1);
  expect(Math.abs(id1.y - id0.y)).toBeLessThan(1);
  await page.screenshot({ path: path.join(SHOTS, '04-scrolled-frozen.png') });
  const visible = await page.$$eval('.g-header .g-hcell .g-htitle', (els) => els.map((e) => e.textContent ?? ''));
  const target = visible.find((t) => t.startsWith('追加項目'))!;
  const tb = (await header(page, target).boundingBox())!;
  const nameBox = (await header(page, '氏名').boundingBox())!;
  await page.mouse.click(tb.x + 10, tb.y + 10);
  await page.mouse.move(tb.x + 10, tb.y + 10);
  await page.mouse.down();
  await page.mouse.move(tb.x - 30, tb.y + 10, { steps: 3 });
  await page.mouse.move(nameBox.x + 4, tb.y + 10, { steps: 10 }); // left half of 氏名 -> gap 1
  await page.waitForTimeout(60);
  await page.mouse.up();
  expect((await titles(page))[1]).toBe(target);
  await expectNoErrors(errors);
});

test('resize by handle (live preview + one commit), Esc restores, double-click autofits, undo', async ({ page }) => {
  const errors = await openDemo(page);
  const id = await idOf(page, '氏名カナ');
  const b = await reveal(page, '氏名カナ');
  await clearCalls(page);
  const hx = b.x + b.width - 1;
  await page.mouse.move(hx, b.y + 20);
  await page.mouse.down();
  await page.mouse.move(hx + 60, b.y + 20, { steps: 6 });
  await page.mouse.up();
  const b2 = (await header(page, '氏名カナ').boundingBox())!;
  expect(Math.abs(b2.width - b.width - 60)).toBeLessThanOrEqual(2);
  const rs = await calls(page, 'onResizeColumn');
  expect(rs.length).toBeGreaterThan(1);
  expect(rs.slice(0, -1).every((c) => c.args[2] === null)).toBe(true); // live previews
  expect(rs[rs.length - 1].args[2]).toEqual({ from: 150 }); // exactly one commit
  expect(rs.filter((c) => c.args[2] !== null)).toHaveLength(1);
  expect(await widthOf(page, id)).toBe(210);
  await page.keyboard.press('Control+z');
  expect(await widthOf(page, id)).toBe(150);
  // Esc during a resize restores the width, no history entry
  const before = (await history(page)).past;
  const b3 = (await header(page, '氏名カナ').boundingBox())!;
  await page.mouse.move(b3.x + b3.width - 1, b3.y + 20);
  await page.mouse.down();
  await page.mouse.move(b3.x + b3.width + 80, b3.y + 20, { steps: 5 });
  await page.keyboard.press('Escape');
  await page.mouse.up();
  expect((await history(page)).past).toBe(before);
  expect((await header(page, '氏名カナ').boundingBox())!.width).toBeCloseTo(150, 0);
  // double-click on the handle = autofit (one committed call)
  await clearCalls(page);
  const b4 = (await header(page, '氏名カナ').boundingBox())!;
  await page.mouse.dblclick(b4.x + b4.width - 1, b4.y + 20);
  const fit = await calls(page, 'onResizeColumn');
  expect(fit).toHaveLength(1);
  expect(fit[0].args[2]).toEqual({ from: 150 });
  const fitted = fit[0].args[1] as number;
  expect(fitted).toBeGreaterThan(90);
  expect(fitted).toBeLessThan(150); // "ヤマモト ショウタ" at 11pt is narrower than 150 px
  expect(await widthOf(page, id)).toBe(fitted);
  await expectNoErrors(errors);
});

test('context menu: 左に列を挿入 / 右に列を挿入 (contiguous count) / 列を削除 / 列幅を元に戻す', async ({ page }) => {
  const errors = await openDemo(page);
  const n0 = (await titles(page)).length;
  await header(page, '大学').click({ button: 'right' });
  const menu = page.getByTestId('context-menu');
  await expect(menu).toBeVisible();
  await expect(menu.getByRole('menuitem')).toHaveText(['左に列を挿入', '右に列を挿入', '列を削除', '列幅を元に戻す']);
  expect(await selectedTitles(page)).toEqual(['大学']); // right-click selects the column
  await page.screenshot({ path: path.join(SHOTS, '05-context-menu.png') });
  await clearCalls(page);
  await menu.getByRole('menuitem', { name: '左に列を挿入' }).click();
  await expect(menu).toHaveCount(0);
  expect((await calls(page, 'onInsertColumns'))[0].args).toEqual([3, 1]);
  let t = await titles(page);
  expect(t.length).toBe(n0 + 1);
  expect(t[t.indexOf('大学') - 1]).toMatch(/^新しい列/);
  // two contiguous selected columns -> insert 2 on the right
  await header(page, '大学').click();
  await header(page, '学籍番号').click({ modifiers: ['Shift'] });
  await header(page, '学籍番号').click({ button: 'right' });
  await page.getByTestId('context-menu-item-insert-right').click();
  expect((await calls(page, 'onInsertColumns'))[1].args).toEqual([6, 2]);
  t = await titles(page);
  expect(t.slice(4, 8).map((x) => x.replace(/\d+$/, ''))).toEqual(['大学', '学籍番号', '新しい列', '新しい列']);
  // delete via the menu
  const birthId = await idOf(page, '生年月日');
  await header(page, '生年月日').click({ button: 'right' });
  await page.getByRole('menuitem', { name: '列を削除' }).click();
  expect((await calls(page, 'onDeleteColumns'))[0].args[0]).toEqual([birthId]);
  expect(await titles(page)).not.toContain('生年月日');
  // reset width
  const id = await idOf(page, '入学日');
  const hb = await reveal(page, '入学日');
  await page.mouse.move(hb.x + hb.width - 1, hb.y + 20);
  await page.mouse.down();
  await page.mouse.move(hb.x + hb.width + 49, hb.y + 20, { steps: 4 });
  await page.mouse.up();
  expect(await widthOf(page, id)).toBe(140);
  await header(page, '入学日').click({ button: 'right' });
  await page.getByRole('menuitem', { name: '列幅を元に戻す' }).click();
  expect((await calls(page, 'onResetColumnWidths'))[0].args[0]).toEqual([id]);
  expect(await widthOf(page, id)).toBe(90);
  // keyboard: menu opens with Shift+F10, arrows move, Esc closes and focus returns to the grid
  await page.keyboard.press('Shift+F10');
  await expect(menu).toBeVisible();
  await page.keyboard.press('ArrowDown');
  await expect(page.getByTestId('context-menu-item-insert-right')).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(menu).toHaveCount(0);
  await expect(page.getByTestId('grid-editor')).toBeFocused();
  // right-click on a cell opens the same column menu
  await page.getByTestId('grid-cell-4-0').click({ button: 'right' });
  await expect(menu).toBeVisible();
  await page.mouse.click(5, 5); // outside closes
  await expect(menu).toHaveCount(0);
  await expectNoErrors(errors);
});

test('Delete key: selected columns -> onDeleteColumns; otherwise clears the active cell; Ctrl+- deletes too', async ({ page }) => {
  const errors = await openDemo(page);
  await header(page, '大学').click();
  await header(page, '入学日').click({ modifiers: ['Control'] });
  await clearCalls(page);
  const ids = [await idOf(page, '大学'), await idOf(page, '入学日')];
  const n0 = (await titles(page)).length;
  await page.keyboard.press('Delete');
  expect((await calls(page, 'onDeleteColumns'))[0].args[0]).toEqual(ids);
  expect((await titles(page)).length).toBe(n0 - 2);
  await page.keyboard.press('Control+z');
  expect((await titles(page)).length).toBe(n0);
  // no column selection: Delete clears the active cell through onCommitEdit(pos, '')
  await page.getByTestId('grid-cell-5-2').click();
  await clearCalls(page);
  await page.keyboard.press('Delete');
  const ce = await calls(page, 'onCommitEdit');
  expect(ce).toHaveLength(1);
  expect(ce[0].args).toEqual([{ row: 5, col: await idOf(page, '氏名カナ') }, '']);
  await expect(page.getByTestId('grid-cell-5-2')).toHaveText('');
  // Delete on an empty cell does nothing
  await page.keyboard.press('Delete');
  expect(await calls(page, 'onCommitEdit')).toHaveLength(1);
  // Ctrl+- with a selection
  const kubunId = await idOf(page, '区分');
  await header(page, '区分').click();
  await page.keyboard.press('Control+-');
  expect((await calls(page, 'onDeleteColumns')).at(-1)!.args[0]).toEqual([kubunId]);
  await expectNoErrors(errors);
});
