/**
 * Grid browser tests, part 2: active cell + keyboard navigation, in-cell editing (typing / F2 / double-click),
 * Japanese IME via CDP composition events, formula bar, undo keys while editing, Ctrl+C, scrollToRow.
 */
import { expect, test, type Page } from '@playwright/test';
import path from 'node:path';
import { SHOTS, address, calls, cellOf, cellValue, clearCalls, expectNoErrors, fbValue, frames, idOf, indexOfTitle, openDemo, reveal, revealRow } from './helpers';

const editor = (page: Page) => page.getByTestId('grid-editor');

test('keyboard navigation: arrows, Ctrl+arrows, Tab, Enter, Home/End, Ctrl+Home/End, PageDown/PageUp', async ({ page }) => {
  const errors = await openDemo(page);
  await page.getByTestId('grid-cell-2-0').click();
  await expect(page.getByTestId('formula-bar-address')).toHaveText('A2');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowRight');
  expect(await address(page)).toBe('B3');
  await page.keyboard.press('Tab');
  expect(await address(page)).toBe('C3');
  await page.keyboard.press('Shift+Tab');
  await page.keyboard.press('Enter');
  expect(await address(page)).toBe('B4');
  await page.keyboard.press('Shift+Enter');
  expect(await address(page)).toBe('B3');
  await page.keyboard.press('End');
  expect(await address(page)).toMatch(/^[A-Z]+3$/);
  await expect(page.locator('.g-hcell.is-activecol')).toContainText('追加項目');
  await page.keyboard.press('Home');
  expect(await address(page)).toBe('A3');
  await page.keyboard.press('Control+End');
  await frames(page, 3);
  expect(await address(page)).toMatch(/^[A-Z]+10001$/);
  await expect(page.getByTestId('grid-rowhdr-10001')).toBeVisible(); // scrolled into view
  await page.keyboard.press('Control+Home');
  await frames(page, 3);
  expect(await address(page)).toBe('A2');
  await page.keyboard.press('PageDown');
  const r = Number((await address(page))!.slice(1));
  expect(r).toBeGreaterThan(20);
  await page.keyboard.press('PageUp');
  expect(await address(page)).toBe('A2');
  await page.keyboard.press('Control+ArrowDown');
  expect(await address(page)).toBe('A10001');
  await page.keyboard.press('Control+ArrowUp');
  await page.keyboard.press('Control+ArrowRight');
  await frames(page, 3);
  expect(await address(page)).toMatch(/^[A-Z]{2}2$/);
  await page.keyboard.press('ArrowUp'); // clamped at the first data row
  expect(await address(page)).toMatch(/^[A-Z]{2}2$/);
  await expectNoErrors(errors);
});

test('typing / Enter / F2 / Esc / Tab / double-click editing, Ctrl+Z inside the editor stays native', async ({ page }) => {
  const errors = await openDemo(page);
  const univ = await idOf(page, '大学');
  const i = await indexOfTitle(page, '大学');
  await (await cellOf(page, 5, '大学')).click();
  await clearCalls(page);
  await page.keyboard.type('abc');
  await expect(editor(page)).toHaveClass(/is-editing/);
  await expect(page.getByTestId('formula-bar-input')).toHaveValue('abc'); // formula bar mirrors the cell editor
  await page.keyboard.press('Enter');
  expect((await calls(page, 'onCommitEdit'))[0].args).toEqual([{ row: 5, col: univ }, 'abc']);
  await expect(page.getByTestId(`grid-cell-5-${i}`)).toHaveText('abc');
  expect(await address(page)).toBe('D6'); // Enter moves down
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('F2');
  await expect(editor(page)).toHaveValue('abc');
  await page.keyboard.type('def');
  await page.keyboard.press('Escape');
  await expect(editor(page)).not.toHaveClass(/is-editing/);
  await expect(page.getByTestId(`grid-cell-5-${i}`)).toHaveText('abc');
  expect(await calls(page, 'onCommitEdit')).toHaveLength(1);
  // F2 then End appends; Tab commits and moves right
  await page.keyboard.press('F2');
  await page.keyboard.press('End');
  await page.keyboard.type('XYZ');
  await page.keyboard.press('Tab');
  await expect(page.getByTestId(`grid-cell-5-${i}`)).toHaveText('abcXYZ');
  expect(await address(page)).toBe('E5');
  // F2 + Enter without a change does not create a commit
  await page.keyboard.press('ArrowLeft');
  await page.keyboard.press('F2');
  await page.keyboard.press('Enter');
  expect(await calls(page, 'onCommitEdit')).toHaveLength(2);
  // Ctrl+Z while editing = native text undo, NOT the grid history
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('F2');
  await page.keyboard.type('q');
  await page.keyboard.press('Control+z');
  expect(await calls(page, 'onUndo')).toHaveLength(0);
  await page.keyboard.press('Escape');
  // Ctrl+Z outside editing -> onUndo -> cell restored
  await page.keyboard.press('Control+z');
  expect(await calls(page, 'onUndo')).toHaveLength(1);
  await expect(page.getByTestId(`grid-cell-5-${i}`)).toHaveText('abc');
  // double-click edits the existing content; a click elsewhere commits
  await page.getByTestId(`grid-cell-5-${i}`).dblclick();
  await expect(editor(page)).toHaveClass(/is-editing/);
  await expect(editor(page)).toHaveValue('abc');
  await page.keyboard.type('!');
  await page.getByTestId('grid-cell-8-0').click();
  await expect(page.getByTestId(`grid-cell-5-${i}`)).toHaveText('abc!');
  expect(await address(page)).toBe('A8');
  // Alt+Enter inserts a line break, Backspace clears and edits, Ctrl+Enter commits in place
  await page.keyboard.press('Backspace');
  await expect(editor(page)).toHaveValue('');
  await page.keyboard.type('x');
  await page.keyboard.press('Alt+Enter');
  await page.keyboard.type('y');
  await page.keyboard.press('Control+Enter');
  expect(await address(page)).toBe('A8');
  expect(await cellValue(page, 8, 'id')).toBe('x\ny');
  // text column keeps leading zeros (the demo parses like the doc: text stays text)
  await (await cellOf(page, 9, '学籍番号')).click();
  await page.keyboard.type('0012345');
  await page.keyboard.press('Enter');
  await expect(await cellOf(page, 9, '学籍番号')).toHaveText('0012345');
  // typing "=..." shows the formula (no cached value -> grey italic)
  await (await cellOf(page, 10, '奨学金')).click();
  await page.keyboard.type('=SUM(A1:B1)');
  await page.keyboard.press('Enter');
  const fcell = await cellOf(page, 10, '奨学金');
  await expect(fcell).toHaveText('=SUM(A1:B1)');
  await expect(fcell).toHaveClass(/is-ftext/);
  await expectNoErrors(errors);
});

test('Japanese IME: composition on a not-yet-editing cell keeps every character; Enter during composition is ignored', async ({ page }) => {
  const errors = await openDemo(page);
  const cdp = await page.context().newCDPSession(page);
  const i = await indexOfTitle(page, '大学');
  await (await cellOf(page, 7, '大学')).click();
  await clearCalls(page);
  // first key of a Japanese IME: the composition starts on the invisible sink
  await cdp.send('Input.imeSetComposition', { text: 'に', selectionStart: 1, selectionEnd: 1 });
  await cdp.send('Input.imeSetComposition', { text: 'にほ', selectionStart: 2, selectionEnd: 2 });
  await cdp.send('Input.imeSetComposition', { text: 'にほんご', selectionStart: 4, selectionEnd: 4 });
  await expect(editor(page)).toHaveClass(/is-editing/);
  await expect(editor(page)).toHaveValue('にほんご');
  // an Enter keydown that belongs to the composition (isComposing / keyCode 229) must not commit
  await editor(page).evaluate((el) => {
    const ev = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true, isComposing: true });
    el.dispatchEvent(ev);
    const ev2 = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true });
    Object.defineProperty(ev2, 'keyCode', { get: () => 229 });
    el.dispatchEvent(ev2);
  });
  expect(await calls(page, 'onCommitEdit')).toHaveLength(0);
  await expect(editor(page)).toHaveClass(/is-editing/);
  await page.screenshot({ path: path.join(SHOTS, '06-ime-composing.png') });
  await cdp.send('Input.insertText', { text: '日本語' }); // conversion confirmed
  await expect(editor(page)).toHaveValue('日本語');
  await page.keyboard.press('Enter');
  await expect(page.getByTestId(`grid-cell-7-${i}`)).toHaveText('日本語');
  expect((await calls(page, 'onCommitEdit')).map((c) => c.args[1])).toEqual(['日本語']);
  // IME in F2 mode appends at the caret
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('F2');
  await cdp.send('Input.imeSetComposition', { text: 'だいがく', selectionStart: 4, selectionEnd: 4 });
  await cdp.send('Input.insertText', { text: '大学' });
  await page.keyboard.press('Enter');
  await expect(page.getByTestId(`grid-cell-7-${i}`)).toHaveText('日本語大学');
  await expectNoErrors(errors);
});

test('formula bar: shows the active cell, Enter commits via onCommitEdit, Esc cancels, IME-safe, focus returns to the grid', async ({ page }) => {
  const errors = await openDemo(page);
  const fb = page.getByTestId('formula-bar-input');
  // formula cell -> formula text (current letters)
  const avgIdx = await indexOfTitle(page, '平均点');
  await (await cellOf(page, 4, '平均点')).click();
  await expect(page.getByTestId('formula-bar-address')).toHaveText(/^[A-Z]+4$/);
  expect(await fbValue(page)).toMatch(/^=AVERAGE\([A-Z]+4:[A-Z]+4\)$/);
  void avgIdx;
  // date cell -> yyyy/m/d like Excel's formula bar; percentage -> 12.3%
  await (await cellOf(page, 4, '入学日')).click();
  expect(await fbValue(page)).toMatch(/^\d{4}\/\d{1,2}\/\d{1,2}$/);
  await (await cellOf(page, 4, '出席率')).click();
  expect(await fbValue(page)).toMatch(/^\d+(\.\d+)?%$/);
  // edit in the formula bar
  const univ = await idOf(page, '大学');
  await (await cellOf(page, 6, '大学')).click();
  await clearCalls(page);
  await fb.click();
  await fb.fill('早稲田');
  await page.keyboard.press('Enter');
  expect((await calls(page, 'onCommitEdit'))[0].args).toEqual([{ row: 6, col: univ }, '早稲田']);
  await expect(await cellOf(page, 6, '大学')).toHaveText('早稲田');
  await expect(editor(page)).toBeFocused(); // onDone -> grid focus
  // Esc restores
  await fb.click();
  await fb.fill('やめる');
  await page.keyboard.press('Escape');
  await expect(fb).toHaveValue('早稲田');
  expect(await calls(page, 'onCommitEdit')).toHaveLength(1);
  // IME in the formula bar: Enter that confirms a conversion does not commit
  const cdp = await page.context().newCDPSession(page);
  await fb.click();
  await fb.fill('');
  await cdp.send('Input.imeSetComposition', { text: 'けいおう', selectionStart: 4, selectionEnd: 4 });
  await fb.evaluate((el) => {
    const ev = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true, isComposing: true });
    el.dispatchEvent(ev);
  });
  expect(await calls(page, 'onCommitEdit')).toHaveLength(1);
  await cdp.send('Input.insertText', { text: '慶應' });
  await expect(fb).toHaveValue('慶應');
  await page.keyboard.press('Enter');
  expect((await calls(page, 'onCommitEdit'))[1].args[1]).toBe('慶應');
  await expect(await cellOf(page, 6, '大学')).toHaveText('慶應');
  // leaving the formula bar with a change commits it to the cell that was active when typing started
  await fb.click();
  await fb.fill('一橋');
  await page.getByTestId('grid-cell-3-0').click();
  expect((await calls(page, 'onCommitEdit'))[2].args).toEqual([{ row: 6, col: univ }, '一橋']);
  expect(await address(page)).toBe('A3');
  await page.screenshot({ path: path.join(SHOTS, '07-formula-bar.png') });
  await expectNoErrors(errors);
});

test('Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z call onUndo / onRedo when not editing', async ({ page }) => {
  const errors = await openDemo(page);
  await page.getByTestId('grid-cell-3-2').click();
  await clearCalls(page);
  await page.keyboard.press('Control+z');
  await page.keyboard.press('Control+y');
  await page.keyboard.press('Control+Shift+Z');
  expect((await calls(page)).map((c) => c.name)).toEqual(['onUndo', 'onRedo', 'onRedo']);
  await expectNoErrors(errors);
});

test('Ctrl+C: the copy event carries the displayed text of the active cell (system clipboard untouched)', async ({ page }) => {
  const errors = await openDemo(page);
  await page.getByTestId('grid-cell-2-0').click();
  // 1) a real Ctrl+C on the sink fires a copy event (stopped in the capture phase so nothing reaches the OS clipboard)
  await page.evaluate(() => {
    (window as unknown as { __copies: number }).__copies = 0;
    window.addEventListener(
      'copy',
      (e) => {
        (window as unknown as { __copies: number }).__copies++;
        e.stopImmediatePropagation();
      },
      { capture: true, once: true },
    );
  });
  await page.keyboard.press('Control+c');
  expect(await page.evaluate(() => (window as unknown as { __copies: number }).__copies)).toBe(1);
  // 2) the grid's copy handler puts the displayed text into the event's DataTransfer (synthetic event)
  const text = await editor(page).evaluate((el) => {
    const dt = new DataTransfer();
    const ev = new ClipboardEvent('copy', { clipboardData: dt, bubbles: true, cancelable: true });
    el.dispatchEvent(ev);
    return { text: dt.getData('text/plain'), prevented: ev.defaultPrevented };
  });
  expect(text).toEqual({ text: '100001', prevented: true });
  await (await cellOf(page, 2, '奨学金')).click();
  const yen = await editor(page).evaluate((el) => {
    const dt = new DataTransfer();
    el.dispatchEvent(new ClipboardEvent('copy', { clipboardData: dt, bubbles: true, cancelable: true }));
    return dt.getData('text/plain');
  });
  expect(yen).toMatch(/^-?¥[\d,]+$/);
  // paste (synthetic event, clipboard untouched): the first cell of the pasted text goes into the active cell
  await clearCalls(page);
  await editor(page).evaluate((el) => {
    const dt = new DataTransfer();
    dt.setData('text/plain', '12345\t999\r\nzzz\r\n');
    el.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  });
  const pc = await calls(page, 'onCommitEdit');
  expect(pc).toHaveLength(1);
  expect(pc[0].args[1]).toBe('12345');
  await expectNoErrors(errors);
});

test('scrollToRow / scrollToColumn handle; rows outside the grid are refused', async ({ page }) => {
  const errors = await openDemo(page);
  await page.getByTestId('demo-jump').fill('5000');
  await page.getByTestId('demo-jump').press('Enter');
  await frames(page, 3);
  await expect(page.getByTestId('grid-rowhdr-5000')).toBeVisible();
  expect(await address(page)).toBe('A5000');
  await expect(page.getByTestId('grid-rowhdr-5000')).toHaveClass(/is-activerow/);
  expect(await page.evaluate(() => window.__gridDemo!.grid()!.scrollToRow(1))).toBe(false); // header row
  expect(await page.evaluate(() => window.__gridDemo!.grid()!.scrollToRow(10002))).toBe(false);
  expect(await page.evaluate(() => window.__gridDemo!.grid()!.scrollToColumn('nope'))).toBe(false);
  const id = await idOf(page, '追加項目15');
  expect(await page.evaluate((c) => window.__gridDemo!.grid()!.scrollToColumn(c, { select: true }), id)).toBe(true);
  await frames(page, 3);
  await expect(page.locator('.g-hcell.is-sel')).toContainText('追加項目15');
  await reveal(page, 'ID');
  await revealRow(page, 2);
  await expectNoErrors(errors);
});

test('read-only mode: typing / F2 / Delete do not edit; column operations still work', async ({ page }) => {
  const errors = await openDemo(page, '?readonly=1');
  await page.getByTestId('grid-cell-3-3').click();
  await clearCalls(page);
  await page.keyboard.type('abc');
  await page.keyboard.press('F2');
  await page.keyboard.press('Delete');
  await expect(editor(page)).not.toHaveClass(/is-editing/);
  expect(await calls(page, 'onCommitEdit')).toHaveLength(0);
  await page.getByTestId('grid-header-3').click();
  await page.keyboard.press('Delete');
  expect(await calls(page, 'onDeleteColumns')).toHaveLength(1);
  await expectNoErrors(errors);
});
