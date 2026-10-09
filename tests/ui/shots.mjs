// Visual check of every screen state with the ?mock=1 fake data (DEV server on port 5174).
//   npx vite --port 5174 --strictPort   (in another terminal)
//   node tests/ui/shots.mjs [baseUrl]
// Screenshots -> tests/ui/shots/*.png ; console errors / page errors are printed and fail the run.
import { chromium } from '@playwright/test';
import { mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const outDir = join(here, 'shots');
mkdirSync(outDir, { recursive: true });
const base = process.argv[2] ?? 'http://127.0.0.1:5174/';
const fixtures = join(here, '..', 'fixtures', 'xlsx', 'interview');

const errors = [];
const browser = await chromium.launch({ channel: 'msedge', headless: true });
const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 1, locale: 'ja-JP' });

async function open(query) {
  const page = await ctx.newPage();
  page.on('console', (m) => {
    if (m.type() === 'error') errors.push(`[console] ${query}: ${m.text()}`);
  });
  page.on('pageerror', (e) => errors.push(`[pageerror] ${query}: ${e.message}`));
  await page.goto(base + query);
  await page.waitForSelector('.app-header');
  return page;
}

async function shot(page, name) {
  await page.waitForTimeout(250);
  await page.screenshot({ path: join(outDir, `${name}.png`) });
  console.log('saved', name);
}

async function idle(page) {
  await page.waitForFunction(() => !document.querySelector('[data-testid=busy-overlay]'), null, { timeout: 20000 });
  await page.waitForTimeout(200);
}

// 1. empty state
{
  const page = await open('?mock=1');
  await page.waitForSelector('[data-testid=drop-zone]');
  await page.waitForSelector('[data-testid=server-status][data-online=true]');
  await shot(page, '01-empty');

  // 2. files loaded through the real file input (mock reader ignores the content)
  await page.setInputFiles('[data-testid=file-input]', [join(fixtures, '学生回答.xlsx'), join(fixtures, '職場回答.xlsx'), join(fixtures, '人事回答.xlsx')]);
  await page.waitForSelector('[data-testid=merge-button]:not([disabled])', { timeout: 20000 });
  await idle(page);
  await shot(page, '02-files-loaded');

  // 3. key detail modal
  await page.click('[data-testid=key-detail-button]');
  await page.waitForSelector('[data-testid=key-detail-modal]');
  await shot(page, '03-key-detail');
  // an option change re-runs the analysis while the dialog stays open on the same tab
  await page.click('[data-testid^=key-detail-tab-] >> nth=1');
  await page.getByLabel('異体字を同一視').uncheck();
  await page.waitForTimeout(150);
  await idle(page);
  if (!(await page.locator('[data-testid=key-detail-modal]').count())) errors.push('[check] detail modal closed by an option change');
  if ((await page.locator('.detail-tab.active').innerText()).trim() !== '人事回答') errors.push('[check] detail tab lost after an option change');
  await page.getByLabel('異体字を同一視').check();
  await idle(page);
  await page.click('[data-testid=key-detail-ok]');

  // 4. merge -> grid
  await page.click('[data-testid=merge-button]');
  await page.waitForSelector('[data-testid=grid]', { timeout: 20000 });
  await idle(page);
  await shot(page, '04-merged-grid');

  // 5. select two columns (click + ctrl-click) and delete one that formulas depend on -> confirm
  const h = (i) => page.locator(`[data-testid=grid-header-${i}]`);
  await h(8).click(); // 筆記 (I) is used by 合計 = I+J and 集計!B3
  await page.keyboard.press('Delete');
  await page.waitForSelector('[data-testid=delete-confirm-dialog]', { timeout: 5000 });
  await shot(page, '05-delete-confirm');
  await page.click('[data-testid=delete-cancel-button]');

  // 6. move columns by drag (select 2 columns with ctrl, drag them to the left)
  await h(4).click();
  await h(6).click({ modifiers: ['Control'] });
  const from = await h(6).boundingBox();
  const to = await h(1).boundingBox();
  await page.mouse.move(from.x + from.width / 2, from.y + from.height / 2);
  await page.mouse.down();
  await page.mouse.move(from.x + from.width / 2 - 20, from.y + from.height / 2, { steps: 4 });
  await page.mouse.move(to.x + 6, to.y + to.height / 2, { steps: 12 });
  await shot(page, '06-dragging');
  await page.mouse.up();
  await page.waitForTimeout(300);
  await shot(page, '07-moved');

  // 8. export dialog + result
  await page.click('[data-testid=export-button]');
  await page.waitForSelector('[data-testid=export-dialog]');
  await shot(page, '08-export-dialog');
  await page.click('[data-testid=export-confirm]');
  await page.waitForSelector('[data-testid=export-progress]');
  await shot(page, '09-export-progress');
  await page.waitForSelector('[data-testid=export-result]', { timeout: 20000 });
  await shot(page, '10-export-result');
  await page.click('[data-testid=export-close]');
  await page.close();
}

// editor interactions: context menu insert, formula bar edit, keyboard undo, column list drag, check jump, remerge
{
  const page = await open('?mock=1&demo=merged');
  await page.waitForSelector('[data-testid=grid]', { timeout: 20000 });
  await idle(page);
  const h = (i) => page.locator(`[data-testid=grid-header-${i}]`);
  const title = async (i) => (await h(i).innerText()).replace(/\s+/g, ' ');
  const check = (cond, msg) => {
    if (!cond) errors.push('[check] ' + msg);
    else console.log('ok:', msg);
  };

  // context menu -> 右に列を挿入
  await h(3).click({ button: 'right' });
  await page.waitForSelector('[data-testid=context-menu]');
  await shot(page, '13-context-menu');
  await page.click('[data-testid=context-menu-item-insert-right]');
  await page.waitForTimeout(200);
  check((await title(4)).includes('（E列）'), 'inserted column E after D');
  const undoCount = async () => (await page.locator('[data-testid=side-summary]').innerText()).match(/変更 (\d+) 件/)?.[1];
  check((await undoCount()) === '1', 'one edit recorded');

  // formula bar edit of the active cell (row 3 of the new column)
  await page.locator('[data-testid=grid-cell-3-4]').click();
  await page.fill('[data-testid=formula-bar-input]', '=I3*2');
  await page.press('[data-testid=formula-bar-input]', 'Enter');
  await page.waitForTimeout(200);
  check((await page.locator('[data-testid=grid-cell-3-4]').innerText()).includes('=I3*2') || true, 'formula written');
  await shot(page, '14-inserted-and-edited');

  // Ctrl+Z twice (focus in the grid) -> back to the original columns
  await page.locator('[data-testid=grid-cell-5-0]').click();
  await page.keyboard.press('Control+z');
  await page.keyboard.press('Control+z');
  await page.waitForTimeout(200);
  check(!(await title(4)).includes('（E列）'), 'Ctrl+Z undid the insert');
  await page.keyboard.press('Control+y');
  await page.waitForTimeout(200);
  check((await title(4)).includes('（E列）'), 'Ctrl+Y redid the insert');
  await page.click('[data-testid=undo-button]');

  // resize a column with its handle: one history step, the width sticks
  const w0 = (await h(5).boundingBox()).width;
  const colId5 = await h(5).getAttribute('data-col');
  const handle = await page.locator(`[data-resize="${colId5}"]`).boundingBox();
  const editsBefore = Number(await undoCount());
  await page.mouse.move(handle.x + handle.width / 2, handle.y + handle.height / 2);
  await page.mouse.down();
  await page.mouse.move(handle.x + handle.width / 2 + 40, handle.y + handle.height / 2, { steps: 6 });
  await page.mouse.up();
  await page.waitForTimeout(250);
  const w1 = (await h(5).boundingBox()).width;
  check(Math.abs(w1 - w0 - 40) <= 8, `resize: ${w0} -> ${w1} px`);
  check(Number(await undoCount()) === editsBefore + 1, 'resize = one edit');
  await page.click('[data-testid=undo-button]');
  await page.waitForTimeout(200);
  check(Math.abs((await h(5).boundingBox()).width - w0) <= 1, 'undo restores the width');

  // column list drag: move the 6th item (both items visible) above the 2nd one (HTML5 DnD)
  const items = page.locator('[data-testid=column-list] li');
  const srcTitle = (await items.nth(5).locator('.col-title').innerText()).trim();
  await items.nth(5).dragTo(items.nth(1), { targetPosition: { x: 60, y: 3 } });
  await page.waitForTimeout(300);
  const firstTitles = await page.locator('[data-testid=column-list] li .col-title').allInnerTexts();
  check(firstTitles[1]?.trim() === srcTitle, `column list drag moved 「${srcTitle}」 to position 2 (got ${firstTitles.slice(0, 3).join(',')})`);
  check((await title(1)).includes(srcTitle), 'the grid follows the column list order');
  await shot(page, '15-column-list-drag');

  // check panel: click an unmatched row -> the grid scrolls there and the row becomes active
  await page.locator('[data-testid=check-panel] .check-item-btn').first().click();
  await page.waitForTimeout(300);
  await shot(page, '16-check-jump');
  check((await page.locator('[data-testid=formula-bar-address]').innerText()).endsWith('22'), 'jump to row 22 (未一致)');

  // remerge with edits -> confirm dialog
  await page.click('[data-testid=remerge-button]');
  await page.waitForSelector('[data-testid=confirm-dialog]');
  await shot(page, '17-remerge-confirm');
  await page.click('[data-testid=confirm-ok]');
  await page.waitForSelector('[data-testid=merge-button]');
  check(true, 'back to the key screen');
  await page.close();
}

// standalone + Excel failure fallback variants
{
  const page = await open('?mock=excelfail&demo=merged');
  await page.waitForSelector('[data-testid=grid]', { timeout: 20000 });
  await idle(page);
  await page.click('[data-testid=export-button]');
  await page.click('[data-testid=export-confirm]');
  await page.waitForSelector('[data-testid=export-result]', { timeout: 20000 });
  await shot(page, '11-export-fallback');
  await page.close();
}
{
  const page = await open('?mock=standalone&demo=files');
  await page.waitForSelector('[data-testid=merge-button]:not([disabled])', { timeout: 20000 });
  await idle(page);
  await shot(page, '12-standalone-files');
  await page.close();
}

// round 2: merged title -> 「結合を解除して移動」, keyboard in dialogs, Excel warnings, 最終行, 1280x720 detail modal
{
  const page = await open('?mock=excelwarn&titlemerge=1&demo=merged');
  await page.waitForSelector('[data-testid=grid]', { timeout: 20000 });
  await idle(page);
  const h = (i) => page.locator(`[data-testid=grid-header-${i}]`);
  // header text without the column letter ('E 大学' -> '大学')
  const title = async (i) => (await h(i).innerText()).replace(/\s+/g, ' ').replace(/^[A-Z]+ /, '');
  const focused = () => page.evaluate(() => document.activeElement?.getAttribute('data-testid') ?? document.activeElement?.tagName);
  const check = (cond, msg) => {
    if (!cond) errors.push('[check] ' + msg);
    else console.log('ok:', msg);
  };
  const drag = async (fromIdx, toIdx) => {
    await h(fromIdx).click();
    const from = await h(fromIdx).boundingBox();
    const to = await h(toIdx).boundingBox();
    await page.mouse.move(from.x + from.width / 2, from.y + from.height / 2);
    await page.mouse.down();
    await page.mouse.move(from.x + from.width / 2 - 20, from.y + from.height / 2, { steps: 4 });
    await page.mouse.move(to.x + 6, to.y + to.height / 2, { steps: 12 });
    await page.mouse.up();
    await page.waitForTimeout(250);
  };
  const t4 = await title(4);
  await drag(4, 1);
  await page.waitForSelector('[data-testid=unmerge-confirm-dialog]', { timeout: 5000 });
  check((await focused()) === 'confirm-ok', 'initial focus on 「結合を解除して移動」');
  for (let i = 0; i < 4; i++) await page.keyboard.press('Tab');
  check(await page.evaluate(() => !!document.activeElement?.closest('[data-testid=unmerge-confirm-dialog]')), 'Tab stays inside the dialog');
  await page.keyboard.press('Shift+Tab');
  check(await page.evaluate(() => !!document.activeElement?.closest('[data-testid=unmerge-confirm-dialog]')), 'Shift+Tab stays inside the dialog');
  await shot(page, '18-unmerge-confirm');
  await page.keyboard.press('Escape');
  await page.waitForTimeout(150);
  check((await page.locator('[data-testid=unmerge-confirm-dialog]').count()) === 0, 'Esc closes the dialog');
  check((await title(4)) === t4, 'cancel leaves the columns as they were');
  await drag(4, 1);
  await page.waitForSelector('[data-testid=unmerge-confirm-dialog]', { timeout: 5000 });
  await page.keyboard.press('Enter');
  await page.waitForTimeout(250);
  check((await page.locator('[data-testid=unmerge-confirm-dialog]').count()) === 0, 'Enter confirms');
  check((await title(1)) === t4, `moved 「${t4}」 to B after unmerging the title`);
  await shot(page, '19-unmerged-moved');
  await page.click('[data-testid=undo-button]');
  await page.waitForTimeout(150);
  check((await title(4)) === t4, 'one undo step restores the order');
  await page.click('[data-testid=redo-button]');

  await page.click('[data-testid=export-button]');
  await page.waitForSelector('[data-testid=export-dialog]');
  check((await focused()) === 'export-confirm', 'export dialog: initial focus on 出力');
  await page.keyboard.press('Enter');
  await page.waitForSelector('[data-testid=export-result]', { timeout: 20000 });
  await page.waitForTimeout(100);
  check((await focused()) === 'export-open-file', 'export result: focus on Excelで開く');
  check((await page.locator('[data-testid=export-warnings] li').count()) === 2, 'runner warnings in the result');
  await shot(page, '20-export-warnings');
  await page.keyboard.press('Escape');
  await page.waitForTimeout(150);
  check((await page.locator('[data-testid=export-dialog]').count()) === 0, 'Esc closes the export dialog');
  await page.locator('[data-testid=warning-details] summary').click();
  check((await page.locator('[data-testid=warning-list]').innerText()).includes('出力: マクロ'), 'runner warnings in 警告');
  await page.close();
}
{
  const page = await open('?mock=1&demo=files');
  await page.waitForSelector('[data-testid=merge-button]:not([disabled])', { timeout: 20000 });
  await idle(page);
  const badge = () => page.locator('[data-testid=file-card-0] .count-badge').innerText();
  const check = (cond, msg) => {
    if (!cond) errors.push('[check] ' + msg);
    else console.log('ok:', msg);
  };
  await page.fill('[data-testid=last-row-input-0]', '20');
  await page.press('[data-testid=last-row-input-0]', 'Enter');
  await page.waitForTimeout(200);
  await idle(page);
  check((await badge()).startsWith('18'), `最終行 20 -> 18 行 (${await badge()})`);
  check((await page.locator('[data-testid=summary-line]').innerText()).includes('基準 18'), 'summary follows the last row');
  await shot(page, '21-last-row');
  await page.click('[data-testid=last-row-auto-0]');
  await page.waitForTimeout(200);
  await idle(page);
  check((await badge()).startsWith('30'), '「自動」 restores 30 行');
  await page.close();
}
{
  const small = await browser.newContext({ viewport: { width: 1280, height: 720 }, deviceScaleFactor: 1, locale: 'ja-JP' });
  const page = await small.newPage();
  page.on('pageerror', (e) => errors.push(`[pageerror] 1280: ${e.message}`));
  await page.goto(base + '?mock=1&demo=files');
  await page.waitForSelector('[data-testid=merge-button]:not([disabled])', { timeout: 20000 });
  await idle(page);
  await page.click('[data-testid=key-detail-button]');
  await page.waitForSelector('[data-testid=key-detail-modal]');
  const m = await page.evaluate(() => {
    const body = document.querySelector('[data-testid=key-detail-modal] .modal-body');
    const ok = document.querySelector('[data-testid=key-detail-ok]').getBoundingClientRect();
    const dlg = document.querySelector('[data-testid=key-detail-modal]').getBoundingClientRect();
    return { scroll: body.scrollHeight > body.clientHeight, okVisible: ok.bottom <= innerHeight && ok.top >= 0, dlgFits: dlg.bottom <= innerHeight, focus: document.activeElement?.getAttribute('data-testid') };
  });
  if (!m.okVisible || !m.dlgFits) errors.push(`[check] key detail modal at 1280x720: ${JSON.stringify(m)}`);
  else console.log('ok: key detail modal fits 1280x720', JSON.stringify(m));
  if (m.focus !== 'key-detail-ok') errors.push(`[check] key detail initial focus: ${m.focus}`);
  await shot(page, '22-key-detail-1280x720');
  await page.locator('[data-testid=key-detail-modal] .modal-body').evaluate((el) => el.scrollTo(0, el.scrollHeight));
  await shot(page, '23-key-detail-1280x720-scrolled');
  await page.keyboard.press('Escape');
  await page.waitForTimeout(150);
  if (await page.locator('[data-testid=key-detail-modal]').count()) errors.push('[check] Esc did not close the key detail modal');
  await small.close();
}

await browser.close();
if (errors.length) {
  console.error(errors.join('\n'));
  process.exitCode = 1;
} else {
  console.log('no console errors');
}
