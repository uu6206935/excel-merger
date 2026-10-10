/**
 * Round 7 (user request) against the REAL app and the real local server, in Microsoft Edge:
 *   - drag the file cards to reorder them (before the merge: the order of the imported column groups; after the merge: the
 *     column group of that file moves in the grid), Alt+←/→, Esc, auto-scroll, the OS file drop stays untouched;
 *   - the source band above the grid headers: click selects the file's columns, drag moves them;
 *   - click on a file card: the grid scrolls so that the file's columns are in the middle, all of them selected;
 *   - 「他のシートも出力に含める」.
 *
 *   npx playwright test -c tests/ui/playwright.config.ts w1-cards       (builds the UI, starts the server on :5192)
 *   UI_BASE_URL=http://localhost:5192 npx playwright test -c tests/ui/playwright.config.ts w1-cards   (a server that is already running)
 *
 * Inputs: the 5 sample files (samples/). 01 is the base (12 columns, 2 frozen), 02 / 04 / 05 bring columns, 03 (CSV) brings none
 * by default (its keys are the base's own columns) — the "file without columns in the table" case. Only the built-in engine is
 * used (no Excel start). Screenshots: tests/ui/shots/w1-*.png.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { expect, test, type Locator, type Page } from '@playwright/test';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SAMPLES = path.resolve(HERE, '..', '..', 'samples');
const SHOTS = path.join(HERE, 'shots');
fs.mkdirSync(SHOTS, { recursive: true });
const FILES = ['01_学生名簿.xlsx', '02_面接評価.xlsx', '03_連絡先.csv', '04_適性検査.xls', '05_研修出欠.xlsx'].map((f) => path.join(SAMPLES, f));
const shot = (page: Page, name: string) => page.screenshot({ path: path.join(SHOTS, `w1-${name}.png`) });

// ------------------------------------------------------------------------------------------------ helpers

async function openApp(page: Page): Promise<string[]> {
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(`pageerror: ${e.message}`));
  page.on('console', (m) => {
    if (m.type() === 'error' && !/Failed to load resource|ERR_/.test(m.text())) errors.push(m.text());
  });
  await page.goto('/');
  await page.getByTestId('server-status').waitFor();
  await expect(page.getByTestId('server-status')).toHaveAttribute('data-online', 'true');
  return errors;
}

async function settle(page: Page, timeout = 120_000) {
  await expect(page.locator('[data-testid=busy-overlay], .card-loading')).toHaveCount(0, { timeout });
}

async function addFiles(page: Page, paths: string[]) {
  await page.getByTestId('file-input').setInputFiles(paths);
  await settle(page);
}

async function mergeNow(page: Page) {
  await page.getByTestId('merge-button').click();
  await page.getByTestId('grid').waitFor({ timeout: 120_000 });
  await settle(page);
}

/** all five samples, optionally merged */
async function start(page: Page, merged = false): Promise<string[]> {
  const errors = await openApp(page);
  await addFiles(page, FILES);
  await expect(page.getByTestId('merge-button')).toBeEnabled();
  if (merged) await mergeNow(page);
  return errors;
}

const card = (page: Page, id: string): Locator => page.locator(`[data-file-id="${id}"]`);
const cardIds = (page: Page) => page.$$eval('[data-file-id]', (els) => els.map((e) => e.getAttribute('data-file-id')));
const activeCards = (page: Page) => page.$$eval('[data-file-id].is-active', (els) => els.map((e) => e.getAttribute('data-file-id')));
const gridScrollLeft = (page: Page) => page.getByTestId('grid').evaluate((e) => e.scrollLeft);
const selectedCount = async (page: Page) => Number(/(\d+) 列を選択中/.exec((await page.locator('.hint-sel').textContent().catch(() => '')) ?? '')?.[1] ?? 0);
const sideText = (page: Page) => page.getByTestId('side-summary').innerText();
const editCount = async (page: Page) => Number(/変更 (\d+) 件/.exec(await sideText(page))?.[1] ?? NaN);
/** colour dots of the planned (imported) columns, in order, collapsed to one entry per run */
const plannedGroups = async (page: Page) => {
  const colors = await page.$$eval('[data-testid=column-list] .col-item.is-new .col-color', (els) => els.map((e) => getComputedStyle(e).backgroundColor));
  return colors.filter((c, i) => i === 0 || c !== colors[i - 1]);
};
/** colour dots of ALL columns of the 列の一覧 (merged phase), collapsed to one entry per run = the order of the file groups in the grid */
const listGroups = async (page: Page) => {
  const colors = await page.$$eval('[data-testid=column-list] .col-item .col-color', (els) => els.map((e) => getComputedStyle(e).backgroundColor));
  return colors.filter((c, i) => i === 0 || c !== colors[i - 1]);
};
const cardColor = (page: Page, id: string) => card(page, id).locator('.card-strip').evaluate((e) => getComputedStyle(e).backgroundColor);

/** the bands of the grid by the text of their label: [{ testid, label, count, x, w }] (only the rendered ones) */
async function bands(page: Page) {
  return page.$$eval('.g-band', (els) =>
    els.map((e) => {
      const r = e.getBoundingClientRect();
      const label = e.querySelector('.g-band-label')?.firstChild?.textContent ?? '';
      const count = Number(/(\d+)列$/.exec(e.textContent ?? '')?.[1] ?? NaN);
      return { testid: e.getAttribute('data-testid') ?? '', label, count, x: r.x, w: r.width };
    }),
  );
}

/** column titles of the grid in column order (the rendered ones; pinned header cells carry their column index too) */
async function headerTitles(page: Page): Promise<string[]> {
  const rows = await page.$$eval('[data-testid^=grid-header-]', (els) =>
    els.map((e) => ({ i: Number((e.getAttribute('data-testid') ?? '').replace('grid-header-', '')), t: e.querySelector('.g-htitle')?.textContent ?? '' })),
  );
  return rows.sort((a, b) => a.i - b.i).map((r) => r.t);
}

/** scroll the grid to the right until the band of that file is rendered, then return its locator */
async function revealBand(page: Page, text: string): Promise<Locator> {
  const band = page.locator('.g-band:not([data-testid$=-pin])', { hasText: text }).first();
  for (let i = 0; i < 14; i++) {
    if (await band.count()) break;
    await page.getByTestId('grid').evaluate((e, step) => (e.scrollLeft += step), 450);
    await page.waitForTimeout(120);
  }
  await expect(band).toBeVisible();
  return band;
}

async function openWarnings(page: Page) {
  const d = page.getByTestId('warning-details');
  if (!(await d.evaluate((e) => e.hasAttribute('open')))) await d.locator('summary').click();
}

/** wait until the grid stopped scrolling (smooth scroll) */
async function scrollSettled(page: Page) {
  let last = -1;
  for (let i = 0; i < 40; i++) {
    const now = await gridScrollLeft(page);
    if (now === last) return now;
    last = now;
    await page.waitForTimeout(160);
  }
  return last;
}

/** |middle of the (single piece) band − middle of the grid viewport| as a share of the viewport width */
async function centreError(page: Page, labelStart: string): Promise<number> {
  const g = (await page.getByTestId('grid').boundingBox())!;
  const b = (await bands(page)).filter((x) => x.label.startsWith(labelStart) && !x.testid.endsWith('-pin'));
  expect(b.length, `band ${labelStart}`).toBe(1);
  return Math.abs(b[0].x + b[0].w / 2 - (g.x + g.width / 2)) / g.width;
}

/** press on a card, move past the drag threshold, then to the drop position (before `target`, or after the last card) */
async function dragCard(page: Page, id: string, target: string | null, opts: { release?: boolean; shotName?: string } = {}) {
  const from = (await card(page, id).boundingBox())!;
  const sx = from.x + 40; // the title line (no control there)
  const sy = from.y + 14;
  let tx: number;
  let ty: number;
  if (target) {
    const t = (await card(page, target).boundingBox())!;
    tx = t.x + 26;
    ty = t.y + 30;
  } else {
    const add = (await page.getByTestId('drop-zone').boundingBox())!;
    tx = add.x - 3;
    ty = add.y + 30;
  }
  await page.mouse.move(sx, sy);
  await page.mouse.down();
  await page.mouse.move(sx - 6, sy + 6, { steps: 3 }); // past the 5 px threshold
  await page.mouse.move(tx, ty, { steps: 14 });
  await page.waitForTimeout(80);
  if (opts.shotName) await shot(page, opts.shotName);
  if (opts.release !== false) {
    await page.mouse.up();
    await page.waitForTimeout(250);
  }
}

/** drop real files like the OS does */
async function osDrop(page: Page, paths: string[]) {
  const cdp = await page.context().newCDPSession(page);
  const data = { items: [], files: paths.map((p) => path.resolve(p)), dragOperationsMask: 1 };
  for (const type of ['dragEnter', 'dragOver', 'drop'] as const) await cdp.send('Input.dispatchDragEvent', { type, x: 600, y: 500, data });
  await cdp.detach();
}

// ------------------------------------------------------------------------------------------------ cards before the merge

test.describe('カードの並べ替え（結合前）', () => {
  test('ドラッグで順番が入れ替わり、結合キー・取り込む列の選択はそのまま、取り込む列の並びは新しい順になる', async ({ page }) => {
    const errors = await start(page);
    expect(await cardIds(page)).toEqual(['s1', 's2', 's3', 's4', 's5']);
    // user choices that a pure reorder must keep: the join type (「他ファイルにしかない行も末尾に追加」), one column of 02 left out
    await page.getByTestId('join-type-select').selectOption('full');
    await settle(page);
    const chipBefore = await page.getByTestId('key-chip-name').getAttribute('aria-pressed');
    const firstChoice = page.locator('.import-group').first().locator('.col-chip.on').first();
    const choiceText = (await firstChoice.textContent()) ?? '';
    await firstChoice.click();
    await expect(page.locator('.import-group').first().locator('.col-chip', { hasText: choiceText.replace('＋ ', '') }).first()).toHaveAttribute('aria-pressed', 'false');
    const summaryBefore = await page.getByTestId('summary-line').innerText();
    const groupsBefore = await plannedGroups(page);
    const c2 = await cardColor(page, 's2');
    const c5 = await cardColor(page, 's5');
    expect(groupsBefore.indexOf(c2)).toBeLessThan(groupsBefore.indexOf(c5));

    // 05 before 02
    await dragCard(page, 's5', 's2', { shotName: 'cards-dragging' });
    expect(await cardIds(page)).toEqual(['s1', 's5', 's2', 's3', 's4']);
    await settle(page);
    await expect(page.getByTestId('merge-button')).toBeEnabled();

    // kept: key plan, join type, the choice of columns, the live summary (same files, same keys)
    expect(await page.getByTestId('key-chip-name').getAttribute('aria-pressed')).toBe(chipBefore);
    await expect(page.getByTestId('join-type-select')).toHaveValue('full');
    expect(await page.getByTestId('summary-line').innerText()).toBe(summaryBefore);
    await expect(page.locator('.import-group').filter({ hasText: '02_面接評価' }).locator('.col-chip', { hasText: choiceText.replace('＋ ', '') }).first()).toHaveAttribute('aria-pressed', 'false');
    // recomputed: the imported columns are listed in the new order (05 before 02), the merge preview rows too
    const groupsAfter = await plannedGroups(page);
    expect(groupsAfter.indexOf(c5)).toBeLessThan(groupsAfter.indexOf(c2));
    const previewRows = await page.locator('.check-table tbody tr').allInnerTexts();
    expect(previewRows.findIndex((t) => t.includes('05_研修出欠'))).toBeLessThan(previewRows.findIndex((t) => t.includes('02_面接評価')));
    // a drag is not a click: the card is not "picked", nothing got selected
    await expect(page.locator('.is-picked')).toHaveCount(0);

    // the merge follows the new order
    await mergeNow(page);
    const grouped = await listGroups(page);
    expect(grouped.indexOf(c5)).toBeGreaterThan(0);
    expect(grouped.indexOf(c5)).toBeLessThan(grouped.indexOf(c2));
    expect(errors).toEqual([]);
  });

  test('挿入バーとゴースト。Esc で取り消せる。ボタン・選択欄から押しても始まらない。ファイルのドロップ表示は出ない', async ({ page }) => {
    const errors = await start(page);
    const from = (await card(page, 's4').boundingBox())!;
    const t = (await card(page, 's2').boundingBox())!;
    await page.mouse.move(from.x + 40, from.y + 14);
    await page.mouse.down();
    await page.mouse.move(from.x + 30, from.y + 22, { steps: 3 });
    await page.mouse.move(t.x + 30, t.y + 40, { steps: 10 });
    await expect(page.getByTestId('card-drop-indicator')).toHaveCount(1);
    await expect(page.getByTestId('card-ghost')).toContainText('04_適性検査.xls を移動');
    // the bar stands in front of the card under the pointer
    const bar = (await page.getByTestId('card-drop-indicator').boundingBox())!;
    expect(Math.abs(bar.x + 2 - t.x)).toBeLessThan(14);
    await expect(card(page, 's4')).toHaveAttribute('data-card-dragging', '');
    // it is NOT a file drop: no drop overlay, no dashed card highlight
    await expect(page.locator('.drop-overlay')).toHaveCount(0);
    await expect(page.locator('.add-card.dragging')).toHaveCount(0);
    await page.keyboard.press('Escape');
    await expect(page.getByTestId('card-drop-indicator')).toHaveCount(0);
    await expect(page.getByTestId('card-ghost')).toHaveCount(0);
    await page.mouse.up();
    expect(await cardIds(page)).toEqual(['s1', 's2', 's3', 's4', 's5']);
    await expect(page.locator('.is-picked')).toHaveCount(0); // the release after Esc is no click either

    // controls do not start a drag
    for (const control of [page.getByTestId('base-toggle-2'), page.getByTestId('file-clear-3'), page.getByTestId('last-row-input-4')]) {
      const b = (await control.boundingBox())!;
      await page.mouse.move(b.x + b.width / 2, b.y + b.height / 2);
      await page.mouse.down();
      await page.mouse.move(b.x + b.width / 2 - 80, b.y + 60, { steps: 6 });
      await expect(page.getByTestId('card-ghost')).toHaveCount(0);
      await page.mouse.up();
    }
    await page.keyboard.press('Escape');
    expect(await cardIds(page)).toEqual(['s1', 's2', 's3', 's4', 's5']);
    expect(errors).toEqual([]);
  });

  test('最後のカードの後ろ（点線のカードの手前）に落とせる。点線のカードは動かせず、落とし先でもない', async ({ page }) => {
    const errors = await start(page);
    await dragCard(page, 's1', null); // after the last card
    expect(await cardIds(page)).toEqual(['s2', 's3', 's4', 's5', 's1']);
    await settle(page);
    // the dashed card stays the last element and does not react to a drag
    const add = (await page.getByTestId('drop-zone').boundingBox())!;
    await page.mouse.move(add.x + 30, add.y + 30);
    await page.mouse.down();
    await page.mouse.move(add.x - 200, add.y + 40, { steps: 8 });
    await expect(page.getByTestId('card-ghost')).toHaveCount(0);
    await page.mouse.up();
    expect(await cardIds(page)).toEqual(['s2', 's3', 's4', 's5', 's1']);
    // dropping a card on itself / right where it is changes nothing
    await dragCard(page, 's3', 's3');
    await dragCard(page, 's3', 's4');
    expect(await cardIds(page)).toEqual(['s2', 's3', 's4', 's5', 's1']);
    expect(errors).toEqual([]);
  });

  test('行がはみ出すとき、端に近づけると横にスクロールする', async ({ page }) => {
    const errors = await start(page);
    const row = page.getByTestId('file-cards');
    const overflow = await row.evaluate((e) => e.scrollWidth - e.clientWidth);
    test.skip(overflow < 40, 'the cards fit into the row at this width');
    const rb = (await row.boundingBox())!;
    const c = (await card(page, 's1').boundingBox())!;
    await page.mouse.move(c.x + 40, c.y + 14);
    await page.mouse.down();
    await page.mouse.move(c.x + 50, c.y + 20, { steps: 3 });
    await page.mouse.move(rb.x + rb.width - 20, c.y + 30, { steps: 10 });
    await expect.poll(() => row.evaluate((e) => e.scrollLeft), { timeout: 8000 }).toBeGreaterThan(100);
    await expect(page.getByTestId('card-drop-indicator')).toHaveCount(1);
    await page.keyboard.press('Escape');
    await page.mouse.up();
    expect(await cardIds(page)).toEqual(['s1', 's2', 's3', 's4', 's5']);
    expect(errors).toEqual([]);
  });

  test('キーボード: カードにフォーカスして Alt+→ / Alt+← で動かせ、フォーカスは残り、結果が読み上げ用に通知される', async ({ page }) => {
    const errors = await start(page);
    await card(page, 's2').focus();
    await page.keyboard.press('Alt+ArrowRight');
    await expect(page.getByTestId('card-live')).toContainText('「02_面接評価.xlsx」を 3 番目に移動しました（全 5 件）');
    expect(await cardIds(page)).toEqual(['s1', 's3', 's2', 's4', 's5']);
    // (a card cannot be moved while the key analysis that follows a move is still running: wait until 「結合する」 is enabled again;
    //  without it the next key press could arrive too early and be ignored — a race that failed this test at the round-9 baseline)
    await expect(page.getByTestId('merge-button')).toBeEnabled();
    await settle(page);
    await expect(card(page, 's2')).toBeFocused();
    await page.keyboard.press('Alt+ArrowLeft');
    await expect(page.getByTestId('merge-button')).toBeEnabled();
    await settle(page);
    await expect(card(page, 's2')).toBeFocused();
    await page.keyboard.press('Alt+ArrowLeft');
    await expect.poll(() => cardIds(page)).toEqual(['s2', 's1', 's3', 's4', 's5']);
    await expect(page.getByTestId('merge-button')).toBeEnabled();
    await settle(page);
    expect(await cardIds(page)).toEqual(['s2', 's1', 's3', 's4', 's5']);
    await expect(page.getByTestId('card-live')).toContainText('1 番目');
    await expect(card(page, 's2')).toBeFocused();
    await page.keyboard.press('Alt+ArrowLeft'); // already first
    await expect(page.getByTestId('card-live')).toContainText('先頭');
    // inside a text field Alt+arrows are left alone
    await settle(page);
    await page.getByTestId('last-row-input-2').focus();
    await page.keyboard.press('Alt+ArrowRight');
    expect(await cardIds(page)).toEqual(['s2', 's1', 's3', 's4', 's5']);
    expect(errors).toEqual([]);
  });

  test('クリックは強調表示だけ。基準の切り替え・シート／見出し行の選択・クリアは今まで通り動く', async ({ page }) => {
    const errors = await start(page);
    await card(page, 's2').click({ position: { x: 40, y: 14 } });
    await expect(card(page, 's2')).toHaveClass(/is-picked/);
    await card(page, 's4').click({ position: { x: 40, y: 14 } });
    await expect(card(page, 's4')).toHaveClass(/is-picked/);
    await expect(card(page, 's2')).not.toHaveClass(/is-picked/);
    // header-row select keeps working (and is not a card click)
    await page.getByTestId('header-row-select-4').selectOption({ index: 0 });
    await settle(page);
    // 基準にする / クリア
    await page.getByTestId('base-toggle-1').click();
    await settle(page);
    await expect(card(page, 's2')).toHaveClass(/is-base/);
    await page.getByTestId('file-clear-4').click();
    await settle(page);
    expect(await cardIds(page)).toHaveLength(4);
    expect(errors).toEqual([]);
  });

  test('OSからのファイルのドロップは、カードを並べ替えたあとも今まで通り使える', async ({ page }) => {
    const errors = await openApp(page);
    await addFiles(page, FILES.slice(0, 4));
    await dragCard(page, 's4', 's2');
    expect(await cardIds(page)).toEqual(['s1', 's4', 's2', 's3']);
    await settle(page);
    await osDrop(page, [FILES[4]]);
    await settle(page);
    await expect(page.locator('[data-file-id]')).toHaveCount(5);
    expect(await cardIds(page)).toEqual(['s1', 's4', 's2', 's3', 's5']); // the new file joins at the end
    await expect(card(page, 's5')).toContainText('05_研修出欠.xlsx');
    expect(errors).toEqual([]);
  });
});

// ------------------------------------------------------------------------------------------------ after the merge

test.describe('結合後: カードのクリック・ドラッグ、帯のクリック・ドラッグ', () => {
  test('カードをクリックすると、そのファイルの列が表の中央付近に来て、列がすべて選択される（基準のカードも）', async ({ page }) => {
    const errors = await start(page, true);
    // display order = the order of the groups in the grid; 03 brings no column and goes last
    expect(await cardIds(page)).toEqual(['s1', 's2', 's4', 's5', 's3']);
    const all = await bands(page);
    expect(all.find((b) => b.label.startsWith('01_学生名簿'))?.count).toBe(12);
    await expect(page.locator('[data-file-id].is-active')).toHaveCount(0);

    // 05 is the 4th group: far to the right of the first screen
    expect(await gridScrollLeft(page)).toBe(0);
    await card(page, 's5').click({ position: { x: 40, y: 14 } });
    const sl = await scrollSettled(page);
    expect(sl).toBeGreaterThan(300);
    expect(await centreError(page, '05_研修出欠')).toBeLessThan(0.15);
    expect(await activeCards(page)).toEqual(['s5']);
    expect(await selectedCount(page)).toBe(5); // all 5 columns of 05
    expect(await page.evaluate(() => document.activeElement?.getAttribute('data-testid'))).toBe('grid-editor'); // Ctrl+C / arrows work
    expect(await page.getByTestId('grid').evaluate((e) => e.scrollTop)).toBe(0);
    await shot(page, 'card-click-centred');

    // 02 (7 columns): replaces the selection, the card highlight follows
    await card(page, 's2').click({ position: { x: 40, y: 14 } });
    await scrollSettled(page);
    expect(await centreError(page, '02_面接評価')).toBeLessThan(0.15);
    expect(await activeCards(page)).toEqual(['s2']);
    expect(await selectedCount(page)).toBe(7);

    // the base (12 columns, 2 of them frozen): all 12 selected, no horizontal scroll needed for the frozen part
    await card(page, 's1').click({ position: { x: 40, y: 14 } });
    await scrollSettled(page);
    expect(await activeCards(page)).toEqual(['s1']);
    expect(await selectedCount(page)).toBe(12);

    // 03 has no column in the table: says so, selects nothing, cannot be dragged
    await card(page, 's3').click({ position: { x: 40, y: 14 } });
    await expect(page.getByTestId('toast')).toContainText('「03_連絡先.csv」の列は表にありません');
    expect(await activeCards(page)).toEqual(['s1']);
    await expect(card(page, 's3')).not.toHaveAttribute('data-reorderable', '');
    await expect(card(page, 's1')).toHaveAttribute('data-reorderable', '');
    expect(errors).toEqual([]);
  });

  test('カードのクリックのあと、矢印キー・Ctrl+C がグリッドで使える（フォーカスがグリッドにある）', async ({ page }) => {
    const errors = await start(page, true);
    await card(page, 's2').click({ position: { x: 40, y: 14 } });
    await scrollSettled(page);
    const addr0 = await page.getByTestId('formula-bar-address').textContent();
    await page.keyboard.press('ArrowDown');
    const addr1 = await page.getByTestId('formula-bar-address').textContent();
    expect(addr1).not.toBe(addr0);
    expect(addr1?.replace(/\d+$/, '')).toBe(addr0?.replace(/\d+$/, '')); // same column, next row
    expect(errors).toEqual([]);
  });

  test('カードをドラッグすると、そのファイルの列がまとめて動く（元に戻す1回で戻る）。同じ場所に落としても何も起きない', async ({ page }) => {
    const errors = await start(page, true);
    const n0 = await editCount(page);
    expect(n0).toBe(0);
    expect(await cardIds(page)).toEqual(['s1', 's2', 's4', 's5', 's3']);

    // 05 before 02
    await dragCard(page, 's5', 's2', { shotName: 'cards-merged-dragging' });
    expect(await cardIds(page)).toEqual(['s1', 's5', 's2', 's4', 's3']);
    expect(await editCount(page)).toBe(1);
    expect(await activeCards(page)).toEqual(['s5']); // the moved group is selected
    expect(await selectedCount(page)).toBe(5);
    await expect(page.getByTestId('undo-button')).toHaveAttribute('title', /元に戻す: 05_研修出欠の列を移動/);
    // the grid shows the moved group (centred, or at least in the viewport)
    await scrollSettled(page);
    expect(await centreError(page, '05_研修出欠')).toBeLessThan(0.15);
    const labels = (await bands(page)).filter((b) => !b.testid.endsWith('-pin')).map((b) => b.label);
    expect(labels.findIndex((l) => l.startsWith('05_研修出欠'))).toBeGreaterThanOrEqual(0);
    expect(labels.findIndex((l) => l.startsWith('05_研修出欠'))).toBeLessThan(labels.findIndex((l) => l.startsWith('02_面接評価')));

    // dropped where it is: nothing
    await dragCard(page, 's5', 's2');
    expect(await editCount(page)).toBe(1);

    // Ctrl+Z (focus is in the grid or anywhere) restores both the columns and the card order; Ctrl+Y does it again
    await page.keyboard.press('Control+z');
    await expect.poll(() => cardIds(page)).toEqual(['s1', 's2', 's4', 's5', 's3']);
    expect(await editCount(page)).toBe(0);
    await page.keyboard.press('Control+y');
    await expect.poll(() => cardIds(page)).toEqual(['s1', 's5', 's2', 's4', 's3']);

    // the base group to the very end of the groups
    await dragCard(page, 's1', null);
    const unmerge = page.getByTestId('unmerge-confirm-dialog');
    if (await unmerge.isVisible().catch(() => false)) {
      // the title merge A1:K1 is in the way: the same confirmation as for a column drag, then ONE undo step
      await expect(unmerge).toContainText('結合を解除して移動');
      await unmerge.getByRole('button', { name: '結合を解除して移動' }).click();
    }
    await expect.poll(() => cardIds(page)).toEqual(['s5', 's2', 's4', 's1', 's3']);
    expect(await editCount(page)).toBe(2);
    expect(errors).toEqual([]);
  });

  test('帯をクリックするとそのファイルの列がすべて選択され、カードも選択中の見た目になる。Ctrl+クリックで足せる', async ({ page }) => {
    const errors = await start(page, true);
    const b02 = await revealBand(page, '02_面接評価');
    await b02.click({ position: { x: 24, y: 8 } });
    expect(await selectedCount(page)).toBe(7);
    expect(await activeCards(page)).toEqual(['s2']);
    await shot(page, 'band-selected');
    const b04 = await revealBand(page, '04_適性検査');
    await b04.click({ position: { x: 24, y: 8 }, modifiers: ['Control'] });
    expect(await selectedCount(page)).toBeGreaterThan(7);
    expect(await activeCards(page)).toEqual([]); // two files selected: no single active card
    await page.getByTestId('grid').evaluate((e) => (e.scrollLeft = 0));
    await page.waitForTimeout(150);
    await page.getByTestId('grid-band-0').click({ position: { x: 24, y: 8 } });
    expect(await selectedCount(page)).toBe(12);
    expect(await activeCards(page)).toEqual(['s1']);
    expect(errors).toEqual([]);
  });

  test('帯をドラッグすると、そのファイルの列がまとめて動く。表のタイトル結合に掛かる移動は、列のドラッグと同じ確認が出る', async ({ page }) => {
    await page.setViewportSize({ width: 1920, height: 1000 }); // 02 and 04 (7 columns each) side by side
    const errors = await start(page, true);
    // 04 before 02 (both bands in view): a click on the 02 card brings its group to the middle of the grid
    await card(page, 's2').click({ position: { x: 40, y: 14 } });
    await scrollSettled(page);
    const b02 = (await page.locator('.g-band:not([data-testid$=-pin])', { hasText: '02_面接評価' }).first().boundingBox())!;
    const b04 = page.locator('.g-band:not([data-testid$=-pin])', { hasText: '04_適性検査' }).first();
    await expect(b04).toBeVisible();
    const b4 = (await b04.boundingBox())!;
    const g = (await page.getByTestId('grid').boundingBox())!;
    const first02 = b02.x + 8; // the left half of 02's first column
    await page.mouse.move(b4.x + 40, b4.y + 8);
    await page.mouse.down();
    await page.mouse.move(b4.x + 20, b4.y + 14, { steps: 3 });
    await page.mouse.move(Math.max(first02, g.x + 300), g.y + 160, { steps: 14 });
    await page.waitForTimeout(80);
    await expect(page.locator('.g-ghost')).toContainText(/04_適性検査（\d+列）を移動/);
    await shot(page, 'band-drag');
    await page.mouse.up();
    await page.waitForTimeout(300);
    expect(await editCount(page)).toBe(1);
    expect(await cardIds(page)).toEqual(['s1', 's4', 's2', 's5', 's3']);
    expect(await activeCards(page)).toEqual(['s4']);
    await page.keyboard.press('Control+z');
    await expect.poll(() => cardIds(page)).toEqual(['s1', 's2', 's4', 's5', 's3']);
    expect(await editCount(page)).toBe(0);

    // a card drop puts 02 in front of the base (no refusal: the edge of the title merge) ...
    await dragCard(page, 's2', 's1');
    expect(await cardIds(page)).toEqual(['s2', 's1', 's4', 's5', 's3']);
    expect(await editCount(page)).toBe(1);
    // ... and a band drop INTO the base group (before フリガナ) is in the way of the title merge A1:K1: the same
    // 「結合を解除して移動」 question as for a column drag, then ONE more undo step
    await page.getByTestId('grid').evaluate((e) => (e.scrollLeft = 0));
    await page.waitForTimeout(250);
    // (the frozen pane now holds 02's 7 columns + 学生ID + 氏名: the doc moved the freeze with the columns)
    const pin = (await page.getByTestId('grid-band-0').boundingBox())!;
    const furi = page.locator('.g-header .g-hcell', { has: page.locator('.g-htitle', { hasText: /^フリガナ$/ }) }).first();
    await expect(furi).toBeVisible();
    const fb = (await furi.boundingBox())!;
    await page.mouse.move(pin.x + 60, pin.y + 8);
    await page.mouse.down();
    await page.mouse.move(pin.x + 80, pin.y + 14, { steps: 3 });
    await page.mouse.move(fb.x + 10, fb.y + 90, { steps: 14 });
    await page.mouse.up();
    const dlg = page.getByTestId('unmerge-confirm-dialog');
    await expect(dlg).toBeVisible();
    await dlg.getByRole('button', { name: '結合を解除して移動' }).click();
    await expect(dlg).toHaveCount(0);
    expect(await editCount(page)).toBe(2);
    await page.keyboard.press('Control+z');
    await expect.poll(() => editCount(page)).toBe(1);
    expect(errors).toEqual([]);
  });

  test('列見出しのドラッグ（従来の列の移動）はそのまま動く', async ({ page }) => {
    const errors = await start(page, true);
    await card(page, 's2').click({ position: { x: 40, y: 14 } }); // brings 02's 7 columns into view (and selects them)
    await scrollSettled(page);
    const before = await headerTitles(page);
    const t = before.slice(12, 19); // 02's columns
    const h = (name: string) => page.locator('.g-header .g-hcell', { has: page.locator('.g-htitle', { hasText: new RegExp(`^${name}$`) }) }).first();
    await h(t[0]).click(); // plain click: the selection is just this column
    await h(t[1]).click({ modifiers: ['Control'] });
    const src = (await h(t[0]).boundingBox())!;
    const dst = (await h(t[3]).boundingBox())!;
    await page.mouse.move(src.x + src.width / 2, src.y + 20);
    await page.mouse.down();
    await page.mouse.move(src.x + src.width / 2 + 30, src.y + 24, { steps: 3 });
    await page.mouse.move(dst.x + dst.width - 6, dst.y + 80, { steps: 12 });
    await page.mouse.up();
    await page.waitForTimeout(300);
    const after = await headerTitles(page);
    expect(after.slice(12, 19)).toEqual([t[2], t[3], t[0], t[1], t[4], t[5], t[6]]);
    expect(await editCount(page)).toBe(1);
    await page.keyboard.press('Control+z');
    await expect.poll(() => editCount(page)).toBe(0);
    expect(await headerTitles(page)).toEqual(before);
    expect(errors).toEqual([]);
  });
});

// ------------------------------------------------------------------------------------------------ single sheet output

test.describe('他のシートも出力に含める', () => {
  test('既定はオフ（結合した1シートだけを出力）。結合後は変えられず、結果にそのことが出る', async ({ page }) => {
    const errors = await start(page);
    const box = page.getByTestId('include-other-sheets');
    await expect(box).not.toBeChecked();
    await expect(box.locator('xpath=ancestor::label')).toHaveAttribute('title', '通常は、結合した1つのシートだけを出力します。基準ファイルにあるほかのシート（集計・グラフなど）も残したいときにチェックします');
    // next to 結合方法
    const join = (await page.getByTestId('join-type-select').boundingBox())!;
    const cb = (await box.boundingBox())!;
    expect(cb.x).toBeGreaterThan(join.x);
    expect(Math.abs(cb.y - join.y)).toBeLessThan(40);
    // the merge preview says what is left out
    await expect(page.getByTestId('preview-sheets')).toContainText('出力は結合した 1 シートのみです（他のシート 3 枚「集計」「ピボット」「マスタ」');

    await mergeNow(page);
    await expect(box).toBeDisabled();
    await expect(box).not.toBeChecked();
    await expect(page.getByTestId('side-dropped-sheets')).toContainText('出力は結合した 1 シートのみです（他のシート 3 枚「集計」「ピボット」「マスタ」');
    await expect(page.getByTestId('side-formulas-detached')).toContainText('他のシートを参照していた数式');
    await expect(page.getByTestId('toast')).toContainText('他のシート 3 枚は含めません');
    // the 警告 list carries the same lines and no longer promises to keep the objects that lived on the dropped sheets
    await openWarnings(page);
    const warn = await page.getByTestId('warning-list').innerText();
    expect(warn).toContain('出力は結合した 1 シートのみです');
    expect(warn).not.toMatch(/グラフ 1 個: /);
    expect(warn).not.toContain('ピボットテーブル');

    // 結合をやり直す: the checkbox is enabled again and keeps its value; ticked: nothing is dropped
    await page.getByTestId('remerge-button').click();
    await expect(box).toBeEnabled();
    await box.check();
    // (round 9) ticking it also ticks 「他のファイルのシートも含める」; this test is about the BASE file's sheets only: clear that one again
    await expect(page.getByTestId('include-source-sheets')).toBeChecked();
    await page.getByTestId('include-source-sheets').uncheck();
    await expect(box).toBeChecked();
    await settle(page);
    await expect(page.getByTestId('preview-sheets')).toContainText('そのまま出力に含めます');
    await mergeNow(page);
    await expect(box).toBeChecked();
    await expect(page.getByTestId('side-dropped-sheets')).toHaveCount(0);
    await openWarnings(page);
    await expect(page.getByTestId('warning-list')).toContainText(/グラフ 1 個: /);
    expect(errors).toEqual([]);
  });
});
