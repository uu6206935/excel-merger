// Full flow of the REAL app (no mock) through the real local server, with the samples:
//   PORT=5192 NO_OPEN=1 node server/server.mjs          (another terminal; DIST_DIR / OUTPUT_DIR may point elsewhere)
//   node tests/ui/real.mjs [baseUrl] [--js-only] [--insert-after=<見出し>] [--prefix=<name>] [--sheets=single|base|all] [--no-shots] [file ...]
// Loads the files through the file input, prints what the cards / summary / check lists show, merges, moves
// columns (an imported one via the side-panel column list, a base one across the merged title -> the
// 「結合を解除して移動」 confirmation), inserts a column, types a date / % / =SUM / =UNIQUE, deletes a column that
// formulas depend on (convert to values), exports with Excel本体 and with the built-in engine, prints the saved
// paths as `OUT excel <path>` / `OUT js <path>` lines (and `SHEETS <mode>`: which sheets the output holds, for verify-outputs.mjs).
// --sheets=single (default, the app's default: ONLY the merged sheet) | base (「他のシートも出力に含める」 only: the base's other sheets) |
// all (both boxes: the base's sheets + the original sheets of every other file, <file>_<sheet>).
// Screenshots: tests/ui/shots/real-*.png. Exit 1 on a failed check or a console error.
import { chromium } from '@playwright/test';
import { mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const outDir = join(here, 'shots');
mkdirSync(outDir, { recursive: true });
const args = process.argv.slice(2);
const flags = new Set(args.filter((a) => a.startsWith('--')));
const pos = args.filter((a) => !a.startsWith('--'));
const base = pos[0] ?? 'http://localhost:5192/';
const samples = join(here, '..', '..', 'samples');
const files = pos.length > 1 ? pos.slice(1) : ['01_学生名簿.xlsx', '02_面接評価.xlsx', '03_連絡先.csv', '04_適性検査.xls', '05_研修出欠.xlsx'].map((f) => join(samples, f));
const excelWanted = !flags.has('--js-only');
// --insert-after=<header>: where the new column goes (default 志望度: Excel extends its CF / DV over the new column)
const insertAfter = [...flags].find((f) => f.startsWith('--insert-after='))?.slice('--insert-after='.length) || '志望度';
// --prefix=<name>: output file names <prefix>_excel.xlsx / <prefix>_js.xlsx
const prefix = [...flags].find((f) => f.startsWith('--prefix='))?.slice('--prefix='.length) || 'u2';
const noShots = flags.has('--no-shots');
// --sheets=single|base|all: which sheets the output holds (the two checkboxes of the settings bar, set before the merge)
const sheetsMode = [...flags].find((f) => f.startsWith('--sheets='))?.slice('--sheets='.length) || 'single';
if (!['single', 'base', 'all'].includes(sheetsMode)) throw new Error(`--sheets must be single, base or all (got ${sheetsMode})`);

const browser = await chromium.launch({ channel: 'msedge', headless: true });
const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 }, acceptDownloads: true, locale: 'ja-JP' });
const page = await ctx.newPage();
const errors = [];
const failures = [];
page.on('console', (m) => m.type() === 'error' && errors.push(m.text()));
page.on('pageerror', (e) => errors.push('pageerror: ' + e.message));
const text = async (sel) => ((await page.locator(sel).count()) ? (await page.locator(sel).first().innerText()).replace(/\s+/g, ' ').trim() : null);
const idle = async () => {
  await page.waitForFunction(() => !document.querySelector('[data-testid=busy-overlay]'), null, { timeout: 180000 });
  await page.waitForTimeout(200);
};
const check = (cond, msg) => {
  if (cond) console.log('ok:', msg);
  else {
    failures.push(msg);
    console.log('FAIL:', msg);
  }
};
const shot = async (name) => {
  if (noShots) return;
  await page.waitForTimeout(250);
  await page.screenshot({ path: join(outDir, `${name}.png`) });
  console.log('saved', name);
};
const toast = () => text('[data-testid=toast]');

// ---------------------------------------------------------------------------------------------- files
await page.goto(base);
await page.waitForSelector('[data-testid=server-status]');
await page.waitForTimeout(500);
console.log('status:', await text('[data-testid=server-status]'), '|', await text('.app-header'));
const t0 = Date.now();
await page.setInputFiles('[data-testid=file-input]', files);
await page.waitForFunction((n) => document.querySelectorAll('[data-testid^=file-card-]').length >= n, files.length, { timeout: 60000 });
await page.waitForFunction(() => !document.querySelector('.card-loading'), null, { timeout: 180000 });
await page.waitForTimeout(300);
await idle();
await page.waitForSelector('[data-testid=summary-line]', { timeout: 60000 }).catch(() => undefined);
await idle();
console.log(`files read + analysed in ${Date.now() - t0} ms`);
for (let i = 0; i < files.length; i++) {
  const card = `[data-testid=file-card-${i}]`;
  console.log(
    `card ${i}:`,
    await text(`${card} .card-head`),
    '|',
    await text(`${card} .file-info`),
    '|',
    await text(`[data-testid=file-keys-${i}]`),
    '| header row',
    await page.locator(`[data-testid=header-row-select-${i}]`).inputValue().catch(() => '?'),
    '| last row',
    await page.locator(`[data-testid=last-row-input-${i}]`).inputValue().catch(() => '?'),
    '|',
    (await text(`${card} .card-error`)) ?? '',
  );
}
const summaryLine = await text('[data-testid=summary-line]');
console.log('summary:', summaryLine);
console.log('side summary:', await text('[data-testid=side-summary]'));
check(/^60 行/.test((await text('[data-testid=file-card-0] .count-badge')) ?? ''), '01_学生名簿: 60 行 (footer rows 人数/平均 are not data)');
check(/基準 60 名/.test(summaryLine ?? ''), `summary 基準 60 名 (${summaryLine})`);
check((await text('[data-testid=file-encoding-2]')) === 'Shift_JIS', `03_連絡先.csv shows the encoding tag Shift_JIS (${await text('[data-testid=file-encoding-2]')})`);
// check lists (before the merge: groups + summary lines for the other files' rows)
const checkText = await text('[data-testid=check-panel]');
console.log('check panel:', checkText);
for (const d of await page.locator('[data-testid=check-panel] details').all()) await d.evaluate((el) => (el.open = true));
const checkItems = await page.locator('[data-testid=check-panel] li').allInnerTexts();
console.log('check items:\n  ' + checkItems.map((s) => s.replace(/\s+/g, ' ')).join('\n  '));
const unmatchedCount = Number((await text('[data-testid=check-unmatched-count]')) ?? '0');
check(!checkItems.some((s) => /人数|平均|合計/.test(s) && /未一致|どのファイル/.test(s)), 'no fake 未一致 rows from the footer (人数・平均)');
console.log('unmatched count:', unmatchedCount, 'toast:', await toast());
await shot('real-01-files');

// key detail modal (also: scrollable + OK reachable at 1280x720)
if (await page.locator('[data-testid=key-detail-button]').count()) {
  await page.click('[data-testid=key-detail-button]');
  await page.waitForSelector('[data-testid=key-detail-modal]');
  check((await page.evaluate(() => document.activeElement?.getAttribute('data-testid'))) === 'key-detail-ok', 'key detail: initial focus on OK');
  await shot('real-02-key-detail');
  await page.setViewportSize({ width: 1280, height: 720 });
  await page.waitForTimeout(300);
  const tabs = page.locator('[data-testid^=key-detail-tab-]');
  const nTabs = await tabs.count();
  for (let t = nTabs - 1; t >= 0; t--) {
    if (nTabs) await tabs.nth(t).click();
    await page.waitForTimeout(200);
    const m = await page.evaluate(() => {
      const body = document.querySelector('[data-testid=key-detail-modal] .modal-body');
      const ok = document.querySelector('[data-testid=key-detail-ok]').getBoundingClientRect();
      return { scrollable: body.scrollHeight > body.clientHeight, overflowY: getComputedStyle(body).overflowY, okVisible: ok.bottom <= innerHeight && ok.top >= 0 };
    });
    check(m.okVisible && (m.overflowY === 'auto' || m.overflowY === 'scroll'), `key detail tab ${t} at 1280x720: body scrolls, OK visible ${JSON.stringify(m)}`);
    if (m.scrollable && t === 1) await shot('real-02c-key-detail-1280x720-tab2');
  }
  await shot('real-02b-key-detail-1280x720');
  await page.keyboard.press('Escape');
  await page.waitForTimeout(200);
  check((await page.locator('[data-testid=key-detail-modal]').count()) === 0, 'Esc closes the key detail modal');
  await page.setViewportSize({ width: 1440, height: 900 });
  await page.waitForTimeout(300);
}

// ---------------------------------------------------------------------------------------------- merge
// the sheets of the output: the app's default is the merged sheet only; 「他のシートも出力に含める」 also ticks 「他のファイルのシートも含める」
if (sheetsMode !== 'single' && (await page.locator('[data-testid=include-other-sheets]').count())) {
  await page.check('[data-testid=include-other-sheets]');
  if (sheetsMode === 'base') await page.uncheck('[data-testid=include-source-sheets]');
  check(
    (await page.isChecked('[data-testid=include-other-sheets]')) && (await page.isChecked('[data-testid=include-source-sheets]')) === (sheetsMode === 'all'),
    `sheet options set to "${sheetsMode}"`,
  );
}
console.log(`SHEETS ${sheetsMode}`);
if (!(await page.locator('[data-testid=merge-button]:not([disabled])').count())) {
  console.log('merge button disabled:', await page.locator('[data-testid=merge-button]').getAttribute('title'));
  failures.push('merge button disabled');
} else {
  const t1 = Date.now();
  await page.click('[data-testid=merge-button]');
  await page.waitForSelector('[data-testid=grid]', { timeout: 180000 });
  await idle();
  console.log(`merge in ${Date.now() - t1} ms; toast:`, await toast());
  console.log('side summary:', await text('[data-testid=side-summary]'));
  await shot('real-03-merged');

  const titles = async () => (await page.locator('[data-testid=column-list] li .col-title').allInnerTexts()).map((s) => s.trim());
  const idxOf = async (t) => (await titles()).indexOf(t);
  const hdr = (i) => page.locator(`[data-testid=grid-header-${i}]`);
  const hdrTitle = async (i) => ((await hdr(i).count()) ? (await hdr(i).locator('.g-htitle').innerText()).trim() : null);
  const letter = (i) => {
    let s = '';
    let x = i + 1;
    while (x > 0) {
      const r = (x - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      x = Math.floor((x - 1) / 26);
    }
    return s;
  };
  /** select a column through the side list (the grid scrolls to it) */
  const reveal = async (i) => {
    await page.locator(`[data-testid=column-list-item-${i}]`).click();
    await page.waitForTimeout(600);
  };
  /** header context menu -> item (the menu closes when the grid scrolls: retry) */
  const menu = async (i, key) => {
    for (let attempt = 0; attempt < 3; attempt++) {
      await hdr(i).click({ button: 'right' });
      const item = page.locator(`[data-testid=context-menu-item-${key}]`);
      try {
        await item.waitFor({ state: 'visible', timeout: 3000 });
        await item.click({ timeout: 3000 });
        return true;
      } catch {
        await page.keyboard.press('Escape');
        await page.waitForTimeout(400);
      }
    }
    return false;
  };
  const editsCount = async () => Number(((await text('[data-testid=side-summary]')) ?? '').match(/変更 (\d+) 件/)?.[1] ?? NaN);
  const before = await titles();
  console.log('columns after merge:', before.join(' | '));

  // 1. move an imported column with the side-panel list (outside the title: moves freely)
  const firstImported = before.length > 13 ? 13 : before.length - 1;
  const movedTitle = before[firstImported + 2];
  if (movedTitle) {
    await page.locator(`[data-testid=column-list-item-${firstImported + 2}]`).dragTo(page.locator(`[data-testid=column-list-item-${firstImported}]`), { targetPosition: { x: 60, y: 3 } });
    await page.waitForTimeout(400);
    check((await titles())[firstImported] === movedTitle, `column list drag: 「${movedTitle}」 -> position ${firstImported + 1} (toast: ${await toast()})`);
  }

  // 2. move a base column across the merged title (A1:K1) in the grid -> 「結合を解除して移動」
  const iUni = await idxOf('大学');
  const iName = await idxOf('氏名');
  const drag = async (fromIdx, toIdx) => {
    await hdr(fromIdx).click();
    const from = await hdr(fromIdx).boundingBox();
    const to = await hdr(toIdx).boundingBox();
    await page.mouse.move(from.x + from.width / 2, from.y + from.height / 2);
    await page.mouse.down();
    await page.mouse.move(from.x + from.width / 2 - 20, from.y + from.height / 2, { steps: 4 });
    await page.mouse.move(to.x + 6, to.y + to.height / 2, { steps: 12 });
    await page.mouse.up();
    await page.waitForTimeout(400);
  };
  await page.locator('[data-testid=grid]').evaluate((el) => el.scrollTo(0, 0));
  await page.waitForTimeout(200);
  await drag(iUni, iName);
  const dlg = await page.waitForSelector('[data-testid=unmerge-confirm-dialog]', { timeout: 5000 }).catch(() => null);
  check(!!dlg, `moving 大学 in front of 氏名 asks 「結合を解除して移動」 (toast: ${await toast()})`);
  if (dlg) {
    console.log('unmerge dialog:', await text('[data-testid=unmerge-confirm-dialog]'));
    check((await page.evaluate(() => document.activeElement?.getAttribute('data-testid'))) === 'confirm-ok', 'unmerge dialog: initial focus on 「結合を解除して移動」');
    await shot('real-04-unmerge-confirm');
    const n0 = await editsCount();
    await page.click('[data-testid=confirm-ok]');
    await page.waitForTimeout(500);
    check((await hdrTitle(iName)) === '大学', `大学 moved to ${letter(iName)} after unmerging the title (got ${await hdrTitle(iName)})`);
    check((await editsCount()) === n0 + 1, 'unmerge + move = one undo step');
    console.log('toast:', await toast());
    await shot('real-05-unmerged-moved');
  } else {
    await shot('real-04-unmerge-refused');
    // the doc refused without the fix: still move a column in front of the title block (no merge is split)
    const iJudge = await idxOf('判定');
    if (iJudge >= 0) {
      await page.locator(`[data-testid=column-list-item-${iJudge}]`).dragTo(page.locator('[data-testid=column-list-item-0]'), { targetPosition: { x: 60, y: 2 } });
      await page.waitForTimeout(500);
      check((await titles())[0] === '判定', `fallback move: 判定 -> A (toast: ${await toast()})`);
      await shot('real-05-moved-to-A');
    }
  }

  // 3. insert a column right of 志望度 and type a date / % / =SUM / =UNIQUE
  const iAfter = await idxOf(insertAfter);
  await reveal(iAfter);
  check(await menu(iAfter, 'insert-right'), `context menu 右に列を挿入 (after ${insertAfter})`);
  await page.waitForTimeout(400);
  const iNew = iAfter + 1;
  console.log('inserted:', await hdrTitle(iNew), 'toast:', await toast());
  const iUni2 = await idxOf('大学');
  // 志望度 AFTER the insert (inserting left of it moves it: the SUM must not include its own cell)
  const iWish = await idxOf('志望度');
  const W = letter(iWish);
  const U = letter(iUni2);
  const edits = [
    [4, '2026/4/1'],
    [5, '50%'],
    [6, `=SUM(${W}4:${W}63)`],
    [8, `=UNIQUE(${U}4:${U}63)`],
  ];
  for (const [row, input] of edits) {
    await page.locator(`[data-testid=grid-cell-${row}-${iNew}]`).click();
    await page.fill('[data-testid=formula-bar-input]', input);
    await page.press('[data-testid=formula-bar-input]', 'Enter');
    await page.waitForTimeout(300);
    const shown = (await page.locator(`[data-testid=grid-cell-${row}-${iNew}]`).innerText()).trim();
    await page.locator(`[data-testid=grid-cell-${row}-${iNew}]`).click();
    const bar = await page.locator('[data-testid=formula-bar-input]').inputValue();
    console.log(`edit ${letter(iNew)}${row} <- ${input}: cell shows "${shown}", bar "${bar}", toast: ${await toast()}`);
  }
  // A new column inherits the number format of its left neighbour (Excel 「書式を左と同じにする」): only a General
  // neighbour (the default 志望度) gets the date / % format from the typed text, like Excel.
  const shownDate = (await page.locator(`[data-testid=grid-cell-4-${iNew}]`).innerText()).trim();
  const shownPct = (await page.locator(`[data-testid=grid-cell-5-${iNew}]`).innerText()).trim();
  if (insertAfter === '志望度') {
    check(shownDate === '2026/4/1', 'date typed into a General cell shows 2026/4/1');
    check(shownPct === '50%', '50% shows as 50%');
  } else {
    console.log(`info: neighbour ${insertAfter} has its own number format; typed date shows "${shownDate}", 50% shows "${shownPct}"`);
  }
  await shot('real-06-edited');

  // 4. delete an imported column that formulas depend on (面接評価 平均 <- 判定) -> convert to values
  const iAvg = await idxOf('平均');
  if (iAvg >= 0) {
    await reveal(iAvg);
    await menu(iAvg, 'delete');
    const del = await page.waitForSelector('[data-testid=delete-confirm-dialog]', { timeout: 5000 }).catch(() => null);
    check(!!del, 'deleting 平均 (referenced by 判定) asks first');
    if (del) {
      console.log('delete dialog:', await text('[data-testid=delete-confirm-dialog]'));
      await shot('real-07-delete-confirm');
      await page.click('[data-testid=delete-convert-button]');
      await page.waitForTimeout(500);
      check((await idxOf('平均')) === -1, `平均 deleted (toast: ${await toast()})`);
    }
  } else console.log('no 平均 column (imports differ?)');
  console.log('columns now:', (await titles()).join(' | '));
  console.log('side summary:', await text('[data-testid=side-summary]'));

  // ---------------------------------------------------------------------------------------------- export
  const exportWith = async (engine, name, shots) => {
    await page.click('[data-testid=export-button]');
    await page.waitForSelector('[data-testid=export-dialog]');
    const eng = page.locator(`[data-testid=export-engine-${engine}]`);
    if (await eng.isDisabled()) {
      console.log(`engine ${engine} disabled`);
      await page.keyboard.press('Escape');
      return null;
    }
    await eng.check();
    await page.fill('[data-testid=export-filename]', name);
    if (shots) await shot(shots[0]);
    const t = Date.now();
    await page.click('[data-testid=export-confirm]');
    // trace the progress texts of the dialog (where does the time go?)
  let lastProgress = '';
  const tracer = setInterval(async () => {
    const pt = await text('[data-testid=export-progress]').catch(() => null);
    if (pt && pt !== lastProgress) {
      lastProgress = pt;
      console.log(`  [${engine} ${Date.now() - t} ms] ${pt}`);
    }
  }, 700);
  await page.waitForSelector('[data-testid=export-result], [data-testid=export-error]', { timeout: 600000 });
  clearInterval(tracer);
    await page.waitForTimeout(200);
    const res = await text('[data-testid=export-result]');
    console.log(`export ${engine} in ${Date.now() - t} ms:`, res, (await text('[data-testid=export-error]')) ?? '');
    console.log(`  warnings (${engine}):`, (await text('[data-testid=export-warnings]')) ?? '（なし）');
    const p = (await page.locator('[data-testid=export-result] .result-path').count()) ? (await page.locator('[data-testid=export-result] .result-path').innerText()).trim() : null;
    if (shots) await shot(shots[1]);
    const usedEngine = await text('[data-testid=export-result-engine]');
    check(!!p, `export ${engine}: saved (${usedEngine})`);
    if (engine === 'excel') check(usedEngine === 'Excel本体で出力しました', 'Excel本体 engine really used (no fallback)');
    await page.keyboard.press('Escape');
    await page.waitForTimeout(300);
    if (p) console.log(`OUT ${engine} ${p}`);
    return p;
  };
  if (excelWanted) await exportWith('excel', `${prefix}_excel.xlsx`, ['real-08-export-dialog', 'real-09-export-result']);
  await exportWith('js', `${prefix}_js.xlsx`, excelWanted ? null : ['real-08-export-dialog', 'real-09-export-result']);
  const det = page.locator('[data-testid=warning-details]');
  if (await det.count()) {
    await det.evaluate((el) => (el.open = true));
    console.log('警告:', (await page.locator('[data-testid=warning-list]').allInnerTexts()).join(' / ').replace(/\s+/g, ' '));
  }
}

console.log(errors.length ? 'console errors:\n' + errors.join('\n') : 'no console errors');
console.log(failures.length ? `${failures.length} failed check(s):\n  ${failures.join('\n  ')}` : 'all checks passed');
await browser.close();
if (errors.length || failures.length) process.exitCode = 1;
