// Round 4 flow of the REAL app (no mock) through the real local server: a file WITHOUT a usable key is excluded from
// the merge instead of blocking it, and the manual key picker of the detail modal.
//   PORT=5192 NO_OPEN=1 node server/server.mjs          (another terminal; DIST_DIR / OUTPUT_DIR may point elsewhere)
//   node tests/ui/real-excluded.mjs [baseUrl] [--js-only] [--no-shots] [--prefix=<name>]
// Files: 01_学生名簿 + 02_面接評価 + 04_適性検査, with 02 as the base (04 has only a kana column: no key against 02).
// Checks: the card badge + 「キーを選ぶ…」, the summary / preview notice / 警告 note, the merge button, the key detail modal
// (opened on 04's tab, focus trap, Esc), the manual picker (1 and 2 columns per side, live hint, 追加, × removal, the file
// joins the merge again), the match-rate sentence under the key table (live), no key at all -> 「結合する」 disabled with a
// tooltip, the merge without 04, both exports (prints `OUT excel <path>` / `OUT js <path>`), and the merge again with a
// key selected for 04. Screenshots: tests/ui/shots/real-1*.png. Exit 1 on a failed check or a console error.
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
const f01 = join(samples, '01_学生名簿.xlsx');
const f02 = join(samples, '02_面接評価.xlsx');
const f04 = join(samples, '04_適性検査.xls');
const excelWanted = !flags.has('--js-only');
const prefix = [...flags].find((f) => f.startsWith('--prefix='))?.slice('--prefix='.length) || 'u4x';
const noShots = flags.has('--no-shots');

const browser = await chromium.launch({ channel: 'msedge', headless: true });
const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 }, acceptDownloads: true, locale: 'ja-JP' });
const page = await ctx.newPage();
const errors = [];
const failures = [];
page.on('console', (m) => m.type() === 'error' && errors.push(m.text()));
page.on('pageerror', (e) => errors.push('pageerror: ' + e.message));
const text = async (sel) => ((await page.locator(sel).count()) ? (await page.locator(sel).first().innerText()).replace(/\s+/g, ' ').trim() : null);
const has = async (sel) => (await page.locator(sel).count()) > 0;
const idle = async () => {
  await page.waitForFunction(() => !document.querySelector('[data-testid=busy-overlay]'), null, { timeout: 180000 });
  await page.waitForTimeout(250);
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
const active = () => page.evaluate(() => document.activeElement?.getAttribute('data-testid') ?? document.activeElement?.tagName ?? null);
const warnings = async () => {
  const det = page.locator('[data-testid=warning-details]');
  if (await det.count()) await det.evaluate((el) => (el.open = true));
  return (await page.locator('[data-testid=warning-list] li').allInnerTexts()).map((s) => s.replace(/\s+/g, ' ').trim());
};
const cardCount = () => page.locator('[data-testid^=file-card-]').count();
const waitLoaded = async (n) => {
  await page.waitForFunction((k) => document.querySelectorAll('[data-testid^=file-card-]').length >= k, n, { timeout: 60000 });
  await page.waitForFunction(() => !document.querySelector('.card-loading'), null, { timeout: 180000 });
  await page.waitForTimeout(300);
  await idle();
  await idle();
};

// ---------------------------------------------------------------------------------------------- files
await page.goto(base);
await page.waitForSelector('[data-testid=server-status]');
await page.waitForTimeout(400);
await page.setInputFiles('[data-testid=file-input]', [f01, f02, f04]);
await waitLoaded(3);
console.log('summary (base 01):', await text('[data-testid=summary-line]'));
await page.click('[data-testid=base-toggle-1]'); // 02_面接評価 becomes the base
await idle();
await page.waitForTimeout(400);
await idle();

const summaryLine = (await text('[data-testid=summary-line]')) ?? '';
console.log('summary (base 02):', summaryLine);
check(/基準 63 名/.test(summaryLine), `02 is the base: 基準 63 名 (${summaryLine})`);
check(/結合対象外 1 ファイル/.test(summaryLine), 'the summary line says 結合対象外 1 ファイル');
check((await has('[data-testid=file-excluded-2]')) && !(await has('[data-testid=file-excluded-0]')) && !(await has('[data-testid=file-excluded-1]')), 'only the 04 card has the excluded badge');
const badge = (await text('[data-testid=file-excluded-2]')) ?? '';
console.log('04 card:', badge);
check(badge.includes('結合対象外（結合キーが見つかりません）') && badge.includes('キーを選ぶ…'), '04 card: 「結合対象外（結合キーが見つかりません）」 + 「キーを選ぶ…」');
check(!(await page.locator('[data-testid=merge-button]').isDisabled()), 'the merge button stays enabled (04 does not block the merge)');
const notice = (await text('[data-testid=excluded-notice]')) ?? '';
check(notice.includes('04_適性検査') && notice.includes('結合対象外'), `the preview says it (${notice.slice(0, 60)}…)`);
check(/結合対象外/.test((await text('[data-testid=side-summary]')) ?? ''), 'the side summary says it');
const notes0 = await warnings();
console.log('警告:', notes0.join(' / '));
const excludedNote = (list) => list.find((n) => n.includes('04_適性検査') && /結合の対象外/.test(n));
check(!!excludedNote(notes0), `警告 has the note (${excludedNote(notes0)?.slice(0, 50)}…)`);
await shot('real-10-files-excluded');

// ---------------------------------------------------------------------------------------------- detail modal on 04's tab
await page.click('[data-testid=file-excluded-2] button');
await page.waitForSelector('[data-testid=key-detail-modal]');
check((await page.locator('[data-testid=key-detail-tab-s3]').getAttribute('aria-selected')) === 'true', '「キーを選ぶ…」 opens the detail modal on 04\'s tab');
check(await has('[data-testid=detail-excluded]'), 'the modal explains that the file is 結合対象外');
check((await text('[data-testid=key-detail-tab-s3]'))?.includes('結合対象外') ?? false, 'the tab carries the 結合対象外 tag');
check((await active()) === 'key-detail-ok', 'initial focus on OK');
check((await has('[data-testid=match-rate-empty]')) && !(await has('[data-testid=match-rate-text]')), 'no key selected: no match-rate sentence');
await shot('real-11-key-detail-excluded');

// focus trap: Tab / Shift+Tab never leave the dialog
let trapped = true;
for (let i = 0; i < 70 && trapped; i++) {
  await page.keyboard.press('Tab');
  trapped = await page.evaluate(() => !!document.activeElement?.closest('[data-testid=key-detail-modal]'));
}
for (let i = 0; i < 70 && trapped; i++) {
  await page.keyboard.press('Shift+Tab');
  trapped = await page.evaluate(() => !!document.activeElement?.closest('[data-testid=key-detail-modal]'));
}
check(trapped, 'focus trap: 70 x Tab and 70 x Shift+Tab stay inside the dialog');

// ---------------------------------------------------------------------------------------------- manual picker
const picker = page.locator('[data-testid=manual-key-picker]');
await picker.scrollIntoViewIfNeeded();
const baseOpts = await page.locator('[data-testid=manual-base-select-0] option').allInnerTexts();
const otherOpts = await page.locator('[data-testid=manual-other-select-0] option').allInnerTexts();
console.log('base options:', baseOpts.join(' | '));
console.log('04 options:', otherOpts.join(' | '));
check(baseOpts.includes('受験番号（A列）') && baseOpts.includes('氏名（B列）'), 'base dropdown lists every column of 02 by header + column letter');
check(otherOpts.includes('学籍番号（A列）') && otherOpts.includes('氏名カナ（B列）') && otherOpts.length === 9, `04 dropdown lists its 8 columns (+ the empty choice): ${otherOpts.length}`);
check((await page.locator('[data-testid=manual-key-add]').isDisabled()) && /選ぶと/.test((await text('[data-testid=manual-key-hint]')) ?? ''), '「追加」 is disabled until both sides are chosen');
await page.selectOption('[data-testid=manual-base-select-0]', { label: '受験番号（A列）' });
check((await page.locator('[data-testid=manual-key-add]').isDisabled()), 'one side only: still disabled');
await page.selectOption('[data-testid=manual-other-select-0]', { label: '学籍番号（A列）' });
const hint1 = (await text('[data-testid=manual-key-hint]')) ?? '';
console.log('hint:', hint1);
check(hint1.includes('検出した種類') && hint1.includes('受験番号') && hint1.includes('学籍番号'), 'live hint shows the detected type of the chosen columns');
check(!(await page.locator('[data-testid=manual-key-add]').isDisabled()), '「追加」 enabled with a column on each side');
// 2 columns per side: one more dropdown each; unequal counts are explained, × takes the extra one away again
await page.click('[data-testid=manual-base-add-column]');
check((await page.locator('[data-testid^=manual-base-select-]').count()) === 2, '「＋列を追加」 adds a second dropdown on the base side');
check((await page.locator('[data-testid=manual-key-add]').isDisabled()), 'an empty second dropdown blocks 「追加」');
await page.selectOption('[data-testid=manual-base-select-1]', { label: '氏名（B列）' });
check(/同じ数の列/.test((await text('[data-testid=manual-key-hint]')) ?? ''), 'unequal number of columns is explained');
await page.click('[data-testid=manual-other-add-column]');
await page.selectOption('[data-testid=manual-other-select-1]', { label: '氏名カナ（B列）' });
check(!(await page.locator('[data-testid=manual-key-add]').isDisabled()), '2 + 2 columns can be added');
await shot('real-12-manual-picker');
// the smallest supported window: the dialog scrolls inside, nothing is cut off sideways, OK stays reachable
await page.setViewportSize({ width: 1280, height: 720 });
await page.waitForTimeout(300);
await picker.scrollIntoViewIfNeeded();
const fit = await page.evaluate(() => {
  const body = document.querySelector('[data-testid=key-detail-modal] .modal-body');
  const dlg = document.querySelector('[data-testid=key-detail-modal]').getBoundingClientRect();
  const ok = document.querySelector('[data-testid=key-detail-ok]').getBoundingClientRect();
  const sides = [...document.querySelectorAll('[data-testid=manual-key-picker] .picker-side')].map((el) => el.getBoundingClientRect());
  return {
    noSideScroll: body.scrollWidth <= body.clientWidth + 1,
    okVisible: ok.bottom <= innerHeight && ok.top >= 0,
    dialogFits: dlg.top >= 0 && dlg.bottom <= innerHeight && dlg.left >= 0 && dlg.right <= innerWidth,
    sidesInside: sides.every((r) => r.left >= dlg.left && r.right <= dlg.right),
  };
});
check(fit.noSideScroll && fit.okVisible && fit.dialogFits && fit.sidesInside, `manual picker at 1280x720: ${JSON.stringify(fit)}`);
await shot('real-12b-manual-picker-1280x720');
await page.setViewportSize({ width: 1440, height: 900 });
await page.waitForTimeout(300);
await page.click('[data-testid=manual-key-add]');
await idle();
check((await page.locator('[data-testid^=key-row-manual]').count()) === 1, 'a composite manual key row was added');
const compRow = page.locator('[data-testid^=key-row-manual]').first();
const compText = (await compRow.innerText()).replace(/\s+/g, ' ');
console.log('composite row:', compText);
check(compText.includes('手動') && (await compRow.locator('input[type=checkbox]').isChecked()), 'the new candidate is marked 手動 and checked');
const compId = await compRow.getAttribute('data-testid');
// remove it with × (analysis runs again)
await page.click(`[data-testid="key-remove-${compId.replace('key-row-', '')}"]`);
await idle();
check((await page.locator('[data-testid^=key-row-manual]').count()) === 0, '× removes the manual candidate');
check(await has('[data-testid=detail-excluded]'), '… and the file is 結合対象外 again');
// the single-column key
await picker.scrollIntoViewIfNeeded();
check((await page.locator('[data-testid^=manual-base-select-]').count()) === 1, 'the picker starts again with one dropdown per side');
await page.selectOption('[data-testid=manual-base-select-0]', { label: '受験番号（A列）' });
await page.selectOption('[data-testid=manual-other-select-0]', { label: '学籍番号（A列）' });
await page.click('[data-testid=manual-key-add]');
await idle();
const rowSel = page.locator('[data-testid^=key-row-manual]');
check((await rowSel.count()) === 1, 'the single-column key was added to the table');
const rowText = ((await rowSel.first().innerText()) ?? '').replace(/\s+/g, ' ');
console.log('row:', rowText);
check((await rowSel.first().locator('input[type=checkbox]').isChecked()) && rowText.includes('手動'), 'checked + 手動');
check((await rowSel.first().locator('[data-testid^=key-remove-]').count()) === 1, 'the manual candidate has a × button');
const autoRemove = await page.locator('[data-testid=key-candidates] tr:not([data-testid^=key-row-manual]) [data-testid^=key-remove-]').count();
check(autoRemove === 0, 'automatic candidates have no × (they can only be unchecked)');
const rate04 = (await text('[data-testid=match-rate-text]')) ?? '';
console.log('match rate (04):', rate04);
check(/^このキーで \d+ \/ 42 行が一致します（基準 63 行中 \d+ 行が見つかりました）$/.test(rate04), 'the match-rate sentence under the key table (Power Query style)');
check(!(await has('[data-testid=detail-excluded]')) && !(((await text('[data-testid=key-detail-tab-s3]')) ?? '').includes('結合対象外')), 'the file is part of the merge again (banner and tab tag gone)');
check(/追加しました/.test((await text('[data-testid=manual-key-message]')) ?? ''), 'the picker confirms the addition');
await picker.scrollIntoViewIfNeeded();
await shot('real-13-manual-added');
// live: uncheck -> excluded again, sentence gone; check -> back, sentence again
const cb = rowSel.first().locator('input[type=checkbox]');
await cb.uncheck();
await idle();
check((await has('[data-testid=detail-excluded]')) && (await has('[data-testid=match-rate-empty]')), 'unchecking the only key: 結合対象外 again, the sentence is replaced by a hint');
await cb.check();
await idle();
check(!(await has('[data-testid=detail-excluded]')) && (await has('[data-testid=match-rate-text]')), 'checking it again: part of the merge again');

// the sentence follows the toggles of another file (01): a used key -> sentence; none -> hint
await page.click('[data-testid=key-detail-tab-s1]');
await page.waitForTimeout(200);
const rate01 = (await text('[data-testid=match-rate-text]')) ?? '';
console.log('match rate (01):', rate01);
// 63 base rows: 60 are found in 01; two base rows share an ID, so only 58 of the 60 rows of 01 are found in the base
check(/^このキーで 58 \/ 60 行が一致します（基準 63 行中 60 行が見つかりました）$/.test(rate01), `01: 58 / 60 rows match, 60 of the 63 base rows found (${rate01})`);
const boxes = page.locator('[data-testid=key-candidates] tbody tr.is-used input[type=checkbox]');
const nUsed = await boxes.count();
for (let i = 0; i < nUsed; i++) {
  await page.locator('[data-testid=key-candidates] tbody tr.is-used input[type=checkbox]').first().uncheck();
  await idle();
}
check((await has('[data-testid=match-rate-empty]')) && (await has('[data-testid=detail-excluded]')), '01: no key left -> the sentence disappears, 結合対象外');
await page.locator('[data-testid=key-candidates] tbody tr:not(.is-used) input[type=checkbox]:not([disabled])').first().check();
await idle();
const rateBack = (await text('[data-testid=match-rate-text]')) ?? '';
check(/^このキーで \d+ \/ 60 行が一致します/.test(rateBack), `01: a key checked again -> the sentence is back (${rateBack})`);
await page.click('[data-testid=key-detail-tab-s3]');
await page.keyboard.press('Escape');
await page.waitForTimeout(300);
check(!(await has('[data-testid=key-detail-modal]')), 'Esc closes the detail modal');
check((await active()) !== 'BODY', `focus returns to the page after Esc (${await active()})`);

// ---------------------------------------------------------------------------------------------- remove the manual key: excluded again
await page.click('[data-testid=key-detail-button]');
await page.waitForSelector('[data-testid=key-detail-modal]');
await page.click('[data-testid=key-detail-tab-s3]');
await page.waitForTimeout(200);
const manualRow = page.locator('[data-testid^=key-row-manual]').first();
const manualId = (await manualRow.getAttribute('data-testid')).replace('key-row-', '');
check(((await manualRow.count()) === 1) && (await manualRow.locator('input[type=checkbox]').isChecked()), 'the manual key survives closing the dialog');
await page.click(`[data-testid="key-remove-${manualId}"]`);
await idle();
await page.click('[data-testid=key-detail-ok]');
await page.waitForTimeout(300);
check(await has('[data-testid=file-excluded-2]'), '× + OK: the 04 card is excluded again');

// ---------------------------------------------------------------------------------------------- no key at all
await page.click('[data-testid=file-clear-0]'); // 01_学生名簿 removed: only 02 (base) and 04 (no key) remain
await idle();
await page.waitForTimeout(400);
await idle();
const mergeBtn = page.locator('[data-testid=merge-button]');
const disabledTitle = (await mergeBtn.getAttribute('title')) ?? '';
console.log('no key at all -> merge button disabled:', await mergeBtn.isDisabled(), '|', disabledTitle);
check((await mergeBtn.isDisabled()) && disabledTitle.includes('どのファイルにも結合キーが選ばれていません'), '「結合する」 is disabled with a tooltip saying why');
check(/どのファイルにも結合キーが選ばれていません/.test((await text('[data-testid=side-summary]')) ?? ''), 'the side summary says so too');
await shot('real-14-no-key-at-all');
// 01 comes back (a new card): the merge is possible again without any other action
await page.setInputFiles('[data-testid=file-input]', [f01]);
await waitLoaded(3);
check(!(await mergeBtn.isDisabled()), 'adding a file that has a key enables 「結合する」 again');
console.log('cards:', await cardCount(), '| summary:', await text('[data-testid=summary-line]'));
// the cards are now 02 (0), 04 (1), 01 (2); the file ids: s2 (02) / s3 (04) / s4 (01)
check(await has('[data-testid=file-excluded-1]'), 'the 04 card (now the second one) is still excluded');

// ---------------------------------------------------------------------------------------------- merge without 04
await mergeBtn.click();
await page.waitForSelector('[data-testid=grid]', { timeout: 180000 });
await idle();
const mergeToast = (await toast()) ?? '';
console.log('merge toast:', mergeToast);
check(mergeToast.includes('結合しました') && mergeToast.includes('結合対象外 1 ファイル（04_適性検査）'), 'the merge goes ahead; the toast names the excluded file');
const titles = async () => (await page.locator('[data-testid=column-list] li .col-title').allInnerTexts()).map((s) => s.trim());
const cols = await titles();
console.log('columns:', cols.join(' | '));
check(cols.slice(0, 9).join('|') === '受験番号|氏名|面接官|評価1|評価2|評価3|平均|判定|コメント', 'the base columns of 02 come first');
check(cols.includes('フリガナ') && cols.includes('志望度'), 'the columns of 01 were imported');
check(!cols.includes('性格タイプ') && !cols.includes('総合判定') && !cols.includes('学籍番号'), 'no column of 04 was imported');
const notesM = await warnings();
check(!!excludedNote(notesM) && excludedNote(notesM).includes('取り込んでいません'), `警告 (after the merge): ${excludedNote(notesM)?.slice(0, 60)}…`);
console.log('side summary:', await text('[data-testid=side-summary]'));
await shot('real-15-merged-excluded');

// ---------------------------------------------------------------------------------------------- export (both engines)
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
  if (shots) await shot(shots);
  const used = await text('[data-testid=export-result-engine]');
  check(!!p, `export ${engine}: saved (${used})`);
  if (engine === 'excel') check(used === 'Excel本体で出力しました', 'Excel本体 engine really used (no fallback)');
  await page.keyboard.press('Escape');
  await page.waitForTimeout(300);
  if (p) console.log(`OUT ${engine} ${p}`);
  return p;
};
if (excelWanted) await exportWith('excel', `${prefix}_excel.xlsx`, null);
await exportWith('js', `${prefix}_js.xlsx`, null);

// ---------------------------------------------------------------------------------------------- back to the keys, select a key for 04, merge again
await page.click('[data-testid=remerge-button]');
await page.waitForSelector('[data-testid=merge-button]', { timeout: 10000 });
await idle();
check(await has('[data-testid=file-excluded-1]'), 'after 「結合をやり直す」 the 04 card is still excluded');
await page.click('[data-testid=file-excluded-1] button');
await page.waitForSelector('[data-testid=key-detail-modal]');
await picker.scrollIntoViewIfNeeded();
await page.selectOption('[data-testid=manual-base-select-0]', { label: '受験番号（A列）' });
await page.selectOption('[data-testid=manual-other-select-0]', { label: '学籍番号（A列）' });
await page.click('[data-testid=manual-key-add]');
await idle();
await page.click('[data-testid=key-detail-ok]');
await page.waitForTimeout(300);
check(!(await has('[data-testid=file-excluded-1]')), 'a key selected in the detail modal: the card is no longer excluded');
check(!/結合対象外/.test((await text('[data-testid=summary-line]')) ?? ''), 'the summary no longer mentions 結合対象外');
check((await text('[data-testid=file-keys-1]'))?.includes('学籍番号') ?? false, `the 04 card shows its key line (${await text('[data-testid=file-keys-1]')})`);
await mergeBtn.click();
await page.waitForSelector('[data-testid=grid]', { timeout: 180000 });
await idle();
const toast2 = (await toast()) ?? '';
console.log('merge toast (04 joined):', toast2);
const cols2 = await titles();
console.log('columns:', cols2.join(' | '));
check(cols2.includes('性格タイプ') && cols2.includes('総合判定'), 'the columns of 04 are imported once it has a key (without any other action)');
check(!toast2.includes('結合対象外'), 'the toast no longer mentions 結合対象外');
await shot('real-16-merged-04-joined');

console.log(errors.length ? 'console errors:\n' + errors.join('\n') : 'no console errors');
console.log(failures.length ? `${failures.length} failed check(s):\n  ${failures.join('\n  ')}` : 'all checks passed');
await browser.close();
if (errors.length || failures.length) process.exitCode = 1;
