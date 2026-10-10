/**
 * E2E helpers (owner: T). Everything goes through the data-testid contract of docs/ARCHITECTURE.md §5.
 *
 * The UI and the core modules are built concurrently. Instead of hard-coding test.fixme, specs call
 *   requireTestIds(page, [...])   -> test.fixme at runtime when an element of the contract is missing
 *   fixmeIfNotImplemented(page, e) -> test.fixme when a core module is still a stub ("... not implemented yet")
 * so they start running for real as soon as the missing piece lands.
 */
import { spawn, spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { expect, test, type Locator, type Page } from '@playwright/test';
import { excelAvailable, inspectFiles, type CellProbe, type InspectResult } from '../tools/lib/inspect.mjs';
import { checkOutput as checkOutputDumps } from '../tools/lib/verdict.mjs';

export const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const SAMPLES_DIR = path.join(ROOT, 'samples');
export const SAMPLE_FILES = ['01_学生名簿.xlsx', '02_面接評価.xlsx', '03_連絡先.csv', '04_適性検査.xls', '05_研修出欠.xlsx'];
export const BASE_SAMPLE = path.join(SAMPLES_DIR, SAMPLE_FILES[0]);

export const tid = (page: Page, id: string): Locator => page.getByTestId(id);

export function samplePaths(names: string[] = SAMPLE_FILES): string[] {
  const paths = names.map((n) => path.join(SAMPLES_DIR, n));
  const missing = paths.filter((p) => !fs.existsSync(p));
  test.skip(missing.length > 0, `サンプルがありません（npm run samples を実行してください）: ${missing.map((p) => path.basename(p)).join(', ')}`);
  return paths;
}

/** Mark the running test fixme when one of the contract's data-testids is not rendered (yet). */
export async function requireTestIds(page: Page, ids: string[], what: string, timeoutMs = 3000): Promise<void> {
  for (const id of ids) {
    const loc = page.locator(`[data-testid="${id}"]`).first();
    try {
      await loc.waitFor({ state: 'attached', timeout: timeoutMs });
    } catch {
      test.fixme(true, `${what}: UI に data-testid="${id}" がまだありません`);
    }
  }
}

/** Collects page errors / console errors so a stub module ("not implemented yet") can be recognised. */
export function watchErrors(page: Page): { messages: string[] } {
  const box = { messages: [] as string[] };
  page.on('pageerror', (e) => box.messages.push(String(e?.message ?? e)));
  page.on('console', (m) => {
    if (m.type() === 'error') box.messages.push(m.text());
  });
  return box;
}

const NOT_IMPLEMENTED = /not implemented yet|未実装/i;

/** Every toast text shown so far (recorded by the init script of openApp; toasts disappear after a few s). */
export async function toastLog(page: Page): Promise<string[]> {
  return page.evaluate(() => (window as unknown as { __e2eToasts?: string[] }).__e2eToasts ?? []).catch(() => []);
}

/** Text of the visible toast / error fallback / export error, if any (+ every toast shown so far). */
export async function visibleErrorText(page: Page): Promise<string> {
  const parts: string[] = [...(await toastLog(page))];
  for (const id of ['toast', 'error-fallback', 'export-error']) {
    const loc = page.getByTestId(id);
    if (await loc.count()) {
      const t = (await loc.first().innerText().catch(() => '')).trim();
      if (t) parts.push(t);
    }
  }
  return parts.join(' / ');
}

/** When `err` (or what the page shows) says a module is still a stub, mark the test fixme; otherwise rethrow. */
export async function fixmeIfNotImplemented(page: Page, errors: { messages: string[] }, what: string, err?: unknown): Promise<void> {
  const shown = await visibleErrorText(page);
  const all = [String(err instanceof Error ? err.message : err ?? ''), shown, ...errors.messages].join('\n');
  const m = all.split('\n').find((l) => NOT_IMPLEMENTED.test(l));
  if (m) test.fixme(true, `${what}: モジュールが未実装です — ${m.trim().slice(0, 200)}`);
  if (err) throw err;
}

/** Runs `fn`; a failure caused by a stub module turns the test into fixme. */
export async function step<T>(page: Page, errors: { messages: string[] }, what: string, fn: () => Promise<T>): Promise<T> {
  try {
    return await fn();
  } catch (e) {
    const shown = await visibleErrorText(page).catch(() => '');
    if (shown) {
      test.info().annotations.push({ type: `${what}: 画面のメッセージ`, description: shown });
      if (e instanceof Error) e.message += `\n画面のメッセージ（トースト等）: ${shown}`;
    }
    await fixmeIfNotImplemented(page, errors, what, e);
    throw e;
  }
}

/** Opens the app. `standalone` pages get no File System Access picker, so exports fall back to a download. */
export async function openApp(page: Page, query = ''): Promise<void> {
  await page.addInitScript(() => {
    try {
      delete (window as unknown as { showSaveFilePicker?: unknown }).showSaveFilePicker;
    } catch {
      /* ignore */
    }
    // remember every toast (they vanish after a few seconds) so a failure can be explained afterwards
    const w = window as unknown as { __e2eToasts: string[] };
    w.__e2eToasts = [];
    new MutationObserver(() => {
      const t = document.querySelector('[data-testid="toast"]');
      const s = t?.textContent?.trim();
      if (s && w.__e2eToasts[w.__e2eToasts.length - 1] !== s) w.__e2eToasts.push(s);
    }).observe(document, { subtree: true, childList: true, characterData: true });
  });
  await page.goto('/' + query);
  await expect(page.locator('#root')).not.toBeEmpty();
}

export async function isServerMode(page: Page): Promise<boolean> {
  const r = await page.request.get('/api/status').catch(() => null);
  if (!r || !r.ok()) return false;
  const j = (await r.json().catch(() => null)) as { app?: string } | null;
  return j?.app === 'excel-merger';
}

export async function excelInstalled(page: Page): Promise<boolean> {
  const r = await page.request.get('/api/status').catch(() => null);
  if (!r || !r.ok()) return false;
  const j = (await r.json().catch(() => null)) as { excel?: { installed?: boolean } } | null;
  return !!j?.excel?.installed;
}

export async function loadFiles(page: Page, files: string[]): Promise<void> {
  await requireTestIds(page, ['file-input'], 'ファイル読み込み');
  await page.getByTestId('file-input').setInputFiles(files);
  for (let i = 0; i < files.length; i++) await expect(page.getByTestId(`file-card-${i}`)).toBeVisible();
}

/** Waits until the key analysis is done and 「結合する」 can be pressed. */
export async function waitMergeReady(page: Page, errors: { messages: string[] }): Promise<void> {
  await requireTestIds(page, ['merge-button'], '結合ボタン');
  await step(page, errors, 'ファイルの読み込み・結合キーの検出', async () => {
    await expect(page.getByTestId('merge-button')).toBeEnabled({ timeout: 90_000 });
  });
}

/** the files on the cards at the time of the merge (the card row is gone afterwards): what check-output --with-copies needs as sources */
const cardMemory = new WeakMap<Page, { base: SourceFile | null; others: SourceFile[] }>();

export async function mergeNow(page: Page, errors: { messages: string[] }): Promise<void> {
  await waitMergeReady(page, errors);
  cardMemory.set(page, await cardSources(page).catch(() => ({ base: null, others: [] })));
  const before = (await toastLog(page)).length;
  await page.getByTestId('merge-button').click();
  await step(page, errors, '結合', async () => {
    // the grid appears, or the merge fails with an error toast (fail fast with its text)
    await expect
      .poll(
        async () => {
          if (await page.getByTestId('grid').isVisible().catch(() => false)) return 'grid';
          const failed = (await toastLog(page)).slice(before).find((t) => /できませんでした|エラー/.test(t));
          return failed ? `failed: ${failed}` : 'merging';
        },
        { message: '結合が 120 秒以内に終わりませんでした', timeout: 120_000, intervals: [100, 250, 500, 1000] },
      )
      .not.toBe('merging');
    if (!(await page.getByTestId('grid').isVisible().catch(() => false))) {
      const failed = (await toastLog(page)).slice(before).find((t) => /できませんでした|エラー/.test(t));
      throw new Error(`結合に失敗しました: ${failed ?? '（理由不明）'}`);
    }
    await expect(page.getByTestId('grid-header-0')).toBeVisible();
  });
}

export interface ColumnInfo {
  index: number;
  letter: string;
  title: string;
  source: string;
}

/** Current columns (in order) from the side panel's 列の一覧 (not virtualised, unlike the grid header). */
export async function columns(page: Page): Promise<ColumnInfo[]> {
  const items = page.locator('[data-testid^="column-list-item-"]');
  const rows = await items.evaluateAll((els) =>
    els.map((e) => ({
      id: e.getAttribute('data-testid') ?? '',
      letter: e.querySelector('.col-letter')?.textContent?.trim() ?? '',
      // the title span; fall back to the tooltip ("E列: 大学（学生回答（基準））— …")
      title: e.querySelector('.col-title')?.textContent?.trim() ?? null,
      tip: e.getAttribute('title') ?? '',
    })),
  );
  return rows
    .map((r) => {
      const title = r.title ?? /^[A-Z]+列: ([\s\S]*?)（/.exec(r.tip)?.[1] ?? r.tip;
      const source = r.tip.startsWith(`${r.letter}列: ${title}（`) ? r.tip.slice(`${r.letter}列: ${title}（`.length).replace(/）— [\s\S]*$/, '') : '';
      return { index: Number(r.id.replace('column-list-item-', '')), letter: r.letter, title, source };
    })
    .sort((a, b) => a.index - b.index);
}

export async function columnTitles(page: Page): Promise<string[]> {
  return (await columns(page)).map((c) => c.title);
}

export async function headerTitle(page: Page, i: number): Promise<string> {
  const h = page.getByTestId(`grid-header-${i}`);
  await h.scrollIntoViewIfNeeded();
  return (await h.locator('.g-htitle').innerText()).trim();
}

/** Brings grid column i into view (the grid virtualises columns). */
export async function revealColumn(page: Page, i: number): Promise<Locator> {
  const h = page.getByTestId(`grid-header-${i}`);
  if (!(await h.count())) {
    // scroll the grid horizontally until the header is rendered
    const grid = page.getByTestId('grid');
    for (let k = 0; k < 40 && !(await h.count()); k++) {
      await grid.evaluate((el) => el.scrollBy({ left: el.clientWidth * 0.6 }));
      await nextFrames(page, 2); // the virtualised columns render on the next frame
    }
  }
  await h.scrollIntoViewIfNeeded();
  return h;
}

/** Resolves after `n` animation frames (the grid reacts to pointer moves per frame; frames are a better clock than milliseconds under load). */
export async function nextFrames(page: Page, n = 2): Promise<void> {
  await page.evaluate(
    (count) =>
      new Promise<void>((resolve) => {
        const step = (left: number) => (left <= 0 ? resolve() : requestAnimationFrame(() => step(left - 1)));
        step(count);
      }),
    n,
  );
}

/** Waits until the grid's horizontal scroll position has not changed for `still` frames (auto-scroll finished; at most ~3 s). */
export async function settleGridScroll(page: Page, still = 4): Promise<void> {
  await page.evaluate(
    (need) =>
      new Promise<void>((resolve) => {
        const grid = document.querySelector('[data-testid="grid"]') as HTMLElement | null;
        const t0 = performance.now();
        let last = -1;
        let same = 0;
        const tick = () => {
          const x = grid?.scrollLeft ?? 0;
          same = x === last ? same + 1 : 0;
          last = x;
          if (same >= need || performance.now() - t0 > 3000) resolve();
          else requestAnimationFrame(tick);
        };
        requestAnimationFrame(tick);
      }),
    still,
  );
}

/**
 * Drags the selected header `fromIdx` so that the columns land before column `gap` (gap = n -> at the end).
 * Uses real mouse down / move / up (pointer events) like a user.
 */
export async function dragHeader(page: Page, fromIdx: number, gap: number): Promise<void> {
  const src = await revealColumn(page, fromIdx);
  const sb = (await src.boundingBox())!;
  const n = await page.locator('[data-testid^="column-list-item-"]').count();
  let x = sb.x + sb.width / 2;
  const y = sb.y + sb.height / 2;
  await page.mouse.move(x, y);
  await page.mouse.down();
  // start the drag (the grid starts it after 4 px and tracks pointermove)
  for (let s = 1; s <= 4; s++) {
    await page.mouse.move(x + s * 3, y);
    await nextFrames(page, 1);
  }
  x += 12;
  // the target may be scrolled out of view (columns are wide, the grid virtualises them): hover in the grid's
  // auto-scroll zone near the left / right edge until the drop point is well inside the visible area
  const deadline = Date.now() + 20_000;
  for (;;) {
    const where = await dropPoint(page, gap, n);
    if (typeof where !== 'string') {
      const steps = 8;
      for (let s = 1; s <= steps; s++) {
        await page.mouse.move(x + ((where.x - x) * s) / steps, y + ((where.y - y) * s) / steps);
        await nextFrames(page, 1);
      }
      await settleGridScroll(page);
      // still the right spot after the last frame of auto-scroll?
      const again = await dropPoint(page, gap, n);
      if (typeof again !== 'string' && Math.abs(again.x - where.x) > 2) {
        x = where.x;
        continue;
      }
      break;
    }
    if (Date.now() > deadline) break;
    const edge = await page.evaluate((side) => {
      const grid = document.querySelector('[data-testid="grid"]') as HTMLElement;
      const r = grid.getBoundingClientRect();
      const pin = grid.querySelector('.g-pin');
      const left = pin ? pin.getBoundingClientRect().right : r.left + 56;
      return side === 'left' ? left + 16 : r.left + grid.clientWidth - 12;
    }, where);
    await page.mouse.move(edge, y);
    x = edge;
    await nextFrames(page, 4); // a few frames of auto-scroll
  }
  await page.mouse.up();
}

/**
 * Viewport point that drops before column `gap` (n = after the last column), or the direction to auto-scroll
 * when that point is hidden (scrolled away / behind the frozen pane) or inside the auto-scroll zone.
 */
async function dropPoint(page: Page, gap: number, n: number): Promise<{ x: number; y: number } | 'left' | 'right'> {
  return page.evaluate(
    ({ gap, n }) => {
      const EDGE = 60;
      const grid = document.querySelector('[data-testid="grid"]') as HTMLElement;
      const r = grid.getBoundingClientRect();
      const pin = grid.querySelector('.g-pin');
      const leftEdge = pin ? pin.getBoundingClientRect().right : r.left + 56;
      const rightEdge = r.left + grid.clientWidth;
      const hdr = (i: number) => grid.querySelector<HTMLElement>(`[data-testid="grid-header-${i}"]`);
      const frozen = (el: Element) => !!pin && pin.contains(el);
      const atEnd = gap >= n;
      const el = atEnd ? hdr(n - 1) : hdr(gap);
      if (!el) {
        const idx = [...grid.querySelectorAll<HTMLElement>('[data-testid^="grid-header-"]')]
          .filter((e) => !frozen(e))
          .map((e) => Number(e.dataset.testid!.slice('grid-header-'.length)));
        return idx.length && gap < Math.min(...idx) ? 'left' : 'right';
      }
      const b = el.getBoundingClientRect();
      const x = atEnd ? b.right - 3 : b.left + Math.min(6, b.width / 4);
      const y = b.top + b.height / 2;
      if (frozen(el)) return { x, y };
      if (x < leftEdge + EDGE) return grid.scrollLeft <= 0 ? { x: Math.max(x, leftEdge + 2), y } : 'left';
      if (x > rightEdge - EDGE) return atEnd && grid.scrollLeft + grid.clientWidth >= grid.scrollWidth - 1 ? { x: Math.min(x, rightEdge - 4), y } : 'right';
      return { x, y };
    },
    { gap, n },
  );
}

/** Result of the export dialog: a saved path (server mode) or a downloaded file. */
export async function runExport(page: Page, errors: { messages: string[] }, engine: 'js' | 'excel', fileName: string, outDir: string): Promise<string> {
  await requireTestIds(page, ['export-button'], 'Excel出力ボタン');
  await page.getByTestId('export-button').click();
  await requireTestIds(page, ['export-dialog', 'export-filename', `export-engine-${engine}`, 'export-confirm'], '出力ダイアログ');
  const dialog = page.getByTestId('export-dialog');
  await expect(dialog).toBeVisible();
  await page.getByTestId('export-filename').fill(fileName);
  await page.getByTestId(`export-engine-${engine}`).check();
  const download = page.waitForEvent('download', { timeout: 5 * 60 * 1000 }).catch(() => null);
  await page.getByTestId('export-confirm').click();
  const result = page.getByTestId('export-result');
  const failed = page.getByTestId('export-error');
  await step(page, errors, `出力（${engine}）`, async () => {
    await expect(result.or(failed)).toBeVisible({ timeout: 5 * 60 * 1000 });
    if (await failed.isVisible()) throw new Error(`出力に失敗しました: ${await failed.innerText()}`);
  });
  const text = await result.innerText();
  let file: string;
  const pathEl = result.locator('.result-path');
  if (await pathEl.count()) {
    file = (await pathEl.innerText()).trim();
  } else {
    const d = await download;
    if (!d) throw new Error(`出力結果にパスもダウンロードもありません: ${text}`);
    fs.mkdirSync(outDir, { recursive: true });
    file = path.join(outDir, d.suggestedFilename());
    await d.saveAs(file);
  }
  // runner / patcher warnings shown with the result (MACROS_DROPPED, REF_ERRORS … in Japanese)
  const warn = result.getByTestId('export-warnings');
  if (await warn.count()) test.info().annotations.push({ type: `出力の警告（${engine}）${fileName}`, description: (await warn.innerText()).trim() });
  if (engine === 'excel') {
    // the dialog says so when the Excel engine failed and the JS engine was used instead
    expect(text, 'Excel本体エンジンが失敗して内蔵エンジンに切り替わっていないこと').not.toContain('内蔵エンジンで出力しました');
  }
  const close = page.getByTestId('export-close');
  if (await close.count()) await close.click();
  else await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  expect(fs.existsSync(file), `出力ファイルがあること: ${file}`).toBe(true);
  return file;
}

export interface CheckOutputResult {
  ok: boolean;
  verdict: 'PASS' | 'FAIL';
  openMode: string;
  items: Array<{ level: string; code: string; message: string; examples?: string[] }>;
  stats: Record<string, unknown>;
}

/** a file whose sheets were copied into the output (check-output --source): its path and the name the app saw (default: the file's own name) */
export interface SourceFile {
  path: string;
  name?: string;
}

export interface CheckOptions {
  /** the merged sheet / its header row (default: guessed from the base) */
  sheet?: string;
  headerRow?: number;
  /** the output holds ONLY the merged sheet (the app's default): check-output --single-sheet */
  singleSheet?: boolean;
  /** the output also holds the sheets of the other files (<file>_<sheet>): check-output --with-copies --source ... */
  withCopies?: boolean;
  /** the other files in card order (to compare the copies with their sources) */
  sources?: SourceFile[];
}

/** the check-output flags of a sheet mode (see setSheetMode): single = --single-sheet, base = none, all = --with-copies */
export function checkOptionsFor(mode: SheetMode | undefined, sources?: SourceFile[]): CheckOptions {
  switch (mode) {
    case 'single':
      return { singleSheet: true };
    case 'all':
      return { withCopies: true, ...(sources ? { sources } : {}) };
    case 'copies':
      return { singleSheet: true, withCopies: true, ...(sources ? { sources } : {}) };
    default:
      return {};
  }
}

/** node tools/check-output.mjs <file> --base <base> --json (needs Excel: our own hidden instance). */
export function checkOutput(file: string, base = BASE_SAMPLE, opts: CheckOptions = {}): Promise<CheckOutputResult> {
  return new Promise((resolve, reject) => {
    const extra = [
      ...(opts.sheet ? ['--sheet', opts.sheet] : []),
      ...(opts.headerRow ? ['--header-row', String(opts.headerRow)] : []),
      ...(opts.singleSheet ? ['--single-sheet'] : []),
      ...(opts.withCopies ? ['--with-copies'] : []),
      ...(opts.sources ?? []).flatMap((s) => ['--source', s.path, ...(s.name ? ['--source-name', s.name] : [])]),
    ];
    const child = spawn(process.execPath, [path.join(ROOT, 'tools', 'check-output.mjs'), file, '--base', base, ...extra, '--json'], { cwd: ROOT, windowsHide: true });
    let out = '';
    let err = '';
    child.stdout.on('data', (c) => (out += c.toString('utf8')));
    child.stderr.on('data', (c) => (err += c.toString('utf8')));
    child.on('close', (code) => {
      if (code === 2 || !out.trim()) return reject(new Error(`check-output failed (exit ${code}): ${err}`));
      try {
        resolve(JSON.parse(out));
      } catch (e) {
        reject(new Error(`check-output: invalid JSON: ${String(e)}\n${out.slice(0, 500)}`));
      }
    });
  });
}

/** Failure codes that must never appear in an export (the others may be caused by the user's own edits). */
export const MUST_NOT_FAIL = [
  'inspect-mismatch', // Excel's reading of the file disagrees with the package even after the fresh-instance retries (tools/lib/dump-check.mjs)
  'open-failed',
  'repair',
  'new-ref-text',
  'new-error-#REF!',
  'new-error-#NAME?',
  'formula-columns',
  'formulas-lost',
  'sheets-lost',
  'names-lost',
  'charts-lost',
  'chart-series',
  'pivots-lost',
  'table-lost',
  // the output shapes of round 9 (tools/lib/verdict.mjs, copies.mjs)
  'not-single-sheet', // --single-sheet: the output holds other sheets
  'dangling-sheet-ref', // --single-sheet: a formula reads a sheet that is not in the output
  'copy-missing', // --with-copies: a copy is not there under the name the rule gives
  'copy-unexpected',
  'copy-order',
  'copy-differs', // a copy is not its source sheet
  'copy-new-error',
];

export function describeVerdict(v: CheckOutputResult): string {
  return v.items
    .filter((i) => i.level === 'fail' || i.level === 'warn')
    .map((i) => `[${i.level}] ${i.code}: ${i.message}${i.examples?.length ? '\n    ' + i.examples.slice(0, 4).join('\n    ') : ''}`)
    .join('\n');
}

/** The assertions every export must pass: opens without repair, none of MUST_NOT_FAIL. */
export function assertValid(v: CheckOutputResult, label: string): void {
  test.info().annotations.push({ type: `check-output ${label}`, description: `${v.verdict}
${describeVerdict(v)}` });
  expect(v.openMode, `${label}: Excel で修復なしに開けること
${describeVerdict(v)}`).toBe('normal');
  const fails = v.items.filter((i) => i.level === 'fail' && MUST_NOT_FAIL.includes(i.code));
  expect(fails, `${label}: 出力の検証
${describeVerdict(v)}`).toEqual([]);
}

/** check-output (the CLI) + the assertions every export must pass. */
export async function expectValidExport(file: string, label: string, base = BASE_SAMPLE, opts: CheckOptions = {}): Promise<CheckOutputResult> {
  const v = await checkOutput(file, base, opts);
  assertValid(v, label);
  return v;
}

// ---------------------------------------------------------------------------------------------------------------
// round 2 helpers
// ---------------------------------------------------------------------------------------------------------------

/** Opens the app, loads the given samples and merges them. */
export async function mergeSamples(page: Page, errors: { messages: string[] }, names: string[] = SAMPLE_FILES, before?: () => Promise<void>): Promise<void> {
  await openApp(page);
  await step(page, errors, 'ファイル読み込み', () => loadFiles(page, samplePaths(names)));
  await waitMergeReady(page, errors);
  if (before) await before();
  await mergeNow(page, errors);
  await requireTestIds(page, ['column-list', 'undo-button', 'redo-button', 'export-button'], '結合後の画面');
}

/** Index of a column by its header title (throws with the list when it is missing). */
export async function columnIndex(page: Page, title: string): Promise<number> {
  const titles = await columnTitles(page);
  const i = titles.indexOf(title);
  if (i < 0) throw new Error(`列「${title}」がありません: ${titles.join(' / ')}`);
  return i;
}

/** Selects the columns (click the first, Ctrl+click the others) and returns their current indices. */
export async function selectColumns(page: Page, titles: string[]): Promise<number[]> {
  const idx: number[] = [];
  for (const [k, t] of titles.entries()) {
    const i = await columnIndex(page, t);
    const h = await revealColumn(page, i);
    await h.click(k === 0 ? {} : { modifiers: ['Control'] });
    await expect(h).toHaveAttribute('aria-selected', 'true');
    idx.push(i);
  }
  return idx;
}

/** Selects `titles` and drags them to gap `gap` (0 = before column A, n = at the end). */
export async function moveColumns(page: Page, titles: string[], gap: number): Promise<void> {
  const idx = await selectColumns(page, titles);
  await dragHeader(page, idx[idx.length - 1], gap);
}

/** Types `text` into cell (row, column index) and commits with Enter. */
export async function typeInCell(page: Page, row: number, colIdx: number, text: string): Promise<Locator> {
  await revealColumn(page, colIdx);
  const cell = page.getByTestId(`grid-cell-${row}-${colIdx}`);
  await cell.scrollIntoViewIfNeeded();
  await cell.click();
  await page.keyboard.type(text);
  await page.keyboard.press('Enter');
  return cell;
}

/** Right-click menu of header `colIdx` → 左に列を挿入 / 右に列を挿入. */
export async function insertColumn(page: Page, colIdx: number, side: 'left' | 'right'): Promise<void> {
  const h = await revealColumn(page, colIdx);
  const menu = page.getByTestId('context-menu');
  const box = (await h.boundingBox())!;
  // the right edge of a header is its resize handle (a separate element): right-click in the left part
  const position = { x: Math.max(3, Math.min(10, box.width / 3)), y: Math.min(14, box.height / 2) };
  for (let attempt = 0; attempt < 3 && !(await menu.isVisible()); attempt++) {
    await h.click({ button: 'right', position });
    await menu.waitFor({ state: 'visible', timeout: 4000 }).catch(() => undefined);
  }
  // a real failure (not a fixme): the grid is implemented, the menu must open
  await expect(menu, `列 ${colIdx} の見出しを右クリックしても列のメニューが出ません`).toBeVisible();
  await requireTestIds(page, [`context-menu-item-insert-${side}`], '列のメニュー');
  await page.getByTestId(`context-menu-item-insert-${side}`).click();
  await expect(menu).toHaveCount(0);
}

/** Number of user edits since the merge, from the side summary 「結合後の変更 N 件」. */
export async function editCount(page: Page): Promise<number> {
  const t = await page.getByTestId('side-summary').innerText();
  const m = /結合後の変更\s*(\d+)\s*件/.exec(t);
  return m ? Number(m[1]) : NaN;
}

export interface CompareReport {
  identical: boolean;
  formulaDiffs: number;
  notes: string[];
  excel: { rows: Array<{ group: string; sheet: string; ok: boolean; detail: string; filtered?: boolean }>; diffs: number } | null;
  package: { rows: Array<{ group: string; sheet: string; ok: boolean; detail: string; filtered?: boolean }>; diffs: number } | null;
  text: string;
}

/** node tools/compare.mjs a b --json (+ the text report for messages). `args` e.g. ['--package-only']. */
export function compareOutputs(a: string, b: string, args: string[] = []): CompareReport {
  const run = (extra: string[]) =>
    spawnSync(process.execPath, [path.join(ROOT, 'tools', 'compare.mjs'), a, b, ...args, ...extra], { cwd: ROOT, encoding: 'utf8', timeout: 10 * 60 * 1000, windowsHide: true });
  const j = run(['--json']);
  if (j.status === 2 || !j.stdout.trim()) throw new Error(`compare.mjs を実行できませんでした (exit ${j.status}): ${j.stderr}`);
  const report = JSON.parse(j.stdout) as Omit<CompareReport, 'text'>;
  // the text form, rebuilt from the JSON rows (no second Excel run)
  const lines: string[] = [];
  for (const [view, r] of [['Excel', report.excel], ['package', report.package]] as const) {
    if (!r) continue;
    for (const row of r.rows) if (!row.ok && !row.filtered) lines.push(`[${view}] ${row.group}${row.sheet ? ' / ' + row.sheet : ''}: ${row.detail}`);
  }
  const filtered = [...(report.excel?.rows ?? []), ...(report.package?.rows ?? [])].filter((r) => r.filtered).map((r) => `  (filtered) ${r.group}: ${r.detail.slice(0, 160)}`);
  return { ...report, text: [...(lines.length ? lines : ['差分なし']), ...report.notes.map((n) => `(${n})`), ...filtered].join('\n') };
}

// ---------------------------------------------------------------------------------------------------------------
// shared by the round 2 / round 3 flows: export with both engines, verify, compare, probe
// ---------------------------------------------------------------------------------------------------------------

export const OUT_DIR = path.join(ROOT, 'test-results', 'e2e-downloads');

export interface Outputs {
  files: Array<{ engine: 'js' | 'excel'; file: string }>;
  /** which sheets the output holds (read from the settings bar when it was exported): decides what check-output is asked (--single-sheet / --with-copies) */
  mode?: SheetMode;
  /** mode 'all': the other files in card order (their sheets are in the output as copies) */
  sources?: SourceFile[];
}

export const engineLabel = (e: 'js' | 'excel') => (e === 'js' ? '内蔵エンジン' : 'Excel本体');

/**
 * Exports with the JS engine and (server mode + Excel installed) with the Excel engine. The sheet mode is read from the settings bar
 * (the two checkboxes keep their state after the merge, greyed out), so verifyAndProbe asks check-output for the right output shape.
 * `sources`: the files whose sheets were copied (mode 'all'); default: the sample files that are on the cards.
 */
export async function exportBoth(page: Page, errors: { messages: string[] }, stem: string, o: { sources?: SourceFile[] } = {}): Promise<Outputs> {
  const mode = await sheetModeOf(page);
  const files: Outputs['files'] = [];
  files.push({ engine: 'js', file: await runExport(page, errors, 'js', `${stem}_内蔵エンジン.xlsx`, OUT_DIR) });
  if (await excelInstalled(page)) files.push({ engine: 'excel', file: await runExport(page, errors, 'excel', `${stem}_Excel本体.xlsx`, OUT_DIR) });
  else test.info().annotations.push({ type: 'Excel本体', description: 'Excel本体エンジンは使えないため内蔵エンジンのみ出力しました' });
  const sources = mode === 'all' || mode === 'copies' ? o.sources ?? cardMemory.get(page)?.others : undefined;
  return { files, mode, ...(sources ? { sources } : {}) };
}

// ---------------------------------------------------------------------------------------------------------------
// round 9: which sheets the output holds (the two checkboxes of the settings bar)
// ---------------------------------------------------------------------------------------------------------------

/**
 * single = 「他のシートも出力に含める」 off (the app's default): ONLY the merged sheet
 * base   = only that box ticked and 「他のファイルのシートも含める」 cleared again: the merged sheet + the base file's other sheets (what rounds 1-6 produced)
 * all    = both ticked (ticking the first one ticks the second one too): + the original sheets of every other file as <file>_<sheet>
 * copies = (API level only: the UI cannot reach it) the merged sheet + the copies, without the base's other sheets
 */
export type SheetMode = 'single' | 'base' | 'all' | 'copies';

/** What the settings bar says now (the checkboxes are greyed out, but keep their state, once the merge is done). */
export async function sheetModeOf(page: Page): Promise<SheetMode> {
  const other = page.getByTestId('include-other-sheets');
  const source = page.getByTestId('include-source-sheets');
  if (!(await other.count())) return 'single';
  const o = await other.isChecked();
  const s = (await source.count()) ? await source.isChecked() : false;
  return o ? (s ? 'all' : 'base') : s ? 'copies' : 'single';
}

/** Sets the two checkboxes BEFORE the merge (they are locked afterwards): single / base / all. */
export async function setSheetMode(page: Page, mode: Exclude<SheetMode, 'copies'>): Promise<void> {
  await requireTestIds(page, ['include-other-sheets', 'include-source-sheets'], '出力するシートの設定');
  const other = page.getByTestId('include-other-sheets');
  const source = page.getByTestId('include-source-sheets');
  await other.setChecked(mode !== 'single');
  // ticking the first box ticks the second one too: 'base' clears it again
  await source.setChecked(mode === 'all');
  await expect(other).toBeChecked({ checked: mode !== 'single' });
  await expect(source).toBeChecked({ checked: mode === 'all' });
}

/** The files on the cards, as check-output wants them: the base file and the other files in card order (only the samples can be found by name). */
export async function cardSources(page: Page): Promise<{ base: SourceFile | null; others: SourceFile[] }> {
  const n = await page.locator('[data-testid^="file-card-"]').count();
  let base: SourceFile | null = null;
  const others: SourceFile[] = [];
  for (let i = 0; i < n; i++) {
    const name = (await page.getByTestId(`file-card-${i}`).locator('h2').first().innerText()).trim();
    const full = path.join(SAMPLES_DIR, name);
    if (!fs.existsSync(full)) continue;
    if ((await page.getByTestId(`file-excluded-${i}`).count()) > 0) continue; // 結合対象外: takes no part
    const isBase = (await page.getByTestId(`base-toggle-${i}`).getAttribute('aria-pressed')) === 'true';
    if (isBase) base = { path: full, name };
    else others.push({ path: full, name });
  }
  return { base, others };
}

/**
 * A dump that needed a second Excel start (tools/lib/inspect.mjs: Excel answered while busy and its reading disagreed with the
 * package) is worth knowing about: it is added to the test's annotations and to test-results/e2e-inspect-retries.log.
 */
export function noteInspectionRetries(dumps: InspectResult[], files: string[]): void {
  const retried = dumps.map((d, i) => ({ d, f: files[i] })).filter((x) => (x.d.attempts ?? 1) > 1);
  if (!retried.length) return;
  const text = retried.map(({ d, f }) => `${path.basename(f)}: ${d.attempts} 回目の Excel で読み取り${d.suspicious?.length ? `（まだ不一致: ${d.suspicious.join(' / ')}）` : ''}`).join('\n');
  test.info().annotations.push({ type: 'inspect の再試行', description: text });
  try {
    fs.appendFileSync(path.join(ROOT, 'test-results', 'e2e-inspect-retries.log'), `${new Date().toISOString()} ${test.info().title}\n${text}\n`);
  } catch {
    /* the log is a convenience */
  }
}

export interface VerifyResult {
  verdicts: Map<string, CheckOutputResult>;
  /** cell probes (`cells` of verifyAndProbe): the base first, then one entry per output (same order as Outputs.files) */
  probes: Array<Record<string, CellProbe>>;
  /** the full inspect dumps, same order as `probes` */
  dumps: InspectResult[];
}

/**
 * check-output verdict + cell probes for EVERY output with ONE Excel instance (inspect base + outputs together, the
 * verdict is computed in-process by the same tools/lib/verdict.mjs the CLI uses). Every Excel start costs ~80 MB of
 * Office diagnostics in %TEMP% and several seconds, so flows call this once. Without Excel (or QA_NO_EXCEL=1) nothing
 * is checked and an empty result comes back.
 */
export async function verifyAndProbe(out: Outputs, o: { base?: string; check?: CheckOptions; cells?: string[] } = {}): Promise<VerifyResult> {
  const res: VerifyResult = { verdicts: new Map(), probes: [], dumps: [] };
  if (!excelAvailable()) {
    test.info().annotations.push({ type: 'check-output', description: 'Excel がないため check-output は実行していません' });
    return res;
  }
  const base = o.base ?? BASE_SAMPLE;
  // the output shape (single sheet / copies of the other files) comes from the settings bar at the time of the export, unless the caller says
  const shape = { ...checkOptionsFor(out.mode, out.sources), ...(o.check ?? {}) };
  const srcFiles = shape.withCopies ? shape.sources ?? [] : [];
  const files = [base, ...out.files.map((f) => f.file), ...srcFiles.map((s) => s.path)];
  const dumps = await inspectFiles(files, { recalc: true, ...(o.cells?.length ? { probeCells: o.cells } : {}) });
  noteInspectionRetries(dumps, files);
  dumps.forEach((d, i) => {
    if (d.error && d.openMode !== 'failed') throw new Error(`${path.basename(files[i])} を調べられませんでした: ${d.error}`);
  });
  if (dumps[0].openMode === 'failed') throw new Error(`基準ファイルを Excel で開けませんでした: ${base}
${dumps[0].error ?? ''}`);
  res.dumps = dumps;
  res.probes = dumps.map((d) => d.cellProbes ?? {});
  const sources = srcFiles.map((s, k) => ({ name: s.name ?? path.basename(s.path), dump: dumps[1 + out.files.length + k] }));
  out.files.forEach((f, i) => {
    const d = dumps[i + 1];
    const verdict = checkOutputDumps(dumps[0], d, { sheet: shape.sheet, headerRow: shape.headerRow, singleSheet: shape.singleSheet, withCopies: shape.withCopies, sources });
    const v: CheckOutputResult = { ...verdict, openMode: d.openMode } as CheckOutputResult;
    assertValid(v, engineLabel(f.engine));
    res.verdicts.set(f.engine, v);
  });
  return res;
}

/** check-output on every output (skipped without Excel on this machine; QA_NO_EXCEL=1 skips it on purpose). */
export async function verifyAll(out: Outputs, base = BASE_SAMPLE, opts: CheckOptions = {}): Promise<Map<string, CheckOutputResult>> {
  return (await verifyAndProbe(out, { base, check: opts })).verdicts;
}

/** The verdict of check-output itself must be PASS (not just free of the MUST_NOT_FAIL codes). */
export function expectPass(verdicts: Map<string, CheckOutputResult>, what: string): void {
  for (const [engine, v] of verdicts) expect(v.verdict, `${what}（${engineLabel(engine as 'js' | 'excel')}）: check-output は PASS\n${describeVerdict(v)}`).toBe('PASS');
}

/** Same op log, two engines: tools/compare.mjs must show no formula differences (other differences are listed). */
export function expectSameFormulas(out: Outputs, args: string[] = []): void {
  const js = out.files.find((o) => o.engine === 'js');
  const xl = out.files.find((o) => o.engine === 'excel');
  if (!js || !xl) return;
  const c = compareOutputs(js.file, xl.file, args);
  test.info().annotations.push({ type: `compare 内蔵エンジン vs Excel本体 ${args.join(' ')}`, description: c.text.slice(0, 6000) });
  expect(c.formulaDiffs, `内蔵エンジンと Excel本体で数式が違います:\n${c.text}`).toBe(0);
}

/** Opens the files in our own Excel (ONE instance) and returns the dumps incl. `cellProbes` for `cells`. */
export async function inspectProbe(files: string[], cells: string[], recalc: boolean): Promise<InspectResult[]> {
  const dumps = await inspectFiles(files, { recalc, probeCells: cells });
  noteInspectionRetries(dumps, files);
  dumps.forEach((d, i) => {
    if (d.error || !d.cellProbes) throw new Error(`${path.basename(files[i])} を調べられませんでした: ${d.error ?? 'cellProbes なし'}`);
  });
  return dumps;
}
