/**
 * Round 8 (user report: "from the 2nd Excel file on the formulas are not loaded and become plain / blank cells"): the UI side, against
 * the REAL app, the real modules and the real local server, in Microsoft Edge.
 *   - the samples (01 base + 02, 04, 05): the merge preview says the formulas will be kept; after the merge the side panel's
 *     「数式の取り込み」 block says they were kept (02_面接評価: formulas kept, 0 empty), the toast detail names the files, the
 *     imported formula columns really hold formulas;
 *   - a hand-made source whose formulas have NO stored result (what openpyxl / pandas write): the block is amber, says how many cells
 *     are empty and why, shows the hint, the details list the reasons and clickable example rows (click = the grid jumps to the cell).
 *
 *   npx playwright test -c tests/ui/playwright.config.ts formula-report       (builds the UI, starts the server on :5192)
 *   UI_BASE_URL=http://localhost:5192 npx playwright test -c tests/ui/playwright.config.ts formula-report   (a server that is already running)
 *
 * Only the built-in engine is involved (no Excel start). Screenshots: tests/ui/shots/f-*.png.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { expect, test, type Page } from '@playwright/test';
import { buildXlsx } from '../golden/build';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SAMPLES = path.resolve(HERE, '..', '..', 'samples');
const SHOTS = path.join(HERE, 'shots');
fs.mkdirSync(SHOTS, { recursive: true });
const sample = (n: string) => path.join(SAMPLES, n);
const shot = (page: Page, name: string) => page.screenshot({ path: path.join(SHOTS, `f-${name}.png`) });

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

type Upload = string | { name: string; mimeType: string; buffer: Buffer };

async function addFiles(page: Page, files: Upload[]) {
  const payloads = files.map((f) => (typeof f === 'string' ? { name: path.basename(f), mimeType: 'application/octet-stream', buffer: fs.readFileSync(f) } : f));
  await page.getByTestId('file-input').setInputFiles(payloads);
  await settle(page);
}

async function mergeNow(page: Page) {
  await page.getByTestId('merge-button').click();
  await page.getByTestId('grid').waitFor({ timeout: 120_000 });
  await settle(page);
}

/** position (0-based) of the rendered grid column with that header text, or -1 (the grid renders the visible columns only) */
const headerIndex = (page: Page, title: string) =>
  page.$$eval(
    '[data-testid^=grid-header-]',
    (els, t) => els.map((e) => ({ i: Number((e.getAttribute('data-testid') ?? '').replace('grid-header-', '')), t: e.querySelector('.g-htitle')?.textContent ?? '' })).find((x) => x.t === t)?.i ?? -1,
    title,
  );

/** the text of a locator with the white space collapsed */
const flat = async (page: Page, testid: string) => ((await page.getByTestId(testid).innerText()) ?? '').replace(/\s+/g, ' ').trim();

// -------------------------------------------------------------------------------- hand-made files (no stored results)

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const colName = (n: number) => String.fromCharCode(64 + n);
type G = string | number | null | { f: string; v?: number | string };

function sheetData(rows: G[][]): string {
  return rows
    .map(
      (row, i) =>
        `<row r="${i + 1}">${row
          .map((g, j) => {
            const ref = `${colName(j + 1)}${i + 1}`;
            if (g === null) return '';
            if (typeof g === 'number') return `<c r="${ref}"><v>${g}</v></c>`;
            if (typeof g === 'string') return `<c r="${ref}" t="inlineStr"><is><t>${esc(g)}</t></is></c>`;
            // a formula; with a stored result only when `v` is given (openpyxl / pandas / ClosedXML store none)
            return typeof g.v === 'number' ? `<c r="${ref}"><f>${esc(g.f)}</f><v>${g.v}</v></c>` : `<c r="${ref}"><f>${esc(g.f)}</f></c>`;
          })
          .join('')}</row>`,
    )
    .join('');
}

const NAMES = ['山田 太郎', '鈴木 花子', '高橋 健', '田中 美咲', '伊藤 大輔', '渡辺 優', '中村 翔', '小林 彩', '加藤 学', '吉田 恵'];

/** base: 名簿 (ID, 氏名, 学年), 10 students */
async function baseFile(): Promise<Upload> {
  const rows: G[][] = [['学生ID', '氏名', '学年'], ...NAMES.map((n, i): G[] => [`S${String(i + 1).padStart(3, '0')}`, n, 1 + (i % 4)])];
  const buffer = Buffer.from(await buildXlsx([{ name: '名簿', data: sheetData(rows) }]));
  return { name: '10_基準名簿.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer };
}

/**
 * source: the same 10 people (other order) with
 *   E 合計   =C+D                      same-row formula of imported columns   -> stays a formula (no stored result: Excel calculates it)
 *   F 区分   =VLOOKUP(A,Master!…)      reads another sheet                    -> cannot be kept; NO stored result -> EMPTY
 *   G 税込   =C*$J$14                  reads a cell outside the table         -> cannot be kept; NO stored result -> EMPTY
 *   H 参考   =C*Master!$B$2            reads another sheet                    -> cannot be kept; stored result -> becomes the value
 */
async function sourceFile(): Promise<Upload> {
  const order = [...NAMES.keys()].reverse();
  const rows: G[][] = [['ID', '氏名', '筆記', '面接', '合計', '区分', '税込', '参考']];
  order.forEach((k, r) => {
    const n = r + 2;
    rows.push([`S${String(k + 1).padStart(3, '0')}`, NAMES[k], 60 + k, 70 + k, { f: `C${n}+D${n}` }, { f: `VLOOKUP(A${n},Master!A:B,2,FALSE)` }, { f: `C${n}*$J$14` }, { f: `C${n}*Master!$B$2`, v: (60 + k) * 2 }]);
  });
  while (rows.length < 13) rows.push([]);
  rows.push([null, null, null, null, null, null, null, null, null, 1.1]); // J14: the rate, outside the table
  const master: G[][] = [['ID', '係数'], ['S001', 2], ['S002', 3]];
  const buffer = Buffer.from(await buildXlsx([{ name: '評価', data: sheetData(rows) }, { name: 'Master', data: sheetData(master) }]));
  return { name: '11_評価（結果なし）.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer };
}

// ------------------------------------------------------------------------------------------------ the samples

test.describe('samples: formulas of the 2nd and later files stay formulas', () => {
  test('preview hint, block after the merge, toast detail, live formulas in the grid', async ({ page }) => {
    const errors = await openApp(page);
    await addFiles(page, [sample('01_学生名簿.xlsx'), sample('02_面接評価.xlsx'), sample('04_適性検査.xls'), sample('05_研修出欠.xlsx')]);
    await expect(page.getByTestId('merge-button')).toBeEnabled();

    // ---- before the merge: one line under 「追加される列」 (no promise about the result: the preview cannot know it)
    const hint = page.getByTestId('preview-formulas-hint');
    await expect(hint).toBeVisible();
    await expect(hint).toHaveText('数式を含む列があります。結合後も数式のまま残るよう付け替えます');
    // the columns that will be imported and hold formulas, per file (data attribute + tooltip; nothing else is added to the preview)
    await expect(hint).toHaveAttribute('data-formula-columns', 's2:2,s4:1'); // 02: 平均, 判定 / 05: 出席回数 / 04 is .xls: values only, none
    expect((await hint.getAttribute('title')) ?? '').toContain('02_面接評価.xlsx: 数式を含む列 2 列');
    await shot(page, 'preview-hint');
    // the hint follows the columns that will be imported: switching 平均 and 判定 off removes 02 from it (and back)
    const group = page.locator('.import-group', { hasText: '02_面接評価' });
    await group.locator('.col-chip', { hasText: '平均' }).click();
    await expect(hint).toHaveAttribute('data-formula-columns', 's2:1,s4:1');
    await group.locator('.col-chip', { hasText: '判定' }).click();
    await expect(hint).toHaveAttribute('data-formula-columns', 's4:1'); // 05_研修出欠 still brings a formula column
    await group.locator('.col-chip', { hasText: '平均' }).click();
    await group.locator('.col-chip', { hasText: '判定' }).click();
    await expect(hint).toHaveAttribute('data-formula-columns', 's2:2,s4:1');
    expect(await page.getByTestId('formula-report').count()).toBe(0); // nothing to report before the merge

    // ---- merge (the toast goes away after a few seconds: read it first)
    await mergeNow(page);
    const toast = page.getByTestId('toast');
    await expect(toast).toContainText('結合しました');
    const toastText = ((await toast.innerText()) ?? '').replace(/s+/g, ' ');
    const toastDetail = (await toast.getAttribute('title')) ?? '';
    const block = page.getByTestId('formula-report');
    await expect(block).toBeVisible();
    await expect(block).toHaveAttribute('data-tone', 'ok');
    const f2 = await flat(page, 'formula-file-s2');
    expect(f2).toContain('02_面接評価.xlsx');
    expect(f2).toMatch(/数式 [\d,]+ 件を数式のまま取り込みました/);
    expect(f2).not.toContain('空白');
    expect(f2).not.toContain('値にしました');
    await expect(page.getByTestId('formula-empty-s2')).toHaveCount(0);
    await expect(page.getByTestId('formula-hint-s2')).toHaveCount(0);
    const f4 = await flat(page, 'formula-file-s4');
    expect(f4).toContain('05_研修出欠.xlsx');
    expect(f4).toMatch(/数式 [\d,]+ 件を数式のまま取り込みました/);
    await expect(page.getByTestId('formula-file-s3')).toHaveCount(0); // 04_適性検査.xls has no formulas
    await expect(block.locator('.count-badge')).not.toHaveClass(/badge-warn/);
    const kept = (s: string) => Number(/数式 ([\d,]+) 件を数式のまま/.exec(s)![1].replace(/,/g, ''));
    expect(kept(f2)).toBeGreaterThan(50); // 63 people x 2 formula columns in the file
    expect(kept(f4)).toBeGreaterThan(20); // 32 people x 1 formula column

    // ---- the merge toast: counts + one sentence per file in its detail (tooltip)
    expect(toastText).toContain(`数式 ${(kept(f2) + kept(f4)).toLocaleString('ja-JP')} 件は数式のまま`);
    expect(toastDetail).toContain(`02_面接評価.xlsx: 数式 ${kept(f2).toLocaleString('ja-JP')} 件を数式のまま取り込みました`);
    expect(toastDetail).toContain('05_研修出欠.xlsx: 数式');
    await shot(page, 'merged-block');

    // ---- the 警告 list does not repeat the formula items of the block (nothing went wrong)
    await page.getByTestId('warning-details').locator('summary').click();
    const warnings = await page.getByTestId('warning-list').innerText();
    expect(warnings).not.toContain('数式のまま残しました');
    expect(warnings).not.toContain('値に変換しました（参照先の列が取り込まれていないため）');

    // ---- the imported formula columns really hold formulas (the cards click centres the file's columns)
    await page.locator('[data-file-id="s2"]').click();
    await expect.poll(() => headerIndex(page, '平均'), { message: '平均 column of 02_面接評価 is scrolled into view' }).toBeGreaterThanOrEqual(0);
    const avg = await headerIndex(page, '平均');
    await page.getByTestId(`grid-cell-4-${avg}`).click();
    await expect(page.getByTestId('formula-bar-input')).toHaveValue(/^=AVERAGE\(/);
    const judge = await headerIndex(page, '判定');
    if (judge >= 0) {
      await page.getByTestId(`grid-cell-4-${judge}`).click();
      await expect(page.getByTestId('formula-bar-input')).toHaveValue(/^=IF\(/);
    }
    await shot(page, 'merged-grid');
    expect(errors).toEqual([]);
  });
});

// ------------------------------------------------------------------------------------------------ no stored results

test.describe('a source whose formulas have no stored result', () => {
  test('amber block: kept / values / empty, the hint, reasons, example rows that jump to the cell', async ({ page }) => {
    const errors = await openApp(page);
    await addFiles(page, [await baseFile(), await sourceFile()]);
    await expect(page.getByTestId('merge-button')).toBeEnabled();
    const hint = page.getByTestId('preview-formulas-hint');
    await expect(hint).toBeVisible();
    await shot(page, 'amber-preview');
    await mergeNow(page);
    const toastDetail = (await page.getByTestId('toast').getAttribute('title')) ?? '';

    const block = page.getByTestId('formula-report');
    await expect(block).toBeVisible();
    await expect(block).toHaveAttribute('data-tone', 'warn');
    await expect(block.locator('.count-badge')).toHaveClass(/badge-warn/);
    const file = page.getByTestId('formula-file-s2');
    await expect(file).toHaveAttribute('data-tone', 'warn');
    const text = await flat(page, 'formula-file-s2');
    expect(text).toContain('11_評価（結果なし）.xlsx');
    expect(text).toContain('数式 10 件を数式のまま取り込みました'); // 合計
    expect(text).toContain('10 件は値にしました'); // 参考 (had a stored result)
    expect(text).toContain('空白 20 件（元のファイルに計算結果がありません）'); // 区分, 税込 (no result, cannot be kept)
    // with empty cells the breakdown of the values alone is not claimed on that line
    expect(await flat(page, 'formula-values-s2')).not.toContain('内訳');
    await expect(page.getByTestId('formula-hint-s2')).toHaveText('元のExcelで一度開いて保存し直すと取り込めることがあります');
    await expect(page.getByTestId('formula-recalc-s2')).toContainText('Excel で開いたときに表示されます');
    await shot(page, 'amber-block');

    // ---- details: reasons (combined) and example rows
    const details = page.getByTestId('formula-details-s2');
    await details.locator('summary').click();
    const dt = (await details.innerText()).replace(/\s+/g, ' ');
    expect(dt).toContain('値にした・空白になった数式の内訳');
    expect(dt).toContain('他のシートを参照 20 件');
    expect(dt).toContain('表の外のセル 10 件');
    await shot(page, 'amber-details');
    const rowButtons = details.locator('.formula-row-btn[data-row]');
    expect(await rowButtons.count()).toBeGreaterThan(0);

    // ---- click an example row: the grid jumps to a cell of that row in a column of this file that held a formula
    const first = rowButtons.first();
    const row = Number(await first.getAttribute('data-row'));
    await first.click();
    await expect(page.getByTestId('formula-bar-address')).toContainText(String(row));
    const addr = (await page.getByTestId('formula-bar-address').innerText()).trim();
    expect(addr).toMatch(new RegExp(`^[A-Z]+${row}$`));
    await shot(page, 'amber-jump');

    // ---- the side lists: 警告 has one entry for the file; the check panel lists the example rows
    await page.getByTestId('warning-details').locator('summary').click();
    const warnings = await page.getByTestId('warning-list').innerText();
    expect(warnings).toContain('11_評価（結果なし）.xlsx: 10 件は値にしました／空白 20 件');
    const toValues = page.locator('[data-group="toValues"]');
    await expect(toValues).toBeVisible();
    await toValues.locator('.check-group-head').click();
    expect(await toValues.locator('.check-item-btn').count()).toBeGreaterThan(0);

    // ---- toast detail: the sentence of the file and the hint
    expect(toastDetail).toContain('11_評価（結果なし）.xlsx: 数式 10 件を数式のまま取り込みました／10 件は値にしました／空白 20 件（元のファイルに計算結果がありません）');
    expect(toastDetail).toContain('元のExcelで一度開いて保存し直すと取り込めることがあります');

    // ---- kept formulas without a stored result are not blank in the grid: the formula is shown (grey, italic) until Excel calculates it
    await page.locator('[data-file-id="s2"]').click();
    await expect.poll(() => headerIndex(page, '合計'), { message: '合計 column is scrolled into view' }).toBeGreaterThanOrEqual(0);
    const total = await headerIndex(page, '合計');
    const cell = page.getByTestId(`grid-cell-2-${total}`);
    await expect(cell).toHaveClass(/is-ftext/);
    expect(((await cell.innerText()) ?? '').trim()).toMatch(/^=/);
    await shot(page, 'amber-grid');
    expect(errors).toEqual([]);
  });
});

// ------------------------------------------------------------------------------------------------ rows and columns move

const letter = (i: number): string => {
  let s = '';
  let x = i + 1;
  while (x > 0) {
    const r = (x - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
};

test.describe('the imported formulas stay consistent when rows are matched by key and columns are moved', () => {
  test('every row reads its OWN row (02 lists the people in another order); moving a column rewrites the formulas that read it', async ({ page }) => {
    const errors = await openApp(page);
    await addFiles(page, [sample('01_学生名簿.xlsx'), sample('02_面接評価.xlsx')]);
    await mergeNow(page);
    const columnTitles = async () => (await page.locator('[data-testid=column-list] li .col-title').allInnerTexts()).map((s) => s.trim());
    const formulaAt = async (title: string, row: number) => {
      await page.locator(`[data-testid=column-list-item-${(await columnTitles()).indexOf(title)}]`).click(); // the grid scrolls to the column
      await expect.poll(() => headerIndex(page, title), { message: `${title} is scrolled into view` }).toBeGreaterThanOrEqual(0);
      await page.getByTestId(`grid-cell-${row}-${await headerIndex(page, title)}`).click();
      await expect(page.getByTestId('formula-bar-address')).toContainText(String(row));
      return page.getByTestId('formula-bar-input').inputValue();
    };

    // ---- rows: the merged row r holds the person of base row r; the formula of that row must read row r (not the source file's row)
    const before = new Map<number, string>();
    for (const row of [4, 11, 25]) {
      const f = await formulaAt('平均', row);
      const m = /^=AVERAGE\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)$/.exec(f);
      expect(m, `平均 of row ${row}: ${f}`).not.toBeNull();
      expect([Number(m![2]), Number(m![4])], `the formula of row ${row} reads row ${row}: ${f}`).toEqual([row, row]);
      before.set(row, f);
      const j = await formulaAt('判定', row);
      expect(j).toMatch(new RegExp(`^=IF[(][A-Z]+${row}>=3[.]5,`));
    }

    // ---- columns: move 平均 in front of 面接官 (the first column of 02): 判定 reads 平均 at its NEW letter
    const titles = await columnTitles();
    const iAvg = titles.indexOf('平均');
    const iFirst = titles.indexOf('面接官');
    expect(iAvg).toBeGreaterThan(iFirst);
    // (Alt+↑ in the column list = the same move as the drag; the list is short in this window, so a drag would have to scroll it)
    await page.locator(`[data-testid=column-list-item-${iAvg}]`).click();
    for (let k = 0; k < iAvg - iFirst; k++) {
      await page.keyboard.press('Alt+ArrowUp');
      await expect.poll(async () => (await columnTitles()).indexOf('平均')).toBe(iAvg - k - 1);
    }
    const moved = await columnTitles();
    const avgLetter = letter(moved.indexOf('平均'));
    for (const row of [4, 11, 25]) {
      const j = await formulaAt('判定', row);
      expect(j, `判定 of row ${row} reads 平均 at its new place`).toBe(`=IF(${avgLetter}${row}>=3.5,"合格","保留")`);
      // the moved formula itself still reads the three 評価 columns of ITS row, wherever they are now
      const f = await formulaAt('平均', row);
      const m = /^=AVERAGE\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)$/.exec(f);
      expect(m, `平均 of row ${row} after the move: ${f}`).not.toBeNull();
      expect([Number(m![2]), Number(m![4])]).toEqual([row, row]);
    }
    expect(before.size).toBe(3);
    await shot(page, 'moved-columns');
    expect(errors).toEqual([]);
  });
});
