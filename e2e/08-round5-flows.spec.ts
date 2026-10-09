// Round 5 flows on the samples (owner: T), the ways a real person uses the tool twice over:
//   (a) three files in three different orders (each file may be the base) — every merge lines up with truth.json
//   (b) export twice in a row (same name, server mode never overwrites) — the same workbook both times
//   (c) export, edit, export, undo all the way back, export — equals the export of the untouched merge; redo all the way — equals
//       the edited export; both engines
//   (d) the 警告 section for sample 01: what it says before the merge, after it, and when files come and go
//   (e) Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z after several operations, with the focus in different places
import fs from 'node:fs';
import path from 'node:path';
import { expect, test, type Page } from '@playwright/test';
import {
  BASE_SAMPLE,
  columnIndex,
  columnTitles,
  compareOutputs,
  editCount,
  excelInstalled,
  insertColumn,
  isServerMode,
  loadFiles,
  mergeNow,
  mergeSamples,
  moveColumns,
  openApp,
  OUT_DIR,
  runExport,
  SAMPLE_FILES,
  SAMPLES_DIR,
  samplePaths,
  selectColumns,
  step,
  typeInCell,
  waitMergeReady,
  watchErrors,
} from './helpers';
import { alignmentProblems, readSheet, samePackage, verifyCases, waitIdle } from './round5';

const TWO = [SAMPLE_FILES[0], SAMPLE_FILES[1]];
const BASE_HEADERS = ['学生ID', '氏名', 'フリガナ', '大学', '学部', '学年', 'メール', '電話', '生年月日', '年齢', '志望度', '重複チェック'];
const samplePath = (name: string) => path.join(SAMPLES_DIR, name);

// ---------------------------------------------------------------------------------------------------------------------------------
// (a) three files, three orders
// ---------------------------------------------------------------------------------------------------------------------------------

interface OrderCase {
  label: string;
  files: string[];
  /** header row of the merged sheet in the output */
  headerRow: number;
  sheet?: string;
  /** the base's own columns come first, in this order */
  baseHeaders: string[];
  /** columns that must be among the imported ones (header text before any （ファイル名） suffix) */
  imported: string[];
  /** joined files whose columns are compared with truth.json */
  compare: string[];
}

const ORDERS: OrderCase[] = [
  {
    label: '01 → 02 → 03（名簿が基準）',
    files: ['01_学生名簿.xlsx', '02_面接評価.xlsx', '03_連絡先.csv'],
    headerRow: 3,
    sheet: '名簿',
    baseHeaders: BASE_HEADERS,
    imported: ['面接官', '評価1', '評価2', '評価3', '平均', '判定', 'コメント'],
    compare: ['02_面接評価.xlsx'],
  },
  {
    label: '03 → 01 → 02（連絡先 CSV が基準）',
    files: ['03_連絡先.csv', '01_学生名簿.xlsx', '02_面接評価.xlsx'],
    headerRow: 1,
    baseHeaders: ['氏名', 'メール', '電話'],
    imported: ['フリガナ', '大学', '学部', '学年', '志望度', '面接官', '評価1', '平均', '判定', 'コメント'],
    compare: ['01_学生名簿.xlsx', '02_面接評価.xlsx'],
  },
  {
    label: '02 → 03 → 01（面接評価が基準）',
    files: ['02_面接評価.xlsx', '03_連絡先.csv', '01_学生名簿.xlsx'],
    headerRow: 1,
    sheet: '面接評価',
    baseHeaders: ['受験番号', '氏名', '面接官', '評価1', '評価2', '評価3', '平均', '判定', 'コメント'],
    imported: ['メール', '電話', 'フリガナ', '大学', '学部', '学年', '志望度'],
    compare: ['03_連絡先.csv', '01_学生名簿.xlsx'],
  },
];

const stripSuffix = (t: string) => t.replace(/（[^（）]*）$/, '');

test.describe.serial('ラウンド5: 3 ファイルを別の順番で投入して結合（どれが基準でも正しい）', () => {
  const made: Array<{ order: OrderCase; file: string }> = [];

  for (const order of ORDERS) {
    test(`(a) ${order.label}: 結合でき、列の並びと行の対応が正しく、出力できる`, async ({ page }) => {
      test.setTimeout(10 * 60_000);
      const errors = watchErrors(page);
      await openApp(page);
      await step(page, errors, 'ファイル読み込み', () => loadFiles(page, samplePaths(order.files)));
      await waitMergeReady(page, errors);
      await waitIdle(page);
      await expect(page.getByTestId('base-toggle-0'), '最初に投入したファイルが基準').toHaveAttribute('aria-pressed', 'true');
      await expect(page.locator('[data-testid^="file-excluded-"]'), 'どのファイルにもキーがある').toHaveCount(0);

      // the first order also imports a KEY column by hand (03's メール is a key against 01: not imported unless asked)
      const imported = [...order.imported];
      const compare = [...order.compare];
      if (order.files[0] === '01_学生名簿.xlsx') {
        const chip = page.locator('.import-group').filter({ hasText: '03_連絡先' }).locator('.col-chip', { hasText: /メール$/ }); // a chip reads 「メール」, once imported 「＋ メール」
        await expect(chip, 'キーの列は既定では取り込まない').toHaveAttribute('aria-pressed', 'false');
        await chip.click();
        await expect(chip).toHaveAttribute('aria-pressed', 'true');
        imported.push('メール（03_連絡先）');
        compare.push('03_連絡先.csv');
      }
      await mergeNow(page, errors);
      const titles = await columnTitles(page);
      test.info().annotations.push({ type: `${order.label} の列`, description: titles.join(' / ') });
      expect(titles.slice(0, order.baseHeaders.length), '基準の列が先頭に、元の順で並ぶ').toEqual(order.baseHeaders);
      expect(titles.filter((t, i) => titles.indexOf(t) !== i), '見出しの重複がない').toEqual([]);
      for (const h of imported) expect(titles.includes(h) || titles.map(stripSuffix).includes(h), `取り込んだ列「${h}」`).toBe(true);

      const name = `e2e_r5a_${order.files.map((f) => f.slice(0, 2)).join('')}.xlsx`;
      const file = await runExport(page, errors, 'js', name, OUT_DIR);
      made.push({ order, file });
      const grid = readSheet(file, order.sheet);
      expect(grid.rows[order.headerRow - 1].slice(0, titles.length), '出力の見出し行が結合後の列と同じ').toEqual(titles);
      const { problems, checked } = alignmentProblems(grid, order.files[0], compare);
      expect(checked, '照合したセルの数').toBeGreaterThan(40);
      expect(problems, `行の対応（truth.json）\n${problems.slice(0, 8).join('\n')}`).toEqual([]);
      expect(errors.messages.filter((m) => !/favicon|DevTools|Failed to load resource/i.test(m))).toEqual([]);
    });
  }

  test('(a) 3 つの出力を Excel で開く（1 回の起動で 3 つとも）: 修復なし・#REF! なし・同じ見出しの列を指す', async () => {
    test.setTimeout(10 * 60_000);
    test.skip(made.length === 0, '前のテストが出力を作れませんでした');
    const v = await verifyCases(
      made.map(({ order, file }) => ({
        label: order.label,
        base: samplePath(order.files[0]),
        file,
        check: { ...(order.sheet ? { sheet: order.sheet } : {}), headerRow: order.headerRow },
      })),
    );
    for (const [i, r] of v.entries()) {
      test.info().annotations.push({ type: `check-output ${made[i].order.label}`, description: r.verdict });
      // a CSV base has no formulas of its own to follow: its check is "opens, no #REF!, nothing lost"
      expect(r.openMode).toBe('normal');
    }
  });
});

// ---------------------------------------------------------------------------------------------------------------------------------
// (b) export twice
// ---------------------------------------------------------------------------------------------------------------------------------

test('(b) 続けて 2 回出力: 同じ名前でも上書きされず（サーバー時）、2 つとも同じブックで、2 回目もそのまま使える', async ({ page }) => {
  test.setTimeout(15 * 60_000);
  const errors = watchErrors(page);
  await mergeSamples(page, errors, TWO);
  const server = await isServerMode(page);
  const name = `e2e_r5b_2回出力_${Date.now()}.xlsx`;

  // the dialog proposes <基準ファイル名>_統合_<yyyyMMdd>.xlsx
  await page.getByTestId('export-button').click();
  const d = new Date();
  const stamp = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
  await expect(page.getByTestId('export-filename'), '既定のファイル名').toHaveValue(`01_学生名簿_統合_${stamp}.xlsx`);
  await page.keyboard.press('Escape');
  await expect(page.getByTestId('export-dialog')).toBeHidden();

  const first = await runExport(page, errors, 'js', name, path.join(OUT_DIR, 'first'));
  await expect(page.getByTestId('export-button'), '出力が終わればボタンは元に戻る').toHaveText('Excel出力');
  const second = await runExport(page, errors, 'js', name, path.join(OUT_DIR, 'second'));
  expect(second).not.toBe(first);
  if (server) {
    // the server never overwrites: the second file gets " (2)"
    expect(path.dirname(second)).toBe(path.dirname(first));
    expect(path.basename(second), '2 回目は別のファイル名（上書きしない）').toMatch(/ \(2\)\.xlsx$/);
    expect(fs.existsSync(first), '1 回目のファイルは残っている').toBe(true);
  }
  const a = samePackage(first, second);
  expect(a.identical, `2 回の出力は同じブック\n${a.text}`).toBe(true);
  expect(fs.statSync(first).size).toBe(fs.statSync(second).size);

  // and with the Excel engine twice (its hidden Excel is started and closed for every export)
  if (await excelInstalled(page)) {
    const x1 = await runExport(page, errors, 'excel', `e2e_r5b_excel_${Date.now()}.xlsx`, path.join(OUT_DIR, 'excel1'));
    const x2 = await runExport(page, errors, 'excel', `e2e_r5b_excel_${Date.now()}_2.xlsx`, path.join(OUT_DIR, 'excel2'));
    const b = samePackage(x1, x2);
    expect(b.identical, `Excel本体で 2 回出力しても同じブック\n${b.text}`).toBe(true);
    const c = compareOutputs(first, x1, ['--package-only']);
    expect(c.formulaDiffs, `内蔵エンジンと Excel本体で数式が同じ\n${c.text}`).toBe(0);
  }
  expect(errors.messages.filter((m) => !/favicon|DevTools|Failed to load resource/i.test(m))).toEqual([]);
});

// ---------------------------------------------------------------------------------------------------------------------------------
// (c) undo all the way back
// ---------------------------------------------------------------------------------------------------------------------------------

/** Clicks 元に戻す until it is disabled; returns the number of steps. */
async function undoAll(page: Page): Promise<number> {
  let n = 0;
  const undo = page.getByTestId('undo-button');
  while (await undo.isEnabled()) {
    const before = await editCount(page);
    await undo.click();
    n++;
    await expect.poll(() => editCount(page), { message: `元に戻す ${n} 回目で結合後の変更が 1 件減る` }).toBe(before - 1);
    if (n > 50) throw new Error('元に戻すが終わりません');
  }
  return n;
}

async function redoAll(page: Page): Promise<number> {
  let n = 0;
  const redo = page.getByTestId('redo-button');
  while (await redo.isEnabled()) {
    await redo.click();
    n++;
    if (n > 50) throw new Error('やり直すが終わりません');
  }
  return n;
}

/** A sequence of different edits (each one undo step): returns the titles after them. */
async function someEdits(page: Page, errors: { messages: string[] }): Promise<string[]> {
  const waitEdits = (n: number, what: string) =>
    step(page, errors, what, async () => {
      await expect.poll(() => editCount(page), { message: `${what}: 結合後の変更 ${n} 件`, timeout: 15_000 }).toBe(n);
    });
  await moveColumns(page, ['メール'], 2); // across the merged title → confirm
  await expect(page.getByTestId('unmerge-confirm-dialog')).toBeVisible();
  await page.getByTestId('confirm-ok').click();
  await waitEdits(1, 'メールを移動（結合を解除）');
  await moveColumns(page, ['判定'], 0);
  await waitEdits(2, '判定を先頭へ');
  await insertColumn(page, await columnIndex(page, '志望度'), 'right');
  await waitEdits(3, '列の挿入');
  const at = (await columnIndex(page, '志望度')) + 1;
  await typeInCell(page, 4, at, 'メモ');
  await waitEdits(4, '文字の入力');
  await selectColumns(page, ['平均']);
  await page.keyboard.press('Delete');
  await expect(page.getByTestId('delete-confirm-dialog')).toBeVisible();
  await page.getByTestId('delete-convert-button').click();
  await waitEdits(5, '平均を値に変換して削除');
  return columnTitles(page);
}

test('(c) 出力 → 編集 → 出力 → 元に戻すを最後まで → 出力は「何もしていない結合」と同じ。やり直すを最後まで → 編集後の出力と同じ（両エンジン）', async ({ page }) => {
  test.setTimeout(20 * 60_000);
  const errors = watchErrors(page);
  await mergeSamples(page, errors, TWO);
  const initial = await columnTitles(page);
  const stem = `e2e_r5c_${Date.now()}`;

  const premerge = await runExport(page, errors, 'js', `${stem}_A_結合直後.xlsx`, OUT_DIR);
  const edited = await someEdits(page, errors);
  expect(edited).not.toEqual(initial);
  const changed = await runExport(page, errors, 'js', `${stem}_X_編集後.xlsx`, OUT_DIR);
  expect(samePackage(premerge, changed).identical, '編集した結果は結合直後の出力と違う').toBe(false);

  // undo all the way: the history gives back every state, in the end the untouched merge
  const steps = await undoAll(page);
  expect(steps, '元に戻せた回数 = 編集の回数').toBe(5);
  expect(await editCount(page)).toBe(0);
  await expect(page.getByTestId('undo-button')).toBeDisabled();
  await expect(page.getByTestId('redo-button')).toBeEnabled();
  expect(await columnTitles(page), '列の並びが結合直後に戻る').toEqual(initial);
  const back = await runExport(page, errors, 'js', `${stem}_B_全部戻した.xlsx`, OUT_DIR);
  const same = samePackage(premerge, back);
  expect(same.identical, `全部元に戻してからの出力は、結合直後の出力と同じ\n${same.text}`).toBe(true);

  // redo all the way: the edited workbook again
  expect(await redoAll(page)).toBe(5);
  expect(await editCount(page)).toBe(5);
  expect(await columnTitles(page)).toEqual(edited);
  const again = await runExport(page, errors, 'js', `${stem}_C_やり直した.xlsx`, OUT_DIR);
  const same2 = samePackage(changed, again);
  expect(same2.identical, `全部やり直してからの出力は、編集後の出力と同じ\n${same2.text}`).toBe(true);

  // the Excel engine replays an EMPTY op log too: same formulas as the built-in engine's untouched merge
  if (await excelInstalled(page)) {
    await undoAll(page);
    expect(await editCount(page)).toBe(0);
    const viaExcel = await runExport(page, errors, 'excel', `${stem}_D_全部戻した_Excel本体.xlsx`, OUT_DIR);
    const c = compareOutputs(premerge, viaExcel);
    test.info().annotations.push({ type: '結合直後(内蔵) vs 全部戻した(Excel本体)', description: c.text.slice(0, 3000) });
    expect(c.formulaDiffs, `Excel本体でも、全部戻した出力の数式は結合直後と同じ\n${c.text}`).toBe(0);
  }
  await verifyCases([{ label: '全部戻した出力', base: BASE_SAMPLE, file: back }]);
});

// ---------------------------------------------------------------------------------------------------------------------------------
// (d) 警告
// ---------------------------------------------------------------------------------------------------------------------------------

async function warningState(page: Page): Promise<{ count: string; badgeClass: string; open: boolean; items: Array<{ text: string; severity: string }> }> {
  const details = page.getByTestId('warning-details');
  const badge = details.locator('summary .count-badge');
  const open = await details.evaluate((el) => (el as HTMLDetailsElement).open);
  await details.evaluate((el) => ((el as HTMLDetailsElement).open = true));
  const items = await page.locator('[data-testid="warning-list"] li').evaluateAll((els) =>
    els.map((e) => ({ text: (e.textContent ?? '').replace(/\s+/g, ' ').trim(), severity: (e.className.match(/sev-(\w+)/) ?? [])[1] ?? '' })),
  );
  return { count: (await badge.innerText()).trim(), badgeClass: (await badge.getAttribute('class')) ?? '', open, items };
}

test.describe('ラウンド5: 警告（サンプル 01 の例）', () => {
  test('(d) 01 だけ → グラフとピボットの注意。ファイルを足すと形式の注意が増え、結合後は結合の注意が増え、外すと消える', async ({ page }) => {
    test.setTimeout(10 * 60_000);
    const errors = watchErrors(page);
    await openApp(page);
    const server = await isServerMode(page);
    const excel = await excelInstalled(page);
    const hint = server && excel ? 'Excel本体モードで確実に保持されます' : 'Excel本体モード（Excel のある PC でアプリを起動）で出力すると確実です';

    // 01 alone: the workbook has one chart and one pivot table
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, samplePaths([SAMPLE_FILES[0]])));
    await waitIdle(page);
    let w = await warningState(page);
    test.info().annotations.push({ type: '01 だけの警告', description: w.items.map((i) => `[${i.severity}] ${i.text}`).join('\n') });
    expect(w.open, '既定では閉じている').toBe(false);
    expect(w.count).toBe('2');
    expect(w.badgeClass, '注意（warning）がある').toContain('badge-warn');
    expect(w.items.map((i) => i.text)).toEqual([`グラフ 1 個: ${hint}`, `ピボットテーブル 1 個: 元データ範囲は ${hint}（出力後に「更新」してください）`]);
    expect(w.items.map((i) => i.severity)).toEqual(['info', 'warning']);

    // the other samples: csv / xls keep only values; 02 and 05 add nothing
    await page.getByTestId('file-input').setInputFiles(samplePaths(SAMPLE_FILES.slice(1)));
    await expect(page.locator('[data-testid^="file-card-"]')).toHaveCount(5);
    await waitMergeReady(page, errors);
    await waitIdle(page);
    w = await warningState(page);
    test.info().annotations.push({ type: '5 ファイルの警告', description: w.items.map((i) => `[${i.severity}] ${i.text}`).join('\n') });
    const texts = w.items.map((i) => i.text);
    expect(w.count).toBe('5');
    expect(texts).toContain('03_連絡先.csv: csv 形式のため値のみ取り込みます（書式・数式は引き継がれません）');
    expect(texts.some((t) => t.startsWith('03_連絡先.csv: CSVファイルを読み込みました（文字コード: Shift_JIS）'))).toBe(true);
    expect(texts.some((t) => t.startsWith('04_適性検査.xls: xls形式のファイルは値（と表示形式）だけを読み込みました'))).toBe(true);
    expect(texts.filter((t) => /02_面接評価|05_研修出欠/.test(t)), '02 / 05 は注意なし').toEqual([]);

    // merge: the merge's own notes are added (the ones above stay)
    await mergeNow(page, errors);
    w = await warningState(page);
    test.info().annotations.push({ type: '結合後の警告', description: w.items.map((i) => `[${i.severity}] ${i.text}`).join('\n') });
    expect(w.count).toBe('8');
    expect(w.items.map((i) => i.text)).toEqual(expect.arrayContaining(texts));
    const added = w.items.map((i) => i.text).filter((t) => !texts.includes(t));
    expect(added.some((t) => /^取り込んだ数式 \d+ 件は数式のまま残しました/.test(t)), '取り込んだ数式を数式のまま残した').toBe(true);
    expect(added.some((t) => t.includes('フィルターの範囲を取り込んだ列まで広げました')), 'フィルターの範囲を広げた').toBe(true);
    expect(added.some((t) => t.includes('「04_適性検査」は値だけを読み込める形式のため、書式なし（標準）で取り込みました')), '04 は書式なしで取り込んだ').toBe(true);

    // back to the files: clearing 03 and 04 takes their notes away
    await page.getByTestId('remerge-button').click();
    // 変更がなければ確認なしで戻る（あれば確認が出る）: どちらかが出るまで待つ
    await expect(page.getByTestId('confirm-ok').or(page.getByTestId('merge-button'))).toBeVisible();
    if (await page.getByTestId('confirm-ok').isVisible()) await page.getByTestId('confirm-ok').click();
    await expect(page.getByTestId('merge-button')).toBeVisible();
    await waitIdle(page);
    await page.getByTestId('file-clear-3').click();
    await page.getByTestId('file-clear-2').click();
    await expect(page.locator('[data-testid^="file-card-"]')).toHaveCount(3);
    await waitMergeReady(page, errors);
    await waitIdle(page);
    w = await warningState(page);
    expect(w.items.map((i) => i.text), '03 と 04 の注意が消えて、01 の注意だけが残る').toEqual([`グラフ 1 個: ${hint}`, `ピボットテーブル 1 個: 元データ範囲は ${hint}（出力後に「更新」してください）`]);
    expect(w.count).toBe('2');
    expect(errors.messages.filter((m) => !/favicon|DevTools|Failed to load resource/i.test(m))).toEqual([]);
  });

  test('(d) 出力後の警告: 出力したときの注意（Excel本体が出す警告）は「出力: …」として加わる', async ({ page }) => {
    test.setTimeout(15 * 60_000);
    const errors = watchErrors(page);
    await mergeSamples(page, errors, TWO);
    const before = (await warningState(page)).items.map((i) => i.text);
    await runExport(page, errors, 'js', `e2e_r5d_${Date.now()}.xlsx`, OUT_DIR);
    const after = (await warningState(page)).items.map((i) => i.text);
    // the built-in engine has no warnings for this merge: the list is unchanged; anything added is an 「出力: …」 note
    expect(after.filter((t) => !before.includes(t)).every((t) => t.startsWith('出力: '))).toBe(true);
    expect(after).toEqual(expect.arrayContaining(before));
  });
});

// ---------------------------------------------------------------------------------------------------------------------------------
// (e) keyboard shortcuts
// ---------------------------------------------------------------------------------------------------------------------------------

async function redoCount(page: Page): Promise<number> {
  const m = /やり直し可能\s*(\d+)/.exec(await page.getByTestId('side-summary').innerText());
  return m ? Number(m[1]) : 0;
}

test('(e) Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z: 何度操作したあとでも 1 回のキーで 1 手ずつ、フォーカスがどこにあっても効く（入力中・ダイアログ中は効かない）', async ({ page }) => {
  test.setTimeout(15 * 60_000);
  const errors = watchErrors(page);
  await mergeSamples(page, errors, TWO);
  const initial = await columnTitles(page);
  const edited = await someEdits(page, errors); // 5 operations
  expect(await editCount(page)).toBe(5);
  const state = async () => ({ edits: await editCount(page), redo: await redoCount(page) });
  const expectState = async (edits: number, redo: number, what: string) => {
    await expect.poll(state, { message: what, timeout: 10_000 }).toEqual({ edits, redo });
  };

  // focus in the grid (a cell is selected)
  await page.getByTestId('grid-cell-5-0').click();
  await page.keyboard.press('Control+z');
  await expectState(4, 1, '(グリッド) Ctrl+Z 1 回で 1 手');
  await page.keyboard.press('Control+z');
  await page.keyboard.press('Control+z');
  await expectState(2, 3, '(グリッド) Ctrl+Z 3 回で 3 手');
  await page.keyboard.press('Control+y');
  await expectState(3, 2, '(グリッド) Ctrl+Y');
  await page.keyboard.press('Control+Shift+z');
  await expectState(4, 1, '(グリッド) Ctrl+Shift+Z もやり直し');
  await page.keyboard.press('Control+Shift+Z');
  await expectState(5, 0, '(グリッド) Ctrl+Shift+Z（大文字のキー）');
  expect(await columnTitles(page), '5 手やり直すと編集後の列に戻る').toEqual(edited);
  // nothing to redo: a no-op, not an error
  await page.keyboard.press('Control+y');
  await page.keyboard.press('Control+Shift+z');
  await expectState(5, 0, 'やり直すものがないときは何も起きない');

  // focus on a button / on plain page text (not in the grid): the window-level shortcut works too
  await page.getByTestId('side-summary').click();
  await page.keyboard.press('Control+z');
  await expectState(4, 1, '(ページ) Ctrl+Z');
  await page.getByTestId('redo-button').focus();
  await page.keyboard.press('Control+y');
  await expectState(5, 0, '(ボタンにフォーカス) Ctrl+Y');
  await page.getByTestId('column-list').click({ position: { x: 5, y: 5 } });
  await page.keyboard.press('Control+z');
  await expectState(4, 1, '(列の一覧) Ctrl+Z');
  await page.keyboard.press('Control+y');
  await expectState(5, 0, '(列の一覧) Ctrl+Y');

  // typing in a text field keeps Ctrl+Z for the text: the formula bar while a cell is being edited
  await page.getByTestId('grid-cell-6-0').click();
  const bar = page.getByTestId('formula-bar-input');
  await bar.click();
  await bar.fill('abc');
  await page.keyboard.press('Control+z');
  await expectState(5, 0, '(数式バーの入力中) Ctrl+Z は文字の取り消しで、結合後の変更は戻らない');
  await page.keyboard.press('Escape');
  await expectState(5, 0, 'Esc で入力をやめても変更は増えない');

  // …and inside a cell that is being edited (the editor is a text field of its own)
  const cell = page.getByTestId('grid-cell-7-2');
  const original = (await cell.innerText()).trim();
  await cell.dblclick();
  await expect(page.getByTestId('grid-editor'), 'セルの編集が始まる').toHaveClass(/is-editing/);
  await page.keyboard.type('xyz');
  await page.keyboard.press('Control+z');
  await expectState(5, 0, '(セルの編集中) Ctrl+Z は文字の取り消しで、結合後の変更は戻らない');
  await page.keyboard.press('Escape');
  await expect(page.getByTestId('grid-editor')).not.toHaveClass(/is-editing/);
  await expect(cell, 'Esc で編集をやめるとセルは元のまま').toHaveText(original);
  await expectState(5, 0, 'Esc で編集をやめても変更は増えない');

  // while a dialog is open the document is not touched
  await page.getByTestId('export-button').click();
  await expect(page.getByTestId('export-dialog')).toBeVisible();
  await page.keyboard.press('Control+z');
  await page.keyboard.press('Control+y');
  await expectState(5, 0, '(出力ダイアログ中) Ctrl+Z / Ctrl+Y は文書に効かない');
  await page.keyboard.press('Escape');
  await expect(page.getByTestId('export-dialog')).toBeHidden();

  // all the way back with the keyboard, one press each; the very first state, then one more press (no-op)
  for (let k = 4; k >= 0; k--) {
    await page.getByTestId('grid-cell-5-1').click();
    await page.keyboard.press('Control+z');
    await expectState(k, 5 - k, `Ctrl+Z → 変更 ${k} 件`);
  }
  expect(await columnTitles(page)).toEqual(initial);
  await page.keyboard.press('Control+z');
  await expectState(0, 5, '最初の状態でさらに Ctrl+Z しても何も起きない');
  await expect(page.getByTestId('undo-button')).toBeDisabled();

  // a new edit drops the redo branch: Ctrl+Y / Ctrl+Shift+Z do nothing afterwards
  await typeInCell(page, 4, await columnIndex(page, '電話'), '000-0000');
  await expectState(1, 0, '新しい編集でやり直しの枝は消える');
  await page.keyboard.press('Control+y');
  await page.keyboard.press('Control+Shift+z');
  await expectState(1, 0, '枝が消えたあとの Ctrl+Y / Ctrl+Shift+Z');
  await expect(page.getByTestId('redo-button')).toBeDisabled();
  expect(errors.messages.filter((m) => !/favicon|DevTools|Failed to load resource/i.test(m))).toEqual([]);
});
