// Round 3 flows on the samples (owner: T):
//   (g) a NON-contiguous set of columns selected with Ctrl and moved with one drag → both engines → check-output PASS
//       and the formulas of 集計 still refer to the columns with the same headers
//   (h) the files dropped in a different order, base = 02_面接評価 chosen with ★基準にする → 04_適性検査 has no key against 02 and
//       is EXCLUDED (badge 「結合対象外（結合キーが見つかりません）」 + 「キーを選ぶ…」, the merge is not blocked), the rest merges, the
//       summary numbers match the truth of the samples, the export passes check-output with --base 02_面接評価.xlsx
//   (i) 「キーを選ぶ…」 → the manual picker 「列を選んで結合キーに追加」: 04 joins the merge again (card, summary, columns, export)
//   (j) the manual picker rescues a file whose automatic key was switched off (フリガナ ⟷ 氏名カナ): the rows line up with truth.json
//   (k) no file has a key at all: 「結合する」 is disabled with the reason, and a manual key enables it
import fs from 'node:fs';
import path from 'node:path';
import { expect, test, type Locator, type Page } from '@playwright/test';
import { excelAvailable } from '../tools/lib/inspect.mjs';
import {
  columns,
  columnIndex,
  columnTitles,
  dragHeader,
  editCount,
  engineLabel,
  exportBoth,
  expectPass,
  expectSameFormulas,
  loadFiles,
  mergeNow,
  mergeSamples,
  openApp,
  requireTestIds,
  revealColumn,
  SAMPLE_FILES,
  SAMPLES_DIR,
  samplePaths,
  selectColumns,
  setSheetMode,
  step,
  toastLog,
  verifyAndProbe,
  waitMergeReady,
  watchErrors,
} from './helpers';
import { alignmentProblems, readSheet, verifyCases, waitIdle } from './round5';

/** base + 02_面接評価 only: quicker merges for the flows that do not need every file */
const TWO = [SAMPLE_FILES[0], SAMPLE_FILES[1]];

/** Titles after moving the columns `moved` (in their current order) in front of the column that is `gap` in `titles`. */
function afterMove(titles: string[], moved: string[], gap: number): string[] {
  const set = new Set(moved);
  const head = titles.slice(0, gap).filter((t) => !set.has(t));
  const rest = titles.slice(gap).filter((t) => !set.has(t));
  return [...head, ...titles.filter((t) => set.has(t)), ...rest];
}

/** Text of a grid cell (scrolls the virtualised grid to it). */
async function readCell(page: Page, row: number, colIdx: number): Promise<string> {
  await revealColumn(page, colIdx);
  const grid = page.getByTestId('grid');
  await grid.evaluate((el, y) => el.scrollTo({ top: y }), Math.max(0, (row - 8) * 22));
  const cell = page.getByTestId(`grid-cell-${row}-${colIdx}`);
  await expect(cell).toBeAttached();
  return (await cell.innerText()).trim();
}

const stripSpaces = (s: string) => s.replace(/\s+/g, '');

/** Text of every 警告 note (the section is collapsed unless there are errors: it is opened first). */
async function notesOf(page: Page): Promise<string[]> {
  await page.getByTestId('warning-details').evaluate((el) => ((el as HTMLDetailsElement).open = true));
  return (await page.locator('[data-testid="warning-list"] li').allInnerTexts()).map((t) => t.replace(/\s+/g, ' ').trim());
}

/** A file without a usable key: badge + 「キーを選ぶ…」 on ITS card only, and the notices in the summary / preview / 警告. */
async function expectExcluded(page: Page, index: number, fileName: string): Promise<void> {
  const stem = fileName.replace(/\.[^.]+$/, '');
  const badge = page.getByTestId(`file-excluded-${index}`);
  await expect(badge, `${fileName} のカードに「結合対象外（結合キーが見つかりません）」`).toContainText('結合対象外（結合キーが見つかりません）');
  await expect(page.getByTestId(`file-pick-key-${index}`)).toHaveText('キーを選ぶ…');
  await expect(page.locator('[data-testid^="file-excluded-"]'), '対象外の表示はそのファイルのカードだけ').toHaveCount(1);
  await expect(page.getByTestId('summary-line')).toContainText('結合対象外 1 ファイル');
  await expect(page.getByTestId('excluded-notice'), '結合プレビューにも出る').toContainText(stem);
  await expect(page.getByTestId('side-summary')).toContainText('結合対象外');
  const note = (await notesOf(page)).find((n) => n.includes(stem) && /結合の対象外/.test(n));
  expect(note, `警告に「${stem}」の結合対象外の注意がある`).toBeTruthy();
}

/** Selects the option of a <select> whose text starts with `text` ('受験番号' → '受験番号（A列）'). */
async function selectStartingWith(select: Locator, text: string): Promise<void> {
  const value = await select.locator('option').evaluateAll((os, t) => (os as HTMLOptionElement[]).find((o) => (o.textContent ?? '').startsWith(t))?.value, text);
  if (value == null) throw new Error(`選択肢「${text}…」がありません: ${(await select.locator('option').allInnerTexts()).join(' | ')}`);
  await select.selectOption(value);
}

/** 「列を選んで結合キーに追加」: base column(s) ⟷ this file's column(s), then 追加. Returns the message shown next to the button. */
async function addManualKey(page: Page, baseCols: string[], otherCols: string[]): Promise<string> {
  await page.getByTestId('manual-key-picker').scrollIntoViewIfNeeded();
  for (let k = 1; k < baseCols.length; k++) await page.getByTestId('manual-base-add-column').click();
  for (let k = 1; k < otherCols.length; k++) await page.getByTestId('manual-other-add-column').click();
  for (const [k, c] of baseCols.entries()) await selectStartingWith(page.getByTestId(`manual-base-select-${k}`), c);
  for (const [k, c] of otherCols.entries()) await selectStartingWith(page.getByTestId(`manual-other-select-${k}`), c);
  await expect(page.getByTestId('manual-key-hint')).toContainText('検出した種類');
  await expect(page.getByTestId('manual-key-add')).toBeEnabled();
  await page.getByTestId('manual-key-add').click();
  const msg = page.getByTestId('manual-key-message');
  await expect(msg).toBeVisible();
  await waitIdle(page);
  return (await msg.innerText()).trim();
}
/** formula text without the $ signs: Excel may write an absolute / relative reference differently after a move */
const noDollar = (f: string | null | undefined) => (f ?? '').replace(/\$/g, '');

test.describe('ラウンド3: 離れた列の一括移動・基準ファイルを変えた結合', () => {
  test('(g) Ctrl で離れた列（学年・年齢・判定）を選んで 1 回のドラッグで移動 → 両エンジンで PASS、集計の数式は同じ見出しの列を指す', async ({ page }) => {
    test.setTimeout(15 * 60_000);
    const errors = watchErrors(page);
    // the formulas of 集計 are checked below: keep the base file's other sheets (the app's default output is the merged sheet only)
    await mergeSamples(page, errors, TWO, () => setSheetMode(page, 'base'));
    const before = await columnTitles(page);
    // two base columns inside the title merge (学年 F, 年齢 J) and one imported column far to the right (判定)
    const moved = ['学年', '年齢', '判定'];
    for (const t of moved) expect(before, `列「${t}」`).toContain(t);
    expect(before.indexOf('判定') - before.indexOf('年齢'), '選ぶ列は互いに離れている').toBeGreaterThan(2);
    const gap = 2; // in front of フリガナ (the two frozen columns 学生ID / 氏名 stay)
    const expected = afterMove(before, moved, gap);
    expect(expected.slice(0, 6)).toEqual(['学生ID', '氏名', '学年', '年齢', '判定', 'フリガナ']);

    // Ctrl+click the three headers: exactly these are selected (neighbours are not)
    const idx = await selectColumns(page, moved);
    for (const i of idx) await expect(await revealColumn(page, i)).toHaveAttribute('aria-selected', 'true');
    for (const t of ['大学', '志望度']) await expect(await revealColumn(page, before.indexOf(t))).toHaveAttribute('aria-selected', 'false');

    // one drag moves the whole selection; 学年 / 年齢 lie inside the merged title A1:K1 → the app asks first
    await dragHeader(page, idx[idx.length - 1], gap);
    const dlg = page.getByTestId('unmerge-confirm-dialog');
    await step(page, errors, 'タイトル結合の確認', async () => {
      await expect(dlg, '学年・年齢はタイトルの結合 (A1:K1) の一部なので「結合を解除して移動」の確認が出る').toBeVisible();
    });
    await page.getByTestId('confirm-ok').click();
    await expect(dlg).toBeHidden();
    await step(page, errors, '離れた列の移動', async () => {
      await expect.poll(() => columnTitles(page), { timeout: 15_000 }).toEqual(expected);
    });
    expect(await editCount(page), '結合の解除＋移動は 1 回の操作').toBe(1);

    // undo / redo bring back both states
    await page.getByTestId('undo-button').click();
    await expect.poll(() => columnTitles(page)).toEqual(before);
    await page.getByTestId('redo-button').click();
    await expect.poll(() => columnTitles(page)).toEqual(expected);

    // the formulas of 集計 expected after the move (letters of the CURRENT order), compared without $
    const cols = await columns(page);
    const L = (title: string) => {
      const c = cols.find((x) => x.title === title);
      if (!c) throw new Error(`列「${title}」がありません`);
      return c.letter;
    };
    const uni = L('大学');
    const age = L('年齢');
    const wish = L('志望度');
    const grade = L('学年');
    const mail = L('メール');
    const dup = L('重複チェック');
    const name = L('氏名');
    const want: Record<string, string> = {
      '集計!B5': `=COUNTIF(名簿!${uni}:${uni},A5)`,
      '集計!C5': `=IFERROR(AVERAGEIF(名簿!${uni}4:${uni}63,A5,名簿!${age}4:${age}63),"-")`,
      '集計!D5': `=COUNTIFS(名簿!${uni}4:${uni}63,A5,名簿!${wish}4:${wish}63,">=4")`,
      '集計!E5': `=IFERROR(SUMPRODUCT((名簿!${uni}4:${uni}63=A5)*名簿!${wish}4:${wish}63)/B5,"-")`,
      '集計!B18': `=COUNTIF(名簿!${grade}:${grade},A18)`,
      '集計!B24': `=COUNTBLANK(名簿!${mail}4:${mail}63)`,
      '集計!B25': `=MAX(名簿!${age}:${age})`,
      '集計!B27': `=名簿!${name}4`,
      '集計!B29': `=INDEX(名簿!${name}4:${name}63,MATCH(MAX(名簿!${wish}4:${wish}63),名簿!${wish}4:${wish}63,0))`,
      '集計!B30': `=COUNTIF(名簿!${dup}4:${dup}63,">1")`,
    };
    const out = await exportBoth(page, errors, 'e2e_r3g_離れた列の移動');
    const v = await verifyAndProbe(out, { cells: Object.keys(want) });
    expectPass(v.verdicts, '離れた列の移動');
    expectSameFormulas(out);
    test.skip(!excelAvailable(), '集計の数式の確認には Excel が必要です');
    v.probes.slice(1).forEach((p, i) => {
      const label = engineLabel(out.files[i].engine);
      test.info().annotations.push({ type: `集計の数式 ${label}`, description: Object.keys(want).map((c) => `${c}: ${p[c]?.formula}`).join('\n') });
      for (const [cell, formula] of Object.entries(want)) expect.soft(noDollar(p[cell]?.formula), `${label}: ${cell} は同じ見出しの列を指す`).toBe(formula);
    });
    // the defined name 志望度列 follows 志望度
    v.dumps.slice(1).forEach((d, i) => {
      const n = (d.names ?? []).find((x) => x.name === '志望度列');
      expect.soft(noDollar(n?.refersTo), `${out.files[i].engine}: 名前「志望度列」は志望度の列`).toBe(`=名簿!${wish}:${wish}`);
    });
    for (const [engine, verdict] of v.verdicts) {
      expect(verdict.items.find((x) => x.code === 'formula-columns')?.level, `${engine}: 数式はすべて同じ見出しの列を参照`).toBe('pass');
      expect(verdict.stats.formulasChecked as number).toBeGreaterThan(100);
    }
  });

  test('(h) ファイルを別の順番で投入し ★基準にする で 02_面接評価 を基準に → 結合でき、数字が正しく、--base 02 で PASS', async ({ page }) => {
    test.setTimeout(20 * 60_000);
    const errors = watchErrors(page);
    // the samples in the REVERSE order: the first dropped file (05_研修出欠) is the default base
    const order = [...SAMPLE_FILES].reverse();
    const iBase = order.indexOf('02_面接評価.xlsx');
    const i04 = order.indexOf('04_適性検査.xls');
    const base02 = path.join(SAMPLES_DIR, '02_面接評価.xlsx');
    const truth = JSON.parse(fs.readFileSync(path.join(SAMPLES_DIR, 'truth.json'), 'utf8')) as {
      base: { students: Array<{ row: number; kana: string; uni: string; email: string }> };
      interview: { rows: Array<{ row: number; baseRow: number | null }>; expected: { unmatchedBaseRows: number[]; extraRows: number[] } };
    };
    const summary = page.getByTestId('summary-line');

    await openApp(page);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, samplePaths(order)));
    await waitMergeReady(page, errors);
    await requireTestIds(page, ['summary-line', 'check-panel'], '結合キーの見込み');
    await expect(page.getByTestId('base-toggle-0'), '最初に投入したファイルが基準').toHaveAttribute('aria-pressed', 'true');
    await expect(page.getByTestId('file-card-0')).toContainText('05_研修出欠.xlsx');
    await expect(summary).toContainText(/基準\s*32\s*名/);

    // ★基準にする on 02_面接評価: the base and every number follow
    await page.getByTestId(`base-toggle-${iBase}`).click();
    await expect(page.getByTestId(`base-toggle-${iBase}`)).toHaveAttribute('aria-pressed', 'true');
    await expect(page.getByTestId('base-toggle-0')).toHaveAttribute('aria-pressed', 'false');
    await expect(summary, '基準は 02_面接評価 の 63 行').toContainText(/基準\s*63\s*名/, { timeout: 60_000 });
    await expect(page.getByTestId(`file-keys-${iBase}`)).toContainText('受験番号');

    // numbers before the merge (truth.json: 63 rows in 02, 3 of them in no other file = 11 / 30 / 49 行目)
    const checkNumbers = async () => {
      await expect(summary, '基準は 02_面接評価 の 63 行').toContainText(/基準\s*63\s*名/, { timeout: 60_000 });
      await expect(summary, '未一致 3（どのファイルにもいない 3 行）').toContainText(/未一致\s*3(?!\d)/);
      await expect(page.getByTestId('check-unmatched-count')).toHaveText('3');
      await expect(page.getByTestId('check-panel')).toContainText(`${truth.interview.expected.extraRows[0]}行目`);
      await expect.soft(page.getByTestId('check-duplicate-count'), '再提出の重複（2 名分）などが数えられる').toHaveText(/[1-9]/);
    };
    await checkNumbers();

    // 02 has no フリガナ: 04_適性検査 (氏名カナ only) has no key against it. That file is EXCLUDED from the merge — it does not
    // block it: the card says 「結合対象外（結合キーが見つかりません）」 and offers 「キーを選ぶ…」, the summary / preview / 警告 say so
    // too, and 「結合する」 merges the other three.
    await expectExcluded(page, i04, order[i04]);
    await expect(page.getByTestId('merge-button'), '04 が対象外でも結合はできる').toBeEnabled();
    await checkNumbers();
    await mergeNow(page, errors);
    // the toast of the merge names the excluded file (toasts vanish: the recorded log is read)
    await expect.poll(async () => (await toastLog(page)).some((t) => t.includes('結合しました') && t.includes('結合対象外 1 ファイル（04_適性検査）')), { message: '結合後のトーストに対象外のファイル名が出る' }).toBe(true);
    await requireTestIds(page, ['column-list', 'undo-button', 'export-button', 'side-summary'], '結合後の画面');
    const titles = await columnTitles(page);
    test.info().annotations.push({ type: '結合後の列', description: titles.join(' / ') });
    // 02's own columns stay in front, in their own order; the other files' columns follow (header conflicts get the file name)
    expect(titles.slice(0, 9)).toEqual(['受験番号', '氏名', '面接官', '評価1', '評価2', '評価3', '平均', '判定', 'コメント']);
    expect(titles.length, '02 の 9 列 + 取り込み 18 列').toBe(27);
    for (const h of ['学籍番号', '氏名カナ', '受検日', '性格タイプ', '総合判定']) expect(titles, `対象外の 04_適性検査 の列「${h}」は取り込まない`).not.toContain(h);
    const notesAfter = await notesOf(page);
    expect(notesAfter.find((n) => n.includes('04_適性検査') && n.includes('取り込んでいません')), '警告（結合後）: 「列は取り込んでいません」').toBeTruthy();
    expect(titles.filter((x, i) => titles.indexOf(x) !== i), '見出しの重複').toEqual([]);
    for (const h of ['フリガナ', '出席回数', 'メール', '志望度']) expect(titles).toContain(h);
    expect.soft(titles.indexOf('出席回数'), '取り込み列は投入した順（05 → 03 → 01）に並ぶ').toBeLessThan(titles.indexOf('メール'));
    expect.soft(titles.indexOf('メール'), '取り込み列は投入した順（05 → 03 → 01）に並ぶ').toBeLessThan(titles.indexOf('フリガナ'));
    await expect(page.getByTestId('side-summary')).toContainText(/63\s*行/);

    // rows line up with the truth: 02 row r ↔ roster row baseRow (extras: nothing imported, resubmitted rows: same student)
    const iKana = await columnIndex(page, 'フリガナ');
    const iUni = await columnIndex(page, '大学（01_学生名簿）');
    expect(await readCell(page, 2, 0)).toBe('000158');
    for (const r of [2, 3, 11, 30, 41, 56, 64]) {
      const t = truth.interview.rows.find((x) => x.row === r)!;
      const stu = t.baseRow == null ? null : truth.base.students.find((s) => s.row === t.baseRow)!;
      expect(stripSpaces(await readCell(page, r, iKana)), `${r} 行目のフリガナ（名簿 ${t.baseRow ?? 'なし'} 行目）`).toBe(stu ? stripSpaces(stu.kana) : '');
      expect(await readCell(page, r, iUni), `${r} 行目の大学（名簿 ${t.baseRow ?? 'なし'} 行目）`).toBe(stu ? stu.uni : '');
    }

    const out = await exportBoth(page, errors, 'e2e_r3h_基準を02に');
    const v = await verifyAndProbe(out, { base: base02, check: { sheet: '面接評価', headerRow: 1 } });
    expectPass(v.verdicts, '基準 02_面接評価');
    for (const [engine, verdict] of v.verdicts) {
      expect(verdict.stats.mergedSheet, `${engine}: 結合した表は 面接評価 シート`).toBe('面接評価');
      const added = verdict.stats.addedHeaders as string[];
      expect(added.length, `${engine}: check-output が数えた追加列`).toBe(18);
      expect(added, `${engine}: 01 の大学は 見出し（ファイル名）`).toContain('大学（01_学生名簿）');
      const rows = verdict.stats.rows as { base: number; out: number };
      expect(rows.out, `${engine}: 行は基準のまま（追加行なし）`).toBe(rows.base);
      expect(verdict.items.some((i) => i.code === 'formula-columns' && i.level === 'pass'), `${engine}: 平均・判定の数式は同じ見出しの列を指す`).toBe(true);
    }
    expectSameFormulas(out);
  });

  test('(i) 「キーを選ぶ…」→ 列を選んで結合キーに追加（受験番号 ⟷ 学籍番号）→ 04 が結合に戻り、列が取り込まれ、出力も正しい', async ({ page }) => {
    test.setTimeout(15 * 60_000);
    const errors = watchErrors(page);
    // 02 is dropped first = the base; 01 joins by ID; 04 (氏名カナ only) has no key against 02
    const order = ['02_面接評価.xlsx', '01_学生名簿.xlsx', '04_適性検査.xls'];
    const base02 = path.join(SAMPLES_DIR, '02_面接評価.xlsx');
    await openApp(page);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, samplePaths(order)));
    await waitMergeReady(page, errors);
    await waitIdle(page);
    await expectExcluded(page, 2, order[2]);

    // 「キーを選ぶ…」 opens the key detail modal on 04's tab (source ids follow the drop order: s3)
    await page.getByTestId('file-pick-key-2').click();
    const modal = page.getByTestId('key-detail-modal');
    await expect(modal).toBeVisible();
    await expect(page.getByTestId('key-detail-tab-s3')).toHaveAttribute('aria-selected', 'true');
    await expect(page.getByTestId('key-detail-tab-s3')).toContainText('結合対象外');
    await expect(page.getByTestId('detail-excluded')).toBeVisible();
    await expect(page.getByTestId('match-rate-empty'), 'キーがないので一致率の文はまだ出ない').toBeVisible();
    await expect(page.getByTestId('match-rate-text')).toHaveCount(0);
    // the picker lists every column of both tables by header + column letter
    await page.getByTestId('manual-key-picker').scrollIntoViewIfNeeded();
    const baseOptions = await page.getByTestId('manual-base-select-0').locator('option').allInnerTexts();
    const otherOptions = await page.getByTestId('manual-other-select-0').locator('option').allInnerTexts();
    expect(baseOptions).toEqual(expect.arrayContaining(['受験番号（A列）', '氏名（B列）', 'コメント（I列）']));
    expect(otherOptions).toEqual(expect.arrayContaining(['学籍番号（A列）', '氏名カナ（B列）', '総合判定（H列）']));
    // 追加 stays disabled until a column is chosen on both sides
    await expect(page.getByTestId('manual-key-add')).toBeDisabled();
    await selectStartingWith(page.getByTestId('manual-base-select-0'), '受験番号');
    await expect(page.getByTestId('manual-key-add')).toBeDisabled();

    const added = await addManualKey(page, [], ['学籍番号']);
    expect(added, 'ピッカーが追加を知らせる').toContain('追加しました');
    const row = page.locator('[data-testid^="key-row-manual"]');
    await expect(row, '手動のキー候補が 1 行').toHaveCount(1);
    await expect(row).toContainText('手動');
    await expect(row.locator('input[type="checkbox"]')).toBeChecked();
    await expect(row.locator('[data-testid^="key-remove-"]'), '手動の候補には × がある').toHaveCount(1);
    await expect(page.getByTestId('detail-excluded'), '04 は結合に戻った').toHaveCount(0);
    await expect(page.getByTestId('key-detail-tab-s3')).not.toContainText('結合対象外');
    const rate = (await page.getByTestId('match-rate-text').innerText()).trim();
    test.info().annotations.push({ type: '一致率の文（受験番号 ⟷ 学籍番号）', description: rate });
    expect(rate).toMatch(/^このキーで \d+ \/ 42 行が一致します（基準 63 行中 \d+ 行が見つかりました）$/);

    // 使う のチェックを外す → 対象外に戻る、入れる → 戻る、× で候補ごと外す → 対象外に戻る
    const cb = row.locator('input[type="checkbox"]');
    await cb.uncheck();
    await waitIdle(page);
    await expect(page.getByTestId('detail-excluded')).toBeVisible();
    await expect(page.getByTestId('match-rate-empty')).toBeVisible();
    await cb.check();
    await waitIdle(page);
    await expect(page.getByTestId('detail-excluded')).toHaveCount(0);
    await expect(page.getByTestId('match-rate-text')).toBeVisible();
    await row.locator('[data-testid^="key-remove-"]').click();
    await waitIdle(page);
    await expect(row).toHaveCount(0);
    await expect(page.getByTestId('detail-excluded')).toBeVisible();
    // the picker starts again (one dropdown per side) and accepts the pair once more
    await expect(page.locator('[data-testid^="manual-base-select-"]')).toHaveCount(1);
    expect(await addManualKey(page, ['受験番号'], ['学籍番号'])).toContain('追加しました');
    await expect(row).toHaveCount(1);

    await page.getByTestId('key-detail-ok').click();
    await expect(modal).toBeHidden();
    await waitIdle(page);
    await expect(page.locator('[data-testid^="file-excluded-"]'), '対象外のカードはもうない').toHaveCount(0);
    await expect(page.getByTestId('file-keys-2'), '04 のカードにキーの列が出る').toContainText('学籍番号');
    await expect(page.getByTestId('summary-line')).not.toContainText('結合対象外');
    await expect(page.getByTestId('excluded-notice')).toHaveCount(0);
    expect((await notesOf(page)).some((n) => n.includes('結合の対象外')), '警告の「結合の対象外」の注意が消える').toBe(false);

    await mergeNow(page, errors);
    const titles = await columnTitles(page);
    test.info().annotations.push({ type: '結合後の列', description: titles.join(' / ') });
    expect(titles.slice(0, 9)).toEqual(['受験番号', '氏名', '面接官', '評価1', '評価2', '評価3', '平均', '判定', 'コメント']);
    for (const h of ['受検日', '言語', '非言語', '合計', '性格タイプ', '総合判定']) expect(titles, `04 の列「${h}」が取り込まれる`).toContain(h);
    for (const h of ['フリガナ', '志望度']) expect(titles, `01 の列「${h}」が取り込まれる`).toContain(h);
    expect(titles.filter((x, i) => titles.indexOf(x) !== i), '見出しの重複').toEqual([]);
    expect((await notesOf(page)).some((n) => n.includes('結合の対象外'))).toBe(false);

    // export (built-in engine) → Excel opens it without repair; the rows line up; 学籍番号 never matches 受験番号, so 04's columns stay empty
    const out = await exportBoth(page, errors, 'e2e_r3i_キーを選ぶ');
    const [v] = await verifyCases([{ label: '内蔵エンジン', base: base02, file: out.files[0].file, check: { sheet: '面接評価', headerRow: 1 } }]);
    if (v) expect(v.verdict, '出力の検証').toBe('PASS');
    const grid = readSheet(out.files[0].file, '面接評価');
    const { problems, checked } = alignmentProblems(grid, '02_面接評価.xlsx', ['01_学生名簿.xlsx']);
    expect(checked, '照合した行数').toBeGreaterThan(50);
    expect(problems, `01 の列が正しい行に入っている\n${problems.slice(0, 8).join('\n')}`).toEqual([]);
    const at = grid.rows[0].indexOf('性格タイプ');
    expect(at, '出力に 04 の「性格タイプ」列がある').toBeGreaterThanOrEqual(0);
    expect(grid.rows.slice(1, 64).map((r) => r[at]).filter((x) => x.trim() !== '').length, '学籍番号は受験番号と一致しないので 04 の列は空').toBe(0);
  });

  test('(j) 自動のキー（フリガナ ⟷ 氏名カナ）を外して対象外にした 04 を、列を選んで結合キーに追加して戻す → 行が正解（truth.json）どおりに並ぶ', async ({ page }) => {
    test.setTimeout(15 * 60_000);
    const errors = watchErrors(page);
    const order = ['01_学生名簿.xlsx', '04_適性検査.xls'];
    await openApp(page);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, samplePaths(order)));
    await waitMergeReady(page, errors);
    await waitIdle(page);
    await expect(page.getByTestId('file-excluded-1'), '最初は 04 にも自動のキーがある').toHaveCount(0);
    await expect(page.getByTestId('file-keys-1')).toContainText('氏名カナ');

    // switch every automatic key of 04 off → 04 is 結合対象外
    await page.getByTestId('key-detail-button').click();
    const modal = page.getByTestId('key-detail-modal');
    await expect(modal).toBeVisible();
    await page.getByTestId('key-detail-tab-s2').click();
    const used = page.locator('[data-testid="key-candidates"] tbody tr.is-used input[type="checkbox"]');
    const before = await used.count();
    expect(before, '04 で使われている自動のキーがある').toBeGreaterThan(0);
    for (let k = 0; k < before; k++) {
      await used.first().uncheck();
      await waitIdle(page);
    }
    await expect(page.getByTestId('detail-excluded')).toBeVisible();

    // …and choose フリガナ ⟷ 氏名カナ by hand: the same pair already sits in the table (unchecked) → it is only checked again
    const msg = await addManualKey(page, ['フリガナ'], ['氏名カナ']);
    test.info().annotations.push({ type: 'ピッカーのメッセージ', description: msg });
    expect(msg).toMatch(/追加しました|すでに表にある/);
    await expect(page.getByTestId('detail-excluded'), '04 は結合に戻った').toHaveCount(0);
    const rate = (await page.getByTestId('match-rate-text').innerText()).trim();
    test.info().annotations.push({ type: '一致率の文（フリガナ ⟷ 氏名カナ）', description: rate });
    const m = /^このキーで (\d+) \/ 42 行が一致します（基準 60 行中 (\d+) 行が見つかりました）$/.exec(rate);
    expect(m, rate).toBeTruthy();
    expect(Number(m![1]), '42 行のうち大半が一致する（truth.json: 40 行）').toBeGreaterThanOrEqual(35);
    await page.getByTestId('key-detail-ok').click();
    await expect(modal).toBeHidden();
    await waitIdle(page);
    await expect(page.locator('[data-testid^="file-excluded-"]')).toHaveCount(0);

    await mergeNow(page, errors);
    const out = await exportBoth(page, errors, 'e2e_r3j_手動キー');
    const grid = readSheet(out.files[0].file, '名簿');
    const { problems, checked } = alignmentProblems(grid, '01_学生名簿.xlsx', ['04_適性検査.xls']);
    expect(checked, '照合した行数（同姓同名の 2 行は 04 では見分けられないので除く）').toBe(58);
    expect(problems, `04 の性格タイプが正しい学生の行に入っている\n${problems.slice(0, 8).join('\n')}`).toEqual([]);
    const at = grid.rows[2].indexOf('性格タイプ');
    expect(grid.rows.slice(3, 63).filter((r) => r[at].trim() !== '').length, '性格タイプが入った行数').toBeGreaterThanOrEqual(35);
    await verifyCases([{ label: '内蔵エンジン', base: path.join(SAMPLES_DIR, '01_学生名簿.xlsx'), file: out.files[0].file }]);
  });

  test('(k) どのファイルにもキーがないと「結合する」は押せず理由が出る。列を選んでキーを足すと押せて、結合できる', async ({ page }) => {
    test.setTimeout(10 * 60_000);
    const errors = watchErrors(page);
    const order = ['02_面接評価.xlsx', '04_適性検査.xls'];
    await openApp(page);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, samplePaths(order)));
    await waitIdle(page);
    const merge = page.getByTestId('merge-button');
    await expect(merge, 'キーのあるファイルが一つもない').toBeDisabled();
    await expect(merge).toHaveAttribute('title', /どのファイルにも結合キーが選ばれていません/);
    await expect(page.getByTestId('summary-line')).toContainText('どのファイルにも結合キーが選ばれていません');
    await expect(page.getByTestId('side-summary')).toContainText('どのファイルにも結合キーが選ばれていません');
    await expect(page.getByTestId('file-excluded-1')).toContainText('結合対象外（結合キーが見つかりません）');

    await page.getByTestId('file-pick-key-1').click();
    await expect(page.getByTestId('key-detail-modal')).toBeVisible();
    expect(await addManualKey(page, ['受験番号'], ['学籍番号'])).toContain('追加しました');
    await page.getByTestId('key-detail-ok').click();
    await expect(page.getByTestId('key-detail-modal')).toBeHidden();
    await expect(merge, 'キーを足すと結合できる').toBeEnabled();
    await expect(page.getByTestId('file-excluded-1')).toHaveCount(0);
    await mergeNow(page, errors);
    const titles = await columnTitles(page);
    expect(titles.slice(0, 9)).toEqual(['受験番号', '氏名', '面接官', '評価1', '評価2', '評価3', '平均', '判定', 'コメント']);
    for (const h of ['受検日', '性格タイプ', '総合判定']) expect(titles).toContain(h);
  });
});
