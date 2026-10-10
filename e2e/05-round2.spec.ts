// Round 2 flows on the samples (owner: T): title-merge moves (「結合を解除して移動」), the key detail modal, join
// 'full', typed dates / % / dynamic arrays, deleting a referenced column, and a long undo/redo sequence — each
// exported with BOTH engines and verified with tools/check-output.mjs, tools/inspect (cell probes) and
// tools/compare.mjs (same op log, two engines: no formula differences).
import { expect, test } from '@playwright/test';
import { numToCol } from '../tools/lib/compare-core.mjs';
import { excelAvailable } from '../tools/lib/inspect.mjs';
import {
  columns,
  columnIndex,
  columnTitles,
  editCount,
  engineLabel,
  exportBoth,
  expectSameFormulas,
  insertColumn,
  loadFiles,
  mergeSamples,
  moveColumns,
  openApp,
  requireTestIds,
  SAMPLE_FILES,
  samplePaths,
  selectColumns,
  setSheetMode,
  step,
  typeInCell,
  verifyAndProbe,
  waitMergeReady,
  watchErrors,
} from './helpers';
import { waitIdle } from './round5';

const BASE_HEADERS = ['学生ID', '氏名', 'フリガナ', '大学', '学部', '学年', 'メール', '電話', '生年月日', '年齢', '志望度', '重複チェック'];
/** base + 02_面接評価 only: quicker merges for the flows that do not need every file */
const TWO = [SAMPLE_FILES[0], SAMPLE_FILES[1]];
const SHEET = '名簿';
const excelSerial = (y: number, m: number, d: number) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5;

test.describe('ラウンド2: タイトル結合・キー詳細・行の追加・入力形式・削除・元に戻す', () => {
  test('(a) タイトルの結合セルをまたぐ移動: 確認 →「結合を解除して移動」→ 移動し、出力してもタイトルが同じに見える', async ({ page }) => {
    test.setTimeout(15 * 60_000);
    const errors = watchErrors(page);
    await mergeSamples(page, errors, TWO);
    const before = await columnTitles(page);
    expect(before.slice(0, 12)).toEqual(BASE_HEADERS);
    const expected = ['学生ID', '氏名', 'フリガナ', '学年', '志望度', '大学', '学部', 'メール', '電話', '生年月日', '年齢', '重複チェック'];

    // 学年 (F) + 志望度 (K) lie inside the title merge A1:K1 → Excel refuses; the app offers to unmerge
    await moveColumns(page, ['学年', '志望度'], 3);
    const dlg = page.getByTestId('unmerge-confirm-dialog');
    await step(page, errors, 'タイトル結合の確認', async () => {
      await expect(dlg, '「結合を解除して移動」の確認ダイアログ').toBeVisible();
    });
    await expect.soft(dlg, '確認の文面に結合範囲 (A1:K1) が出ること').toContainText('A1:K1');
    await expect(page.getByTestId('confirm-ok')).toHaveText('結合を解除して移動');
    // cancel: nothing happens
    await page.getByTestId('confirm-cancel').click();
    await expect(dlg).toBeHidden();
    expect(await columnTitles(page)).toEqual(before);
    expect(await editCount(page)).toBe(0);

    await moveColumns(page, ['学年', '志望度'], 3);
    await expect(dlg).toBeVisible();
    await page.getByTestId('confirm-ok').click();
    await expect(dlg).toBeHidden();
    await step(page, errors, '結合を解除して移動', async () => {
      await expect.poll(async () => (await columnTitles(page)).slice(0, 12), { timeout: 15_000 }).toEqual(expected);
    });
    // unmerge + move = ONE edit / ONE undo step
    expect(await editCount(page)).toBe(1);
    await page.getByTestId('undo-button').click();
    await expect.poll(() => columnTitles(page)).toEqual(before);
    await page.getByTestId('redo-button').click();
    await expect.poll(async () => (await columnTitles(page)).slice(0, 12)).toEqual(expected);
    // the title is no longer merged: another move across it needs no confirmation
    await moveColumns(page, ['電話'], 1);
    await expect(dlg).toHaveCount(0);
    await expect.poll(async () => (await columnTitles(page))[1]).toBe('電話');
    await page.getByTestId('undo-button').click();
    await expect.poll(async () => (await columnTitles(page)).slice(0, 12)).toEqual(expected);

    const out = await exportBoth(page, errors, 'e2e_r2a_タイトル結合');
    const cells = ['A1', 'B1', 'D1', 'F1', 'K1', 'L1'].map((a) => `${SHEET}!${a}`);
    // ONE Excel instance: check-output verdicts of both outputs + the title cells of the base and both outputs
    const [b, ...outs] = (await verifyAndProbe(out, { cells })).probes;
    expectSameFormulas(out);
    test.skip(!excelAvailable(), 'セルの見た目の確認には Excel が必要です');
    const t = b[`${SHEET}!A1`];
    test.info().annotations.push({ type: '元のタイトル (A1)', description: JSON.stringify(t) });
    outs.forEach((p, i) => {
      const label = engineLabel(out.files[i].engine);
      test.info().annotations.push({ type: `出力のタイトル ${label}`, description: cells.map((c) => `${c}: ${JSON.stringify(p[c])}`).join('\n') });
      const a1 = p[`${SHEET}!A1`];
      expect(a1.text, `${label}: タイトルの文字は A1 のまま`).toBe(t.text);
      expect(a1.merge, `${label}: A1 の結合は解除されている`).toBeNull();
      expect([a1.bold, a1.fontSize, a1.fontColor, a1.fill], `${label}: A1 の書式（太字・サイズ・文字色・塗り）はそのまま`).toEqual([t.bold, t.fontSize, t.fontColor, t.fill]);
      for (const c of ['A1', 'B1', 'D1', 'F1', 'K1']) expect(p[`${SHEET}!${c}`].hAlign, `${label}: ${c} は「選択範囲内で中央」(7)`).toBe(7);
      expect(p[`${SHEET}!L1`].hAlign, `${label}: 元の結合範囲の外 (L1) は変わらない`).not.toBe(7);
      // while merged, Excel paints the top-left cell's fill over the whole area; unmerged cells show their own fill
      for (const c of ['B1', 'D1', 'F1', 'K1'])
        expect.soft(p[`${SHEET}!${c}`].fill, `${label}: ${c} にもタイトルの塗りつぶし（結合していたときの見た目）が必要（白い文字が白地に消えないこと）`).toBe(t.fill);
      // the band stops at the old merge area: the cell next to it keeps its own (base) fill
      expect.soft(p[`${SHEET}!L1`].fill, `${label}: 元の結合範囲の外 (L1) の塗りは元のまま`).toBe(b[`${SHEET}!L1`].fill);
    });
    // both engines replay the same op log: the title cells must look the same in the two outputs
    if (outs.length === 2) {
      for (const c of cells) {
        const [js, xl] = outs.map((p) => p[c]);
        expect.soft([js.fill, js.hAlign, js.bold, js.fontColor, js.merge ?? null], `${c}: 内蔵エンジンと Excel本体でタイトルの見た目（塗り・配置・太字・文字色・結合）が同じ`).toEqual([xl.fill, xl.hAlign, xl.bold, xl.fontColor, xl.merge ?? null]);
      }
    }
  });

  test('(b) 詳細設定で 02_面接評価 の ID キーを外すと、見込みの数字が変わる', async ({ page }) => {
    const errors = watchErrors(page);
    await openApp(page);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, samplePaths()));
    await waitMergeReady(page, errors);
    await requireTestIds(page, ['summary-line', 'key-detail-button', 'key-chip-id'], '結合キーの設定');
    const summary = page.getByTestId('summary-line');
    await expect(summary).toContainText(/基準\s*60\s*名/);
    const before = (await summary.innerText()).trim();

    await page.getByTestId('key-detail-button').click();
    const modal = page.getByTestId('key-detail-modal');
    await expect(modal).toBeVisible();
    await modal.locator('[data-testid^="key-detail-tab-"]').filter({ hasText: '02_面接評価' }).click();
    const detail = modal.locator('.detail-summary');
    await expect(detail).toContainText('一致');
    const detailBefore = (await detail.innerText()).trim();
    // the ID candidate in use: base 学生ID ⟷ 受験番号
    const used = modal.locator('[data-testid="key-candidates"] tbody tr.is-used').filter({ hasText: '受験番号' });
    await expect(used.first(), '受験番号の ID キーが使われていること').toBeVisible();
    const label = (await used.first().locator('.c-key b').innerText()).trim();
    const esc = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const row = modal
      .locator('[data-testid="key-candidates"] tbody tr')
      .filter({ has: page.locator('.c-key b', { hasText: new RegExp(`^${esc}$`) }) })
      .filter({ hasText: '受験番号' })
      .first();
    const cb = row.locator('input[type="checkbox"]');
    await expect(cb).toBeChecked();
    await cb.uncheck();
    await expect(cb).not.toBeChecked();

    await expect.poll(async () => (await detail.innerText()).trim(), { message: 'このファイルの一致件数が変わること', timeout: 30_000 }).not.toBe(detailBefore);
    await expect.poll(async () => (await summary.innerText()).trim(), { message: '見込み（summary-line）の数字が変わること', timeout: 30_000 }).not.toBe(before);
    const after = (await summary.innerText()).trim();
    test.info().annotations.push({ type: '見込み', description: `前: ${before}\n後: ${after}\n02 前: ${detailBefore}\n02 後: ${(await detail.innerText()).trim()}` });
    expect(after).toMatch(/基準\s*60\s*名/);
    await page.getByTestId('key-detail-ok').click();
    await expect(modal).toBeHidden();
    // 02 is the only other file with an ID column, so the ID chip is now off everywhere
    await expect(page.getByTestId('key-chip-id')).toHaveAttribute('aria-pressed', 'false');
  });

  test('(c) 「他ファイルにしかない行も末尾に追加」: 追加行が表の末尾に入り、集計の範囲も追加行まで広がる', async ({ page }) => {
    test.setTimeout(15 * 60_000);
    const errors = watchErrors(page);
    let expectedRows = 0;
    const summary = page.getByTestId('summary-line');
    const appendedIn = async () => Number(/追加\s*(\d+)/.exec(await summary.innerText())?.[1] ?? NaN);
    await mergeSamples(page, errors, SAMPLE_FILES, async () => {
      await requireTestIds(page, ['join-type-select'], '結合方法');
      await page.getByTestId('join-type-select').selectOption('full');
      // changing the join type drops the analysis: the preview is recomputed and 「結合する」 is enabled again once the number is final
      await waitMergeReady(page, errors);
      await waitIdle(page);
      await expect(summary).toContainText(/追加\s*\d+/);
      expectedRows = await appendedIn();
    });
    expect(expectedRows, '他ファイルにしかない行がある（02 の未登録 3 名など）').toBeGreaterThan(0);
    // after the merge: the side summary 「追加行 N」 = the rows really appended to the merged sheet
    const side = await page.getByTestId('side-summary').innerText();
    const appended = Number(/追加行\s*(\d+)/.exec(side)?.[1] ?? NaN);
    test.info().annotations.push({ type: '追加行', description: `結合前の見込み: 追加 ${expectedRows} / 結合後: ${side.replace(/\s+/g, ' ')}` });
    expect.soft(appended, `結合前の見込み（追加 ${expectedRows}）と実際に追加された行数（${side.replace(/\s+/g, ' ')}）が同じ`).toBe(expectedRows);
    expect(appended).toBeGreaterThan(0);
    // appended rows sit right after the base's last data row (63); the footer moves down by the same amount
    const last = 63 + appended;
    const grid = page.getByTestId('grid');
    const scrollToRow = async (r: number) => {
      await grid.evaluate((el, y) => el.scrollTo({ top: y }), Math.max(0, (r - 8) * 22));
      await expect(page.getByTestId(`grid-cell-${r}-0`)).toBeAttached();
    };
    for (let r = 64; r <= last; r++) {
      await scrollToRow(r);
      const texts = await Promise.all([0, 1, 2, 3].map((i) => page.getByTestId(`grid-cell-${r}-${i}`).innerText()));
      expect(texts.join('').trim(), `追加した ${r} 行目にキー（学生ID / 氏名 / フリガナ / 大学）が入っている`).not.toBe('');
    }
    // the grid shows the data rows only (the footer 人数 / 平均 below the blank row is not a grid row)
    await expect(page.getByTestId(`grid-cell-${last + 1}-0`), '追加行のあとは表示しない（集計行はグリッドの外）').toHaveCount(0);

    const out = await exportBoth(page, errors, 'e2e_r2c_行の追加');
    // the footer (人数 / 志望度4以上, rows 65-66 of the base) moved below the appended rows and its ranges grew
    const f = [`${SHEET}!A${last + 2}`, `${SHEET}!B${last + 2}`, `${SHEET}!B${last + 3}`];
    const { verdicts, probes: allProbes } = await verifyAndProbe(out, { cells: f });
    if (excelAvailable()) {
      const probes = allProbes.slice(1);
      probes.forEach((p, i) => {
        const label = engineLabel(out.files[i].engine);
        test.info().annotations.push({ type: `集計行 ${label}`, description: f.map((c) => `${c}: ${JSON.stringify(p[c])}`).join('\n') });
        expect(p[f[0]].value, `${label}: 集計行（人数）は追加行の下 (${last + 2} 行目)`).toBe('人数');
        expect(p[f[1]].formula, `${label}: 人数の COUNTA が追加行まで広がる`).toBe(`=COUNTA(B4:B${last})`);
        expect(p[f[2]].formula, `${label}: 志望度4以上の COUNTIF も広がる`).toMatch(new RegExp(`^=COUNTIF\\([A-Z]+4:[A-Z]+${last},`));
      });
    }
    for (const [engine, v] of verdicts) {
      const st = v.stats.appended as { rows: number; ranges: number; extended: number; notExtended: number } | undefined;
      const info = v.items.filter((i) => i.code.startsWith('appended')).map((i) => `[${i.level}] ${i.message}\n    ${(i.examples ?? []).join('\n    ')}`);
      test.info().annotations.push({ type: `追加行 ${engine}`, description: `${JSON.stringify(st)}\n${info.join('\n')}` });
      expect(st?.rows, `${engine}: check-output が数えた追加行`).toBe(appended);
      expect(st?.extended ?? 0, `${engine}: 集計（COUNTIF / AVERAGEIF … の範囲）が追加行まで広がる`).toBeGreaterThan(0);
      expect.soft(st?.notExtended, `${engine}: 追加行を含まない範囲\n${info.join('\n')}`).toBe(0);
    }
    expectSameFormulas(out);
  });

  test('(d) 日付・%・=UNIQUE(...) を入力 → 両エンジンで出力 → Excel で日付・%・スピルになる', async ({ page }) => {
    test.setTimeout(15 * 60_000);
    const errors = watchErrors(page);
    await mergeSamples(page, errors, TWO);
    const n = (await columnTitles(page)).length;
    await insertColumn(page, n - 1, 'right');
    await expect.poll(async () => (await columnTitles(page)).length).toBe(n + 1);
    const x = n; // the new, empty column at the end
    const letter = numToCol(x + 1);
    const c4 = await typeInCell(page, 4, x, '2026/4/1');
    await step(page, errors, '日付の入力', async () => {
      await expect(c4).toHaveText('2026/4/1');
    });
    const c5 = await typeInCell(page, 5, x, '50%');
    await expect(c5).toHaveText('50%');
    const c8 = await typeInCell(page, 8, x, '=UNIQUE(D4:D63)');
    await c8.click();
    await expect(page.getByTestId('formula-bar-input')).toHaveValue('=UNIQUE(D4:D63)');
    await c4.click();
    await expect(page.getByTestId('formula-bar-input')).toHaveValue(/^2026\/0?4\/0?1$/);

    const out = await exportBoth(page, errors, 'e2e_r2d_入力形式');
    const cells = [4, 5, 8, 9].map((r) => `${SHEET}!${letter}${r}`);
    const probes = (await verifyAndProbe(out, { cells })).probes.slice(1);
    expectSameFormulas(out);
    test.skip(!excelAvailable(), 'セルの表示形式の確認には Excel が必要です');
    probes.forEach((p, i) => {
      const label = engineLabel(out.files[i].engine);
      test.info().annotations.push({ type: `入力したセル ${label}`, description: cells.map((c) => `${c}: ${JSON.stringify(p[c])}`).join('\n') });
      const [d, pc, u, u2] = cells.map((c) => p[c]);
      expect.soft(d.value, `${label}: ${letter}4 は日付のシリアル値`).toBe(excelSerial(2026, 4, 1));
      expect.soft(d.numFmt, `${label}: ${letter}4 は日付の表示形式`).toMatch(/y+.*m+.*d+/i);
      expect.soft(d.text, `${label}: ${letter}4 の表示`).toBe('2026/4/1');
      expect.soft(pc.value, `${label}: ${letter}5 は 0.5`).toBe(0.5);
      expect.soft(pc.numFmt, `${label}: ${letter}5 は % の表示形式`).toBe('0%');
      expect.soft(pc.text, `${label}: ${letter}5 の表示`).toBe('50%');
      expect.soft(u.formula, `${label}: ${letter}8 の数式`).toBe('=UNIQUE(D4:D63)');
      expect.soft(u.spill ?? '', `${label}: ${letter}8 はスピルする（動的配列）`).toMatch(new RegExp(`^${letter}8:${letter}\\d+$`));
      expect.soft(u.error, `${label}: ${letter}8 はエラーにならない`).toBeUndefined();
      expect.soft(u2.value ?? null, `${label}: ${letter}9 にスピルした値`).not.toBeNull();
    });
  });

  test('(e) 数式が参照している列（年齢）を値に変換して削除 → どちらのエンジンでも #REF! がない', async ({ page }) => {
    test.setTimeout(15 * 60_000);
    const errors = watchErrors(page);
    // 集計 is a sheet of the base file: keep it (the app's default output is the merged sheet only)
    await mergeSamples(page, errors, TWO, () => setSheetMode(page, 'base'));
    const before = await columnTitles(page);
    await selectColumns(page, ['年齢']);
    await page.keyboard.press('Delete');
    const dlg = page.getByTestId('delete-confirm-dialog');
    await expect(dlg, '集計シートの AVERAGEIF(名簿!$J$4:$J$63 …) が年齢を参照しているので確認が出る').toBeVisible();
    await expect(dlg).toContainText('参照');
    await page.getByTestId('delete-convert-button').click();
    await expect(dlg).toBeHidden();
    await step(page, errors, '列の削除', async () => {
      await expect.poll(async () => (await columnTitles(page)).includes('年齢')).toBe(false);
    });
    expect((await columnTitles(page)).length).toBe(before.length - 1);

    const out = await exportBoth(page, errors, 'e2e_r2e_参照列の削除');
    const { verdicts, probes: e2probes } = await verifyAndProbe(out, { cells: ['集計!C5', '集計!B5'] });
    for (const [engine, v] of verdicts) {
      const ref = v.stats.refText as { base: number; out: number };
      expect(ref.out, `${engine}: #REF! を含む数式・名前・書式が増えていない`).toBeLessThanOrEqual(ref.base);
      const errs = v.stats.errors as { base: Record<string, number>; out: Record<string, number> };
      expect(errs.out['#REF!'] ?? 0, `${engine}: #REF! の値が増えていない`).toBeLessThanOrEqual(errs.base['#REF!'] ?? 0);
      expect(v.items.some((i) => i.code === 'formulas-to-values'), `${engine}: 年齢を参照していた数式は値に変換されている`).toBe(true);
    }
    expectSameFormulas(out);
    test.skip(!excelAvailable(), '値の確認には Excel が必要です');
    // 集計!C5 = AVERAGEIF(…年齢…) keeps its value (now a constant)
    const [b, ...outs] = e2probes;
    outs.forEach((p, i) => {
      const label = engineLabel(out.files[i].engine);
      expect(p['集計!C5'].formula, `${label}: 集計!C5 は値`).toBeNull();
      expect(p['集計!C5'].value, `${label}: 集計!C5 の値は削除前と同じ`).toBe(b['集計!C5'].value);
      expect(p['集計!B5'].formula, `${label}: 年齢を参照しない集計!B5 は数式のまま`).toBe(b['集計!B5'].formula);
    });
  });

  test('(f) 10 回以上の操作と元に戻す / やり直す → 両エンジンで出力 → 数式の差がない', async ({ page }) => {
    test.setTimeout(20 * 60_000);
    const errors = watchErrors(page);
    await mergeSamples(page, errors, TWO);
    const undo = async () => page.getByTestId('undo-button').click();
    const redo = async () => page.getByTestId('redo-button').click();
    const waitEdits = async (n: number, what: string) => {
      await step(page, errors, what, async () => {
        await expect.poll(() => editCount(page), { message: `${what}: 結合後の変更 ${n} 件`, timeout: 15_000 }).toBe(n);
      });
    };

    // 1. across the merged title: confirm 「結合を解除して移動」
    await moveColumns(page, ['メール'], 2);
    await expect(page.getByTestId('unmerge-confirm-dialog')).toBeVisible();
    await page.getByTestId('confirm-ok').click();
    await waitEdits(1, 'メールを移動（結合を解除）');
    // 2. an imported column to the very front
    await moveColumns(page, ['判定'], 0);
    await waitEdits(2, '判定を先頭へ');
    // 3. two non-adjacent imported columns to the end
    await moveColumns(page, ['評価1', '評価3'], (await columnTitles(page)).length);
    await waitEdits(3, '評価1・評価3 を末尾へ');
    // 4. insert a column right of 志望度 (a visible column: 重複チェック is a hidden helper column whose header is a
    //    few pixels wide, nobody right-clicks it)
    await insertColumn(page, await columnIndex(page, '志望度'), 'right');
    await waitEdits(4, '列の挿入');
    const newIdx = (await columnIndex(page, '志望度')) + 1;
    const letterOf = async (title: string) => (await columns(page)).find((c) => c.title === title)!.letter;
    // 5. a formula referring to 志望度 (current letter), 6. a number
    await typeInCell(page, 4, newIdx, `=${await letterOf('志望度')}4*2`);
    await waitEdits(5, '数式の入力');
    await typeInCell(page, 5, newIdx, '100');
    await waitEdits(6, '数値の入力');
    // 7. undo ×2, 8. redo ×1
    await undo();
    await undo();
    await waitEdits(4, '元に戻す ×2');
    await redo();
    await waitEdits(5, 'やり直す');
    // 9. a new edit drops the redo branch
    await typeInCell(page, 6, newIdx, 'メモ');
    await waitEdits(6, '文字の入力');
    await expect(page.getByTestId('redo-button')).toBeDisabled();
    // 10. delete 平均 (判定 refers to it) with conversion, undo, redo
    await selectColumns(page, ['平均']);
    await page.keyboard.press('Delete');
    await expect(page.getByTestId('delete-confirm-dialog')).toBeVisible();
    await page.getByTestId('delete-convert-button').click();
    await waitEdits(7, '平均を値に変換して削除');
    await undo();
    await waitEdits(6, '削除を元に戻す');
    expect(await columnTitles(page)).toContain('平均');
    await redo();
    await waitEdits(7, '削除をやり直す');
    expect(await columnTitles(page)).not.toContain('平均');
    // 11. move 面接官 into the (unmerged) title area, 12. edit 志望度, 13. insert + undo
    await moveColumns(page, ['面接官'], 3);
    await waitEdits(8, '面接官を移動');
    const shibo = await columnIndex(page, '志望度');
    await typeInCell(page, 7, shibo, '5');
    await waitEdits(9, '志望度の編集');
    await insertColumn(page, await columnIndex(page, '大学'), 'left');
    await waitEdits(10, '列の挿入（大学の左）');
    await undo();
    await waitEdits(9, '挿入を元に戻す');
    await expect(page.getByTestId('redo-button')).toBeEnabled();
    const finalTitles = await columnTitles(page);
    test.info().annotations.push({ type: '最終の列', description: finalTitles.join(' / ') });
    expect(finalTitles[0]).toBe('判定');
    expect(finalTitles.slice(-2)).toEqual(['評価1', '評価3']);

    const out = await exportBoth(page, errors, 'e2e_r2f_元に戻す');
    await verifyAndProbe(out);
    // both views: Excel (Formula2 after opening, recalculated) and the package (file formulas)
    expectSameFormulas(out);
  });
});

