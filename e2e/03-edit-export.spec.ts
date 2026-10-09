// The main flow on the samples: merge → select (Ctrl/Shift) → drag columns → undo/redo → delete a referenced
// column → edit a cell → export (内蔵エンジン / Excel本体) → tools/check-output.mjs verdict (owner: T).
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { expect, test, type Page } from '@playwright/test';
import { excelAvailable } from '../tools/lib/inspect.mjs';
import {
  checkOutput,
  columns,
  columnTitles,
  describeVerdict,
  dragHeader,
  excelInstalled,
  inspectProbe,
  loadFiles,
  mergeNow,
  MUST_NOT_FAIL,
  openApp,
  requireTestIds,
  revealColumn,
  ROOT,
  runExport,
  samplePaths,
  step,
  watchErrors,
  type CheckOutputResult,
} from './helpers';

const BASE_HEADERS = ['学生ID', '氏名', 'フリガナ', '大学', '学部', '学年', 'メール', '電話', '生年月日', '年齢', '志望度', '重複チェック'];
const OUT_DIR = path.join(ROOT, 'test-results', 'e2e-downloads');

test.describe.serial('結合 → 列の編集 → 出力', () => {
  let page: Page;
  let errors: { messages: string[] };
  let moved: string[] = [];
  let telLetter = '';
  let merged = false;
  const needMerged = () => test.fixme(!merged, '前のテストで結合まで進めなかったため実行できません');
  const outputs: Partial<Record<'js' | 'excel', string>> = {};

  test.beforeAll(async ({ browser }) => {
    page = await browser.newPage({ acceptDownloads: true, viewport: { width: 1440, height: 900 }, locale: 'ja-JP' });
    errors = watchErrors(page);
  });
  test.afterAll(async () => {
    await page?.close();
  });

  test('結合すると、基準の列の後ろに他ファイルの列が並ぶ', async () => {
    await openApp(page);
    await loadFiles(page, samplePaths());
    await mergeNow(page, errors);
    await requireTestIds(page, ['column-list', 'undo-button', 'redo-button', 'export-button'], '結合後の画面');
    const titles = await columnTitles(page);
    expect(titles.slice(0, 12)).toEqual(BASE_HEADERS);
    expect(titles.length).toBeGreaterThan(20);
    for (const h of ['面接官', '評価1', '平均', '判定', 'コメント']) expect(titles).toContain(h);
    // an imported header equal to an existing one gets the file name appended -> headers stay unique
    expect.soft(titles.filter((t, i) => titles.indexOf(t) !== i), '見出しの重複').toEqual([]);
    // the ID keeps its leading zeros, names are shown as in the base file
    await expect(page.getByTestId('grid-cell-4-0')).toHaveText('000100');
    await expect(page.getByTestId('grid-cell-4-1')).toHaveText('石川　莉子');
    await expect(page.getByTestId('remerge-button')).toBeVisible();
    merged = true;
  });

  test('Ctrl / Shift で離れた列を選んでドラッグで先頭へ移動できる', async () => {
    needMerged();
    const before = await columnTitles(page);
    const h12 = await revealColumn(page, 12);
    await h12.click();
    await expect(h12).toHaveAttribute('aria-selected', 'true');
    await (await revealColumn(page, 14)).click({ modifiers: ['Control'] });
    await (await revealColumn(page, 16)).click({ modifiers: ['Control', 'Shift'] });
    for (const i of [12, 14, 15, 16]) await expect(page.getByTestId(`grid-header-${i}`)).toHaveAttribute('aria-selected', 'true');
    await expect(page.getByTestId('grid-header-13')).toHaveAttribute('aria-selected', 'false');
    moved = [before[12], before[14], before[15], before[16]];

    await dragHeader(page, 14, 0);
    await step(page, errors, '列の移動', async () => {
      await expect.poll(async () => (await columnTitles(page)).slice(0, 4), { timeout: 15_000 }).toEqual(moved);
    });
    const after = await columnTitles(page);
    expect(after.slice(4, 16)).toEqual(BASE_HEADERS);
    expect(after.length).toBe(before.length);
    expect([...after].sort()).toEqual([...before].sort());
  });

  test('元に戻す / やり直す（キーボードとボタン）', async () => {
    needMerged();
    await page.getByTestId('grid-cell-5-0').click();
    await page.keyboard.press('Control+z');
    await expect.poll(async () => (await columnTitles(page))[0]).toBe('学生ID');
    await page.keyboard.press('Control+y');
    await expect.poll(async () => (await columnTitles(page)).slice(0, 4)).toEqual(moved);
    await page.keyboard.press('Control+z');
    await expect.poll(async () => (await columnTitles(page))[0]).toBe('学生ID');
    await page.keyboard.press('Control+Shift+z');
    await expect.poll(async () => (await columnTitles(page)).slice(0, 4)).toEqual(moved);

    await page.getByTestId('undo-button').click();
    await expect.poll(async () => (await columnTitles(page))[0]).toBe('学生ID');
    await expect(page.getByTestId('redo-button')).toBeEnabled();
    await page.getByTestId('redo-button').click();
    await expect.poll(async () => (await columnTitles(page)).slice(0, 4)).toEqual(moved);
  });

  test('タイトルの結合セルの一部だけを動かす移動は確認が出て、キャンセルすれば何も変わらない', async () => {
    needMerged();
    const before = await columnTitles(page);
    const from = before.indexOf('志望度');
    expect(from).toBeGreaterThan(4);
    const h = await revealColumn(page, from);
    await h.click();
    await dragHeader(page, from, 6);
    // Excel refuses to split a merged area; the only obstacle is the title above the header row, so the app asks
    // whether to unmerge it (「選択範囲内で中央」) and move — see 05-round2 (a) for the confirmed path
    const dlg = page.getByTestId('unmerge-confirm-dialog');
    await step(page, errors, 'タイトル結合の確認', async () => {
      await expect(dlg).toBeVisible();
    });
    await expect(page.getByTestId('confirm-ok')).toHaveText('結合を解除して移動');
    await page.getByTestId('confirm-cancel').click();
    await expect(dlg).toBeHidden();
    expect(await columnTitles(page)).toEqual(before);
  });

  test('数式が参照している列を削除すると確認が出て、値に変換して削除できる', async () => {
    needMerged();
    const before = await columnTitles(page);
    const idx = before.indexOf('平均');
    expect(idx, '取り込んだ「平均」列があること').toBeGreaterThanOrEqual(0);
    const h = await revealColumn(page, idx);
    await h.click();
    await expect(h).toHaveAttribute('aria-selected', 'true');
    await page.keyboard.press('Delete');
    const dlg = page.getByTestId('delete-confirm-dialog');
    // 判定 (=IF(平均>=3.5,…)) is imported as a live formula (premerge rule 6), so a confirmation is required
    await expect(dlg, '参照されている列の削除では確認ダイアログが出ること（取り込んだ「判定」の数式が「平均」を参照）').toBeVisible();
    // 判定 (=IF(平均>=3.5,…)) depends on 平均
    await expect.soft(dlg).toContainText('参照');
    await expect.soft(page.getByTestId('delete-convert-button')).toContainText('値に変換');

    // cancel first: nothing happens
    await page.getByTestId('delete-cancel-button').click();
    await expect(dlg).toBeHidden();
    expect(await columnTitles(page)).toEqual(before);

    await h.click();
    await page.keyboard.press('Delete');
    await expect(dlg).toBeVisible();
    await page.getByTestId('delete-convert-button').click();
    await expect(dlg).toBeHidden();
    await step(page, errors, '列の削除', async () => {
      await expect.poll(async () => (await columnTitles(page)).includes('平均')).toBe(false);
    });
    const after = await columnTitles(page);
    expect(after).toContain('判定');
    expect(after.length).toBe(before.length - 1);
    // the 判定 cells still show their results (now values)
    const j = after.indexOf('判定');
    await revealColumn(page, j);
    await expect(page.getByTestId(`grid-cell-4-${j}`)).toHaveText(/合格|保留|^$/);
  });

  test('セルを編集できる（入力 → Enter、数式バーにも表示）', async () => {
    needMerged();
    const cols = await columns(page);
    const tel = cols.find((c) => c.title === '電話');
    expect(tel, '基準の「電話」列').toBeTruthy();
    telLetter = tel!.letter;
    await revealColumn(page, tel!.index);
    const cell = page.getByTestId(`grid-cell-4-${tel!.index}`);
    await cell.click();
    await page.keyboard.type('03-1234-5678');
    await page.keyboard.press('Enter');
    await step(page, errors, 'セルの編集', async () => {
      await expect(cell).toHaveText('03-1234-5678');
    });
    await cell.click();
    await requireTestIds(page, ['formula-bar-input'], '数式バー');
    await expect(page.getByTestId('formula-bar-input')).toHaveValue('03-1234-5678');
  });

  test('内蔵エンジンで出力 → 修復なしで開け、#REF! がなく、集計の数式が同じ見出しの列を指す', async () => {
    test.setTimeout(6 * 60_000);
    needMerged();
    const file = await runExport(page, errors, 'js', 'e2e_内蔵エンジン.xlsx', OUT_DIR);
    outputs.js = file;
    expect(fs.statSync(file).size).toBeGreaterThan(10_000);
    await verify(file);
  });

  test('Excel本体で出力 → 同じチェックに合格し、内蔵エンジンと同じ数式になる', async () => {
    test.setTimeout(10 * 60_000); // Excel engine export + check-output + compare + probe = several Excel starts
    needMerged();
    test.skip(!(await excelInstalled(page)), 'サーバーモードで Excel がインストールされている場合のみ');
    const file = await runExport(page, errors, 'excel', 'e2e_Excel本体.xlsx', OUT_DIR);
    outputs.excel = file;
    await verify(file);
    if (outputs.js) {
      // same log, same semantics: formulas of both engines must agree
      const r = spawnSync(process.execPath, [path.join(ROOT, 'tools', 'compare.mjs'), outputs.js, file, '--no-values'], { cwd: ROOT, encoding: 'utf8', timeout: 5 * 60 * 1000 });
      test.info().annotations.push({ type: 'compare js vs excel', description: (r.stdout || r.stderr || '').slice(0, 4000) });
      expect.soft(r.status, `内蔵エンジンと Excel本体の出力の差分:\n${r.stdout}`).toBe(0);
      // the typed text '03-1234-5678' is plain text in both engines: no quotePrefix (the apostrophe) is stored or
      // shown unless the doc asked for it (CellWrite.quote) — the two engines must not differ here
      if (excelAvailable() && telLetter) {
        const cell = `名簿!${telLetter}4`;
        const dumps = await inspectProbe([outputs.js, file], [cell], false);
        dumps.forEach((d, i) => {
          const pr = d.cellProbes![cell];
          const label = i === 0 ? '内蔵エンジン' : 'Excel本体';
          test.info().annotations.push({ type: `入力したセル ${label}`, description: `${cell}: ${JSON.stringify(pr)}` });
          expect.soft(pr.value, `${label}: ${cell} は入力した文字のまま`).toBe('03-1234-5678');
          expect.soft(pr.prefix || null, `${label}: ${cell} に quotePrefix（先頭のアポストロフィ）が付いていない`).toBeNull();
        });
      }
    }
  });
});

async function verify(file: string): Promise<CheckOutputResult> {
  test.skip(!(await hasExcel()), 'check-output には Excel が必要です');
  const v = await checkOutput(file);
  test.info().annotations.push({ type: 'check-output', description: `${v.verdict}\n${describeVerdict(v)}` });
  expect(v.openMode, `Excel で修復なしに開けること
${describeVerdict(v)}`).toBe('normal');
  const fails = v.items.filter((i) => i.level === 'fail' && MUST_NOT_FAIL.includes(i.code));
  expect(fails, describeVerdict(v)).toEqual([]);
  expect(v.items.some((i) => i.code === 'formula-columns' && i.level === 'pass'), '集計の数式が同じ見出しの列を参照していること').toBe(true);
  expect(v.stats.formulasChecked as number).toBeGreaterThan(100);
  return v;
}

async function hasExcel(): Promise<boolean> {
  return excelAvailable();
}
