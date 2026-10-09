// Round 5: files real users actually drop (owner: T) — a workbook cut off in the middle, a password-protected one, an empty file,
// a zip that is not a workbook, .txt / pictures, Excel's lock files, the same file twice. The app must say what is wrong ON THE
// CARD of that file, never throw, never lose the other files, and stay usable (merge, export, clear the bad card).
import path from 'node:path';
import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';
import {
  columnTitles,
  exportBoth,
  loadFiles,
  mergeNow,
  openApp,
  SAMPLE_FILES,
  SAMPLES_DIR,
  samplePaths,
  step,
  toastLog,
  waitMergeReady,
  watchErrors,
} from './helpers';
import { alignmentProblems, dropFiles, passwordProtectedXlsx, payload, readSheet, truncatedXlsx, verifyCases, waitIdle } from './round5';

const TWO = [SAMPLE_FILES[0], SAMPLE_FILES[1]];

/** console / page errors that are not the app's (favicon, devtools) */
const appErrors = (errors: { messages: string[] }) => errors.messages.filter((m) => !/favicon|DevTools|Failed to load resource/i.test(m));

/** a zip that is a perfectly good archive but no workbook */
async function zipNotWorkbook(): Promise<Buffer> {
  const zip = new JSZip();
  zip.file('readme.txt', 'これは Excel ブックではありません');
  zip.file('data/list.csv', 'a,b\n1,2\n');
  return zip.generateAsync({ type: 'nodebuffer' });
}

/** The error text of a card (the red line under the header), or null when the card has none. */
async function cardError(page: Page, index: number): Promise<string | null> {
  const loc = page.getByTestId(`file-card-${index}`).locator('.card-error');
  return (await loc.count()) ? (await loc.first().innerText()).trim() : null;
}

test.describe('ラウンド5: 壊れた・パスワード付き・空のファイル', () => {
  test('(a) 途中で切れた xlsx / パスワード付き / 空 / zip だが Excel ではない → カードに理由が出て、アプリはそのまま使える', async ({ page }) => {
    const errors = watchErrors(page);
    await openApp(page);
    await dropFiles(
      page,
      [
        payload('途中で切れた.xlsx', truncatedXlsx()),
        payload('パスワード付き.xlsx', passwordProtectedXlsx()),
        payload('空.xlsx', Buffer.alloc(0)),
        payload('zipだがExcelではない.xlsx', await zipNotWorkbook()),
      ],
      4,
    );
    const expected: Array<[string, RegExp[]]> = [
      ['途中で切れた.xlsx', [/途中で切れた\.xlsx: Excelファイルとして読み込めませんでした/, /ファイルが壊れているか、Excel形式ではありません/]],
      ['パスワード付き.xlsx', [/パスワード付き\.xlsx: パスワードで保護されているため読み込めません/, /Excelで保護を解除して保存し直してください/]],
      ['空.xlsx', [/空\.xlsx: 空のファイルです/]],
      ['zipだがExcelではない.xlsx', [/zipだがExcelではない\.xlsx: Excelファイルとして読み込めませんでした/]],
    ];
    for (const [i, [name, patterns]] of expected.entries()) {
      const card = page.getByTestId(`file-card-${i}`);
      await expect(card.locator('h2'), `${i + 1} 枚目のカードはファイル名を出す`).toHaveText(name);
      await expect(card, `${name}: エラーのカード`).toHaveClass(/has-error/);
      const msg = await cardError(page, i);
      expect(msg, `${name} のカードに理由が出る`).not.toBeNull();
      for (const re of patterns) expect(msg!, name).toMatch(re);
      expect(msg!, `${name}: 画面に英語の技術的なメッセージ（スタック）を出さない`).not.toMatch(/\bat \w+.*\(|Error:|undefined|\[object/);
      // nothing can be done with the file but remove it
      await expect(page.getByTestId(`base-toggle-${i}`), `${name}: 基準にはできない`).toBeDisabled();
      await expect(page.getByTestId(`file-clear-${i}`), `${name}: クリアできる`).toBeEnabled();
      await expect(card.locator('.mini-preview'), `${name}: プレビューは出さない`).toHaveCount(0);
    }
    // no merge possible, and the app says why (two readable files are needed); the drop zone is still there
    await expect(page.getByTestId('merge-button')).toBeDisabled();
    await expect(page.getByTestId('merge-button')).toHaveAttribute('title', /ファイルを2つ以上読み込んでください/);
    await expect(page.getByTestId('drop-zone')).toBeVisible();
    await expect(page.getByTestId('busy-overlay')).toHaveCount(0);
    expect(appErrors(errors), 'ページのエラーがない').toEqual([]);

    // a good pair afterwards: the first READABLE file is the base, everything merges, the bad cards stay where they are
    await dropFiles(page, samplePaths(TWO), 6);
    await waitMergeReady(page, errors);
    await expect(page.getByTestId('base-toggle-4'), '最初に読めたファイルが基準').toHaveAttribute('aria-pressed', 'true');
    await expect(page.getByTestId('base-toggle-0')).toHaveAttribute('aria-pressed', 'false');
    await expect(page.getByTestId('summary-line')).toContainText(/基準\s*60\s*名/);
    for (let i = 0; i < 4; i++) expect(await cardError(page, i), `エラーのカード ${i + 1} はそのまま`).not.toBeNull();
    await mergeNow(page, errors);
    const titles = await columnTitles(page);
    expect(titles.slice(0, 12)).toEqual(['学生ID', '氏名', 'フリガナ', '大学', '学部', '学年', 'メール', '電話', '生年月日', '年齢', '志望度', '重複チェック']);
    for (const h of ['面接官', '評価1', '平均', '判定', 'コメント']) expect(titles).toContain(h);
    expect(appErrors(errors), '結合までページのエラーがない').toEqual([]);
  });

  test('(b) 正しいファイルと壊れたファイルを一緒に投げる → 読めたものだけ結合でき、出力も正しい。エラーのカードは「クリア」で外せる', async ({ page }) => {
    test.setTimeout(10 * 60_000);
    const errors = watchErrors(page);
    await openApp(page);
    const [f01, f02] = samplePaths(TWO);
    // one drop: a good file, a cut-off file, a good file, a protected file (the order of a real folder)
    await dropFiles(page, [f01, payload('壊れている.xlsx', truncatedXlsx('02_面接評価.xlsx', 0.5)), f02, payload('保護されている.xlsx', passwordProtectedXlsx())], 4);
    await waitMergeReady(page, errors);
    await expect(page.getByTestId('base-toggle-0')).toHaveAttribute('aria-pressed', 'true');
    expect(await cardError(page, 0)).toBeNull();
    expect(await cardError(page, 1)).toContain('読み込めませんでした');
    expect(await cardError(page, 2)).toBeNull();
    expect(await cardError(page, 3)).toContain('パスワード');
    await expect(page.getByTestId('summary-line')).toContainText(/基準\s*60\s*名/);

    // clear the bad cards (the cards renumber): the numbers do not change
    await page.getByTestId('file-clear-3').click();
    await expect(page.locator('[data-testid^="file-card-"]')).toHaveCount(3);
    await page.getByTestId('file-clear-1').click();
    await expect(page.locator('[data-testid^="file-card-"]')).toHaveCount(2);
    for (let i = 0; i < 2; i++) expect(await cardError(page, i)).toBeNull();
    await waitMergeReady(page, errors);
    await expect(page.getByTestId('summary-line')).toContainText(/基準\s*60\s*名/);

    await mergeNow(page, errors);
    const out = await exportBoth(page, errors, 'e2e_r5b_壊れたファイルと一緒');
    const grid = readSheet(out.files[0].file, '名簿');
    const { problems, checked } = alignmentProblems(grid, '01_学生名簿.xlsx', ['02_面接評価.xlsx']);
    expect(checked, '照合した行数').toBe(60);
    expect(problems, problems.slice(0, 5).join('\n')).toEqual([]);
    await verifyCases(out.files.map((f) => ({ label: f.engine, base: path.join(SAMPLES_DIR, SAMPLE_FILES[0]), file: f.file })));
    expect(appErrors(errors)).toEqual([]);
  });

  test('(c) 結合のあとで壊れたファイルを足そうとすると確認が出て、やめれば何も変わらない。続ければキーの画面に戻ってエラーのカードが増える', async ({ page }) => {
    test.setTimeout(10 * 60_000);
    const errors = watchErrors(page);
    await openApp(page);
    await step(page, errors, 'ファイル読み込み', () => loadFiles(page, samplePaths(TWO)));
    await waitMergeReady(page, errors);
    await mergeNow(page, errors);
    const before = await columnTitles(page);

    await page.getByTestId('file-input').setInputFiles([payload('あとから.xlsx', truncatedXlsx())]);
    const dialog = page.getByTestId('confirm-dialog');
    await expect(dialog, '結合後にファイルを足すので確認が出る').toBeVisible();
    await expect(dialog).toContainText('ファイルを追加しますか？');
    // cancel: the merged sheet and the file list are untouched
    await page.getByTestId('confirm-cancel').click();
    await expect(dialog).toBeHidden();
    await expect(page.getByTestId('grid')).toBeVisible();
    expect(await columnTitles(page)).toEqual(before);
    await expect(page.locator('[data-testid^="file-card-"]'), 'カードは増えない').toHaveCount(2);

    // go on: back to the key screen, the broken file shows its reason, the two good ones can be merged again
    await page.getByTestId('file-input').setInputFiles([payload('あとから.xlsx', truncatedXlsx())]);
    await expect(dialog).toBeVisible();
    await page.getByTestId('confirm-ok').click();
    await expect(dialog).toBeHidden();
    await expect(page.locator('[data-testid^="file-card-"]')).toHaveCount(3);
    await expect(page.getByTestId('grid'), 'キーの画面に戻る').toHaveCount(0);
    expect(await cardError(page, 2)).toContain('読み込めませんでした');
    await waitMergeReady(page, errors);
    await mergeNow(page, errors);
    expect(await columnTitles(page)).toEqual(before);
    expect(appErrors(errors)).toEqual([]);
  });
});

test.describe('ラウンド5: 対応していないファイル・ロックファイル・同じファイル', () => {
  test('(d) .txt と画像は受け付けず理由を出し、Excel のロックファイル (~$) は黙って無視し、同じファイルを 2 回入れると知らせる', async ({ page }) => {
    const errors = watchErrors(page);
    await openApp(page);
    const [f01] = samplePaths([SAMPLE_FILES[0]]);
    await page.getByTestId('file-input').setInputFiles([
      payload('メモ.txt', Buffer.from('memo')),
      payload('写真.png', Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])),
      payload('~$名簿.xlsx', Buffer.alloc(165, 0x41)),
    ]);
    await expect.poll(async () => (await toastLog(page)).some((t) => t.includes('メモ.txt') && t.includes('写真.png') && t.includes('対応していない形式です')), { message: '対応していない形式の理由が出る' }).toBe(true);
    const toast = (await toastLog(page)).find((t) => t.includes('メモ.txt'))!;
    expect(toast, '対応している形式の一覧が案内される').toContain('.xlsx');
    expect(toast, 'ロックファイル (~$) の名前は出さない（黙って無視）').not.toContain('~$');
    await expect(page.locator('[data-testid^="file-card-"]'), '受け付けなかったファイルのカードは作らない').toHaveCount(0);
    await expect(page.getByTestId('drop-zone')).toBeVisible();

    await dropFiles(page, [f01], 1);
    await waitIdle(page);
    const n = (await toastLog(page)).length;
    await page.getByTestId('file-input').setInputFiles([f01]);
    await expect.poll(async () => (await toastLog(page)).slice(n).some((t) => t.includes('同じファイルはすでに読み込まれています')), { message: '同じファイルを 2 回入れたことを知らせる' }).toBe(true);
    await expect(page.locator('[data-testid^="file-card-"]'), 'カードは増えない').toHaveCount(1);
    expect(appErrors(errors)).toEqual([]);
  });

  test('(e) HTML を .xls の名前で保存したファイル（業務システムの出力によくある）と、中身が文字だけの .xlsx は、落ちずに表として読み込まれる', async ({ page }) => {
    const errors = watchErrors(page);
    await openApp(page);
    const html = Buffer.from(
      '<html><head><meta charset="utf-8"></head><body><table><tr><th>学生ID</th><th>氏名</th><th>点数</th></tr>' +
        '<tr><td>000100</td><td>石川　莉子</td><td>80</td></tr><tr><td>000101</td><td>鈴木 花子</td><td>75</td></tr></table></body></html>',
      'utf8',
    );
    await dropFiles(page, [payload('システム出力.xls', html)], 1);
    const card = page.getByTestId('file-card-0');
    expect(await cardError(page, 0), 'エラーにならない').toBeNull();
    await expect(card).toContainText('2 行');
    await expect(card.locator('.mini-preview')).toContainText('石川');
    await expect(card, '書式のない「値のみ」として読むことを伝える').toContainText('値のみ');

    await dropFiles(page, [payload('ただの文字.xlsx', Buffer.from('これはExcelではありません。\n'.repeat(5), 'utf8'))], 2);
    await expect(page.getByTestId('file-card-1').locator('h2')).toHaveText('ただの文字.xlsx');
    await expect(page.getByTestId('file-card-1'), '文字だけのファイルも値のみの表として読む').toContainText('値のみ');
    expect(appErrors(errors)).toEqual([]);
  });
});
