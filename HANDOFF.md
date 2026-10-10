# 引き継ぎ書 — Excel 統合ツール（2026-10-10 時点・ラウンド9「案A」の実装と検証が完了）

別セッションで作業を続けるための資料です。**まずこのファイル、次に `docs/ARCHITECTURE.md`（末尾の Round 9 の項）を読んでください。**

## 1. 保存場所（フルパス）

| もの | フルパス |
|---|---|
| **プロジェクト本体** | `D:\ダウンロード\ExcelMerger\excel-merger\` |
| 引き継ぎ書（このファイル） | `D:\ダウンロード\ExcelMerger\excel-merger\HANDOFF.md` |
| 設計書（契約・仕様・テストID一覧） | `D:\ダウンロード\ExcelMerger\excel-merger\docs\ARCHITECTURE.md` |
| 共通型・契約 | `src\core\types.ts` / `src\core\contracts.ts`（`PremergeOptions.includeSourceSheets`、`PremergeLayout.addedSheets`、`MergeReport.addedSheets`） |
| シート名の決まり（契約） | `src\core\xlsx\sheetnames.ts`（`planSheetNames`）。JS 版の写し: `tools\lib\sheetnames.mjs`（`tests\qa\sheetnames-port.test.ts` が一致を保証） |
| 他のファイルのシートをコピーする本体 | `src\core\xlsx\patch\addsheets*.ts`, `copysheet.ts`, `copyformula.ts`, `copyrules.ts`, `copynames.ts`, `bookinfo.ts`（呼び出しは `src\core\xlsx\premerge.ts`） |
| 数式をコピー側へ付け替える | `src\core\formula\remap.ts`（`RemapSpec.sheetCopies`） |
| 画面（2 つのチェック・一覧・警告） | `src\state\sheetPlan.ts`, `src\components\OutputSheets.tsx`, `SettingsBar.tsx`, `SidePanel.tsx`, `MergePreview.tsx` |
| **ラウンド9の総合検査（本物の Excel）** | `tests\oracle\copied-sheets\`（`run.ps1`・`RESULTS.md`） |
| 出力の検査ツール | `tools\check-output.mjs`（`--single-sheet` / `--with-copies --source`）、`tools\lib\verdict.mjs`・`copies.mjs`、`tools\compare.mjs`。説明: `tools\README.md` |
| 画面の通し試験 | `e2e\`（`12-round9-sheets.spec.ts` がラウンド9）、`tests\ui\`（U9 の `sheets-option.spec.ts` ほか。設定 `tests\ui\playwright.config.ts`、ポート 5192） |
| サンプル 5 本・デモ手順 | `samples\`（`samples\README.md`） |
| 起動用バッチ | `start.bat`（ASCII のみ） |
| exe（**ラウンド9まで入った最新版・2026-10-10 17:59 ビルド**） | `ExcelMerger.exe`（プロジェクト直下） |
| 配布フォルダ / zip（**同じ最新版**） | `release\ExcelMerger\` / `release\ExcelMerger_20261010.zip`（33.5MB。作り直しは `npm run release`） |
| 作業のコピー（ラウンド8のコミット時点の全体） | `D:\ダウンロード\ExcelMerger\excel-merger_copy_20261010_0456\` |
| 作業用の一時フォルダ（消してよい） | `D:\ダウンロード\ExcelMerger\.tmp\`（検査の途中成果は `.tmp\v9\`） |
| Claude のメモリ | `C:\Users\81802\.claude\projects\D--20260705-----------\memory\project-excel-merger.md` |
| ※ユーザーが別に持っているコピー（触らない・2026-10-02 時点の版） | `D:\ダウンロード\excel-merger\`（ここの exe が起動中のことがある） |

## 2. 現在の状態

### ラウンド1〜8（完了・コミット済み）
複数Excelのドラッグ＆ドロップ → キー自動検出 → 結合 → グリッド編集（列の複数選択・ドラッグ移動・挿入・削除・元に戻す／やり直す）→ 出力（内蔵エンジン／Excel本体エンジン）。カードのドラッグ入れ替え・帯のクリック／ドラッグ（ラウンド7）、
出力は既定で「結合した 1 シートだけ」（`dropOtherSheets`）、2 つ目以降のファイルの数式を数式のまま取り込む（ラウンド8、`formulaStats`）。

### ラウンド9「案A」（完了・コミット済み・exe と配布 zip も作り直し済み）
ユーザーの依頼: 出力に**ほかのシートも全部**入れる。結合した表が先頭、そのほかのシートは**各ファイルの元のシートをそのまま**。シート名は「ファイル名_シート名」、重なったら**あとから追加したほうに (1)(2)…**。
別のシートを参照する数式は、コピーしたシートに付け替えてエラーにしない。

* 画面: 設定バーに 2 つのチェック。「他のシートも出力に含める」（基準ファイルのほかのシート）を入れると、下の「他のファイルのシートも含める」も入る（外せば基準のシートだけ）。**既定はどちらもオフ（1 シートだけ）**。結合後は変えられない。結合前に「出力されるシート」の一覧、結合後は「シート一覧」。
* コピー本体: セル・書式・数式（シート参照を付け替え）・条件付き書式・入力規則・結合・名前・ハイパーリンク・図・グラフ・画像・メモ・表・ピボット・グラフシート・非表示のシート・1904 年の日付など。コピーできないもの
  （スライサー・SmartArt・フォームコントロール・マクロシートなど）は `sheetCopyLimits` の警告に、シートごとに書く。限界の一覧は `README.md`（利用者向け）と `tests\golden\RESULTS.md`（Round 9）。
* 数式: 結合した表に入れられない数式（別のシートを引く VLOOKUP、表の外のセル、名前、構造化参照など）は、**コピーしたシートを読む数式のまま**残す（`formulaStats.reasons.viaCopy`、項目 `formulasViaCopy`）。

#### 検証した結果（ラウンド9・V9）
| 何を | 結果 |
|---|---|
| 総合検査（本物の Excel・両エンジン・ランダム編集） `tests\oracle\copied-sheets\` | 72 通りの組み合わせ・78 本のランダム編集・コピーしたシート 355 枚（約 12 万セル）・Excel で開いて再計算したブック 337 冊。修復なし・新しい #REF!/#NAME? なし・シート名と並びは決まりどおり・コピーは元と同じ・編集でコピーは変わらない・取り込んだ数式の値は元と同じ・両エンジンの数式の差 0。わざと入れた不具合 27 件を全部見つけた。**見つけた本物の不具合は 1 件（ウィンドウ枠の固定）で、直して回帰試験を付けた**。詳細・再現手順・直っていない差は `tests\oracle\copied-sheets\RESULTS.md` |
| `tools\check-output.mjs --single-sheet` / `--with-copies --source …` | 実装済み（単体試験 `tests\qa\verdict-shapes.test.ts`、実 Excel の E2E `e2e\12-round9-sheets.spec.ts`、総合検査の `-CheckOutput`） |
| E2E（`E2E_BUILD=1 npx playwright test`、サーバーモード）／スタンドアロン | サーバーモード 2 回とも **60 件合格・9 件スキップ・失敗 0**（約 18 分）。スタンドアロン（`E2E_BASE_URL=http://127.0.0.1:5176`）2 回とも **59 件合格・10 件スキップ・失敗 0**（11〜15 分）。スキップは、モードが違うと意味のない試験（`?mock` はスタンドアロンだけ、Excel エンジンはサーバーモードだけ）。不安定な試験は出なかった |
| `tests\ui`（U9・Playwright） | 82 件合格・6 件はモック用でスキップ。U9 が「コピー本体が来るまで fixme」にしていた 3 件（`sheets-option` の real engine）は、今回すべて合格 |

### 未了（次のセッションでやること）— 優先順
（コミット・`ARCHITECTURE.md` への追記・exe のビルド・配布 zip の作り直しは 2026-10-10 に済み。最新の exe で `tests\ui\real.mjs --sheets=single|all` を通し、`tools\check-output.mjs`（`--single-sheet` / `--with-copies`）で 4 つの出力が合格、内蔵エンジンと Excel本体の出力の差は 0、シート名は
`名簿 | 集計 | ピボット | マスタ | 02_面接評価_面接評価 | 02_面接評価_評価テーブル | 03_連絡先_03_連絡先 | 04_適性検査_適性検査結果 | 05_研修出欠_出欠` を確認済み。）
1. 実機未確認: 実際の IME 入力、Explorer からの本物のドラッグ、他の PC での Defender の挙動、ユーザーの実データ（大きなファイル）でのコピーの速さ（10 シート×2 万行×30 列で約 17〜21 秒、メモリ約 2 GB と測っている）。
2. 直っていない既知の問題（最小の再現つき）は **`tests\oracle\copied-sheets\RESULTS.md`** と、このファイルの「4. 既知の制限」。

## 3. 守ってほしいこと（失敗から学んだ点）
- **Excel の扱い**: テストで使うExcelは**自分専用の非表示インスタンス**だけ。Excel を名前で kill しない（ユーザーのExcelが開いている）。止めるのは自分の PID だけ。クリップボードと PageSetup は使わない。
- **ディスク**: C: は逼迫しやすい（Excel 起動のたびに診断ログ約 80MB が `%TEMP%\Diagnostics\EXCEL` に出る）。`$env:TEMP='D:\ダウンロード\ExcelMerger\.tmp'; $env:TMP=$env:TEMP` を設定してから vitest／node／playwright を動かし、Excel の起動回数を減らす。総合検査は Excel の起動ごとに専用の TEMP を作って消す。
- **起動中の exe**: ユーザーが `ExcelMerger.exe` を開いていることがある。kill しない。上書きしない。
- **契約ファイル**（`types.ts`, `contracts.ts`, 各 `index.ts`, `sheetnames.ts`, `ARCHITECTURE.md`）は、変更の影響が大きい。変えるときは影響範囲を全部確認する。`tools\lib\sheetnames.mjs` は `sheetnames.ts` の写しなので、片方を変えたら両方変える（試験が落ちる）。
- 両エンジン（内蔵／Excel本体）は**同じ操作ログ（`LoggedOp`）を再生**する。片方だけ仕様を変えない。`.bat` は ASCII のみ。UI の URL は `http://localhost`（`127.0.0.1` は不可）。
- コードは Write/Edit で書く（シェルの heredoc・`node -e`・`sed` はバックスラッシュを消す／変える。今回も `\d` や `\\` が何度も壊れた）。正規表現と Windows のパスは編集のあとに必ず見直す。
- PowerShell のスクリプトは UTF-8 BOM 付きか ASCII のみ（日本語のパスを `.ps1` に直接書かない。パラメータで渡す）。
- E2E・検査を同時に走らせるときは、ポート 5199 / 5176（`playwright.config.ts`）と 5192（`tests\ui`）を他の作業と分ける。

## 4. 既知の制限（ユーザーに伝えてよいもの／直せていないもの）
- Excel本体で列を動かすとクリップボードの中身が消える。元に戻せなかったときは警告を出す。
- VLOOKUP の列番号は Excel も自動で直さない（該当時は警告）。
- exe に署名がない（他のPCでは SmartScreen が出ることがある。対処は README に記載）。
- 画面は 1280px 未満だと横スクロール。
- コピーされないもの（README の「コピーされないもの・変わるもの」と同じ）: スライサー・タイムライン・SmartArt・インク・フォームコントロール／ActiveX／OLE・ヘッダー／フッターの画像・マクロシート・VBA（他ファイルの）、Strict と通常ファイルの間の図・グラフ。
  外部ブック／3-D／構造化参照の数式は値になる。スレッド形式のコメントは普通のメモになる。外部データのピボットは値のみ。
- 総合検査で見つかって直っていないもの（どれも見た目か、README に書いた制限。RESULTS.md の 3.4）: 既定の縦位置・テーマのフォント名がコピー先の既定になる／外部ブック・3-D・構造化参照の数式は値になる／スライサー等はコピーされない／Excel エンジンは表の列を表の外へ動かせず内蔵エンジンに切り替わる（1/78）／
  **元データの列を動かしたピボットは、内蔵エンジンでは保存済みの表示のまま、Excel エンジンでは作り直された表示になる**（開いたときの更新で揃う。既存の動作）／Strict 形式のファイルの日付セル（`t="d"`）はそのままコピーされる（Excel 2010 以降は読める）。行の高さ・図の位置は検査していない。
- 直したもの（コミットされていません。確認してください）: `src\core\xlsx\patch\views.ts`（`frozenAfter`）と `src\core\doc\struct.ts`（`freezeOp.cut`）— 固定した列・行をすべて削除／移動で外すと、Excel は 1 列（行）だけ残す。内蔵エンジンと画面の文書は元の本数のままだった。`views.ts` は P の持ち物ですが、2 行の変更なので V9 が直しました。
- 実機未確認: 実際の IME 入力、Explorer からの本物のドラッグ、他PCでの Defender の挙動。

## 5. 新しいセッションに貼る文（コピーして使う）

> `D:\ダウンロード\ExcelMerger\excel-merger\HANDOFF.md` と `docs\ARCHITECTURE.md` を読んで、HANDOFF の「未了」を上から順に完了させてください。ラウンド9（他のシート・他のファイルのシートを出力に入れる「案A」）まですべてコミット・ビルド済みです。残りは「未了」の実機未確認の項目と、ユーザーの新しい依頼だけです。Excel を名前で kill しない・起動中の ExcelMerger.exe を上書きしない・TEMP を D: に向ける、を守ってください。

---

## 6. 再開のしかた（コマンド集）

```
cd D:\ダウンロード\ExcelMerger\excel-merger
$env:TEMP='D:\ダウンロード\ExcelMerger\.tmp'; $env:TMP=$env:TEMP
npx tsc --noEmit
npx vitest run                                              # 約 3,000 件。時間を測る試験は 1 回だけ落ちることがある（単独で再実行）
E2E_BUILD=1 npx playwright test                             # サーバーモード（ポート 5199 / 5176）。Excel を使う。約 20〜30 分
npx vite --config e2e/vite.e2e.config.ts                    # スタンドアロン（別の窓で。ポート 5176）
$env:E2E_BASE_URL='http://127.0.0.1:5176'; npx playwright test
npx playwright test -c tests/ui/playwright.config.ts        # U9 の画面試験（内蔵エンジンのみ。ポート 5192。tests\ui\shots の画像が書き換わるので git checkout で戻す）
powershell -NoProfile -ExecutionPolicy Bypass -File tests\oracle\copied-sheets\run.ps1 -Only '^(smp-01-all-keep|ob-all-keep)$' -CheckOutput 4 -Negative
powershell -NoProfile -ExecutionPolicy Bypass -File tests\oracle\copied-sheets\run.ps1      # 全部（70 通り以上・約 1.5 時間）
EXCEL_TESTS=1 npx vitest run tests/oracle/copied-sheets     # 検査を vitest から（EXCEL_ORACLE_ONLY で絞る）
node tools\check-output.mjs <出力.xlsx> --base samples\01_学生名簿.xlsx --single-sheet
node tools\check-output.mjs <出力.xlsx> --base samples\01_学生名簿.xlsx --with-copies --source samples\02_面接評価.xlsx --source samples\03_連絡先.csv --source samples\04_適性検査.xls --source samples\05_研修出欠.xlsx
```

---

## 付録: 履歴（過去の依頼）

### ラウンド7（ユーザーの追加依頼・実装済み）
1. 画面上部の**ファイルカード**をドラッグして順番を入れ替えられる。2. グリッド上の**ファイルごとの色の帯**を掴んで、その列グループごと任意の位置へ移動できる。3. 帯を**クリック**するとその列グループが全選択される。
4. **カードをクリック**するとグリッドがそのファイルの列が中央付近に来るよう移動し、その列が全選択される。5. 色は説明用。実装は `DocColumn.origin` で判定する。6. **出力は結合した 1 シートだけ**（今回の案A で、設定により他のシートも入れられるようになった）。
ラウンド7の未了だった項目は、ラウンド9で次のように片づいた: 実機での通し確認＝`e2e\` と総合検査、1 シート出力の実 Excel 検証＝総合検査（`dropOtherSheets` の組み合わせ）と `e2e\12` (a)、検査ツール・E2E の更新＝V9、ビルドと配布物のみ未了（上の「未了」3）。

### ラウンド8（完了・コミット済み）: 2つ目以降のファイルの数式
`src\core\formula\remap.ts` が、行の対応・表全体の範囲・取り込まないキー列・名前付き範囲・表の項目名参照を付け替える。実Excelで 579 件の取り込み検証、554 通りの編集で値の食い違い 0、内蔵エンジンと Excel本体の出力の差 0。
詳細: `tests\xlsx\formula-import-findings.md`, `tests\oracle\consistency\RESULTS.md`, `tests\golden\RESULTS.md`。未解決（V1 が報告）: join 'full' で、基準の最終データ行にある累計 `=SUM($D$2:D31)` が追加行まで伸びる件は、ラウンド9で直った（`regress.test.ts`）。
元に計算結果が無い数式は、画面では灰色の式で表示され、「値に変換して削除」すると空白になる（計算エンジンが無いため）。
