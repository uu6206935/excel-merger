# 引き継ぎ書 — Excel 統合ツール（2026-10-07 時点）

別セッションで作業を続けるための資料です。**まずこのファイル、次に `docs/ARCHITECTURE.md` を読んでください。**

## 1. 保存場所（フルパス）

| もの | フルパス |
|---|---|
| **プロジェクト本体** | `D:\ダウンロード\ExcelMerger\excel-merger\` |
| 引き継ぎ書（このファイル） | `D:\ダウンロード\ExcelMerger\excel-merger\HANDOFF.md` |
| 設計書（契約・仕様・テストID一覧） | `D:\ダウンロード\ExcelMerger\excel-merger\docs\ARCHITECTURE.md` |
| 調査レポート R1〜R6 | `D:\ダウンロード\ExcelMerger\excel-merger\docs\research\` |
| ラウンド7の依頼文（詳細仕様つき） | `D:\ダウンロード\ExcelMerger\excel-merger\docs\handoff\round7-wave1-workflow.js` |
| 共通型・契約 | `D:\ダウンロード\ExcelMerger\excel-merger\src\core\types.ts` / `src\core\contracts.ts` |
| 起動用バッチ | `D:\ダウンロード\ExcelMerger\excel-merger\start.bat` |
| exe（**ラウンド6までの版**） | `D:\ダウンロード\ExcelMerger\excel-merger\ExcelMerger.exe` |
| 配布フォルダ / zip（**ラウンド6までの版**） | `D:\ダウンロード\ExcelMerger\excel-merger\release\ExcelMerger\` / `…\release\ExcelMerger_20261002.zip` |
| サンプル 5 本 | `D:\ダウンロード\ExcelMerger\excel-merger\samples\` |
| 検証ツール（出力チェック・比較） | `D:\ダウンロード\ExcelMerger\excel-merger\tools\check-output.mjs` / `tools\compare.mjs` |
| 参考にした元アプリ（見た目の基準） | `D:\ダウンロード\ExcelMerger\interview-scheduler\` |
| 作業用の一時フォルダ（消してよい） | `D:\ダウンロード\ExcelMerger\.tmp\` |
| Claude のメモリ | `C:\Users\81802\.claude\projects\D--20260705-----------\memory\project-excel-merger.md` |
| ※ユーザーが別に持っているコピー（触らない・2026-10-02 時点の版） | `D:\ダウンロード\excel-merger\`（ここの exe が起動中のことがある） |

## 2. 現在の状態

### 完成・検証済み（ラウンド1〜6）
複数Excelのドラッグ＆ドロップ → キー自動検出（ID・氏名・学籍番号・メール・電話・カナ・氏名＋大学など、手動キー追加も可）→ 結合 → グリッド編集（列の複数選択・ドラッグ移動・挿入・削除・元に戻す／やり直す）→ 出力（内蔵エンジン／Excel本体エンジン）。実際のExcelで「修復なし・`#REF!` なし」を確認済み。`ExcelMerger.exe` と `start.bat` で配布可能。

### ラウンド7（ユーザーの追加依頼）— **コードは入っているが、最終検証と再ビルドが未了**
ユーザーの依頼（原文の趣旨）:
1. 画面上部の**ファイルカード**をドラッグして順番を入れ替えられる。
2. グリッド上の**ファイルごとの色の帯**（「01_学生名簿（基準）2列」など）を掴んで、その列グループごと任意の位置へ移動できる。
3. 帯を**クリック**するとその列グループが全選択される。
4. **カードをクリック**するとグリッドがそのファイルの列が中央付近に来るよう移動し、その列が全選択される。
5. 色は説明用。実装は「どのファイル由来の列か（`DocColumn.origin`）」で判定する。
6. **出力は結合した 1 シートだけ**。他のシート・グラフ・統計は出力に入れない。

実装状況（コードで確認済み。2026-10-07 時点で `npx tsc --noEmit` クリーン、`npx vitest run` は 2,512 件合格・45 件スキップ）:

| 項目 | 状況 | 主なファイル |
|---|---|---|
| カードのドラッグ入れ替え（結合前・結合後） | 実装あり。画面写真で挿入位置の表示を確認 | `src\components\FileCardRow.tsx`, `FileCard.tsx`, `src\state\groups.ts`, `reducer.ts`, `controller.ts` |
| 帯のクリック選択・ドラッグ移動 | 実装あり（テスト `tests\grid\runs.test.ts`, `tests\grid\e2e\bands.spec.ts`） | `src\components\grid\runs.ts`, `Grid.tsx`, `selection.ts` |
| カードのクリック→中央へ移動＋全選択 | 実装あり。写真で 5 列選択を確認 | `Grid.tsx`（`scrollToColumnsCentered`）, `EditorPane.tsx` |
| 1 シート出力 | 実装あり。`PremergeOptions.dropOtherSheets`、画面のチェック「他のシートも出力に含める」（既定オフ）、「出力は結合した1シートのみ」の注記 | `src\core\xlsx\patch\dropsheets.ts`, `premerge.ts`, `src\App.tsx`, `SettingsBar.tsx`, `tests\xlsx\dropsheets-*.test.ts` |

担当エージェントは**報告を出す前にセッション上限で停止**したため、下の「未了」が残っています。

### 未了（次のセッションでやること）— 優先順
1. **ビルド**: `npm run build` → `npm run build:runner` → `npm run build:exe`。*起動中の exe は上書きできない*ので、まず `node scripts\build-exe.mjs --out <別パス>\ExcelMerger.exe` で作り、閉じてから差し替える。
2. **実機での通し確認**（実サーバー＋実Excel）: サンプル 5 本を投入 → カードのドラッグ入れ替え（結合前／結合後）→ 帯のクリック・ドラッグ → カードのクリック → 結合 → **内蔵エンジンと Excel本体の両方で出力** → 出力が**1シート**で、Excelが修復なしで開き、`#REF!` がないこと。`node tests\ui\real.mjs http://localhost:<port>/ samples\01_学生名簿.xlsx …` が叩き台。
3. **1シート出力を実Excelで検証**（担当が未実施。記録なし）: 複数シート入りのブック（`samples\01_学生名簿.xlsx`、`tests\fixtures\golden\torture.xlsx`、`tests\fixtures\foreign\corpus\` の数本）で、修復なし・1シート・孤立パーツなしを確認し、結果を `tests\golden\RESULTS.md` に記録。準備スクリプト: `tests\golden\com\check-dropsheets.ps1`, `make-fixtures-dropsheets.ps1`。
4. **古い仕様に合わせた検証ツール・E2Eの更新**: `tools\check-output.mjs`（`tools\lib\verdict.mjs`）は「シートが消えたら FAIL」「集計シートの数式の参照先を確認」という作りなので、1 シート出力では**必ず FAIL する**。`--single-sheet` のようなオプション（またはレポートの `droppedSheets` からの自動判定）を足し、1 シート出力を合格にする。`e2e\*.spec.ts`（02〜06 が「集計」を前提）と `tests\ui\real.mjs`、`samples\README.md`、`README.md` の記述も更新。
5. **ドキュメント**: `docs\ARCHITECTURE.md` の Round 7 の項に、新しい `data-testid`（カード・帯・「他のシートも出力に含める」チェックなど。`tests\ui\w1-cards.spec.ts` 参照）を追記。
6. **配布物の作り直し**: `npm run release`（exe は 1 の新ビルドを `--exe` で指定）。
7. 全体の最終確認: `npx tsc --noEmit` と `npx vitest run`、`E2E_BUILD=1 npx playwright test`（サーバーモード）。

## 3. 守ってほしいこと（失敗から学んだ点）
- **Excel の扱い**: テストで使うExcelは**自分専用の非表示インスタンス**だけ。Excel を名前で kill しない（ユーザーのExcelが開いている）。止めるのは自分の PID だけ。クリップボードと PageSetup は使わない。
- **ディスク**: C: は逼迫しやすい（Excel 起動のたびに診断ログ約 80MB が `%TEMP%\Diagnostics\EXCEL` に出る）。`$env:TEMP='D:\ダウンロード\ExcelMerger\.tmp'; $env:TMP=$env:TEMP` を設定してから vitest／node を動かし、Excel の起動回数を減らす。
- **起動中の exe**: ユーザーが `ExcelMerger.exe` を開いていることがある（2026-10-07 時点では `D:\ダウンロード\excel-merger\` 側が起動中）。kill しない。上書きしない。
- **契約ファイル**（`types.ts`, `contracts.ts`, 各 `index.ts`, `ARCHITECTURE.md`）は、変更の影響が大きい。変えるときは影響範囲を全部確認する。
- 両エンジン（内蔵／Excel本体）は**同じ操作ログ（`LoggedOp`）を再生**する。片方だけ仕様を変えない。`.bat` は ASCII のみ。UI の URL は `http://localhost`（`127.0.0.1` は不可）。
- コードは Write/Edit で書く（シェルの heredoc はバックスラッシュを消す）。正規表現は自動編集のあとに必ず見直す。

## 4. 既知の制限（ユーザーに伝えてよいもの）
- Excel本体で列を動かすとクリップボードの中身が消える。元に戻せなかったときは警告を出す。
- VLOOKUP の列番号は Excel も自動で直さない（該当時は警告）。
- exe に署名がない（他のPCでは SmartScreen が出ることがある。対処は README に記載）。
- 画面は 1280px 未満だと横スクロール。
- 実機未確認: 実際の IME 入力、Explorer からの本物のドラッグ、マクロ入りファイル、他PCでの Defender の挙動。

## 5. 新しいセッションに貼る文（コピーして使う）

> `D:\ダウンロード\ExcelMerger\excel-merger\HANDOFF.md` と `docs\ARCHITECTURE.md` を読んで、HANDOFF の「未了」を上から順に完了させてください。ラウンド7（カード／帯のドラッグ・クリック選択・1シート出力）はコードが入っていますが、最終検証と再ビルドが未了です。Excel を名前で kill しない・起動中の ExcelMerger.exe を上書きしない・TEMP を D: に向ける、を守ってください。

---

## 追記（2026-10-10）— ラウンド8 完了、次はラウンド9

### ラウンド8（完了・コミット済み）: 2つ目以降のファイルの数式
- 原因: 2つ目以降のファイルは、同じ行の取り込む列を参照する数式しか数式のまま残らず、他は値に、元に計算結果が無いと空白になっていた。
- 対策: `src\core\formula\remap.ts`（新規）が、行の対応（結合キーで並び替わった行）・表全体の範囲・取り込まないキー列（`JoinResult.keyPairs`）・名前付き範囲・表の項目名参照を付け替える。`MergeReport.formulaStats` と画面の「数式の取り込み」欄で件数と理由を表示。
- 検証: 実Excelで 579 件の取り込み検証（37,851 件の数式が元と一致）、554 通りの編集（列の移動・挿入・削除・行の追加）で値の食い違い 0、内蔵エンジンと Excel本体の出力の差 0。
  詳細: `tests\xlsx\formula-import-findings.md`, `tests\oracle\consistency\RESULTS.md`, `tests\golden\RESULTS.md`。
- 未解決（V1 が報告）: join 'full' で、基準の最終データ行にある累計 `=SUM($D$2:D31)` が追加行まで伸びる（`tests\oracle\consistency\regress.test.ts` の `it.fails`）。元に計算結果が無い数式は、画面では灰色の式で表示され、「値に変換して削除」すると空白になる（計算エンジンが無いため）。

### 次: ラウンド9「案A」（ユーザー依頼）
元のExcelの全シートを出力に入れる。シート名は「ファイル名_シート名」、重なったら後から追加した方に (1)(2)…。他のシートを参照する数式は、コピーしたシートを指すよう付け替えて数式のまま残す。設定バーに「他のファイルのシートも含める」を追加（既定オフ）。
- シート名の決め方（`planSheetNames`: 31文字制限・禁止文字・(n) の付け方）は別の場所で試作し、12件のテストが合格している。ラウンド9の最初に `src\core\xlsx\sheetnames.ts` と `tests\xlsx\sheetnames.test.ts` として取り込む（このコミットの時点では、まだリポジトリに入っていない）。
- 作業の分担: A1 コピー本体 → A2 図・表・コメント・ピボット／F2 数式のコピー側への付け替え → V9 実Excelでの検証・ツール・E2E・文書。指示書は `docs\handoff\round9-plan-a-workflow.js` に置く予定（同上、まだ無い）。
- ラウンド7の未了（HANDOFF 上の「未了」1〜7）は、ラウンド9の V9 と最終ビルドで消化する。
