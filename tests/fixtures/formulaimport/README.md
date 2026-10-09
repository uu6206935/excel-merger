# formulaimport fixtures (round 8, task F1)

Source workbooks for the formula-import tests and the real-Excel oracle (`tests/oracle/import-values/`). A subset of the generated corpus
(shuffled order of every item, plus the "same" order of f01 / f05 / f06):

* `base.xlsx` — the generated base (sheet 名簿, 30 students S001…S030); `base_samples.xlsx` — a copy of `samples/01_学生名簿.xlsx`.
* `nocache/*.xlsx` — written by openpyxl / XlsxWriter (`gen_corpus.py`): formulas **without** stored results (XlsxWriter: placeholder zeros).
* `cached/*.xlsx` — the same workbooks opened, fully recalculated and saved by real Excel (`author.ps1`; f02 has its shared formulas created
  there). `cached/f01_samerow__same.xlsm` — the macro-enabled twin.
* `manifest.json` — one entry per file (`base` = the base workbook of the pair when it is not `base.xlsx`).
* `expected.json` — counters (kept / value / empty / reasons / widened / uncached / substituted) of every file under four configurations;
  `tests/xlsx/formula-import-corpus.test.ts` recomputes them. Rewrite with `npx tsx tests/oracle/import-values/make-expected.ts` after a
  deliberate change of behaviour and review the diff.

Persons: 27 of the 31 source people match the base, 4 exist only in the source (join "full" appends them), 3 base people are missing in
the source. Items: f01 same-row · f02 shared formulas · f03 key columns · f04 lookups · f05 ranges · f06 running totals / side cell / ROW · f07 other
sheet · f08 defined names · f09 Excel Table · f10 arrays · f11 result types · f12 footer · f13 HYPERLINK · f14 two-row header · f16 unsafe ·
g01 the app's sample base. The whole corpus (48 files) is rebuilt by `powershell -File tests\oracle\import-values\run.ps1 -Regenerate`.
