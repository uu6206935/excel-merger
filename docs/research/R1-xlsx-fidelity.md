# R1: xlsx read/write library fidelity for the JS export engine

Scratch folder: `C:\Users\81802\AppData\Local\Temp\claude\D--20260705-----------\2d64adba-cbb0-49f4-bf69-6262ff0972bb\scratchpad\research\r1`
(all paths below are relative to it). Date of test: 2026-09-30. Machine: Windows 11, Excel 16 (M365, build with LAMBDA/dynamic arrays), Node 24.16.0.

## TL;DR

1. **No JS library can load and re-save a real Excel workbook without damage, except xlsx-populate (almost) and a raw ZIP pass-through (always).** With ExcelJS 4.4.0 (and every ExcelJS fork), a normal Excel-authored workbook came back needing **repair** (table autoFilter), with charts, pivots, threaded comments, cross-sheet list validations, whole-column names, sheet-scoped names, internal hyperlinks and dynamic arrays lost or degraded. The newer "high-fidelity" libraries (documonster/excelts, hucre, @office-kit/xlsx, @protobi/exceljs) wrote files **Excel refused to open at all** (even with repair) for this workbook.
2. **ExcelJS `spliceColumns` is unusable for our purpose**: it moves cell values but does not adjust formulas (same sheet or cross-sheet), CF, DV, autoFilter, images, tables, charts or pivots. Shared-formula pointers get corrupted, and `writeBuffer()` then throws. documonster behaves the same way (merges are the only thing it adjusts).
3. **Recommended JS engine: a package-preserving OOXML patcher** built on JSZip (plus our own formula reference shifter). It rewrites only the XML nodes that contain coordinates or references and copies every other part byte-for-byte. I built a prototype (`rawpatch/`) that inserts a column. **Its output matches Excel's own `Columns("C").Insert()` result on all 28 checked feature groups** (formulas 110/110, values, names, CF, DV incl. x14, table, chart, pivot, shapes, notes, threaded comments, hyperlinks, merges, freeze panes, autofilter, print area, styles), and it opens without repair. The only differences are Excel's "format same as left" behaviour: the inserted column inherits the left column's width and CF range. Bundle cost is about 32 kB gzip, against 272 kB for ExcelJS.
4. Use **SheetJS CE for reading values** (already used by the reference app). ExcelJS is optional, for reading styles for the preview. **Never** re-save user workbooks with SheetJS CE: it destroys styles, tables, drawings and CF, and it double-encodes Japanese font names.
5. If ExcelJS must be used for writing anyway, use **exceljs-community 5.2.0** (a maintained drop-in; its output bytes are identical to 4.4.0, so it has the same bugs) and apply the workaround list in §6.

---

## 1. The torture workbook (`torture.xlsx`, built by `build-torture.ps1` through Excel COM)

Sheets: `Data` (tab theme color 6), `集計 シート` (Japanese name with a space, tab RGB), `Tbl`, `Pivot`, `Lists` (hidden). Custom theme: accent1 changed to CC9933, so a replaced theme is detectable.

`Data` contains:
- title `A1` merged `A1:F1`, 16pt bold, fill theme 5 with tint +0.4, font theme color with tint
- header row 3 (height 30): theme fill with tint -0.25, medium red bottom border
- 20 data rows with these number formats: `yyyy/mm/dd`, `m"月"d"日"(aaa)`, `[$¥-411]#,##0`, `0.0%`, `#,##0 "個"`, `#,##0.00`
- thin borders everywhere, wrap and top alignment in J, a rich-text cell (J6), a line break (J7)
- column widths, hidden column K, grouped (outline) columns L:M, custom height on row 10
- freeze panes at C4, autofilter A3:J23
- conditional formats: cellIs, expression `=$F4>0.5` (bold, dark red), 3-color scale, data bar (x14 ext), icon set
- validations: list `=Lists!$A$1:$A$5` (Excel writes this as **x14 extLst**, because it points to another sheet), and whole number 1..100 with an error message
- names: `TaxRate` (workbook scope, points to the Japanese sheet), `AmountCol` = `Data!$E:$E` (whole column), `LocalRange` (sheet scope), Print_Area, Print_Titles, `_FilterDatabase`
- formulas:
  - names in formulas: `=E4*(1+TaxRate)`
  - cross-sheet `='集計 シート'!$B$2*E4`
  - **shared formulas**: `S4:S23` (range assignment) and `T5:T23` (FillDown)
  - legacy CSE array `{=SUM(E4:E23*G4:G23)}`
  - dynamic array `=SORT(UNIQUE(C4:C23))` spilling to P4:P8
  - error results `=1/0` and `=NA()`
  - `=SUM(AmountCol)`, `=SUM(SalesTbl[金額])`, `=COUNTA(Lists!A:A)`
- objects: a chart, a picture, a rectangle shape, a legacy note (B4), a threaded comment (B5), an external hyperlink with a tooltip, an internal hyperlink (`'集計 シート'!A1`)

`Tbl` holds ListObject `SalesTbl` with a calculated column `=[@数量]*[@単価]`, a totals row, and `=SalesTbl[[#Totals],[金額]]`. `集計 シート` holds `SUMIF(Data!C:C,...)`, `VLOOKUP`, `INDEX/MATCH`, `=ROWS(Data!P4#)` (stored as `_xlfn.ANCHORARRAY`) and `=TaxRate*100`. `Pivot` holds PivotTable PT1 over `Data!A3:J23`.

> Note: Excel 365 does **not** write shared formulas for formulas that contain defined names or cross-sheet refs (I4/O4 were written one by one), but does for plain relative formulas (S, T columns). `sharedtest.xlsx` shows the variants: FillDown / AutoFill / `Range.Formula=` on a range / `FormulaR1C1` / Copy all produce `t="shared"`.

## 2. How fidelity was measured

- `roundtrip.mjs`: load then write unchanged for each library → `rt/torture__<lib>.xlsx` (also run on `samples/`, `variants/`).
- `inspect.ps1`: opens each file in **our own** hidden Excel instance and dumps a JSON snapshot to `snap*/`. The snapshot covers:
  - sheets, visibility, tab colors, theme accent1, every Name
  - every cell's `Formula2` and `Value2`, `HasArray`, `HasSpill`
  - FormatConditions (type / AppliesTo / Formula1), `SpecialCells(xlCellTypeAllValidation)`, validation details
  - ListObjects, ChartObjects plus series formula, Shapes, PivotTables, Comments, CommentsThreaded, Hyperlinks
  - merged areas, widths / hidden / outline of columns 1-20, heights of rows 1-12
  - freeze panes, AutoFilter
  - style probes on 16 cells: number format, font, theme color, fill, tint, borders, alignment, wrap
  - rich-text runs
- `compare.js`: diffs each snapshot against the original → `compare.out.md`.
- **Repair detection:** with `DisplayAlerts=$false`, a plain `Workbooks.Open(path,0,$true)` on a file that needs repair throws `Workbooks クラスの Open プロパティを取得できません。` The script then retries with `CorruptLoad=1` (xlRepairFile), then `2` (xlExtractData). Excel writes the repair log to `%TEMP%\error*.xml` (removedRecord/repairedRecord). `opentest.ps1` does only this open test.
  - **Gotcha:** pass real values for every positional argument before CorruptLoad, not `[Reflection.Missing]::Value`. With Missing, the repair open also fails. Working call: `Open($p, 0, $true, 5, '', '', $true, 2, '', $false, $false, $Missing, $false, $false, 1)`.
- Variants with one feature removed (`make-variants.ps1` → `variants/v-*.xlsx`) were used to isolate why a library's output is unopenable.

## 3. Results matrix (full torture workbook)

Legend: OK = identical to the original in Excel. x = lost or changed. n/o = Excel cannot open the file, so it cannot be measured.

| Library (version) | Excel open | Formulas | CSE | Dyn-array spill | CF (5) | DV list (x14) | DV whole | Table | Chart | Shapes | Pivot | Note | Threaded | Hyperlinks | Names | Styles / numFmt | Col / row / freeze / filter / print | Rich text | Theme |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| **JSZip pass-through** (control) | normal | 110/110 | OK | OK | OK | OK | OK | OK | OK | OK | OK | OK | OK | OK | OK | OK | OK | OK | OK |
| **xlsx-populate 1.21.0** | normal | 110/110 | OK | **x** (P5:P8 → `#SPILL!`) | OK | OK | OK | OK | OK | OK | OK | OK | OK | OK | OK | OK | OK | OK | OK |
| exceljs 4.4.0 | **repair** (table autoFilter removed) | 110 (+P5:P8 now CSE) | OK | x (becomes legacy CSE) | OK* | **x** | OK | OK after repair | **x** | x (rectangle, chart) | **x** | OK | **x** (and an empty note) | x (internal link, tooltip) | **x** (AmountCol → `#NAME?`, scope lost) | x (Normal style vAlign; CF italic*) | OK (print area rewritten `$A1:$J23`) | OK | OK |
| exceljs 4.4.0 + table fix | normal | same as above | | | | | | OK | | | | | | | | | | | |
| exceljs-community 5.2.0 | repair | identical bytes to exceljs 4.4.0 → identical results | | | | | | | | | | | | | | | | | |
| @zurmokeeper/exceljs 4.4.9 | repair | as exceljs (but keeps the rectangle shape) | | | | | | | | | | | | | | | | | |
| devextreme-exceljs-fork 4.5.0 | repair | as exceljs | | | | | | | | | | | | | | | | | |
| @protobi/exceljs 4.4.0-protobi.10 | **n/o** (only "extract data") | on `v-notable`: as exceljs but **keeps chart, pivot, rectangle** | | | | | | | | | | | | | | | | | |
| documonster 0.17.1 (= excelts successor) | **n/o** | on `v-nodyn`: formulas 108/108, CSE OK, CF OK*, DV list x, DV whole OK, **table OK, chart OK, pivot OK, threaded OK**, rectangle lost, internal link OK, tooltip lost, AmountCol lost, sheet scope OK, Normal vAlign x | | | | | | | | | | | | | | | | | |
| @cj-tech-master/excelts 10.2.0 | **n/o** | older name of documonster, same failure; needs `installChartSupport()` or it throws on load | | | | | | | | | | | | | | | | | |
| hucre 1.1.0 (`openXlsx`/`saveXlsx`) | **n/o** | on `v-nopivot`: **number formats → General, borders/fills/fonts lost**, row heights lost, rich text lost, DV list lost, chart series name → literal | | | | | | | | | | | | | | | | | |
| @office-kit/xlsx 0.23.0 | **n/o** (even with extract, and even for `v-lite`) | not measurable | | | | | | | | | | | | | | | | | |
| SheetJS CE 0.20.3 | normal | 94/110 (`SalesTbl[...]` → `#REF!`, SORT → `_xlws.SORT` `#NAME?`) | OK | x | **0/5** | x | x | **x** | x | x | x | OK | OK | OK | OK | **x all** (font name mojibake `æ¸¸ã´ã·ãã¯`) | freeze x, row heights x | x | OK |

\* **CF italic bug:** the expression rule's dxf contains `<i val="0"/>`. ExcelJS (and documonster) read that as italic=true, so after a round trip the rule renders **bold+italic**. Verified via COM (`FormatConditions(n).Font.Italic = True`).

What made each "n/o" library fail (found by stripping one feature at a time, `rt-variants/` + `opentest.ps1`):

| Library | Opens after removing… | Cause |
|---|---|---|
| exceljs, forks | table (`v-notable` opens normally) | writes `headerRowCount="0"` + `hiddenButton="1"` filter columns (reader bug, see §6-1) |
| protobi | table | same table bug, and repair cannot fix it for protobi |
| documonster / excelts | dynamic array (`v-nodyn`) | rewrites `xl/metadata.xml` (adds `adjust="1"`); swapping in only that part makes the original unopenable |
| hucre | pivot (`v-nopivot`) | pivot part rewrite |
| office-kit | nothing (`v-lite` fails too) | sheet XML / sharedStrings (alpha quality) |

**Simple real-world files** (the reference app's 4 `input/*.xlsx`, copied as `samples/sample1..4.xlsx`): every library's round trip opened normally. The damage only appears with "real Excel" features (tables, dynamic arrays, pivots, x14 ext, threaded comments), which are exactly what users of this tool will drop in.

## 4. Structural edits: what the libraries do (critical for requirement 9)

`exceljs-splice.js`: `ws.spliceColumns(3, 0, [])` (insert at C) on `Data`.

| Item | Result |
|---|---|
| I4 (moves to J4) | formula stays `E4*(1+TaxRate)`: **not adjusted** (now points to the wrong column) |
| Shared formula dependents | point at the wrong master. S5 shows `SUM(LocalRange)`, U6 shows `SUM(AmountCol)` (garbage). `writeBuffer()` then **throws** `Shared Formula master must exist above and or left of clone for cell U6` |
| Cross-sheet formula `集計 シート!B3 = SUM(Data!E4:E23)` | not adjusted |
| merges, autoFilter, CF sqref + formulas, DV keys, image anchors | not moved |
| `_FilterDatabase` defined name | split into two areas (`Data!$A$3:$B$23`, `Data!$D$3:$K$23`) |
| Column widths | shifted, but the new column has no width |

documonster `Worksheet.spliceColumns(ws,3,0,[])` gives the same results. It moves merges (A1:F1 → A1:G1) and nothing else. Its DV model is range-keyed (`"range:G4:G23"`), unlike ExcelJS.

**Conclusion: whichever library we pick, the reference-adjusting logic must be our own.** Once we own it, the library's only job is "load and save without losing anything". JSZip does that perfectly.

## 5. Prototype: package-preserving patcher (recommended engine)

Files: `rawpatch/formula.js` (reference tokenizer / shifter / translator) and `rawpatch/insert-cols.js` (package patcher).
Run: `node rawpatch/insert-cols.js torture.xlsx out.xlsx Data C 1`

What it patches (string/regex over the XML; untouched parts keep their bytes):

- **Every worksheet**
  1. **Unshare shared formulas.** Each dependent `<f t="shared" si=".."/>` becomes an explicit `<f>` using `translateFormula(masterFormula, dRow, dCol)` (only relative parts move).
  2. Shift references in `<f>` text (unqualified refs count as the home sheet), in CF `<formula>`, in DV `<formula1/2>`, in x14 `<xm:f>`, and in `<hyperlink location=…>`.
- **Target sheet only**, additionally:
  - `<c r>`, `<f ref>` (array / spill ranges), `<dimension>`, `<mergeCell>`, `<autoFilter>`, `<hyperlink ref>`
  - `sqref` / `topLeftCell` / `activeCell`, `<xm:sqref>`
  - `<col min/max>` (split if the insert falls inside a span)
  - drop `<row spans>`
- **Tables** on the target sheet:
  - `ref` and autoFilter `ref`
  - when inserting *inside* the table: add `<tableColumn id=max+1 name="列1">` **and write the header cell text `列1`**. Without the header text, Excel "repairs" the table.
  - `calculatedColumnFormula` refs
- **Drawings:** `<xdr:from|to><xdr:col>` (0-based). **Charts** (all): `<c:f>`.
- **Pivot caches:** `<worksheetSource sheet="Data" ref>`, plus `refreshOnLoad="1"`.
- **Notes, threaded comments, VML:** comments `ref`, threadedComments `ref`, VML `<x:Anchor>` columns 1 and 5, and `<x:Column>`.
- **Workbook:** `<definedName>` text. Delete `xl/calcChain.xml` together with its rel and content-type override. Add `fullCalcOnLoad="1"` to `<calcPr>`.

Verification against Excel itself (`golden-insert.ps1` saves Excel's own insert as the golden file; `snap-ins/`):

```
insert column C on Data : rawpatch opens normal; formulas 110/110, values 412/412, names, CF formulas,
                          DV (incl. x14), table, chart series (Data!$F$4:$F$13), pivot source (A3:K23),
                          shapes/anchors, notes, threaded, hyperlinks, merges (A1:G1), freeze, filter,
                          print area, styles: ALL identical to Excel.
                          Differences = Excel's "format like left column": new col C gets width 14 and
                          CF B4:B23 grows to B4:C23.
insert column C inside SalesTbl (Tbl): identical to Excel on every checked item, opens normal
                          (after adding the header-cell fix).
ExcelJS spliceColumns   : output could not even be written (shared-formula exception).
```

Performance in the browser (Chromium; the Claude browser pane, bundle `vitetest/src/lib-rawxml.ts`, a DOMParser variant that shifts every cell of the 10k×40 sheet):

| Step | Time |
|---|---|
| unzip | 7 ms |
| inflate sheet | 366 ms |
| DOMParser | 1.5 s |
| shift 400k `<c>` + formulas | 2.0 s |
| XMLSerializer | 1.0 s |
| zip | 1.1 s |
| **total** | **about 6.0 s** |

A regex/string pass (as in `insert-cols.js`) avoids the DOM and is faster. Only the target sheet is large; other parts are small.

Known limits of the prototype, for the implementers to finish:
- Delete-column semantics: refs fully inside the deleted range → `#REF!`; ranges shrink. Excel golden files should be generated the same way.
- Move = cut + insert, where refs *to* the moved cells follow them.
- 3D refs (`Sheet1:Sheet3!A1`), external refs `[1]Sheet!A1`, `colBreaks`, sparklines (`x14:sparklineGroups` `<xm:f>`/`<xm:sqref>` – add to the regex list), slicers/timelines, form controls (`x14:formControlPr` + `ctrlProps`), data-table `<f t="dataTable">`, `ignoredErrors sqref`, `protectedRanges`, `<selection pane>` of split panes.
- Excel's VLOOKUP `col_index_num` is not adjusted (Excel doesn't either). Warn the user when an insert/move lands inside a lookup range.
- Pivot caches whose source is a *defined name* or table (worksheetSource `name=`) need no change. Inserting inside a pivot source range needs `refreshOnLoad` (already set).

## 6. ExcelJS: exact APIs and verified gotchas (`exceljs-api.js`, `exceljs-api.out.txt`, `exceljs-table.js`, `exceljs-splice.js`, `exceljs-style-hazard.js`, `dv-wholecol.js`)

All of these were run on exceljs 4.4.0. exceljs-community 5.2.0 behaves identically (same `boolean-xform.js`, same bytes out).

**Reading formulas.** `cell.type === ValueType.Formula (6)`. `cell.formulaType`: 1 = master/normal, 2 = shared dependent.

```js
ws.getCell('S4').value   // {formula:'E4*G4', result:12000, shareType:'shared', ref:'S4:S23'}   (master)
ws.getCell('S5').value   // {sharedFormula:'S4', result:228494}                             (dependent)
ws.getCell('S5').formula // 'E5*G5'  <- getter translates the master for you (verified)
ws.getCell('N4').value   // {formula:'SUM(E4:E23*G4:G23)', result:46133461, shareType:'array', ref:'N4'}
ws.getCell('P4').value   // {formula:'_xlfn._xlws.SORT(_xlfn.UNIQUE(C4:C23))', result:'SFC', shareType:'array', ref:'P4:P8'}
ws.getCell('P5').value   // '京大'  (spilled cells are plain values; cm="1" is NOT modelled -> written back as CSE)
ws.getCell('Q4').value   // {formula:'1/0', result:{error:'#DIV/0!'}}
wb.getWorksheet('集計 シート').getCell('B8').value // {formula:'ROWS(_xlfn.ANCHORARRAY(Data!P4))', result:5}  (= P4#)
```

- The shared master need not be the first formula of the column: T4 is a plain formula, T5 is the master of T5:T23.
- Formula text has no leading `=` and keeps the `_xlfn.` / `_xlws.` prefixes. Keep them when writing, or Excel shows `#NAME?`.
- **Unshare before any edit** (verified; 39 cells unshared, the output then has 0 `t="shared"`):
  ```js
  ws.eachRow(row => row.eachCell(cell => {
    if (cell.type !== ExcelJS.ValueType.Formula) return;
    const v = cell.value;
    if (v.sharedFormula || v.shareType === 'shared') cell.value = { formula: cell.formula, result: cell.result };
  }));
  ```
- `wb.calcProperties.fullCalcOnLoad = true` writes `<calcPr calcId="171027" fullCalcOnLoad="1"/>` (verified). By default `calcProperties` is `{}` and calcChain is dropped.

**Values:**
- Dates: date-formatted numbers come back as a JS `Date` at **UTC midnight** (`2025-12-09T00:00:00.000Z`). The epoch is `wb.properties.date1904`; the reference app's `xlsx-util.ts` already handles the 1462 offset.
- Rich text: `{richText:[{text:'重要:'}, {text:' 要確認のメモ', font:{...}}]}`. A run without `font` inherits the cell font.
- External hyperlink: `{text:'Example', hyperlink:'https://...'}`. **The tooltip is dropped. Internal links (`location=`) are dropped**: J5 reads as a plain string.
- Notes: `cell.note = {texts:[{font, text}]}`. The legacy fallback note of a threaded comment reads as `{texts:[]}` and is written back as an **empty note**. Threaded comments and persons are lost.

**Styles.**
- `cell.style` is `{font, border, fill, alignment, numFmt, protection}`. Colors are `{argb}` or `{theme, tint}`. Theme index = XML `theme` attribute (0-based; COM ThemeColor = xml+1).
- **Mutation hazard (verified):** after load, all cells with the same `s` index share **one style object**. `B6.font = {...B6.font, bold:true}` also bolds B7 and C6. The same happens with `.fill =`, and with `Z1.style = B7.style` then editing Z1. Safe pattern:
  ```js
  cell.style = structuredClone(cell.style); // first
  cell.font = { ...cell.font, bold: true };
  ```
- **Normal style damaged:** `cellStyleXfs` / `xf[0]` lose `<alignment vertical="center"/>`, `標準` is renamed `Normal`, and other named cell styles (ハイパーリンク…) are dropped. Every unstyled cell in a Japanese workbook changes from vertically centred to bottom.
- `BooleanXform` ignores `val`: `<b val="0"/>`, `<i val="0"/>`, `<strike val="0"/>` are read as **true** (root cause of the CF italic bug; it also affects fonts and rich-text runs). Fix in a patch: `this.model = !['0','false'].includes(node.attributes.val)`.

**Columns and rows:**
- `ws.getColumn(n)`: `.width` (e.g. 8.625), `.hidden`, `.outlineLevel`, `.style`. A hidden column reads width 0 and hidden true.
- Grouped columns are written back with `collapsed="1"`.
- `ws.getRow(n)`: `.height`, `.hidden`, `.outlineLevel`.

**Merges:** `ws.model.merges` gives `['A1:F1']`. `ws._merges` is `{A1: Range}`. `cell.isMerged`, `cell.master.address`.

**Conditional formatting:** `ws.conditionalFormattings` is `[{ref:'B4:B23', rules:[{type:'expression', priority, formulae:['$F4>0.5'], style:{font,...}}]}]`.
- cellIs rules add `operator`; colorScale / dataBar / iconSet rules have `cfvo` and `color`.
- The dataBar keeps its x14 ext (`x14Id` is regenerated).
- The iconSet name is kept (`'3TrafficLights'`).

**Data validation:**
- `ws.dataValidations.model` is keyed by **single cell address**: sqref G4:G23 becomes 20 keys.
- A whole-column DV (`G1:G1048576`) becomes **1,048,576 keys**, adding about 250 MB heap and 4 s of write time *per rule* (verified in `dv-wholecol.js`).
- On write, the sqrefs overlap (`G10:G23` **and** `G4:G23`).
- **x14 extLst validations are dropped.** Excel always writes a list DV that refers to another sheet this way.

**Defined names:**
- `wb.definedNames.model` is `[{name:'TaxRate', ranges:["'集計 シート'!$B$2"]}, ...]`.
- **Whole-column/row names are silently dropped** (`AmountCol` → `=SUM(AmountCol)` becomes `#NAME?`).
- `localSheetId` is lost: sheet-scoped names become global.
- Print_Area and Print_Titles move to `ws.pageSetup.printArea = 'A1:J23'` and `printTitlesRow = '3:3'`, and are written back as `'Data'!$A1:$J23` (mixed absolute/relative).

**Tables:**
- `ws.tables.SalesTbl.table` is `{name, displayName, tableRef:'A1:D12', totalsRow:true, headerRow:false (BUG), columns:[{name, totalsRowLabel, totalsRowFunction}], autoFilterRef:'A1:D11', style:{theme:'TableStyleMedium2',...}}`.
- `ws.getTables()` returns the array.
- **Reader bug:** `headerRow: attributes.headerRowCount === '1'`, but an absent attribute means 1. The writer then emits `headerRowCount="0"` and `hiddenButton="1"` filter columns, so Excel repairs and removes the table autoFilter. Fix after load (verified; the file then opens normally):
  ```js
  for (const w of wb.worksheets) for (const t of Object.values(w.tables)) {
    t.table.headerRow = true;                       // unless the source XML really had headerRowCount="0"
    for (const c of t.table.columns) if (c.filterButton === undefined) c.filterButton = true;
  }
  ```
- `calculatedColumnFormula` and `dataDxfId` are lost.

**Images:**
- `ws.getImages()` is `[{type:'image', imageId:'0', range:{tl:{nativeCol:8,nativeRow:0,...}, br:{...}, editAs:'oneCell'}}]`.
- `wb.getImage(id)` / `wb.model.media` give `{type:'image', name, extension:'png', buffer}`.
- Non-picture shapes and charts are lost (zurmokeeper keeps simple shapes; protobi keeps charts).

**Sheet settings:**
- autoFilter: `ws.autoFilter === 'A3:J23'`.
- Views: `ws.views = [{state:'frozen', xSplit:2, ySplit:3, topLeftCell:'C4', ...}]`. The selection/activeCell is lost.
- Hidden sheets: `ws.state` is `'hidden'` or `'veryHidden'`.
- Tab color: `ws.properties.tabColor` is `{theme:5}` or `{argb}`.

**Theme:** `wb._themes.theme1` holds the raw XML string, written back verbatim, so the custom accent survives.

**Pivot tables:** not read (`pivotTables` is 0) and dropped on write.

## 7. Performance: 10,000 rows × 40 styled columns (`perf.mjs`, file 1.77 MB)

Node 24, fresh process per library:

| Library | Load | Write | Peak heap / rss |
|---|---|---|---|
| exceljs 4.4.0 | 2.5–2.9 s | 2.5–2.7 s | ~480 MB / 760 MB |
| exceljs-community 5.2.0 | 2.1 s | 3.1 s | 515 MB / 750 MB |
| xlsx-populate | 4.4 s | 2.8 s | 700 MB / 930 MB |
| documonster | 1.7 s | 4.9 s | **950 MB** / 1.2 GB |
| hucre | 1.3 s | 0.7 s | 144 MB |
| office-kit | 2.5 s | 1.3 s | 118 MB |
| SheetJS | 3.3 s | 1.5 s | 250 MB |

Generating the file with ExcelJS (styles + write) took 6.7 s.

Browser (Chromium in the Claude browser pane, Vite single-file bundles, load + write):

| Engine | Time |
|---|---|
| exceljs | 9.2 s (266 MB heap) |
| xlsx-populate UMD | 8.8 s |
| documonster | 5.0 s |
| hucre | 1.6 s |
| raw DOM patch | 6.0 s |

**Always run export in a Web Worker** (or in the Node server), with a progress indicator.

## 8. Browser / Vite feasibility (`vitetest/`, Vite 6 + vite-plugin-singlefile, the same config as the reference app)

| Import | Single-file `index.html` | Works in browser |
|---|---|---|
| `import ExcelJS from 'exceljs'` | 941 kB (272 kB gz) | yes. Vite picks the `"browser": "./dist/exceljs.min.js"` field automatically; output identical to Node |
| `import ExcelJS from 'exceljs-community'` | 966 kB (280 kB gz) | yes (`exports.browser` → dist) |
| `import XlsxPopulate from 'xlsx-populate'` | 323 kB | **no**: `ReferenceError: Buffer is not defined` |
| `import XP from 'xlsx-populate/browser/xlsx-populate-no-encryption.min.js'` | 358 kB (111 kB gz) | yes |
| `import { Workbook } from 'documonster/excel'` | 984 kB (261 kB gz) | yes |
| `import { openXlsx, saveXlsx } from 'hucre/xlsx'` | 291 kB (81 kB gz) | yes |
| `jszip` only (patcher) | **100 kB (32 kB gz)** | yes |

How the reference app does it:
- `src/core/export.ts` and `convert.ts` use `import ExcelJS from 'exceljs'` (default import, write path).
- `src/core/xlsx-util.ts` uses `import type {...} from 'exceljs'` only, so the read path stays cheap.
- `parse.ts` reads with SheetJS (`import * as XLSX from 'xlsx'`, installed from the CDN tgz).
- There is no alias in `vite.config.ts`, and `dist/index.html` is 1.7 MB with ExcelJS, SheetJS and React inlined.
- `vitest.config.ts` uses `environment:'node'`, which resolves the Node entry. That is fine for exceljs 4.4.0 (CJS). **exceljs-community 5.x** Node entry needs Node ≥ 22.12 (archiver@8 ESM). For Vitest, alias it to `exceljs-community/dist/exceljs.min.js` (per its README).
- ExcelJS `writeBuffer()` returns a Buffer in Node and an ArrayBuffer in the browser. The reference app's `excelJsOutputToUint8Array()` already normalises this.
- `ModelContainer` and `stream.*` are undefined in the browser bundle.

## 9. Library verdicts

| Library | Verdict |
|---|---|
| JSZip 3.10.1 | **use.** Foundation of the export engine (byte-preserving). fflate is a faster drop-in alternative if needed |
| SheetJS CE 0.20.3 | **use for reading only** (values, formulas, .xls/.csv inputs, fast). **Reject for writing** |
| exceljs-community 5.2.0 | **use only if needed**: read styles for the GUI preview, or build brand-new workbooks. Maintained drop-in for 4.4.0 with identical fidelity. Never round-trip user workbooks through it without the §6 fixes plus XML post-patching |
| exceljs 4.4.0 | reference-only (unmaintained, last push 2025-01, 809 open issues). Prefer the community fork |
| xlsx-populate 1.21.0 | reference-only / fallback. Best round-trip fidelity of any library, but it drops formula cached values, rewrites spilled cells as shared strings (`#SPILL!`), has no structural operations, is unmaintained, and in the browser needs the prebuilt UMD |
| documonster 0.17.1 / @cj-tech-master/excelts 10.2.0 | reject for now, watch. Best preservation in the ExcelJS family (charts, pivots, threaded comments, names scope) and actively developed, but its output with any dynamic-array formula is **unopenable** in Excel. Also inherits the italic, Normal-style and whole-column-name bugs; 950 MB heap on 10k×40 |
| @protobi/exceljs | reject (table bug makes the file unopenable; otherwise ExcelJS plus charts and pivots) |
| @zurmokeeper/exceljs, devextreme-exceljs-fork, other forks | reject (same as 4.4.0) |
| hucre 1.1.0 | reject (drops number formats, borders and fills; pivot output unopenable) |
| @office-kit/xlsx 0.23.0 | reject (output unopenable) |

## 10. Recommendation for the implementers

- **JS export engine = OOXML patcher over the original package** (JSZip + `rawpatch/`-style rewriting). Base it on the primary input workbook, and model the user's operations as a list of structural ops (insert/delete/move columns, append rows). Apply each op with one reference shifter across every part listed in §5. Verify with the Excel-golden method (`golden-insert.ps1` + `inspect.ps1` + `compare.js`). The same op list is what the planned "Excel本体" engine replays through COM, so both engines share one op model and one golden-test harness.
- **New cells** (merged-in columns from other files):
  - Write `<c>` nodes with `t="inlineStr"` (or append to sharedStrings), numbers, and formulas with `fullCalcOnLoad`.
  - For styles, append font/fill/border/numFmt/xf entries to the base `styles.xml` with dedupe. That is a small style-merger; ExcelJS style objects are a convenient in-memory shape for it.
- Always unshare shared formulas in sheets being edited. Always delete calcChain and set `fullCalcOnLoad`. Keep `_xlfn.` prefixes. Keep `cm="1"`, `metadata.xml` and `t="str"` on spilled cells.
- Run in a Worker. Keep ExcelJS lazy-loaded (dynamic `import()`), because the singlefile build inlines it anyway.

## 11. Environment pitfalls met while doing this (for whoever automates Excel next)

- In this PowerShell 5.1 + Excel setup, assigning `Range.Value2 = <Int32>` or `= object[,]` failed (`Unable to cast Int32 to String`, then `0x800A03EC`). Setting **`Range.Formula = [string]`** per cell works.
- The default printer (Brother, offline) makes `PageSetup.*` setters fail and **getters hang for minutes**. Avoid PageSetup; read print areas from the `Print_Area` / `Print_Titles` Names.
- `Get-Content -Raw` strings carry PS* NoteProperties. `ConvertTo-Json` on them exploded to 1 GB (OOM). Use `[IO.File]::ReadAllText`.
- `Split-Path -LeafBase` does not exist in PS 5.1.
- A Git-Bash `tail -f` on a log file locks it against `Add-Content`. Log with `[IO.File]::AppendAllText` inside try/catch.
- Record the PID of your own Excel instance (`GetWindowThreadProcessId(Application.Hwnd)`) in a file, so a hung run can be cleaned up without touching other Excel processes. Other agents' and the user's Excel instances were running during this work.
- Excel repair logs go to `%TEMP%\errorNNNNNN_NN.xml`.

## 12. Artifacts

| Path (under r1\) | What |
|---|---|
| `torture.xlsx`, `build-torture.ps1`, `com-common.ps1` | torture fixture + builder + safe COM helper |
| `variants\v-*.xlsx`, `make-variants.ps1` | one-feature-removed fixtures |
| `samples\sample1-4.xlsx` | copies of the reference app inputs |
| `roundtrip.mjs`, `rt\`, `rt-variants\`, `rt-samples\`, `rt-fix\` | round-trip runner + outputs |
| `inspect.ps1`, `compare.js`, `snap\`, `snap-v\`, `snap-ins\`, `compare.out.md` | Excel-COM snapshotter, differ, results |
| `opentest.ps1`, `bisect.js`, `bisect\`, `diag\` | open/repair tester, part bisection |
| `exceljs-api.js` (+`.out.txt`), `exceljs-table.js`, `exceljs-splice.js`, `exceljs-style-hazard.js`, `dv-wholecol.js`, `xlsx-populate-api.js`, `documonster-splice.mjs` | API probes |
| `rawpatch\formula.js`, `rawpatch\insert-cols.js`, `golden-insert.ps1`, `ins\` | recommended-engine prototype + Excel golden files |
| `perf.mjs`, `perf_10000x40.xlsx` | performance test |
| `vitetest\` (vite.config.ts, src\lib-*.ts, serve.mjs, dist-*) | browser bundle tests |
