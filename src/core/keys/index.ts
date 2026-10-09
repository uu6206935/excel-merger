/**
 * keys facade: data-sheet / header detection, key candidates, join.
 * Owner: keys module (normalize.ts, detect.ts, match.ts, api.ts ported from research R6, plus multi.ts).
 *
 *  pickDataSheet(book: SourceBook): string
 *    - the sheet that most looks like a list (dictionary hits in a header row, many rows).
 *
 *  detectRegion(book: SourceBook, sheet: string, opts?: { headerRow?: number; lastDataRow?: number }): TableRegion
 *    - header row (dictionary hits first, not densest row), first/last data row, first/last column.
 *      The table ends at trailing summary rows (合計/小計/計/人数/平均/件数 … labels and/or aggregate formulas such
 *      as SUM/COUNTA/COUNTIF/AVERAGE over the column above), also when they sit below blank rows.
 *    - opts.headerRow: the user picked the header row → columns from that row, data below it (same rules);
 *      opts.lastDataRow: the user picked the last data row (clamped to the sheet).
 *    - KeyAnalysis.internal holds closures/class instances: analyzeSources and joinSources must run in the same
 *      thread (not structured-cloneable).
 *
 *  analyzeSources(base: KeySource, others: KeySource[], options?: Partial<NormalizeOptions>): KeyAnalysis
 *    - pairwise analysis base <-> each other (R6 analyzeKeys) + main-screen chips (ID / 氏名).
 *
 *  addManualCandidate(analysis, sourceId, baseCols: number[], otherCols: number[]): KeyAnalysis
 *    - the user picked columns themselves (detail modal 「列を選んで追加」); returns a new analysis whose pair for
 *      sourceId has one more candidate with real match statistics (selected:false; the UI then adds its id to plan.tiers).
 *      Throws an Error with a Japanese message for invalid columns.
 *
 *  defaultPlan(analysis: KeyAnalysis): JoinPlan
 *    - tiers = auto-selected candidates per pair, joinType 'left', duplicatePolicy 'first'.
 *
 *  joinSources(analysis: KeyAnalysis, plan: JoinPlan): JoinResult
 *    - rows in base order (+ appended rows for 'full'; rows appended from one source can be matched by later sources),
 *      per-pair reports (unmatched, duplicates, ambiguous, conflicts, suggestions), check items, keyFill, defaultImports.
 */
export { pickDataSheet, detectRegion, analyzeSources, defaultPlan, joinSources, addManualCandidate } from './multi';
