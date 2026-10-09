/**
 * doc facade implementation (owner: doc D). Contract: see ./index.ts and src/core/contracts.ts.
 *
 * Every operation returns a NEW EditorSnapshot sharing all unchanged parts with the input (undo/redo keeps whole
 * snapshots), appends what it did to the op log (the same log is replayed by the JS package patcher and by the
 * Excel COM runner), or refuses with a Japanese reason. A no-op returns the input snapshot itself.
 *
 * Module map:
 *   create.ts   createEditor (column-major doc from the premerged book)
 *   struct.ts   Excel-semantics structural ops over the whole doc (+ refusals)
 *   compiled.ts pre-tokenized formulas (fast path of excelRef.adjustFormula)
 *   deps.ts     #REF! dry runs, VLOOKUP-style column-number warnings
 *   delete.ts   analyzeDelete / deleteColumns (+ convert to values)
 *   edit.ts     setCellInput (+ totals-row labels, quote flag) / setColumnWidth (customWidth) / formulaForDisplay
 *   autowiden.ts Excel's auto-widening of auto-width columns for typed dates / numbers
 *   parse.ts    Excel-like input parsing (+ needsQuotePrefix)
 *   tables.ts   ListObject helpers
 *   unmerge.ts  title merges blocking a move: detection + the 'unmerge' op (centerContinuous + the top-left cell's
 *               fill / border sides spread over the former area)
 */
import { planMove, MAX_COL } from '../formula/excelRef';
import type { StructOp } from '../formula/excelRef';
import type { EditorSnapshot, LoggedOp, MergedDoc, SourceBook } from '../types';
import type { DeleteImpact, OpResult, PremergeLayout } from '../contracts';
import { buildEditor } from './create';
import { analyzeDeleteImpl, deleteColumnsImpl, indicesOf } from './delete';
import { lookupIssues, lookupWarningTexts } from './deps';
import { formulaForDisplayImpl, setCellInputImpl, setColumnWidthImpl } from './edit';
import { applyStructOps, collapseWarnings } from './struct';
import { titleBlock, unmergeInDoc, unmergeQuestion } from './unmerge';
import { boxText } from './util';

export function createEditor(
  premerged: SourceBook,
  layout: PremergeLayout,
  sourceInfo: MergedDoc['sourceInfo'],
  region: { headerRow: number; firstDataRow: number; lastDataRow: number },
): EditorSnapshot {
  return buildEditor(premerged, layout, sourceInfo, region);
}

const structLog = (doc: MergedDoc, ops: readonly StructOp[]) => ops.map((op) => ({ kind: 'struct' as const, sheet: doc.sheetName, op }));

/**
 * Move the columns `ids` (any selection, in any order) so that they sit, in their current order, at slot `gap`.
 * When the only obstacles are title merges (merged areas entirely above the header row) the refusal carries
 * fix {kind:'unmergeTitles'}; with opts.unmergeTitles those merges are first unmerged keeping their look
 * (LoggedOp 'unmerge', centerAcross) and then the columns are moved — one snapshot, one undo step.
 */
export function moveColumns(s: EditorSnapshot, ids: string[], gap: number, opts?: { unmergeTitles?: boolean }): OpResult {
  const doc = s.doc;
  if (!Array.isArray(ids) || !ids.length) return { ok: true, snapshot: s, warnings: [] };
  const idx = indicesOf(doc, ids);
  if (typeof idx === 'string') return { ok: false, reason: idx };
  const n = doc.columns.length;
  if (!Number.isInteger(gap) || gap < 0 || gap > n) return { ok: false, reason: '移動先の位置が正しくありません。' };
  const ops = planMove(idx.map((i) => i + 1), gap + 1, 'col');
  if (!ops.length) return { ok: true, snapshot: s, warnings: [] };
  let base = doc;
  const pre: LoggedOp[] = [];
  const warnings: string[] = [];
  const block = titleBlock(doc, ops);
  if (block.kind === 'other') return { ok: false, reason: block.reason };
  if (block.kind === 'titles') {
    const refs = block.boxes.map(boxText);
    if (!opts?.unmergeTitles) return { ok: false, reason: block.reason, fix: { kind: 'unmergeTitles', refs, message: unmergeQuestion(block.boxes) } };
    base = unmergeInDoc(doc, block.boxes, true);
    pre.push({ kind: 'unmerge', sheet: doc.sheetName, refs, centerAcross: true });
    warnings.push(`タイトル行の結合（${refs.join('、')}）を解除し、「選択範囲内で中央」にしました。`);
  }
  const r = applyStructOps(base, ops);
  if (!r.ok) return { ok: false, reason: r.reason };
  warnings.push(...lookupWarningTexts(lookupIssues(base, ops), '移動'));
  warnings.push(...collapseWarnings(r.collapsed));
  if (r.refErrors > 0) warnings.push(`${r.refErrors.toLocaleString()} 個の数式が #REF! になりました。`);
  const label = (pre.length ? '結合を解除して' : '') + (idx.length === 1 ? '列を移動' : `${idx.length} 列を移動`);
  return { ok: true, snapshot: { doc: r.doc, log: [...s.log, ...pre, ...structLog(doc, ops)], label }, warnings };
}

/** Next free ids 'n<k>' for user-inserted columns. */
function newIds(doc: MergedDoc, count: number): string[] {
  let max = 0;
  for (const c of doc.columns) {
    const m = /^n(\d+)$/.exec(c.id);
    if (m) max = Math.max(max, Number(m[1]));
  }
  const out: string[] = [];
  for (let k = 1; k <= count; k++) out.push(`n${max + k}`);
  return out;
}

/** Insert `count` blank columns before 0-based position `index` (Excel "format from left"). */
export function insertColumns(s: EditorSnapshot, index: number, count: number): OpResult {
  const doc = s.doc;
  const n = doc.columns.length;
  if (!Number.isInteger(index) || index < 0 || index > n) return { ok: false, reason: '挿入する位置が正しくありません。' };
  if (!Number.isInteger(count) || count < 1) return { ok: false, reason: '挿入する列数が正しくありません。' };
  if (n + count > MAX_COL) return { ok: false, reason: `列が多すぎるため挿入できません（Excel の上限は ${MAX_COL.toLocaleString()} 列です）。` };
  const ops: StructOp[] = [{ type: 'insert', axis: 'col', at: index + 1, count }];
  const r = applyStructOps(doc, ops, newIds(doc, count));
  if (!r.ok) return { ok: false, reason: r.reason };
  const warnings = [...lookupWarningTexts(lookupIssues(doc, ops), '挿入'), ...collapseWarnings(r.collapsed)];
  const label = count === 1 ? '列を挿入' : `${count} 列を挿入`;
  return { ok: true, snapshot: { doc: r.doc, log: [...s.log, ...structLog(doc, ops)], label }, warnings };
}

export function analyzeDelete(s: EditorSnapshot, ids: string[]): DeleteImpact {
  return analyzeDeleteImpl(s, ids);
}

export function deleteColumns(s: EditorSnapshot, ids: string[], opts: { convertDependents: boolean }): OpResult {
  return deleteColumnsImpl(s, ids, opts ?? { convertDependents: false });
}

export function setCellInput(s: EditorSnapshot, colId: string, row: number, text: string): OpResult {
  return setCellInputImpl(s, colId, row, text);
}

/** width in the STORED <col width> unit (ColProps.width), not the character count Excel's UI shows */
export function setColumnWidth(s: EditorSnapshot, colId: string, width: number): OpResult {
  return setColumnWidthImpl(s, colId, width);
}

export function formulaForDisplay(doc: MergedDoc, colIndex: number, row: number): string | null {
  return formulaForDisplayImpl(doc, colIndex, row);
}
