/**
 * replayLog (owner: patcher P): replays the user's op log on the premerged package with Excel semantics.
 * See ./index.ts and docs/ARCHITECTURE.md §2. The heavy lifting lives in ./patch/engine.ts.
 *
 * Op encodings as the patcher reads them (types.ts LoggedOp):
 *  - struct   : excelRef StructOp on `sheet` (positions as of that point of the log). Table rules beyond excelRef (Excel-measured,
 *               goldens h2-del-*): a row delete through a table's totals row removes the totals row (totalsRowCount, totals
 *               functions / labels / formulas, filter range); a data row may only be reordered between the header and the end of
 *               the table (never at / above the header, the header / totals row itself never moves) — else a Japanese PatchError;
 *  - cells    : CellWrite[] with reason 'edit' | 'toValues' (see Engine.writeCells for the Excel rules: arrays, spill
 *               masters, table header renames, totals-row entries, numFmt, dynamic arrays). `quote` is honoured exactly:
 *               quote === true on a text constant gives the cell quotePrefix="1", every other write adds none (a
 *               prefixed cell overwritten by a non-quote write loses it — Excel through COM keeps the flag, an invisible
 *               difference, see Engine.setCell; a clear keeps the style). A text constant containing a line feed (LF,
 *               CR LF) gets wrapText on its xf like Excel does by itself (a lone CR does not in Excel; formulas / numbers
 *               never) — golden q-wrap. A value / formula written INTO the spill range of a dynamic array (not at its master)
 *               collapses that spill (master ref = its own cell, cached spilled values cleared; Excel shows #SPILL! and re-spills on
 *               load) — golden r-spill-write;
 *  - colWidth : `width` is the value STORED in <col width> (Excel character units INCLUDING the ~5 px padding, the
 *               same unit as ColProps.width / what the reader returns). It is written verbatim with customWidth="1" (and
 *               ends an auto-fit: bestFit is dropped, as Excel does); the Excel runner converts it to Range.ColumnWidth
 *               (stored width - padding);
 *  - unmerge  : merged areas removed. ALWAYS: the top-left cell's fill goes to every other cell of the former area and each
 *               border side it has to the cells on that edge (top: first row, bottom: last row, left: first column, right:
 *               last column); centerAcross also gives every cell horizontal="centerContinuous".
 *
 * Call-out formulas (round 5): a formula logged by a cells op that would make Excel reach outside the workbook by itself
 * (`isUnsafeFormula`, src/core/export/comJob.ts: DDE `|`, [book.xlsx] / UNC / URL references, WEBSERVICE / IMAGE / RTD / CALL / REGISTER /
 * EXEC) is written as TEXT with the quote prefix — never as a formula — exactly like the Excel runner (UNSAFE_FORMULA); a text constant
 * that Excel would parse as such a formula gets the quote prefix. `onWarning` (optional, additive) receives one UNSAFE_FORMULA warning
 * per cell (at most 20 named + one summary line, like the runner's warn lines).
 */
import type { LoggedOp } from '../types';
import type { ProgressFn } from '../contracts';
import { Engine } from './patch/engine';
import { PatchError, toPatchError } from './patch/pkg';

/** A non-fatal note of the replay (same code / wording as the Excel runner's warn lines). */
export interface ReplayWarning {
  code: 'UNSAFE_FORMULA';
  message: string;
}

/**
 * Failure contract (hostile / damaged files included): the result is a correct package or a PatchError whose message
 * is Japanese. An op that cannot be applied says which one ("3 番目の操作を反映できませんでした: …"); anything the
 * patcher did not expect (a TypeError from a strange file ...) is converted to a PatchError too (`internal` = true, the
 * original error in `cause`). XML parts the patcher wrote are checked for well-formedness before the package is built.
 */
export async function replayLog(premerged: Uint8Array, log: LoggedOp[], onProgress?: ProgressFn, onWarning?: (w: ReplayWarning) => void): Promise<Uint8Array> {
  const report = (ratio: number | null, message: string) => { try { onProgress?.({ ratio, message }); } catch { /* UI callback errors never abort the export */ } };
  try {
    report(0, 'ファイルを読み込み中…');
    const eng = await Engine.open(premerged);
    const ops = Array.isArray(log) ? log : [];
    const n = ops.length;
    for (let i = 0; i < n; i++) {
      const op = ops[i];
      report(0.05 + (0.75 * i) / Math.max(1, n), `操作を反映中…（${i + 1}/${n}）`);
      try {
        if (!op || typeof op !== 'object') throw new PatchError('不明な操作が含まれています。');
        if (op.kind !== 'struct' && op.kind !== 'cells' && op.kind !== 'colWidth' && op.kind !== 'unmerge') throw new PatchError('不明な操作が含まれています。');
        if (typeof op.sheet !== 'string') throw new PatchError('操作の対象シートが指定されていません。');
        if (op.kind === 'struct') await eng.struct(op.sheet, op.op);
        else if (op.kind === 'cells') await eng.writeCells(op.sheet, Array.isArray(op.writes) ? op.writes : [], op.reason === 'toValues' ? 'toValues' : 'edit');
        else if (op.kind === 'colWidth') await eng.setColWidth(op.sheet, op.col, op.width);
        else await eng.unmerge(op.sheet, Array.isArray(op.refs) ? op.refs : [], !!op.centerAcross);
      } catch (e) {
        const pe = toPatchError(e, `${i + 1} 番目の操作を反映できませんでした`);
        if (pe.internal) throw pe;
        throw new PatchError(`${i + 1} 番目の操作を反映できませんでした: ${pe.message}`);
      }
    }
    report(0.85, '数式・参照を更新中…');
    const out = await eng.finish();
    for (const w of eng.warnings) { try { onWarning?.(w); } catch { /* a UI callback never aborts the export */ } }
    report(1, '完了');
    return out;
  } catch (e) {
    throw toPatchError(e, 'ファイルを書き出せませんでした');
  }
}
