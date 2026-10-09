/**
 * export facade.
 *
 *  exportWithJs(premerged: Uint8Array, log: LoggedOp[], onProgress?: ProgressFn): Promise<Uint8Array>
 *    - = xlsx.replayLog (package patcher). Works everywhere (no Excel needed).
 *
 *  logToRunnerJob(log: LoggedOp[], opts: { format: 'xlsx' | 'xlsm'; sheet: string }): RunnerJob
 *    - converts the op log into excel-runner.ps1 ops (struct ops -> moveCols/insertCols/deleteCols/insertRows/
 *      deleteRows with the clipboard-free move method; cells -> setCells with UI-form formulas for Formula2;
 *      colWidth -> setColumnWidth). The server runs it on the premerged package with the installed Excel.
 */
export { exportWithJs } from './js';
export { logToRunnerJob } from './comJob';
