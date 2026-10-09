/**
 * Excel本体 engine against the REAL Excel (skipped unless EXCEL_TESTS=1).
 *   1. An Excel-authored fixture (formulas on two sheets, names) goes through the server's jobs API with the real
 *      excel-runner.ps1: column moves (left / right / multi-column / to the end), insert, delete, row insert / move,
 *      cell writes (literal '=' strings, '001', formulas) and a column width. Every resulting formula (both sheets)
 *      and every defined name must equal excelRef.adjustFormulaSeq's prediction; the user's clipboard text must be
 *      unchanged afterwards.
 *   2. Refusals (merged cells / tables) come back as Japanese errors and write nothing.
 *   3. When the JS engine (xlsx.replayLog) is implemented, its output for the same log must give the same formulas
 *      when opened in Excel (ARCHITECTURE §6).
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { adjustFormula, mapIndex, type StructOp } from '../../src/core/formula/excelRef';
import { logToRunnerJob } from '../../src/core/export/comJob';
import type { CellWrite, LoggedOp } from '../../src/core/types';
import { runRunner } from '../../server/excel/run-job.mjs';
import { EXCEL_TESTS, cellAt, clipboardText, dumpBook, makeFixture, tempDir, type Dump, type FixtureSpec, type SpecCell } from './lib';
import { makeSandbox, raw, startServer, waitJob, type RunningServer, type Sandbox } from '../server/helpers';

const MAIN = '名簿';
const SUM_SHEET = '集計';
const RUNNER = path.resolve(__dirname, '..', '..', 'server', 'excel', 'excel-runner.ps1');

function mainSpec(): FixtureSpec {
  const cells: SpecCell[] = [];
  const hdr = ['ID', '氏名', '点数1', '点数2', '合計', '比率', '備考', '部署', 'X', 'Y', '範囲和', '検索'];
  hdr.forEach((h, i) => cells.push([`${String.fromCharCode(65 + i)}1`, 's', h]));
  for (let r = 2; r <= 11; r++) {
    cells.push([`A${r}`, 'n', 1000 + r], [`B${r}`, 's', `氏名${r}`], [`C${r}`, 'n', r * 10], [`D${r}`, 'n', r * 3]);
    cells.push([`E${r}`, 'f', `=C${r}+D${r}`], [`F${r}`, 'f', `=E${r}/$E$12`], [`G${r}`, 's', `メモ${r}`], [`H${r}`, 's', r % 2 ? '営業' : '開発']);
    cells.push([`I${r}`, 'n', r], [`J${r}`, 'f', `=I${r}*2`], [`K${r}`, 'f', `=SUM($C${r}:D${r})`], [`L${r}`, 'f', `=VLOOKUP(A${r},$A$2:$K$11,5,FALSE)`]);
  }
  cells.push(['A12', 's', '計'], ['C12', 'f', '=SUM(C2:C11)'], ['D12', 'f', '=SUM(D2:D11)'], ['E12', 'f', '=SUM(E2:E11)']);
  cells.push(['G12', 'f', '=COUNTA(B:B)'], ['H12', 'f', '=SUMPRODUCT(C2:C11,D2:D11)'], ['J12', 'f', '=AVERAGE(J2:J11)'], ['K12', 'f', '=SUM(C2:D11)']);
  cells.push(['L12', 'f', `=C12+${SUM_SHEET}!B3`]);
  cells.push(['A13', 'f', '=ROWS(A2:A11)'], ['B13', 'f', '=SUM(C:C)'], ['C13', 'f', '=SUM(A2:L2)'], ['D13', 'f', '=$I$2+$J$2'], ['E13', 'f', '=SUM(I2:J11)']);
  const sum: SpecCell[] = [
    ['A1', 's', '合計'],
    ['B1', 'f', `=${MAIN}!E12`],
    ['B2', 'f', `=SUM(${MAIN}!C2:C11)`],
    ['B3', 'f', `=${MAIN}!$D$12`],
    ['B4', 'f', `=SUM(${MAIN}!C:D)`],
    ['B5', 'f', `=COUNTIF(${MAIN}!H2:H11,"営業")`],
    ['B6', 'f', '=SUM(点数範囲)'],
    ['B7', 'f', `=INDEX(${MAIN}!A2:L11,3,2)`],
    ['B8', 'f', `=${MAIN}!I2+${MAIN}!J3`],
    ['B9', 'f', `=SUM(${MAIN}!$A$2:$L$2)`],
    ['B10', 'f', `=${MAIN}!H2&"!"`],
  ];
  return {
    sheets: [
      { name: MAIN, cells, widths: [[2, 14]] },
      { name: SUM_SHEET, cells: sum },
    ],
    names: [
      { name: '点数範囲', refersTo: `=${MAIN}!$C$2:$C$11` },
      { name: '合計列', refersTo: `=${MAIN}!$E$2:$E$11` },
      { name: '全体', refersTo: `=${MAIN}!$A$1:$L$13` },
    ],
  };
}

/** Tracks every cell / name of the fixture through the log with excelRef (the prediction). */
class Tracker {
  cells = new Map<string, { sheet: string; r: number; c: number; f?: string; literal?: unknown }>();
  names = new Map<string, string>();
  constructor(d: Dump) {
    for (const s of d.sheets) for (const c of s.cells) if (c.hf && c.f) this.cells.set(`${s.name}|${c.r}|${c.c}`, { sheet: s.name, r: c.r, c: c.c, f: c.f.slice(1) });
    for (const n of d.names) this.names.set(n.name, n.refersTo.replace(/^=/, ''));
  }
  apply(entry: LoggedOp) {
    if (entry.kind === 'struct') this.struct(entry.sheet, entry.op);
    else if (entry.kind === 'cells') for (const w of entry.writes) this.write(entry.sheet, w);
  }
  private struct(sheet: string, op: StructOp) {
    const next = new Map<string, { sheet: string; r: number; c: number; f?: string; literal?: unknown }>();
    for (const cell of this.cells.values()) {
      let { r, c } = cell;
      if (cell.sheet === sheet) {
        const m = mapIndex(op.axis === 'col' ? c : r, op);
        if (m == null) continue; // deleted
        if (op.axis === 'col') c = m;
        else r = m;
      }
      const f = cell.f !== undefined ? adjustFormula(cell.f, op, { hostSheet: cell.sheet, targetSheet: sheet }).formula : undefined;
      next.set(`${cell.sheet}|${r}|${c}`, { ...cell, r, c, f });
    }
    this.cells = next;
    for (const [k, v] of this.names) this.names.set(k, adjustFormula(v, op, { hostSheet: null, targetSheet: sheet, isName: true }).formula);
  }
  private write(sheet: string, w: CellWrite) {
    const key = `${sheet}|${w.row}|${w.col}`;
    if (!w.cell) this.cells.delete(key);
    else if (w.cell.f) this.cells.set(key, { sheet, r: w.row, c: w.col, f: w.cell.f });
    else this.cells.set(key, { sheet, r: w.row, c: w.col, literal: w.cell.v });
  }
  /** column / row where an ORIGINAL position ended up (null = deleted) */
  static follow(ops: StructOp[], axis: 'col' | 'row', v: number): number | null {
    let x: number | null = v;
    for (const op of ops) if (x != null && op.axis === axis) x = mapIndex(x, op);
    return x;
  }
}

function compare(t: Tracker, d: Dump) {
  const mismatches: string[] = [];
  let checked = 0;
  for (const cell of t.cells.values()) {
    const got = cellAt(d, cell.sheet, cell.r, cell.c);
    if (cell.f !== undefined) {
      checked++;
      const want = `=${cell.f}`;
      if (!got || !got.hf || got.f !== want) mismatches.push(`${cell.sheet}!R${cell.r}C${cell.c}: want ${want} got ${got ? (got.hf ? got.f : `value ${JSON.stringify(got.v)}`) : 'empty'}`);
    } else {
      checked++;
      if (!got || got.hf || got.v !== cell.literal) mismatches.push(`${cell.sheet}!R${cell.r}C${cell.c}: want literal ${JSON.stringify(cell.literal)} got ${JSON.stringify(got)}`);
    }
  }
  for (const [name, want] of t.names) {
    checked++;
    const got = d.names.find((n) => n.name === name);
    if (!got || got.refersTo !== `=${want}`) mismatches.push(`name ${name}: want =${want} got ${got?.refersTo}`);
  }
  return { mismatches, checked };
}

describe.skipIf(!EXCEL_TESTS)('Excel本体 engine with the real Excel', () => {
  let dir: string;
  let fixture: string;
  let before: Dump;
  let sb: Sandbox;
  let srv: RunningServer;

  beforeAll(async () => {
    dir = tempDir('main');
    fixture = path.join(dir, 'fixture.xlsx');
    await makeFixture(mainSpec(), fixture);
    before = await dumpBook(fixture);
    sb = makeSandbox('com');
    const env: Record<string, string> = { EXCEL_RUNNER: RUNNER, EXCEL_DETECT: '' };
    srv = await startServer(sb, env);
  }, 180_000);
  afterAll(async () => {
    await srv?.stop();
    if (sb) fs.rmSync(sb.dir, { recursive: true, force: true });
    if (dir) fs.rmSync(dir, { recursive: true, force: true });
  });

  it('fixture sanity: Excel wrote the formulas and names', () => {
    expect(cellAt(before, MAIN, 2, 5)?.f).toBe('=C2+D2');
    expect(cellAt(before, SUM_SHEET, 6, 2)?.f).toBe('=SUM(点数範囲)');
    expect(before.names.map((n) => n.name).sort()).toEqual(['全体', '合計列', '点数範囲'].sort());
  });

  it('replays moves / insert / delete / row ops / cell writes exactly like excelRef predicts, keeps the clipboard', async () => {
    const t = new Tracker(before);
    const log: LoggedOp[] = [];
    const structOps: StructOp[] = [];
    const push = (e: LoggedOp) => {
      log.push(e);
      t.apply(e);
      if (e.kind === 'struct') structOps.push(e.op);
    };
    const struct = (op: StructOp) => push({ kind: 'struct', sheet: MAIN, op });
    struct({ type: 'move', axis: 'col', from: 3, count: 1, to: 8 }); // 点数1 -> before 部署 (right)
    struct({ type: 'move', axis: 'col', from: 9, count: 2, to: 2 }); // X,Y -> before 氏名 (left, 2 columns)
    struct({ type: 'insert', axis: 'col', at: 5, count: 1 });
    // the doc marks the texts Excel would parse with quote: true (the runner no longer guesses); the phone number needs no flag
    const writes: CellWrite[] = [
      { row: 2, col: 5, cell: { t: 's', v: '=not a formula', quote: true } },
      { row: 3, col: 5, cell: { t: 's', v: '001', quote: true } },
      { row: 4, col: 5, cell: { t: 'n', v: 42 } },
      { row: 5, col: 5, cell: { t: 'n', v: 0, f: 'SUM(A2:B2)' } },
      { row: 6, col: 5, cell: { t: 'b', v: true } },
      { row: 7, col: 5, cell: { t: 's', v: '1/2', quote: true } },
      { row: 8, col: 5, cell: { t: 's', v: '+81 3', quote: true } },
      { row: 9, col: 5, cell: { t: 's', v: '03-1234-5678' } },
    ];
    push({ kind: 'cells', sheet: MAIN, reason: 'edit', writes });
    struct({ type: 'move', axis: 'col', from: 1, count: 1, to: 14 }); // ID -> the very end
    const memo = Tracker.follow(structOps, 'col', 7)!; // original G (備考): nothing refers to it
    struct({ type: 'delete', axis: 'col', at: memo, count: 1 });
    struct({ type: 'move', axis: 'col', from: 2, count: 3, to: 10 }); // 3 columns to the right
    struct({ type: 'insert', axis: 'row', at: 5, count: 1 });
    struct({ type: 'move', axis: 'row', from: 3, count: 1, to: 9 });
    push({ kind: 'cells', sheet: MAIN, reason: 'edit', writes: [{ row: 14, col: 3, cell: { t: 'n', v: 0, f: 'B2*2' } }] });
    struct({ type: 'move', axis: 'col', from: 6, count: 1, to: 1 });
    push({ kind: 'colWidth', sheet: MAIN, col: 2, width: 20 });

    const job = logToRunnerJob(log, { format: 'xlsx', sheet: MAIN });
    const clipBefore = await clipboardText();
    const created = await raw(srv.port, 'POST', '/api/jobs', { body: JSON.stringify(job) });
    expect(created.status).toBe(200);
    const { id } = created.json();
    const put = await raw(srv.port, 'PUT', `/api/jobs/${id}/input`, { body: fs.readFileSync(fixture) });
    expect(put.status).toBe(200);
    const v = await waitJob(srv.port, id, 300_000);
    expect(v.state, JSON.stringify(v)).toBe('done');
    const res = await raw(srv.port, 'GET', `/api/jobs/${id}/result`);
    expect(res.status).toBe(200);
    const outFile = path.join(dir, 'excel-out.xlsx');
    fs.writeFileSync(outFile, res.body);
    const clipAfter = await clipboardText();
    expect(clipAfter.text).toBe(clipBefore.text);

    const after = await dumpBook(outFile);
    const { mismatches, checked } = compare(t, after);
    expect(mismatches, mismatches.join('\n')).toEqual([]);
    expect(checked).toBeGreaterThan(60);
    // literal strings were not parsed by Excel
    const lit = [...t.cells.values()].filter((c) => c.literal !== undefined).map((c) => cellAt(after, c.sheet, c.r, c.c)!);
    expect(lit.map((c) => c.v)).toEqual(expect.arrayContaining(['=not a formula', '001', '1/2', '+81 3', '03-1234-5678', 42, true]));
    // nothing was converted behind the log's back: no TEXT_COERCED safety-net warning for the flagged texts or the phone number
    expect(v.warnings.filter((w: string) => /文字列として書き直しました/.test(w))).toEqual([]);
    // LoggedOp colWidth = the value STORED in <col width> (20): the runner sets ColumnWidth = 20 - padding
    // (0.625 for the 11pt 游ゴシック default font, 5/7 for Calibri; measured in the workbook's own Excel)
    const columnWidth = after.sheets.find((s) => s.name === MAIN)!.widths[1];
    expect(20 - columnWidth).toBeGreaterThan(0.3);
    expect(20 - columnWidth).toBeLessThan(1.2);
    expect(v.warnings.filter((w: string) => /#REF!/.test(w))).toEqual([]);
  }, 600_000);

  it('JS engine (replayLog) output gives the same formulas in Excel', async () => {
    const { replayLog } = await import('../../src/core/xlsx/replay');
    const log: LoggedOp[] = [
      { kind: 'struct', sheet: MAIN, op: { type: 'move', axis: 'col', from: 3, count: 1, to: 8 } },
      { kind: 'struct', sheet: MAIN, op: { type: 'move', axis: 'col', from: 9, count: 2, to: 2 } },
      { kind: 'struct', sheet: MAIN, op: { type: 'insert', axis: 'col', at: 5, count: 1 } },
      { kind: 'struct', sheet: MAIN, op: { type: 'delete', axis: 'col', at: 9, count: 1 } },
    ];
    let js: Uint8Array;
    try {
      js = await replayLog(new Uint8Array(fs.readFileSync(fixture)), log);
    } catch (e) {
      if (/not implemented/i.test(String((e as Error).message))) {
        console.warn('[tests/com] replayLog is still a stub: JS-vs-Excel comparison skipped');
        return;
      }
      throw e;
    }
    const jsFile = path.join(dir, 'js-out.xlsx');
    fs.writeFileSync(jsFile, js);
    const t = new Tracker(before);
    for (const e of log) t.apply(e);
    const { mismatches } = compare(t, await dumpBook(jsFile));
    expect(mismatches, mismatches.join('\n')).toEqual([]);
  }, 300_000);
});

describe.skipIf(!EXCEL_TESTS)('Excel本体 engine refusals (real Excel)', () => {
  let dir: string;
  let fixture: string;
  beforeAll(async () => {
    dir = tempDir('refuse');
    fixture = path.join(dir, 'refuse.xlsx');
    const cells: SpecCell[] = [
      ['A1', 's', 'h1'], ['B1', 's', 'h2'], ['C1', 's', 'h3'],
      ['A2', 'n', 1], ['B2', 'n', 2], ['C2', 'n', 3],
      ['A3', 'n', 4], ['B3', 'n', 5], ['C3', 'n', 6],
      ['E1', 's', '結合'], ['H1', 's', 'free'], ['H2', 'n', 9], ['J2', 'f', '=SUM(Tbl[h1])'],
    ];
    await makeFixture({ sheets: [{ name: 'T', cells, merges: ['E1:F1'], tables: [{ name: 'Tbl', range: 'A1:C3' }] }] }, fixture);
  }, 180_000);
  afterAll(() => {
    if (dir) fs.rmSync(dir, { recursive: true, force: true });
  });

  const cases: Array<[string, Record<string, unknown>, string, RegExp]> = [
    ['part of a merged area', { op: 'moveCols', sheet: 'T', start: 5, count: 1, to: 1, method: 'clipFree' }, 'REFUSED_MERGE', /結合セル/],
    ['a column into a table', { op: 'moveCols', sheet: 'T', start: 8, count: 1, to: 2, method: 'clipFree' }, 'REFUSED_TABLE', /テーブル/],
    ['a table column out of the table', { op: 'moveCols', sheet: 'T', start: 1, count: 1, to: 9, method: 'clipFree' }, 'REFUSED_TABLE', /テーブル/],
    ['a range partly over a table', { op: 'moveCols', sheet: 'T', start: 3, count: 2, to: 9, method: 'clipFree' }, 'REFUSED_TABLE', /テーブル/],
  ];
  cases.forEach(([label, op, code, re], index) => {
    it(`refuses to move ${label} (Japanese message, nothing written)${index === 0 ? ' and removes the scratch folder it made itself' : ''}`, async () => {
      const out = path.join(dir, `out-${code}-${Math.random().toString(36).slice(2)}.xlsx`);
      const jobFile = `${out}.job.json`;
      // case 0: no workDir in the job -> the runner creates <TEMP>/xlrunner-<guid> itself (TEMP points to an own folder here) and must delete it
      const ownTemp = path.join(dir, 'own-temp');
      if (index === 0) fs.mkdirSync(ownTemp, { recursive: true });
      const job: Record<string, unknown> = { version: 1, format: 'xlsx', input: fixture, output: out, ops: [op], options: { recalc: true, clipboard: 'restoreEachOp' } };
      if (index !== 0) job.workDir = path.join(dir, 'w');
      fs.writeFileSync(jobFile, JSON.stringify(job));
      const env: Record<string, string> = { XLRUNNER_CACHE: path.join(dir, 'cache') };
      if (index === 0) Object.assign(env, { TEMP: ownTemp, TMP: ownTemp });
      const r = await runRunner({ runner: RUNNER, jobFile, env });
      expect(r.ok).toBe(false);
      expect(r.code).toBe(code);
      expect(r.message).toMatch(re);
      expect(r.exitCode).toBe(5);
      expect(fs.existsSync(out)).toBe(false);
      expect(r.excelPid).toBeGreaterThan(0);
      if (index === 0) expect(fs.readdirSync(ownTemp).filter((n) => /^xlrunner-/.test(n))).toEqual([]);
    }, 180_000);
  });

  it('rejects a bad job before starting Excel', async () => {
    const jobFile = path.join(dir, 'bad.job.json');
    fs.writeFileSync(jobFile, JSON.stringify({ version: 1, format: 'xlsx', input: fixture, output: path.join(dir, 'x.xlsm'), ops: [] }));
    const r = await runRunner({ runner: RUNNER, jobFile, env: { XLRUNNER_CACHE: path.join(dir, 'cache') } });
    expect(r.ok).toBe(false);
    expect(r.code).toBe('BAD_JOB');
    expect(r.message).toMatch(/拡張子/);
    expect(r.excelPid).toBe(0);
  }, 120_000);
});
