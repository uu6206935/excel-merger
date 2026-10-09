/**
 * Column moves INSIDE an Excel table through the runner, against the REAL Excel (skipped unless EXCEL_TESTS=1).
 *
 * Inside a ListObject the runner cannot use its clipboard-free sequence (it would delete table columns), so it uses
 * Excel's own "Cut, then Insert Cut Cells", verifies the moved window and retries when another program took the
 * clipboard in between (Insert then adds BLANK columns). The clipboard is shared with other programs (and other
 * engineers' tests), so this runs the same job TABLE_MOVE_RUNS times (default 8) and reports the success rate.
 *
 * Every successful output must be exact: header order = the predicted permutation, table range unchanged, every
 * column's values followed it (no blank column), calculated columns / structured references outside the table /
 * names unchanged, A1 references adjusted exactly like excelRef predicts. A failed job must be a clean refusal
 * (MOVE_VERIFY_FAILED, nothing saved). Finally the JS engine (replayLog) must give the same workbook.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { adjustFormula, mapIndex, numToCol, type StructOp } from '../../src/core/formula/excelRef';
import { logToRunnerJob } from '../../src/core/export/comJob';
import type { LoggedOp } from '../../src/core/types';
import { runRunner, type RunnerEvent, type RunnerResult } from '../../server/excel/run-job.mjs';
import { EXCEL_TESTS, cellAt, clipboardText, diffDumps, dumpBook, makeFixture, packageDiff, tempDir, type Dump, type FixtureSpec, type SpecCell } from './lib';

const DATA = 'データ';
const SUM = '集計';
const RUNNER = path.resolve(__dirname, '..', '..', 'server', 'excel', 'excel-runner.ps1');
const RUNS = Math.max(1, Number(process.env.TABLE_MOVE_RUNS) || 8);

/** Table Sales = B2:I22 (header row 2, rows 3..22); column A and K.. are outside the table. */
const HEAD = ['No', '担当', '地域', '単価', '数量', '金額', '税込', '備考'];
const T1 = 2;
const T2 = 9;
const REGIONS = ['東', '西', '南', '北'];

function spec(): FixtureSpec {
  const cells: SpecCell[] = [['A2', 's', '行']];
  HEAD.forEach((h, i) => cells.push([`${numToCol(T1 + i)}2`, 's', h]));
  for (let r = 3; r <= 22; r++) {
    cells.push([`A${r}`, 'n', r - 2], [`B${r}`, 'n', r - 2], [`C${r}`, 's', `担当${(r % 4) + 1}`], [`D${r}`, 's', REGIONS[r % 4]]);
    cells.push([`E${r}`, 'n', 100 + ((r * 37) % 400)], [`F${r}`, 'n', ((r * 7) % 10) + 1]);
    if (r % 3 === 0) cells.push([`I${r}`, 's', `メモ${r}`]); // partly blank column
  }
  cells.push(['G3:G22', 'f', '=[@単価]*[@数量]'], ['H3:H22', 'f', '=[@金額]*1.1']);
  cells.push(['K2', 's', '集計欄'], ['K3', 'f', '=SUM(Sales[金額])'], ['K4', 'f', '=SUM(Sales[[単価]:[数量]])']);
  cells.push(['K5', 'f', '=COUNTIFS(Sales[地域],"東",Sales[数量],">3")'], ['K6', 'f', '=SUMPRODUCT(Sales[単価],Sales[数量])']);
  cells.push(['K7', 'f', '=SUM(E3:E22)'], ['K8', 'f', '=AVERAGE(F3:F22)+E3'], ['K9', 'f', '=COUNTA(B2:I2)'], ['K10', 'f', '=SUM($B$3:$C$22)']);
  const sum: SpecCell[] = [
    ['A1', 's', '合計'],
    ['B1', 'f', '=SUM(Sales[金額])'],
    ['B2', 'f', '=VLOOKUP(3,Sales,6,FALSE)'],
    ['B3', 'f', `=SUM(${DATA}!E3:E22)`],
    ['B4', 'f', '=SUM(単価列)'],
    ['B5', 'f', '=SUM(範囲)'],
    ['B6', 'f', '=Sales[[#Headers],[数量]]'],
    ['B7', 'f', `=${DATA}!F3*2`],
  ];
  return {
    sheets: [
      {
        name: DATA,
        cells,
        tables: [{ name: 'Sales', range: 'B2:I22' }],
        widths: [[3, 12], [9, 16]],
        cf: [{ range: 'F3:F22', type: 'expression', formula: '=F3>5' }],
        dv: [{ range: 'D3:D22', type: 'list', formula: '東,西,南,北' }],
      },
      { name: SUM, cells: sum },
    ],
    names: [
      { name: '単価列', refersTo: '=Sales[単価]' },
      { name: '範囲', refersTo: `=${DATA}!$E$3:$F$22` },
    ],
  };
}

/** The moves (1-based, `to` = insert-before index in the coordinates before the op, like excelRef's StructOp). */
const MOVES: StructOp[] = [
  { type: 'move', axis: 'col', from: 6, count: 1, to: 5 }, // 数量 before 単価 (left by one)
  { type: 'move', axis: 'col', from: 3, count: 2, to: 9 }, // 担当+地域 before 備考 (two columns to the right)
  { type: 'move', axis: 'col', from: 2, count: 1, to: T2 + 1 }, // No to the end of the table
  { type: 'move', axis: 'col', from: 8, count: 1, to: 2 }, // 備考 (now 8th) to the start of the table
  { type: 'move', axis: 'col', from: 4, count: 3, to: T2 + 1 }, // a block of three columns to the gap right after the table (to = c2 + 1)
];
/**
 * Moves whose destination is the gap right after the table's last column (the 3rd and the 5th one) are accepted by the doc layer
 * and by the patcher as an in-table reorder since round 3 (the column stays the last table column, as in Excel's own Cut + Insert
 * Cut Cells), so one log serves both engines.
 */
const LOG: LoggedOp[] = MOVES.map((op) => ({ kind: 'struct', sheet: DATA, op }));

const followWith = (moves: StructOp[]) => (c: number) => moves.reduce<number | null>((x, op) => (x == null ? null : mapIndex(x, op)), c)!;
const follow = followWith(MOVES);

function expectedHeaders(): string[] {
  const out: string[] = [];
  HEAD.forEach((h, i) => (out[follow(T1 + i) - T1] = h));
  return out;
}

const sheetOf = (d: Dump, name: string) => d.sheets.find((s) => s.name === name)!;

/** Every problem of one Excel output (empty = exact). */
function verify(before: Dump, after: Dump, moves: StructOp[] = MOVES): { problems: string[]; spans: string[] } {
  const follow = followWith(moves);
  const problems: string[] = [];
  const spans: string[] = [];
  const t = sheetOf(after, DATA).tables[0];
  const want = expectedHeaders();
  if (!t || t.name !== 'Sales' || t.range !== 'B2:I22') problems.push(`table: ${JSON.stringify(t)}`);
  if (JSON.stringify(t?.columns) !== JSON.stringify(want)) problems.push(`header order: ${JSON.stringify(t?.columns)} want ${JSON.stringify(want)}`);
  // values follow their column (a degraded insert would show a blank column and shifted values)
  for (let c = T1; c <= T2; c++) {
    const to = follow(c);
    for (let r = 2; r <= 22; r++) {
      const a = cellAt(before, DATA, r, c);
      const b = cellAt(after, DATA, r, to);
      const va = a?.v ?? null;
      const vb = b?.v ?? null;
      const same = typeof va === 'number' && typeof vb === 'number' ? Math.abs(va - vb) < 1e-9 : va === vb;
      if (!same) problems.push(`${HEAD[c - T1]} row ${r}: ${JSON.stringify(va)} -> ${numToCol(to)}${r} ${JSON.stringify(vb)}`);
      if (a?.hf && b?.f !== a.f) problems.push(`calculated column ${HEAD[c - T1]} row ${r}: ${a.f} -> ${b?.f}`);
    }
  }
  // formulas outside the table: structured refs unchanged, A1 refs adjusted as excelRef predicts
  for (const sheet of [DATA, SUM]) {
    for (const cell of sheetOf(before, sheet).cells) {
      if (!cell.hf || !cell.f) continue;
      if (sheet === DATA && cell.c >= T1 && cell.c <= T2) continue; // calculated columns: checked above
      let f = cell.f.slice(1);
      for (const op of moves) f = adjustFormula(f, op, { hostSheet: sheet, targetSheet: DATA }).formula;
      const got = cellAt(after, sheet, cell.r, cell.c);
      if (got?.f === `=${f}`) continue;
      if (/\[\[[^\]]+\]:\[[^\]]+\]\]/.test(cell.f)) spans.push(`${sheet}!${numToCol(cell.c)}${cell.r}: ${cell.f} -> ${got?.f}`);
      else problems.push(`${sheet}!${numToCol(cell.c)}${cell.r}: want =${f} got ${got?.f}`);
    }
  }
  for (const n of before.names) {
    let f = n.refersTo.slice(1);
    for (const op of moves) f = adjustFormula(f, op, { hostSheet: null, targetSheet: DATA, isName: true }).formula;
    const got = after.names.find((x) => x.name === n.name)?.refersTo;
    if (got !== `=${f}`) problems.push(`name ${n.name}: want =${f} got ${got}`);
  }
  return { problems, spans };
}

interface RunStat {
  run: number;
  ok: boolean;
  code?: string;
  message?: string;
  attempts: number[];
  retries: number;
  clipWarnings: string[];
  moveMs: number[];
  problems: string[];
  spans: string[];
}

describe.skipIf(!EXCEL_TESTS)('table column moves with the real Excel (Cut + Insert Cut Cells, verified, retried)', () => {
  let dir: string;
  let fixture: string;
  let before: Dump;
  let firstOk: string | null = null;
  const stats: RunStat[] = [];

  beforeAll(async () => {
    dir = tempDir('tbl');
    fixture = path.join(dir, 'table.xlsx');
    await makeFixture(spec(), fixture);
    before = await dumpBook(fixture);
  }, 240_000);
  afterAll(() => {
    if (dir && process.env.KEEP_COM_OUTPUT !== '1') fs.rmSync(dir, { recursive: true, force: true });
    else if (dir) console.log(`[tests/com] outputs kept in ${dir}`);
  });

  it('fixture sanity: an Excel-authored table with calculated columns and structured references', () => {
    const t = sheetOf(before, DATA).tables[0];
    expect(t).toMatchObject({ name: 'Sales', range: 'B2:I22', columns: HEAD });
    expect(cellAt(before, DATA, 3, 7)?.f).toBe('=[@単価]*[@数量]');
    expect(cellAt(before, DATA, 4, 11)?.f).toBe('=SUM(Sales[[単価]:[数量]])');
    expect(before.names.find((n) => n.name === '単価列')?.refersTo).toBe('=Sales[単価]');
  });

  it(`moves columns inside the table ${RUNS} times: every output exact, success rate reported`, async () => {
    const job = logToRunnerJob(LOG, { format: 'xlsx', sheet: DATA });
    expect(job.ops.map((o) => o.op)).toEqual(['moveCols', 'moveCols', 'moveCols', 'moveCols', 'moveCols']);
    // the two moves to the gap after the table reach the runner unchanged (it recognises the table itself)
    expect(job.ops[2]).toMatchObject({ start: 2, count: 1, to: T2 + 1 });
    expect(job.ops[4]).toMatchObject({ start: 4, count: 3, to: T2 + 1 });
    for (let run = 1; run <= RUNS; run++) {
      const out = path.join(dir, `run${run}.xlsx`);
      const jobFile = path.join(dir, `run${run}.job.json`);
      fs.writeFileSync(jobFile, JSON.stringify({ ...job, input: fixture, output: out, workDir: path.join(dir, `w${run}`) }));
      const events: RunnerEvent[] = [];
      const r: RunnerResult = await runRunner({ runner: RUNNER, jobFile, env: { XLRUNNER_CACHE: path.join(dir, 'cache') }, onEvent: (e) => events.push(e) });
      const moves = events.filter((e) => e.type === 'progress' && e.phase === 'op' && e.op === 'moveCols');
      const warns = events.filter((e) => e.type === 'warn').map((e) => String(e.code));
      const stat: RunStat = {
        run,
        ok: r.ok,
        code: r.code,
        message: r.ok ? undefined : r.message,
        attempts: moves.map((e) => Number(e.attempts)),
        retries: warns.filter((w) => w === 'MOVE_RETRY').length,
        clipWarnings: warns.filter((w) => /^CLIPBOARD_/.test(w)),
        moveMs: moves.map((e) => Number(e.moveMs)),
        problems: [],
        spans: [],
      };
      if (r.ok) {
        expect(moves.map((e) => e.method)).toEqual(['cutInsert', 'cutInsert', 'cutInsert', 'cutInsert', 'cutInsert']);
        const v = verify(before, await dumpBook(out));
        stat.problems = v.problems;
        stat.spans = v.spans;
        if (!firstOk && v.problems.length === 0) firstOk = out;
      } else {
        // a failed verification must be a clean refusal: nothing saved
        stat.problems = r.code === 'MOVE_VERIFY_FAILED' && !fs.existsSync(out) ? [] : [`unexpected failure ${r.code}: ${r.message}`];
      }
      stats.push(stat);
    }
    const okRuns = stats.filter((s) => s.ok).length;
    const ops = stats.flatMap((s) => s.attempts);
    const firstTry = ops.filter((a) => a === 1).length;
    const ms = stats.flatMap((s) => s.moveMs).sort((a, b) => a - b);
    console.log(
      `[tests/com] table moves: ${okRuns}/${RUNS} jobs ok (${Math.round((100 * okRuns) / RUNS)}%), ` +
        `${firstTry}/${ops.length} Cut+Insert ops right on the first attempt, ${stats.reduce((n, s) => n + s.retries, 0)} retries, ` +
        `failed jobs: ${JSON.stringify(stats.filter((s) => !s.ok).map((s) => `run ${s.run} ${s.code}: ${s.message}`))}, clipboard warnings: ${JSON.stringify(stats.flatMap((s) => s.clipWarnings))}, move ms median ${ms[Math.floor(ms.length / 2)]} max ${ms[ms.length - 1]}` +
        `${stats.some((s) => s.spans.length) ? `\n  column spans rewritten by Excel: ${[...new Set(stats.flatMap((s) => s.spans))].join(' | ')}` : ''}`,
    );
    for (const s of stats) expect(s.problems, `run ${s.run}`).toEqual([]);
    // every span reference Excel rewrote still names the same two columns (it may swap their order)
    for (const sp of new Set(stats.flatMap((s) => s.spans))) expect(sp).toMatch(/Sales\[\[(単価|数量)\]:\[(単価|数量)\]\]\) -> =SUM\(Sales\[\[(単価|数量)\]:\[(単価|数量)\]\]\)$/);
    expect(okRuns / RUNS).toBeGreaterThanOrEqual(0.75);
  }, 60_000 + RUNS * 90_000);

  /** One runner job on the fixture; returns the output path (asserts the job succeeded). */
  async function runOnce(log: LoggedOp[], label: string): Promise<string> {
    const job = logToRunnerJob(log, { format: 'xlsx', sheet: DATA });
    const out = path.join(dir, `${label}.xlsx`);
    const jobFile = path.join(dir, `${label}.job.json`);
    fs.writeFileSync(jobFile, JSON.stringify({ ...job, input: fixture, output: out, workDir: path.join(dir, `w-${label}`) }));
    const r: RunnerResult = await runRunner({ runner: RUNNER, jobFile, env: { XLRUNNER_CACHE: path.join(dir, 'cache') } });
    expect(r.ok, JSON.stringify(r)).toBe(true);
    return out;
  }

  it('JS engine (replayLog) gives the same table and formulas for the same log', async () => {
    // the SAME log (with the two moves to the gap right after the table) through both engines
    const excelOut = await runOnce(LOG, 'excel-js-same-log');
    const v = verify(before, await dumpBook(excelOut), MOVES);
    expect(v.problems, 'the Excel result of the log is exact').toEqual([]);
    const { replayLog } = await import('../../src/core/xlsx/replay');
    const js = await replayLog(new Uint8Array(fs.readFileSync(fixture)), LOG);
    const jsFile = path.join(dir, 'js.xlsx');
    fs.writeFileSync(jsFile, js);
    const [xd, jd] = await Promise.all([dumpBook(excelOut), dumpBook(jsFile)]);
    const diffs: string[] = [];
    for (const name of [DATA, SUM]) {
      const xs = sheetOf(xd, name);
      const js2 = sheetOf(jd, name);
      if (JSON.stringify(xs.tables) !== JSON.stringify(js2.tables)) diffs.push(`${name} tables: ${JSON.stringify(xs.tables)} | ${JSON.stringify(js2.tables)}`);
      const at = new Map(js2.cells.map((c) => [`${c.r},${c.c}`, c]));
      for (const c of xs.cells) {
        const o = at.get(`${c.r},${c.c}`);
        at.delete(`${c.r},${c.c}`);
        const addr = `${name}!${numToCol(c.c)}${c.r}`;
        if (!o) diffs.push(`${addr}: only in Excel ${JSON.stringify(c)}`);
        else {
          if (c.f !== o.f) diffs.push(`${addr} formula: ${c.f} | ${o.f}`);
          const same = typeof c.v === 'number' && typeof o.v === 'number' ? Math.abs(c.v - o.v) < 1e-9 : c.v === o.v;
          if (!same) diffs.push(`${addr} value: ${JSON.stringify(c.v)} | ${JSON.stringify(o.v)}`);
          if (c.nf !== o.nf) diffs.push(`${addr} format: ${c.nf} | ${o.nf}`);
        }
      }
      for (const o of at.values()) diffs.push(`${name}!${numToCol(o.c)}${o.r}: only in JS ${JSON.stringify(o)}`);
      for (let i = 0; i < 12; i++) if (Math.abs(xs.widths[i] - js2.widths[i]) > 1e-6) diffs.push(`${name} width ${numToCol(i + 1)}: ${xs.widths[i]} | ${js2.widths[i]}`);
    }
    const names = (d: Dump) => d.names.map((n) => `${n.name}=${n.refersTo}`).sort();
    if (JSON.stringify(names(xd)) !== JSON.stringify(names(jd))) diffs.push(`names: ${JSON.stringify(names(xd))} | ${JSON.stringify(names(jd))}`);
    diffs.push(...(await packageDiff(new Uint8Array(fs.readFileSync(excelOut)), js))); // CF / DV / tables / names in the files
    if (diffs.length) console.warn(`[tests/com] table moves, Excel runner vs replayLog: ${diffs.length} differences\n${diffs.join('\n')}`);
    expect(diffs).toEqual([]);
  }, 300_000);

  it('the user clipboard is put back after the Cut + Insert Cut Cells moves inside the table', async () => {
    // Read-only check of the shared clipboard (text): another program copying during the ~40 s run would make it differ,
    // so a difference is retried before it counts.
    let ok = false;
    let last = '';
    for (let attempt = 1; attempt <= 3 && !ok; attempt++) {
      const b = await clipboardText();
      await runOnce(LOG, `clip${attempt}`);
      const a = await clipboardText();
      ok = a.text === b.text;
      if (!ok) last = `attempt ${attempt}: before ${JSON.stringify(b.text)} after ${JSON.stringify(a.text)}`;
    }
    expect(ok, last).toBe(true);
  }, 600_000);

  it('a block moved to the gap right after the table (to = last column + 1) is accepted by BOTH engines as an in-table reorder', async () => {
    // Excel (and the runner) keep the column inside the table as its last column; since round 3 the doc layer and replayLog
    // accept the same move (the JS engine reorders the table columns; the table range stays B2:I22)
    const { replayLog } = await import('../../src/core/xlsx/replay');
    const js = await replayLog(new Uint8Array(fs.readFileSync(fixture)), LOG);
    expect(js.length).toBeGreaterThan(1000);
    const { readBook } = await import('../../src/core/xlsx/read');
    const book = await readBook(js, 'js.xlsx');
    const t = book.sheets.find((x) => x.name === DATA)!.features.tables[0];
    expect(t.ref).toBe('B2:I22');
    expect(t.columns).toEqual(expectedHeaders());
    expect(firstOk, 'the Excel runs above produced exact outputs for this log').toBeTruthy();
  }, 60_000);
});

/**
 * Cells above / below the table in the table's own columns (a title row, a totals row): Excel's Cut + Insert Cut Cells
 * moves the TABLE rows only when the destination lies inside the table, so those cells stay behind (the app and
 * replayLog move whole columns). The runner verifies every move, so such a job ends as a clean refusal (nothing saved);
 * a move to the table's edge moves whole columns and is exact. Column widths: Excel shifts the widths of the columns
 * right of the insertion point, the runner puts them where the app has them.
 */
describe.skipIf(!EXCEL_TESTS)('table moves with cells above / below the table in the same columns', () => {
  const S = 'S';
  let dir: string;
  let fixture: string;
  beforeAll(async () => {
    dir = tempDir('tbl2');
    fixture = path.join(dir, 'outside.xlsx');
    await makeFixture(
      {
        sheets: [
          {
            name: S,
            cells: [
              ['B1', 's', 'T-b'], ['C1', 's', 'T-c'], ['D1', 's', 'T-d'], ['E1', 's', 'T-e'],
              ['B3', 's', 'a'], ['C3', 's', 'b'], ['D3', 's', 'c'], ['E3', 's', 'd'],
              ['B4', 'n', 1], ['C4', 'n', 2], ['D4', 'n', 3], ['E4', 'n', 4],
              ['B6', 's', 'Bb'], ['C6', 's', 'Bc'], ['D6', 's', 'Bd'], ['E6', 's', 'Be'],
              ['G2', 's', 'outside'], ['G4', 'f', '=SUM(B4:E4)'],
            ],
            tables: [{ name: 'T', range: 'B3:E4' }],
            widths: [[2, 10], [3, 20], [4, 30], [5, 40], [7, 50]],
          },
        ],
      },
      fixture,
    );
  }, 240_000);
  afterAll(() => {
    if (dir && process.env.KEEP_COM_OUTPUT !== '1') fs.rmSync(dir, { recursive: true, force: true });
  });

  /** runner result of one move; retried when another program kept the clipboard busy */
  async function runMove(op: StructOp, label: string): Promise<{ r: RunnerResult; out: string }> {
    const log: LoggedOp[] = [{ kind: 'struct', sheet: S, op }];
    const job = logToRunnerJob(log, { format: 'xlsx', sheet: S });
    let last: RunnerResult | undefined;
    for (let attempt = 1; attempt <= 3; attempt++) {
      const out = path.join(dir, `${label}-${attempt}.xlsx`);
      const jobFile = `${out}.job.json`;
      fs.writeFileSync(jobFile, JSON.stringify({ ...job, input: fixture, output: out, workDir: path.join(dir, `w-${label}-${attempt}`) }));
      last = await runRunner({ runner: RUNNER, jobFile, env: { XLRUNNER_CACHE: path.join(dir, 'cache') } });
      if (last.ok || !/クリップボード/.test(String(last.message))) return { r: last, out };
    }
    return { r: last!, out: '' };
  }

  async function jsDump(op: StructOp, label: string): Promise<Dump> {
    const { replayLog } = await import('../../src/core/xlsx/replay');
    const js = await replayLog(new Uint8Array(fs.readFileSync(fixture)), [{ kind: 'struct', sheet: S, op }]);
    const file = path.join(dir, `js-${label}.xlsx`);
    fs.writeFileSync(file, js);
    return dumpBook(file);
  }

  it('a move to the table edge moves whole columns: exact, widths and cells above / below included', async () => {
    const op: StructOp = { type: 'move', axis: 'col', from: 5, count: 1, to: 2 }; // column d to the table's left edge
    const { r, out } = await runMove(op, 'edge');
    expect(r.ok, JSON.stringify(r)).toBe(true);
    const diffs = diffDumps(await dumpBook(out), await jsDump(op, 'edge'));
    expect(diffs).toEqual([]);
  }, 600_000);

  it('a move to a destination inside the table: exact or a clean refusal (nothing saved)', async () => {
    const op: StructOp = { type: 'move', axis: 'col', from: 3, count: 1, to: 5 };
    const { r, out } = await runMove(op, 'inside');
    if (r.ok) {
      console.log('[tests/com] table move with cells above / below, destination inside the table: exact result');
      expect(diffDumps(await dumpBook(out), await jsDump(op, 'inside'))).toEqual([]);
    } else {
      console.log(`[tests/com] table move with cells above / below, destination inside the table: refused (${r.code}): ${r.message}`);
      expect(r.code).toBe('MOVE_VERIFY_FAILED');
      // the message tells the user (Japanese) what to do: the UI falls back to the built-in engine for this code
      expect(r.message).toMatch(/内蔵エンジン/);
      expect(r.message).toMatch(/テーブル/);
      expect(r.message).toMatch(/保存していません/);
      expect(fs.existsSync(out)).toBe(false);
    }
  }, 600_000);
});
