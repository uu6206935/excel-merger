/**
 * Round 3 parity fixes against the REAL Excel (skipped unless EXCEL_TESTS=1). ONE runner job (one Excel start) on a workbook built
 * as a package (tests/com/r3.ts, openpyxl style: fill and borders only on the top-left cell of the merged areas):
 *
 *   1. quote flag: CellWrite.quote === true is the only thing that gives a text the quotePrefix; "03-1234-5678" (phone) and
 *      every other unflagged text are written as they are (the e2e spec 03 "Excel本体 vs 内蔵エンジン" difference). A text Excel
 *      converts anyway is rewritten as text WITH the apostrophe, keeps the number format the cell had, and is reported (TEXT_COERCED).
 *   2. unmerge: fill + borders of the top-left cell are copied to the other cells of each former area, then 「選択範囲内で中央」.
 *   3. numFmt 'General' works (Japanese Excel rejects the English word through COM: the runner translates it).
 *   and the same log through replayLog (the JS engine) must give the same package (tools/lib/package-diff: resolved styles incl.
 *   quotePrefix, fills, borders, values) — sheet 名簿 only; sheet 防御 (the runner's safety net) is not part of the log.
 *   Disk hygiene: after the job the runner has removed its working copy and the saved workbook from the work folder.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { logToRunnerJob } from '../../src/core/export/comJob';
import { runRunner, type RunnerEvent, type RunnerResult } from '../../server/excel/run-job.mjs';
import { comparePackages } from '../../tools/lib/package-diff.mjs';
import { EXCEL_TESTS, diffRawCells, rawCells, tempDir, type RawCell } from './lib';
import { DEF_TEXTS, R3_DEF, R3_SHEET, buildR3Fixture, buildR3Log, defensiveOps } from './r3';

const RUNNER = path.resolve(__dirname, '..', '..', 'server', 'excel', 'excel-runner.ps1');

/**
 * Differences between the two engines that are understood and accepted: [regex over the row detail, who is right + evidence].
 * Empty = the two engines agree on everything compared.
 */
const KNOWN: Array<[RegExp, string]> = [];

describe.skipIf(!EXCEL_TESTS)('round 3 parity: quote flag, unmerge formats, General (real Excel, one job)', () => {
  let dir: string;
  let work: string;
  let r: RunnerResult;
  let events: RunnerEvent[];
  let excelBytes: Uint8Array;
  let jsBytes: Uint8Array | null = null;
  let jsError = '';
  let main: Map<string, RawCell>;
  let def: Map<string, RawCell>;
  const warns = () => events.filter((e) => e.type === 'warn').map((e) => ({ code: String(e.code), message: String(e.message) }));

  beforeAll(async () => {
    dir = tempDir('r3');
    const fixture = path.join(dir, 'fixture.xlsx');
    const fixtureBytes = await buildR3Fixture();
    fs.writeFileSync(fixture, fixtureBytes);
    const log = buildR3Log();
    const job = logToRunnerJob(log, { format: 'xlsx', sheet: R3_SHEET });
    job.ops.push(...defensiveOps());
    work = path.join(dir, 'work');
    const out = path.join(dir, 'excel.xlsx');
    const jobFile = path.join(dir, 'job.json');
    fs.writeFileSync(jobFile, JSON.stringify({ ...job, input: fixture, output: out, workDir: work }));
    events = [];
    r = await runRunner({ runner: RUNNER, jobFile, env: { XLRUNNER_CACHE: path.join(dir, 'cache') }, onEvent: (e) => events.push(e) });
    if (r.ok) {
      excelBytes = new Uint8Array(fs.readFileSync(out));
      main = await rawCells(excelBytes, R3_SHEET);
      def = await rawCells(excelBytes, R3_DEF);
    }
    try {
      const { replayLog } = await import('../../src/core/xlsx/replay');
      jsBytes = await replayLog(fixtureBytes, log);
      fs.writeFileSync(path.join(dir, 'js.xlsx'), jsBytes); // kept with KEEP_COM_OUTPUT=1: node tools/compare.mjs excel.xlsx js.xlsx
    } catch (e) {
      jsError = String((e as Error).message);
    }
  }, 300_000);
  afterAll(() => {
    if (dir && process.env.KEEP_COM_OUTPUT !== '1') fs.rmSync(dir, { recursive: true, force: true });
    else if (dir) console.log(`[tests/com] outputs kept in ${dir}`);
  });

  it('the job ran (no BAD_NUMFMT: General was accepted) and the runner cleaned its own files', () => {
    expect(r.ok, JSON.stringify(r)).toBe(true);
    const codes = warns().map((w) => w.code);
    expect(codes).not.toContain('BAD_NUMFMT');
    expect(codes).not.toContain('UNMERGE_FORMATS');
    // disk hygiene: working copy and saved workbook are gone from the (caller owned) work folder, also the output is complete
    expect(fs.existsSync(work) ? fs.readdirSync(work) : []).toEqual([]);
    expect(excelBytes.length).toBeGreaterThan(2000);
  });

  it('quote: true gets the quotePrefix, nothing else does (phone number stays a plain text)', () => {
    const c = (a: string) => main.get(a)!;
    expect(c('C4')).toMatchObject({ value: '03-1234-5678', t: 's', quotePrefix: false, numFmtId: 0 }); // the e2e spec 03 difference
    expect(c('C5')).toMatchObject({ value: '001', t: 's', quotePrefix: true, numFmtId: 0 });
    expect(c('C6')).toMatchObject({ value: '1/2', t: 's', quotePrefix: true, numFmtId: 0 }); // text, and no date format left behind
    expect(c('C7')).toMatchObject({ value: '=text', t: 's', quotePrefix: true });
    expect(c('C7').formula).toBeUndefined();
    expect(c('C8')).toMatchObject({ value: '007', t: 's', quotePrefix: false, numFmt: '@' }); // text format, no apostrophe needed
    expect(c('C9')).toMatchObject({ value: '005', t: 's', quotePrefix: true, numFmt: '@' }); // both: exactly what was asked
    expect(c('D4')).toMatchObject({ value: 'メモ', quotePrefix: false });
  });

  it('numFmt General resets the format (Japanese Excel rejects the English word through COM)', () => {
    const c = (a: string) => main.get(a)!;
    expect(c('E4')).toMatchObject({ value: '5', numFmtId: 0 }); // was a date format
    expect(c('E9')).toMatchObject({ value: 'abc', numFmtId: 0 }); // was the text format
    expect(c('E6')).toMatchObject({ numFmt: 'yyyy/m/d' }); // a real format still works
    expect(c('E5').numFmt).toBe('yyyy/m/d'); // untouched cells keep theirs
  });

  it('unmerge copies the fill and the borders of the top-left cell, then centres across the selection', () => {
    const c = (a: string) => main.get(a)!;
    for (const a of ['A1', 'B1', 'C1', 'D1', 'E1', 'F1']) {
      expect(c(a), a).toMatchObject({ fill: 'solid rgb=FF1F4E78', horizontal: 'centerContinuous' });
      expect(c(a).border.top, `${a} top`).toMatch(/^thin/);
      expect(c(a).border.bottom, `${a} bottom`).toMatch(/^thin/);
    }
    expect(c('F1').border.right).toMatch(/^thin/); // the last cell of the row carries the right line
    expect(c('A1').border.left).toMatch(/^thin/);
    // the two-row band (theme fill, no centre across): every cell got the fill
    for (const a of ['H1', 'I1', 'J1', 'H2', 'I2', 'J2']) expect(c(a)?.fill, a).toMatch(/^solid (theme=5|rgb=)/);
    expect(c('I1').horizontal).toBeUndefined();
    // the copies were counted by the runner: the one-row title (fill + top + bottom + right) and the two-row band
    const copies = events.filter((e) => e.type === 'progress' && e.op === 'unmerge').map((e) => Number(e.formatCopies));
    expect(copies.length).toBe(2);
    expect(copies[0]).toBe(4);
    expect(copies[1]).toBe(7); // row 1: fill + top + right; row 2: fill + bottom + left + right (the outline of the area)
  });

  it('a plain string Excel converts anyway is rewritten as text with the apostrophe, keeps its number format, and is reported', () => {
    const cell = (i: number) => def.get(`A${i + 1}`)!;
    const coerced = ['TRUE', '1/2', '=(', '2026-01-02', '007', '=abc', ' 12'];
    DEF_TEXTS.forEach((text, i) => {
      const x = cell(i);
      // Excel keeps a line break as CR LF inside the cell (the file escapes the CR as _x000D_): compare without CR
      const plain = (t: string | undefined) => (t ?? '').replace(/_x000D_/g, '').replace(/\r/g, '');
      expect(x.t, `${text} is a string`).toBe('s');
      expect(plain(x.value), text).toBe(plain(text));
      expect(x.formula, `${text} is not a formula`).toBeUndefined();
      if (coerced.includes(text) || text === "'x") expect(x.quotePrefix, `${text} quotePrefix`).toBe(true);
      else expect(x.quotePrefix, `${text} quotePrefix`).toBe(false);
    });
    // number formats the cells had before: General / date / text / General / date / General ... (Excel set a date format for "1/2")
    const fmts = DEF_TEXTS.map((_, i) => cell(i).numFmt);
    expect(fmts.slice(0, 6)).toEqual(['General', 'yyyy/m/d', '@', 'General', 'yyyy/m/d', 'General']);
    expect(fmts.slice(6)).toEqual(['General', 'General', 'General', 'General', 'General', 'General']);
    const coercedWarns = warns().filter((w) => w.code === 'TEXT_COERCED');
    expect(coercedWarns.length).toBe(coerced.length);
    for (const w of coercedWarns) expect(w.message).toMatch(/文字列として書き直しました/);
    for (const a of ['A1', 'A2', 'A4', 'A5', 'A6', 'A7', 'A11']) expect(coercedWarns.some((w) => w.message.includes(`セル ${a} `)), a).toBe(true);
    for (const a of ['A3', 'A8', 'A9', 'A10', 'A12']) expect(coercedWarns.some((w) => w.message.includes(`セル ${a} `)), a).toBe(false);
  });

  it('JS engine (replayLog) gives the same sheet as the runner for the same log: values, quotePrefix, formats, fills, alignment, drawn lines', async () => {
    expect(jsBytes, `replayLog failed: ${jsError}`).toBeTruthy();
    const js = await rawCells(jsBytes!, R3_SHEET);
    const diffs = diffRawCells(main, js, ['excel', 'js']);
    const accepted = diffs.filter((d) => KNOWN.some(([re]) => re.test(d)));
    const real = diffs.filter((d) => !accepted.includes(d));
    if (real.length) console.warn(`[tests/com] parity: Excel runner vs replayLog, ${real.length} differences\n${real.join('\n')}`);
    else console.log(`[tests/com] parity: runner and replayLog agree on sheet ${R3_SHEET} (${main.size} cells checked, ${accepted.length} accepted differences)`);
    expect(real).toEqual([]);
  });

  it('tools/lib/package-diff (the package view of tools/compare.mjs) finds no other difference than how a line between two cells is stored', async () => {
    expect(jsBytes, `replayLog failed: ${jsError}`).toBeTruthy();
    const { rows } = await comparePackages(excelBytes, jsBytes!);
    const bad = rows.filter((x) => !x.ok && (x.sheet === R3_SHEET || !x.sheet));
    // Excel stores a line shared by two neighbours on ONE of them (A1.right -> B1.left; H1.bottom -> H2.top): that is a per-cell style
    // row only; the drawn lines are compared above. Everything else (merges, widths, names, constants, formulas ...) must be equal.
    const other = bad.filter((x) => x.group !== 'styles').map((x) => `[${x.group}] ${x.sheet}: ${x.detail}`);
    const styles = bad.filter((x) => x.group === 'styles').map((x) => x.detail);
    if (styles.length) console.log(`[tests/com] parity: package view style rows (representation of shared lines): ${styles.join(' || ').slice(0, 1500)}`);
    expect(other).toEqual([]);
  });
});
