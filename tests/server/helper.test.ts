/**
 * The C# half of the Excel runner (server/excel/XlRunnerHelper.cs) without Excel: tests/server/ps/helper-check.ps1 compiles it with
 * Add-Type (Windows PowerShell 5.1, like the runner does) and runs the cases below.
 *   - quote prefix (round 3): CellWrite.quote === true is the ONLY thing that puts an apostrophe in front of a string ("'001");
 *     no guessing ("03-1234-5678", "=abc", "1/2" are written as they are; the runner verifies them afterwards)
 *   - Verify: what Excel converted anyway is found (number / date / bool / error in place of the text); CRLF vs LF is not a difference
 *   - Localize: the English "General" is translated to this Excel's name ("G/標準" in Japanese Excel rejects the English word)
 */
import { describe, expect, it } from 'vitest';
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { POWERSHELL } from '../../server/excel/run-job.mjs';

const CS = path.resolve(__dirname, '..', '..', 'server', 'excel', 'XlRunnerHelper.cs');
const SCRIPT = path.join(__dirname, 'ps', 'helper-check.ps1');

interface BuildResult {
  cells: Array<Array<{ kind: number; val: unknown; prefix: boolean }>>;
  prefixed: number;
  riskyPlain: number[];
  bad: number[];
}
const KEmpty = 0;
const KText = 1;
const KNumber = 2;

function runCases(cases: unknown[]): Array<BuildResult & { fmt?: string }> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'em-helper-'));
  try {
    const cf = path.join(dir, 'cases.json');
    const out = path.join(dir, 'out.json');
    fs.writeFileSync(cf, JSON.stringify(cases), 'utf8');
    const r = spawnSync(POWERSHELL, ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', SCRIPT, '-Cs', CS, '-Cases', cf, '-Out', out], {
      encoding: 'utf8',
      windowsHide: true,
      timeout: 120_000,
    });
    if (r.status !== 0 || !fs.existsSync(out)) throw new Error(`helper-check failed (${r.status}): ${r.stdout}\n${r.stderr}`);
    return JSON.parse(fs.readFileSync(out, 'utf8'));
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

describe.skipIf(process.platform !== 'win32')('XlRunnerHelper.cs (compiled by PowerShell, no Excel)', () => {
  const res = runCases([
    // 0: quote flag in every object form; plain strings stay plain
    {
      kind: 'build',
      rows: [
        [{ t: 's', v: '001', quote: true }],
        ['03-1234-5678'],
        ['=abc'],
        [{ t: 's', v: '1/2' }],
        [{ t: 's', v: '1/2', quote: false }],
        [{ s: 'TRUE', quote: true }],
        [{ t: 's', v: '@', numFmt: '@' }],
        [{ t: 's', v: '007', quote: true, numFmt: '@' }],
        [{ t: 's', v: '', quote: true }],
        [{ t: 'n', v: 5, quote: true }],
        ["'x"],
        [{ t: 's', v: 'abc', quote: true }],
      ],
    },
    // 1: mode 1 (prefixAll)
    { kind: 'build', mode: 1, rows: [['abc'], ['001']] },
    // 2: Verify
    {
      kind: 'build',
      rows: [['03-1234-5678'], ['=abc'], ['1/2'], ['TRUE'], ['a\r\nb'], ['abc'], [{ t: 's', v: '001', quote: true }], ['x']],
      read: [['03-1234-5678'], [-2146826259], [46024], [true], ['a\nb'], ['abc'], ['001'], [null]],
    },
    { kind: 'localize', fmt: 'General', general: 'G/std' },
    { kind: 'localize', fmt: 'general', general: 'G/std' },
    { kind: 'localize', fmt: '0.0_ ;General', general: 'G/std' },
    { kind: 'localize', fmt: '"General"0 General', general: 'G/std' },
    { kind: 'localize', fmt: 'General', general: 'General' },
    { kind: 'localize', fmt: 'yyyy/m/d', general: 'G/std' },
    { kind: 'localize', fmt: '@', general: 'G/std' },
  ]);

  it('quote === true puts the apostrophe in front, nothing else does (no guessing)', () => {
    const b = res[0];
    const d = b.cells.map((r) => r[0]);
    expect(d[0]).toMatchObject({ kind: KText, val: "'001", prefix: true });
    expect(d[1]).toMatchObject({ kind: KText, val: '03-1234-5678', prefix: false });
    expect(d[2]).toMatchObject({ kind: KText, val: '=abc', prefix: false });
    expect(d[3]).toMatchObject({ kind: KText, val: '1/2', prefix: false });
    expect(d[4]).toMatchObject({ kind: KText, val: '1/2', prefix: false });
    expect(d[5]).toMatchObject({ kind: KText, val: "'TRUE", prefix: true });
    expect(d[6]).toMatchObject({ kind: KText, val: '@', prefix: false }); // a text-format cell never needs one
    expect(d[7]).toMatchObject({ kind: KText, val: "'007", prefix: true }); // quote + "@": the contract says quotePrefix, exactly
    expect(d[8]).toMatchObject({ kind: KEmpty }); // nothing to prefix
    expect(d[9]).toMatchObject({ kind: KNumber, val: 5, prefix: false });
    expect(d[10]).toMatchObject({ kind: KText, val: "''x", prefix: true }); // Excel eats a leading apostrophe: one more
    expect(d[11]).toMatchObject({ kind: KText, val: "'abc", prefix: true }); // explicit, even if Excel would keep "abc" as text
    expect(b.prefixed).toBe(5);
  });

  it('textMode prefixAll still prefixes every string', () => {
    expect(res[1].cells.map((r) => r[0].val)).toEqual(["'abc", "'001"]);
  });

  it('riskyPlain = the plain strings Excel might convert (only used to remember their number format)', () => {
    // rows 1 "03-1234-5678" (digit first), 2 "=abc", 3 and 4 "1/2", 6 "@" (a superset heuristic: it decides nothing about the text)
    expect(res[0].riskyPlain).toEqual([1, 2, 3, 4, 6]);
  });

  it('Verify finds the strings Excel converted (number, date, TRUE, error) and ignores CRLF vs LF', () => {
    expect(res[2].bad).toEqual([1, 2, 3, 7]); // =abc (error), 1/2 (date), TRUE (bool), "x" read back as empty
  });

  it('Localize translates the English General (outside quotes) and leaves everything else alone', () => {
    expect(res.slice(3).map((r) => r.fmt)).toEqual(['G/std', 'G/std', '0.0_ ;G/std', '"General"0 G/std', 'General', 'yyyy/m/d', '@']);
  });
});

/**
 * Round 5: text that came from an untrusted file must never become a live external call when it is written to a cell:
 * DDE (=cmd|' /C calc'!A0), external workbook references (also UNC / URL), WEBSERVICE / RTD / CALL / EXEC. The runner writes such
 * a formula - and such a typed text - as plain text with the quote prefix and warns UNSAFE_FORMULA.
 */
describe.skipIf(process.platform !== 'win32')('XlRunnerHelper.cs: UnsafeFormula / neutralized cells (compiled by PowerShell, no Excel)', () => {
  // shared with src/core/export/comJob.ts isUnsafeFormula (tests/server/comJob.test.ts): one rule, two implementations, the same vectors
  const VECTORS = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'unsafe-formulas.json'), 'utf8')) as { dangerous: string[]; harmless: string[] };
  const DANGEROUS = VECTORS.dangerous;
  const HARMLESS = VECTORS.harmless;
  const res = runCases([
    ...DANGEROUS.map((f) => ({ kind: 'unsafe', f })),
    ...HARMLESS.map((f) => ({ kind: 'unsafe', f })),
    // build: bare text, typed formulas, Lotus-style signs
    {
      kind: 'build',
      rows: [
        ["=cmd|' /C calc'!A0"], // 0 bare text that Excel would enter as a DDE formula
        [{ f: "=cmd|' /C calc'!A0" }], // 1 a formula object
        [{ f: '=SUM(A1:A3)' }], // 2 stays a formula
        ["+cmd|' /C calc'!A0"], // 3
        ['-cmd|x'], // 4
        ['=1+1'], // 5 stays plain text as before (the runner verifies it afterwards)
        ['=abc'], // 6
        [{ f: "='[x.xlsx]S'!A1", dyn: true }], // 7 dynamic-array formula object
        [{ t: 's', v: "=cmd|' /C calc'!A0", quote: true }], // 8 already quoted: one prefix only, not counted as neutralized
        [{ f: '=HYPERLINK("http://example.com","x")' }], // 9 stays a formula
        ['plain | text'], // 10 no leading formula character
        ['@SUM(A1)'], // 11 a plain "@..." text without danger stays plain
      ],
    },
  ]);
  const unsafe = res.slice(0, DANGEROUS.length + HARMLESS.length) as unknown as Array<{ unsafe: boolean }>;
  const b = res[res.length - 1] as BuildResult & { neutralized: number[] };

  it('classifies DDE, external workbook references and call-out functions as unsafe', () => {
    DANGEROUS.forEach((f, i) => expect(unsafe[i].unsafe, f).toBe(true));
  });
  it('leaves ordinary formulas alone (pipes in strings / sheet names, structured references, HYPERLINK, look-alike names)', () => {
    HARMLESS.forEach((f, i) => expect(unsafe[DANGEROUS.length + i].unsafe, f).toBe(false));
  });
  it('Build turns unsafe formulas and unsafe typed text into quoted text and reports the cells', () => {
    const d = b.cells.map((r) => r[0]);
    expect(d[0]).toMatchObject({ kind: KText, val: "'=cmd|' /C calc'!A0", prefix: true });
    expect(d[1]).toMatchObject({ kind: KText, val: "'=cmd|' /C calc'!A0", prefix: true });
    expect(d[2]).toMatchObject({ kind: 4, val: '=SUM(A1:A3)', prefix: false }); // KFormula
    expect(d[3]).toMatchObject({ kind: KText, prefix: true });
    expect(d[4]).toMatchObject({ kind: KText, prefix: true });
    expect(d[5]).toMatchObject({ kind: KText, val: '=1+1', prefix: false });
    expect(d[6]).toMatchObject({ kind: KText, val: '=abc', prefix: false });
    expect(d[7]).toMatchObject({ kind: KText, prefix: true }); // the dynamic-array route is closed as well
    expect(d[8]).toMatchObject({ kind: KText, val: "'=cmd|' /C calc'!A0", prefix: true });
    expect(d[9]).toMatchObject({ kind: 4 });
    expect(d[10]).toMatchObject({ kind: KText, val: 'plain | text', prefix: false });
    expect(d[11]).toMatchObject({ kind: KText, val: '@SUM(A1)', prefix: false });
    expect(b.neutralized).toEqual([0, 1, 3, 4, 7]);
  });
});
