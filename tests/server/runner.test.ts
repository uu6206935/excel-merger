/**
 * Excel runner packaging (no Excel needed): build-runner merge, the built excel-runner.ps1 (BOM, CRLF, up to date),
 * launcher .bat rules, PowerShell invocation (ScriptBlock form works under a Restricted execution policy),
 * job validation errors in Japanese, runRunner with a fake runner (cancel, crash).
 */
import { describe, expect, it } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { mergeRunner, runnerIsStale, buildRunnerFile, PLACEHOLDER } from '../../server/excel/merge-runner.mjs';
import { POWERSHELL, childEnv, runRunner, runnerCommand } from '../../server/excel/run-job.mjs';
import { spawnSync } from 'node:child_process';
import { FAKE_RUNNER, ROOT } from './helpers';

const RUNNER_DIR = path.join(ROOT, 'server', 'excel');
const RUNNER = path.join(RUNNER_DIR, 'excel-runner.ps1');
const WIN = process.platform === 'win32';

describe('build-runner (merge-runner.mjs)', () => {
  it('merges the C# helper into the placeholder with CRLF line endings', () => {
    const out = mergeRunner(`\uFEFFline1\n$x = @'\n${PLACEHOLDER}\n'@\nend\n`, '\uFEFFclass A {}\r\n// ü $&\n\n');
    expect(out).toBe("line1\r\n$x = @'\r\nclass A {}\r\n// ü $&\r\n'@\r\nend\r\n");
    expect(() => mergeRunner('no placeholder', 'x')).toThrow(/placeholder/);
    expect(() => mergeRunner(`${PLACEHOLDER}`, "a\n'@ bad")).toThrow(/'@/);
  });

  it('the committed excel-runner.ps1 is built, up to date, UTF-8 with BOM and CRLF', () => {
    expect(runnerIsStale(RUNNER_DIR)).toBe(false);
    const b = fs.readFileSync(RUNNER);
    expect([b[0], b[1], b[2]]).toEqual([0xef, 0xbb, 0xbf]);
    const text = b.subarray(3).toString('utf8');
    expect(text).not.toMatch(/[^\r]\n/);
    expect(text).toContain('namespace XlRunner');
    expect(text).not.toContain(PLACEHOLDER);
    const expected = mergeRunner(fs.readFileSync(path.join(RUNNER_DIR, 'excel-runner.src.ps1'), 'utf8'), fs.readFileSync(path.join(RUNNER_DIR, 'XlRunnerHelper.cs'), 'utf8'));
    expect(text).toBe(expected);
  });

  it('buildRunnerFile writes a fresh copy atomically', () => {
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'em-runner-'));
    try {
      fs.copyFileSync(path.join(RUNNER_DIR, 'excel-runner.src.ps1'), path.join(tmp, 'excel-runner.src.ps1'));
      fs.copyFileSync(path.join(RUNNER_DIR, 'XlRunnerHelper.cs'), path.join(tmp, 'XlRunnerHelper.cs'));
      expect(runnerIsStale(tmp)).toBe(true);
      const r = buildRunnerFile(tmp);
      expect(r.bytes).toBe(fs.statSync(RUNNER).size);
      expect(runnerIsStale(tmp)).toBe(false);
      expect(fs.readdirSync(tmp).filter((f) => f.endsWith('.tmp'))).toEqual([]);
    } finally {
      fs.rmSync(tmp, { recursive: true, force: true });
    }
  });
});

describe('runner sources: round 3 contract (static checks, no Excel)', () => {
  const src = fs.readFileSync(path.join(RUNNER_DIR, 'excel-runner.src.ps1'), 'utf8');
  const cs = fs.readFileSync(path.join(RUNNER_DIR, 'XlRunnerHelper.cs'), 'utf8');

  it('every MOVE_VERIFY_FAILED failure tells the user to use the built-in engine (the UI falls back for this code)', () => {
    const fails = src.split('\n').filter((l) => l.includes("Fail 'MOVE_VERIFY_FAILED'"));
    expect(fails.length).toBeGreaterThanOrEqual(5);
    for (const l of fails) expect(l, l.trim().slice(0, 80)).toMatch(/UseBuiltIn|内蔵エンジン/);
    expect(src).toMatch(/\$script:UseBuiltIn = '[^']*内蔵エンジン[^']*'/);
  });

  it('the text of a cell is never guessed: only the quote flag (or a leading apostrophe) gives the apostrophe', () => {
    expect(cs).toMatch(/dict\.TryGetValue\("quote"/);
    const putText = cs.slice(cs.indexOf('static void PutText'), cs.indexOf('public static bool WasPrefixed'));
    expect(putText.length).toBeGreaterThan(100);
    expect(putText).not.toMatch(/Risky\(/);
    expect(putText).toMatch(/quote \|\| s\[0\] == '\\''/);
  });

  it('the runner removes its own temp files (working copy, saved workbook, scratch folder) in the final block, also after a failure', () => {
    expect(src).toMatch(/\[void\]\$script:TempFiles\.Add\(\$copy\)/);
    expect(src).toMatch(/\[void\]\$script:TempFiles\.Add\(\$tmpOut\)/);
    expect(src).toMatch(/\$script:OwnWork = \$work/);
    const finallyBlock = src.slice(src.lastIndexOf('} finally {'));
    expect(finallyBlock).toMatch(/Remove-RunnerTemp/);
    expect(finallyBlock.indexOf('Stop-Excel')).toBeLessThan(finallyBlock.indexOf('Remove-RunnerTemp')); // Excel first: it holds the files
  });

  it('round 5: no waiting for Excel to exit, PID + start time reported, the upload opened in place, unsafe formulas written as text', () => {
    expect(src).toContain('[int]$QuitWaitMs = 0,');
    expect(src).toContain("Emit ([ordered]@{ type = 'excel'; pid = $script:XlPid; started = $startedMs })");
    expect(src).toContain('$ownCopy = -not (Test-SamePath $inPath $copy)');
    expect(src).toContain('$direct = Test-SamePath ([IO.Path]::GetDirectoryName($output)) $work');
    expect(src).toContain("Warn 'UNSAFE_FORMULA'");
    expect(src).toContain('-notmatch \'^\\$?[A-Za-z]{1,3}\\$?[0-9]{1,7}$\'');
    expect(cs).toContain('public static bool UnsafeFormula(string f)');
    expect(cs).toContain('if (UnsafeFormula(f))');
    expect(cs).toContain('StartsLikeFormula(s) && UnsafeFormula(s)');
    // the result line must not wait for Excel: Stop-Excel waits only when a caller asks for it
    const stop = src.slice(src.indexOf('function Stop-Excel'), src.indexOf('# ---------------------------------------------------------------------------- workbook helpers'));
    expect(stop).toContain('if ([int]$script:QuitWaitMs -gt 0)');
  });

  it("'General' is translated for Excel (the ja-JP Excel rejects the English word through COM)", () => {
    expect(src).toMatch(/International\(26\)/);
    expect(src).toMatch(/\[XlRunner\.Cells\]::Localize\(/);
  });
});

describe('launchers', () => {
  for (const bat of ['start.bat', 'build.bat']) {
    it(`${bat} is ASCII-only with CRLF line endings and no BOM`, () => {
      const b = fs.readFileSync(path.join(ROOT, bat));
      const s = b.toString('latin1');
      expect(/[^\x00-\x7f]/.test(s)).toBe(false);
      expect(/[^\r]\n/.test(s)).toBe(false);
      expect(s.startsWith('@echo off')).toBe(true);
      // no echo inside ( ... ) blocks: an unescaped ")" would end the block
      expect(/\([^)]*\becho\b/i.test(s.replace(/^rem .*$/gim, ''))).toBe(false);
    });
  }
  it('start.bat tries the exe, the runtime layout, installed node, then standalone', () => {
    const s = fs.readFileSync(path.join(ROOT, 'start.bat'), 'latin1');
    const order = ['ExcelMerger.exe', 'runtime\\node.exe', 'server\\server.mjs', 'dist\\index.html'].map((k) => s.indexOf(k));
    expect(order.every((i) => i > 0)).toBe(true);
    expect([...order].sort((a, b) => a - b)).toEqual(order);
    expect(s).toContain('--detach');
    expect(s).toContain('LSS 18');
  });
});

describe('PowerShell invocation', () => {
  it('builds a ScriptBlock command with PowerShell quoting (Japanese, spaces, apostrophes)', () => {
    const { cmd, args } = runnerCommand("C:\\ア プリ\\it's\\excel-runner.ps1", [{ name: 'JobPath', value: "D:\\j o'b.json" }], 'scriptblock');
    expect(cmd.toLowerCase()).toContain('powershell.exe');
    expect(args.slice(0, 3)).toEqual(['-NoLogo', '-NoProfile', '-NonInteractive']);
    expect(args).not.toContain('-ExecutionPolicy');
    const c = args[args.length - 1];
    expect(c).toContain("ReadAllText('C:\\ア プリ\\it''s\\excel-runner.ps1'");
    expect(c).toContain("-JobPath 'D:\\j o''b.json'");
    const file = runnerCommand('C:\\r.ps1', [{ name: 'Probe', switch: true }], 'file');
    expect(file.args).toEqual(['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', 'C:\\r.ps1', '-Probe']);
  });

  it('childEnv: an override replaces every case variant of the name (Windows env names are case-insensitive)', () => {
    // vitest workers see inherited names upper-cased ('PSEXECUTIONPOLICYPREFERENCE' from a parent PowerShell started
    // with -ExecutionPolicy Bypass); a plain {...process.env, PSExecutionPolicyPreference} spread kept both and the
    // child got 'Bypass' (the old flaky policy test)
    const base = { PSEXECUTIONPOLICYPREFERENCE: 'Bypass', Path: 'C:/x', Keep: 'y', Gone: undefined };
    const env = childEnv({ PSExecutionPolicyPreference: 'Restricted', path: 'C:/z', Drop: null }, base);
    if (WIN) {
      expect(env).toEqual({ Keep: 'y', PSExecutionPolicyPreference: 'Restricted', path: 'C:/z' });
    } else {
      expect(env.PSExecutionPolicyPreference).toBe('Restricted');
      expect(env.Keep).toBe('y');
    }
    expect(Object.keys(childEnv({}, base)).sort()).toEqual(['Keep', 'PSEXECUTIONPOLICYPREFERENCE', 'Path']);
  });

  const POLICY_ENV = { PSExecutionPolicyPreference: 'Restricted', XLRUNNER_CACHE: path.join(os.tmpdir(), 'em-runner-cache') };

  it.skipIf(!WIN)('the ScriptBlock form runs under a Restricted execution policy; bad jobs fail in Japanese', async () => {
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'em-policy-'));
    try {
      const jobFile = path.join(tmp, 'ジョブ.json');
      fs.writeFileSync(jobFile, JSON.stringify({ version: 1, format: 'xlsx', input: path.join(tmp, 'missing.xlsx'), output: path.join(tmp, 'o.xlsx'), ops: [] }));
      const env = POLICY_ENV;
      const sb = await runRunner({ runner: RUNNER, jobFile, invoke: 'scriptblock', env, timeoutMs: 120_000 });
      expect(sb.started).toBe(true);
      expect(sb.ok).toBe(false);
      expect(sb.code).toBe('FILE_NOT_FOUND');
      expect(sb.message).toContain('入力ファイルが見つかりません');
      expect(sb.excelPid).toBe(0); // Excel is never started for a bad job
      const auto = await runRunner({ runner: RUNNER, jobFile, invoke: 'auto', env, timeoutMs: 120_000 });
      expect(auto.started).toBe(true);
      expect(auto.code).toBe('FILE_NOT_FOUND');
      expect(auto.firstAttempt).toBeUndefined(); // the ScriptBlock form started: no -File fallback

      const bad = path.join(tmp, 'bad.json');
      const cases: Array<[unknown, RegExp]> = [
        [{ version: 2 }, /version/],
        [{ version: 1, format: 'csv', input: 'a.xlsx', output: 'b.csv' }, /xlsx か xlsm/],
        [{ version: 1, format: 'xlsx', input: 'a.xls', output: 'b.xlsx' }, /入力は/],
        [{ version: 1, format: 'xlsx', input: jobFile.replace('.json', '.xlsx'), output: 'b.xlsx', ops: [{ op: 'format C:' }] }, /未知の操作/],
      ];
      for (const [job, re] of cases) {
        fs.writeFileSync(bad, JSON.stringify(job));
        const r = await runRunner({ runner: RUNNER, jobFile: bad, env, timeoutMs: 120_000 });
        expect(r.code, JSON.stringify(job)).toBe('BAD_JOB');
        expect(r.message).toMatch(re);
        expect(r.exitCode).toBe(2);
      }
      fs.writeFileSync(bad, '{ not json');
      const nj = await runRunner({ runner: RUNNER, jobFile: bad, env, timeoutMs: 120_000 });
      expect(nj.code).toBe('BAD_JOB');
      expect(nj.message).toContain('JSON');
    } finally {
      fs.rmSync(tmp, { recursive: true, force: true });
    }
  }, 300_000);

  it.skipIf(!WIN)('the Restricted policy really applies to the runner processes: a plain -File is refused there', (ctx) => {
    // Which policy does a child actually get? A Group Policy (MachinePolicy / UserPolicy) wins over the process scope
    // set by PSExecutionPolicyPreference; then the refusal cannot be demonstrated on this machine (the ScriptBlock
    // test above still proves the runner works whatever the policy is).
    const env = childEnv(POLICY_ENV);
    const probe = spawnSync(POWERSHELL, ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', 'Get-ExecutionPolicy'], { env, encoding: 'latin1', windowsHide: true, timeout: 60_000 });
    const effective = String(probe.stdout || '').trim();
    ctx.skip(
      effective !== 'Restricted',
      `the process-scope execution policy does not apply in this environment (effective policy: "${effective || probe.error?.message || 'unknown'}"), so a Restricted policy cannot be simulated here`,
    );
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'em-policy-file-'));
    try {
      const jobFile = path.join(tmp, 'job.json');
      fs.writeFileSync(jobFile, JSON.stringify({ version: 1, format: 'xlsx', input: path.join(tmp, 'missing.xlsx'), output: path.join(tmp, 'o.xlsx'), ops: [] }));
      const plain = spawnSync(POWERSHELL, ['-NoLogo', '-NoProfile', '-NonInteractive', '-File', RUNNER, '-JobPath', jobFile], {
        env,
        encoding: 'latin1',
        windowsHide: true,
        timeout: 60_000,
      });
      expect(plain.status).not.toBe(0);
      expect(plain.stdout).not.toContain('"type":"start"');
    } finally {
      fs.rmSync(tmp, { recursive: true, force: true });
    }
  }, 120_000);
});

describe('runRunner (fake runner)', () => {
  const job = (ops: unknown[]) => {
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'em-fake-'));
    const input = path.join(tmp, 'in.xlsx');
    fs.writeFileSync(input, 'PK');
    const jobFile = path.join(tmp, 'job.json');
    fs.writeFileSync(jobFile, JSON.stringify({ version: 1, format: 'xlsx', input, output: path.join(tmp, 'out.xlsx'), ops }));
    return { tmp, jobFile };
  };

  it('streams events and resolves with the final result', async () => {
    const { tmp, jobFile } = job([{ op: 'moveCols' }]);
    const events: string[] = [];
    const r = await runRunner({ runner: FAKE_RUNNER, jobFile, onEvent: (e) => events.push(String(e.type)) });
    expect(r.ok).toBe(true);
    expect(events[0]).toBe('start');
    expect(events).toContain('progress');
    expect(events[events.length - 1]).toBe('result');
    fs.rmSync(tmp, { recursive: true, force: true });
  });

  it('cancels via AbortSignal and reports a crash without a final line in Japanese', async () => {
    const a = job([{ op: 'debugWait', seconds: 60 }]);
    const ctrl = new AbortController();
    setTimeout(() => ctrl.abort(), 500);
    const t = Date.now();
    const r = await runRunner({ runner: FAKE_RUNNER, jobFile: a.jobFile, signal: ctrl.signal });
    expect(Date.now() - t).toBeLessThan(10_000);
    expect(r.cancelled).toBe(true);
    expect(r.code).toBe('CANCELLED');
    expect(r.message).toBe('中止しました');
    const b = job([{ op: 'x', sheet: 'CRASH' }]);
    const c = await runRunner({ runner: FAKE_RUNNER, jobFile: b.jobFile });
    expect(c.ok).toBe(false);
    expect(c.code).toBe('RUNNER_CRASHED');
    expect(c.message).toMatch(/異常終了.*3/);
    for (const d of [a.tmp, b.tmp]) fs.rmSync(d, { recursive: true, force: true });
  });
});
