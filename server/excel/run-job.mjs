/**
 * run-job.mjs - Node side of the "Excel本体" engine (Node >= 18, built-in modules only; bundled into the
 * server / SEA exe by esbuild, so: no top-level await, no dynamic import, no CLI code in here - see cli.mjs).
 *
 *   detectExcel()                     registry only (~50-150 ms, never starts Excel)
 *   runRunner({ runner, jobFile })    runs excel-runner.ps1 on a job file, streams its JSON lines
 *   killExcelPid(pid, startedMs)      kills ONLY that PID, and only if it is an EXCEL.EXE started at that time
 *
 * Excel launched through COM is NOT a child of powershell.exe (DCOM starts it), so killing PowerShell never
 * stops Excel: the runner reports its Excel PID first ({"type":"excel","pid":N,"started":ms}) and this module kills
 * exactly that PID on timeout / cancel, or if it is still alive a few seconds after the result has arrived.
 * The user's own Excel windows (other PIDs) are never touched.
 *
 * Round 5 (measured): a hidden Excel that has just saved keeps running after Quit() until powershell.exe lets go of it
 * (the PowerShell COM binder keeps references), and then shuts down by talking to Microsoft's cloud services (telemetry,
 * identity: 8-9 open HTTPS connections). When that network work stalls, Excel needs ~60 s to exit and powershell.exe
 * cannot exit before it (COM rundown): the runner's final {"type":"result"} line was already out after ~5 s, but the
 * child 'close' event - what this module used to wait for - came ~58 s later (4 of 12 jobs in a soak test).
 * `earlyResult: true` therefore settles the promise AS SOON AS the final line arrived; the process shutdown goes on in
 * the background (`closed` promise) under a watchdog that kills our own Excel PID after `exitGraceMs`.
 */
import { spawn, execFile } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const SYSROOT = process.env.SystemRoot || process.env.windir || 'C:\\Windows';
export const POWERSHELL = path.join(SYSROOT, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
const REG = path.join(SYSROOT, 'System32', 'reg.exe');
const TASKLIST = path.join(SYSROOT, 'System32', 'tasklist.exe');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function execText(file, args, timeoutMs = 5000) {
  return new Promise((resolve) => {
    try {
      execFile(file, args, { windowsHide: true, timeout: timeoutMs, encoding: 'buffer' }, (err, stdout, stderr) => {
        resolve({ code: err ? (typeof err.code === 'number' ? err.code : 1) : 0, stdout: decode(stdout), stderr: decode(stderr) });
      });
    } catch (e) {
      resolve({ code: 1, stdout: '', stderr: String(e && e.message) });
    }
  });
}

/** Console tools print in the OEM/ANSI code page (cp932 on Japanese Windows) unless told otherwise. */
export function decode(buf) {
  if (!buf || buf.length === 0) return '';
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(buf);
  } catch {
    /* not UTF-8 */
  }
  try {
    return new TextDecoder('shift_jis').decode(buf);
  } catch {
    return Buffer.from(buf).toString('latin1');
  }
}

async function regValue(key, name) {
  const args = ['query', key, ...(name ? ['/v', name] : ['/ve'])];
  const r = await execText(REG, args);
  if (r.code !== 0) return null;
  const line = r.stdout.split(/\r?\n/).find((l) => /\sREG_(SZ|EXPAND_SZ|DWORD)\s/.test(l));
  if (!line) return null;
  return line.replace(/^\s*\S.*?\s+REG_\w+\s+/, '').trim();
}

/**
 * "Is Excel usable?" from the registry only (no Excel start).
 * installed=false -> the UI disables the Excel本体 engine and uses the built-in (JS) engine.
 */
export async function detectExcel() {
  const forced = String(process.env.EXCEL_DETECT || '').trim().toLowerCase();
  if (forced === 'none') return { installed: false, reason: 'forced' };
  if (forced === 'installed') return { installed: true, version: '16.0 (test)', reason: 'forced' };
  if (process.platform !== 'win32') return { installed: false, reason: 'not-windows' };
  const [curVer, clsid, exe, c2r] = await Promise.all([
    regValue('HKCR\\Excel.Application\\CurVer'), // e.g. Excel.Application.16
    regValue('HKCR\\Excel.Application\\CLSID'),
    regValue('HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe'),
    regValue('HKLM\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration', 'VersionToReport'),
  ]);
  const exePath = exe ? exe.replace(/^"|"$/g, '') : null;
  const exeExists = exePath ? fs.existsSync(exePath) : false;
  const major = curVer ? Number(curVer.split('.').pop()) : NaN;
  const installed = Boolean(curVer && clsid && (exeExists || !exePath));
  const version = c2r || (Number.isFinite(major) ? `${major}.0` : undefined);
  return { installed, version: installed ? version : undefined, progId: curVer, exePath, exeExists };
}

export async function isExcelPid(pid) {
  if (!pid || !Number.isInteger(pid) || pid <= 0) return false;
  const r = await execText(TASKLIST, ['/FI', `PID eq ${pid}`, '/FO', 'CSV', '/NH']);
  return /"EXCEL\.EXE"/i.test(r.stdout);
}

/** Cheap "does a process with this PID exist" (no child process). A recycled PID also answers true: see killExcelPid. */
export function pidAlive(pid) {
  if (!pid || !Number.isInteger(pid) || pid <= 0) return false;
  try {
    process.kill(pid, 0);
    return true;
  } catch (e) {
    return !!(e && e.code === 'EPERM'); // exists, but belongs to someone else
  }
}

/** Start time (Unix ms) of a process, or null when it cannot be read. One PowerShell start (~0.3 s): only used before a kill. */
export async function processStartMs(pid) {
  if (!pid || !Number.isInteger(pid) || pid <= 0) return null;
  const cmd = `$ErrorActionPreference='Stop'; [Console]::Out.Write([long]([DateTimeOffset](Get-Process -Id ${pid}).StartTime).ToUnixTimeMilliseconds())`;
  const r = await execText(POWERSHELL, ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', cmd], 15000);
  const n = Number(String(r.stdout).trim());
  return r.code === 0 && Number.isFinite(n) && n > 0 ? n : null;
}

/**
 * Kill ONLY the given PID, and only if it is an EXCEL.EXE (never the user's other Excel windows). Windows recycles PIDs
 * quickly: when the runner told us when its Excel was started (`startedMs`), a process with a different start time is
 * somebody else's (for example the user's Excel that got the same PID after ours exited) and is left alone.
 */
export async function killExcelPid(pid, startedMs) {
  if (!(await isExcelPid(pid))) return false;
  if (startedMs) {
    const now = await processStartMs(pid);
    if (now !== null && Math.abs(now - startedMs) > 3000) return false;
  }
  try {
    process.kill(pid);
    return true;
  } catch {
    return false;
  }
}

async function waitPidGone(pid, ms) {
  const end = Date.now() + ms;
  while (Date.now() < end) {
    if (!pidAlive(pid)) return true;
    await sleep(200);
  }
  return !pidAlive(pid);
}

/**
 * PowerShell single-quoted string literal. PowerShell treats U+2018 / U+2019 / U+201A / U+201B like an ASCII apostrophe
 * (a file name such as "x’; calc; ’.xlsx" would otherwise end the literal), and inside a literal any of the five written
 * twice stands for itself.
 */
export const psQuote = (s) => `'${String(s).replace(/['\u2018\u2019\u201A\u201B]/g, (m) => m + m)}'`;

/**
 * process.env + overrides for a child process. Windows environment names are case-insensitive, but a plain
 * object spread keeps both 'PSEXECUTIONPOLICYPREFERENCE' (vitest workers upper-case inherited names) and
 * 'PSExecutionPolicyPreference', and Windows then uses whichever comes first in the sorted block - the override
 * could silently lose. Keys of `extra` replace every case variant; undefined / null removes the variable.
 */
export function childEnv(extra = {}, base = process.env) {
  const out = {};
  const win = process.platform === 'win32';
  const drop = new Set(Object.keys(extra || {}).map((k) => (win ? k.toUpperCase() : k)));
  for (const [k, v] of Object.entries(base || {})) {
    if (v === undefined || drop.has(win ? k.toUpperCase() : k)) continue;
    out[k] = v;
  }
  for (const [k, v] of Object.entries(extra || {})) if (v !== undefined && v !== null) out[k] = String(v);
  return out;
}

/**
 * Command line for the runner.
 *  - 'scriptblock' (default): the script text is loaded into a ScriptBlock, which is not subject to the
 *    execution policy (works under a GPO "Restricted" policy, no "-ExecutionPolicy Bypass" flag, R5 §6).
 *  - 'file': classic -ExecutionPolicy Bypass -File (fallback when the ScriptBlock form fails to start).
 *  - a .mjs/.js/.cjs runner (tests) runs with the current node.
 */
export function runnerCommand(runner, runnerArgs, invoke) {
  if (/\.(mjs|cjs|js)$/i.test(runner)) return { cmd: process.execPath, args: [runner, ...runnerArgs.flatMap(argToFake)] };
  const base = ['-NoLogo', '-NoProfile', '-NonInteractive'];
  if (invoke === 'file') return { cmd: POWERSHELL, args: [...base, '-ExecutionPolicy', 'Bypass', '-File', runner, ...runnerArgs.flatMap(argToFile)] };
  const psArgs = runnerArgs.map((a) => (a.switch ? `-${a.name}` : `-${a.name} ${psQuote(a.value)}`)).join(' ');
  const cmd = `$ErrorActionPreference='Stop'; & ([ScriptBlock]::Create([IO.File]::ReadAllText(${psQuote(runner)}, [Text.Encoding]::UTF8))) ${psArgs}; exit $LASTEXITCODE`;
  return { cmd: POWERSHELL, args: [...base, '-Command', cmd] };
}
function argToFile(a) {
  return a.switch ? [`-${a.name}`] : [`-${a.name}`, a.value];
}
function argToFake(a) {
  return a.switch ? [`--${a.name}`] : [a.value];
}

/**
 * Runs the runner. Resolves with its final {type:'result'|'error'} object plus bookkeeping (never rejects
 * for job errors; rejects only when the process cannot be started at all).
 *
 * options.runner      path of excel-runner.ps1 (or a fake .mjs in tests)
 * options.jobFile     job JSON path (omit with probe: true)
 * options.timeoutMs   whole job limit (default 30 min)
 * options.idleMs      max silence between JSON lines (default 5 min; a hidden modal dialog shows up as silence)
 * options.onEvent(ev) every JSON line (start / excel / progress / warn / result / error)
 * options.signal      AbortSignal to cancel (kills this job's Excel PID)
 * options.invoke      'auto' (scriptblock, then file if it did not start) | 'scriptblock' | 'file'
 * options.env         extra environment variables (e.g. XLRUNNER_CACHE, TEMP)
 * options.earlyResult true = resolve when the final line has arrived (see the header); the result then has
 *                     `earlyResult: true`, `exitCode: null`, `closed` (Promise of { exitCode, excelKilled, stderr, wallMs },
 *                     resolved when powershell.exe and our Excel are gone) and `reap()` (kill them now).
 *                     false (default) = resolve when the process has ended and Excel is gone (CLI, tests).
 * options.exitGraceMs earlyResult only: how long the process may keep running after the final line before our
 *                     Excel PID is killed (default 6000; Excel normally exits ~1 s after powershell.exe)
 * options.extraArgs   more runner arguments, e.g. [{ name: 'QuitWaitMs', value: '0' }]
 */
export async function runRunner(options) {
  const invoke = options.invoke || 'auto';
  const first = await runOnce({ ...options, invoke: invoke === 'file' ? 'file' : 'scriptblock' });
  if (invoke === 'auto' && !first.started && !first.cancelled && !/\.(mjs|cjs|js)$/i.test(options.runner)) {
    // PowerShell could not even start the ScriptBlock (unexpected host restrictions): try -File once.
    const second = await runOnce({ ...options, invoke: 'file' });
    return { ...second, firstAttempt: { exitCode: first.exitCode, stderr: first.stderr } };
  }
  return first;
}

function runOnce({ runner, jobFile, probe = false, timeoutMs = 30 * 60_000, idleMs = 5 * 60_000, onEvent = () => {}, signal, invoke, env, earlyResult = false, exitGraceMs = 6000, extraArgs = [] }) {
  return new Promise((resolve, reject) => {
    const runnerArgs = [...(probe ? [{ name: 'Probe', switch: true }] : [{ name: 'JobPath', value: jobFile }]), ...extraArgs];
    const { cmd, args } = runnerCommand(runner, runnerArgs, invoke);
    let child;
    try {
      child = spawn(cmd, args, { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'], env: childEnv(env) });
    } catch (e) {
      reject(e);
      return;
    }
    let excelPid = 0;
    let excelStartedMs = 0;
    let final = null;
    let started = false;
    let buf = '';
    const stderrChunks = [];
    let killedFor = null;
    let settled = false;
    let closedFlag = false;
    let excelKilled = false;
    let graceTimer = null;
    let reaping = null;
    const t0 = Date.now();
    let resolveClosed;
    const closed = new Promise((r) => {
      resolveClosed = r;
    });
    let idleTimer = null;
    const armIdle = () => {
      clearTimeout(idleTimer);
      idleTimer = setTimeout(() => abort('idle'), idleMs);
    };
    const totalTimer = setTimeout(() => abort('timeout'), timeoutMs);
    armIdle();

    /** Kills our Excel PID (verified) and, a moment later, the powershell.exe child if it is still there. Idempotent. */
    function reap() {
      if (closedFlag) return Promise.resolve();
      if (reaping) return reaping;
      reaping = (async () => {
        if (excelPid && pidAlive(excelPid)) excelKilled = (await killExcelPid(excelPid, excelStartedMs)) || excelKilled;
        await sleep(1500);
        if (!closedFlag) {
          try {
            child.kill();
          } catch {
            /* already gone */
          }
        }
      })();
      return reaping;
    }

    async function abort(reason) {
      if (settled) {
        // the job is over (result delivered): only the process shutdown is still going on - end it now
        reap();
        return;
      }
      if (killedFor) return;
      killedFor = reason;
      // Kill our Excel first: the pending COM call in PowerShell then fails (RPC_E_DISCONNECTED) and the runner
      // exits by itself. Then make sure powershell.exe is gone too.
      if (excelPid) excelKilled = (await killExcelPid(excelPid, excelStartedMs)) || excelKilled;
      setTimeout(() => {
        try {
          child.kill();
        } catch {
          /* already gone */
        }
      }, excelPid ? 3000 : 0);
    }
    const onAbort = () => abort('cancelled');
    if (signal) {
      if (signal.aborted) onAbort();
      else signal.addEventListener('abort', onAbort, { once: true });
    }

    const minutes = (ms) => (ms >= 60_000 ? `${Math.round(ms / 60_000)} 分` : `${Math.round(ms / 1000)} 秒`);
    /** The object runRunner resolves with. `code` = exit code of powershell.exe (null while it is still running). */
    function compose(code) {
      const stderr = decode(Buffer.concat(stderrChunks)).trim();
      let res = final;
      if (!res) {
        res = {
          type: 'error',
          ok: false,
          code: killedFor === 'cancelled' ? 'CANCELLED' : killedFor ? 'TIMEOUT' : 'RUNNER_CRASHED',
          message:
            killedFor === 'idle'
              ? `Excel が ${minutes(idleMs)}間応答しなかったため中止しました（Excel が確認画面を出して止まった可能性があります）`
              : killedFor === 'timeout'
                ? `処理が ${minutes(timeoutMs)}を超えたため中止しました`
                : killedFor === 'cancelled'
                  ? '中止しました'
                  : `Excel本体エンジンが異常終了しました (終了コード ${code})${stderr ? `: ${stderr.split(/\r?\n/)[0].slice(0, 200)}` : ''}`,
        };
      } else if (killedFor) {
        res = { ...res, ok: false, code: killedFor === 'cancelled' ? 'CANCELLED' : 'TIMEOUT', message: killedFor === 'cancelled' ? '中止しました' : res.message };
      }
      return {
        ...res,
        exitCode: code,
        excelPid,
        excelStartedMs: excelStartedMs || undefined,
        excelKilled,
        started,
        cancelled: killedFor === 'cancelled',
        stderr: stderr || undefined,
        wallMs: Date.now() - t0,
      };
    }

    /** earlyResult: the final line is here - deliver it now, let the process shutdown finish in the background. */
    function settleEarly() {
      if (settled || closedFlag) return;
      settled = true;
      clearTimeout(totalTimer);
      clearTimeout(idleTimer);
      const res = compose(null);
      res.earlyResult = true;
      res.closed = closed;
      res.reap = reap;
      graceTimer = setTimeout(() => reap(), exitGraceMs);
      resolve(res);
    }

    child.stdout.setEncoding('utf8'); // runner output is pure ASCII JSON
    child.stdout.on('data', (chunk) => {
      buf += chunk;
      let nl;
      while ((nl = buf.indexOf('\n')) >= 0) {
        const line = buf.slice(0, nl).trim();
        buf = buf.slice(nl + 1);
        if (!line) continue;
        let ev;
        try {
          ev = JSON.parse(line);
        } catch {
          ev = { type: 'log', text: line };
        }
        if (!settled) armIdle();
        if (ev.type === 'start') started = true;
        if (ev.type === 'excel' && Number.isInteger(ev.pid) && ev.pid > 0) {
          excelPid = ev.pid;
          if (Number.isFinite(ev.started) && ev.started > 0) excelStartedMs = Number(ev.started);
        }
        if (ev.type === 'result' || ev.type === 'error') final = ev;
        try {
          onEvent(ev);
        } catch {
          /* listener errors must not break the job */
        }
        if (final && earlyResult) settleEarly();
      }
    });
    child.stderr.on('data', (c) => stderrChunks.push(c));
    child.on('error', (e) => {
      clearTimeout(totalTimer);
      clearTimeout(idleTimer);
      clearTimeout(graceTimer);
      if (!settled) {
        settled = true;
        reject(e);
      }
      resolveClosed({ exitCode: null, excelKilled, error: String(e && e.message), wallMs: Date.now() - t0 });
    });
    child.on('close', async (code) => {
      closedFlag = true;
      clearTimeout(totalTimer);
      clearTimeout(idleTimer);
      clearTimeout(graceTimer);
      signal?.removeEventListener('abort', onAbort);
      // Excel usually exits ~0.5-1 s after powershell.exe ends; give it a few seconds, then kill our PID only.
      if (excelPid && !(await waitPidGone(excelPid, killedFor ? 3000 : settled ? 4000 : 8000))) excelKilled = (await killExcelPid(excelPid, excelStartedMs)) || excelKilled;
      const res = compose(code);
      if (!settled) {
        settled = true;
        resolve(res);
      }
      resolveClosed({ exitCode: code, excelKilled, stderr: res.stderr, wallMs: Date.now() - t0 });
    });
  });
}
