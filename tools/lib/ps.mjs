// Node helpers to run the tools' PowerShell/Excel scripts safely (owner: T).
// - every .ps1 must be UTF-8 WITH BOM (Windows PowerShell 5.1 reads BOM-less files as ANSI): ensureBom() fixes that
// - the scripts print ASCII-only JSON lines; {"type":"excel","pid":N} announces the script's OWN Excel PID so a
//   watchdog can kill exactly that process (never Excel by name) when the script hangs.
import { spawn, spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const SYSROOT = process.env.SystemRoot || process.env.windir || 'C:\\Windows';
export const POWERSHELL = path.join(SYSROOT, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');

/** Prepend a UTF-8 BOM to the given .ps1 files if it is missing. */
export function ensureBom(files) {
  for (const f of files) {
    const buf = fs.readFileSync(f);
    if (buf.length >= 3 && buf[0] === 0xef && buf[1] === 0xbb && buf[2] === 0xbf) continue;
    fs.writeFileSync(f, Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), buf]));
  }
}

function isAlive(pid) {
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

function killPid(pid) {
  if (!pid) return;
  try {
    process.kill(pid, 'SIGKILL');
  } catch {
    /* already gone */
  }
}

/** true when `pid` is (still) an EXCEL.EXE process — guards against PID reuse before killing. */
function isExcelPid(pid) {
  try {
    const r = spawnSync(path.join(SYSROOT, 'System32', 'tasklist.exe'), ['/FI', `PID eq ${pid}`, '/FO', 'CSV', '/NH'], {
      windowsHide: true,
      encoding: 'latin1',
      timeout: 5000,
    });
    return /^"EXCEL\.EXE"/im.test(r.stdout || '');
  } catch {
    return false;
  }
}

/**
 * Run a PowerShell script with a watchdog.
 * @param {string} script absolute path of the .ps1
 * @param {string[]} args  arguments after -File <script>
 * @param {{timeoutMs?: number, onLine?: (obj: any, raw: string) => void, cwd?: string}} [opts]
 * @returns {Promise<{code: number, lines: any[], stderr: string, timedOut: boolean, excelPids: number[]}>}
 */
export function runPs(script, args, opts = {}) {
  const timeoutMs = opts.timeoutMs ?? 10 * 60 * 1000;
  ensureBom([script]);
  return new Promise((resolve) => {
    const child = spawn(POWERSHELL, ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', script, ...args], {
      windowsHide: true,
      cwd: opts.cwd,
    });
    const lines = [];
    const excelPids = [];
    let buf = '';
    let stderr = '';
    let timedOut = false;
    const onData = (chunk) => {
      buf += chunk.toString('utf8');
      let i;
      while ((i = buf.indexOf('\n')) >= 0) {
        const raw = buf.slice(0, i).replace(/\r$/, '');
        buf = buf.slice(i + 1);
        if (!raw.trim()) continue;
        let obj = null;
        try {
          obj = JSON.parse(raw);
        } catch {
          obj = { type: 'text', text: raw };
        }
        if (obj && obj.type === 'excel' && obj.pid) excelPids.push(obj.pid);
        lines.push(obj);
        opts.onLine?.(obj, raw);
      }
    };
    child.stdout.on('data', onData);
    child.stderr.on('data', (c) => {
      stderr += c.toString('utf8');
    });
    const timer = setTimeout(() => {
      timedOut = true;
      for (const pid of excelPids) if (isExcelPid(pid)) killPid(pid); // our own Excel only (PID announced by the script)
      killPid(child.pid);
    }, timeoutMs);
    child.on('close', (code) => {
      clearTimeout(timer);
      if (buf.trim()) onData('\n');
      // PowerShell normally quits Excel itself; if our instance is still alive a few seconds later, kill that PID.
      const leftovers = excelPids.filter(isAlive);
      if (leftovers.length) {
        setTimeout(() => {
          for (const pid of leftovers) if (isAlive(pid) && isExcelPid(pid)) killPid(pid);
          resolve({ code: code ?? 1, lines, stderr, timedOut, excelPids });
        }, 4000);
      } else {
        resolve({ code: code ?? 1, lines, stderr, timedOut, excelPids });
      }
    });
  });
}
