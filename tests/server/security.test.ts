/**
 * Round 5 security review of server/server.mjs, server/safe.mjs and the PowerShell hand-over (no Excel needed):
 * hostile file names (device names, ADS, bidi characters), hostile paths for /api/open (junctions, UNC, device paths, ADS), Host /
 * Origin / Sec-Fetch guards (also for reads and the SSE stream), request smuggling, oversize bodies, job flooding, job-id guessing,
 * hostile op content (quotes, $( ), backticks, newlines, external references in a cell address), PowerShell quoting with the
 * typographic quotes PowerShell treats like apostrophes, disk-full handling.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import fs from 'node:fs';
import http from 'node:http';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { buildOpenCommand, cleanOutputName, isInside, isReservedDeviceName, textCheckOpenPath, validateRunnerOps } from '../../server/safe.mjs';
import { POWERSHELL, psQuote, runRunner } from '../../server/excel/run-job.mjs';
import { makePackage, makeSandbox, openEvents, raw, sleep, startServer, type RunningServer, type Sandbox } from './helpers';

const WIN = process.platform === 'win32';

// ---------------------------------------------------------------------------------------------------------
// safe.mjs: file names
// ---------------------------------------------------------------------------------------------------------
describe('cleanOutputName: hostile names from the x-file-name header', () => {
  const ok = (raw: string) => {
    const r = cleanOutputName(raw);
    if (!r.name) throw new Error(`refused: ${raw} -> ${r.error}`);
    return r.name;
  };

  it('Windows device names are prefixed, also behind a second extension, trailing spaces and the legacy superscript digits', () => {
    const cases: Array<[string, string]> = [
      ['CON.xlsx', '_CON.xlsx'],
      ['con.xlsx', '_con.xlsx'],
      ['CON.tar.xlsx', '_CON.tar.xlsx'], // "CON.anything" is the console device as well
      ['NUL .xlsx', '_NUL .xlsx'],
      ['PRN.xlsx', '_PRN.xlsx'],
      ['AUX.xlsx', '_AUX.xlsx'],
      ['COM1.xlsx', '_COM1.xlsx'],
      ['LPT9.xlsx', '_LPT9.xlsx'],
      ['COM\u00b9.xlsx', '_COM\u00b9.xlsx'],
      ['LPT\u00b2.xlsx', '_LPT\u00b2.xlsx'],
      ['CONIN$.xlsx', '_CONIN$.xlsx'],
      ['CONOUT$.xlsx', '_CONOUT$.xlsx'],
      ['CLOCK$.xlsx', '_CLOCK$.xlsx'],
      ['COM10.xlsx', 'COM10.xlsx'], // not a device (only COM1-9)
      ['console.xlsx', 'console.xlsx'],
      ['CONSOLE.tar.xlsx', 'CONSOLE.tar.xlsx'],
      ['x.CON.xlsx', 'x.CON.xlsx'], // only the part before the FIRST dot counts
    ];
    for (const [input, want] of cases) expect(ok(input), input).toBe(want);
    expect(isReservedDeviceName('nul')).toBe(true);
    expect(isReservedDeviceName('nullable')).toBe(false);
  });

  it('alternate data streams, drive letters, folders and traversal never produce a name', () => {
    for (const bad of ['report.xlsx:evil.exe', 'a.xlsx::$DATA', 'a.xlsx:Zone.Identifier', 'C:evil.xlsx', 'C:\\evil.xlsx', '..\\evil.xlsx', '../evil.xlsx', 'a/b.xlsx', '\\\\server\\share\\x.xlsx', '/etc/passwd.xlsx', 'a..b.xlsx']) {
      const r = cleanOutputName(bad);
      expect(r.name, bad).toBeUndefined();
      expect(typeof r.error, bad).toBe('string');
    }
  });

  it('trailing dots / spaces, control and invisible characters are removed or replaced', () => {
    expect(ok('x.xlsx.')).toBe('x.xlsx');
    expect(ok('x.xlsx  ')).toBe('x.xlsx');
    expect(ok('x.xlsx . . ')).toBe('x.xlsx');
    expect(ok('x.xlsx\u3000')).toBe('x.xlsx');
    expect(ok('a\u0000b\nc.xlsx')).toBe('a_b_c.xlsx');
    expect(ok('\u202egnp.xlsx')).toBe('gnp.xlsx'); // right-to-left override (display spoofing)
    expect(ok('a\u200bb\ufeffc.xlsx')).toBe('abc.xlsx');
    expect(ok('a<b>c|d?e*f"g.xlsx')).toBe('a_b_c_d_e_f_g.xlsx');
    expect(ok('  .hidden.xlsx')).toBe('.hidden.xlsx');
    expect(cleanOutputName('.xlsx').name).toBeUndefined(); // no stem: a dotfile has no extension for Node
    expect(cleanOutputName('   ').error).toMatch(/ファイル名/);
    expect(cleanOutputName(undefined).error).toMatch(/ファイル名/);
    expect(cleanOutputName(['a.xlsx'] as unknown).error).toMatch(/ファイル名/);
    expect(ok('\ud800.xlsx')).toBe('_.xlsx'); // lone surrogate
    expect(ok('x.XLSX')).toBe('x.xlsx');
  });

  it('only .xlsx / .xlsm survive, whatever else is appended', () => {
    for (const bad of ['a.exe', 'a.xlsx.exe', 'a.xlsx.bat', 'a.bat.xlsx.txt', 'a.xls', 'a.xlsb', 'a', 'a.xlsx.lnk', 'a.xlsx:.exe']) expect(cleanOutputName(bad).name, bad).toBeUndefined();
  });

  it('property: 3000 random hostile strings either fail or give a plain, safe Windows file name', () => {
    const alphabet = ['a', 'Z', '1', ' ', '.', '..', ':', '/', '\\', '*', '?', '"', '<', '>', '|', 'CON', 'NUL', 'COM1', '.xlsx', '.xlsm', '\u202e', '\u200b', '\u0000', '\n', '日本', '\ud83d\ude00', '\ud83d', '$', '%', '~', '\u00b9'];
    let seed = 12345;
    const rnd = () => (seed = (seed * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
    for (let n = 0; n < 3000; n++) {
      let s = '';
      const len = 1 + Math.floor(rnd() * 12);
      for (let i = 0; i < len; i++) s += alphabet[Math.floor(rnd() * alphabet.length)];
      if (rnd() < 0.7) s += rnd() < 0.5 ? '.xlsx' : '.XLSM';
      if (rnd() < 0.2) {
        try {
          s = encodeURIComponent(s);
        } catch {
          /* a lone surrogate cannot be URI-encoded: leave the raw string */
        }
      }
      const r = cleanOutputName(s);
      if (!r.name) continue;
      const name = r.name;
      expect(name, s).toMatch(/\.(xlsx|xlsm)$/);
      expect(/[<>:"|?*\\/\u0000-\u001f\u007f]/.test(name), `${JSON.stringify(s)} -> ${JSON.stringify(name)}`).toBe(false);
      expect(/[. ]$/.test(name.replace(/\.(xlsx|xlsm)$/, '.'))).toBe(true); // the stem is followed by the extension only
      expect(isReservedDeviceName(name.slice(0, -5)), name).toBe(false);
      expect(name.length).toBeLessThanOrEqual(156);
      expect(path.basename(name)).toBe(name);
      expect(name.normalize('NFC')).toBe(name);
    }
  });
});

describe('textCheckOpenPath / isInside / buildOpenCommand', () => {
  it('refuses UNC paths, device paths, alternate data streams, control characters and non-strings before touching the disk', () => {
    const refused = ['\\\\server\\share\\x.xlsx', '//server/share/x.xlsx', '\\\\?\\C:\\out\\x.xlsx', '\\\\.\\pipe\\x', 'C:\\out\\x.xlsx:stream', 'C:\\out\\x.txt:y.xlsx', 'C:\\out\\x.xlsx::$DATA', 'x.xlsx:y'];
    for (const p of refused) expect(textCheckOpenPath(p)?.status, p).toBe(403);
    for (const p of ['C:\\out\\x\u0000.xlsx', 'C:\\out\\x\n.xlsx', 'a'.repeat(1001), '']) expect(textCheckOpenPath(p)?.status, JSON.stringify(p).slice(0, 40)).toBe(400);
    for (const p of [undefined, null, 5, {}, [], true]) expect(textCheckOpenPath(p)?.status).toBe(400);
    for (const p of ['C:\\out\\x.xlsx', 'D:\\ダウンロード\\x y.xlsx', 'relative\\x.xlsx']) expect(textCheckOpenPath(p)).toBeNull();
  });

  it('isInside: strictly inside, no prefix trick, no traversal', () => {
    expect(isInside('C:\\out', 'C:\\out\\a.xlsx')).toBe(true);
    expect(isInside('C:\\out', 'C:\\out\\sub\\a.xlsx')).toBe(true);
    expect(isInside('C:\\out', 'C:\\out')).toBe(false);
    expect(isInside('C:\\out', 'C:\\output\\a.xlsx')).toBe(false);
    expect(isInside('C:\\out', 'C:\\out\\..\\x.xlsx')).toBe(false);
    expect(isInside('C:\\out', 'D:\\out\\a.xlsx')).toBe(false);
    if (WIN) expect(isInside('C:\\out', 'c:\\OUT\\a.xlsx')).toBe(true);
  });

  const HOSTILE_PATHS = [
    "C:\\out\\x’; Start-Process calc; ’.xlsx",
    "C:\\out\\it's.xlsx",
    'C:\\out\\a`b$(calc)c.xlsx',
    'C:\\out\\a&b^c%d!e.xlsx',
    'C:\\out\\a,b;c.xlsx',
    'C:\\out\\‘q‘ ‚r‚ ‛s‛.xlsx',
    'D:\\ダウンロード\\日本語 フォルダ\\統合 結果.xlsx',
  ];

  it('opening a file never goes through a shell string: no cmd.exe / start, the path travels in an environment variable', () => {
    for (const p of HOSTILE_PATHS) {
      const c = buildOpenCommand(p, false, 'C:\\Windows');
      expect(c.cmd.toLowerCase()).toBe('c:\\windows\\system32\\windowspowershell\\v1.0\\powershell.exe');
      expect(c.args).toEqual(['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', 'Invoke-Item -LiteralPath $env:EM_OPEN_TARGET']);
      expect(c.args.join(' ')).not.toContain(p);
      expect(c.env).toEqual({ EM_OPEN_TARGET: p });
      expect(c.verbatim).toBe(false);
      expect(JSON.stringify(c).toLowerCase()).not.toContain('cmd.exe');
    }
    // folder selection: explorer.exe with the quoted path (a double quote cannot be part of a Windows file name; it is refused)
    const sel = buildOpenCommand('C:\\out\\a b.xlsx', true, 'C:\\Windows');
    expect(sel.cmd.toLowerCase()).toBe('c:\\windows\\explorer.exe');
    expect(sel.args).toEqual(['/select,"C:\\out\\a b.xlsx"']);
    for (const bad of ['C:\\out\\a" /e,"C:\\Windows.xlsx', 'C:\\out\\a\nb.xlsx', 'C:\\out\\a\u0000b.xlsx']) {
      expect(() => buildOpenCommand(bad, true, 'C:\\Windows'), JSON.stringify(bad)).toThrow();
      expect(() => buildOpenCommand(bad, false, 'C:\\Windows'), JSON.stringify(bad)).toThrow();
    }
  });

  it.skipIf(!WIN)('the real PowerShell receives every hostile path unchanged from the environment and runs nothing', () => {
    const marker = path.join(os.tmpdir(), `em-inj-${process.pid}-${Date.now()}.txt`);
    for (const p of HOSTILE_PATHS) {
      const c = buildOpenCommand(p, false, process.env.SystemRoot || 'C:\\Windows');
      // same command line, only the action is replaced by "print the code points of the target"
      const args = [...c.args.slice(0, -1), '($env:EM_OPEN_TARGET.ToCharArray() | ForEach-Object { [int]$_ }) -join \',\''];
      const r = spawnSync(POWERSHELL, args, { env: { ...process.env, ...c.env, INJECTION_MARKER: marker }, encoding: 'utf8', windowsHide: true, timeout: 60_000 });
      expect(r.status, r.stderr).toBe(0);
      expect(r.stdout.trim()).toBe([...p].flatMap((ch) => [...Array(ch.length).keys()].map((i) => ch.charCodeAt(i))).join(','));
    }
    expect(fs.existsSync(marker)).toBe(false);
  }, 120_000);

  it.skipIf(!WIN)('psQuote: apostrophes of every kind (incl. the typographic U+2018 / U+2019 / U+201A / U+201B) stay inside the literal', () => {
    const marker = path.join(os.tmpdir(), `em-inj2-${process.pid}-${Date.now()}.txt`);
    const samples = [
      "it's",
      'it’s',
      '‘x‘',
      'a‚b',
      'a‛b',
      `x’; New-Item -ItemType File -Path '${marker.replace(/'/g, "''")}'; ’y`,
      "x'; New-Item -ItemType File -Path '" + marker.replace(/'/g, "''") + "'; '",
      '$(New-Item -ItemType File -Path x)',
      'a`nb',
      'a"b',
      'a;b&c|d>e',
    ];
    for (const s of samples) {
      const cmd = `($(${psQuote(s)}).ToCharArray() | ForEach-Object { [int]$_ }) -join ','`;
      const r = spawnSync(POWERSHELL, ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', cmd], { encoding: 'utf8', windowsHide: true, timeout: 60_000 });
      expect(r.status, `${s}: ${r.stderr}`).toBe(0);
      expect(r.stdout.trim(), s).toBe([...Array(s.length).keys()].map((i) => s.charCodeAt(i)).join(','));
    }
    expect(fs.existsSync(marker)).toBe(false);
    expect(psQuote("a'b")).toBe("'a''b'");
    expect(psQuote('a’b')).toBe("'a’’b'");
    expect(psQuote('plain')).toBe("'plain'");
  }, 240_000);

  it.skipIf(!WIN)('a runner and a job file inside folders named like PowerShell code are still just paths (scriptblock invocation, real PowerShell)', async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'em-psq-'));
    const marker = path.join(root, 'PWNED.txt');
    try {
      const dir = path.join(root, `it's ’a’; New-Item -ItemType File -Path PWNED.txt; ‘b‘ $(1+1)`);
      fs.mkdirSync(dir, { recursive: true });
      const runner = path.join(dir, 'fake runner.ps1');
      fs.writeFileSync(runner, 'param([string]$JobPath)\n[Console]::Out.WriteLine(\'{"type":"start"}\')\n[Console]::Out.WriteLine(\'{"type":"result","ok":true}\')\n');
      const jobFile = path.join(dir, 'job ‘1’.json');
      fs.writeFileSync(jobFile, '{}');
      const events: string[] = [];
      const r = await runRunner({ runner, jobFile, invoke: 'scriptblock', onEvent: (e) => events.push(e.type), timeoutMs: 120_000 });
      expect(r.ok, JSON.stringify(r)).toBe(true);
      expect(r.started).toBe(true);
      expect(events).toEqual(['start', 'result']);
      expect(fs.existsSync(marker)).toBe(false);
      expect(fs.existsSync(path.join(process.cwd(), 'PWNED.txt'))).toBe(false);
    } finally {
      fs.rmSync(root, { recursive: true, force: true });
    }
  }, 180_000);
});

// ---------------------------------------------------------------------------------------------------------
// safe.mjs: the op whitelist
// ---------------------------------------------------------------------------------------------------------
describe('validateRunnerOps: the op log is rebuilt from known, typed fields', () => {
  const SHEET = "S'h\"e$(calc)`e\\t;&|<>%日本";
  const good = (op: Record<string, unknown>) => {
    const r = validateRunnerOps([op]);
    if (r.error) throw new Error(r.error);
    return r.ops![0];
  };
  const bad = (op: unknown, re: RegExp = /.+/) => {
    const r = validateRunnerOps([op]);
    expect(r.error, JSON.stringify(op).slice(0, 120)).toMatch(re);
  };

  it('keeps hostile sheet names and cell text as DATA (nothing is interpreted by anything before PowerShell reads the JSON file)', () => {
    const rows = [["=cmd|' /C calc'!A0", { f: '=HYPERLINK("http://evil","x")' }, { t: 's', v: "'; Start-Process calc; '", quote: true }, null, 5, true, { e: '#N/A' }, { t: 'd', v: 46000, numFmt: 'yyyy/m/d' }]];
    const op = good({ op: 'setCells', sheet: SHEET, at: 'C5', rows, textMode: 'prefixAll', toValues: true });
    expect(op).toEqual({ op: 'setCells', sheet: SHEET, at: 'C5', rows, textMode: 'prefixAll', toValues: true });
  });

  it('drops every field it does not know (a client cannot smuggle input / output / workDir into an op)', () => {
    const op = good({ op: 'moveCols', sheet: 'S', start: 2, count: 1, to: 5, method: 'clipFree', input: 'C:\\Windows\\win.ini', output: 'C:\\x.xlsx', workDir: 'C:\\', evil: { a: 1 }, __proto__x: 1 });
    expect(op).toEqual({ op: 'moveCols', sheet: 'S', start: 2, count: 1, to: 5, method: 'clipFree' });
    expect(good({ op: 'deleteRows', sheet: 'S', at: 3, count: 2, rows: [[1]] })).toEqual({ op: 'deleteRows', sheet: 'S', at: 3, count: 2 });
  });

  it('cell addresses are plain A1 addresses: no names, no external references (Range("\\\\host\\share\\[x.xlsx]S!A1") would reach for the network)', () => {
    for (const at of ["'\\\\evil\\share\\[x.xlsx]S'!A1", '[x.xlsx]S!A1', "S!A1", 'A1:B2', 'MyName', 'A1; calc', '', 'A0 ', '=A1', 'A1\n', 'ABCD1', 'A12345678', 5, null, ['A1']]) {
      bad({ op: 'setCells', sheet: 'S', at, rows: [[1]] }, /setCells の at/);
    }
    for (const at of ['A1', '$A$1', 'XFD1048576', 'c5']) expect(good({ op: 'setCells', sheet: 'S', at, rows: [[1]] }).at).toBe(at);
    bad({ op: 'unmerge', sheet: 'S', refs: ["'\\\\evil\\s\\[x.xlsx]S'!A1:B2"] }, /unmerge/);
    bad({ op: 'unmerge', sheet: 'S', refs: 'A1:B2' }, /unmerge/);
    expect(good({ op: 'unmerge', sheet: 'S', refs: ['A1:F1', 'B2'], centerAcross: true })).toEqual({ op: 'unmerge', sheet: 'S', refs: ['A1:F1', 'B2'], centerAcross: true });
  });

  it('sheet names: a string of 1-255 characters without control characters (a newline cannot end a command line)', () => {
    for (const sheet of ['', 'x'.repeat(256), 'a\nb', 'a\rb', 'a\u0000b', 'a\tb', 5, null, undefined, ['S'], { a: 1 }]) bad({ op: 'insertCols', sheet, at: 1, count: 1 }, /シート名/);
    expect(good({ op: 'insertCols', sheet: 'x'.repeat(255), at: 1, count: 1 }).sheet).toHaveLength(255);
  });

  it('numbers are integers in range, enums are enums', () => {
    for (const v of [0, -1, 1.5, 1048577, '1', null, undefined, NaN, Infinity, 1e999, {}, [], true]) {
      bad({ op: 'moveCols', sheet: 'S', start: v, count: 1, to: 3 }, /start \/ count \/ to/);
      bad({ op: 'deleteCols', sheet: 'S', at: v, count: 1 }, /at \/ count/);
    }
    bad({ op: 'moveCols', sheet: 'S', start: 1, count: 1, to: 3, method: 'rm -rf' }, /method/);
    bad({ op: 'insertCols', sheet: 'S', at: 1, count: 1, formatFrom: 'sideways' }, /formatFrom/);
    bad({ op: 'setColumnWidth', sheet: 'S', col: 1, width: 300, unit: 'chars' }, /width/);
    bad({ op: 'setColumnWidth', sheet: 'S', col: 1, width: -1 }, /width/);
    bad({ op: 'setColumnWidth', sheet: 'S', col: 1, width: 10, unit: 'furlongs' }, /unit/);
    bad({ op: 'setColumnWidth', sheet: 'S', col: 16385, width: 10 }, /col/);
    bad({ op: 'setCells', sheet: 'S', at: 'A1', rows: [[1]], textMode: 'eval' }, /textMode/);
    expect(good({ op: 'setColumnWidth', sheet: 'S', col: 3, width: 19.375, unit: 'stored' })).toEqual({ op: 'setColumnWidth', sheet: 'S', col: 3, count: 1, width: 19.375, unit: 'stored' });
  });

  it('cells: primitives and known-key objects only', () => {
    for (const cell of [{ f: { a: 1 } }, { cmd: 'x' }, { v: [1] }, [1], { numFmt: 'x'.repeat(256) }, { f: '=1', extra: 1 }, { t: 's', v: {} }, 1e999]) {
      bad({ op: 'setCells', sheet: 'S', at: 'A1', rows: [[cell]] }, /書き込めない形式/);
    }
    bad({ op: 'setCells', sheet: 'S', at: 'A1', rows: [1] }, /二重配列/);
    bad({ op: 'setCells', sheet: 'S', at: 'A1', rows: 'x' }, /配列/);
    expect(validateRunnerOps([{ op: 'setCells', sheet: 'S', at: 'A1', rows: [] }]).error).toBeUndefined();
  });

  it('unknown ops and the debug hook (only with XLRUNNER_ALLOW_DEBUG)', () => {
    for (const op of ['format', 'exec', '', undefined, 5, '__proto__', 'constructor']) bad({ op, sheet: 'S' }, /未知の操作/);
    bad({ op: 'debugWait', seconds: 1 }, /未知の操作/);
    expect(validateRunnerOps([{ op: 'debugWait', seconds: 2 }], { allowDebug: true }).ops?.[0]).toEqual({ op: 'debugWait', seconds: 2 });
    for (const x of [null, 5, 'op', [1]]) expect(validateRunnerOps([x as unknown]).error).toMatch(/オブジェクト/);
  });
});

// ---------------------------------------------------------------------------------------------------------
// The server
// ---------------------------------------------------------------------------------------------------------
let sb: Sandbox;
let srv: RunningServer;
let port: number;
const DUMP = () => path.join(sb.dir, 'dumped-job.json');

beforeAll(async () => {
  sb = makeSandbox('sec');
  srv = await startServer(sb, { FAKE_DUMP_JOB: path.join(sb.dir, 'dumped-job.json') });
  port = srv.port;
});
afterAll(async () => {
  await srv?.stop();
  fs.rmSync(sb.dir, { recursive: true, force: true });
});

const JSON_H = { 'Content-Type': 'application/json' };
const openReq = (p: unknown) => raw(port, 'POST', '/api/open', { headers: JSON_H, body: JSON.stringify({ path: p }) });
const saveReq = (name: string, body: Buffer | string = 'PKdata', headers: Record<string, string> = {}) =>
  raw(port, 'POST', '/api/save', { headers: { 'x-file-name': encodeURIComponent(name), ...headers }, body });
const listDir = (d: string) => (fs.existsSync(d) ? fs.readdirSync(d).sort() : []);

describe('guards: Host, Origin, Sec-Fetch-Site (also for reads)', () => {
  it('refuses every Host that is not this machine, including the usual rebinding / IPv6 / encoding tricks', async () => {
    const hosts = [
      'evil.example',
      `evil.example:${port}`,
      'localhost.evil.example',
      `localhost.evil.example:${port}`,
      `127.0.0.1.evil.example:${port}`,
      `localhost:${port}@evil.example`,
      `evil.example#localhost:${port}`,
      `localhost,evil.example:${port}`,
      `localhost:${port}, evil.example`,
      'localhost.',
      '127.1',
      '0x7f.1',
      '2130706433',
      '0177.0.0.1',
      `[::ffff:127.0.0.1]:${port}`,
      `[::1].evil.example`,
      `[0:0:0:0:0:0:0:1]:${port}`,
      '[::1',
      'localhost:5190:5190',
      'LOCALHOST.EVIL.EXAMPLE',
      'localhost%2eevil.example',
      `local\u0068ost.evil.example`,
      '0.0.0.0',
      `${os.hostname()}:${port}`,
    ];
    for (const host of hosts) {
      for (const p of ['/api/ping', '/api/status', '/']) {
        const r = await raw(port, 'GET', p, { headers: { Host: host } });
        expect(r.status, `${host} ${p}`).toBe(403);
      }
    }
    for (const host of [`localhost:${port}`, `127.0.0.1:${port}`, `[::1]:${port}`, 'LOCALHOST', ` localhost:${port} `, `LocalHost:${port}`]) {
      expect((await raw(port, 'GET', '/api/ping', { headers: { Host: host } })).status, host).toBe(200);
    }
  });

  it('API reads are refused when they come from another site; pages can still be navigated to', async () => {
    const crossSite: Array<Record<string, string>> = [
      { 'Sec-Fetch-Site': 'cross-site' },
      { 'Sec-Fetch-Site': 'same-site' },
      { Origin: 'http://evil.example' },
      { Origin: `http://localhost:${port + 1}` },
      { Origin: 'null' },
      { Origin: `https://localhost:${port}` },
      { Origin: `http://evil.example:${port}`, 'Sec-Fetch-Site': 'same-origin' },
    ];
    for (const h of crossSite) {
      for (const p of ['/api/ping', '/api/status', '/api/events', '/api/jobs/00000000-0000-4000-8000-000000000000', '/api/jobs/00000000-0000-4000-8000-000000000000/result']) {
        const r = await raw(port, 'GET', p, { headers: h });
        expect(r.status, `${JSON.stringify(h)} ${p}`).toBe(403);
        expect(r.json().error).toContain('拒否');
      }
    }
    const fromOwnPage: Array<Record<string, string>> = [{}, { 'Sec-Fetch-Site': 'same-origin' }, { 'Sec-Fetch-Site': 'none' }, { Origin: `http://localhost:${port}`, 'Sec-Fetch-Site': 'same-origin' }];
    for (const h of fromOwnPage) {
      expect((await raw(port, 'GET', '/api/status', { headers: h })).status, JSON.stringify(h)).toBe(200);
    }
    // a page of the tool itself may be navigated to from anywhere (a link in an e-mail)
    const page = await raw(port, 'GET', '/', { headers: { 'Sec-Fetch-Site': 'cross-site', 'Sec-Fetch-Mode': 'navigate' } });
    expect(page.status).toBe(200);
  });

  it('cross-site writes are refused for every verb and endpoint, before any body is read', async () => {
    for (const [method, p] of [
      ['POST', '/api/save'],
      ['PUT', '/api/save'],
      ['POST', '/api/open'],
      ['POST', '/api/open-folder'],
      ['POST', '/api/jobs'],
      ['PUT', '/api/jobs/00000000-0000-4000-8000-000000000000/input'],
      ['DELETE', '/api/jobs/00000000-0000-4000-8000-000000000000'],
      ['POST', '/api/bye'],
    ]) {
      const r = await raw(port, method, p, { headers: { Origin: 'http://evil.example', 'x-file-name': 'a.xlsx', 'Content-Type': 'text/plain' }, body: '{"version":1}' });
      expect(r.status, `${method} ${p}`).toBe(403);
    }
    expect(listDir(sb.output).filter((f) => f.includes('a.xlsx'))).toEqual([]);
  });

  it('security headers: no framing, nosniff, same-origin resource policy, no referrer', async () => {
    const page = await raw(port, 'GET', '/');
    expect(page.headers['x-frame-options']).toBe('DENY');
    expect(page.headers['content-security-policy']).toContain("frame-ancestors 'none'");
    expect(page.headers['x-content-type-options']).toBe('nosniff');
    expect(page.headers['referrer-policy']).toBe('no-referrer');
    expect(page.headers['cross-origin-resource-policy']).toBe('same-origin');
    const api = await raw(port, 'GET', '/api/status');
    expect(api.headers['x-content-type-options']).toBe('nosniff');
    expect(api.headers['cross-origin-resource-policy']).toBe('same-origin');
    expect(api.headers['access-control-allow-origin']).toBeUndefined();
    const res404 = await raw(port, 'GET', '/missing.js');
    expect(res404.headers['x-content-type-options']).toBe('nosniff');
    const js = await raw(port, 'GET', '/extra.js');
    expect(js.headers['cross-origin-resource-policy']).toBe('same-origin');
  });

  it('the SSE stream is bounded (64 windows) and refused for other sites', async () => {
    const opened: Array<{ close: () => void }> = [];
    try {
      for (let i = 0; i < 64; i++) opened.push(await openEvents(port));
      const r = await raw(port, 'GET', '/api/events');
      expect(r.status).toBe(429);
      opened.pop()!.close();
      await sleep(300);
      const again = await openEvents(port);
      expect(again.firstChunk).toContain('retry');
      again.close();
    } finally {
      for (const o of opened) o.close();
    }
  });
});

describe('the server listens on the loopback interface only', () => {
  it('a HOST environment variable that names the PC / 0.0.0.0 / a LAN address is ignored (it would publish the API to the network)', async () => {
    const lan = Object.values(os.networkInterfaces())
      .flat()
      .find((a) => a && a.family === 'IPv4' && !a.internal)?.address;
    const hosts = ['0.0.0.0', os.hostname(), lan ?? '192.0.2.1'];
    for (const HOST of hosts) {
      const sb2 = makeSandbox('host');
      const s2 = await startServer(sb2, { HOST });
      try {
        expect((await raw(s2.port, 'GET', '/api/ping')).status, HOST).toBe(200); // reachable on 127.0.0.1
        expect(s2.out(), HOST).toContain('無視しました');
        if (lan) {
          const code = await new Promise<string>((resolve) => {
            const c = net.connect({ host: lan, port: s2.port });
            c.on('connect', () => {
              c.destroy();
              resolve('CONNECTED');
            });
            c.on('error', (e: NodeJS.ErrnoException) => resolve(String(e.code)));
            setTimeout(() => {
              c.destroy();
              resolve('TIMEOUT');
            }, 4000);
          });
          expect(code, `${HOST}: a connection to ${lan}:${s2.port} must not succeed`).not.toBe('CONNECTED');
        }
      } finally {
        await s2.stop();
        fs.rmSync(sb2.dir, { recursive: true, force: true });
      }
    }
  }, 60_000);
});

/** Sends raw bytes on a socket and returns everything the server answers until it closes the connection. */
function rawSocket(p: number, payload: string | Buffer, wait = 1500): Promise<string> {
  return new Promise((resolve) => {
    const s = net.connect(p, '127.0.0.1');
    let buf = '';
    s.on('data', (d) => (buf += d.toString('latin1')));
    s.on('close', () => resolve(buf));
    s.on('error', () => resolve(buf));
    s.write(payload);
    setTimeout(() => s.destroy(), wait);
  });
}

describe('request smuggling and malformed requests', () => {
  it('Content-Length + Transfer-Encoding in one request is refused and the smuggled second request is never executed', async () => {
    const h = `Host: localhost:${port}\r\n`;
    const clte = `POST /api/save HTTP/1.1\r\n${h}x-file-name: smuggle.xlsx\r\nContent-Length: 4\r\nTransfer-Encoding: chunked\r\n\r\n0\r\n\r\nGET /api/status HTTP/1.1\r\n${h}\r\n`;
    const a = await rawSocket(port, clte);
    expect(a).toMatch(/^HTTP\/1\.1 400/);
    expect((a.match(/HTTP\/1\.1 \d{3}/g) ?? []).length, 'exactly one answer').toBe(1);
    expect(listDir(sb.output).filter((f) => f.startsWith('smuggle'))).toEqual([]);
    const tecl = `POST /api/save HTTP/1.1\r\n${h}x-file-name: smuggle2.xlsx\r\nTransfer-Encoding: xchunked\r\nContent-Length: 4\r\n\r\nPK\r\n\r\n`;
    const b = await rawSocket(port, tecl);
    expect(b).toMatch(/^HTTP\/1\.1 400/);
    expect(listDir(sb.output).filter((f) => f.startsWith('smuggle'))).toEqual([]);
    const dup = await rawSocket(port, `POST /api/save HTTP/1.1\r\n${h}x-file-name: smuggle3.xlsx\r\nContent-Length: 2\r\nContent-Length: 40\r\n\r\nPK`);
    expect(dup).toMatch(/^HTTP\/1\.1 400/);
    expect(listDir(sb.output).filter((f) => f.startsWith('smuggle'))).toEqual([]);
  });

  it('bare LF line endings, absurd headers and an absolute-form target do not get around the guards', async () => {
    const lf = await rawSocket(port, `GET /api/ping HTTP/1.1\nHost: evil.example\n\n`);
    expect(lf).not.toMatch(/"app":"excel-merger"/);
    const big = await rawSocket(port, `GET /api/ping HTTP/1.1\r\nHost: localhost:${port}\r\nX-Big: ${'a'.repeat(200_000)}\r\n\r\n`);
    expect(big === '' || /^HTTP\/1\.1 431/.test(big), 'a 431 answer or a closed connection: never processed').toBe(true);
    expect((await raw(port, 'GET', '/api/ping')).status).toBe(200); // the server is still fine
    const abs = await rawSocket(port, `GET http://evil.example/api/ping HTTP/1.1\r\nHost: evil.example\r\n\r\n`);
    expect(abs).toMatch(/^HTTP\/1\.1 403/);
    const absOk = await rawSocket(port, `GET http://evil.example/api/ping HTTP/1.1\r\nHost: localhost:${port}\r\n\r\n`);
    expect(absOk).toMatch(/^HTTP\/1\.1 200/); // routed by path; nothing is fetched from evil.example
    const longUrl = await rawSocket(port, `GET /${'a'.repeat(100_000)} HTTP/1.1\r\nHost: localhost:${port}\r\n\r\n`);
    expect(longUrl === '' || /^HTTP\/1\.1 4\d\d/.test(longUrl)).toBe(true);
  });
});

describe('POST /api/save: hostile names end up as plain files inside the output folder, nothing else', () => {
  it('device names, ADS, bidi characters, trailing dots', async () => {
    const before = new Set(listDir(sb.output));
    const cases: Array<[string, number, string?]> = [
      ['CON.tar.xlsx', 200, '_CON.tar.xlsx'],
      ['NUL .xlsx', 200, '_NUL .xlsx'],
      ['lpt1.xlsx', 200, '_lpt1.xlsx'],
      ['a.xlsx:evil.exe', 400],
      ['a.xlsx::$DATA', 400],
      ['\u202etxt.xlsx', 200, 'txt.xlsx'],
      ['trail.xlsx.   ', 200, 'trail.xlsx'],
      ['x\u0000y.xlsx', 200, 'x_y.xlsx'],
      ['..', 400],
      ['a/../../b.xlsx', 400],
    ];
    for (const [name, status, want] of cases) {
      const r = await saveReq(name, 'PK');
      expect(r.status, name).toBe(status);
      if (want) {
        expect(r.json().name, name).toBe(want);
        expect(path.dirname(r.json().path)).toBe(path.resolve(sb.output));
        expect(fs.readFileSync(r.json().path, 'utf8')).toBe('PK');
      }
    }
    for (const f of listDir(sb.output)) {
      if (before.has(f)) continue;
      const st = fs.statSync(path.join(sb.output, f));
      expect(st.isFile(), f).toBe(true);
      expect(f).toMatch(/\.(xlsx|xlsm)$/);
    }
    // nothing was created next to / above the output folder
    expect(listDir(sb.dir).filter((f) => /evil|txt\.xlsx|^b\.xlsx$/.test(f))).toEqual([]);
  });

  it('a body is streamed to disk byte-exact; an aborted upload leaves no file', async () => {
    const body = Buffer.alloc(5 * 1024 * 1024 + 123);
    for (let i = 0; i < body.length; i += 4099) body[i] = i & 0xff;
    const r = await saveReq('stream.xlsx', body);
    expect(r.status).toBe(200);
    expect(Buffer.compare(fs.readFileSync(r.json().path), body)).toBe(0);

    const before = listDir(sb.output);
    await new Promise<void>((resolve) => {
      const req = http.request({ host: '127.0.0.1', port, method: 'POST', path: '/api/save', headers: { Host: `localhost:${port}`, 'x-file-name': 'aborted.xlsx', 'Content-Length': 10_000_000 } });
      req.on('error', () => resolve());
      req.write(Buffer.alloc(200_000));
      setTimeout(() => {
        req.destroy();
        resolve();
      }, 300);
    });
    await sleep(500);
    expect(listDir(sb.output)).toEqual(before);
    expect(listDir(sb.output).some((f) => f.startsWith('aborted'))).toBe(false);
  });
});

describe('POST /api/open and /api/open-folder: only real .xlsx / .xlsm files inside the output folder', () => {
  let outside: string;
  let secret: string;
  beforeAll(() => {
    fs.mkdirSync(sb.output, { recursive: true });
    outside = path.join(sb.dir, 'outside');
    fs.mkdirSync(outside, { recursive: true });
    secret = path.join(outside, 'secret.xlsx');
    fs.writeFileSync(secret, 'PK-secret');
    fs.writeFileSync(path.join(sb.output, 'inside.xlsx'), 'PK-inside');
  });

  it('a junction inside the output folder that leads outside is refused (real-path check)', async (ctx) => {
    if (!WIN) return ctx.skip();
    const link = path.join(sb.output, 'linkdir');
    try {
      fs.symlinkSync(outside, link, 'junction');
    } catch (e) {
      return ctx.skip(`cannot create a junction here: ${e}`);
    }
    try {
      const viaJunction = await openReq(path.join(link, 'secret.xlsx'));
      expect(viaJunction.status).toBe(403);
      expect(viaJunction.json().error).toMatch(/出力フォルダ/);
      // open-folder with such a selection: only the output folder is opened, the selection is ignored
      const folder = await raw(port, 'POST', '/api/open-folder', { headers: JSON_H, body: JSON.stringify({ path: path.join(link, 'secret.xlsx') }) });
      expect(folder.status).toBe(200);
      expect(path.resolve(folder.json().path)).toBe(path.resolve(sb.output));
    } finally {
      fs.rmSync(link, { recursive: true, force: true });
    }
    expect(fs.existsSync(secret)).toBe(true); // removing the junction did not touch its target
  });

  it('a symbolic link to a file outside is refused (when the test account may create one)', async (ctx) => {
    const link = path.join(sb.output, 'lnk.xlsx');
    try {
      fs.symlinkSync(secret, link, 'file');
    } catch (e) {
      return ctx.skip(`cannot create a file symlink here: ${e}`);
    }
    try {
      expect((await openReq(link)).status).toBe(403);
    } finally {
      fs.rmSync(link, { force: true });
    }
  });

  it('text tricks: UNC, device paths, alternate data streams, traversal, other extensions, trailing dots', async () => {
    const inside = path.join(sb.output, 'inside.xlsx');
    expect((await openReq(inside)).status).toBe(200);
    expect((await openReq(inside.toUpperCase())).status).toBe(WIN ? 200 : 404); // case-insensitive file system
    const refused: Array<[unknown, number | number[]]> = [
      [`\\\\localhost\\c$\\Windows\\win.ini`, 403],
      [`\\\\?\\${inside}`, 403],
      [`\\\\.\\${inside}`, 403],
      [`${inside}:stream`, 403],
      [`${inside}::$DATA`, 403],
      [path.join(sb.output, 'x.txt:y.xlsx'), 403],
      [`${inside}.`, 403],
      [`${inside} `, 403],
      [path.join(sb.output, '..', 'outside', 'secret.xlsx'), 403],
      [secret, 403],
      [path.join(sb.output, 'sub', '..', '..', 'outside', 'secret.xlsx'), 403],
      [sb.output, 403],
      [path.join(sb.output, 'inside.xlsx.exe'), 403],
      [path.join(sb.output, 'inside.bat'), 403],
      [path.join(sb.output, 'NUL.xlsx'), [403, 404]], // a device name: never opened, never hangs
      [path.join(sb.output, 'CON'), 403],
      [path.join(sb.output, 'missing.xlsx'), 404],
      ['C:\\Windows\\System32\\cmd.exe', 403],
      ['inside.xlsx', 403], // relative: resolved against the server's folder, which is not the output folder
      [`${inside}\u0000.exe`, 400],
      [`${inside}\n`, 400],
      [42, 400],
      [null, 400],
      [['x'], 400],
      [{ path: inside }, 400],
      ['', 400],
    ];
    for (const [p, status] of refused) {
      const r = await openReq(p);
      if (Array.isArray(status)) expect(status, JSON.stringify(p)).toContain(r.status);
      else expect(r.status, JSON.stringify(p)).toBe(status);
      expect(typeof r.json().error).toBe('string');
    }
    const dir = path.join(sb.output, 'folder.xlsx');
    fs.mkdirSync(dir, { recursive: true });
    expect((await openReq(dir)).status).toBe(404); // a directory that merely has the extension
    fs.rmSync(dir, { recursive: true, force: true });
  });

  it.skipIf(!WIN)('an 8.3 short alias of the output folder cannot be used to reach anything else', async () => {
    const fso = spawnSync(POWERSHELL, ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', `(New-Object -ComObject Scripting.FileSystemObject).GetFolder($env:EM_SB_DIR).ShortPath`], {
      env: { ...process.env, EM_SB_DIR: sb.dir },
      encoding: 'utf8',
      windowsHide: true,
      timeout: 60_000,
    });
    const shortDir = fso.stdout.trim();
    if (!shortDir || shortDir.toLowerCase() === sb.dir.toLowerCase()) return; // 8.3 names are disabled on this volume
    for (const p of [path.join(shortDir, 'output', 'inside.xlsx'), path.join(shortDir, 'outside', 'secret.xlsx')]) {
      const r = await openReq(p);
      expect([200, 403], p).toContain(r.status);
      if (p.includes('outside')) expect(r.status).toBe(403);
    }
  }, 90_000);

  it('open-folder ignores a selection outside the output folder and always answers with the output folder', async () => {
    for (const body of [JSON.stringify({ path: secret }), JSON.stringify({ path: '\\\\evil\\share\\x.xlsx' }), '{ not json', JSON.stringify({ path: 5 }), '']) {
      const r = await raw(port, 'POST', '/api/open-folder', { headers: JSON_H, body });
      expect(r.status, body).toBe(200);
      expect(path.resolve(r.json().path)).toBe(path.resolve(sb.output));
      expect(r.json().opened).toBe(false); // NO_OPEN in tests: nothing is launched
    }
  });
});

describe('jobs: hostile content, flooding, id guessing', () => {
  const createJob = (job: unknown) => raw(port, 'POST', '/api/jobs', { headers: JSON_H, body: typeof job === 'string' ? job : JSON.stringify(job) });
  const created: string[] = [];
  afterAll(async () => {
    for (const id of created) await raw(port, 'DELETE', `/api/jobs/${id}`).catch(() => {});
  });

  it('what reaches the runner is the whitelisted copy, and server-chosen paths inside the temp folder', async () => {
    const sheet = "S'h\"e$(calc)`e\\t;&|<>%日本";
    const hostileOps = [
      { op: 'setCells', sheet, at: 'A1', rows: [["=cmd|' /C calc'!A0", { f: '=HYPERLINK("http://evil","x")' }]], evil: 'x', input: 'C:\\Windows\\win.ini' },
      { op: 'moveCols', sheet, start: 1, count: 1, to: 3, workDir: 'C:\\', output: 'C:\\x.xlsx' },
      { op: 'unmerge', sheet, refs: ['A1:B2'], centerAcross: true, extra: 1 },
    ];
    const { id } = (await createJob({ version: 1, format: 'xlsx', ops: hostileOps, input: 'C:\\Windows\\win.ini', output: path.join(sb.dir, 'evil.xlsx'), workDir: sb.dir, options: { recalc: true, clipboard: 'restoreEachOp', extra: 1 } })).json();
    created.push(id);
    expect((await raw(port, 'PUT', `/api/jobs/${id}/input`, { body: await makePackage() })).status).toBe(200);
    const end = Date.now() + 15_000;
    while (!fs.existsSync(DUMP()) && Date.now() < end) await sleep(100);
    const job = JSON.parse(fs.readFileSync(DUMP(), 'utf8'));
    expect(job.ops).toEqual([
      { op: 'setCells', sheet, at: 'A1', rows: [["=cmd|' /C calc'!A0", { f: '=HYPERLINK("http://evil","x")' }]] },
      { op: 'moveCols', sheet, start: 1, count: 1, to: 3, method: 'clipFree' },
      { op: 'unmerge', sheet, refs: ['A1:B2'], centerAcross: true },
    ]);
    expect(job.options).toEqual({ recalc: true, clipboard: 'restoreEachOp' });
    for (const k of ['input', 'output', 'workDir']) expect(isInside(sb.temp, job[k]), `${k}: ${job[k]}`).toBe(true);
    expect(path.basename(job.input)).toMatch(/^work\.xlsx$/);
    expect(path.basename(job.output)).toBe('output.xlsx');
    expect(fs.existsSync(path.join(sb.dir, 'evil.xlsx'))).toBe(false);
  });

  it('refuses hostile ops with a 400 that names the op', async () => {
    const cases: Array<[unknown, RegExp]> = [
      [{ op: 'setCells', sheet: 'S', at: "'\\\\evil\\share\\[x.xlsx]S'!A1", rows: [[1]] }, /setCells の at/],
      [{ op: 'insertCols', sheet: 'a\nb', at: 1, count: 1 }, /シート名/],
      [{ op: 'deleteCols', sheet: 'S', at: -1, count: 1 }, /at \/ count/],
      [{ op: 'moveCols', sheet: 'S', start: 1, count: 1, to: 2, method: 'x' }, /method/],
      [{ op: 'setCells', sheet: 'S', at: 'A1', rows: [[{ f: { x: 1 } }]] }, /書き込めない/],
      [{ op: 'powershell', sheet: 'S' }, /未知の操作/],
    ];
    for (const [op, re] of cases) {
      const r = await createJob({ version: 1, format: 'xlsx', ops: [op] });
      expect(r.status, JSON.stringify(op)).toBe(400);
      expect(r.json().error).toMatch(re);
    }
  });

  it('job ids are random v4 UUIDs; guessing, traversal and other ids never reach another job', async () => {
    const ids: string[] = [];
    for (let i = 0; i < 12; i++) {
      const { id } = (await createJob({ version: 1, format: 'xlsx', ops: [] })).json();
      ids.push(id);
      created.push(id);
      expect(id).toMatch(/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/);
    }
    expect(new Set(ids).size).toBe(ids.length);
    const prefixes = new Set(ids.map((i) => i.slice(0, 8)));
    expect(prefixes.size, 'the folder name uses the first 8 hex digits: they must not repeat within one server').toBe(ids.length);
    for (const bad of ['00000000-0000-4000-8000-000000000000', ids[0].slice(0, 35), ids[0].toUpperCase(), '..%2f..%2fstatus', `${ids[0]}%00`, `${ids[0]}x`, '-'.repeat(36)]) {
      const r = await raw(port, 'GET', `/api/jobs/${bad}`);
      expect([404, 403], bad).toContain(r.status);
      expect(r.text(), bad).not.toContain('"state"');
    }
    // the result of a job that is not done is refused; other jobs' results are not reachable through this id
    expect((await raw(port, 'GET', `/api/jobs/${ids[0]}/result`)).status).toBe(409);
    expect((await raw(port, 'GET', `/api/jobs/${ids[0]}/input`)).status).toBe(405);
  });

  it('flooding: at most 20 unfinished jobs are kept; deleting one makes room again', async () => {
    const extra: string[] = [];
    let refused = 0;
    for (let i = 0; i < 30; i++) {
      const r = await createJob({ version: 1, format: 'xlsx', ops: [] });
      if (r.status === 200) extra.push(r.json().id);
      else {
        refused++;
        expect(r.status).toBe(429);
        expect(r.json().error).toMatch(/多すぎ/);
      }
    }
    created.push(...extra);
    expect(refused).toBeGreaterThan(0);
    expect((await raw(port, 'DELETE', `/api/jobs/${extra[0]}`)).status).toBe(200);
    const again = await createJob({ version: 1, format: 'xlsx', ops: [] });
    expect(again.status).toBe(200);
    created.push(again.json().id);
    for (const id of created.splice(0)) await raw(port, 'DELETE', `/api/jobs/${id}`);
  });
});

describe('oversize bodies and a full disk', () => {
  it('bodies above the limit are refused while streaming (no file left), declared or not (chunked)', async () => {
    const sb2 = makeSandbox('big');
    const s2 = await startServer(sb2, { EM_MAX_BODY_MB: '2' });
    try {
      const big = Buffer.alloc(3 * 1024 * 1024, 1);
      // declared size (Content-Length) and streamed size (chunked: raw() sends no Content-Length)
      const declared = await raw(s2.port, 'POST', '/api/save', { headers: { 'x-file-name': 'big.xlsx', 'Content-Length': big.length }, body: big });
      expect(declared.status).toBe(413);
      const streamed = await raw(s2.port, 'POST', '/api/save', { headers: { 'x-file-name': 'big2.xlsx' }, body: big });
      expect(streamed.status).toBe(413);
      expect(listDir(sb2.output).filter((f) => f.startsWith('big'))).toEqual([]);
      // JSON job above the limit
      const job = await raw(s2.port, 'POST', '/api/jobs', { headers: JSON_H, body: JSON.stringify({ version: 1, format: 'xlsx', ops: [{ op: 'setCells', sheet: 'S', at: 'A1', rows: [[big.toString('latin1')]] }] }) });
      expect(job.status).toBe(413);
      // chunked (no Content-Length): refused in the middle of the stream
      const status = await new Promise<number>((resolve, reject) => {
        const req = http.request({ host: '127.0.0.1', port: s2.port, method: 'POST', path: '/api/save', headers: { Host: `localhost:${s2.port}`, 'x-file-name': 'chunked.xlsx', 'Transfer-Encoding': 'chunked' } }, (res) => {
          resolve(res.statusCode || 0);
          res.resume();
        });
        req.on('error', (e) => (/ECONNRESET|hang up|EPIPE/.test(String(e)) ? undefined : reject(e)));
        for (let i = 0; i < 6; i++) req.write(Buffer.alloc(512 * 1024, 2));
        req.end();
      });
      expect(status).toBe(413);
      await sleep(300);
      expect(listDir(sb2.output).filter((f) => f.startsWith('chunked'))).toEqual([]);
      // a job input above the limit: 413 and no leftover folder
      const { id } = (await raw(s2.port, 'POST', '/api/jobs', { headers: JSON_H, body: JSON.stringify({ version: 1, format: 'xlsx', ops: [] }) })).json();
      const put = await raw(s2.port, 'PUT', `/api/jobs/${id}/input`, { body: big });
      expect(put.status).toBe(413);
      await raw(s2.port, 'DELETE', `/api/jobs/${id}`);
      await sleep(300);
      expect(listDir(sb2.temp).filter((d) => d.startsWith('job-'))).toEqual([]);
      // the server is still fine
      expect((await raw(s2.port, 'GET', '/api/ping')).status).toBe(200);
    } finally {
      await s2.stop();
      fs.rmSync(sb2.dir, { recursive: true, force: true });
    }
  }, 60_000);

  it('a decompression bomb in [Content_Types].xml is cut off (the package check reads at most 4 MB of it) and the server stays responsive', async () => {
    const JSZip = (await import('jszip')).default;
    const zip = new JSZip();
    zip.file('[Content_Types].xml', Buffer.alloc(96 * 1024 * 1024, 0x20), { compression: 'DEFLATE' }); // ~100 KB on disk, 96 MB inflated
    zip.file('xl/workbook.xml', '<workbook/>');
    const bomb = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
    expect(bomb.length).toBeLessThan(2 * 1024 * 1024);
    const { id } = (await raw(port, 'POST', '/api/jobs', { headers: JSON_H, body: JSON.stringify({ version: 1, format: 'xlsx', ops: [] }) })).json();
    const t0 = Date.now();
    const put = await raw(port, 'PUT', `/api/jobs/${id}/input`, { body: bomb });
    expect([200, 400]).toContain(put.status);
    expect(Date.now() - t0).toBeLessThan(10_000);
    expect((await raw(port, 'GET', '/api/ping')).status).toBe(200);
    await raw(port, 'DELETE', `/api/jobs/${id}`);
  }, 60_000);

  it('status.diskLow / diskFreeMB / diskMessage: a clear Japanese warning when the temp drive is nearly full', async () => {
    const normal = (await raw(port, 'GET', '/api/status')).json();
    expect(typeof normal.diskLow).toBe('boolean');
    expect(normal.outputDiskLow).toBeDefined();
    if (normal.diskFreeMB !== null) expect(normal.diskFreeMB).toBeGreaterThan(0);
    expect(normal.jobs).toBeDefined();

    const sb2 = makeSandbox('low');
    const s2 = await startServer(sb2, { EM_DISK_LOW_MB: '999999999' }); // "less than a few petabytes free" = always low
    try {
      const st = (await raw(s2.port, 'GET', '/api/status')).json();
      if (st.diskFreeMB === null) return; // Node without statfs: nothing to check
      expect(st.diskLow).toBe(true);
      expect(st.outputDiskLow).toBe(true);
      expect(st.diskMessage).toMatch(/空き容量が少なくなっています/);
      expect(st.diskMessage).toMatch(/内蔵エンジン/);
      const end = Date.now() + 5000;
      while (!s2.out().includes('警告: ') && Date.now() < end) await sleep(100);
      expect(s2.out(), 'the server log carries the same warning').toContain('空き容量が少なくなっています');
    } finally {
      await s2.stop();
      fs.rmSync(sb2.dir, { recursive: true, force: true });
    }
  }, 60_000);

  it('an Excel job is refused up front when the temp drive cannot hold it: DISK_LOW + fallbackToJs (the UI uses the built-in engine), no files left', async () => {
    const sb2 = makeSandbox('nospace');
    const s2 = await startServer(sb2, { EM_MIN_FREE_MB: '999999999' });
    try {
      if ((await raw(s2.port, 'GET', '/api/status')).json().diskFreeMB === null) return;
      const { id } = (await raw(s2.port, 'POST', '/api/jobs', { headers: JSON_H, body: JSON.stringify({ version: 1, format: 'xlsx', ops: [] }) })).json();
      const put = await raw(s2.port, 'PUT', `/api/jobs/${id}/input`, { body: await makePackage() });
      expect(put.status).toBe(507);
      expect(put.json().errorCode).toBe('DISK_LOW');
      expect(put.json().fallbackToJs).toBe(true);
      expect(put.json().error).toMatch(/空き容量が足りません/);
      expect(put.json().error).toMatch(/内蔵エンジン/);
      await raw(s2.port, 'DELETE', `/api/jobs/${id}`);
      await sleep(300);
      expect(listDir(sb2.temp)).toEqual([]);
    } finally {
      await s2.stop();
      fs.rmSync(sb2.dir, { recursive: true, force: true });
    }
  }, 60_000);
});
