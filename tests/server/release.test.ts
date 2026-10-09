/**
 * scripts/make-release.mjs (npm run release): builds release\ExcelMerger\ + ExcelMerger_<yyyyMMdd>.zip from the project files.
 * Runs the script as a child process with a SMALL FAKE exe (--min-exe-size lowers the 50 MB limit) and --out into
 * <os.tmpdir()>\release-test\... (a folder with Japanese characters, spaces and parentheses; never the project's release\).
 * Checks the folder, start.bat (byte-identical, ASCII, CRLF), the BOM / CRLF of the txt, the UTF-8 names of the zip
 * (flag + valid UTF-8 + JSZip), that no developer folder is included, the refusals and, on Windows, that Windows PowerShell 5.1
 * Expand-Archive extracts every Japanese file name correctly.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import { ROOT } from './helpers';

const SCRIPT = path.join(ROOT, 'scripts', 'make-release.mjs');
const BASE = path.join(os.tmpdir(), 'release-test');
const WIN = process.platform === 'win32';
const TXT = 'はじめにお読みください.txt';
const DATE = '20260102';
/** Folders / files that must never be in the hand-out. */
const FORBIDDEN = ['node_modules', 'src', 'tests', 'docs', 'dist', '.tmp', 'test-results', 'output', 'e2e', 'tools', 'build', 'server', 'scripts', 'release', 'truth.json', 'package.json', 'package-lock.json', 'tsconfig.json'];

let work = '';
let fakeExe = '';
let out = '';
let firstRun = { status: -1, stdout: '', stderr: '' };

function run(args: string[]) {
  const r = spawnSync(process.execPath, [SCRIPT, ...args], { encoding: 'utf8', timeout: 120_000, windowsHide: true, env: { ...process.env, NODE_OPTIONS: '' } });
  return { status: r.status ?? -1, stdout: r.stdout || '', stderr: r.stderr || '' };
}

/** All files below dir (posix-style relative paths), sorted. */
function walk(dir: string, prefix = ''): string[] {
  const res: string[] = [];
  for (const e of fs.readdirSync(path.join(dir, prefix), { withFileTypes: true })) {
    const rel = prefix ? `${prefix}/${e.name}` : e.name;
    if (e.isDirectory()) res.push(...walk(dir, rel));
    else res.push(rel);
  }
  return res.sort();
}

/** Central directory of a zip: name bytes, UTF-8 flag. */
function central(buf: Buffer) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  expect(eocd, 'end of central directory').toBeGreaterThanOrEqual(0);
  const count = buf.readUInt16LE(eocd + 10);
  let off = buf.readUInt32LE(eocd + 16);
  const list: { name: string; raw: Buffer; utf8: boolean }[] = [];
  for (let n = 0; n < count; n++) {
    expect(buf.readUInt32LE(off)).toBe(0x02014b50);
    const flags = buf.readUInt16LE(off + 8);
    const nl = buf.readUInt16LE(off + 28);
    const el = buf.readUInt16LE(off + 30);
    const cl = buf.readUInt16LE(off + 32);
    const raw = Buffer.from(buf.subarray(off + 46, off + 46 + nl));
    list.push({ name: raw.toString('utf8'), raw, utf8: (flags & 0x800) !== 0 });
    off += 46 + nl + el + cl;
  }
  return list;
}

beforeAll(() => {
  expect(path.resolve(BASE).startsWith(path.resolve(ROOT, 'release')), 'the test must never write into the project release folder').toBe(false);
  fs.mkdirSync(BASE, { recursive: true });
  work = fs.mkdtempSync(path.join(BASE, '配布 テスト (1)-'));
  fakeExe = path.join(work, 'fake', 'ExcelMerger.exe');
  fs.mkdirSync(path.dirname(fakeExe), { recursive: true });
  const exe = Buffer.alloc(200_000, 0x41);
  exe.write('MZ', 0, 'latin1');
  for (let i = 100; i < exe.length; i += 7) exe[i] = (i * 31) & 0xff; // not trivially compressible
  fs.writeFileSync(fakeExe, exe);
  out = path.join(work, 'out ダミー');
  firstRun = run(['--exe', fakeExe, '--out', out, '--date', DATE, '--min-exe-size', '1000']);
}, 120_000);

afterAll(() => {
  if (work) fs.rmSync(work, { recursive: true, force: true });
  try {
    fs.rmdirSync(BASE); // only when nothing else is in it
  } catch {
    /* not empty / gone */
  }
});

describe('make-release: refusals', () => {
  it('refuses a missing exe with a clear message and creates nothing', () => {
    const o = path.join(work, 'out-missing');
    const r = run(['--exe', path.join(work, 'does-not-exist.exe'), '--out', o, '--date', DATE]);
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('was not found');
    expect(r.stderr).toContain('does-not-exist.exe');
    expect(fs.existsSync(path.join(o, 'ExcelMerger'))).toBe(false);
    expect(fs.existsSync(path.join(o, `ExcelMerger_${DATE}.zip`))).toBe(false);
  });

  it('refuses an exe smaller than 50 MB (the default limit) and says why', () => {
    const o = path.join(work, 'out-small');
    const r = run(['--exe', fakeExe, '--out', o, '--date', DATE]);
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/only 0\.2 MB/);
    expect(r.stderr).toContain('50 MB');
    expect(fs.existsSync(path.join(o, 'ExcelMerger'))).toBe(false);
    expect(fs.existsSync(path.join(o, `ExcelMerger_${DATE}.zip`))).toBe(false);
  });

  it('refuses a file that is not a Windows exe', () => {
    const notExe = path.join(work, 'not-an-exe.exe');
    fs.writeFileSync(notExe, Buffer.alloc(5000, 0x42));
    const r = run(['--exe', notExe, '--out', path.join(work, 'out-notexe'), '--min-exe-size', '1000']);
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('MZ');
  });

  it('refuses unknown options, a bad --date and an exe that lives inside the folder it would replace', () => {
    expect(run(['--bogus', 'x']).status).toBe(1);
    const bad = run(['--exe', fakeExe, '--out', path.join(work, 'out-baddate'), '--date', '2026-01-02', '--min-exe-size', '1000']);
    expect(bad.status).toBe(1);
    expect(bad.stderr).toContain('yyyyMMdd');
    // --exe pointing into <out>\ExcelMerger: the script would delete its own input
    const o = path.join(work, 'out-inside');
    fs.mkdirSync(path.join(o, 'ExcelMerger'), { recursive: true });
    const inside = path.join(o, 'ExcelMerger', 'ExcelMerger.exe');
    fs.copyFileSync(fakeExe, inside);
    const r = run(['--exe', inside, '--out', o, '--min-exe-size', '1000']);
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('inside the folder');
    expect(fs.existsSync(inside)).toBe(true);
  });

  it('does not delete a folder that does not look like a hand-out folder', () => {
    const o = path.join(work, 'out-foreign');
    fs.mkdirSync(path.join(o, 'ExcelMerger'), { recursive: true });
    const precious = path.join(o, 'ExcelMerger', 'my-notes.txt');
    fs.writeFileSync(precious, 'keep me');
    const r = run(['--exe', fakeExe, '--out', o, '--min-exe-size', '1000']);
    expect(r.status).toBe(1);
    expect(r.stderr).toContain('does not look like a hand-out folder');
    expect(fs.readFileSync(precious, 'utf8')).toBe('keep me');
  });
});

describe('make-release: the hand-out folder', () => {
  const folder = () => path.join(out, 'ExcelMerger');

  it('ran successfully', () => {
    expect(firstRun.status, firstRun.stderr + firstRun.stdout).toBe(0);
    expect(firstRun.stdout).toContain('zip checked');
  });

  it('contains exactly the exe, start.bat, README.md, the txt and samples\\', () => {
    expect(fs.readdirSync(folder()).sort()).toEqual(['ExcelMerger.exe', 'README.md', 'samples', 'start.bat', TXT].sort());
  });

  it('copies the exe and start.bat byte for byte; start.bat is ASCII with CRLF only and no BOM', () => {
    expect(fs.readFileSync(path.join(folder(), 'ExcelMerger.exe')).equals(fs.readFileSync(fakeExe))).toBe(true);
    const copy = fs.readFileSync(path.join(folder(), 'start.bat'));
    expect(copy.equals(fs.readFileSync(path.join(ROOT, 'start.bat')))).toBe(true);
    expect(copy.some((b) => b > 0x7f)).toBe(false);
    expect(copy.subarray(0, 3).equals(Buffer.from([0xef, 0xbb, 0xbf]))).toBe(false);
    const s = copy.toString('latin1');
    expect(s.startsWith('@echo off\r\n')).toBe(true);
    expect(/(^|[^\r])\n/.test(s)).toBe(false);
    expect(/\r(?!\n)/.test(s)).toBe(false);
  });

  it('writes 「はじめにお読みください.txt」 as UTF-8 with BOM and CRLF, 6-10 short friendly lines with the key facts', () => {
    const b = fs.readFileSync(path.join(folder(), TXT));
    expect([...b.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
    expect(b[3]).not.toBe(0xef); // exactly one BOM
    const text = b.subarray(3).toString('utf8');
    expect(text.includes('�')).toBe(false);
    expect(/(^|[^\r])\n/.test(text)).toBe(false); // CRLF only
    expect(text.endsWith('\r\n')).toBe(true);
    const lines = text.slice(0, -2).split('\r\n');
    const filled = lines.filter((l) => l.trim() !== '');
    expect(lines.length).toBeLessThanOrEqual(10);
    expect(filled.length).toBeGreaterThanOrEqual(6);
    for (const l of lines) expect(l.length, l).toBeLessThanOrEqual(80);
    for (const must of ['start.bat', 'ダブルクリック', 'output', 'プロパティ', 'ブロックの解除', 'WindowsによってPCが保護されました', '詳細情報', '実行', '削除', 'インターネット', '送られ']) {
      expect(text, must).toContain(must);
    }
    expect(text).not.toMatch(/#|\*\*|```/); // plain text, no Markdown
  });

  it('ships README.md without the developer / maker sections, and the samples without truth.json', () => {
    const readme = fs.readFileSync(path.join(folder(), 'README.md'), 'utf8');
    const src = fs.readFileSync(path.join(ROOT, 'README.md'), 'utf8');
    expect(readme).toContain('## 起動のしかた');
    expect(readme).toContain('## zip で受け取ったとき');
    expect(readme).toContain('ブロックの解除');
    expect(readme).toContain('はじめにお読みください.txt');
    expect(readme).not.toMatch(/^## .*(開発|作る人向け)/m);
    expect(readme).not.toContain('npm run');
    expect(readme.length).toBeLessThan(src.length);
    expect(readme.trimEnd().endsWith('---')).toBe(false);

    const samples = fs.readdirSync(path.join(folder(), 'samples')).sort();
    const want = fs
      .readdirSync(path.join(ROOT, 'samples'))
      .filter((n) => /\.(xlsx|xls|csv)$/i.test(n) && !n.startsWith('~$'))
      .sort();
    expect(want.length).toBe(5);
    expect(samples).toEqual([...want, 'README.md'].sort());
    expect(samples).not.toContain('truth.json');
    for (const n of want) expect(fs.readFileSync(path.join(folder(), 'samples', n)).equals(fs.readFileSync(path.join(ROOT, 'samples', n))), n).toBe(true);
    const sr = fs.readFileSync(path.join(folder(), 'samples', 'README.md'), 'utf8');
    expect(sr).toContain('01_学生名簿.xlsx');
    expect(sr).not.toMatch(/truth\.json|npm run|tools[\\/]|--base /);
  });

  it('includes no developer folders or files', () => {
    for (const rel of walk(folder())) {
      for (const part of rel.split('/')) expect(FORBIDDEN, rel).not.toContain(part);
      expect(rel).not.toMatch(/\.(ts|tsx|mjs|map|ps1)$/);
      expect(rel).not.toMatch(/\.old-\d+\.exe$/);
    }
  });

  it('prints the zip size (and it is the real size of the file)', () => {
    const zipPath = path.join(out, `ExcelMerger_${DATE}.zip`);
    const m = /zip size: ([\d.]+) MB \(([\d,]+) bytes\)/.exec(firstRun.stdout);
    expect(m, firstRun.stdout).toBeTruthy();
    expect(Number(m![2].replace(/,/g, ''))).toBe(fs.statSync(zipPath).size);
    const sha = /zip sha256: ([0-9a-f]{64})/.exec(firstRun.stdout);
    expect(sha, firstRun.stdout).toBeTruthy();
    expect(sha![1]).toBe(createHash('sha256').update(fs.readFileSync(zipPath)).digest('hex'));
  });

  it('is repeatable: a second run replaces the folder and drops leftovers', () => {
    fs.writeFileSync(path.join(folder(), 'leftover.txt'), 'old');
    const r = run(['--exe', fakeExe, '--out', out, '--date', DATE, '--min-exe-size', '1000']);
    expect(r.status, r.stderr).toBe(0);
    expect(fs.existsSync(path.join(folder(), 'leftover.txt'))).toBe(false);
    expect(fs.readdirSync(folder()).sort()).toEqual(['ExcelMerger.exe', 'README.md', 'samples', 'start.bat', TXT].sort());
    expect(fs.readdirSync(out).sort()).toEqual(['ExcelMerger', `ExcelMerger_${DATE}.zip`]); // no .partial left
  });

  it('uses today (local time) for the zip name when --date is not given', () => {
    const o = path.join(work, 'out-today');
    const r = run(['--exe', fakeExe, '--out', o, '--min-exe-size', '1000']);
    expect(r.status, r.stderr).toBe(0);
    const d = new Date();
    const p = (n: number) => String(n).padStart(2, '0');
    const today = `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}`;
    const yesterday = new Date(d.getTime() - 86_400_000);
    const prev = `${yesterday.getFullYear()}${p(yesterday.getMonth() + 1)}${p(yesterday.getDate())}`; // the run may straddle midnight
    const zips = fs.readdirSync(o).filter((n) => n.endsWith('.zip'));
    expect(zips).toHaveLength(1);
    expect(zips[0]).toMatch(/^ExcelMerger_\d{8}\.zip$/);
    expect([`ExcelMerger_${today}.zip`, `ExcelMerger_${prev}.zip`]).toContain(zips[0]);
  });
});

describe('make-release: the zip', () => {
  const zipPath = () => path.join(out, `ExcelMerger_${DATE}.zip`);

  it('has ONE top-level folder ExcelMerger/ with exactly the files of the folder (and its folder entries)', async () => {
    const zip = await JSZip.loadAsync(fs.readFileSync(zipPath()), { checkCRC32: true });
    const names = Object.keys(zip.files).sort();
    const files = walk(path.join(out, 'ExcelMerger'));
    const dirs = new Set<string>(['ExcelMerger/']);
    for (const f of files) {
      const parts = f.split('/');
      for (let i = 1; i < parts.length; i++) dirs.add(`ExcelMerger/${parts.slice(0, i).join('/')}/`);
    }
    expect(names).toEqual([...files.map((f) => `ExcelMerger/${f}`), ...dirs].sort());
    for (const n of names) expect(n.startsWith('ExcelMerger/')).toBe(true);
    for (const n of names) for (const part of n.split('/')) expect(FORBIDDEN, n).not.toContain(part);
    expect(names).toContain(`ExcelMerger/${TXT}`);
    expect(names).toContain('ExcelMerger/samples/01_学生名簿.xlsx');
    expect(names).toContain('ExcelMerger/samples/04_適性検査.xls');
    for (const f of files) {
      const bytes = await zip.file(`ExcelMerger/${f}`)!.async('nodebuffer');
      expect(bytes.equals(fs.readFileSync(path.join(out, 'ExcelMerger', ...f.split('/')))), f).toBe(true);
    }
  });

  it('stores every non-ASCII name as valid UTF-8 with the UTF-8 flag (so Windows shows it correctly)', () => {
    const entries = central(fs.readFileSync(zipPath()));
    const japanese = entries.filter((e) => /[^\x00-\x7f]/.test(e.name));
    expect(japanese.length).toBeGreaterThanOrEqual(6); // the txt + 5 samples
    for (const e of entries) {
      expect(Buffer.from(e.name, 'utf8').equals(e.raw), e.name).toBe(true);
      expect(e.name.includes('�'), e.name).toBe(false);
      expect(e.name.includes('\\'), e.name).toBe(false);
      if (/[^\x00-\x7f]/.test(e.name)) expect(e.utf8, `UTF-8 flag of ${e.name}`).toBe(true);
    }
  });

  it.skipIf(!WIN)('extracts with Windows PowerShell 5.1 Expand-Archive into a folder with Japanese characters, spaces and parentheses, every name intact', () => {
    const dest = path.join(work, '展開 先 (1)');
    const r = spawnSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', "$ErrorActionPreference='Stop'; Expand-Archive -LiteralPath $env:RT_ZIP -DestinationPath $env:RT_DEST -Force"], {
      encoding: 'utf8',
      timeout: 90_000,
      windowsHide: true,
      env: { ...process.env, RT_ZIP: zipPath(), RT_DEST: dest },
    });
    expect(r.status, `${r.stdout}\n${r.stderr}`).toBe(0);
    const extracted = walk(path.join(dest, 'ExcelMerger'));
    expect(extracted).toEqual(walk(path.join(out, 'ExcelMerger')));
    expect(extracted).toContain(TXT);
    expect(extracted).toContain('samples/01_学生名簿.xlsx');
    expect(extracted).toContain('samples/05_研修出欠.xlsx');
    expect(fs.readFileSync(path.join(dest, 'ExcelMerger', 'start.bat')).equals(fs.readFileSync(path.join(ROOT, 'start.bat')))).toBe(true);
    expect(fs.readFileSync(path.join(dest, 'ExcelMerger', TXT)).equals(fs.readFileSync(path.join(out, 'ExcelMerger', TXT)))).toBe(true);
  }, 120_000);
});
