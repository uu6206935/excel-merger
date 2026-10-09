#!/usr/bin/env node
/**
 * build-exe.mjs (npm run build:exe) - packages the local server + single-file UI + Excel runner into
 * ExcelMerger.exe (Node SEA, docs/research/R5-packaging.md).
 *
 *   node scripts/build-exe.mjs [--out <file.exe>] [--icon <file.ico>] [--index <index.html>] [--release [dir]]
 *                              [--layout sea|runtime] [--no-code-cache] [--gui-subsystem] [--keep-signature]
 *                              [--classic] [--no-smoke]
 *
 * Inputs (project root = the folder above scripts/):
 *   server/server.mjs               ESM entry (Node built-ins + ./excel/*.mjs), bundled to ONE CommonJS file
 *   server/excel/excel-runner.ps1   Excel本体 engine (rebuilt from its sources when stale), SEA asset
 *   dist/index.html                 vite-plugin-singlefile build (npm run build), SEA asset "index.html"
 *                                   (--index <file> embeds another page, e.g. a placeholder while the UI is not built)
 * Output:
 *   --layout sea (default)  ExcelMerger.exe in the project root (start.bat prefers it)
 *                           --release [dir]: also a hand-out folder (default release\ExcelMerger) with the exe,
 *                           start.bat, README.md and dist\index.html (standalone fallback)
 *   --layout runtime        release\ExcelMerger-runtime\ : runtime\node.exe (the UNMODIFIED, OpenJS-signed node.exe),
 *                           app\server.cjs, server\excel\excel-runner.ps1, dist\index.html, start.bat, README.md.
 *                           Use it where Smart App Control / an antivirus blocks the unsigned SEA exe.
 *
 * Steps (Node 24 "classic" flow; Node >= 25.5 has `node --build-sea`, used automatically):
 *   1 esbuild bundle  2 sea-config  3 blob  4 copy node.exe  5 drop signature (signtool, optional)
 *   6 postject  7 icon + version (resedit, AFTER postject)  8 optional GUI subsystem  9 smoke test (READY + /api/ping)
 * Run it with the x64 node.exe that should be embedded (it copies process.execPath).
 */
import fs from 'node:fs';
import http from 'node:http';
import path from 'node:path';
import { spawn, spawnSync, execFileSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, '..');
const require = createRequire(import.meta.url);
const SENTINEL_FUSE = 'NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2';
const EXE_NAME = 'ExcelMerger.exe';

// ----------------------------------------------------------------- args
const argv = process.argv.slice(2);
const flag = (n) => argv.includes(n);
const opt = (n, d) => {
  const i = argv.indexOf(n);
  return i >= 0 && argv[i + 1] && !argv[i + 1].startsWith('--') ? argv[i + 1] : d;
};
const LAYOUT = opt('--layout', 'sea');
const BUILD = path.join(ROOT, 'build', 'sea');
const OUT = path.resolve(opt('--out', path.join(ROOT, EXE_NAME)));
const USE_CODE_CACHE = !flag('--no-code-cache');
const ICON = opt('--icon', fs.existsSync(path.join(HERE, 'app.ico')) ? path.join(HERE, 'app.ico') : null);
const GUI = flag('--gui-subsystem');
const INDEX = path.resolve(opt('--index', path.join(ROOT, 'dist', 'index.html')));
const RELEASE = flag('--release') ? path.resolve(opt('--release', path.join(ROOT, 'release', 'ExcelMerger'))) : null;
const RUNNER_DIR = path.join(ROOT, 'server', 'excel');
const RUNNER = path.join(RUNNER_DIR, 'excel-runner.ps1');
const VERSION = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version || '1.0.0';

const t0 = Date.now();
const step = (s) => console.log(`\n[build-exe] ${s}  (+${((Date.now() - t0) / 1000).toFixed(1)}s)`);
const die = (msg) => {
  console.error(`\n[build-exe] ERROR: ${msg}`);
  process.exit(1);
};

if (process.platform !== 'win32') console.warn('[build-exe] WARNING: this script targets Windows (.exe).');
const major = Number(process.versions.node.split('.')[0]);
if (major < 22) die(`Node ${process.version}: use Node 22 or 24 LTS (x64) to build the exe`);
const hasBuildSea = spawnSync(process.execPath, ['--help'], { encoding: 'utf8' }).stdout.includes('--build-sea');
console.log(`[build-exe] node ${process.version} (${process.execPath})  root=${ROOT}  layout=${LAYOUT}  version=${VERSION}`);

if (!fs.existsSync(INDEX)) die(`${INDEX} is missing: run "npm run build" first (or pass --index <placeholder.html>)`);
// Excel runner: rebuild from sources when stale, and it must carry the UTF-8 BOM (Windows PowerShell 5.1)
{
  const { buildRunnerFile, runnerIsStale } = await import(pathToFileURL(path.join(RUNNER_DIR, 'merge-runner.mjs')).href);
  if (fs.existsSync(path.join(RUNNER_DIR, 'excel-runner.src.ps1')) && runnerIsStale(RUNNER_DIR)) {
    buildRunnerFile(RUNNER_DIR);
    console.log('[build-exe] rebuilt server/excel/excel-runner.ps1');
  }
  if (!fs.existsSync(RUNNER)) die(`${RUNNER} is missing: run "npm run build:runner"`);
  const b = fs.readFileSync(RUNNER);
  if (!(b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf)) die('excel-runner.ps1 has no UTF-8 BOM (Windows PowerShell 5.1 would misread it)');
}
for (const bat of ['start.bat', 'build.bat']) {
  const p = path.join(ROOT, bat);
  if (fs.existsSync(p) && /[^\x00-\x7f]/.test(fs.readFileSync(p, 'latin1'))) die(`${bat} must be ASCII-only`);
}

// ----------------------------------------------------------------- 1. bundle
step('1/9 esbuild bundle -> build/sea/server.cjs');
fs.mkdirSync(BUILD, { recursive: true });
const esbuild = require('esbuild');
const bundle = await esbuild.build({
  entryPoints: [path.join(ROOT, 'server', 'server.mjs')],
  outfile: path.join(BUILD, 'server.cjs'),
  bundle: true,
  platform: 'node',
  format: 'cjs',
  target: `node${Math.min(major, 24)}`,
  // import.meta is empty in CJS output: feed it from __filename (== process.execPath inside SEA).
  define: { 'import.meta.url': '__import_meta_url', __APP_VERSION__: JSON.stringify(VERSION) },
  banner: { js: 'const __import_meta_url = require("node:url").pathToFileURL(__filename).href;' },
  legalComments: 'none',
  logLevel: 'warning',
  metafile: true,
});
console.log('  bundle size:', fs.statSync(path.join(BUILD, 'server.cjs')).size, 'bytes; inputs:', Object.keys(bundle.metafile.inputs).join(', '));

if (LAYOUT === 'runtime') {
  await buildRuntimeLayout();
  process.exit(0);
}
if (LAYOUT !== 'sea') die(`unknown --layout ${LAYOUT} (sea | runtime)`);

// ----------------------------------------------------------------- 2. sea-config
step('2/9 sea-config.json');
const assets = { 'index.html': INDEX, 'excel-runner.ps1': RUNNER };
const seaConfig = {
  main: path.join(BUILD, 'server.cjs'),
  output: path.join(BUILD, 'sea-prep.blob'),
  disableExperimentalSEAWarning: true,
  useSnapshot: false,
  useCodeCache: USE_CODE_CACHE, // note: dynamic import() does not work when true (the server uses none)
  execArgvExtension: 'none', // a NODE_OPTIONS set on the user's PC cannot change the exe (verified in R5)
  assets,
};
if (hasBuildSea && !flag('--classic')) seaConfig.output = OUT; // --build-sea writes the exe itself
fs.writeFileSync(path.join(BUILD, 'sea-config.json'), JSON.stringify(seaConfig, null, 2));
console.log(`  assets: ${Object.entries(assets).map(([k, v]) => `${k} (${fs.statSync(v).size} B)`).join(', ')}  useCodeCache=${USE_CODE_CACHE}`);

fs.mkdirSync(path.dirname(OUT), { recursive: true });
replaceOldExe(OUT);

if (hasBuildSea && !flag('--classic')) {
  step('3-6/9 node --build-sea (Node >= 25.5: no postject needed)');
  run(process.execPath, ['--build-sea', path.join(BUILD, 'sea-config.json')]);
} else {
  step('3/9 node --experimental-sea-config');
  run(process.execPath, ['--experimental-sea-config', path.join(BUILD, 'sea-config.json')]);
  const blob = fs.readFileSync(seaConfig.output);
  console.log('  blob size:', blob.length, 'bytes');

  step('4/9 copy node.exe');
  fs.copyFileSync(process.execPath, OUT);

  step('5/9 remove Authenticode signature');
  const signtool = findSigntool();
  if (flag('--keep-signature')) console.log('  --keep-signature: postject will warn "signature seems corrupted" (harmless)');
  else if (signtool) run(signtool, ['remove', '/s', OUT]);
  else console.log('  signtool.exe not found (Windows SDK not installed): skipped; postject warns, the exe still runs');

  step('6/9 postject inject NODE_SEA_BLOB (15-30 s)');
  const { inject } = require('postject');
  await inject(OUT, 'NODE_SEA_BLOB', blob, { sentinelFuse: SENTINEL_FUSE });
}

if (ICON) {
  step('7/9 icon + version info (resedit, after injection)');
  await setIconAndVersion(OUT, path.resolve(ICON));
} else step('7/9 icon: skipped (no scripts/app.ico, use --icon file.ico)');

if (GUI) {
  step('8/9 PE subsystem -> WINDOWS (2): no console window');
  setPeSubsystem(OUT, 2);
} else step('8/9 PE subsystem: console (default; start.bat uses --detach)');

step('9/9 smoke test');
console.log('  exe:', OUT, (fs.statSync(OUT).size / 1048576).toFixed(1), 'MB');
if (!flag('--no-smoke')) {
  const r = await smokeTest(OUT);
  console.log(`  READY after ${r.ms} ms, /api/ping -> ${r.ping} (first run of a new exe includes the Defender scan)`);
}

if (RELEASE) {
  step(`release folder -> ${RELEASE}`);
  fs.mkdirSync(path.join(RELEASE, 'dist'), { recursive: true });
  fs.copyFileSync(OUT, path.join(RELEASE, EXE_NAME));
  for (const f of ['start.bat', 'README.md']) if (fs.existsSync(path.join(ROOT, f))) fs.copyFileSync(path.join(ROOT, f), path.join(RELEASE, f));
  fs.copyFileSync(INDEX, path.join(RELEASE, 'dist', 'index.html')); // standalone fallback of start.bat
}
console.log(`\n[build-exe] done in ${((Date.now() - t0) / 1000).toFixed(1)}s -> ${OUT}`);

// =================================================================== helpers
async function buildRuntimeLayout() {
  const dir = path.resolve(opt('--out', path.join(ROOT, 'release', 'ExcelMerger-runtime')));
  step(`2/3 runtime layout -> ${dir}`);
  for (const sub of ['runtime', 'app', path.join('server', 'excel'), 'dist']) fs.mkdirSync(path.join(dir, sub), { recursive: true });
  fs.copyFileSync(process.execPath, path.join(dir, 'runtime', 'node.exe'));
  fs.copyFileSync(path.join(BUILD, 'server.cjs'), path.join(dir, 'app', 'server.cjs'));
  fs.copyFileSync(RUNNER, path.join(dir, 'server', 'excel', 'excel-runner.ps1'));
  fs.copyFileSync(INDEX, path.join(dir, 'dist', 'index.html'));
  fs.writeFileSync(path.join(dir, 'package.json'), JSON.stringify({ name: 'excel-merger', private: true, version: VERSION }, null, 2));
  for (const f of ['start.bat', 'README.md']) if (fs.existsSync(path.join(ROOT, f))) fs.copyFileSync(path.join(ROOT, f), path.join(dir, f));
  step('3/3 smoke test');
  if (!flag('--no-smoke')) {
    const r = await smokeTest(path.join(dir, 'runtime', 'node.exe'), [path.join(dir, 'app', 'server.cjs')]);
    console.log(`  READY after ${r.ms} ms, /api/ping -> ${r.ping}`);
  }
  console.log(`\n[build-exe] done in ${((Date.now() - t0) / 1000).toFixed(1)}s -> ${dir}`);
}

/**
 * Makes room for the new exe. Windows cannot delete or overwrite an exe that is running (a window of the app is open),
 * but it CAN rename it: the running copy is moved aside as "<name>.old-<time>.exe" and keeps working, and the build
 * goes on instead of failing. Leftovers of earlier builds ("<name>.old-*.exe") are removed when they are no longer in use.
 */
function replaceOldExe(out) {
  const dir = path.dirname(out);
  const stem = path.basename(out, path.extname(out));
  const ext = path.extname(out);
  try {
    if (fs.existsSync(out)) fs.rmSync(out);
  } catch (e) {
    const aside = path.join(dir, `${stem}.old-${Date.now()}${ext}`);
    try {
      fs.renameSync(out, aside);
      console.log(`  ${path.basename(out)} is in use (${e.code}): renamed to ${path.basename(aside)}; the running copy keeps working. Close it and delete the old file.`);
    } catch {
      die(`cannot replace ${out} (${e.code}): close the running ExcelMerger first`);
    }
  }
  const oldBuild = (name) => name.toLowerCase().startsWith(`${stem.toLowerCase()}.old-`) && name.toLowerCase().endsWith(ext.toLowerCase()) && /^\d+$/.test(name.slice(stem.length + 5, name.length - ext.length));
  for (const f of fs.readdirSync(dir)) {
    if (!oldBuild(f)) continue;
    try {
      fs.rmSync(path.join(dir, f));
      console.log(`  removed the old build ${f}`);
    } catch {
      /* still running: stays until it has exited */
    }
  }
}

function run(cmd, args) {
  const r = spawnSync(cmd, args, { stdio: 'inherit' });
  if (r.status !== 0) die(`${path.basename(cmd)} ${args.join(' ')} failed (exit ${r.status})`);
}

function findSigntool() {
  try {
    const p = execFileSync('where.exe', ['signtool'], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }).split(/\r?\n/)[0].trim();
    if (p) return p;
  } catch {
    /* not on PATH */
  }
  const kits = path.join(process.env['ProgramFiles(x86)'] || 'C:\\Program Files (x86)', 'Windows Kits', '10', 'bin');
  if (!fs.existsSync(kits)) return null;
  const vers = fs
    .readdirSync(kits)
    .filter((d) => /^10\./.test(d))
    .sort((a, b) => a.localeCompare(b, 'en', { numeric: true }))
    .reverse();
  for (const v of vers) {
    const p = path.join(kits, v, 'x64', 'signtool.exe');
    if (fs.existsSync(p)) return p;
  }
  return null;
}

/** Patch IMAGE_OPTIONAL_HEADER.Subsystem (offset 0x5C from the PE signature): 2 = GUI, 3 = console. */
function setPeSubsystem(file, value) {
  const fd = fs.openSync(file, 'r+');
  try {
    const hdr = Buffer.alloc(0x40);
    fs.readSync(fd, hdr, 0, 0x40, 0);
    const peOff = hdr.readUInt32LE(0x3c);
    const sig = Buffer.alloc(4);
    fs.readSync(fd, sig, 0, 4, peOff);
    if (sig.toString('latin1') !== 'PE\0\0') throw new Error('not a PE file');
    const b = Buffer.alloc(2);
    b.writeUInt16LE(value);
    fs.writeSync(fd, b, 0, 2, peOff + 0x5c);
  } finally {
    fs.closeSync(fd);
  }
}

async function setIconAndVersion(file, iconPath) {
  const ResEdit = await import('resedit'); // ESM-only package
  const PELibrary = await import('pe-library');
  const exe = PELibrary.NtExecutable.from(fs.readFileSync(file), { ignoreCert: true });
  const res = PELibrary.NtExecutableResource.from(exe);
  const iconFile = ResEdit.Data.IconFile.from(fs.readFileSync(iconPath));
  const groups = ResEdit.Resource.IconGroupEntry.fromEntries(res.entries);
  const groupId = groups.length ? groups[0].id : 1;
  const lang = groups.length ? groups[0].lang : 1033;
  ResEdit.Resource.IconGroupEntry.replaceIconsForResource(res.entries, groupId, lang, iconFile.icons.map((i) => i.data));
  const vi = ResEdit.Resource.VersionInfo.fromEntries(res.entries)[0];
  if (vi) {
    const [a, b, c] = String(VERSION).split('.').map((x) => Number.parseInt(x, 10) || 0);
    for (const l of vi.getAllLanguagesForStringValues()) {
      vi.setStringValues(l, {
        FileDescription: 'Excel 統合ツール',
        ProductName: 'Excel 統合ツール',
        CompanyName: '',
        OriginalFilename: path.basename(file),
        InternalName: 'ExcelMerger',
        LegalCopyright: '',
        FileVersion: VERSION,
        ProductVersion: VERSION,
      });
      vi.removeStringValue(l, 'Comments');
    }
    vi.setFileVersion(a, b, c, 0);
    vi.setProductVersion(a, b, c, 0);
    vi.outputToResourceEntries(res.entries);
  }
  res.outputResource(exe);
  fs.writeFileSync(file, Buffer.from(exe.generate()));
}

/** Starts the exe (or node + bundle) headless on a random port, waits for READY, checks /api/ping, stops it. */
function smokeTest(exe, pre = []) {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    const c = spawn(exe, [...pre, '--foreground'], {
      env: { ...process.env, APP_MODE: 'none', NO_OPEN: '1', PORT: '0', NODE_OPTIONS: '', IDLE_EXIT_MS: '0' },
      stdio: ['ignore', 'pipe', 'pipe'],
      windowsHide: true,
    });
    let buf = '';
    let done = false;
    const finish = (err, val) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      process.stdout.write(buf.replace(/^/gm, '  | '));
      try {
        c.kill();
      } catch {
        /* gone */
      }
      if (err) reject(err);
      else resolve(val);
    };
    const timer = setTimeout(() => finish(new Error(`no READY within 60s. output:\n${buf}`)), 60000);
    const onData = (d) => {
      buf += d.toString('utf8');
      const m = /READY http:\/\/localhost:(\d+)\//.exec(buf);
      if (m && !done) {
        const ms = Date.now() - start;
        http
          .get({ host: '127.0.0.1', port: Number(m[1]), path: '/api/ping', headers: { Host: `localhost:${m[1]}` } }, (res) => {
            let body = '';
            res.on('data', (x) => (body += x));
            res.on('end', () => {
              try {
                if (JSON.parse(body).app !== 'excel-merger') throw new Error(body);
                finish(null, { ms, ping: body });
              } catch (e) {
                finish(new Error(`/api/ping answered ${body}`));
              }
            });
          })
          .on('error', (e) => finish(e));
      }
    };
    c.stdout.on('data', onData);
    c.stderr.on('data', onData);
    c.on('exit', (code) => {
      if (!done) finish(new Error(`exited with ${code} before READY:\n${buf}`));
    });
  });
}
