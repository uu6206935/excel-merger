#!/usr/bin/env node
/**
 * make-release.mjs (npm run release) - builds the folder you hand to someone, and a zip of it.
 * It does NOT rebuild ExcelMerger.exe (build.bat / npm run build:exe does that); it only packs the CURRENT project files.
 *
 *   node scripts/make-release.mjs [--exe <ExcelMerger.exe>] [--out <dir>] [--date <yyyyMMdd>] [--min-exe-size <bytes>]
 *
 *   --exe   the exe to ship (default: ExcelMerger.exe in the project root)
 *   --out   where the results go (default: <project>\release)
 *   --date  the yyyyMMdd stamp of the zip name (default: today, local time)
 *   --min-exe-size  smallest exe accepted (default 50 MB; only the tests lower it, to work with a tiny fake exe)
 *
 * Result (<out> = release by default):
 *   <out>\ExcelMerger\                 ExcelMerger.exe, start.bat (byte-identical copy), README.md (without the
 *                                      developer sections), はじめにお読みください.txt, samples\ (5 samples + README.md)
 *   <out>\ExcelMerger_<yyyyMMdd>.zip   that folder, as ONE top-level folder "ExcelMerger/", entry names flagged UTF-8
 *
 * Not included on purpose: node_modules, src, tests, docs, dist, .tmp, test-results, output, truth.json.
 * The script refuses to run when the exe is missing or smaller than 50 MB, re-reads the finished zip to check every
 * entry (name, UTF-8 flag, content) and prints the zip size. Only <out>\ExcelMerger and the zip of that date are replaced.
 */
import { createHash } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(HERE, '..');
const APP = 'ExcelMerger';
const EXE_NAME = 'ExcelMerger.exe';
const TXT_NAME = 'はじめにお読みください.txt';
const MIN_EXE_BYTES = 50 * 1024 * 1024;
const SAMPLE_EXT = /\.(xlsx|xlsm|xls|csv)$/i;
const SAMPLE_COUNT = 5;
/** `## ` sections whose heading matches this are for whoever BUILDS the tool, not for the person who gets it. */
const DEV_HEADING = /開発|作る人向け/;

/** The plain-text note next to start.bat: friendly, short, no Markdown. Written as UTF-8 with BOM and CRLF. */
const TXT_LINES = [
  'Excel 統合ツール  はじめにお読みください',
  '',
  '1. フォルダの中の「start.bat」をダブルクリックすると、ツールが開きます。インストールは不要です。',
  '2. 使い終わったら、ツールのウィンドウを閉じるだけです。',
  '3. 出力したExcelファイルは、このフォルダの「output」フォルダに入ります（保存するときに自動で作られます）。',
  '4. zipで受け取ったときは、展開する前に zip を右クリック →「プロパティ」→「ブロックの解除」（または「許可する」）にチェック → OK。',
  '   そのあと「すべて展開」で展開してから、start.bat を使ってください。',
  '5. 「WindowsによってPCが保護されました」と出たら、「詳細情報」→「実行」を押してください。',
  '6. 使わなくなったら、このフォルダを削除するだけでアンインストール完了です。',
  '7. インターネットがなくても動きます。ファイルが外部に送られることはありません。',
];

// ----------------------------------------------------------------------------------------------- args
const argv = process.argv.slice(2);
const KNOWN = new Set(['--exe', '--out', '--date', '--min-exe-size']);
const args = {};
for (let i = 0; i < argv.length; i++) {
  const a = argv[i];
  if (a === '--help' || a === '-h') {
    console.log('Usage: node scripts/make-release.mjs [--exe <ExcelMerger.exe>] [--out <dir>] [--date <yyyyMMdd>]\n(see the comment at the top of scripts/make-release.mjs)');
    process.exit(0);
  }
  if (!KNOWN.has(a)) die(`unknown option ${a}. Usage: node scripts/make-release.mjs [--exe <ExcelMerger.exe>] [--out <dir>] [--date <yyyyMMdd>]`);
  if (i + 1 >= argv.length || argv[i + 1].startsWith('--')) die(`${a} needs a value`);
  args[a] = argv[++i];
}
const EXE = path.resolve(args['--exe'] ?? path.join(ROOT, EXE_NAME));
const OUT = path.resolve(args['--out'] ?? path.join(ROOT, 'release'));
const MIN_EXE = args['--min-exe-size'] !== undefined ? Number(args['--min-exe-size']) : MIN_EXE_BYTES;
const STAMP = args['--date'] ?? localStamp(new Date());
if (!/^\d{8}$/.test(STAMP)) die(`--date must be yyyyMMdd, got "${STAMP}"`);
if (!Number.isFinite(MIN_EXE) || MIN_EXE < 0) die(`--min-exe-size must be a number of bytes, got "${args['--min-exe-size']}"`);
const FOLDER = path.join(OUT, APP);
const ZIP = path.join(OUT, `${APP}_${STAMP}.zip`);

const t0 = Date.now();
const say = (s) => console.log(`[make-release] ${s}`);
function die(msg) {
  console.error(`\n[make-release] ERROR: ${msg}`);
  process.exit(1);
}
function localStamp(d) {
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}`;
}
const mb = (n) => (n / 1048576).toFixed(1);
const isInside = (child, parent) => {
  const rel = path.relative(path.resolve(parent), path.resolve(child));
  return rel === '' || (!rel.startsWith('..') && !path.isAbsolute(rel));
};

// ----------------------------------------------------------------------------------------------- inputs
if (process.platform !== 'win32') console.warn('[make-release] WARNING: this tool is for Windows (the exe cannot be tested here).');
if (!fs.existsSync(EXE) || !fs.statSync(EXE).isFile()) {
  die(`${EXE} was not found. Build it first (build.bat, or npm run build:exe) or pass --exe <path to ExcelMerger.exe>.`);
}
const exeStat = fs.statSync(EXE);
if (exeStat.size < MIN_EXE) {
  die(`${EXE} is only ${mb(exeStat.size)} MB; the real ExcelMerger.exe is about 90 MB (anything under ${Number((MIN_EXE / 1048576).toFixed(1))} MB is refused). It looks like a broken or partial build: run build.bat again.`);
}
{
  const fd = fs.openSync(EXE, 'r');
  const head = Buffer.alloc(2);
  fs.readSync(fd, head, 0, 2, 0);
  fs.closeSync(fd);
  if (head.toString('latin1') !== 'MZ') die(`${EXE} is not a Windows executable (no "MZ" header).`);
}
if (isInside(EXE, FOLDER)) die(`--exe (${EXE}) is inside the folder that would be replaced (${FOLDER}). Copy the exe somewhere else first.`);

const START_BAT = path.join(ROOT, 'start.bat');
const README = path.join(ROOT, 'README.md');
const SAMPLES = path.join(ROOT, 'samples');
for (const f of [START_BAT, README]) if (!fs.existsSync(f)) die(`${f} is missing.`);
{
  // start.bat must stay ASCII with CRLF only (cmd.exe misreads anything else); it is copied byte for byte
  const b = fs.readFileSync(START_BAT);
  if (b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) die('start.bat has a UTF-8 BOM (it breaks "@echo off").');
  if (b.some((x) => x > 0x7f)) die('start.bat must be ASCII-only.');
  const s = b.toString('latin1');
  if (/(^|[^\r])\n/.test(s) || /\r(?!\n)/.test(s)) die('start.bat must use CRLF line endings only.');
}
const sampleFiles = fs.existsSync(SAMPLES)
  ? fs
      .readdirSync(SAMPLES, { withFileTypes: true })
      .filter((e) => e.isFile() && SAMPLE_EXT.test(e.name) && !e.name.startsWith('~$') && !e.name.startsWith('.'))
      .map((e) => e.name)
      .sort()
  : [];
if (sampleFiles.length < SAMPLE_COUNT) die(`samples\\ needs the ${SAMPLE_COUNT} sample files (.xlsx/.xls/.csv), found ${sampleFiles.length} in ${SAMPLES}.`);
if (!fs.existsSync(path.join(SAMPLES, 'README.md'))) die(`${path.join(SAMPLES, 'README.md')} is missing.`);

warnIfExeIsOlderThanSources();

// ----------------------------------------------------------------------------------------------- the folder
say(`exe: ${EXE} (${mb(exeStat.size)} MB)`);
say(`folder: ${FOLDER}`);
removeOldResults();
try {
  fs.mkdirSync(path.join(FOLDER, 'samples'), { recursive: true });
  fs.copyFileSync(EXE, path.join(FOLDER, EXE_NAME));
  fs.copyFileSync(START_BAT, path.join(FOLDER, 'start.bat'));
  fs.writeFileSync(path.join(FOLDER, 'README.md'), releaseReadme(fs.readFileSync(README, 'utf8')));
  fs.writeFileSync(path.join(FOLDER, TXT_NAME), Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(TXT_LINES.join('\r\n') + '\r\n', 'utf8')]));
  for (const f of sampleFiles) fs.copyFileSync(path.join(SAMPLES, f), path.join(FOLDER, 'samples', f));
  fs.writeFileSync(path.join(FOLDER, 'samples', 'README.md'), samplesReadme(fs.readFileSync(path.join(SAMPLES, 'README.md'), 'utf8')));

  // ------------------------------------------------------------------------------------------- the zip
  const files = listFiles(FOLDER); // posix-style relative paths
  say(`zipping ${files.length} files ...`);
  const zip = new JSZip();
  const now = new Date();
  // JSZip writes the DOS time from the UTC fields; Explorer reads it as local time, so shift it
  const stamp = new Date(now.getTime() - now.getTimezoneOffset() * 60000);
  const order = [TXT_NAME, 'start.bat', 'README.md', EXE_NAME, ...files.filter((f) => f.startsWith('samples/'))];
  for (const dir of [`${APP}/`, `${APP}/samples/`]) zip.file(dir, null, { dir: true, date: stamp });
  for (const rel of order) zip.file(`${APP}/${rel}`, fs.readFileSync(path.join(FOLDER, ...rel.split('/'))), { date: stamp, createFolders: true });
  let lastPct = -1;
  const buf = await zip.generateAsync(
    { type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 }, platform: 'DOS' },
    (m) => {
      const pct = Math.floor(m.percent / 20) * 20;
      if (pct !== lastPct && pct > 0) {
        lastPct = pct;
        say(`  zip ${pct}%`);
      }
    },
  );
  const partial = `${ZIP}.partial`;
  fs.writeFileSync(partial, buf);
  fs.renameSync(partial, ZIP);

  // ------------------------------------------------------------------------------------------- check the zip
  await verifyZip(fs.readFileSync(ZIP), files);
} catch (e) {
  fs.rmSync(FOLDER, { recursive: true, force: true });
  fs.rmSync(ZIP, { force: true });
  fs.rmSync(`${ZIP}.partial`, { force: true });
  die(`${e && e.stack ? e.stack : e}`);
}

const zipBytes = fs.statSync(ZIP).size;
console.log('');
say(`done in ${((Date.now() - t0) / 1000).toFixed(1)}s`);
say(`folder: ${FOLDER}`);
for (const f of listFiles(FOLDER)) say(`  ${APP}/${f}  (${fs.statSync(path.join(FOLDER, ...f.split('/'))).size.toLocaleString('en-US')} bytes)`);
say(`zip:    ${ZIP}`);
say(`zip size: ${mb(zipBytes)} MB (${zipBytes.toLocaleString('en-US')} bytes)`);
say(`zip sha256: ${createHash('sha256').update(fs.readFileSync(ZIP)).digest('hex')}`);

// =================================================================================== helpers

/** Removes the previous results (folder + zip of this date). Refuses to delete a folder that is not a hand-out folder. */
function removeOldResults() {
  fs.mkdirSync(OUT, { recursive: true });
  if (fs.existsSync(FOLDER)) {
    const names = fs.readdirSync(FOLDER);
    const looksRight = names.length === 0 || names.includes(EXE_NAME) || names.includes('start.bat') || names.includes(TXT_NAME);
    if (!looksRight) die(`${FOLDER} exists but does not look like a hand-out folder (no ${EXE_NAME} / start.bat). Check it and delete it by hand.`);
    try {
      fs.rmSync(FOLDER, { recursive: true, force: true, maxRetries: 3, retryDelay: 300 });
    } catch (e) {
      die(`cannot replace ${FOLDER} (${e.code}): close the tool that was started from this folder (ExcelMerger.exe / Explorer window) and run again.`);
    }
  }
  try {
    fs.rmSync(ZIP, { force: true });
    fs.rmSync(`${ZIP}.partial`, { force: true });
  } catch (e) {
    die(`cannot replace ${ZIP} (${e.code}): close the program that has the zip open and run again.`);
  }
}

/** Relative file paths (forward slashes) below dir, sorted. */
function listFiles(dir, prefix = '') {
  const out = [];
  for (const e of fs.readdirSync(path.join(dir, prefix), { withFileTypes: true })) {
    const rel = prefix ? `${prefix}/${e.name}` : e.name;
    if (e.isDirectory()) out.push(...listFiles(dir, rel));
    else out.push(rel);
  }
  return out.sort();
}

/** Warns (never refuses) when the exe is older than the sources it was built from, e.g. server/ src/ dist/index.html. */
function warnIfExeIsOlderThanSources() {
  let newest = { t: 0, file: '' };
  const visit = (p) => {
    let st;
    try {
      st = fs.statSync(p);
    } catch {
      return;
    }
    if (st.isDirectory()) {
      for (const n of fs.readdirSync(p)) if (n !== 'node_modules' && !n.startsWith('.')) visit(path.join(p, n));
    } else if (st.mtimeMs > newest.t) newest = { t: st.mtimeMs, file: p };
  };
  for (const p of [path.join(ROOT, 'server'), path.join(ROOT, 'src'), path.join(ROOT, 'dist', 'index.html'), path.join(ROOT, 'package.json')]) visit(p);
  if (newest.t > exeStat.mtimeMs + 2000) {
    console.warn(
      `[make-release] WARNING: ${path.basename(EXE)} (${exeStat.mtime.toISOString()}) is older than ${path.relative(ROOT, newest.file)} (${new Date(newest.t).toISOString()}). ` +
        'If that change should reach the recipient, rebuild the exe first (build.bat).',
    );
  }
}

/**
 * Removes every level-2 (`## `) section whose heading matches `re`, together with the `---` rule in front of it.
 * Headings inside ``` fences are ignored. Keeps the line ending style of the input.
 */
function stripSections(text, re) {
  const eol = text.includes('\r\n') ? '\r\n' : '\n';
  const out = [];
  let fence = false;
  let skipping = false;
  for (const line of text.split(/\r?\n/)) {
    if (/^\s*```/.test(line)) fence = !fence;
    const heading = !fence && /^## /.test(line);
    if (heading) {
      skipping = re.test(line);
      if (skipping) {
        // drop the separator rule (and blank lines) just before the removed section
        while (out.length && out[out.length - 1].trim() === '') out.pop();
        if (out.length && out[out.length - 1].trim() === '---') out.pop();
        while (out.length && out[out.length - 1].trim() === '') out.pop();
        out.push('');
      }
    }
    if (!skipping) out.push(line);
  }
  while (out.length && out[out.length - 1].trim() === '') out.pop();
  return out.join(eol) + eol;
}

/** README.md for the recipient: the same text without the sections for developers / for whoever builds the zip. */
function releaseReadme(text) {
  return stripSections(text, DEV_HEADING);
}

/** samples/README.md for the recipient: no developer sections, no lines about npm / tools\ / truth.json. */
function samplesReadme(text) {
  const kept = stripSections(text, DEV_HEADING)
    .split(/\r?\n/)
    .filter((l) => !/node tools[\\/]|tools[\\/]|npm run|truth\.json|--base /.test(l));
  // collapse the blank lines a removed line may leave behind
  const out = [];
  for (const l of kept) {
    if (l.trim() === '' && out.length && out[out.length - 1].trim() === '') continue;
    out.push(l);
  }
  while (out.length && out[out.length - 1].trim() === '') out.pop();
  const s = out.join('\n') + '\n';
  if (/truth\.json|tools[\\/]|npm /.test(s)) console.warn('[make-release] WARNING: samples/README.md still mentions developer commands; edit samples/README.md.');
  return s;
}

/** Parses the central directory of a zip: [{ name, rawName, utf8, method, size }]. No zip64 (the package is far below 4 GB). */
function readCentralDirectory(buf) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 22 - 65535); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error('zip: end of central directory not found');
  const count = buf.readUInt16LE(eocd + 10);
  let off = buf.readUInt32LE(eocd + 16);
  const entries = [];
  for (let n = 0; n < count; n++) {
    if (buf.readUInt32LE(off) !== 0x02014b50) throw new Error(`zip: bad central directory entry #${n}`);
    const flags = buf.readUInt16LE(off + 8);
    const nameLen = buf.readUInt16LE(off + 28);
    const extraLen = buf.readUInt16LE(off + 30);
    const commentLen = buf.readUInt16LE(off + 32);
    const rawName = Buffer.from(buf.subarray(off + 46, off + 46 + nameLen));
    entries.push({ name: rawName.toString('utf8'), rawName, utf8: (flags & 0x800) !== 0, method: buf.readUInt16LE(off + 10), size: buf.readUInt32LE(off + 24) });
    off += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

/** Re-reads the finished zip: one top-level folder, UTF-8 names (flag + valid UTF-8), exactly the files of the folder, same bytes. */
async function verifyZip(zipBuf, files) {
  const entries = readCentralDirectory(zipBuf);
  const expected = new Set([`${APP}/`]);
  for (const f of files) {
    expected.add(`${APP}/${f}`);
    const parts = f.split('/');
    for (let i = 1; i < parts.length; i++) expected.add(`${APP}/${parts.slice(0, i).join('/')}/`);
  }
  const got = new Set();
  for (const e of entries) {
    if (e.name.includes('\\') || e.name.startsWith('/') || e.name.split('/').includes('..')) throw new Error(`zip: bad entry name ${JSON.stringify(e.name)}`);
    if (/[^\x00-\x7f]/.test(e.name) && !e.utf8) throw new Error(`zip: entry ${e.name} has a non-ASCII name without the UTF-8 flag (Windows would show it garbled)`);
    if (!Buffer.from(e.name, 'utf8').equals(e.rawName)) throw new Error(`zip: entry name is not valid UTF-8: ${e.name}`);
    if (!expected.has(e.name)) throw new Error(`zip: unexpected entry ${e.name}`);
    got.add(e.name);
  }
  for (const n of expected) if (!got.has(n)) throw new Error(`zip: missing entry ${n}`);
  if (got.size !== entries.length) throw new Error('zip: duplicate entry names');
  const zip = await JSZip.loadAsync(zipBuf, { checkCRC32: true });
  for (const f of files) {
    const entry = zip.file(`${APP}/${f}`);
    if (!entry) throw new Error(`zip: cannot read ${f}`);
    const bytes = await entry.async('nodebuffer');
    if (!bytes.equals(fs.readFileSync(path.join(FOLDER, ...f.split('/'))))) throw new Error(`zip: ${f} differs from the folder`);
  }
  say(`zip checked: ${entries.length} entries, names are UTF-8, contents match the folder`);
}
