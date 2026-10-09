/**
 * Builds excel-runner.ps1 from excel-runner.src.ps1 + XlRunnerHelper.cs (Node built-ins only).
 * Used by build-runner.mjs (npm run build:runner) and by the server as a fallback when the built
 * file is missing or older than its sources (dev checkouts).
 *
 * Output: UTF-8 WITH BOM (Windows PowerShell 5.1 reads BOM-less files as the ANSI code page and would
 * garble the Japanese messages) and CRLF line endings.
 */
import fs from 'node:fs';
import path from 'node:path';

export const PLACEHOLDER = '#__CSHARP__';

/** Returns the merged runner text (no BOM; CRLF line endings). Throws on malformed sources. */
export function mergeRunner(srcText, csText) {
  const src = stripBom(srcText).replace(/\r\n/g, '\n');
  const cs = stripBom(csText).replace(/\r\n/g, '\n').replace(/\s+$/, '');
  if (!src.includes(PLACEHOLDER)) throw new Error(`excel-runner.src.ps1 has no ${PLACEHOLDER} placeholder`);
  // The C# text lives in a PowerShell single-quoted here-string: a line starting with '@ would end it.
  if (/^'@/m.test(cs)) throw new Error("XlRunnerHelper.cs must not contain a line starting with '@");
  const merged = src.replace(PLACEHOLDER, () => cs);
  return merged.replace(/\n/g, '\r\n');
}

/** Writes <dir>/excel-runner.ps1 from the sources in <dir>. Returns { file, bytes }. */
export function buildRunnerFile(dir, outFile = path.join(dir, 'excel-runner.ps1')) {
  const src = fs.readFileSync(path.join(dir, 'excel-runner.src.ps1'), 'utf8');
  const cs = fs.readFileSync(path.join(dir, 'XlRunnerHelper.cs'), 'utf8');
  const text = mergeRunner(src, cs);
  const buf = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(text, 'utf8')]);
  const tmp = `${outFile}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, buf);
  fs.renameSync(tmp, outFile);
  return { file: outFile, bytes: buf.length };
}

/** true when <dir>/excel-runner.ps1 is missing or older than one of its sources. */
export function runnerIsStale(dir) {
  const out = path.join(dir, 'excel-runner.ps1');
  let outTime;
  try {
    outTime = fs.statSync(out).mtimeMs;
  } catch {
    return true;
  }
  for (const f of ['excel-runner.src.ps1', 'XlRunnerHelper.cs']) {
    try {
      if (fs.statSync(path.join(dir, f)).mtimeMs > outTime) return true;
    } catch {
      /* source missing (release layout): the built file is all there is */
    }
  }
  return false;
}

function stripBom(s) {
  return s.charCodeAt(0) === 0xfeff ? s.slice(1) : s;
}
