#!/usr/bin/env node
/**
 * Developer CLI for the Excel本体 engine (not bundled into the exe).
 *   node server/excel/cli.mjs --detect          registry check (no Excel start)
 *   node server/excel/cli.mjs --probe           starts a hidden Excel once, prints version / PID
 *   node server/excel/cli.mjs <job.json>        runs a job (paths inside the job must be absolute)
 * Env: XLRUNNER_INVOKE=file|scriptblock|auto, JOB_TIMEOUT_MS, JOB_IDLE_MS
 */
import fsp from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { detectExcel, runRunner } from './run-job.mjs';
import { buildRunnerFile, runnerIsStale } from './merge-runner.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const arg = process.argv[2];
const show = (ev) => {
  if (ev.type !== 'result' && ev.type !== 'error') console.log(JSON.stringify(ev));
};

async function main() {
  if (!arg || arg === '--help') {
    console.log('usage: node server/excel/cli.mjs --detect | --probe | <job.json>');
    return 0;
  }
  if (arg === '--detect') {
    const t = Date.now();
    console.log(JSON.stringify({ ...(await detectExcel()), ms: Date.now() - t }, null, 2));
    return 0;
  }
  if (runnerIsStale(here)) buildRunnerFile(here);
  const runner = path.join(here, 'excel-runner.ps1');
  const env = { XLRUNNER_CACHE: path.join(process.env.LOCALAPPDATA || os.tmpdir(), 'ExcelMerger', 'cache') };
  const invoke = process.env.XLRUNNER_INVOKE || 'auto';
  if (arg === '--probe') {
    const r = await runRunner({ runner, probe: true, onEvent: show, timeoutMs: 120_000, idleMs: 120_000, env, invoke });
    console.log(JSON.stringify(r, null, 2));
    return r.ok ? 0 : 1;
  }
  const jobFile = path.resolve(arg);
  JSON.parse(await fsp.readFile(jobFile, 'utf8')); // validate early
  const r = await runRunner({
    runner,
    jobFile,
    onEvent: show,
    env,
    invoke,
    timeoutMs: Number(process.env.JOB_TIMEOUT_MS) || undefined,
    idleMs: Number(process.env.JOB_IDLE_MS) || undefined,
  });
  console.log(JSON.stringify(r, null, 2));
  return r.ok ? 0 : 1;
}

main().then(
  (code) => {
    process.exitCode = code;
  },
  (e) => {
    console.error(e);
    process.exitCode = 1;
  },
);
