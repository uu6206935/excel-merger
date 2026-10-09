#!/usr/bin/env node
/**
 * npm run build:runner
 * Merges XlRunnerHelper.cs into excel-runner.src.ps1 and writes excel-runner.ps1 (UTF-8 with BOM, CRLF).
 * Replaces the research-time build-runner.py.
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildRunnerFile } from './merge-runner.mjs';

const dir = path.dirname(fileURLToPath(import.meta.url));
try {
  const { file, bytes } = buildRunnerFile(dir);
  console.log(`[build-runner] ${file} (${bytes} bytes, UTF-8 BOM, CRLF)`);
} catch (e) {
  console.error(`[build-runner] 失敗しました / failed: ${e && e.message}`);
  process.exit(1);
}
