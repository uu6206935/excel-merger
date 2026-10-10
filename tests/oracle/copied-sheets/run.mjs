// Entry point of the copied-sheets oracle (see oracle.ts). Run through run.ps1, or directly:
//   node tests/oracle/copied-sheets/run.mjs --out <dir> [--only <regex>] [--limit N] [--check-output N] [--negative] [--seed-base 0] [--no-runner] [--no-excel] [--reuse]
// (TypeScript sources are loaded through tsx.)
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { register } from 'tsx/esm/api';

// tests/com/lib.ts (shared with the vitest suites) uses CommonJS's __dirname, which vitest provides and a plain ES module does not
globalThis.__dirname ??= path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', 'com');
register();
const { main } = await import('./oracle.ts');
process.exitCode = await main(process.argv.slice(2));
