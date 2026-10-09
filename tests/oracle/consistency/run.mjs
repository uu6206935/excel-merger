// Entry point of the consistency oracle (see oracle.ts). Run through run.ps1, or directly:
//   node tests/oracle/consistency/run.mjs --out <dir> [--count 320] [--seed-base 1] [--only <regex>] [--no-runner] [--no-excel] [--reuse]
// (TypeScript sources are loaded through tsx.)
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { register } from 'tsx/esm/api';

// tests/com/lib.ts (shared with the vitest suites) uses CommonJS's __dirname, which vitest provides and a plain ES module does not
globalThis.__dirname ??= path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', 'com');
register();
const { main } = await import('./oracle.ts');
process.exitCode = await main(process.argv.slice(2));
