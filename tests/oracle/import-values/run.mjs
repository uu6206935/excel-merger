// Entry point of the import-values oracle (see oracle.ts). Run through run.ps1, or directly:
//   node tests/oracle/import-values/run.mjs --corpus <dir> --out <dir> [--only <regex>] [--configs all|quick] [--no-excel]
// (TypeScript sources are loaded through tsx.)
import { register } from 'tsx/esm/api';

register();
const { main } = await import('./oracle.ts');
await main(process.argv.slice(2));
