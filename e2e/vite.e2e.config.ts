/**
 * Vite dev server for E2E runs against the standalone UI (owner: T):
 *   npx vite --config e2e/vite.e2e.config.ts          -> http://127.0.0.1:5176
 *   E2E_BASE_URL=http://127.0.0.1:5176 npx playwright test
 * Same as the project config, but on the E2E port and without watching build outputs / samples (the file
 * watcher crashes with EBUSY while ExcelMerger.exe is being written by a concurrent build) and without HMR.
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig, mergeConfig } from 'vite';
import base from '../vite.config';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

export default mergeConfig(
  base,
  defineConfig({
    root,
    server: {
      port: 5176,
      strictPort: true,
      host: '127.0.0.1',
      // no HMR: other engineers edit src/ all the time and a full reload would wipe the page mid-test
      hmr: false,
      watch: {
        ignored: ['**/ExcelMerger.exe', '**/*.exe', '**/dist/**', '**/output/**', '**/test-results/**', '**/samples/**', '**/runtime/**', '**/build/**'],
      },
    },
  }),
);
