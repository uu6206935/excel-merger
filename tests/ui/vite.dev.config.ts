// Dev-server config for the UI checks (port 5174): the project config + a file watcher that ignores build
// outputs (a locked ExcelMerger.exe being written by the exe build crashes Vite's watcher with EBUSY).
//   npx vite --config tests/ui/vite.dev.config.ts --port 5174 --strictPort
import { fileURLToPath } from 'node:url';
import { defineConfig, mergeConfig } from 'vite';
import base from '../../vite.config';

export default mergeConfig(
  base,
  defineConfig({
    root: fileURLToPath(new URL('../..', import.meta.url)),
    server: {
      port: 5174,
      strictPort: true,
      watch: { ignored: ['**/*.exe', '**/*.blob', '**/dist/**', '**/build/**', '**/release/**', '**/tests/**', '**/samples/**', '**/server/**', '**/tools/**', '**/node_modules/**'] },
    },
  }),
);
