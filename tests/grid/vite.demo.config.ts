/**
 * Vite config for the grid demo (owner: grid G), used by the grid's browser tests.
 *
 *  dev   : npx vite --config tests/grid/vite.demo.config.ts        -> http://127.0.0.1:5175/grid-demo.html
 *          (same as `npx vite --port 5175 --strictPort`, but the file watcher ignores build outputs such as
 *           ExcelMerger.exe: watching a locked .exe crashes Vite with EBUSY while the exe is being built)
 *  build : npx vite build --config tests/grid/vite.demo.config.ts  -> tests/grid/.demo-dist/grid-demo.html
 *          (production single file, works from file://; used by tests/grid/e2e/perf.spec.ts)
 */
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { viteSingleFile } from 'vite-plugin-singlefile';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../..', import.meta.url));

export default defineConfig({
  root,
  base: './',
  logLevel: 'warn',
  plugins: [react(), viteSingleFile()],
  server: {
    port: 5175,
    strictPort: true,
    host: '127.0.0.1',
    watch: {
      ignored: ['**/*.exe', '**/*.blob', '**/dist/**', '**/samples/**', '**/.demo-dist/**', '**/.pw-results/**', '**/shots/**', '**/*.xlsx'],
    },
  },
  build: {
    outDir: fileURLToPath(new URL('./.demo-dist', import.meta.url)),
    emptyOutDir: true,
    target: 'es2020',
    chunkSizeWarningLimit: 8000,
    rollupOptions: { input: fileURLToPath(new URL('../../grid-demo.html', import.meta.url)) },
  },
});
