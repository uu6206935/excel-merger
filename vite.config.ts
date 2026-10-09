import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { viteSingleFile } from 'vite-plugin-singlefile';

// Single-file build so that dist/index.html also works when opened directly
// from the file system (file://) without the local server (standalone mode).
export default defineConfig({
  base: './',
  plugins: [react(), viteSingleFile()],
  server: {
    port: 5174,
    host: '127.0.0.1',
    // dev server only: answer requests whose Host header names this PC (DNS-rebinding guard, like server/server.mjs)
    allowedHosts: ['localhost', '127.0.0.1'],
    // Build outputs, samples and the packaged runtime are not sources: watching them crashed the dev server with
    // EBUSY while ExcelMerger.exe / the SEA blob were being written (and wastes watchers on big folders).
    watch: {
      ignored: [
        '**/*.exe',
        '**/*.blob',
        '**/dist/**',
        '**/output/**',
        '**/samples/**',
        '**/runtime/**',
        '**/app/**',
        '**/build/**',
        '**/release/**',
        '**/test-results/**',
      ],
    },
  },
  build: { outDir: 'dist', emptyOutDir: true, target: 'es2020', chunkSizeWarningLimit: 8000 },
});
