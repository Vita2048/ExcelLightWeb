import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vite';

export default defineConfig({
  root: fileURLToPath(new URL('./tests/adapter-browser', import.meta.url)),
  base: './',
  build: {
    outDir: fileURLToPath(new URL('./test-results/adapter-probe-build', import.meta.url)),
    emptyOutDir: true,
  },
});
