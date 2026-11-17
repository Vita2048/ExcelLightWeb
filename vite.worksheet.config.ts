import react from '@vitejs/plugin-react';
import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vite';

export default defineConfig({
  plugins: [react()],
  root: fileURLToPath(new URL('./tests/worksheet-browser', import.meta.url)),
  base: './',
  build: {
    outDir: fileURLToPath(new URL('./test-results/worksheet-build', import.meta.url)),
    emptyOutDir: true,
  },
});
