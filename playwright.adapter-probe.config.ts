import { defineConfig } from '@playwright/test';
import baseline from './playwright.config';

export default defineConfig({
  ...baseline,
  testDir: './tests/adapter-browser',
  outputDir: './test-results/adapter-probe-tests',
  webServer: {
    command: 'node scripts/serve-test-build.mjs test-results/adapter-probe-build',
    url: 'http://127.0.0.1:4173',
    reuseExistingServer: false,
  },
});
