import { defineConfig } from '@playwright/test';
import baseline from './playwright.config';

export default defineConfig({
  ...baseline,
  testDir: './tests/python-browser',
  outputDir: './test-results/python-probe-tests',
  timeout: 420_000,
  webServer: {
    command: 'node scripts/serve-test-build.mjs test-results/python-probe-build',
    url: 'http://127.0.0.1:4173',
    reuseExistingServer: false,
  },
});
