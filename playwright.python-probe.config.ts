import { defineConfig } from '@playwright/test';
import baseline from './playwright.config';

const testPort = process.env.PLAYWRIGHT_TEST_PORT ?? '4173';

export default defineConfig({
  ...baseline,
  testDir: './tests/python-browser',
  outputDir: './test-results/python-probe-tests',
  timeout: 420_000,
  use: { ...baseline.use, baseURL: `http://127.0.0.1:${testPort}` },
  webServer: {
    command: 'node scripts/serve-test-build.mjs test-results/python-probe-build',
    url: `http://127.0.0.1:${testPort}`,
    reuseExistingServer: false,
  },
});
