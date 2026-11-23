import { defineConfig } from '@playwright/test';
import baseline from './playwright.config';

const testPort = process.env.PLAYWRIGHT_TEST_PORT ?? '4173';

export default defineConfig({
  ...baseline,
  testDir: './tests/worksheet-browser',
  outputDir: './test-results/worksheet-browser',
  webServer: {
    ...baseline.webServer,
    command: 'node scripts/serve-test-build.mjs test-results/worksheet-build',
    url: `http://127.0.0.1:${testPort}`,
  },
});
