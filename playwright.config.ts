import { defineConfig, devices } from '@playwright/test';

const testPort = process.env.PLAYWRIGHT_TEST_PORT ?? '4173';

export default defineConfig({
  testDir: './tests/e2e',
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 1 : 0,
  reporter: process.env.CI ? 'github' : 'list',
  use: { baseURL: `http://127.0.0.1:${testPort}`, trace: 'retain-on-failure' },
  webServer: {
    command: 'node scripts/serve-test-build.mjs',
    url: `http://127.0.0.1:${testPort}`,
    reuseExistingServer: false,
  },
  projects: [
    { name: 'chromium', use: { ...devices['Desktop Chrome'] } },
    { name: 'firefox', use: { ...devices['Desktop Firefox'] } },
    { name: 'webkit', use: { ...devices['Desktop Safari'] } },
  ],
});
