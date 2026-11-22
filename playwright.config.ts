import { defineConfig, devices } from '@playwright/test';

const testPort = process.env.PLAYWRIGHT_TEST_PORT ?? '4173';
const chromiumExecutablePath = process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE;

export default defineConfig({
  testDir: './tests/e2e',
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 1 : 0,
  workers: process.env.PLAYWRIGHT_WORKERS ? Number(process.env.PLAYWRIGHT_WORKERS) : 2,
  reporter: process.env.CI ? 'github' : 'list',
  use: { baseURL: `http://127.0.0.1:${testPort}`, trace: 'retain-on-failure' },
  webServer: {
    command: 'node scripts/serve-test-build.mjs',
    url: `http://127.0.0.1:${testPort}`,
    reuseExistingServer: false,
  },
  projects: [
    { name: 'chromium', use: { ...devices['Desktop Chrome'], ...(chromiumExecutablePath && {
      launchOptions: { executablePath: chromiumExecutablePath },
    }) } },
    { name: 'firefox', use: { ...devices['Desktop Firefox'] } },
    { name: 'webkit', use: { ...devices['Desktop Safari'] } },
  ],
});
