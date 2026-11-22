import { defineConfig, devices } from '@playwright/test';

const port = process.env.PLAYWRIGHT_PYTHON_PORT ?? '4175';
const executablePath = process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE;

export default defineConfig({
  testDir: './tests/e2e',
  testMatch: 'python-execution.spec.ts',
  outputDir: './test-results/python-execution',
  timeout: 480_000,
  workers: 1,
  reporter: process.env.CI ? 'github' : 'list',
  use: {
    baseURL: `http://127.0.0.1:${port}`,
    trace: 'retain-on-failure',
    launchOptions: executablePath ? { executablePath } : {},
  },
  webServer: {
    command: 'node scripts/serve-test-build.mjs',
    url: `http://127.0.0.1:${port}`,
    reuseExistingServer: false,
    env: { PLAYWRIGHT_TEST_PORT: port },
  },
  projects: [{ name: 'chromium-python-runtime', use: { ...devices['Desktop Chrome'] } }],
});
