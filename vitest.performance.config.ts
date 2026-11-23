import { defineConfig } from 'vitest/config';

// Maximum-bound measurements are deliberate, slow workloads. Run them with `npm run test:performance`.
export default defineConfig({
  test: { include: ['tests/performance/**/*.test.ts'], testTimeout: 300_000, hookTimeout: 300_000 },
});
