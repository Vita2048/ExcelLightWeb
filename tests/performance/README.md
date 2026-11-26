# performance

Deterministic workloads, environment metadata and p50/p95 measurements; budgets documented in Development.

Run `npm run test:performance`. The separate `vitest.performance.config.ts` keeps these deliberate maximum-bound workloads out of the fast `npm test` and `npm run check` loops. Each run records OS, CPU, core count, RAM, Node version, fixture dimensions and p50/p95 samples.

| Workload | Fixture | Recorded evidence |
| --- | --- | --- |
| Workbook stress | 2 sheets x 1,000 rows x 100 columns = 200,000 populated cells and 10,000 formulas | validation and session construction, single-cell session edit and undo p50/p95, full generation and dirty-chain recalculation |
| Fixed dependency chain | 1,000-formula chain (`A2:A1000` plus `SUM`) | warm recalculation p50/p95 for the Development under-500 ms target |
| Browser responsiveness | production build, playlist of typing and scrolling interactions on a populated worksheet | typing/selection p95 and dropped scroll frames in `tests/e2e/performance.spec.ts` |

Results are written to `test-results/performance/*.json` (ignored by Git). The recorded reference measurements and their machines are summarized in [Development](../../docs/Development.md#cross-browser-parity-and-performance). A measurement on different hardware is not a substitute for the recorded baseline: re-run and record the new environment instead of comparing unlabelled numbers.
