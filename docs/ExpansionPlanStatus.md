# Expansion Plan Status

Live status tracker for [ExpansionPlan.md](ExpansionPlan.md). Update after every sprint with exact verification counts and current risks.

---

## Sprint Matrix

| Sprint | Scope | Status | Verification | Date |
| --- | --- | --- | --- | --- |
| **1.1** | Excel COM oracle harness, `src/calculation/functions/` modularization, `#SPILL!`/`#CALC!` errors | Done | `npm run check` (45/45 files: 1177 passed, 5 skipped, 0 failed) + Excel 16.0.17932 smoke oracle (15 captured, 10 run & passed, 5 skipped pending Sprints 1.3–1.4) | 2026-11-27 |
| **1.2** | Parser & reference upgrade (`#`, `@`, structured refs, 3D refs, higher-order calls) | Done | `npm run check` (46/46 files: 1184 passed, 5 skipped, 0 failed; `vite build` 72 files, 43.99 MiB) | 2026-11-27 |
| **1.3** | Dynamic array spilling, broadcasting, `LET`/`LAMBDA`/helpers, volatile/dynamic refs | Done | `npm run check` (47/47 files: 1221 passed, 5 skipped, 0 failed) + Excel 16.0.17932 dynamic-arrays oracle (34/34 passed) + Chromium e2e (5/5 passed) | 2026-11-27 |
| **1.4** | Core Function Wave A (Lookup, Dynamic Arrays, Logical, Info — 57 functions, 97 total) | Done | `npm run check` (48/48 files: 1413 passed, 2 skipped, 0 failed; `vite build` 72 files, 44.07 MiB) + Excel 16.0.17932 Wave A oracles (75 logical/info + 63 lookup + 47 shaping = 185/185 passed) | 2026-11-27 |
| **1.5** | Core Function Wave B (Math, Aggregates, Text, Date/Time, Core Stats/Financial — 128 functions, 225 total) + `_xlfn.` XLSX + sparse lookups | Done | `npm run check` (49/49 files: 1693 passed, 2 skipped, 0 failed; `vite build` 72 files, 44.18 MiB) + Excel 16.0.17932 Wave B oracles (91 math + 61 text + 42 datetime + 82 stats/financial = 276/276 passed) + `test:performance` (5/5) + Chromium e2e (10/10) + `test:worksheet` (12/12 across Chromium/Firefox/WebKit) | 2026-11-27 |
| **Phase 2** | PivotTables *(detail sprints after Phase 1)* | Planned (detail after prior phase) | — | — |
| **Phase 3** | Charts — Excel variety *(detail sprints after Phase 2)* | Planned (detail after prior phase) | — | — |
| **Phase 4** | Remaining Long-Tail Excel Functions *(detail waves after Phase 1/3)* | Planned (detail after prior phase) | — | — |

---

## Latest Verification Ledger

- **Date**: 2026-11-27 (Sprint 1.5 — Phase 1 Complete)
- **Excel Oracle Build Version**: Excel `16.0` build `17932` (`smoke-excel.json`, `dynamic-arrays.json`, `logical-info.json`, `lookup.json`, `shaping.json`, `math.json`, `text.json`, `datetime.json`, `stats-financial.json`)
- **Commands Run**:
  - `powershell -ExecutionPolicy Bypass -File scripts/generate-excel-oracle.ps1` for `math.spec.json` (`91`), `text.spec.json` (`61`), `datetime.spec.json` (`42`), and `stats-financial.spec.json` (`82`) -> exit 0 (`276` Wave B cases captured).
  - `npm run check` (`tsc --noEmit -p tsconfig.model.json && tsc --noEmit`, `vitest run`, `vite build` + `finalize-static-build.mjs`) -> exit 0 (`72` files, `44.18 MiB` uncompressed / `33.58 MiB` gzip).
  - `npm run test:performance` (`vitest run --config vitest.performance.config.ts`) -> exit 0 (`5/5` passed, `24.82s`).
  - `npm run test:e2e -- tests/e2e/dynamic-arrays.spec.ts tests/e2e/ribbon-parity.spec.ts tests/e2e/formulas.spec.ts tests/e2e/xlsx-files.spec.ts --project=chromium` -> exit 0 (`10/10` passed, `49.5s`).
  - `npm run test:worksheet` (`vite build --config vite.worksheet.config.ts && playwright test --config playwright.worksheet.config.ts`) -> exit 0 (`12/12` passed across Chromium, Firefox, WebKit, `33.3s`).
- **Pass / Fail / Skipped**: Vitest `49/49` files passed (`1693` passed, `0` failed, `2` skipped in `smoke-excel.json` for `=SEQUENCE(2)` and `=IFERROR(SEQUENCE(2),"caught")` which require an external setup blocker); Wave B Excel COM oracles `276/276` passed (`91` math, `61` text, `42` datetime, `82` stats/financial), Wave A `185/185`, `dynamic-arrays.json` `34/34`, `smoke-excel.json` `13/15` (`2` skipped); `test:performance` `5/5`; Chromium e2e `10/10`; `test:worksheet` `12/12`.
- **Unrun**: Full Playwright cross-browser e2e matrix for all remaining non-formula/worksheet specs (`firefox`, `webkit`), `npm run test:python-execution`, and `npm run test:python-probe` were not run in this session (local Windows Firefox launches fail outside `worksheet-fixture` due to `nsIWindowsRegKey.create` noted in `Known-Issues.md`).

---

## Open Issues & Risks

1. **Performance budgets under dynamic arrays (`npm run test:performance`)**: Verified passing (`5/5`) at Phase 1 completion (10,000-cell interaction target and 200k-cell/10k-formula recalculation); monitor regression on future Pivot/Chart additions.
2. **XLSX `_xlfn.` & `.exlt` v1 additive compatibility**: Resolved for Phase 1 (`src/persistence/xlsx/xlsx-formulas.ts` strips `_xlfn.` / `_xlfn._xlws.` on import and re-applies on export, and native spill fields remain optional on `.exlt` v1); future Excel functions (Phase 4) must extend the `_xlfn.` tables alongside registration.
3. **Pure-ECMAScript model boundary (`tsconfig.model.json`)**: Maintained (`tsc --noEmit -p tsconfig.model.json` exited 0); `src/core` and `src/application` keep zero DOM/React/browser dependencies.
