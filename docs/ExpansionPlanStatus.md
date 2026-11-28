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
| **2.1** | PivotTable core model, aggregation engine, `'pivot'` generated-region ownership & `.exlt` v1 persistence | Done | `npm run check` (53/53 files: 1712 passed, 2 skipped, 0 failed; `pivot-engine` 4/4, `pivot-commands` 3/3; `vite build` 72 files, 44.25 MiB) | 2026-11-27 |
| **2.2** | Create PivotTable dialog, docked `PivotTable Fields` pane & `Value Field Settings` | Done | `npm run check` (53/53 files: 1712 passed, 2 skipped, 0 failed; `vite build` 72 files, 44.29 MiB) + Chromium Playwright (`pivot-tables.spec.ts` 1/1 + `ribbon-parity.spec.ts` 5/5 = 6/6 passed) | 2026-11-27 |
| **2.3** | Compact/Outline/Tabular layouts, expand/collapse `+/-`, in-sheet filters/sort & Show Details drill-down | Done | `npm run check` (53/53 files: 1716 passed, 2 skipped, 0 failed; `pivot-engine` 7/7, `pivot-commands` 4/4; `vite build` 72 files, 44.31 MiB) + Chromium Playwright (`pivot-tables.spec.ts` 4/4 + `ribbon-parity.spec.ts` 5/5 = 9/9 passed) | 2026-11-27 |
| **2.4** | Contextual `PivotTable Analyze` & `Design` ribbon tabs, style options & PivotTable styles | Done | `npm run check` (53/53 files: 1719 passed, 2 skipped, 0 failed; `pivot-engine` 9/9, `pivot-commands` 5/5; `vite build` 72 files, 44.35 MiB) + Chromium Playwright (`pivot-tables.spec.ts` 5/5 + `ribbon-parity.spec.ts` 6/6 = 11/11 passed) | 2026-11-27 |
| **2.5** | `GETPIVOTDATA`, mixed recalc propagation, XLSX materialized boundary & Phase 2 completion gate | Done | `npm run check` (53/53 files: 1741 passed, 2 skipped, 0 failed; `vite build` 72 files, 44.37 MiB) + Excel 16.0.17932 `GETPIVOTDATA` oracle (20/20 passed) + `test:performance` (6/6) + Chromium full e2e (108/108) + WebKit Pivot/Ribbon/Formulas (14/14 with `--timeout=60000` on `formulas.spec.ts`) + `test:worksheet` (12/12 across Chromium/Firefox/WebKit) | 2026-11-27 |
| **3.1** | Chart core model, data binding (`range`/`A1#`/`table`/`pivot`), structural rewrites & `.exlt` v1 persistence | Done | `npm run check` (55/55 files: 1751 passed, 2 skipped, 0 failed; `chart-model` 5/5, `chart-commands` 3/3; `vite build` 72 files, 44.41 MiB) | 2026-11-28 |
| **3.2** | Core 2D chart engine, floating worksheet overlay & `Insert → Charts` Excel UI | Done | `npm run check` (56/56 files: 1755 passed, 2 skipped, 0 failed; `chart-layout` 4/4; `vite build` 72 files, 44.48 MiB) + Chromium Playwright (`charts.spec.ts` 1/1 + `ribbon-parity.spec.ts` 6/6 = 7/7 passed) | 2026-11-28 |
| **3.3** | Combo, Statistical, Hierarchy, Financial, Radar & analytics overlays | Planned (next) | — | — |
| **3.4** | Contextual `Chart Design` & `Format` tabs, `+`/Brush/Funnel buttons, dialogs & docked `Format` pane | Planned | — | — |
| **3.5** | PivotCharts, OOXML XLSX chart boundary & Phase 3 completion gate | Planned | — | — |
| **Phase 4** | Remaining Long-Tail Excel Functions *(detail waves after Phase 1/3)* | Planned (detail after prior phase) | — | — |

---

## Latest Verification Ledger

- **Date**: 2026-11-28 (Sprint 3.2 — Core 2D chart engine, floating worksheet overlay & `Insert → Charts` Excel UI)
- **Excel Oracle Build Version**: Excel `16.0` build `17932` (`smoke-excel.json`, `dynamic-arrays.json`, `logical-info.json`, `lookup.json`, `shaping.json`, `math.json`, `text.json`, `datetime.json`, `stats-financial.json`, `getpivotdata.json`)
- **Commands Run**:
  - `npm run check` (`tsc --noEmit -p tsconfig.model.json && tsc --noEmit`, `vitest run`, `vite build` + `finalize-static-build.mjs`) -> exit 0 (`56/56` Vitest files: `1755` passed, `0` failed, `2` pre-existing skipped in `smoke-excel.json`; `72` files, `44.48 MiB` uncompressed / `33.65 MiB` gzip).
  - `npx playwright test tests/e2e/charts.spec.ts tests/e2e/ribbon-parity.spec.ts --project=chromium` -> exit 0 (`7/7` passed: `charts.spec.ts` `1/1`, `ribbon-parity.spec.ts` `6/6`).
- **Pass / Fail / Skipped**: `npm run check` `56/56` files (`1755` passed, `0` failed, `2` skipped); Chromium `charts.spec.ts` + `ribbon-parity.spec.ts` `7/7` passed.
- **Unrun Since Last Code Edit**: `npm run test:performance`, remaining Playwright specs (`test:e2e`, `test:worksheet`) across Chromium/Firefox/WebKit. Prior Sprint 2.5 baseline on 2026-11-27 passed `test:performance` (`6/6`), `test:worksheet` (`12/12`), and Chromium full E2E (`108/108`).

---

## Open Issues & Risks

1. **Performance budgets under dynamic arrays (`npm run test:performance`)**: Verified passing (`5/5`) at Phase 1 completion (10,000-cell interaction target and 200k-cell/10k-formula recalculation); monitor regression on future Pivot/Chart additions.
2. **XLSX `_xlfn.` & `.exlt` v1 additive compatibility**: Resolved for Phase 1 (`src/persistence/xlsx/xlsx-formulas.ts` strips `_xlfn.` / `_xlfn._xlws.` on import and re-applies on export, and native spill fields remain optional on `.exlt` v1); future Excel functions (Phase 4) must extend the `_xlfn.` tables alongside registration.
3. **Pure-ECMAScript model boundary (`tsconfig.model.json`)**: Maintained (`tsc --noEmit -p tsconfig.model.json` exited 0); `src/core` and `src/application` keep zero DOM/React/browser dependencies.
