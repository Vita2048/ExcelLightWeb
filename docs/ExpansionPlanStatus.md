# Expansion Plan Status

Live status tracker for [ExpansionPlan.md](ExpansionPlan.md). Update after every sprint with exact verification counts and current risks.

---

## Sprint Matrix

| Sprint | Scope | Status | Verification | Date |
| --- | --- | --- | --- | --- |
| **1.1** | Excel COM oracle harness, `src/calculation/functions/` modularization, `#SPILL!`/`#CALC!` errors | Done | `npm run check` (45/45 files: 1177 passed, 5 skipped, 0 failed) + Excel 16.0.17932 smoke oracle (15 captured, 10 run & passed, 5 skipped pending Sprints 1.3–1.4) | 2026-11-27 |
| **1.2** | Parser & reference upgrade (`#`, `@`, structured refs, 3D refs, higher-order calls) | Done | `npm run check` (46/46 files: 1184 passed, 5 skipped, 0 failed; `vite build` 72 files, 43.99 MiB) | 2026-11-27 |
| **1.3** | Dynamic array spilling, broadcasting, `LET`/`LAMBDA`/helpers, volatile/dynamic refs | Done | `npm run check` (47/47 files: 1221 passed, 5 skipped, 0 failed) + Excel 16.0.17932 dynamic-arrays oracle (34/34 passed) + Chromium e2e (5/5 passed) | 2026-11-27 |
| **1.4** | Core Function Wave A (Lookup, Dynamic Arrays, Logical, Info ~50 functions) | Not started | — | — |
| **1.5** | Core Function Wave B (Math, Aggregates, Text, Date/Time, Core Stats/Financial ~85 functions) | Not started | — | — |
| **Phase 2** | PivotTables *(detail sprints after Phase 1)* | Planned (detail after prior phase) | — | — |
| **Phase 3** | Charts — Excel variety *(detail sprints after Phase 2)* | Planned (detail after prior phase) | — | — |
| **Phase 4** | Remaining Long-Tail Excel Functions *(detail waves after Phase 1/3)* | Planned (detail after prior phase) | — | — |

---

## Latest Verification Ledger

- **Date**: 2026-11-27 (Sprint 1.3)
- **Excel Oracle Build Version**: Excel `16.0` build `17932` (`tests/fixtures/formulas/smoke-excel.json`, `tests/fixtures/formulas/dynamic-arrays.json`)
- **Commands Run**:
  - `powershell -ExecutionPolicy Bypass -File scripts/generate-excel-oracle.ps1 -SpecPath tests/fixtures/formulas/specs/dynamic-arrays.spec.json -OutPath tests/fixtures/formulas/dynamic-arrays.json` -> exit 0 (`34` cases captured).
  - `npm run check` (`tsc --noEmit -p tsconfig.model.json && tsc --noEmit`, `vitest run`, `vite build` + `finalize-static-build.mjs`) -> exit 0 (`72` files, `44.01 MiB` uncompressed / `33.54 MiB` gzip).
  - `npm run test:e2e -- tests/e2e/dynamic-arrays.spec.ts tests/e2e/ribbon-parity.spec.ts --project=chromium` -> exit 0.
- **Pass / Fail / Skipped**: Vitest `47/47` files passed (`1221` passed, `0` failed, `5` skipped in `smoke-excel.json` for Sprint 1.4 `FILTER`/`SEQUENCE` cases; `34/34` `dynamic-arrays.json` oracle cases passed); Playwright Chromium `5/5` passed.
- **Unrun**: Full cross-browser e2e matrix (`firefox`, `webkit`), `npm run test:worksheet`, and `npm run test:performance` were not run in Sprint 1.3 (scheduled for Phase 1 completion in Sprint 1.5).

---

## Open Issues & Risks

1. **Performance budgets under dynamic arrays (`npm run test:performance`)**: `native-generation.ts` must stay under 500 ms p95 for warm 1,000-formula recalculation and 100 ms p95 for 10,000-cell edits. Broadcasting and spill-graph updates must stay sparse and avoid $O(N^2)$ array copying.
2. **XLSX `_xlfn.` & `.exlt` v1 additive compatibility**: Modern functions (`XLOOKUP`, `FILTER`, `LET`, `LAMBDA`, `TEXTSPLIT`, etc.) require `_xlfn.` prefix handling on XLSX import/export, and native spill fields must remain optional on `.exlt` v1 so existing fixtures round-trip cleanly.
3. **Pure-ECMAScript model boundary (`tsconfig.model.json`)**: `src/core` and `src/application` compile with ECMAScript types only (no DOM/React/browser APIs); all new AST, structured-ref, and shared spill-ownership types must preserve this boundary.
