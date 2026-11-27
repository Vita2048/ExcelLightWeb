# Expansion Plan Status

Live status tracker for [ExpansionPlan.md](ExpansionPlan.md). Update after every sprint with exact verification counts and current risks.

---

## Sprint Matrix

| Sprint | Scope | Status | Verification | Date |
| --- | --- | --- | --- | --- |
| **1.1** | Excel COM oracle harness, `src/calculation/functions/` modularization, `#SPILL!`/`#CALC!` errors | Done | `npm run check` (45/45 files: 1177 passed, 5 skipped, 0 failed) + Excel 16.0.17932 smoke oracle (15 captured, 10 run & passed, 5 skipped pending Sprints 1.3–1.4) | 2026-11-27 |
| **1.2** | Parser & reference upgrade (`#`, `@`, structured refs, 3D refs, higher-order calls) | Not started | — | — |
| **1.3** | Dynamic array spilling, broadcasting, `LET`/`LAMBDA`/helpers, volatile/dynamic refs | Not started | — | — |
| **1.4** | Core Function Wave A (Lookup, Dynamic Arrays, Logical, Info ~50 functions) | Not started | — | — |
| **1.5** | Core Function Wave B (Math, Aggregates, Text, Date/Time, Core Stats/Financial ~85 functions) | Not started | — | — |
| **Phase 2** | PivotTables *(detail sprints after Phase 1)* | Planned (detail after prior phase) | — | — |
| **Phase 3** | Charts — Excel variety *(detail sprints after Phase 2)* | Planned (detail after prior phase) | — | — |
| **Phase 4** | Remaining Long-Tail Excel Functions *(detail waves after Phase 1/3)* | Planned (detail after prior phase) | — | — |

---

## Latest Verification Ledger

- **Date**: 2026-11-27 (Sprint 1.1)
- **Excel Oracle Build Version**: Excel `16.0` build `17932` (`tests/fixtures/formulas/smoke-excel.json`)
- **Commands Run**:
  - `powershell -ExecutionPolicy Bypass -File scripts/generate-excel-oracle.ps1 -SpecPath tests/fixtures/formulas/specs/smoke.spec.json -OutPath tests/fixtures/formulas/smoke-excel.json` -> exit 0, captured 15 cases (scalar, error, and 2x2 `SEQUENCE` array).
  - `npm run check` (`tsc --noEmit -p tsconfig.model.json && tsc --noEmit`, `vitest run`, `vite build` + `finalize-static-build.mjs`) -> exit 0.
- **Pass / Fail / Skipped**: `45/45` Vitest files passed (`1177` passed, `0` failed, `5` skipped in `tests/parity/native-formulas.test.ts` for `smoke-excel.json` `FILTER`/`SEQUENCE`/array cases pending Sprints 1.3–1.4).
- **Unrun**: `scripts/extract-xlsx-oracle.mjs` was not executed against a manual `.xlsx` file (Excel COM succeeded directly); `npm run test:e2e` and `npm run test:worksheet` were not run (no UI changes in Sprint 1.1).

---

## Open Issues & Risks

1. **Performance budgets under dynamic arrays (`npm run test:performance`)**: `native-generation.ts` must stay under 500 ms p95 for warm 1,000-formula recalculation and 100 ms p95 for 10,000-cell edits. Broadcasting and spill-graph updates must stay sparse and avoid $O(N^2)$ array copying.
2. **XLSX `_xlfn.` & `.exlt` v1 additive compatibility**: Modern functions (`XLOOKUP`, `FILTER`, `LET`, `LAMBDA`, `TEXTSPLIT`, etc.) require `_xlfn.` prefix handling on XLSX import/export, and native spill fields must remain optional on `.exlt` v1 so existing fixtures round-trip cleanly.
3. **Pure-ECMAScript model boundary (`tsconfig.model.json`)**: `src/core` and `src/application` compile with ECMAScript types only (no DOM/React/browser APIs); all new AST, structured-ref, and shared spill-ownership types must preserve this boundary.
