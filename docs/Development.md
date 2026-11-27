# Development and automated verification

## Prerequisites

Node 22.12+ (Node 22 line), npm and the checked-in package lockfile. Production uses no Node server. Install with `npm ci`; start with `npm run dev`. Exact direct build/runtime versions are in `package.json`; all transitive resolutions are in `package-lock.json`. Keep both together. No credentials belong in Vite variables.

## Commands

```powershell
npm ci
npm run typecheck
npm test
npm run build
npx playwright install chromium firefox webkit
npm run test:e2e
npm run test:worksheet
```

`npm run check` combines typecheck, unit/integration tests and the production build. E2E is separate because browsers must be installed. The E2E server serves the existing `dist` directory and does not rebuild it; after source changes, run `npm run build` or `npm run check` before `npm run test:e2e`. CI runs both and installs the three engines. The Pages workflow repeats them before uploading the exact tested `dist` directory. The E2E server serves real static files at `/` and `/ExcelLightWeb/`, with no missing-asset SPA fallback, which catches relative asset-path errors hidden by a development server.

Playwright defaults to two workers so local runs and the GitHub-hosted two-core runner use the same bounded concurrency. Set `PLAYWRIGHT_WORKERS` to override this. The server port defaults to 4173; set `PLAYWRIGHT_TEST_PORT` if that port is occupied.

The production browser suite verifies hosting paths, dark rendering, worksheet input, virtualization and accessibility, plus fixture integrity and expectation loading.

## Workbook model and session

`npm run typecheck` first compiles core and application with `tsconfig.model.json` using only the ECMAScript library, then typechecks the entire project. Run the focused model suites with `npx vitest run tests/unit/workbook-model.test.ts tests/integration/workbook-session.test.ts`.

The unit suite exercises the source limits and uses deterministic generated cell edits and reorders to verify forward and inverse deltas. The integration suite covers atomic edits, bounded history, asynchronous save completions and native cache reply guards. A Node worker thread returns a fixed test reply to exercise structured cloning and edit-before-reply ordering.

## Worksheet checks

Run `npm run test:worksheet` to build and serve the fixed two-sheet worksheet fixture under the repository subpath. It checks rendered cell/header alignment at device-pixel boundaries, explicit white fills, hidden and resized tracks, cached formula display, selection statistics, sheet focus and the stored two-axis freeze fixture after scrolling and zooming. The fixture builds under `test-results/` and is separate from `dist/`.

The production suite injects the pinned axe-core development dependency to check the worksheet and ribbon. Manual screen-reader verification has not been run; see [Known Issues](Known-Issues.md).

## Editing and clipboard

`npx vitest run tests/integration/editing-commands.test.ts` covers source-derived typed entry, TSV limits, copy and cut identity, reference rewrites, merges, shared borders, Fill results and AutoSum. `npm run test:e2e` also exercises the production editor in all three engines, including clipboard event payloads, denied API fallback and stale reads. A Chromium-specific case uses the browser clipboard and CDP composition commands; cross-engine composition guard cases dispatch DOM events.

`npx vitest run tests/unit/formula-hints.test.ts` checks all supported signatures and scans incomplete drafts with quoted text, array constants and nested calls. The browser suite verifies caret movement and tooltip dismissal in both editors.

## Structure, formats and appearance

`npx vitest run tests/integration/structure-commands.test.ts` covers lane-shift scenarios and checks calculated results, stable identities, axis metadata, overflow and undo. The Playwright suite exercises clicked-sheet context commands, cell-shift dialogs, header cursor zones and direct header sizing and visibility.

Number formats:

```powershell
npx vitest run tests/unit/number-formats.test.ts tests/integration/number-format-command.test.ts tests/parity/number-formats.test.ts tests/parity/format-exchange.test.ts tests/unit/native-generation.test.ts
```

After a production build, `npm run test:e2e -- tests/e2e/number-formats.spec.ts` covers dialog drafts, undo, localized entry, cached formula display and canvas color. The captured desktop formatter fixture is checked under `en-US` with explicit region tags, and the independent 1904 fixture is checked through the web XLSX adapter.

`npx vitest run tests/integration/appearance-commands.test.ts` covers patterns and gradients, font effects, shared drafts, painter tiling, merge and formula preservation, atomic undo and workbook adoption. `npm run test:e2e -- tests/e2e/appearance.spec.ts` adds production ribbon and dialog interactions, canvas pixels, font fallback and locked and single painting across sheets.

`npx vitest run tests/integration/border-commands.test.ts tests/unit/border-geometry.test.ts` covers shared segments, mixed drafts, merged anchor mapping, outline replacement and removal, preset and stroke undo, and zoomed hit testing. `npm run test:e2e -- tests/e2e/borders.spec.ts` covers the ribbon, preview controls, all four visible sides of a replaced and cleared outline, remembered tools and drawing/erasing gestures.

## Alignment, merges, names and tables

`npm run check` includes alignment and merge transactions in `tests/integration/alignment-commands.test.ts`. `npm run test:e2e -- tests/e2e/alignment.spec.ts` covers merge discard confirmation, the combined Format Cells draft, merged keyboard navigation, canvas clipping at zoom, menu artwork and the Insert/Delete radio dialogs. `tests/e2e/dialog-placement.spec.ts` covers stacked decimal alignment, draggable dialogs and flyouts at different viewport heights.

`tests/integration/table-commands.test.ts` covers reserved and duplicate names, copied-sheet identities, header normalization, cached headers, occupied totals, explicit style overrides, whole-record deletion and resize. `npm run test:e2e -- tests/e2e/names-tables.spec.ts` covers the Create Table and New Name dialogs, Name Manager, the range picker, Table Design, contextual-tab memory, the resize grip and the total-row menu in all three engines.

## Sort, filter, search and view

`tests/integration/data-commands.test.ts` covers sort type ordering with blanks last in both directions, stable and case-sensitive collation, relative formula relocation, merged/table/Python rejections, formatted filter keys, other-column values, hidden manual rows, atomic undo, SUBTOTAL recalculation and table filter extents. `tests/e2e/data-sort-filter.spec.ts` covers the Data tab group, funnel checklist with Search and Select All, Clear and Reapply, undo, selection statistics, the Sort dialog with Options and Sort Warning, Home's Sort & Filter menu, table funnel toggles, the caret-scoped Filter toggle, filtering a selected area or the data block around a populated cell, and the empty-selection refusal.

`tests/integration/search-commands.test.ts` and `tests/unit/view-commands.test.ts` cover sheet and workbook scopes, row/column match order, formula-cell versus formatted-value searches, case and whole-cell modes, escape and blank-query handling, deduplicated 15-entry histories, stale-result refusal, one-transaction Replace All, and freeze boundaries. `tests/e2e/search-view.spec.ts` adds the Find & Select menu artwork, the option selects with the Values default, workbook results across two sheets, Find Next wrapping and reveal, Replace/Replace All undo, Zoom presets and custom clamp, hidden rows and columns at zoom, per-sheet freeze with pinned bands and the thicker divider, a narrow split viewport, and Escape dismissal.

Visual regressions are checked in the same suites: `names-tables.spec.ts` verifies Create Table checkbox alignment and filter controls clear of both frozen split lines; `search-view.spec.ts` measures at least 10 px between the Zoom custom value and its action row; `grid-corner-marker.test.ts` verifies equal triangle legs.

## Browser Python

The probe needs Node, Python 3, curl and the Playwright browsers. Set the process variable `PYTHON` if your Python executable uses another name.

```powershell
npm run prepare:python-probe
npm run test:python-probe
```

Preparation verifies cached files against the runtime lock and downloads any missing or changed file. It also stages the same assets under ignored `public/runtime/pyodide/<version>/`; `npm run build` runs this preparation before Vite bundles the production app. The production app initializes a real module worker and loads NumPy, pandas and Matplotlib on page startup. The test build serves its files and module worker under the repository subpath. Playwright checks the package handshake and reports the time to initialize a fresh worker in its runtime attachment. Add `-- --project chromium` to focus on one browser.

`npm run test:python-execution` builds and serves the production bundle in real Chromium with same-origin Pyodide. It checks exact Figure PNG dimensions, Save Image, placement over populated cells, separate placement sizes, resize/delete, deterministic same-size redraws while the worker is live, cached PNG placement/move after Stop, resize-triggered Python rerun without a separate confirmation dialog, new-size redraw, continued geometry preview during calculation, cancellation/failure retention and explicit authorized rebuild.

`npm run test:e2e -- tests/e2e/mixed-showcases.spec.ts --project=chromium` runs the Sales and SLA `.exlt` cases through their Python Figures onto Dashboard, checks the `[plot]` grid label, verifies the eight directional handle cursors and verifies redraw after resize. These are real worker/browser journeys, not protocol mocks.

## Native `.exlt` persistence

`npx vitest run tests/parity/native-package.test.ts` checks the v1 native package adapter: ZIP structure and size limits, duplicate/traversal rejection, workbook/cache validation, PNG payload integrity, web round trips of all JSON parts/styles/Python metadata, and preservation of a live session after rejected parsing. `npx vitest run tests/parity/file-adapter-probe.test.ts` exercises the bounded ZIP probe and style interchange fixtures. No Python code in a package is executed by these checks.

## Browser file lifecycle and local recovery

`npx vitest run tests/unit/local-workspace.test.ts tests/unit/native-files.test.ts` covers v1 preference migration, opt-in consent validation, background lease expiry/reclaim, file permission denial, external-file version conflicts and writer aborts. After a production build, `npm run test:e2e -- tests/e2e/files.spec.ts` runs in Chromium, Firefox and WebKit. The journeys cover Quick Access order and Save, first-Save browser naming, Recent removal without disk deletion, File backstage geometry and the single-page calculation Options dialog, the two-card Home gallery, creation of a fresh unsaved Support SLA copy, normal Save/Save As naming, absence of file/recovery messages above Recent, native package open, download Export dirty-state behavior, Save/Don't Save/Cancel, recovery after reload, two-tab ownership and removal protection, lease reclaim after expiry, IndexedDB quota failure and transactional migration success/failure.

Geometry checks pin the 200 px navigation rail, 32 px Home inset and 520 px Options dialog. IndexedDB stores package payloads as `ArrayBuffer` for WebKit compatibility while the adapter returns `Blob` packages to the native reader. Accessible worksheet cell labels live in a clipped projection; assert `aria-label` instead of visual visibility for those cells.

## XLSX exchange

`npx vitest run tests/unit/xlsx-styles.test.ts tests/integration/xlsx-import.test.ts tests/integration/xlsx-export.test.ts` covers the style/property matrix, scalar and cache types, shared strings, shared formulas, merges/dimensions/freeze, unsupported formula freshness, omissions, 1904 offsets, bounds/macro/encryption/DTD rejection, reviewed export reports, no invented zero and dirty-state preservation. `npm run test:e2e -- tests/e2e/xlsx-files.spec.ts` opens the sales fixture through File → Open with the download fallback, cancels and continues the import report, verifies sheet names and imported cells, edits a cell, then runs the reviewed Export and checks the downloaded name and retained dirty state.

## Offline packaging and hardening

`npm run check` verifies the runtime lock, extracts wheel license texts into the staged runtime, builds `dist/offline-manifest.json` and the base-scoped `dist/service-worker.js`, and reports full artifact sizes. `tests/unit/offline-service-worker.test.ts` exercises checksum rejection, interrupted-cache retry, quota failure and dirty-tab update deferral in a worker harness; these supplement the real browser checks.

After the production build, `npm run test:e2e -- tests/e2e/offline.spec.ts` runs in Chromium, Firefox and WebKit. All three check root and repository-subpath service-worker scope, CSP and same-origin requests; Chromium and Firefox also reopen the shell offline and check Monaco and escaped worksheet text. Playwright WebKit cannot navigate a service-worker response after `context.setOffline(true)` ([upstream issue #42775](https://github.com/microsoft/playwright/issues/42775)), so its offline-navigation case is skipped while its online scope checks still run. The Chromium-only runtime case caches the full locked package set, disconnects the browser, reloads the app and runs Python in the actual Pyodide WASM worker from Cache Storage.

## Test layers

| Layer | What it proves | Examples |
| --- | --- | --- |
| Unit (Vitest) | Pure model and semantic contracts | typed blanks, reference transforms, formats, formula functions, inverse commands |
| Integration | Boundaries and failure transactions | ZIP bounds, failed IndexedDB transaction, worker stale replies, saved identity race |
| Parity | Agreement with captured desktop behavior | formula scalar/error tables, native package round trips, XLSX style/loss matrix |
| Browser (Playwright) | Real interaction and browser APIs | IME/clipboard fallback, selection/focus, File API, actual Pyodide WASM, chart drag/redraw |
| Performance | Measured workload on recorded hardware/browser | model bounds, scrolling, cold/warm Python, mixed generation latency |

The fixture manifest records source path, SHA-256, expected-value source, permitted use and capture date. Use the existing expected-values JSON and captured assertions; never generate the expected result with the implementation under test. Tests cover cancellation and error paths as well as happy paths. Reference transformations include generated inputs; native archive readers get corrupt and decompression-limit fixtures. Deterministic fake clocks cover TODAY/debounce, and real workers cover concurrency races. Real Pyodide tests run in browser workers against packaged same-origin assets; protocol mocks are supplementary. Figure checks compare dimensions and structural image properties, not byte-exact PNG equality across Matplotlib versions.

### Excel formula oracle

On Windows with desktop Excel installed, generate differential oracle fixtures under `tests/fixtures/formulas/` from category spec files (`tests/fixtures/formulas/specs/*.spec.json`):

```powershell
powershell -ExecutionPolicy Bypass -File scripts/generate-excel-oracle.ps1 -SpecPath tests/fixtures/formulas/specs/smoke.spec.json -OutPath tests/fixtures/formulas/smoke-excel.json
```

If COM automation is unavailable, save an `.xlsx` workbook with the spec expressions in column `T` (`T1`, `T2`, ...) from Excel so cached `<v>` values are present, then extract the fixture JSON with:

```powershell
node scripts/extract-xlsx-oracle.mjs --xlsx path/to/evaluated.xlsx --spec tests/fixtures/formulas/specs/smoke.spec.json --out tests/fixtures/formulas/smoke-excel.json
```

Vitest (`tests/parity/native-formulas.test.ts`) runs the checked-in oracle JSON files locally and in CI without requiring Excel.

## Performance

Record OS/CPU/RAM/browser/runtime, fixture dimensions, warm/cold state and sample count. Report p50/p95 over at least 20 warm interactions.

- Typing/selection p95 under 100 ms on a 10,000-populated-cell fixture; ordinary input avoids main-thread tasks above 50 ms.
- Scrolling targets 60 Hz for the visible window of the `A1:XFD1048576` grid; report dropped frames and viewport/zoom rather than conflating model limits with rendered cells. The spacer's CSS length stays under `16,777,216` px, header included. Ctrl+Down on an empty column lands on the last row inside that window, and the row-number column is wide enough for `1048576`.
- Warm native recalculation target under 500 ms for a fixed 1,000-formula dependency fixture; stress 10,000 formulas and 200,000 populated cells without data loss or tab termination, recording observed time/memory.
- Python default timeout remains 30 seconds. Cold runtime load and showcase cold/warm runs are measured separately.
- First shell JS/CSS gzip target below 500 KiB excluding lazy Monaco/Pyodide. Runtime asset sizes are reported and must stay below actual Pages constraints.

`npm run test:performance` builds a deterministic 2-sheet, 200,000-populated-cell / 10,000-formula workbook and a 10,000-cell fixture, records the environment and writes p50/p95 samples to ignored `test-results/performance/`. It asserts no data loss at the exact limits, the 10,000-cell session-edit p95 under 100 ms, and the fixed 1,000-formula warm recalculation p95 under 500 ms. The 200,000-cell measurements are recorded rather than asserted against an interaction budget; a single-cell edit there measured p95 669 ms and is a known main-thread cost.

Browser input and scrolling are measured after a production build with `npm run test:e2e -- tests/e2e/performance.spec.ts`. Key events are dispatched inside the page so Playwright protocol overhead is excluded; typing and selection p95 are asserted under 100 ms on the 10,000-cell fixture, Chromium long tasks must stay under 50 ms, and scroll frame times, dropped-at-60-Hz counts and idle cadence are reported. The scroll check samples 15 idle and 60 active frames; each animation-frame wait has a 750 ms fallback and active sampling stops on a frame over 500 ms. Chromium and Firefox retain a 60 ms p95 ceiling. WebKit uses a 350 ms p95 ceiling and the same hard `<500 ms` maximum, keeping a margin over the reported CI p95 of 301 ms while still rejecting half-second stalls. Recorded values: Chromium 16.8 ms, Firefox 25 ms; two Windows WebKit reruns on 2026-11-26 measured p95 116 ms and 147 ms with no stalls; the Linux CI matrix passed the same assertion without a retry. Each run prints its frame numbers to the test log. Software-rendered WebKit at 2x device pixels remains an environment limitation, not Safari parity evidence.

The Python timing case runs with the probe harness:

```powershell
$env:PLAYWRIGHT_TEST_PORT='4211'
npm run test:python-probe -- --project=chromium -g "cold runtime startup"
Remove-Item Env:PLAYWRIGHT_TEST_PORT
```

It records cold worker start plus 20 warm unchanged sales-script runs under the pinned runtime in ignored `test-results/performance/python-timings.json`.

## Workflow

Update the affected [Features](Features.md) rows with test paths in the same change. Record pass/fail/skipped/unrun counts where available, not just "tested". A missing real-runtime or interop gate is a blocker for that capability claim. Fix introduced regressions before advancing. Do not add redundant implementation-mirroring tests merely to increase counts.

If npm encounters a corporate certificate chain, use an approved CA configuration or Node's supported system CA option for the process. Never disable TLS verification or commit machine-specific certificates/settings. Browser downloads may similarly need the approved trust configuration.
