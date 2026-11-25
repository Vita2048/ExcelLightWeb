# Development and automated verification

## Scaffold prerequisites

Node 22.12+ (Node 22 line), npm and the checked-in package lockfile. Production uses no Node server. Install with `npm ci`; start with `npm run dev`. Exact direct build/runtime versions are in package.json; all transitive resolutions are in package-lock.json. Keep both together. No credentials belong in Vite variables.

## Commands

```powershell
npm ci
npm run typecheck
npm test
npm run build
npx playwright install chromium firefox webkit
npm run test:e2e
```

`npm run check` combines typecheck, unit/integration tests and production build. E2E is separate because browsers must be installed. The E2E server serves the existing `dist` directory and does not rebuild it; after source changes, run `npm run build` or `npm run check` before `npm run test:e2e`. CI runs both and installs the three browser engines. The Pages workflow repeats them before uploading the exact tested dist directory. The E2E server serves real static files at `/` and `/ExcelLightWeb/`, with no missing-asset SPA fallback. This catches relative asset-path errors hidden by a development server.

Playwright defaults to two workers so local runs and the GitHub-hosted two-core runner use the same bounded concurrency. Set `PLAYWRIGHT_WORKERS` to override this for a machine with different capacity. The server port defaults to 4173; set `PLAYWRIGHT_TEST_PORT` if that port is occupied.

The production browser suite verifies hosting paths, dark rendering, worksheet input, virtualization and accessibility. S01 adds fixture integrity, snapshot and expectation loading checks. Vite's manual preview server is for local development.

## Workbook model and session

`npm run typecheck` first compiles core and application with `tsconfig.model.json` using only the ECMAScript library. It then typechecks the entire project. Run the focused S04 suites with `npx vitest run tests/unit/workbook-model.test.ts tests/integration/workbook-session.test.ts`.

The unit suite exercises the source limits and uses deterministic generated cell edits and reorders to verify forward and inverse deltas. The integration suite covers atomic edits, bounded history, asynchronous save completions and native cache reply guards. A Node worker thread returns a fixed test reply to exercise structured cloning and edit-before-reply ordering.

## Worksheet checks

Run `npm run test:worksheet` to build and serve the fixed two-sheet worksheet fixture under the repository subpath. It checks rendered cell/header alignment at device-pixel boundaries, explicit white fills, hidden and resized tracks, cached formula display, selection statistics and sheet focus. Verify CI runs it after the production browser suite. The fixture builds under `test-results/` and is separate from `dist/`.

The production suite injects the pinned axe-core development dependency to check the worksheet and ribbon. Manual screen-reader verification is scheduled in S25.

## Editing and clipboard checks

`npx vitest run tests/integration/editing-commands.test.ts` covers source-derived typed entry, TSV limits, copy and cut identity, reference rewrites, merges, shared borders, Fill results and AutoSum. `npm run test:e2e` also exercises the production editor in all three engines, including clipboard event payloads, denied API fallback and stale reads. A Chromium-specific case uses the browser clipboard and CDP composition commands. The cross-engine composition guard cases dispatch DOM events.

`npx vitest run tests/unit/formula-hints.test.ts` checks all supported signatures and scans incomplete drafts with quoted text, array constants and nested calls. The production browser suite verifies caret movement and tooltip dismissal in the formula bar and inline editor.

## Structural edit checks

`npx vitest run tests/integration/structure-commands.test.ts` covers the four desktop lane-shift scenarios and checks calculated results, stable identities, axis metadata, overflow and undo. The production Playwright suite exercises clicked-sheet context commands, cell-shift dialogs, header cursor zones and direct header sizing and visibility.

## Number format checks

Run the S10 model and fixture checks with the following command.

```powershell
npx vitest run tests/unit/number-formats.test.ts tests/integration/number-format-command.test.ts tests/parity/number-formats.test.ts tests/parity/format-exchange.test.ts tests/unit/native-generation.test.ts
```

After a production build, run `npm run test:e2e -- tests/e2e/number-formats.spec.ts` for dialog drafts, undo, localized entry, cached formula display and canvas color across the installed engines.

The desktop format oracle uses existing Release binaries and .NET 10. It checks the captured formatter results under `en-US` with explicit tags for each tested region, and verifies the independent 1904 fixture through the desktop XLSX importer. Pass `-p:DesktopBin=<absolute path>` if the binaries are in another directory.

```powershell
dotnet run --project tests/desktop-formats/Oracle.csproj -- tests/fixtures/formats/S10-desktop.json
dotnet run --project tests/desktop-formats/Oracle.csproj -- --check-1904 tests/fixtures/formats/Desktop1904.xlsx
```

The oracle's `--locales` mode writes the Gregorian locale table with the format patterns and standard separators. Use a scratch file to compare a new capture with `src/core/formats/locales.json`. The production bundle includes this data locally.

## Appearance and Format Painter checks

Run `npx vitest run tests/integration/appearance-commands.test.ts` for all modeled patterns and gradients, font effects, shared drafts, painter tiling, merge and formula preservation, atomic undo and workbook adoption. After building, run `npm run test:e2e -- tests/e2e/appearance.spec.ts` for production ribbon and dialog interactions, canvas pixels, font fallback and locked and single painting across sheets. The dialog cases also cover desktop tab order, narrow panes, font catalog selection, the inline palette, the visual pattern picker and Fill Effects radios.

Run the existing desktop Font/Fill and Format Painter source assertions against the desktop Release assemblies with the following command. The harness links the read-only source test files. Pass `-p:DesktopSource=<absolute path>` and, when needed, `-p:DesktopBin=<absolute path>` for a different checkout.

```powershell
dotnet run --project tests/desktop-appearance/Oracle.csproj
```

## Border and drawing tool checks

Run `npx vitest run tests/integration/border-commands.test.ts tests/unit/border-geometry.test.ts` for shared segments, mixed drafts, merged anchor mapping, dialog outline replacement and removal, preset and stroke undo, and zoomed hit testing. After building, run `npm run test:e2e -- tests/e2e/borders.spec.ts` for the production ribbon, dialog preview controls, all four visible sides of a replaced and cleared outline, remembered tools, and drawing and erasing gestures. `npm run test:worksheet` includes a stored two-axis freeze fixture that checks canvas pixels after scrolling and zooming.

Run the desktop source assertions against the local Release assemblies with `dotnet run --project tests/desktop-borders/Oracle.csproj`. Pass `-p:DesktopSource=<absolute path>` and, when needed, `-p:DesktopBin=<absolute path>` for another desktop checkout.

## Alignment, merges and ribbon checks

`npm run check` includes [alignment and merge transactions](../tests/integration/alignment-commands.test.ts). `npm run test:e2e -- tests/e2e/alignment.spec.ts` covers merge discard confirmation, the combined Format Cells draft, merged keyboard navigation and canvas clipping at zoom, menu artwork and keyboard operation, and the Insert and Delete radio dialogs. The full production suite exercises the updated styling and controls for Clipboard, Font, Borders, Number, Cells, Editing, Formulas and View. The scaffold tests request the favicon from both deployment paths. Run the relevant scenarios in [alignment](../tests/e2e/alignment.spec.ts) and [dialog placement](../tests/e2e/dialog-placement.spec.ts) for focused checks of stacked decimal alignment, draggable dialogs and flyouts at different viewport heights.

`dotnet run --project tests/desktop-alignment/Oracle.csproj` executes the desktop alignment and merge assertions against the existing Release assemblies. Pass `-p:DesktopSource=<absolute path>` and, when needed, `-p:DesktopBin=<absolute path>` for another checkout.

## Names and tables

`npm run check` includes [name and table transactions](../tests/integration/table-commands.test.ts): reserved and duplicate names, copied-sheet identities, header normalization, cached headers, occupied totals, explicit style overrides, whole-record deletion and resize. After a production build, run `npm run test:e2e -- tests/e2e/names-tables.spec.ts` for the Create Table and New Name dialogs, Name Manager, the range picker, Table Design, contextual-tab memory, the resize grip and the total-row menu in Chromium, Firefox and WebKit. The keyboard ribbon check in [worksheet](../tests/e2e/worksheet.spec.ts) expects the Insert Table button. No separate desktop oracle project was added for this sprint. S15 adds the Data tab funnels and checklist that operate on that owned filter.

## Sort and filter

`npm run check` includes [data transactions](../tests/integration/data-commands.test.ts): sort type ordering with blanks last in both directions, stable and case-sensitive collation, relative formula relocation, merged, table and Python rejections, formatted filter keys, other-column values, hidden manual rows, atomic undo of sort and filter edits, SUBTOTAL recalculation and table filter extents. After a production build, run `npm run test:e2e -- tests/e2e/data-sort-filter.spec.ts` for the Data tab group, funnel checklist with Search and Select All, Clear and Reapply, undo, selection statistics and SUBTOTAL under a filter, the Sort dialog with Options and the Sort Warning, Home's Sort & Filter menu, table funnel toggles, the caret-scoped Filter toggle, filtering a selected area or the data block around a populated cell outside a table, and the empty-selection refusal in Chromium, Firefox and WebKit. No separate desktop oracle project was added for this sprint; behavior was checked against `DataCommands.cs`, `DataRibbon.cs` and the desktop S12/S18 checks.

## Find, zoom and freeze

`npm run check` includes [search transactions](../tests/integration/search-commands.test.ts) and [freeze commands](../tests/unit/view-commands.test.ts): sheet and workbook scopes, row or column match order, formula-cell versus formatted-value searches (and the Replace tab's stored-source breadth), case and whole-cell modes, escape and blank-query handling, exact-deduplicated 15-entry histories, stale-result refusal, one-transaction Replace All, and the freeze boundary and status contract. After a production build, run `npm run test:e2e -- tests/e2e/search-view.spec.ts` for the Home → Editing → Find & Select menu with its desktop artwork, the Find panes and their option selects with the Values default and SearchDialog option rows, workbook results across two sheets, Find Next wrapping and reveal, Replace and Replace All undo, the Zoom dialog presets and custom clamp, hidden rows and columns at zoom, per-sheet freeze and unfreeze with pinned bands and their thicker edge divider, or a narrow split viewport, and Escape dismissal in Chromium, Firefox and WebKit. No separate desktop oracle project was added; behavior was checked against `SearchDialog.cs`, `EverydayCommands.Find/Replace`, `DataRibbon.cs` and `FreezeChecks.cs`.

S26 visual regressions are checked in the same production browser suites: `names-tables.spec.ts` verifies the Create Table checkbox alignment and filter controls clear of both frozen split lines; `search-view.spec.ts` measures at least 10px between the Zoom custom value and its action row, following the desktop `DialogLayout` footer inset; `grid-corner-marker.test.ts` verifies equal triangle legs.

## Browser Python feasibility

Run the S02 probe with Node, Python 3, curl and the Playwright browsers listed above. Set the process variable `PYTHON` if your Python executable uses another name.

```powershell
npm run prepare:python-probe
npm run test:python-probe
```

Preparation verifies the cached files against the runtime lock and downloads any missing or changed file. It also stages the same assets under ignored `public/runtime/pyodide/<version>/`; `npm run build` runs this preparation before Vite bundles the production app. The test build serves its files and module worker under the repository subpath. Playwright checks the package handshake and reports the time to initialize a fresh worker in its runtime attachment. Add `-- --project chromium` to focus on one browser. Python feasibility still runs against the dedicated probe build in `test-results/python-probe-build/`.

## Offline packaging and hardening

`npm run check` verifies the runtime lock, extracts wheel license texts into the staged runtime, builds `dist/offline-manifest.json` and the base-scoped `dist/service-worker.js`, and reports full artifact sizes. [The offline service-worker contract tests](../tests/unit/offline-service-worker.test.ts) exercise checksum rejection, interrupted-cache retry, quota failure and dirty-tab update deferral in a worker harness; they supplement, not replace, the real browser checks.

After the production build, run `npm run test:e2e -- tests/e2e/offline.spec.ts` in Chromium, Firefox and WebKit. All three check root and repository-subpath service-worker scope, CSP and same-origin requests; Chromium and Firefox also reopen the shell offline and check Monaco and escaped worksheet text. Playwright WebKit cannot navigate a service-worker response after `context.setOffline(true)` ([upstream issue #42775](https://github.com/microsoft/playwright/issues/42775)), so its offline-navigation case is skipped while its online scope checks still run. The Chromium-only runtime case caches the full locked package set, disconnects the browser, reloads the app, and runs Python in the actual Pyodide WASM worker from Cache Storage. Run the complete `npm run test:e2e` matrix as the sprint gate. [Package notices and generation](dependencies.md)

## Formula and file feasibility

Run `npx vitest run tests/unit/native-generation.test.ts tests/unit/formula-presentation.test.ts` for dirty-node pruning, scoped recalculation, volatile clock use, cycles and locale presentation. Run `npm run test:e2e -- tests/e2e/formulas.spec.ts` after building to exercise the production module worker, stale replies, Manual labels and formula commands in three engines. The 1,000-formula test holds delivery of actual worker replies so it can verify that worksheet input remains available and an older revision never publishes its cache batch. It attaches cold and warm worker turnaround samples for diagnosis. Those single-run samples are not a percentile benchmark.

Run `npx vitest run tests/unit/reference-pick.test.ts` for operand and replacement detection. After building, run `npm run test:e2e -- tests/e2e/reference-pick.spec.ts` to verify mouse clicks, drag direction, reference replacement, sheet changes and commit or cancel at the original cell.

## Python authoring and references

Run `npx vitest run tests/unit/python-references.test.ts tests/unit/python-xl-pick.test.ts tests/unit/python-editor-model.test.ts tests/unit/grid-selection.test.ts tests/integration/python-authoring.test.ts tests/integration/python-execution.test.ts` for literal lexing/rewrites, caret slot insertion, inert `=PY` creation, cached-value versus Show Formulas display, stable source commits, input/header/date validation, result limits, spill ownership, native package round trips and Insert as Values. After building, run `npm run test:e2e -- tests/e2e/python-authoring.spec.ts` in Chromium, Firefox and WebKit for Monaco drafts, the dedicated Python ribbon group and glyph, Python cell display, cross-sheet picks, local undo, barriers and Save/reopen. Run `npm run test:python-execution` for the production build in a real Chromium browser with the actual same-origin Pyodide worker; it verifies consent, no execution on open/reopen, scalar/Series/DataFrame/array/Figure results, non-finite errors, input freshness, spill collision/shrink, Stop, timeout/restart and values insertion. The default three-engine suite skips this long runtime journey in Firefox and WebKit; Python authoring still runs in all three engines. Set `PLAYWRIGHT_PYTHON_PORT` when its default port 4175 is occupied. This is a real runtime check, not a protocol mock.

Run the S08 differential suite with `npx vitest run tests/parity/native-formulas.test.ts`. Its synthetic workbook fixtures were evaluated using the desktop Release `NativeCalculator` under `cs-CZ`. The fixture manifest records the assembly hash and runtime. To verify the stored results with the same desktop binaries, run the following from the web repository with .NET 10. Pass `-p:DesktopBin=<absolute path>` for another Release binary directory.

## Mixed native/Python calculation

Run `npx vitest run tests/integration/mixed-calculation.test.ts` for native prerequisites, Python-to-native chains, spill-created edges, cycle blocking, upstream failure and a single atomic history generation. After a production build, run `npm run test:e2e -- tests/e2e/mixed-calculation.spec.ts tests/e2e/mixed-showcases.spec.ts --project=chromium` with the real same-origin Pyodide worker. The first journey covers Manual Run, local Automatic consent, edit coalescing, stale-result handling and Undo/Redo. The Sales and SLA journeys open the checked-in `.exlt` fixtures, confirm import and edits do not start Python in Manual mode, then explicitly Run the edited anchor and check the spill and dependent native KPI. No fixture code runs just because it was imported. The Pyodide runtime version and package hashes are in [runtime-lock.json](../src/python/runtime-lock.json).

## Calculation progress feedback

Explicit Python runs report runtime startup and named NumPy, pandas and Matplotlib loading stages in the worksheet status bar. The status then identifies the current Python anchor and dependent native formulas. A small canvas indicator marks each calculating cell, and the accessible gridcell exposes `aria-description="Calculating"`. Native worker progress and final results stay within the worker protocol; UI state does not enter the calculation model. The loader uses package-stage messages rather than a percentage because the bundled Pyodide package API exposes promise completion, not a determinate per-package byte-progress callback ([Pyodide loading packages](https://pyodide.org/en/stable/usage/loading-packages.html)). The real-worker SLA journey asserts the package stages and Python-cell state; the held native-worker reply asserts the formula-cell state.

## Figure placements and redraw

Run `npm test -- --run tests/unit/figure-geometry.test.ts tests/integration/figure-commands.test.ts` for zoomed/hidden-track geometry, bounded render sizes, multiple placement dimensions, stale-basis rejection, history, native Save/reopen and exclusion of worker-only Figure identity. `npm run test:python-execution` builds/serves the production bundle in real Chromium with same-origin Pyodide; it checks exact Figure PNG dimensions, Save Image, placement over populated cells, separate placement sizes, resize/delete, deterministic same-size redraws while the worker is live, cached PNG placement/move after Stop, resize-triggered Python rerun without a separate confirmation dialog, new-size redraw after resize, continued geometry preview during calculation, cancellation/failure retention and explicit authorized rebuild. `npm run test:e2e -- tests/e2e/mixed-showcases.spec.ts --project=chromium` runs the Sales and SLA `.exlt` cases through their Python Figures onto Dashboard, checks the `[plot]` grid label, verifies the eight directional handle cursors, and verifies redraw after resize. Run `npx vitest run tests/unit/grid-selection.test.ts` for the Python cached scalar, `[plot]`, and Show Formulas display cases. These are real worker/browser journeys, not protocol mocks. The entire production browser matrix remains `npm run test:e2e` across Chromium, Firefox and WebKit.

```powershell
dotnet run --project tests/desktop-formulas/Oracle.csproj -- tests/fixtures/formulas/S08-desktop.json
```

`npm test` includes the S03 parity cases. `npm run test:adapter-probe` bundles the probes separately and tests the desktop fixtures in all three Playwright engines. For independent desktop verification, use .NET 10 and the existing Release desktop test binaries. Run the following from the web repository. Pass `-p:DesktopBin=<absolute path>` to the .NET project if the desktop build is in another location.

```powershell
$env:S03_INTERCHANGE_OUT='node_modules/.cache/s03-verified'
npx vitest run tests/parity/file-adapter-probe.test.ts
dotnet run --project tests/desktop-interchange/Interchange.csproj -- check node_modules/.cache/s03-verified
Remove-Item Env:S03_INTERCHANGE_OUT
```

Use `export` in place of `check` with a scratch directory to generate the original samples. The checked-in files are enough for Node and browser verification on other machines. [Adapter decision and budgets](decisions/ADR-003-formula-and-file-adapters.md)

## Native .exlt persistence

Run `npx vitest run tests/parity/native-package.test.ts` for the v1 native package adapter. It checks ZIP structure and size limits, duplicate/traversal rejection, workbook/cache validation, PNG payload integrity, and preservation of a live session after rejected parsing. When desktop Release assemblies are available, it also verifies web-save-desktop equality for DesktopStyles, Showcase and Support-SLA. Set `DESKTOP_BIN` to the absolute directory containing `ExcelLite.Core.dll` and `ExcelLite.Persistence.dll`; Windows defaults to `C:/Temp/ExcelLight/tests/ExcelLite.Tests/bin/Release/net10.0`. The two C# assertions are explicitly skipped when these external assemblies are absent, as in GitHub's Linux job; a skip is not desktop parity evidence. `tests/desktop-interchange/Interchange.csproj` targets .NET 10. No Python code in a package is executed by these checks.

## Browser file lifecycle and local recovery

Run `npx vitest run tests/unit/local-workspace.test.ts tests/unit/native-files.test.ts` for v1 preference migration, opt-in consent validation, background lease expiry/reclaim, file permission denial, external-file version conflicts and writer aborts. After a production build, run `npm run test:e2e -- tests/e2e/files.spec.ts` in Chromium, Firefox and WebKit. The real browser journeys cover the C# Quick Access order (Save, Undo, Redo) and Save action, first-Save browser naming, Recent removal without disk deletion, desktop File backstage geometry and the single-page calculation Options dialog, the two-card Home gallery, creation of a fresh unsaved Support SLA template copy, normal Save/Save As naming for that copy, absence of file/recovery messages above Recent with feedback retained on the worksheet status bar, local Save and Recent, native package open, download Export dirty-state behavior, Save/Don't Save/Cancel, recovery after reload, two-tab ownership including removal protection, single-tab lease reclaim after expiry, IndexedDB quota failure and transactional migration success/failure. Geometry checks pin the desktop 200px navigation rail, 32px Home inset and 520px Options dialog; the locked dark theme has no General Options category. IndexedDB stores package payloads as `ArrayBuffer` for WebKit compatibility while the adapter returns `Blob` packages to the native reader. Accessible worksheet cell labels live in a clipped projection; assert `aria-label` instead of visual visibility for those cells. The test server defaults to port 4173; set `PLAYWRIGHT_TEST_PORT` to use another local port when that port is occupied.

## XLSX exchange

Run `npx vitest run tests/unit/xlsx-styles.test.ts tests/integration/xlsx-import.test.ts tests/integration/xlsx-export.test.ts` for the style/property matrix, scalar and cache types, shared strings, shared formulas, merges/dimensions/freeze, unsupported formula freshness, omissions, 1904 offsets, bounds/macro/encryption/DTD rejection, reviewed export reports, no invented zero and dirty-state preservation. After a production build, run `npm run test:e2e -- tests/e2e/xlsx-files.spec.ts`; the real Chromium, Firefox and WebKit journeys open the sales fixture through File → Open with the download fallback, cancel and continue the import report, verify sheet names and imported cells, edit a cell, then run the reviewed Export and check the downloaded name and retained dirty state.

The desktop XLSX oracle reimports web-exported files with the Release ClosedXML adapter. Generate exports with `S19_XLSX_OUT`, run Vitest, then the .NET project; set `DesktopBin` when the Release binaries are elsewhere, as with the other desktop oracles. It checks sheet names, a formula source and cache, a number format, date serial 60, alignment, gradients, borders and freeze. It does not check row/column dimensions: the desktop's current ClosedXML reader returns the default width for the width attribute its own exporter writes, while the TypeScript import/export roundtrip covers dimensions.

```powershell
$env:S19_XLSX_OUT='node_modules/.cache/s19-xlsx'
npx vitest run tests/integration/xlsx-export.test.ts
dotnet run --project tests/desktop-xlsx/DesktopXlsx.csproj -- node_modules/.cache/s19-xlsx
Remove-Item Env:S19_XLSX_OUT
```

The oracle requires the external desktop Release assemblies and is not part of `npm run check` or GitHub CI. A skip or an unavailable assembly is not interoperability evidence.

## Test layers

| Layer | What it proves | Examples |
| --- | --- | --- |
| Unit (Vitest) | Pure model and semantic contracts | typed blanks, reference transforms, formats, formula functions, inverse commands |
| Integration | Boundaries and failure transactions | ZIP bounds, failed IndexedDB transaction, worker stale replies, saved identity race |
| Parity | Agreement with captured desktop behavior | formula scalar/error tables, native package field equality, XLSX style/loss matrix |
| Browser (Playwright) | Real interaction and browser APIs | IME/clipboard fallback, selection/focus, File API, actual Pyodide WASM, chart drag/redraw |
| Performance | Measured workload on recorded hardware/browser | model bounds, scrolling, cold/warm Python, mixed generation latency |

S01 creates a fixture manifest with source path, SHA-256, expected-value source, permitted use and capture date. Use existing expected-values JSON and source assertions; where unavailable, record desktop execution output with its runtime manifest. Never generate the expected result with the new implementation under test. Desktop source remains read-only; any temporary oracle harness/output belongs in this workspace. Do not claim desktop verification when its Windows runtime could not be run.

Tests must cover cancellation/error paths as well as happy paths. Reference transformations should include generated inputs/property tests; native archive readers need corrupt and decompression-limit fixtures. Use deterministic fake clocks for TODAY/debounce and real workers for concurrency races. Real Pyodide tests run in browser workers against packaged same-origin assets; protocol mocks are supplementary. Figure checks compare dimensions and structural image properties/tolerances, not unjustified byte-exact PNG equality across Matplotlib versions.

Use axe with keyboard, focus and accessible active-cell assertions. Keep manual screen-reader smoke testing in the release checks. Screenshots are optional diagnostics. Verify dark UI even when OS preference is light. Preserve forced colors.

## Proposed performance targets

These are targets to calibrate in S02/S03 and enforce in S25, not measured claims. Record OS/CPU/RAM/browser/runtime, fixture dimensions, warm/cold state and sample count. Start with a desktop device with 8 GiB RAM; report p50/p95 over at least 20 warm interactions.

- Typing/selection p95 under 100 ms on a 10,000-populated-cell fixture; ordinary input avoids main-thread tasks above 50 ms.
- Scrolling targets 60 Hz for the visible A1:Z1000 projection; report dropped frames and viewport/zoom rather than conflating model limits with rendered cells.
- Warm native recalculation target under 500 ms for a fixed 1,000-formula dependency fixture; stress 10,000 formulas and 200,000 populated cells without data loss or tab termination, recording observed time/memory.
- Python default timeout remains 30 seconds. Cold runtime load and both showcase cold/warm runs are measured separately; approve realistic browser targets after S02.
- First shell JS/CSS gzip target below 500 KiB excluding lazy Monaco/Pyodide. Runtime compressed/uncompressed asset sizes must be reported and stay below actual Pages constraints. Never include user workbooks in public assets to hit a demo goal.

## Cross-browser parity and performance

S25 adds a desktop-derived ribbon manifest in [tests/fixtures/parity/ribbon-order.json](../tests/fixtures/parity/ribbon-order.json). After a production build, run `npm run test:e2e -- tests/e2e/ribbon-parity.spec.ts` in Chromium, Firefox and WebKit. It checks the tab order, each tab's group order, every menu's command order, the nested Format → Hide & Unhide submenu, split-control pairing, expected glyph presence and explicitly iconless rows (AutoFit Row/Column and Hide/Unhide rows in the current web glyph map, plus the AutoSum menu), viewport placement, an axe scan with an open flyout, and the dark-only and forced-colors palettes. The F2/double-click selection contract is covered by [editing](../tests/e2e/editing.spec.ts).

The maximum-bound workloads are separate from the fast check: run `npm run test:performance`. It builds a deterministic 2-sheet, 200,000-populated-cell / 10,000-formula workbook and a 10,000-cell fixture, records OS, CPU, cores, RAM and Node, and writes p50/p95 samples to ignored `test-results/performance/`. It asserts no data loss at the exact limits, the 10,000-cell session-edit p95 under 100 ms, and the fixed 1,000-formula warm recalculation p95 under 500 ms. The 200,000-cell measurements are recorded rather than asserted against an interaction budget.

Browser input and scrolling are measured after a production build with `npm run test:e2e -- tests/e2e/performance.spec.ts`. Key events are dispatched inside the page so Playwright's protocol overhead is excluded; typing and selection p95 are asserted under 100 ms on the 10,000-cell fixture, Chromium long tasks must stay under 50 ms, and scroll frame times, dropped-at-60-Hz counts and idle cadence are reported. The scroll check samples 15 idle and 60 active frames; each animation-frame wait has a 750 ms fallback and active sampling stops on a frame over 500 ms. Chromium and Firefox retain a 60 ms p95 ceiling. WebKit uses a 350 ms p95 ceiling and the same hard `<500 ms` maximum; this adds a 50 ms margin over the reported CI p95 of 301 ms while still rejecting half-second stalls. The S25 machine measured Chromium at 16.8 ms and Firefox at 25 ms; software-rendered WebKit at 2x device pixels stayed near 47–60 ms. That WebKit value is an environment limitation, not Safari parity evidence, and the 60 Hz target is not claimed for it.

The Python timing case runs with the existing probe harness, now honoring `PLAYWRIGHT_TEST_PORT`:

```powershell
$env:PLAYWRIGHT_TEST_PORT='4211'
npm run test:python-probe -- --project=chromium -g "cold runtime startup"
Remove-Item Env:PLAYWRIGHT_TEST_PORT
```

It records cold worker start plus 20 warm unchanged sales-script runs under the pinned runtime in ignored `test-results/performance/S25-python-timings.json`.

## Check outcomes and workflow

Every sprint updates its parity rows with test paths, then writes docs/handovers/SNN.md. Include pass/fail/skipped/unrun counts where available, not just "tested". A missing real-runtime/interop gate is a blocker for that capability. Fix introduced regressions before advancing. Do not add redundant implementation-mirroring tests merely to increase counts.

If npm encounters a corporate certificate chain, use an approved CA configuration or Node's supported system CA option for the process. Never disable TLS verification or commit machine-specific certificates/settings. Browser downloads may similarly need the approved trust configuration.
