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

`npm run check` combines typecheck, unit/integration tests and production build. E2E is separate because browsers must be installed. CI runs both and installs the three browser engines. The Pages workflow repeats them before uploading the exact tested dist directory. The E2E server serves real static files at `/` and `/ExcelLightWeb/`, with no missing-asset SPA fallback. This catches relative asset-path errors hidden by a development server.

The production browser suite verifies hosting paths, dark rendering, worksheet input, virtualization and accessibility. S01 adds fixture integrity, snapshot and expectation loading checks. Vite's manual preview server is for local development.

## Workbook model and session

`npm run typecheck` first compiles core and application with `tsconfig.model.json` using only the ECMAScript library. It then typechecks the entire project. Run the focused S04 suites with `npx vitest run tests/unit/workbook-model.test.ts tests/integration/workbook-session.test.ts`.

The unit suite exercises the source limits and uses deterministic generated cell edits and reorders to verify forward and inverse deltas. The integration suite covers atomic edits, bounded history, asynchronous save completions and native cache reply guards. A Node worker thread returns a fixed test reply to exercise structured cloning and edit-before-reply ordering. S08 verifies calculation through the browser editor. S09 introduces production worker scheduling.

## Worksheet checks

Run `npm run test:worksheet` to build and serve the fixed two-sheet worksheet fixture under the repository subpath. It checks rendered cell/header alignment at device-pixel boundaries, explicit white fills, hidden and resized tracks, cached formula display, selection statistics and sheet focus. Verify CI runs it after the production browser suite. The fixture builds under `test-results/` and is separate from `dist/`.

The production suite injects the pinned axe-core development dependency to check the worksheet and ribbon. Manual screen-reader verification is scheduled in S25.

## Editing and clipboard checks

`npx vitest run tests/integration/editing-commands.test.ts` covers source-derived typed entry, TSV limits, copy and cut identity, reference rewrites, merges, shared borders, Fill results and AutoSum. `npm run test:e2e` also exercises the production editor in all three engines, including clipboard event payloads, denied API fallback and stale reads. A Chromium-specific case uses the browser clipboard and CDP composition commands. The cross-engine composition guard cases dispatch DOM events.

`npx vitest run tests/unit/formula-hints.test.ts` checks all supported signatures and scans incomplete drafts with quoted text, array constants and nested calls. The production browser suite verifies caret movement and tooltip dismissal in the formula bar and inline editor.

## Structural edit checks

`npx vitest run tests/integration/structure-commands.test.ts` covers the four desktop lane-shift scenarios and checks calculated results, stable identities, axis metadata, overflow and undo. The production Playwright suite exercises clicked-sheet context commands, cell-shift dialogs, header cursor zones and direct header sizing and visibility.

## Browser Python feasibility

Run the S02 probe with Node, Python 3, curl and the Playwright browsers listed above. Set the process variable `PYTHON` if your Python executable uses another name.

```powershell
npm run prepare:python-probe
npm run test:python-probe
```

Preparation verifies the cached files against the runtime lock and downloads any missing or changed file. The test build serves these files and its module worker under the repository subpath. Playwright checks the package handshake and reports the time to initialize a fresh worker in its runtime attachment. Add `-- --project chromium` to focus on one browser. The default CI and Pages workflows use the production build. Python feasibility runs against the dedicated probe build in `test-results/python-probe-build/`.

## Formula and file feasibility

Run the S08 differential suite with `npx vitest run tests/parity/native-formulas.test.ts`. Its synthetic workbook fixtures were evaluated using the desktop Release `NativeCalculator` under `cs-CZ`. The fixture manifest records the assembly hash and runtime. To verify the stored results with the same desktop binaries, run the following from the web repository with .NET 10. Pass `-p:DesktopBin=<absolute path>` for another Release binary directory.

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

## Check outcomes and workflow

Every sprint updates its parity rows with test paths, then writes docs/handovers/SNN.md. Include pass/fail/skipped/unrun counts where available, not just "tested". A missing real-runtime/interop gate is a blocker for that capability. Fix introduced regressions before advancing. Do not add redundant implementation-mirroring tests merely to increase counts.

If npm encounters a corporate certificate chain, use an approved CA configuration or Node's supported system CA option for the process. Never disable TLS verification or commit machine-specific certificates/settings. Browser downloads may similarly need the approved trust configuration.
