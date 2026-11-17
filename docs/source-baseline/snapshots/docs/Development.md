# Development

## Build and run

Windows 11 x64, Visual Studio 2026 with .NET desktop/Windows development tools, and .NET SDK **10.0.402** are the current toolchain. `global.json` disables SDK roll-forward. NuGet package versions and seven project lockfiles pin the dependency graph; first restore needs network access when packages are absent from cache.

From the repository root:

```powershell
./scripts/build.ps1 -Configuration Debug
./scripts/build.ps1 -Configuration Release
./src/ExcelLite.App/bin/Release/net10.0-windows10.0.19041.0/win-x64/ExcelLite.App.exe
```

Or open `ExcelLite.sln`, select ExcelLite.App/x64 and build in Visual Studio. Use normal output directories. Ask the user to close the app only after a build actually fails because its executable/output files are locked; do not bypass locks with a separate build folder.

Build scripts use single-node compilation, disabled shared compilation and the user NuGet cache. They disable NuGet advisory auditing for the restricted/offline invocation; a successful build is not a dependency-security assessment. Normal developer output bundles Windows App SDK but requires the .NET 10 runtime and WebView2. Folder publishing uses self-contained .NET.

## Repository layout

Versioned build/preparation/package tools live in `scripts/`. `build/` is entirely ignored and holds generated runtimes, packages, logs and diagnostic images. Standard per-project `bin/` and `obj/` directories remain local ignored .NET output; they are not source-controlled. UI reference assets live under `docs/references/`; the showcase stays under `docs/showcase/`.

To clean a configuration without running tests:

```powershell
dotnet clean ExcelLite.sln -c Debug -m:1 -p:BuildInParallel=false
# Use -c Release for the other configuration.
```

## Dependencies and offline assets

Current pins include Windows App SDK 1.8.260317003, Windows SDK BuildTools 10.0.26100.4654, Win2D 1.4.0, ClosedXML 0.105.0, Apache.Arrow 23.0.0, and Monaco 0.57.0. Prepared Python is CPython 3.14.10, NumPy 2.5.4, pandas 3.1.0, Matplotlib 3.11.2 and pyarrow 23.0.0. Actual csproj/lockfiles and asset/runtime manifests are authoritative if pins change. [Dependency inventories](dependencies/README.md) preserve package licenses and acquired prerequisite hashes.

```powershell
./scripts/prepare-python.ps1
./scripts/prepare-monaco.ps1
./scripts/prepare-prerequisites.ps1
```

Preparation downloads pinned dependencies; the app uses local Monaco assets and worksheet cursors and private Python, without runtime pip/download requirements or system Python/Excel. Python is generated under `build/runtime/python`; vendored Monaco assets, manifests and licenses are preserved under source control. Do not remove licenses or dependency notices during cleanup.

```powershell
./scripts/package-offline.ps1 -Name ExcelLite-offline
```

Packaging refuses an existing output directory. The folder under `build/artifacts/ExcelLite-offline` contains `app/`, `prerequisites/`, `notices/`, `Setup.ps1` and `PACKAGE.json`. Python is copied beside the app under `runtime/python`; bundled VC++/WebView2 installers handle prerequisites without downloads. Setup/install privileges depend on Windows policy. Developer-host packaging is not proof of clean-machine offline installation, signing, upgrade or uninstall behavior.

## Verification commands

Build Debug and Release. Run the shared .NET suite, the Win2D desktop suite, and the Monaco browser suite from the repository root.

```powershell
./scripts/test.ps1 -Configuration Debug
./scripts/test.ps1 -Configuration Release
./scripts/test-win2d.ps1 -Configuration Release -ForegroundInput
npm ci --prefix tests/MonacoPrototype
npm test --prefix tests/MonacoPrototype
```

Use an interactive desktop for `-ForegroundInput` and leave the test window in the foreground. Those cases send real mouse and keyboard input for scrollbar and header sizing, chart placement and resizing, and ribbon focus recovery. The other desktop cases exercise the actual WinUI controls, worksheet commands and Win2D rendering. The runner writes JSON reports under `build/diagnostics/win2d-<Configuration>/` and exits unsuccessfully if any case fails. Use `-Case` with a case name or wildcard to isolate a failure.

The .NET suite covers workbook history, formulas, clipboard, formatting, table/filter/freeze behavior, persistence and Python scheduling and transport. Focused switches are defined in the test project. The Monaco fixture uses Playwright and installed Chrome to check the bundled offline editor and worker assets.

Run `./scripts/test-win2d.ps1 -Configuration Release -Case '*dropdown-smoke'` to check ribbon Font and Font Size placement, selection visibility and reopening. Its JSON report records the time from programmatic opening to the first rendering callback with the list arranged below the field. Use those timings for local comparisons.

The `--alignment-smoke` report records the time from requesting Format Cells to the dialog loading on the Alignment tab. Run `--number-ribbon-smoke` to verify that the other panes load on demand and that edits survive switching tabs.

Run `./scripts/test-win2d.ps1 -Configuration Release -ForegroundInput -Case '--border-opening-smoke'` to measure Home → Borders in dark mode. The report records the first arranged frame and the first frame with icon sources ready, with one initial opening and five reopens. It identifies whether each sample used a foreground mouse click or invoked the same arrow through UI Automation.

The [SLA showcase](showcase/sla/README.md) has an end-to-end calculation fixture. Use `--sla-showcase-only` on the .NET test executable for its focused checks. Add `--sla-arrow` to compare Arrow inputs with the default JSON transport. Run `tests/Python/chart_redraw.py` and `tests/Python/chart_cache_recovery.py` with `build/runtime/python/python.exe` to check chart pixels, draw counts, PNG metadata and redraw after a worker restart. Both write under `build/diagnostics/`.

Regenerate cursor artwork with `scripts/generate-header-cursors.py` or `scripts/generate-border-cursors.py`. Install Pillow in the development Python environment for generation. The app loads the checked-in Windows cursor assets.
Use `python scripts/generate-border-cursors.py --tool erase` to regenerate just the Erase Border cursor and `python tests/Python/border_eraser_cursor.py` to check its artwork and hotspot at all four shipped sizes.

## Worksheet and recalculation diagnostics

```powershell
$app = Resolve-Path ./src/ExcelLite.App/bin/Release/net10.0-windows10.0.19041.0/win-x64/ExcelLite.App.exe
$workbook = Resolve-Path ./docs/showcase/sla/Support-SLA.exlt
& $app --workbook $workbook
```

On **Dashboard**, change **P9** from **4.0** to **2.0**. In Python Manual mode, choose **Recalculate workbook** and confirm Python enablement if prompted. The breach rate and both charts update. For a repeatable measurement through the full shell, use `--sla-recalc-benchmark --report "<absolute JSON path>"`. The report records the first and repeat explicit recalculations and the Automatic-mode edit through the dashboard frame. The edit timeline shows the debounce, each Python worker request, result commit, worksheet publish, and frame presentation as milliseconds from the start of the edit. The focused `--sla-showcase-only` timing omits the debounce and UI and renders at each script’s original canvas size.

Use `--native-grid-benchmark --workbook "<absolute workbook path>" --report "<absolute JSON path>"` for sheet-switch and scroll timings. The timing includes a Win2D draw and the next compositor callback. Keep the app in the foreground. The focused switches `--check-rendering`, `--check-worksheet-parity`, `--check-input`, `--check-chart-input`, `--check-chart-resize`, `--check-monaco` and `--check-recalculation` exercise the worksheet and Python integration.

## Verification boundaries

Automated desktop checks cover the app's authored rendering and interaction behavior. Offline packaging requires a separate installation check on a clean Windows machine. Python source, spills and linked chart placements are preserved in `.exlt`; see [XLSX compatibility](Xlsx-Compatibility.md) for the exchange boundaries.
