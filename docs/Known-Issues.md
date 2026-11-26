# Known issues and unverified checks

This file records limitations that remain open after release acceptance. An item here is not parity evidence and must not be described as verified until it is closed with recorded evidence.

## Release and deployment

- The live site is <https://vita2048.github.io/ExcelLightWeb/>. Repository **Settings → Pages → Source** must stay **GitHub Actions**; the default `pages build and deployment` branch workflow publishes raw Vite source and causes `/src/main.tsx` 404s. Every push to `main` deploys automatically, so `main` must stay releasable.
- Manual assistive-technology verification (screen reader) and the Windows Edge smoke were never run. Automated axe, keyboard, focus and forced-colors checks cover the same surfaces, but neither manual check is claimed as parity evidence.

## Verification gaps

- Desktop C# oracle projects were retired with the desktop code. The captured fixtures and expectations remain, but desktop document-equality, formatter and XLSX reimport comparisons can no longer be re-run from this repository. Existing evidence is historical, not a current gate.
- Playwright WebKit cannot navigate a service-worker response after `context.setOffline(true)` ([upstream issue #42775](https://github.com/microsoft/playwright/issues/42775)). Its offline-navigation case is skipped; online scope and CSP checks still run.
- Local Playwright Firefox on the development machine fails to launch (`nsIWindowsRegKey.create`). The GitHub Linux matrix is the cross-browser authority; a local Firefox run is unavailable.

## Performance

- WebKit scroll frame times come from a software rasterizer at 2x device pixels. The suite uses a 350 ms p95 ceiling with a hard `<500 ms` maximum and prints the measured frames on every run. Recorded p95 values are 116 ms and 147 ms locally and 301 ms in Linux CI. The 60 Hz target is not claimed for WebKit, and these numbers are an environment limitation rather than Safari parity evidence.
- A single-cell edit on the 200,000-cell / 10,000-formula maximum-bound workbook measured p95 669 ms. The 100 ms interaction target applies to the 10,000-cell fixture (measured p95 33.7 ms); the maximum-bound cost is a recorded main-thread limitation, not a passing budget.
- Cold Pyodide startup is about 20 seconds on the recorded machine. Python is preloaded on page load, so this is background startup rather than blocking input, but first-run availability still waits for it.

## Compatibility limits

- The native `.exlt` writer emits format version 1 only. There is no migration decision for a future incompatible schema; older legacy fields are read and migrated.
- XLSX support is the reviewed desktop subset, not general Excel round-tripping. Export materializes scalar Python results only; Python code, bindings, tables and figures stay native and appear in the loss report. CSV, macros, pivots, external workbooks, native dynamic arrays and full chart/table models remain excluded.
- The desktop's current ClosedXML reader drops the column-width attribute its own exporter writes, so desktop width fidelity cannot be asserted; the web adapter preserves the attribute on import and export.
- Six ribbon rows render without glyphs because the web glyph map has none: AutoFit Row Height, AutoFit Column Width, Hide Rows, Hide Columns, Unhide Rows and Unhide Columns. Their command order is still checked.
- Browser storage is origin-scoped and subject to quota, eviction and private-mode loss. Export is the portable backup; project paths on one `github.io` origin are not a security boundary.

## Security and privacy

- Pyodide executes workbook code in a worker but is **not** a security sandbox: user code can reach network and browser capabilities. Automatic is the default calculation mode and runs Python when a workbook opens and after relevant edits; Manual keeps imported source inert until Run.
- First visit and first offline preparation require connectivity. The service worker never caches workbooks or user Python results.
