# Feature parity and browser adaptations

S01 captured the [command inventory](source-baseline/Command-Inventory.md), [desktop test inventory](source-baseline/Test-Inventory.md), and [source hashes](source-baseline/source-manifest.json). Product rows remain planned. The S01 checks confirm the fixture baseline. [Verification](../tests/parity/baseline.test.ts) · [Handover](handovers/S01.md)

References below are filenames under the source src/ExcelLite.*, tests or docs tree. Locate them from C:/Temp/ExcelLight. Treat current source and assertions as authority when high-level documentation differs. Columns for evidence/status must be updated by each sprint.

| ID | Required existing behavior | Source / oracle | Sprint | Status / evidence |
| --- | --- | --- | --- | --- |
| F01 | File Home/New/Open/Save/Save As/Export/Close; Recent 10; native/import identity | FileLifecycle.cs; DocumentSession.cs; FileLifecycleSmoke.cs | S17–S19 | Planned; no parity evidence yet |
| F02 | Sparse types, limits, stable identities, history 100/64MiB, dirty/save races | DocumentModel.cs; DocumentContracts.cs; WorkbookSession.cs | S04 | Planned; no parity evidence yet |
| F03 | Black top bars/dark panels, command families, Options drafts, Quick Access | Menu-Structure.md; Design-Guidelines.md; ShellMenus.cs | S05 | Planned; no parity evidence yet |
| F04 | Grid virtual rendering, rectangle/multi-area selection, keyboard/IME, status | NativeGrid*.cs; GridContextMenu.cs | S05–S06 | Planned; no parity evidence yet |
| F05 | Inline/formula entry; paste/all/values/formulas; cut/copy; clear variants; AutoSum; Fill | GridClipboard.cs; FillChecks.cs; ClearChecks.cs | S06 | Planned; no parity evidence yet |
| F06 | Sheet create/delete/rename/move/copy; axis and lane insert/delete; reference transforms | StructuralReferences.cs; CellShiftChecks.cs; StructureSmoke.cs | S07 | Planned; no parity evidence yet |
| F07 | Row/column resize/autofit/hide/unhide, header cursors, clicked-sheet contexts | NavigationSizingSmoke.cs; StructureCommands.cs | S07 | Planned; no parity evidence yet |
| F08 | 31 formulas, coercions/errors, exact lookup, lazy branches, subtotals | FormulaSubset.cs; NativeCalculator.cs; SubtotalEvaluation.cs | S08 | S03 verifies selected lazy branches, rounding, coercion and errors. Full evaluation remains in S08. [Formula tests](../tests/parity/formula-probe.test.ts) |
| F09 | Native Automatic/Manual; Calculate Sheet/Now; dependency cycles; TODAY; locale separators | NativeCalculationGeneration.cs; FormulaPresentation.cs; FormulaRibbon.cs | S09 | Planned; no parity evidence yet |
| F10 | Number/currency/accounting/date/time/fraction/scientific/custom/special format catalogs; temporal editing | NumberFormatChecks.cs; TemporalFormatOptions.cs; SpecialFormatOptions.cs | S10 | Planned; no parity evidence yet |
| F11 | Fonts 6–72pt, effects, underline variants, custom colors, patterns/gradients; painter | FontFillChecks.cs; FormatPainterChecks.cs; CellAppearanceDraft.cs | S11 | Planned; no parity evidence yet |
| F12 | 13 border styles, independent edges/diagonals, presets, draw/grid/erase, preview | BorderChecks.cs; BorderToolChecks.cs; BorderFormatDraft.cs | S12 | Planned; no parity evidence yet |
| F13 | 8 horizontal/5 vertical alignments, indent/wrap/shrink, reading order, stacked/rotated text | AlignmentChecks.cs; FormatAlignmentTab.cs | S13 | Planned; no parity evidence yet |
| F14 | Merge & Center/Across/Cells/Unmerge; discard confirmation; covered-cell constraints | MergedCells.cs; AlignmentCommands.cs; AlignmentChecks.cs | S13 | Planned; no parity evidence yet |
| F15 | Workbook names and Name Manager; name/table uniqueness; copied identities | WorkbookNames.cs; NameCommands.cs | S14 | Planned; no parity evidence yet |
| F16 | Tables explicit extents, resize grip, cached headers/totals/styles, contextual tab memory | TableDesignChecks.cs; S18TableUiChecks.cs; TableDesignCommands.cs | S14 | Planned; no parity evidence yet |
| F17 | Stable single-key sort, expansion and header options; formatted filters, search/checklist, reapply | DataCommands.cs; DataRibbon.cs; S12Checks.cs | S15 | Planned; no parity evidence yet |
| F18 | Selection stats with manual/filter hidden rows; nested SUBTOTAL visibility | SubtotalEvaluation.cs; EverydayCommands.cs | S15 | Planned; no parity evidence yet |
| F19 | Find/Replace scopes/options/navigation, 15-entry histories, Replace All undo | SearchDialog.cs; EverydayUiSmoke.cs | S16 | Planned; no parity evidence yet |
| F20 | Formula Bar/Gridlines/Headings; Show Formulas; zoom 10–400; fit; freeze two axes | FreezeChecks.cs; FormulaZoomSmoke.cs; NativeGridViewCommands.cs | S09, S16 | Planned; no parity evidence yet |
| F21 | Full native .exlt v1 incl legacy fields, Python caches/spills/placements/images; hostile ZIP rejection | NativePackage.cs; tests/TestSheet/TestSheet.exlt | S17 | S03 verifies GUIDs and freeze encodings, byte-equivalent repackaging and streamed ZIP limits. S17 owns full mapping and validation. [File tests](../tests/parity/file-adapter-probe.test.ts) |
| F22 | Local settings, independent native/Python modes, Recent and save lifecycle | LocalSettings.cs; FileLifecycle.cs | S18 | Planned; no parity evidence yet |
| F23 | XLSX scalar/formula caches, full supported styles, merges/dimensions/visibility/freeze/date conversion | Xlsx-Compatibility.md; XlsxExchange.cs; XlsxStyles.cs; XlsxGradients.cs | S19 | S03 verifies desktop reimport of extended styles and serial 60. S19 implements data mapping, the full style subset and loss reports. [Desktop harness](../tests/desktop-interchange/Program.cs) |
| F24 | XLSX reviewed loss reports/rejections; Python materialized values only; export never executes | Xlsx-Compatibility.md; XlsxExchange.cs | S19 | Planned; no parity evidence yet |
| F25 | =PY anchors, Monaco drafts/local undo, panel clearing, xl caret range picking, legacy bindings | EditorSession.cs; PythonBlocks.cs; PythonReferences.cs; MonacoPrototype tests | S20 | Planned; no parity evidence yet |
| F26 | Literal xl references/names/#All/#Data; preflight syntax/headers; source rewrite spans | PythonReferences.cs; Python-Capabilities.md; worker/excellite_worker.py | S20–S21 | Planned; no parity evidence yet |
| F27 | Fresh namespace; final expression/fallback; scalar/Series/DataFrame/2D/Figure; errors/logs/limits | PythonExecution.cs; worker/excellite_worker.py | S02, S21 | S02 feasibility covers result conversion and Arrow IPC. Application limits and error lifecycle are planned in S21. [Probe cases](../tests/python-browser/python.spec.ts) |
| F28 | Spill occupancy/shrink/ownership; read-only children; Copy/Insert as Values; freshness/last success | PythonSpills.cs; PythonTablePreview.cs; S17Checks.cs | S21 | Planned; no parity evidence yet |
| F29 | Consent, local Automatic/Manual, 75/220ms debounce, stop/restart, stale generations | PythonScheduling.cs; MixedCalculationChecks.cs | S21–S22 | S02 exercises the 30-second deadline, worker termination and fresh execution after restart. Scheduling and consent remain planned. [Probe cases](../tests/python-browser/python.spec.ts) |
| F30 | Mixed native/Python graph, prerequisites, spill replanning, atomic undo/save; no execute on undo | MixedPythonCalculation.cs; MixedCalculationChecks.cs | S22 | Planned; no parity evidence yet |
| F31 | Figure preview/save/place/move/resize/delete; multiple linked placements; cache recovery | FigureInteractions.cs; tests/Python/chart_redraw.py; chart_cache_recovery.py | S23 | S02 redraws a simple Figure from a private in-memory serialization. Placement and recovery across worker restarts remain planned. [Probe cases](../tests/python-browser/python.spec.ts) |
| F32 | Sales and SLA showcases end-to-end, including edits, KPIs, spills and chart redraw | docs/showcase/Expected-values.json; showcase/sla; SlaShowcaseChecks.cs | S02, S22–S25 | S02 runs the unchanged Python scripts with explicit snapshots and checks numeric output and PNGs. Workbook-level integration remains planned. [Probe cases](../tests/python-browser/python.spec.ts) |
| F33 | Keyboard/focus/dialog semantics, accessible grid, forced colors, font/scaling fidelity | NativeGridAutomation.cs; Dialog*.cs; Design-Guidelines.md | S05–S16, S25 | Planned; no parity evidence yet |
| F34 | Offline bundled runtime/assets/notices; browser cache/recovery adaptation | docs/dependencies; scripts; PythonRuntime.cs | S24 | S02 pins and verifies the same-origin probe closure. Offline production packaging remains planned. [Runtime lock](../src/python/runtime-lock.json) |

## Formula allowlist

SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, IF, AND, OR, NOT, IFERROR, ABS, ROUND, ROUNDUP, ROUNDDOWN, SUMIF, COUNTIF, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, DATE, YEAR, MONTH, DAY, TODAY, VLOOKUP, SUBTOTAL. Exactly 31 names; capture FormulaSubset.cs in S01 and flag any baseline difference. SUBTOTAL codes 1–11 and 101–111 are required, even though the total-row picker offers a smaller list.

## Explicit adaptations and exclusions

| Desktop behavior | Browser treatment | Acceptance evidence |
| --- | --- | --- |
| Light / Follow system theme | Removed by user request; dark only, preserve authored cell colors and forced-colors accessibility | OS light preference test and no theme selector |
| WinUI/Win2D/WebView2 and Windows automation | React/Canvas/DOM/Monaco, semantic accessible grid and Playwright | Rendering/input/focus and accessibility suites |
| CPython process with arbitrary OS/native extensions | Pyodide WASM; supported package/workflow matrix, explicit unsupported errors | S02 actual-runtime report; cannot claim arbitrary desktop script parity |
| Local file paths, atomic replacement and Recent paths | IndexedDB local save plus permissioned file handles where available; upload/download fallback, Recent workspace IDs | Quota/cancel/reload/save race and fallback tests; download is not confirmed disk save |
| Windows process stop and retained disk figure snapshots | Worker termination on Stop; in-memory bounded retention; authorized rebuild after worker loss | S02/S23 restart/redraw tests and visible fallback |
| Installed-font enumeration | Curated safe font choices + typed names; optional permissioned enumeration where supported | Font fallback and source-family persistence tests |
| Windows regional overrides | Browser locale and explicit app locale; custom OS separator overrides may require an explicit preference | cs-CZ/en-US golden formatting tests and documented difference |
| Offline Windows installer | Static app with downloaded/cached versioned runtime after first online load | S24 offline cache and update tests |
| Desktop no autosave claim | Explicit browser recovery snapshots added as adaptation | S18 reload/conflict/quota tests |
| Full Excel features already excluded by desktop | Keep excluded: CSV, macros, pivots, external workbook execution, native dynamic arrays, general structured formula names, full Excel chart/table model | Unsupported-format/function tests and loss report |

These adaptations are architectural consequences of browser-local hosting, not permission to quietly discard product workflows. If S02/S03 cannot preserve a required workflow, mark blocked and present the incompatibility and alternatives. Do not count a cached screenshot or a mocked Python result as execution parity.

## XLSX losses that must remain visible

Preserve current reports for AutoFilter/sort indicators, unsupported names/tables/charts/drawings/pivots/comments/validation/conditional formatting/hyperlinks/protection, theme/indexed colors/rich text, hidden-sheet visibility and unsupported formula caches. Export only existing materialized Python cells; omit code/bindings/trust/figures with a reviewed report. Do not broaden scope to full Excel roundtripping. Reject .xls/.xlsb/.xlsm, encrypted or renamed macro files and malformed/out-of-bounds packages. File extension alone is not sufficient validation.

