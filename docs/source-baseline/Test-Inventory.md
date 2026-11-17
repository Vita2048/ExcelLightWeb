# Desktop test inventory

These are existing test entry points and source oracles. Refer to [source-manifest.json](source-manifest.json) for individual file hashes and the frozen paths. Suites below are targets for web coverage. A source assertion is evidence of intended behavior.

| Desktop test family | Where to find it | Web coverage |
| --- | --- | --- |
| Command dispatch and workbook/session contracts | tests/ExcelLite.Tests/Program.cs, S03–S07 checks, WorksheetRefreshChecks | S04–S07 |
| Formula prototype, native formulas and presentation | Program.cs, S08–S10 checks, FormulaPresentationChecks, MixedCalculationChecks | S08–S09, S22 |
| Sheet and cell operations, formatting and view | CellShift, Clear, Fill, Alignment, Border, BorderTool, FontFill, FormatPainter, FormatDialogOptions, NumberFormat, SpecialFormat, TemporalFormat and Freeze checks | S06–S16 |
| Sort, filter, names and table design | S12/S18/S18TableUi, TableDesign and UxRefinement checks | S14–S16 |
| Native and Excel exchange, including Office probe | S11/S13/S15 checks, CreateExcelFixture.ps1 and VerifyExcelExport.ps1. Search NativePackage/XlsxExchange in the full test tree for additional cases | S17–S19 |
| Python protocol, typed conversion, formulas and mixed execution | S16–S21, ArrowTransport and MixedCalculation checks, worker test fixtures | S02, S20–S22 |
| Editor protocol and range picking | tests/MonacoPrototype/editor.test.cjs and NativeGridPythonChecks | S20 |
| Figure layout and recovery | tests/Python/chart_redraw.py, chart_cache_recovery.py and FigureRecoveryProbe | S23 |
| Cursor artwork | tests/Python/border_eraser_cursor.py and BorderToolChecks | S12 |
| Desktop application UI checks and performance | src/ExcelLite.App/*Smoke.cs, NativeGrid*Checks.cs, SlaRecalculationBenchmark.cs and scripts/test-win2d.ps1 | S05–S25 |
| Sales values and chart placements | docs/showcase/tools/verify_showcase.py and sales Expected-values.json | S02, S22–S25 |
| SLA KPIs, pivots, failure, cache and redraw | SlaShowcaseChecks.cs, SlaShowcase.cs and synthetic Support-SLA.exlt | S02, S22–S25 |

S08 and S09 should use the current NativeCalculator assertions for production error codes. FeasibilityWorkbook used historical prototype display codes. The selected formula fixture cases identify their evaluator.
