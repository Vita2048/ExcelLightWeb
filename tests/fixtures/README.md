# fixtures

The [manifest](manifest.json) records the origin, SHA-256, and expected-value source for each fixture. Paths in the manifest are relative to the repository root. These files are for local development and automated tests.

The two `.exlt` showcase workbooks live in [`docs/showcase`](../../docs/showcase): `Showcase.exlt` (sales) and `Support-SLA.exlt` (SLA). Only the Support SLA example is bundled in the production build; it was authored for this repository by the project owner, so publishing it carries no third-party redistribution review. The sales workbook is test-only.

- `sales/` holds the original XLSX input, unchanged Python examples, and independently derived expectations supplied by the desktop project. The saved workbook has blue styles on all three tables; the desktop verifier expected Products green and Sales orange. Use the saved workbook for observed import behavior and the desktop table tests for intended styling behavior.
- `sla/` holds the three Python scripts and KPIs independently calculated from the input rows for both target values. The saved Dashboard caches agree with the initial KPIs. The desktop SLA assertions also specify both pivots and chart scenarios.
- `formulas/` holds selected desktop source assertions and synthetic workbook cases evaluated by the desktop NativeCalculator. The captured output is checked by the web evaluator tests; the capture harness has been retired.
- `formats/` holds synthetic desktop number format outputs and a five-cell 1904 XLSX generated with ClosedXML. The desktop importer verified the converted date and date-time serials and the original clock, elapsed and ordinary numeric serials. The 1900 comparison uses the existing DesktopStyles XLSX with serial 60. The capture harness has been retired; the checked-in outputs are the record.
- `interchange/` holds the style interchange pair described in its own [README](interchange/README.md).

Run `npm test` to verify fixture loads and provenance hashes. Git attributes preserve the captured line endings across checkouts.
