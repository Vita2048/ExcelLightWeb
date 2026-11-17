# fixtures

The [manifest](manifest.json) records the origin, SHA-256, and expected-value source for each fixture. These files are for local development and automated tests. The desktop tree had no repository-wide license file. Verify redistribution rights for the copied assets and source snapshots before publishing them in a public repository.

- `sales/` holds the supplied showcase workbook, original XLSX, unchanged Python examples, and independently derived expectations supplied by the desktop project. The desktop verifier reaches the style check after passing its numeric and XLSX input checks. The saved workbook has blue styles on all three tables. The verifier expects Products green and Sales orange. Use the saved workbook for observed import behavior and the desktop table tests for intended styling behavior.
- `sla/` holds the synthetic tickets workbook, its three Python scripts, and KPIs independently calculated from the input rows for both target values. The saved Dashboard caches agree with the initial KPIs. The desktop SLA assertions also specify both pivots and chart scenarios.
- `formulas/` holds selected source assertions for use by S03 and the formula sprints. Compare current NativeCalculator assertions for production behavior and error codes.

Run `npm test` to verify fixture loads, provenance hashes and frozen snapshot bytes. Git attributes preserve the captured line endings across checkouts.
