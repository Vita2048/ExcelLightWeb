# ADR-003 — Formula and file adapters

Accepted 2026-11-17. The decision applies to the web implementation of the desktop-supported subset.

## Formula engine

Use a TypeScript tokenizer and AST with explicit calculation values and lazy branch evaluation. Dependency extraction walks parsed references independently of evaluation. Keep source text and reference spans for editing, and transform references using sheet identities and parsed coordinates in S07. S08 must extend the grammar and function support, and S09 must implement the dependency graph and recalculation.

The probe matches selected production `NativeCalculator` assertions under ClosedXML 0.105.0. It covers signed half-away rounding, negative places, numeric text in ABS, text rejection in ROUND, inactive IF and IFERROR branches, typed errors and mixed absolute references. Serial 60 stays numeric through interchange and projects to the explicit date tuple 1900-02-29. S19 applies the 1462-day shift only to imported 1904 date values. Broader coercion, aggregate ranges, graph behavior and display formatting need their own fixtures. [Formula tests](../../tests/parity/formula-probe.test.ts)

## ZIP and native workbook

Pin `@zip.js/zip.js` 2.25.0 (BSD-3-Clause). Read from a Blob using a sequential entry visitor with duplicate, path, CRC and overlap checks. Apply the byte budget to the emitted chunks before passing them to the entry consumer. The typed stream test and the synthetic ZIP with understated metadata exercise this boundary. Production should run the visitor inside the file worker with internal ZIP workers disabled. [ZIP reader API](https://gildas-lormeau.github.io/zip.js/api/classes/ZipReader.html)

Use the existing version-1 ZIP/JSON contract. The feasibility adapter checks the version, manifest/workbook/sheet GUIDs, ordered names and freeze encodings. It re-emits the original parts byte for byte inside a new archive. Three fixtures round-trip with identical decompressed parts. Desktop `NativePackage.Load` reports document equality for the native style fixture repacked by the web adapter. S17 must implement full native validation, shape mapping and cell authoring, including JSON depth, names, spills, placement caches and desktop migrations. [File tests](../../tests/parity/file-adapter-probe.test.ts) · [Desktop harness](../../tests/desktop-interchange/Program.cs)

## XLSX exchange

Use zip.js for container I/O and `@xmldom/xmldom` 0.9.12 (MIT) for small namespace-aware parts such as `styles.xml`. The probe rewrites extended style records while preserving the other parts and their style indexes. Desktop reimport preserved signed and stacked rotation, last-line justification and reading direction, explicitly colored and Automatic borders, shared diagonal color, linear three-stop and path gradients, and the fictitious 1900 leap-day value.

ExcelJS 4.4.0 (MIT) read the sales input but failed on the exact desktop exporter fixture. Its workbook parser expects unqualified names and the fixture uses an `x:` prefix for spreadsheet elements. An ExcelJS-generated control retained path gradient stops and rotation but dropped `justifyLastLine`. A full-workbook JSZip decode would also require a separate bounded reading path. The selected OOXML adapter gives S19 direct access to the needed data. [ExcelJS source](https://github.com/exceljs/exceljs/tree/v4.4.0/lib/xlsx/xform) · [xmldom project](https://github.com/xmldom/xmldom)

S19 still needs relationship and content-type resolution, shared text, cell values and formula caches, dimensions, merges and freeze boundaries. Reuse namespaces instead of assuming the literal XML prefix. Add controlled stream parsing for large sheet parts, explicit date/time classification, canonical style projections and loss reporting for workbook objects and unsupported color and fill data. Handle safe directory records when needed for XLSX. Build and test a visible review for every existing desktop reportable loss.

## Byte and memory budgets

The S03 fixture collector accepts up to 16 MiB compressed, 32 MiB expanded, 4 MiB per entry and 512 entries. At most one entry is inflated at a time. The collector retains decoded bytes and copies the active part once, so the raw decoded-buffer target is at most 36 MiB, excluding codec buffers, ZIP output, JSON objects and the DOM. Keep DOM parsing to one small metadata part at a time. These are enforced byte policies, with no claim about process RSS.

| Fixture | Compressed bytes | Expanded bytes | Largest part |
| --- | ---: | ---: | ---: |
| Sales EXLT | 137,512 | 518,076 | 180,795 |
| SLA EXLT | 403,579 | 3,564,229 | 2,910,137 |
| Styles EXLT | 1,194 | 2,964 | 2,684 |
| Styles XLSX | 6,221 | 19,115 | 7,096 |

Production retains the desktop outer limits of 128 MiB compressed, 512 MiB expanded, 32 MiB per part and 10,000 entries. S17/S19 should consume parts and release buffers one at a time, with at most one 32 MiB raw decoded part retained independently of the validated workbook. Coordinate the string and object budget with the workbook model and bound decompression cumulatively. The production limits require browser memory and timing checks in S25 before a support claim.
