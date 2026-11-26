# ADR-003 — Formula and file adapters

Accepted 2026-11-17. The decision applies to the web implementation of the desktop-supported subset.

## Formula engine

Use a TypeScript tokenizer and AST with explicit calculation values and lazy branch evaluation. Dependency extraction walks parsed references independently of evaluation. Source text and reference spans are kept for editing, and references are transformed using sheet identities and parsed coordinates rather than string replacement.

The captured desktop assertions cover signed half-away rounding, negative places, numeric text in ABS, text rejection in ROUND, inactive IF and IFERROR branches, typed errors and mixed absolute references. Serial 60 stays numeric through interchange and projects to the explicit date tuple 1900-02-29. The 1462-day shift is applied only to imported 1904 date values. Broader coercion, aggregate ranges, graph behavior and display formatting have their own fixtures and suites. [Formula tests](../../tests/parity/formula-probe.test.ts) · [Differential tests](../../tests/parity/native-formulas.test.ts)

## ZIP and native workbook

Pin `@zip.js/zip.js` 2.25.0 (BSD-3-Clause). Read from a Blob using a sequential entry visitor with duplicate, path, CRC and overlap checks. Apply the byte budget to the emitted chunks before passing them to the entry consumer. The typed stream test and the synthetic ZIP with understated metadata exercise this boundary. Production runs the visitor inside the file worker with internal ZIP workers disabled. [ZIP reader API](https://gildas-lormeau.github.io/zip.js/api/classes/ZipReader.html)

The adapter implements the existing version-1 ZIP/JSON contract: version, manifest/workbook/sheet GUIDs, ordered names, freeze encodings, JSON depth, names, spills, placement caches and desktop legacy migrations. Three fixtures round-trip with identical decompressed parts, and the web writer's output reloads to an equal document. [File tests](../../tests/parity/file-adapter-probe.test.ts) · [Native package tests](../../tests/parity/native-package.test.ts)

## XLSX exchange

Use zip.js for container I/O and `@xmldom/xmldom` 0.9.12 (MIT) for small namespace-aware parts such as `styles.xml`. The interchange fixture rewrites extended style records while preserving the other parts and their style indexes. Desktop reimport preserved signed and stacked rotation, last-line justification and reading direction, explicitly colored and Automatic borders, shared diagonal color, linear three-stop and path gradients, and the fictitious 1900 leap-day value.

ExcelJS 4.4.0 (MIT) read the sales input but failed on the exact desktop exporter fixture: its workbook parser expects unqualified names while the fixture uses an `x:` prefix for spreadsheet elements. An ExcelJS-generated control retained path gradient stops and rotation but dropped `justifyLastLine`. A full-workbook JSZip decode would also require a separate bounded reading path. The selected OOXML adapter gives direct access to the needed data. [ExcelJS source](https://github.com/exceljs/exceljs/tree/v4.4.0/lib/xlsx/xform) · [xmldom project](https://github.com/xmldom/xmldom)

The implemented adapter resolves relationships and content types, shared text, cell values and formula caches, dimensions, merges and freeze boundaries, reusing namespaces instead of assuming the literal XML prefix. It adds controlled stream parsing for large sheet parts, explicit date/time classification, canonical style projections and loss reporting for workbook objects and unsupported color and fill data, and handles safe directory records. Every desktop reportable loss has a visible review.

## Byte and memory budgets

The probe fixture collector accepts up to 16 MiB compressed, 32 MiB expanded, 4 MiB per entry and 512 entries. At most one entry is inflated at a time. The collector retains decoded bytes and copies the active part once, so the raw decoded-buffer target is at most 36 MiB, excluding codec buffers, ZIP output, JSON objects and the DOM. DOM parsing stays to one small metadata part at a time. These are enforced byte policies, with no claim about process RSS.

| Fixture | Compressed bytes | Expanded bytes | Largest part |
| --- | ---: | ---: | ---: |
| Sales EXLT | 137,512 | 518,076 | 180,795 |
| SLA EXLT | 403,579 | 3,564,229 | 2,910,137 |
| Styles EXLT | 1,194 | 2,964 | 2,684 |
| Styles XLSX | 6,221 | 19,115 | 7,096 |

Production retains the desktop outer limits of 128 MiB compressed, 512 MiB expanded, 32 MiB per part and 10,000 entries. Parts are consumed and released one at a time, with at most one 32 MiB raw decoded part retained independently of the validated workbook. Decompression is bounded cumulatively; browser memory and timing are checked by the performance suites rather than asserted from these byte policies alone.
