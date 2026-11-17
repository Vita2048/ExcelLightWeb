# XLSX supported subset

The six ribbon currency choices use existing literal number-format codes, preserved by native/XLSX persistence: Czech Kč and Polish zł trail the amount, US $ and UK £ precede it, and Euro/Slovak € trail it. They do not introduce new numeric types or locale-specific separator behavior. Decimal buttons preserve these preset symbols and placement.

| Content | Behavior | Limits / report |
| --- | --- | --- |
| Sheets / scalar values | Sheet order/names, numeric/text/boolean values and explicit empty text exchange | New stable ids; constant Excel errors become literal error text with a report |
| Formulas / caches | Source and typed caches exchange without export-time evaluation | Native supported-function subset; imported unsupported/array/reference formulas keep unverified caches, never supply a current dependent result; array/spill export is scalar-formula only |
| Cell styles | Bold, family/size, italic, font effects, five underline styles, number format, RGB fill/text, patterns/gradients, all thirteen border line styles, independent RGB edges and diagonals, all eight horizontal and five vertical alignment modes, wrap, shrink-to-fit, indentation, reading order, justify-last-line and signed/stacked text rotation | Theme/indexed colors, shadow, unsupported fills and rich-text runs are reported losses; font sizes clamp to native bounds |
| Merged cells | Rectangular merged ranges import/export with upper-left content | Bounds, overlap, table and Python-spill conflicts reject adoption; unmerge does not restore discarded values |
| Dimensions / visibility | Explicit row/column dimensions and manual hiding | Row points convert to logical pixels; column widths approximate Calibri 11 / 96 DPI; Excel export limits clamp with a report; nonstandard defaults and axis-wide styles retain only explicit axes/cells |
| Freeze | Frozen row and column boundaries on each worksheet | Independent split windows and out-of-bounds freezes are reported as omissions. Imported frozen boundaries are relative to the worksheet origin. |
| AutoFilter / sort indicators | Native-only | XLSX import/export reports omissions; stored hidden rows import as manual hiding |
| Dates | 1900 raw serials/fractions including serial 60; date-formatted 1904 values/caches convert +1462 | Time-only/plain numerics unchanged; formula source/date constants unchanged and require review |
| Objects / metadata | Unsupported names, tables, charts/drawings, pivots, comments, validations, conditional formats, hyperlinks and protection are reported | No full-fidelity Excel round trip; hidden sheets retain cells but become visible |
| Python | Already materialized scalar/table cells export | Source/bindings/trust/figures omitted; export never runs Python |
| Formats / bounds | Unencrypted `.xlsx` only; bounded compressed/expanded parts, cell/formula counts and coordinates | Reject `.xls`, `.xlsb`, `.xlsm`, encrypted/renamed macro workbooks, malformed/out-of-bounds packages before adoption |
| File lifecycle | Continue/Cancel report before adoption/export; atomic separate output | Export does not mark native content saved and cannot replace the imported source during its session, even after native Save As |

Import/export is implemented; historical adapter checks do not establish full Excel compatibility. Office 2024 was used locally; Excel 2021 compatibility is not inferred. CSV remains outside scope. See [Development](Development.md#verification-boundaries) for current verification limits and [Architecture](Architecture.md#native-files-and-lifecycle) for native persistence and safe file adoption.

## Number formatting

Number presets now include Accounting, Short/Long Date, Fraction and Scientific. Format Cells writes standard Excel codes, including grouping, negative sections, elapsed time, fixed/variable fractions, special types and custom codes (up to 128 characters). Native/XLSX persistence preserves these codes; General can import as an implicit default. Formats without an explicit locale retain invariant display; Date/Time locale choices localize tagged formats. Excel’s complete localized format catalog is not implemented. Accounting/Comma fill directives are preserved, but grid text does not expand asterisk fills to Excel column-width alignment.

## Date/time locales and editing

Date/Time locale choices use standard Excel `[$-LCID]` tags. Rendering and formatted-value filter keys use that culture, including localized month/day names. Native and XLSX roundtrips retain the tag. Localized edits convert back to the 1900 numeric serial, including Excel’s fictitious 29 February 1900; formats and numeric types remain intact.

## Cell editing and Special formats

Special formats now include curated country-specific numeric postal, phone and identity masks for Czech, Slovak, US, UK, German, French and Polish locales. Locale tags, literal separators and prefixes persist in native/XLSX files; leading zeroes are presentation only. The UK catalog contains numeric telephone masks; alphanumeric postcodes remain text.

## Font and Fill formatting

Strikethrough, superscript/subscript, single/double accounting underlines, eighteen pattern choices (including None) and two-color gradients now exchange. Supported gradient directions and path origins are stored in standard fill XML; center variants may canonicalize to equivalent reversed color order on import. Unsupported stop counts/colors/angles are reported or approximated as stated by the import report. Font size remains 6–72 pt. Solid ribbon fills replace any pattern/gradient. Unsupported shadow and rich-text runs retain loss reporting.

## Format Cells control layout

Numeric formats chosen in Format Cells carry the regional locale tag while retaining invariant Excel placeholder syntax. Windows custom grouping/decimal overrides are used when the tag matches the current culture. Untouched existing codes retain their authored representation; previews use the regional culture. Czech/Slovak negative alternatives include explicit minus-sign sections. Color selector layout changes do not alter stored RGB values.


Border exchange preserves all thirteen Excel line styles, independent edge RGB/Automatic colors and both diagonal directions. As in Excel/OOXML, both diagonals share a line style and color. Legacy thin borders remain readable.

Alignment exchange maps directly to ClosedXML/OOXML horizontal/vertical modes, TextRotation, Indent, ShrinkToFit, ReadingOrder and JustifyLastLine. These attributes no longer produce extended-formatting omission warnings. Center Across Selection retains separate cells. Browser text measurement and clipping can differ from desktop Excel; native row heights remain explicit.
