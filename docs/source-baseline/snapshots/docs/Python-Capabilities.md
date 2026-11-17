# Python capabilities

This describes implemented behavior and supported limits. [Development](Development.md#verification-boundaries) distinguishes observed checks from unverified UI behavior; [the showcase guide](showcase/README.md) describes the delivered workbook.

## Authoring

Excel-Lite is an offline spreadsheet application with Python analysis cells. Select a cell, type `=PY`, write a self-contained script using `xl()` to read worksheet data, and return a number, text, table or chart through the final expression. The app manages reference tracking, output placement and source persistence. No Create block or Apply bindings step is required for normal use.

This is a documented compatibility subset, not 100% Microsoft Python in Excel compatibility. Scripts run locally under the current account; files/network access are not sandboxed or blocked. No internet, Excel or system Python is required for the bundled workflows. Accessing outside files/network from user scripts is possible, but external data changes are not tracked automatically and are outside the verified showcase workflows.

## Workbook access API

| Form | Behavior |
| --- | --- |
| `xl("A1")` | Typed scalar from the current sheet; missing cell is None |
| `xl("A1:C10")` | DataFrame of the rectangle, no inferred headers; integer column labels |
| `xl("A1:C10", headers=True)` | First row supplies column labels, remaining rows are data |
| `xl("'Sales Sheet'!A1:C10", headers=True)` | Same behavior on an explicitly named sheet |
| `xl("TaxRate")` | Workbook-scoped named cell/range; scalar or DataFrame by shape |
| `xl("SalesTable[#All]", headers=True)` | The named table's explicit extent including its header row, converted to a DataFrame |

References are literal Python strings and `headers` is a literal bool, default False. A single cell with headers=False returns a scalar; headers=True always means DataFrame conversion, including a one-cell header-only empty DataFrame. A multi-cell one-row/one-column rectangle returns a DataFrame. Header labels must be nonblank unique text; invalid headers produce an actionable error rather than guessed names. A header-only table has zero data rows and retains its columns.

The parser accepts one positional reference and the optional keyword `headers=True` or `headers=False`, including multiline calls/comments and trailing commas. Ordinary, raw and triple-quoted reference strings are supported; quote/backslash and hexadecimal Unicode escapes are decoded. Concatenated strings, f/template/byte reference arguments, computed arguments, positional headers, aliases and shadowing the reserved `xl` name are unsupported. Calls inside interpolated strings are also unsupported: assign an xl value to a variable before formatting it. Unrelated string text is not a dependency. The host lexes strings/comments to discover source spans without running Python; the worker validates Python AST syntax and all discovered header rows before executing user statements. All supported references are captured, including calls in branches/functions that are not invoked. A transformed deleted input is `xl("#REF!")` and rejects on Run until repaired. Source drafts, including incomplete scripts, remain saveable.

Support A1 relative/absolute/mixed coordinates and quoted cross-sheet names within current workbook limits. Copy/move/structure/rename update actual xl reference literals using Python syntax spans and parsed workbook references; unrelated strings/comments remain untouched. Names are workbook-scoped and case-insensitively unique across named ranges/tables, with no ambiguous A1-like names. Names use ASCII letters/underscore followed by letters/digits/underscore (up to 255 characters); A1/R1C1-like identifiers and R/C are reserved. At most 1000 workbook definitions are supported. Tables have explicit extents, optional visible headers/totals and stored column labels. Bare table names denote the full extent; `[#Data]` excludes header and total while retaining labels for `headers=True`. Header Row off caches labels, clears the former header cells and excludes that row from table membership; enabling it restores labels above the data only when safe. Visible labels must be nonblank unique text. Invalid edits/split extents require explicit repair. See [Table Design](Menu-Structure.md#tables). Cell copies reuse workbook names; sheet copies receive independent suffixed names and retargeted Python source. Sorting requires whole records across all table columns without moving the header. They do not provide full Excel table features.

Dynamic reference strings, aliases of xl, arbitrary external workbook references, union/whole-axis xl references, structured table selectors other than #All/#Data, native-formula structured names and a general spill-reference syntax such as `xl("A1#")` are unsupported. Reject unsupported xl forms clearly before execution. To consume generated output, reference a supported explicit rectangle or keep intermediate DataFrames within one script. The dependency system conservatively tracks all supported literal xl calls, including unexecuted branches.

## Types and results

- Numbers, text and booleans retain their workbook types. Text-formatted numeric strings stay text; no implicit numeric/culture conversion. Merged ranges expose content only at the upper-left coordinate; covered coordinates read as None. Python anchors/spill outputs cannot overlap merged ranges. Missing input cells become None; pandas may represent missing numeric column entries as its normal missing values.
- Valid date/time-formatted numbers become documented ISO strings, not Python datetime objects. Fictitious Excel 1900-02-29/serial 60 as a date, invalid dates, spreadsheet errors and unverified/stale formula inputs reject before execution with the source location where available.
- The final expression supplies the result. For compatibility with older scripts, `result = ...` remains a fallback only when there is no final expression. A final expression takes precedence; None is valid blank. Missing output fails visibly.
- A trailing `print(...)` is an expression returning None, so it produces a blank even if `result` was assigned earlier. Put diagnostics before the final assignment, or finish with `result`. Existing source is never silently rewritten on load. The conversion-version fingerprint makes older caches non-current until an explicit run.
- Supported outputs: finite numeric/text/bool/None scalars and normalized NumPy scalars; DataFrame/Series/rectangular 2D arrays; explicit Matplotlib Figure. Lists/arrays that do not meet the documented rectangular 2D contract, arbitrary custom objects, infinite values and higher dimensions fail clearly.
- DataFrame worksheet output includes headers and omits the index by default. Series receive a documented one-column header. NaN/NA/NaT output values become blanks with a conversion notice. Optional index inclusion is explicit, never guessed.
- DataFrames spill into adjacent cells from their anchor as one validated transaction; occupied cells produce #SPILL!, never overwrite. Shrinking removes only owned generated cells. Generated children are read-only and direct the user to the source anchor. Last successful outputs remain inspectable but non-current after failure or blocked refresh.
- Clear All removes a selected Python anchor and its complete spill in one undoable operation. Clearing an isolated generated child is rejected without changing the worksheet.
- Figures have a panel preview, Save Image and bounded linked in-sheet static PNG placements for Charts/Dashboard. They are bounded PNG results, not native or interactive Excel charts. Scalar/table/figure outputs survive native Save/reopen.

## Execution and editing

Each cell/run has an independent namespace. Import modules and define helpers in the script that uses them. There is no shared workbook variable/import state or Microsoft row-major calculation-order promise.

Python Automatic/Manual is saved locally in File → Options → Formulas, separately from native calculation. Automatic enables Python and refreshes stale cells on Opening/New/import/reopen. An isolated, committed cell edit queues a refresh after 75 ms. Additional edits while queued or calculating, paste, and compound edits coalesce for 220 ms. Affected native and Python cells recalculate in dependency order. Missing older preferences retain manual execution until a mode is chosen. Manual retains stale results until **Recalculate workbook**. The first explicit run in a disabled session prompts for consent. Selection and uncommitted typing do not execute source. Source drafts flush at selection/file/run barriers; committed source changes can run in enabled Automatic. Stop clears queued work until another committed input edit or explicit run/enable. No parallel Python blocks.

Opening/New/import/reopen starts the interpreter when the saved Python preference is Automatic, even for an empty workbook. Current caches are skipped except placed chart sources missing live/archive rendering state, which are prepared once for resizing. Manual resets session enablement on adoption. Workbook files never contain authorization; the local Automatic preference explicitly authorizes adoption-time enablement. Undo/Redo cancel queued work and never run code; fingerprints determine restored freshness. Mixed results apply as one history transaction only if their starting snapshot still matches; saves see complete generations, and edits during execution discard superseded results.

Selecting a Python anchor or owned spill child opens its source. Selecting any other cell clears the editor/results; background results do not refill that cleared panel. Typing `=PY` creates its anchor and opens Monaco. With the Monaco caret inside an `xl` reference string, mouse range selection fills that literal. Ctrl+Enter in Monaco runs **Recalculate workbook**. **Hide** next to **Recalculate** returns focus to the worksheet and unchecks **Formulas → Python**. **Python options** can commit or discard a source draft or change Automatic Python. Text Undo stays in Monaco. The panel remains resizable, and full errors/output can be read, selected and copied (I40).

Type `xl("")` in Monaco and put the caret between the quotes; then drag a worksheet range to fill that literal. Picking another sheet inserts a quoted sheet name. Completing a pick moves the caret beyond the literal so ordinary later cell selections open their source. Reference insertion participates in Monaco's local Undo. Enter commits a cell's `=PY`/`=PY()` entry; Escape cancels it before an anchor is created. Text-formatted entries remain text. The panel omits the legacy block selector, Create/Apply bindings buttons and binding field. Create through `=PY`, select worksheet anchors to open source and use `xl()` for new input references. Existing saved binding metadata still executes and persists internally.

## Available libraries versus verified workflows

The mixed dependency scheduler combines native formulas, Python references and spill ownership. A generation reuses one native calculation context and the dependency graph from the edit. It refreshes the native prerequisites of each ready block, and recalculates native formulas invalidated by a new Python result. The graph is rebuilt only when a spill changes cell topology, and the scheduler replans when those edges change. Recalculate workbook commits pending edits and requests a workbook calculation. Newly emitted spill children can reveal ownership edges, so the detached scheduler replans before continuing. Intermediate generation edits validate the changed anchor, spill and figure records; the live workbook commits the finished generation as one undo. Cycles and failed upstream inputs block dependent source, while last successful outputs stay inspectable. See [Architecture](Architecture.md) for snapshot/commit boundaries.

Pinned baseline from the existing runtime manifest: CPython 3.14.10, NumPy 2.5.4, pandas 3.1.0, Matplotlib 3.11.2 and pyarrow 23.0.0. Final release documentation must read the actual packaged manifest and update these values if an approved change is made. Packages are prepared during build, with no runtime pip/download requirement. Python inputs use JSON by default. The diagnostic Arrow input path falls back to JSON for heterogeneous columns. Eligible DataFrame/Series results use Arrow IPC in mapped session files. Figure PNGs use session files for worker-to-host delivery.

The packages' APIs are available subject to runtime, input/output and resource limits. This does not guarantee that every library object can be displayed, every third-party usage is verified, or results match different package versions. Import standard-library/package modules explicitly when required.

Common operations targeted by the retained verified examples include pandas merge, grouping, aggregation, sorting and top-five selection, date-to-month grouping, and Matplotlib revenue bar charts. Other standard-library/NumPy/pandas APIs remain available but are not automatically verified workflows. A function catalog should link to the exact packaged library versions and distinguish available APIs from the app workflows actually tested. Native spreadsheet formulas have their own exact supported-function list from FormulaSubset; package installation UI is outside scope.

The [sales showcase](showcase/README.md) demonstrates joins, both top-five rankings, monthly and regional revenue charts, regional statistics and linked Dashboard KPIs. Its sheets are Customers, Products, Sales, Sales Analysis, Rankings, Charts and Dashboard. Completely empty source records are skipped; partial records and duplicate keys remain validation errors. Existing user scripts are never silently rewritten.

The [support SLA showcase](showcase/sla/README.md) uses `pandas.pivot_table()` for breach rate and 90th-percentile resolution time by team and month. Editing a priority target refreshes the spill, linked KPIs, and two Matplotlib charts in Python Automatic mode.

## Figure placement and resizing

Select a figure result and use **Place on Sheet**. A preview follows the pointer; switch sheets before placing, click for a default size, drag for a custom rectangle, or Escape to cancel. Charts can cover populated cells and overlap other charts. Covered cells retain their values. The selected chart appears in front with an accent outline and eight resize handles. Selection moves the caret to the source only when it is on the current sheet. Drag the body to move the chart or use a handle to resize it. Delete removes only that placement. There is no source caption or Remove Placement button.

The execution that creates a placed chart rasterizes it at the placement size and keeps the Figure for a later resize. On mouse release Matplotlib redraws the retained Figure into the new rectangle without rerunning the analysis. Standard subplot layout and crowded labels adapt on a disposable copy. The worker encodes the final pixels directly and keeps a final draw when the layout or artists can still change. During dragging or on redraw failure, the cached PNG fits without distortion and may leave margins. Custom fixed-position artists retain their authored layout. Errors are summarized in status with full details in the source cell's Python Errors panel.

Private session snapshots restore live Figures after RAM eviction or worker restart. A fresh Automatic session prepares missing placed-chart sources through normal calculation. Resize waits for queued work. If the Figure is unavailable when a resize ends, the app requests consent if needed, runs the source and its prerequisites, and renders at the new size before saving the geometry. Native Save and reopen preserve the source, PNG and placement, while the serialized Figure lives in the session. See [Architecture](Architecture.md#floating-figures) for cache and identity rules.

Dashboard has no independent graph formula: its regional chart is linked to **Charts!E22**, also displayed on Charts. Multiple placements can share one source.

## Persistence, exchange and limits

Native `.exlt` preserves source, references, names/tables, last scalar/table/figure output, spill ownership and freshness. It does not preserve execution authorization. XLSX export is reviewed value snapshots with explicit stale/native-metadata/figure losses; it is not an executable Microsoft Python-in-Excel roundtrip. CSV remains excluded.

Retain existing limits unless measured and explicitly changed: model 10 sheets / 10,000 rows / 100 columns / 200,000 populated cells; current UI viewport A1:Z1000 per sheet; Python input/output tables 100,000 cells; preview 500 rows; source/captured logs 1 MiB; protocol frames 32 MiB; PNG 16 MiB and 16 million pixels; default execution timeout 30 seconds. A spill's headers/generated cells must also fit document limits. These are supported bounds, not measured performance guarantees or clean-machine installation evidence.

## Native formulas

The current allowlist is `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `IF`, `AND`, `OR`, `NOT`, `IFERROR`, `ABS`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `SUMIF`, `COUNTIF`, `LEFT`, `RIGHT`, `MID`, `LEN`, `TRIM`, `UPPER`, `LOWER`, `DATE`, `YEAR`, `MONTH`, `DAY`, `TODAY`, `VLOOKUP` and `SUBTOTAL`. The source authority is `ExcelLite.Core/FormulaSubset.cs`. VLOOKUP supports exact match only (FALSE/0); unsupported functions fail visibly rather than silently returning zero.

SUBTOTAL supports codes 1–11 and 101–111. Filtered rows are excluded; the 100-series also excludes manually hidden rows, and nested subtotals are not counted again. The total-row picker offers None, Average, Count, Count Numbers, Max, Min and Sum; Count counts nonempty cells, Count Numbers numeric cells. It generates data-only SUBTOTAL references without including the total itself.

## Date/time locales and editing

Date/Time locale choices change grid presentation and editing. Stored numeric serials and the existing locale-neutral ISO date/time values supplied to Python remain unchanged; Excel’s fictitious 1900-02-29 remains rejected as a Python date input.
