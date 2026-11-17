# Excel-Lite sales showcase

For a second Python workbook, see the [support ticket SLA showcase](sla/README.md).

This is a local Windows showcase, using the supplied 20 customers, 20 products and 25 orders. Start the current Excel-Lite developer build, choose **File → Open**, and open `Showcase.exlt` in this folder. No source-data download is needed to run the saved workbook. Installation on a clean machine is not qualified by this delivery.

## Five-minute walkthrough

1. Select **Sales Analysis**. The Python cell at A3 joins the three source tables and spills all 25 transactions plus a Total row: **333 units / $63,391.00**. Select the anchor or a generated cell to inspect its script. The other source sheets retain the supplied data; Sales row 29 is its table total; data-only Python references exclude it.
2. Select **Rankings**. Both top-five lists are live analyses. Acme Corporation leads customer spend at **$11,000.95**; Keystone Consulting is fourth. UltraBook Pro 15 leads product revenue at **$17,999.88**.
3. Select **Charts**. Monthly revenue is **$22,509.45 / $20,925.04 / $19,956.51** for January–March 2026. The region plot excludes the Total row. Both plots appear inside the sheet; select their owner cells E3/E22 for the Python preview and **Save Image**. This folder also contains the original PNG results.
4. Select **Dashboard**. Four linked KPIs show **$63,391.00**, **333**, **25** and **$2,535.64**. Regional statistics and top customers link to their analysis sheets. The regional image shares its owner with the Charts sheet.
5. In **File → Options → Formulas**, choose Python **Automatic**. In **Sales!D4**, change order 1001's quantity from **5 to 6** and commit. Without pressing Run, revenue becomes **$64,890.99**, units **334**, average **$2,595.64**, North revenue **$17,287.94** and January revenue **$24,009.44**. Restore D4 to **5**, let refresh finish, then save if desired.

Use a maximized window and hide the Python panel when presenting. Start with 65% zoom for the wide Sales Analysis report and 75% for Charts/Dashboard; adjust to your display. Zoom and the active sheet are session settings, so select the desired sheet after opening. Titles wrap within supported cells; merged cells are not required. The application top bars remain black.

## Table Design and cleared records

Select any cell in Customers, Products or Sales to activate **Table Design**. The tab disappears outside a table and restores the previous ordinary tab. Returning to a table activates Table Design only if it was that table's last active tab; otherwise your current tab stays selected. Newly inserted tables activate Table Design. Use **Table Name / Resize Table**, the seven style-option checkboxes and **Blue / Green / Orange / Purple** styles. Explicit cell formatting remains authoritative. Header Row off clears the former header cells, excludes that row from the table and caches its column labels for Python and later restoration; filter buttons are disabled until the header returns. Total Row adds/removes a separate summary and refuses to overwrite occupied cells below the table. Resize ranges include only table rows: the header when enabled, data and any enabled total. Reattaching the header refuses to overwrite changed worksheet cells above the table.

Sales has an Orange style and its existing row 29 is now the table total. Select **D29** for **None / Average / Count / Count Numbers / Max / Min / Sum**. Count counts nonempty values; Count Numbers counts numeric values. The generated SUBTOTAL updates when rows are filtered or manually hidden. Filtering affects this worksheet subtotal; Python analyses read all records through `[#Data]`. Only one active filter range per sheet is supported.

To reproduce the fixed bug on a disposable copy, select **Sales!A4:F4** and press **Delete**. Completely empty records are skipped. Revenue becomes **$55,891.05**, units **328**, orders **24**, average **$2,328.79**, North revenue **$8,288.00**, and January revenue **$15,009.50**. The shorter report and Dashboard stay consistent. Partial records and duplicate IDs still report errors. Reopen the original workbook without saving this experiment to restore baseline. This fix is in the updated showcase scripts/examples; existing user scripts are never silently rewritten.

## Execution, saving and recovery

- Python **Automatic** is a local preference: opening a workbook enables the interpreter and refreshes stale outputs. Current caches are skipped except placed chart sources missing live/archive rendering state, which Automatic prepares for resizing. Only open trusted workbooks; Python runs under your Windows account and can access files/network.
- Python **Manual** opens cached output without executing scripts and resets enablement on document adoption. Use the existing panel's **Enable Python**, **Run Cell** or **Run All** when ready. Selecting cells, Undo/Redo and Restart never execute scripts.
- **Stop** cancels queued/active work. **Restart** resets the interpreter without running cells. Correct an error and run again. Generated spill children are read-only; edit their source anchor instead. A spill obstruction must be moved or cleared before retrying; existing data is never overwritten.
- Save through **File → Save / Save As** as `.exlt`. The native workbook retains scripts, names/tables, outputs, formatting and linked images. Work on a copy for failure experiments. XLSX is a supported-subset snapshot with reported losses; this is not executable Python-in-Excel exchange or a full-fidelity XLSX round trip.

`Expected-values.json` contains independent source-derived expectations. The original input is [showcase_data.xlsx](showcase_data.xlsx), SHA-256 `274acbb92a5300368da7a1a887fa3789c2c97e628931c6790b1fc95c697108a3`. [Development](../Development.md#verification-boundaries) records the scope of observed verification, including the distinction between the saved workbook and later chart-runtime refinements.

This bounded delivery does not qualify clean offline installation, large-workbook performance, broad accessibility/DPI behavior, web hosting or a production release. The latest native dropdown/chart interaction has not been verified by a full UI walkthrough.


## Python panel, figure placement and search histories

The Python panel clears when you select an ordinary cell or native formula. Select a Python cell or one of its spill cells to see its source/results again. Pending edits are flushed before clearing; selecting a cell does not run code. To insert independent values, select the result, click **Insert as Values**, then choose the destination in the worksheet picker.

**Place on Sheet** enters direct placement mode. A translucent preview follows the mouse. Switch sheets before placement if needed, then click for a default size or drag to choose a size; **Esc** cancels. Select a chart to show its outline and eight resize handles. Drag its body to move it freely over cells, or drag a handle to resize. **Delete** removes only the selected placement; its Python source and other placements remain. Clicking a worksheet cell deselects the chart.

The Dashboard's regional chart comes from **Charts!E22** and is linked into **Dashboard!F7:L25**, reusing the same figure also displayed at Charts!F22:L38. There is no separate Python cell on Dashboard. Select Charts!E22 to inspect its code. Selecting a chart moves the caret to its source only when that source is on the current sheet; selecting the Dashboard chart stays on Dashboard.

**Find what** and **Replace with** are editable combo boxes. Their separate 15-value histories persist locally, newest first; values are recorded when a Find/Replace action is used.

Chart placement and resizing support independent width and height. After mouse release, Matplotlib redraws a fresh copy of the source Figure at the new rectangle dimensions, recalculating plot layout and crowded tick labels without rerunning analysis. Each placement has its own rendered image. Automatic mode prepares missing live Figures for placed charts when opening/calculating the workbook, even when their saved PNGs are current. Resizing waits for scheduled preparation; session snapshots recover worker restarts/cache eviction. Manual mode still needs **Run Cell** when no live or archived Figure exists; the status bar explains unavailable redraws.

## Folder contents

- [Showcase.exlt](Showcase.exlt): saved seven-sheet native workbook.
- [Specification](Specification.md): exact layouts, schemas and baseline values.
- [Python examples](SHOWCASE-EXAMPLES.md): six independent scripts used in the saved workbook.
- `python/`: three packaged chart/statistics examples, linked into the app output as `Assets/PythonExamples`.
- [showcase_data.xlsx](showcase_data.xlsx): unchanged source workbook.
- [Expected-values.json](Expected-values.json): independent baseline expectations.
- [Monthly-Revenue.png](Monthly-Revenue.png) and [Revenue-by-Region.png](Revenue-by-Region.png): original chart results.
- `tools/`: independent source inspector, expectation generator and native artifact verifier.

## Tools

Run from the repository root with a suitable Python interpreter. The prepared private runtime can be used in place of `python`. These commands are reference instructions; running app harnesses/tests still follows the user's authorized verification scope.

```powershell
python docs/showcase/tools/inspect_showcase_data.py
python docs/showcase/tools/showcase_expectations.py docs/showcase/Expected-values.json
python docs/showcase/tools/verify_showcase.py docs/showcase/Showcase.exlt
```

The inspector derives Decimal expectations directly from the XLSX. The verifier reads the native ZIP without executing workbook Python. Its baseline assertions check values, preserved source cells, tables/scripts, Dashboard links and original PNGs; they do not qualify mouse resizing or arbitrary edited workbooks.
