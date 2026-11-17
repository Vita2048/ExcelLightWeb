# Desktop command inventory

Each row captures a menu group from the desktop baseline and links it to a web parity requirement. The full source snapshot retains context, shortcuts and dialog behavior. See [Feature parity](../Feature-Parity.md).

| Surface | Group or path | Commands and controls | Parity |
| --- | --- | --- | --- |
| Title bar | Quick Access Toolbar | Undo, Redo | F02 |
| File | Backstage, reference order | Home, New, Open, Save, Save As, Export, Close; Options at bottom; Back returns to workbook | F01 |
| File → Home | New / Recent | Blank workbook and up to 10 recent native/imported workbooks | F01 |
| File → Options | General | Office Theme: Dark, Light, Follow system | F03 |
| File → Options | Formulas → Calculation options → Workbook Calculation | Automatic, Manual (mutually exclusive; OK applies, Cancel discards) | F09 |
| File → Options | Formulas → Python Calculation | Automatic, Manual; saved locally, with Automatic startup explained | F29 |
| Home | Clipboard | Paste dropdown: Paste, Paste Values, Paste Formulas; Cut; Copy | F05 |
| Home | Font | Editable family/size; Increase/Decrease Font Size; Bold, Italic; Underline/Double Underline; Borders (No/Bottom/All/Outside); Fill Color; Font Color | F11/F12 |
| Home | Alignment | Top/Middle/Bottom Align above Align Left/Center/Align Right; Wrap Text; Merge & Center split control with Merge & Center, Merge Across, Merge Cells, Unmerge Cells | F13/F14 |
| Home | Number | Selected-style menu: General, Number, Currency, Accounting, Short Date, Long Date, Time, Percentage, Fraction, Scientific, Text, More Number Formats...; Currency dropdown: Kč Czech, $ English (United States), € Euro (123 €), £ English (United Kingdom), € Slovak, zł Polish; Percentage; Comma Style; Increase/Decrease Decimal | F10 |
| Home | Cells → Insert | Insert Cells...; Insert Sheet Rows; Insert Sheet Columns; Insert Sheet | F06 |
| Home | Cells → Delete | Delete Cells...; Delete Sheet Rows; Delete Sheet Columns; Delete Sheet | F06 |
| Home | Cells → Format → Organize Sheets | Rename Sheet; Move or Copy Sheet... | F06 |
| Home | Cells → Format → Cell Size | Row Height...; AutoFit Row Height; Column Width...; AutoFit Column Width | F07 |
| Home | Cells → Format → Visibility → Hide & Unhide | Hide Rows; Hide Columns; Unhide Rows; Unhide Columns | F07 |
| Sheet-tab right-click | Reference order, existing commands | Insert...; Delete; Rename; Move or Copy... (target the clicked sheet) | F06 |
| Bottom sheet strip | Sheet navigation | Previous/Next sheet; named active-sheet tabs; New sheet (+) | F06 |
| Home | Editing | AutoSum split control: Sum, Average, Count Numbers, Max, Min; Fill dropdown: Down, Right, Up, Left; Clear dropdown: Clear All, Clear Formats, Clear Contents | F05 |
| Home | Editing → Sort & Filter | Sort A to Z; Sort Z to A; Custom Sort...; Filter; Clear (filter criteria); Reapply | F17 |
| Home | Editing → Find & Select | Find...; Replace... | F19 |
| Formulas | Function Library | Insert Function (implemented functions only); AutoSum with the same five options | F09 |
| Formulas | Formula Auditing | Show Formulas (checked when enabled) | F20 |
| Formulas | Defined Names | Name Manager; Define Name → Define Name... (existing named-range dialog) | F15 |
| Formulas | Calculation | Calculation Options → Automatic / Manual (same workbook mode and saved preference as Options); Calculate Now (workbook); Calculate Sheet | F09 |
| Data | Sort & Filter | A to Z; Z to A; Sort; Filter; Clear; Reapply | F17 |
| Filter-range header dropdown | Sort / selected values | Type-appropriate ascending/descending sort; Clear Filter From "header"; Search; (Select All); value checkboxes; OK / Cancel | F17 |
| Formulas | Python | Python (checked when its panel is shown) | F25 |
| View | Show | Formula Bar; Gridlines; Headings (checked while visible) | F20 |
| View | Zoom | Zoom... (10–400%); 100%; Zoom to Selection | F20 |
| View | Window → Freeze Panes | Freeze Panes / Unfreeze Panes; Freeze Top Row; Freeze First Column | F20 |
| Status bar | Selection statistics / zoom | Visible selected Sum, Average, Count; − / percentage / + | F18/F20 |
| Python panel | Execution / results | Recalculate (Ctrl+Enter); Hide; Python options for Automatic Python, Commit Python edit and Discard uncommitted Python edit; Copy / Insert as Values beside the table result preview | F25/F28/F29 |
| Python panel | Figure results | Save Image and Place on Sheet beside the figure preview; select a chart and press Delete to remove that placement; Result / Output / Errors tabs below the editor | F31 |
| Table Design and totals | Table Design → Properties | Table Name; Resize Table | F16 |
| Table Design and totals | Table Design → Table Style Options | Header Row; Total Row; Banded Rows; First Column; Last Column; Banded Columns; Filter Button | F16 |
| Table Design and totals | Table Design → Table Styles | Blue; Green; Orange; Purple | F16 |
| Table Design and totals | Selected total-row cell | None; Average; Count; Count Numbers; Max; Min; Sum | F16 |

## Additional command and interaction groups

| Surface | Commands and behavior | Parity |
| --- | --- | --- |
| Quick Access and grid | Save, Ctrl+S, inline edit, formula bar edit, Ctrl+click areas, contextual selection, clipboard fallback | F01 / F04 / F05 |
| Row, column and cell context menus | Insert/Delete axes and shift lanes, resize drag, double-click AutoFit, Hide/Unhide, Format Cells | F06 / F07 / F10–F14 |
| Clipboard | Format Painter once or locked, across sheets, Escape to release | F11 |
| Format Cells dialog | Number, Alignment, Font, Border and Fill. Includes draft preview, mixed-field preservation, locale and custom colors | F10–F14 |
| Alignment | Orientation presets, stacked text, increase/decrease indent, format launcher, merge discard confirmation | F13 / F14 |
| Border tools | Draw Border, Draw Border Grid, Erase Border, Line Style, Line Color, More Borders, remembered action | F12 |
| Insert and table design | Insert Table, range picker, resize grip, headers, totals, contextual ribbon, Delete vs Clear behavior | F15 / F16 |
| Search dialog | Find Next, Find All, Replace, Replace All, Options, histories and Escape | F19 |
| Python editor/results | =PY / =PY(), cross-sheet xl range picking, source commit/discard, Stop/Restart lifecycle and result/status tabs | F25–F30 |
| Chart gestures | Pointer preview, Escape, cross-sheet placement, drag, eight-handle resize, delete placement | F31 |

For shortcuts, enabled states, focus and exact command ordering, use the archived [Menu structure](snapshots/docs/Menu-Structure.md) and [Design guidelines](snapshots/docs/Design-Guidelines.md).
