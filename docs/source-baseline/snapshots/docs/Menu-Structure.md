# Menu structure

The [Excel menu references](references/ExcelMenu) define approved command families. Omit unsupported commands and empty tabs; do not invent a new location. Before implementing an added or moved command whose placement is unclear, ask the user. All application top bars remain black in Dark/Light themes, with blue accents and native themed controls. The Font group includes the approved icons/split controls; there is no standalone Python ribbon tab.

| Surface | Group / path | Existing commands |
| --- | --- | --- |
| Title bar | Quick Access Toolbar | Undo, Redo |
| File | Backstage, reference order | Home, New, Open, Save, Save As, Export, Close; Options at bottom; Back returns to workbook |
| File → Home | New / Recent | Blank workbook and up to 10 recent native/imported workbooks |
| File → Options | General | Office Theme: Dark, Light, Follow system |
| File → Options | Formulas → Calculation options → Workbook Calculation | Automatic, Manual (mutually exclusive; OK applies, Cancel discards) |
| File → Options | Formulas → Python Calculation | Automatic, Manual; saved locally, with Automatic startup explained |
| Home | Clipboard | Paste dropdown: Paste, Paste Values, Paste Formulas; Cut; Copy |
| Home | Font | Editable family/size; Increase/Decrease Font Size; Bold, Italic; Underline/Double Underline; Borders (No/Bottom/All/Outside); Fill Color; Font Color |
| Home | Alignment | Top/Middle/Bottom Align above Align Left/Center/Align Right; Wrap Text; Merge & Center split control with Merge & Center, Merge Across, Merge Cells, Unmerge Cells |
| Home | Number | Selected-style menu: General, Number, Currency, Accounting, Short Date, Long Date, Time, Percentage, Fraction, Scientific, Text, More Number Formats...; Currency dropdown: Kč Czech, $ English (United States), € Euro (123 €), £ English (United Kingdom), € Slovak, zł Polish; Percentage; Comma Style; Increase/Decrease Decimal |
| Home | Cells → Insert | Insert Cells...; Insert Sheet Rows; Insert Sheet Columns; Insert Sheet |
| Home | Cells → Delete | Delete Cells...; Delete Sheet Rows; Delete Sheet Columns; Delete Sheet |
| Home | Cells → Format → Organize Sheets | Rename Sheet; Move or Copy Sheet... |
| Home | Cells → Format → Cell Size | Row Height...; AutoFit Row Height; Column Width...; AutoFit Column Width |
| Home | Cells → Format → Visibility → Hide & Unhide | Hide Rows; Hide Columns; Unhide Rows; Unhide Columns |
| Sheet-tab right-click | Reference order, existing commands | Insert...; Delete; Rename; Move or Copy... (target the clicked sheet) |
| Bottom sheet strip | Sheet navigation | Previous/Next sheet; named active-sheet tabs; New sheet (+) |
| Home | Editing | AutoSum split control: Sum, Average, Count Numbers, Max, Min; Fill dropdown: Down, Right, Up, Left; Clear dropdown: Clear All, Clear Formats, Clear Contents |
| Home | Editing → Sort & Filter | Sort A to Z; Sort Z to A; Custom Sort...; Filter; Clear (filter criteria); Reapply |
| Home | Editing → Find & Select | Find...; Replace... |
| Formulas | Function Library | Insert Function (implemented functions only); AutoSum with the same five options |
| Formulas | Formula Auditing | Show Formulas (checked when enabled) |
| Formulas | Defined Names | Name Manager; Define Name → Define Name... (existing named-range dialog) |
| Formulas | Calculation | Calculation Options → Automatic / Manual (same workbook mode and saved preference as Options); Calculate Now (workbook); Calculate Sheet |
| Data | Sort & Filter | A to Z; Z to A; Sort; Filter; Clear; Reapply |
| Filter-range header dropdown | Sort / selected values | Type-appropriate ascending/descending sort; Clear Filter From "header"; Search; (Select All); value checkboxes; OK / Cancel |
| Formulas | Python | Python (checked when its panel is shown) |
| View | Show | Formula Bar; Gridlines; Headings (checked while visible) |
| View | Zoom | Zoom... (10–400%); 100%; Zoom to Selection |
| View | Window → Freeze Panes | Freeze Panes / Unfreeze Panes; Freeze Top Row; Freeze First Column |
| Status bar | Selection statistics / zoom | Visible selected Sum, Average, Count; − / percentage / + |
| Python panel | Execution / results | Recalculate (Ctrl+Enter); Hide; Python options for Automatic Python, Commit Python edit and Discard uncommitted Python edit; Copy / Insert as Values beside the table result preview |
| Python panel | Figure results | Save Image and Place on Sheet beside the figure preview; select a chart and press Delete to remove that placement; Result / Output / Errors tabs below the editor |

## File and settings

Backstage order is Home, New, Open, Save, Save As, Export, Close, with Options at the bottom. Recent belongs to Home. XLSX import is File/Open; native Save/Save As uses `.exlt`, with separate reviewed XLSX Export. CSV commands are absent. Options edits are drafts until OK; Cancel discards them. Python Automatic/Manual is separate from native calculation and explains adoption-time execution.

## Worksheet and dialogs

Row/column right-click preserves a selection when clicked within it and exposes whole-axis Insert/Delete. Cell Insert/Delete dialogs offer shift right/down or left/up, and entire row/column choices; each mutation undoes once. Double-click a sheet tab, use its Rename context command, or choose Rename Sheet from Home → Cells → Format to edit the tab name in place. The full name is selected. Enter or leaving the editor applies a valid name, and Escape cancels. Move or Copy has a Before sheet list, end position and Create a copy.

Header-body cursors use a thick right arrow for rows and down arrow for columns; edges use the supplied double-arrow/bar resize cursor. Edge drag sizes selected axes together. Double-click an edge to AutoFit width or height, or use the matching Home → Cells → Format command for the selected columns or rows.

Find/Replace is under the Editing group in the [Home reference](references/ExcelMenu/ExcelHome.png). The dialog uses native tabs, editable Find what/Replace with combos, Options expansion, Sheet/Workbook scope, row/column order, Values/Formulas, case/whole-cell matching and result navigation. Each editable combo remembers 15 distinct values, newest first, separately and case-sensitively in local preferences. Actions record history, including empty replacement. Native text submission retains drafts on focus loss and dropdown expansion.

Font controls follow the two-row [Home reference](references/ExcelMenu/ExcelHome.png). Underline exposes single/double only; colors offer Automatic/No Fill, Theme/Standard Colors and More Colors. Sort supports a single cell-value key, case sensitivity and top-to-bottom only. Partial selections offer expansion before sorting. Plain ranges infer a header when text in the first row precedes a numeric or Boolean value in the same column. Custom Sort allows an explicit header choice. Filter is checked while enabled and header buttons display sort/funnel state.

Each column’s dropdown lists the values in rows that pass the other columns’ filters. The current column’s filter determines which of those choices are checked. Search narrows the checklist and **(Select All)** acts on the search results. **Clear Filter From** removes that column’s criterion.

Ctrl+click selects separate cells. Ribbon controls for font, fill color, alignment, and number formats use every selected area, as do Clear and status-bar statistics. Delete clears the selected contents in one Undo step and retains table definitions for a selection with multiple areas. Clipboard, Fill, Borders, Merge, Format Cells and other commands that require a rectangle prompt for one contiguous range.

Freeze Panes keeps the rows above and columns left of the active cell visible while scrolling. For example, select D5 to freeze rows 1–4 and columns A–C. At A1 the split defaults near the middle of the viewport. Freeze Top Row and Freeze First Column apply their respective worksheet-edge shortcuts. Unfreeze Panes replaces the primary command while the sheet is frozen. Each sheet has its own settings, and Undo and Redo include freeze changes.

## Tables

Insert → Tables → Table uses a name, explicit range, **My table has headers** and the **...** worksheet range picker. Formulas → Defined Names → Define Name / Name Manager follows [New Name](references/ExcelMenu/ExcelNewName.png) and [Name Manager](references/ExcelMenu/ExcelNameManager.png), with fixed Workbook scope. Name Manager can edit names/comments/extents; its Delete is for named ranges and is disabled for tables. Worksheet Delete removes a fully selected table and its selected contents in one step; Clear All, Clear Formats, Clear Contents and Backspace retain the table definition. Clear All removes the selected cell contents and cell-level formatting in one Undo step.

| Contextual surface | Controls |
| --- | --- |
| Table Design → Properties | Table Name; Resize Table |
| Table Design → Table Style Options | Header Row; Total Row; Banded Rows; First Column; Last Column; Banded Columns; Filter Button |
| Table Design → Table Styles | Blue; Green; Orange; Purple |
| Selected total-row cell | None; Average; Count; Count Numbers; Max; Min; Sum |

The [circled ribbon](references/ExcelMenu/D58-Table-Design.png) and [total dropdown](references/ExcelMenu/D58-Table-Total-Dropdown.png) establish this subset. No slicers, external-data tools, PivotTables or style customization menu.

New table insertion activates Table Design. Leaving a table hides the contextual tab and restores the last ordinary tab. Re-entering activates Table Design only if it was the last active tab for that table; otherwise retain the current ordinary tab. Movement within a table must not steal tab selection. This memory is session-local.

Header Row off caches column labels, clears their former cells while preserving formatting, and excludes that row from table extent/styling. Header on restores the labels above the data only when safe; it refuses to overwrite changed neighboring values. Filter Button is disabled while the header is hidden and restored with it. Total Row is separate from data, refuses occupied destinations and uses SUBTOTAL for selected functions. Resize extents include enabled header/total rows. Styles preserve explicit cell formatting; tables have a direct lower-right resize grip.

## Python and figures

Create Python by entering `=PY`/`=PY()` in a cell; anchors and owned spill children open their source. Ordinary cell selection clears the panel. The removed block selector/Create/Apply bindings controls stay removed; legacy metadata remains internally compatible. Put the Monaco caret inside an `xl` literal to pick a range, including another sheet.

Table results spill automatically. Copy and Insert as Values use the complete last successful result, with freshness labeling; insertion asks for its destination after invocation through the retained-draft worksheet picker. Result/Output/Errors remain below Monaco.

Place on Sheet enters direct pointer-preview mode: switch sheets, click for default size or drag a rectangle, Escape to cancel. Drag the chart body to move it over cells or other charts. Eight handles resize width/height independently and redraw on release. Selection uses the dialog-accent outline and moves the source caret only if it is on the current sheet. Delete removes the placement alone. No source caption or Remove Placement button. Save Image exports the figure stored on the block. [Python capabilities](Python-Capabilities.md) describes calculation and redraw availability.

## Quick Access and Clipboard

The top-left title area shows the app icon, followed by icon-only Save, Undo and Redo before the workbook title. Save uses the same native `.exlt` save/picker path as File > Save and Ctrl+S. Home > Clipboard has a large Paste action with a separate options arrow and stacked Cut, Copy and Format Painter controls. Single-click Format Painter applies the selected cell styles once on destination click/drag release; double-click keeps it active until Escape or another click on the brush. A single destination cell expands to the source rectangle; larger destinations repeat its style pattern. Values, formulas, Python content and dimensions are retained. Painting works across sheets within the same workbook and each application undoes once. Switching workbooks cancels the brush. Only supported authored cell styles are copied; conditional formatting and merge geometry are not copied by Format Painter.

Find/Replace closes on Escape from either tab, including while an editable history field or dropdown has focus; the key dismisses the dialog without running Find Next or Replace.

## Alignment and merged cells

The six alignment buttons use Excel line silhouettes in the existing native theme. Selecting an active explicit alignment returns it to General (horizontal) or Bottom (vertical). Align Left and Align Right clear the stored indent when applied. Increase/Decrease Indent and the Alignment tab set indentation explicitly. Orientation follows Excel menu order: Angle Counterclockwise (+45°), Angle Clockwise (−45°), Vertical Text, Rotate Text Up (+90°), Rotate Text Down (−90°), separator, Format Cell Alignment. Each orientation preset toggles back to horizontal when selected again; choosing another replaces it. Increase/Decrease Indent use 0–250 levels and establish Left alignment unless Left, Right or Distributed is already set. The group launcher and Format Cell Alignment open the Alignment tab. Wrap Text uses a return-arrow line icon with its full label. Vertical alignment defaults to Bottom; horizontal General remains implicit until chosen. Merge & Center centers one rectangle; Merge Across creates one merge per row; Merge Cells retains alignment. The first populated cell in row order, from left to right within each row, supplies the upper-left anchor. Merging prompts before discarding any other populated cells. The split action repeats the last merge mode and unmerges when checked. Unmerge keeps the upper-left value and existing alignment. Each action undoes once. Tables, Python anchors and spill results cannot be merged.

AutoSum accepts one selected row or column. If its final cell is blank, that cell receives the chosen aggregate over the preceding selected cells; otherwise the formula goes in the blank cell immediately after the selection. A single blank cell infers contiguous numeric cells above, then to its left. Count Numbers inserts COUNT. The operation is one Undo step and rejects occupied or out-of-view destinations.

Insert Function starts a draft for the active cell, with the caret inside the argument list. When the formula bar is hidden, it opens the in-cell editor. Enter commits and Escape cancels.

Fill Down/Up copies the first/last selected row through the selection; Fill Right/Left copies the first/last selected column. Each action translates relative formula references, copies cell formatting and blanks, and creates one Undo step. Merged selections require unmerge first.

Home Sort & Filter follows Data command availability: Clear requires active criteria and Reapply requires a filter range. Unavailable items remain visible with muted labels and icons. A separator divides sorting from filtering.

## Number formatting

Home > Number displays the active cell’s format style. Its menu offers General, Number, Currency, Accounting, Short Date, Long Date, Time, Percentage, Fraction, Scientific, Text and More Number Formats. The toolbar includes Currency, Percentage, Comma Style and decimal controls. More Number Formats opens Format Cells with Number, Alignment, Font, Border and Fill tabs: General, Number, Currency, Accounting, Date, Time, Percentage, Fraction, Scientific, Text, Special and Custom. The dialog previews the active value; OK formats the captured selection in one Undo step and Cancel discards edits. Decimal places, grouping, currency symbol, negative style and type/code controls appear by category. Custom types include codes already used in the workbook.

## Date/time locales and editing

Date and Time categories include a Locale (location) selector. Type lists show localized examples; changing locale keeps the selected type position and updates the sample. Locale tags are saved in the format code. F2/double-click and the formula bar show localized dates, with the hidden calendar date included for time-formatted serials of at least one day (for example `10.01.1900 0:00:00`).

## Cell editing and Special formats

Cell right-click includes Format Cells... with a format icon, opening the same Number/Alignment/Font/Border/Fill dialog for the captured selection. OK and Cancel return keyboard focus to the worksheet so F2 works immediately. Special includes Locale (location) and named postal/telephone/identity formats for Czech, Slovak, US, UK, German, French and Polish locales.

The cell menu includes clipboard, insert and delete, number format, clear, and history commands. Heading menus add dimension sizing, AutoFit, hide and unhide, and row or column insertion and deletion.

The Formulas tab’s **Python** button opens the Monaco panel. **Recalculate** runs the workbook from the panel or through Ctrl+Enter in Monaco. **Hide** closes the panel and unchecks the ribbon button. **Python options** contains Automatic Python, Commit Python edit, and Discard uncommitted Python edit. The panel retains the table and figure result actions.

## Font and Fill formatting

Format Cells now has Number, Alignment, Font, Border and Fill tabs. Font offers installed/typed family, Regular/Italic/Bold/Bold Italic, 6–72 pt size, five underline choices (including accounting), Automatic/custom RGB color, Normal font, strikethrough and mutually exclusive superscript/subscript. Fill offers No Color, theme/standard swatches, More Colors with inline RGB/Hex and Apply, pattern color, a graphical pattern palette and two-color Fill Effects. Gradient shading includes Horizontal, Vertical, Diagonal up/down, From corner/center as radio options, alongside clickable variant thumbnails (three linear, four corner, one center) and a final sample below. All tabs share one OK/Cancel draft and one Undo transaction.

## Format Cells control layout

Format Cells uses compact bordered Number/Alignment/Font/Border/Fill tabs. Number has an edit-height read-only Sample, compact decimal/symbol rows, a persistent negative-number ListBox with real red/black examples and named fraction choices. The separator label, numeric samples and chosen numeric format locale follow Windows regional settings. Font uses Color: with Automatic or a color swatch, a Normal font checkbox, and captioned Effects/Preview frames. Color dropdowns contain Theme Colors, Standard Colors and inline More Colors with a picker swatch, #RRGGBB field and Apply.


Border sits immediately after Font. It offers None plus thirteen line styles, Automatic/theme/standard/custom colors through the same selector as Fill Pattern Color, None/Outline/Inside presets, a clickable preview and individual outer/inside/diagonal controls. Inside controls are disabled when the selection has no boundary in that direction. Style and color changes update enabled, uniform preview edges; mixed edges remain unchanged until explicitly targeted. Color changes preserve line styles and style changes preserve colors.

Border uses Excel’s graphic None/Outline/Inside presets and compact edge/diagonal icon buttons along the left and bottom of the preview. All fourteen style choices fit in two columns without scrolling. The preview displays one Text, two horizontal Text labels, two vertical labels or four labels according to selection shape; a single merged cell uses one label. Its canvas matches the worksheet without gridlines, and Automatic ink follows the worksheet foreground.

Home > Font > Borders follows Excel’s thirteen presets (Bottom/Top/Left/Right; No/All/Outside/Thick Outside; Bottom Double/Thick Bottom; Top and Bottom/Thick Bottom/Double Bottom). Draw Borders includes Draw Border, Draw Border Grid, Erase Border, Line Color (Automatic, theme/standard/recent colors and shared More Colors picker) and all fourteen Line Style choices. More Borders opens Format Cells on Border. Choosing ink/style starts or updates the drawing tool. Draw Border supports individual edges and rectangular outlines; Draw Border Grid includes inside lines. Erase Border removes clicked or left-button-dragged segments. Tools persist across strokes and exit with Escape, normal editing, range picking or a workbook/sheet change. Each completed stroke is one Undo step.

Borders is a split button: the main icon repeats the remembered preset or activates the remembered Draw Border/Draw Border Grid/Erase Border tool; only the arrow opens the menu. Escape ends the active tool and restores the normal cursor while retaining its icon/action for the next primary click. Straight horizontal/vertical Draw Border gestures preview one line; leaving that grid edge switches to a rectangular outline.

Format Cells orders its tabs Number, Alignment, Font, Border, Fill. Alignment offers all eight horizontal and five vertical choices, Indent, Justify distributed, Wrap text, Shrink to fit, Merge cells, Context/Left-to-Right/Right-to-Left direction, stacked text and −90°…90° orientation with a pointer-adjustable dial and live preview while typing degrees. Text control and Right-to-left share one row beneath the alignment fields. Edits remain in the dialog draft until OK. Merge warnings precede applying the complete proposal; Cancel preserves the workbook, and combined formatting/merge undoes once. Unedited fields preserve mixed selections.

Home > Font uses the same sorted installed-font family catalog as Format Cells > Font; both retain custom family entry. Its corner-arrow dialog launcher opens Format Cells directly on Font.

Cells > Format groups only supported actions: Cell Size (Row Height, AutoFit Row Height, Column Width, AutoFit Column Width), Visibility (Hide/Unhide Rows/Columns), Organize Sheets (Rename, Move or Copy), then Format Cells. Default Width, Tab Color and protection commands are absent. Number format previews are read-only; selecting a preset applies the existing code, and the Number footer launcher opens Format Cells on Number.
