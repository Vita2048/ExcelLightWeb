# Excel-Lite design guidelines

Current visual and interaction contracts. [Menu structure](Menu-Structure.md) defines command locations; [Python capabilities](Python-Capabilities.md) defines execution/results. These rules are not evidence that every accessibility, scaling or visual state has been verified; see [Development](Development.md#verification-boundaries).

## 1. Design intent

Make a spreadsheet workspace that feels precise, capable, and comfortable for long sessions. Use the Fluent surfaces, controls, blue selection, and spacing in `docs/references/UI3Example` as the starting point. Preserve spreadsheet density and familiarity while giving Python a clear, integrated workspace.

The grid is visually dominant. A slim command area provides common actions without a full-size Excel ribbon. Avoid marketing-style cards, oversized headings, saturated backgrounds, and persistent introductory panels once a workbook is open. Use subtle layering, confident typography, consistent icons, and excellent interaction states to make the app feel professional.

**User-required correction:** do not copy the white window title strip from the reference screenshots. The window title bar, command tab strip, toolbar, and formula-bar container use black backgrounds (`#000000`). This applies in both Dark and Light themes; Light changes the sheet and supporting panels, not the top bars. Caption buttons, labels, and icons must be readable on black, including in the inactive window. Windows High Contrast may override colors for accessibility.

`docs/references/ExcelMenu` defines familiar command families. It does not require matching Excel's ribbon height, icon positions, green branding, or unsupported buttons.

The selected ribbon tab keeps its blue underline through hover, press and keyboard focus. The button's native interaction states style the surface independently of that selection indicator.

## 2. Tokens and resources

Define semantic resources once and share their resolved values with the grid and Monaco. Do not scatter hardcoded colors through XAML, templates, and JavaScript. Values below are starting design choices; validate actual rendered contrast and Windows High Contrast behavior before release.

### Color

| Token | Dark | Light | Use |
| --- | --- | --- | --- |
| TopBarSurface | `#000000` | `#000000` | Title bar, command tabs, toolbar, formula-bar container |
| TopBarText | `#F5F5F5` | `#F5F5F5` | Labels, icons, and window caption buttons on top bars |
| TopBarField | `#242424` | `#242424` | Editable fields and selected/hover surfaces within black top bars |
| Canvas | `#1F1F1F` | `#F5F5F5` | Window workspace |
| SheetSurface | `#202020` | `#FFFFFF` | Unformatted cell background |
| PanelSurface | `#2B2B2B` | `#FFFFFF` | Drawer, dialogs, command surface |
| RaisedSurface | `#333333` | `#F0F0F0` | Hover and selected tab surface |
| TextPrimary | `#F5F5F5` | `#1B1B1B` | Content and primary labels |
| TextSecondary | `#C8C8C8` | `#595959` | Metadata, secondary labels |
| BorderSubtle | `#444444` | `#D5D5D5` | Dividers; not the sole control boundary |
| Gridline | `#383838` | `#E4E4E4` | Quiet sheet structure |
| Accent | `#60CDFF` | `#0067C0` | Focus, active selection, primary action |
| OnAccent | `#101010` | `#FFFFFF` | Text on filled primary buttons |
| SelectionFill | Accent at 14% | Accent at 10% | Range tint; keep cell content readable |
| Error | `#FF99A4` | `#B42318` | Failed execution / validation |
| Warning | `#FFD37A` | `#805400` | Stale / unverified result |
| Success | `#8BD6A3` | `#226B38` | Completed operation indicator |

Theme changes alter supporting application surfaces and unformatted defaults; the top bars remain black. Fields within them use TopBarField with TopBarText, and disabled/focus states must remain legible. Explicit user cell fills/text colors are document data and must not be rewritten by switching themes. Test high-contrast cell formatting in both themes; provide Reset Formatting instead of silently altering saved colors.

Use system high-contrast resources when active. Normal text target is at least 4.5:1 contrast; large text and meaningful non-text controls target at least 3:1. Color alone never distinguishes failure, selection, or staleness: include text, icon, border, or accessible state.

Dialog and popup controls use the same application palette, independent of the Windows desktop accent. Default action buttons use Accent/OnAccent; checked radio rings and checkbox backgrounds/glyphs use the corresponding Accent/OnAccent colors. Selected list/combo items and menu hover/press states use a readable blue-tinted surface; preserve focus, checked state, disabled-state distinctions and keyboard operation. Apply shared semantic resources at popup boundaries as needed; changing a resource dictionary alone is not visual verification. Check the rendered controls in Dark and Light. High Contrast retains the system palette.

Worksheet selectors use standard WinUI 3 TabView/TabViewItem controls at the bottom. The approved minimal orientation layout change turns the native strip toward the sheet while keeping header text upright; no custom drawing. Keep the native add/overflow behavior and context menu. Rename places a focused text box over the tab label with the existing name selected. The editor is positioned from the tab’s outer bounds so the native header’s pointer states cannot move its outline or text. Keep the tab size stable while the editor is open, then restore sizing to the label on the same native tab. Worksheet tabs keep their normal selected and hover states and suppress the native drag appearance because reordering is disabled. Active selection and keyboard focus use the app blue, with visible inactive tab boundaries. High Contrast keeps the system palette.

### Type, spacing, and shape

| Item | Specification |
| --- | --- |
| UI typeface | Segoe UI Variable with Segoe UI fallback |
| Body/controls | 14 effective px; 13 px for dense metadata |
| Sheet default | 14 effective px, regular; numeric alignment right |
| Drawer title / dialog title | 18 px / 16 px semibold |
| Section label | 13 px semibold; avoid all-caps paragraph labels |
| Code | Cascadia Mono or Consolas, 13–14 px |
| Spacing scale | 4, 8, 12, 16, 24, 32 effective px |
| Outer panel padding | 16 px; 8 px between adjacent controls |
| Toolbar/field height | 32 px minimum; dialogs may use 36 px |
| Sheet row / column | Default row 26 px, column 96 px at 100% sheet zoom |
| Corner radius | 4 px small controls, 8 px drawer/dialog surfaces; square cells |
| Focus / active-cell border | 2 px accent with visible contrast |
| Icons | Consistent Fluent family, 16–20 px; 32 px hit area for toolbar icons |

All layout measurements are device-independent/effective units. Sheet zoom scales sheet content, headers, and editor overlay consistently; it does not shrink toolbars or the Python drawer. Respect OS text scaling. Native control templates should retain their focus, disabled, and pointer-state behavior.

### Dropdown and menu direction

Every combo box opens below its field unless a specific exception is documented. When less than one row fits below the field, allow the system to keep the dropdown accessible at the window edge. Keep the current value visible in the field and include it as the selected row in the list. Use the selected row’s background and focus outline in the standard themes without the side marker. Selection-only fields use the native non-editable face with a plain value label. While open, the field outline spans the label and chevron in the standard themes. Use an editable face for fields that accept typed values. Long lists scroll within a bounded dropdown. Preserve native selection, keyboard navigation and system high-contrast styling. Coalesce size and placement corrections at normal dispatcher priority after native callbacks return. Context menus stay pointer-anchored.

After a worksheet command completes or its menu closes, return keyboard focus to the current worksheet cell when the worksheet is available. Open menus, dialogs, range pickers, editable fields and the Python workspace retain their own focus. Ribbon tabs remain keyboard navigable.

### Dialog layout (approved application-wide)

These rules apply to app-owned dialogs of every purpose: settings, editing, dimensions, colors, confirmations, errors and informational summaries. Find/Replace was the review reference, not a required arrangement of fields. Preserve each dialog's task, result semantics and appropriate content hierarchy. Windows file pickers retain their system layout.

- Put short operation status text in an existing action/form row or a reserved status area, so showing a success message does not resize the dialog. Never append a new blank/status row after an action when existing space is available.
- Size the body to its content. Remove decorative minimum heights, filler rows and unused space below the final control. Choose the smallest practical width that fits the title, labels, fields and actions; do not stretch controls merely to fill a wide dialog. Allow larger content such as color pickers, settings and reports to determine a wider size. Constrain the dialog to the available window and scroll long bodies.
- Use 16 effective pixels of horizontal padding in both body and footer. Use 12 pixels at the body's top, 10 between related form sections, and 10 below the final control before the footer. The footer has 10 pixels of vertical padding. Controls sharing the last form row share a bottom edge, so they have the same bottom margin. Avoid a blank row under an action placed beside other fields.
- Put labels in a separate row or column, then vertically center sibling editors, dropdowns and inline checkboxes in one shared input row. Align the control surfaces, not label-plus-control containers.
- Keep labels and editors aligned in form columns. Use consistent label-to-field gaps, normally 10 pixels; widen the label column to fit the longest label rather than clipping it. Repeated combos use matching compact widths. Allow enough room for the longest displayed choice, the native chevron and balanced text insets; adapt widths when a selection changes. Do not add a second large blank gutter before the chevron.
- Keep a 16-pixel semibold title and a native top-right close button. Use a 28-pixel close button with an explicit 14-pixel glyph inside a 16-pixel box so it cannot clip. Give it an accessible name. Closing with the X or Escape has the same result as Cancel/Close and must not apply edits or discard a workbook implicitly.
- Use the app's theme surfaces, an 8-pixel dialog corner radius and a thin app-accent outline (1 effective pixel by default; 0 deliberately hides it). Preserve system colors and focus in High Contrast. The footer is visually distinct from the body without adding empty height.
- Put task completion actions on one footer baseline, right aligned, with 8 pixels between buttons. Give visible footer buttons equal widths, computed from the longest current label plus approximately 16 pixels on each side, with a minimum width of 88 and height of 32. Do not stretch buttons across the footer. Include the native border and a small pixel-rounding allowance when calculating button width, and calculate the footer minimum width from the realized button fonts. Choose a dialog width that fits the action group on one line before considering wrapping; Save/Don't Save/Cancel stays on one line at normal desktop widths. Only if the available window/text scale prevents this, wrap or stack the group instead of clipping or shrinking text.
- The default action uses the native ContentDialog primary/default path and app-blue styling; it must invoke the real command, validation and result behavior. Retain secondary/destructive choices without making them the default. Cancel/Close remains a native footer action. Additional task actions may be grouped separately, but align with the footer baseline when space permits.
- Start focus in the task's first editable field when appropriate. Keep native keyboard navigation, validation, disabled states and automation semantics. Use an app-blue focus indication in normal themes. Do not show irrelevant inherited tab-close shortcut tooltips for nonclosable tabs. All tabbed dialogs use the Find/Replace TabView treatment through `ApplyDialogTabStyle`: identical header sizing, spacing, selected/inactive borders, colors and focus visuals, with inherited close-shortcut tooltips suppressed. Do not create per-dialog tab skins. Dialogs without tabs keep their appropriate arrangement.
- Numeric fields use a standard NumberBox with its native up/down repeat buttons stacked vertically inside the right edge of the field. Remove the X clear affordance. Keep increment/decrement, press-and-hold, keyboard input, bounds and validation in NumberBox; do not implement a parallel numeric-value model. No detached spin-button popup. A minimal layout adjustment to the existing native parts is acceptable; no custom drawing.

Tables in dialogs keep a visible header outside vertical row scrolling. Where users compare variable-length fields, provide native column resize handles that keep header and rows aligned. A resizable dialog must retain its initial shown size as its minimum (subject to the host window being physically smaller). Provide native horizontal, vertical and corner resize handles with standard resize cursors. Keep handles unobtrusive; do not add a decorative fullscreen/resize glyph that can appear as a corner artifact. All fields, borders and preview captions must fit the realized tab-content width, including padding, chevrons and native minimum sizes; reflow or scroll instead of clipping. A resizable dialog must resize its actual background, content and constraints together; increasing only the content width clips the title and fields. Keep resizing within the host window.

Shared implementation/tuning: `src/ExcelLite.App/DialogLayout.cs` (`DialogLayout` tokens and native dialog/NumberBox layout). Dialog-specific widths and composition remain in their own views. `SearchDialogLayout.cs` uses the shared tokens for the approved search dialog; its expanded search options are specific to that feature; `DialogTabs.cs` supplies the common tab treatment for all dialogs. Apply these rules consistently when adding a dialog, and review Dark/Light rendering, narrow bounds, text scaling, focus and destructive/cancel behavior before broader release qualification.

### Checkbox spacing (application-wide)

Use **Format Cells → Font → Effects** as the reference for new app-owned checkboxes in dialogs, panels and settings. Apply the same geometry to checked, unchecked, indeterminate and disabled states.

| Measurement | Standard |
| --- | --- |
| Control sizing | Native WinUI CheckBox sizing, with its 32-effective-pixel minimum row height; allow growth for text scaling or wrapped labels |
| Label-to-glyph gap | Preserve the native template gap used by Font Effects; do not replace the label with a separately positioned TextBlock or tune each checkbox's Padding |
| Vertical gap between checkbox controls | 4 effective pixels, normally a parent StackPanel with Spacing=4 |
| Captioned group-body padding | Left 10, top 8, right 10, bottom 10 effective pixels; use FormatGroupBox's default padding |
| Clearance below the final checkbox | At least 10 effective pixels to the group border, including a standalone checkbox such as Justify distributed |
| Inline checkbox alignment | Vertically center both the checkbox and its label with sibling editors in the shared input row |

Keep all checkbox glyphs in a group on one left edge. Use the parent container for group insets and row gaps. Preserve the native label gap when an inline checkbox needs positioning within an input row. Avoid fixed heights below the native minimum, zero group-body padding, and individual margin adjustments that produce inconsistent rows. If the window is too small, reflow or scroll the body while retaining these dimensions. Any compact variant should use a shared, documented style.

For a standard checkbox group:

```csharp
var options = new StackPanel { Spacing = 4 };
options.Children.Add(new CheckBox { Content = "First option" });
options.Children.Add(new CheckBox { Content = "Second option" });
var group = FormatGroupBox("Options", options);
```

When adding or changing checkboxes, compare realized label offsets, control heights, row gaps and group-body padding with Font Effects. Check the final checkbox's border clearance separately from extra space caused by a stretched sibling preview. Verify disabled/indeterminate states and increased text scaling without clipping; keep automated GUI verification Dark-only as specified in Development.md.

## 3. Workspace anatomy

```text
┌───────────────────────────────────────────────────────────────────────────────┐
│ Excel-Lite   Budget.exlt •                         quick Save     window keys │
├───────────────────────────────────────────────────────────────────────────────┤
│ File  Home  Insert  Formulas  Data  View                              │
│ [Paste ▾] [Font / Style] [Align] [Number ▾] [Rows/Columns] [Find]              │
├───────────────────────────────────────────────────────────────────────────────┤
│ [ C8 ▾ ]  [fx]  [ =SUM(C2:C7)                                      ] [✓][×] │
├───────────────────────────────────────────────────┬───────────────────────────┤
│       A            B             C                │ Python · Summary      [×] │
│ 1  Product       Units         Revenue             │ Anchor: Summary!C8        │
│ 2  ...                                            │ Source: xl() references     │
│                                                   │ [Run ▷] [Stop] [⋯]        │
│           virtualized spreadsheet                 │                           │
│                                                   │ Monaco editor             │
│                                                   │                           │
│                                                   ├───────────────────────────┤
│                                                   │ Result | Output | Errors  │
│                                                   │ table / value / plot      │
├───────────────────────────────────────────────────┴───────────────────────────┤
│ [◀][▶]  Sales  Summary  [+]                                                   │
├───────────────────────────────────────────────────────────────────────────────┤
│ Ready          Sum: 1250   Average: 125   Count: 10        [−] 100% [+]       │
└───────────────────────────────────────────────────────────────────────────────┘
```

This is an information-layout sketch, not a pixel-perfect screen. All top horizontal bars in the sketch have black container backgrounds. Configure the WinUI title-bar/caption area accordingly so no default white strip remains above the workspace. Retain working drag regions, minimize/maximize/close controls, and readable active/inactive caption states. In normal cell navigation, confirmation controls can be hidden until an edit begins. The drawer appears only when requested or when editing a Python block.

Keep the worksheet dominant and the Python panel resizable. Sheet zoom scales cells/headers/edit overlays, not application toolbars. Layout must respect available effective pixels, OS text scaling and native control constraints. Do not treat the layout sketch as proof of responsive behavior at every window size.

## 4. Navigation and command organization

Use the approved [Menu structure](Menu-Structure.md) and docs/references/ExcelMenu images. Ask before implementing new or moved functionality whose placement is unclear. Theme belongs in File → Options → General; native and Python calculation modes are in Options → Formulas. Formulas → Calculation Options also exposes the same native Automatic/Manual setting, applying and saving immediately; reflect the current mode with mutually exclusive checkmarks. Python visibility is Formulas → Python; execution controls stay inside its panel. No standalone Python tab or developer Check shell action.

Formulas uses the shared foreground-bound vector icon style: plain fx for Insert Function and Show Formulas, blue calculator displays for calculation, a tag and list badge for Name Manager, and blue/yellow interlocking snakes for the Python visibility toggle. Insert → Table uses a table grid with header dropdown marks. Keep native checked/disabled/focus states. Define Name has only its existing Define Name action; omit unsupported Apply Names. Insert Function retains its implemented-function picker.

Home → Number has its existing seven format types with vector menu icons, above a compact icon row for Currency, Percentage and Increase/Decrease Decimal. Currency offers Czech Kč, US $, Euro €, UK £, Slovak € and Polish zł. These use existing Currency formatting with literal symbols; grouping/decimal separators retain the app's invariant rendering. Euro and Slovak use the same trailing € format. Decimal changes retain each preset's symbol, grouping and placement. Keep tooltips and accessible command names on icon-only buttons; omit unsupported format types and More Number Formats.

Use one command definition for menu, toolbar, shortcut, and context menu. Disabled actions explain why when the reason is not obvious. Hide unsupported commands rather than suggesting they work. Do not hide supported functionality solely to achieve a cleaner screenshot.

## 5. Spreadsheet interaction

### Selection and editing

Single click selects; double click or F2 edits. Typing starts replacement; formula bar and inline edit buffer represent the same draft. Enter commits and moves down; Tab commits and moves right; Shift reverses direction. Escape cancels only the edit in progress, then returns focus to the active cell.

The active cell has a 2 px border. A rectangular selection uses a subtle tint and emphasized row and column headers. Ctrl+click adds or removes separate cells, treating a merged cell as one area. Earlier areas keep a thin accent outline and their authored fills. The name box shows the active cell for a selection with multiple areas. Escape keeps the active area. Switching sheets starts the destination at A1 and clears the previous selection and scroll position.

Formula edit mode visibly begins with `=`. Provide lightweight supported-function suggestions and argument hints if the chosen component supplies them reliably; the required function picker remains available regardless. Invalid formulas display a clear error after commit; source stays editable. Formula result tooltips distinguish an actual computation error from a stale cached value.

Python anchors show a compact `Py` marker and result type indicator, plus text/accessibility information. Double-clicking a Python anchor opens its script drawer rather than editing the internal script ID. The cell entry is `=PY`/`=PY()`; selecting an anchor or spill child opens its source. Ordinary selection clears the Python panel. Empty cells retain a visible blank formula edit box. Keep implementation IDs out of the drawer title; use a user-friendly block name and address.

### Formula parameter hints and regional entry

When editing a supported function, show a compact hint under the cell editor or below the formula bar near its text caret. Bold the current argument and track nested calls and quoted strings. Close the hint on commit, cancel or focus loss. Insert Function starts the draft at the active cell and places the caret inside the parentheses. It edits in the formula bar when visible and in the worksheet cell otherwise. Use the Windows regional list separator for function argument entry and signatures through .NET [TextInfo.ListSeparator](https://learn.microsoft.com/dotnet/api/system.globalization.textinfo.listseparator). Preserve persisted formula syntax and literal input in Text-formatted cells.

The status-bar zoom percentage opens the same Zoom dialog as View → Zoom. Use native radio buttons for 200%, 100%, 75%, 50%, 25%, Fit selection and Custom. Custom uses the shared compact NumberBox with native vertical spin buttons and no clear X; all digits must remain visible. Zoom spans 10–400% and scales the sheet. Dialog changes apply on OK; Cancel leaves zoom unchanged. View → Zoom to Selection immediately fits and reveals the current range, preserving the selection. Both fit commands exclude hidden/filtered cells and reserve room for headings, frozen panes and scrollbars within those zoom bounds.

Data → Sort & Filter uses blue A/Z letters, a split sort grid and outlined funnels. View uses a magnifier for Zoom, a blue 100 badge and selection grid, plus blue frozen-row/column accents and snowflakes for Freeze Panes. The Unfreeze Panes icon uses the same grid with a cleared split. Frozen cells use inset boundary lines and retain their formatting and pointer interaction. View → Show uses a vertical stack of native checkboxes whose glyphs display visibility state. Preserve native disabled, checked and keyboard-focus states.

### Clipboard, structure, and data operations

- Copy outlines the source; cut uses a distinct dashed outline until committed/cancelled. Paste respects the selected paste mode and produces one Undo step.
- Before a destructive overwrite by Python table insertion, show destination range, dimensions, and non-empty cell count. The final action is Insert Values, not Run again.
- Sort dialog shows range, header state, key column, direction, and that all records in that range move together. Offer Expand to adjacent data when a single selected column is part of a larger contiguous block.
- Filtered headers display a funnel marker; status shows visible/total records. Manual hidden state and filter-hidden state remain distinct so clearing a filter does not unhide manually hidden rows.
- Filter dropdowns use a bounded, scrollable checklist that creates checkboxes as needed. Search and selection apply to values throughout the list, including values outside the scroll viewport.
- Selected-range statistics count visible selected cells only, excluding both manually hidden and filtered-out rows/columns. Sum/Average use numeric cells; Count means non-empty cells, including text. Label statistics accessibly and use the same policy in both themes.
- Deleting a sheet shows its name and any dependent formula/Python references. Undo remains available. Routine individual cell edits do not trigger confirmation dialogs.

### Shortcut routing

| Context | Keys | Behavior |
| --- | --- | --- |
| Global | Ctrl+N / O / S / Shift+S | New / Open / Save / Save As |
| Grid | Ctrl+Z / Y | Workbook Undo / Redo |
| Grid | Ctrl+C / X / V | Clipboard commands |
| Grid | Ctrl+F / H | Find / Replace |
| Grid | Ctrl+D / R | Fill Down / Right |
| Grid | Ctrl+Home, Ctrl+arrows, Shift+movement | Navigation/selection within supported navigation |
| Grid | F2, Enter, Tab, Escape, Delete | Edit/commit/move/cancel/clear |
| Monaco | Ctrl+Z / Y, Ctrl+F | Editor-local Undo / Redo / Find |
| Monaco | Ctrl+Enter | Commit source and Run selected block |
| Application | F6 / Shift+F6 | Cycle major regions, including back from Monaco |
| Menus/dialogs | Escape | Close current transient surface and restore prior focus |

Global Save must flush editor drafts before writing. Do not intercept normal code-editor keystrokes for sheet actions. Document any native control conflicts when changing controls and resolve them before adding more shortcuts.

## 6. Python workspace and result states

The drawer identifies the selected Python cell and opens its source automatically. Typing =PY creates the anchor/editor without exposing block creation. Inputs are discovered from xl() reference literals; show references as optional explanatory details, not mandatory editable bindings. Picking a range inserts an xl() call while preserving the editor draft; Enter confirms and Escape cancels. The block selector, Create and Apply bindings fields are removed; old binding metadata remains internally supported.

Keep Python Automatic/Manual in File → Options → Formulas, separately from native calculation. Explain that saved Automatic starts/enables Python on Opening/New/import/reopen and refreshes stale cells. Missing older preferences retain manual execution until chosen. Enable Python and Run Cell/Run All/Stop/Restart stay in the existing panel (Formulas → Python), as approved by the user. No new Python tab. Stop is available while queued/running; Restart cancels work and clears process state while retaining document results.

| State | Cell / drawer presentation | Available next action |
| --- | --- | --- |
| Not run | `Not run` label; missing-result marker | Run |
| Current scalar | Normal value with `Py` marker; timestamp in details | Edit / Run |
| Current table/plot | Owned worksheet spill or `Plot` indicator; source remains anchored | Edit / Preview / Copy as Values / Save Image |
| Stale | Last value plus warning; distinguish Manual from pending Automatic refresh | Run / inspect last result / enable Python |
| Queued/running | Status text and indeterminate progress; elapsed time | Stop; continue unrelated sheet work |
| Failed | Error marker; concise message; traceback in Errors | Edit / Run; inspect last successful result |
| Cancelled/timed out | Explicit label; last success never presented as newly computed | Run / inspect error details |
| Upstream stale | Identify upstream block; no misleading result | Run Cell / Run All |
| Obsolete completion | `Inputs changed during execution; result discarded` | Run again |

Keep Result, Output, and Errors tabs stable; use badges for output/error counts instead of moving focus automatically. A table preview states `Showing 500 of 2,000 rows`; Worksheet spilling is the normal DataFrame output path; Copy/Insert as Values is secondary and uses the full validated result, not only the preview. A blocked spill shows #SPILL! and an actionable destination conflict without overwriting cells. Figures fit the panel with a zoom/open preview option if available. Save Image exports the figure stored on the block. The run that creates a placed chart stores that figure at the placement size; a later resize keeps a separate render.

Worksheet charts float over cell content and can overlap. Draw the selected chart above other charts so its outline and resize handles stay visible, and use that same order for pointer hit testing.

Manual enablement copy: “Python code can access files and the network using your Windows account. Only run code you trust.” State whether enabling runs stale cells. Workbook authorization remains session-local; the saved local Automatic preference authorizes enablement on workbook adoption. Manual and Stop remain accessible. Selection, Undo/Redo and Restart execute no code. Do not claim sandbox protection.

Error copy should identify the cause and a next step: “The input contains an invalid date in Sales!B8. Correct the value or import that column as text.” Avoid raw protocol errors or stack traces in the main notification.

## 7. File workflows

### New and Open

Start with an empty worksheet grid spanning A1:Z1000. File/Home offers Blank workbook and Recent. Open the supplied [showcase workbook](showcase/README.md) through File/Open. Open/import shows a cancellable busy state when needed and retains the current document until validation succeeds. Adoption follows the saved Python preference: Automatic enables the interpreter and refreshes stale cells; Manual does not run code. Unsaved document dialog: Save, Don't Save, Cancel. “Don't Save” is not the default focused destructive action.

### Save and export

Title shows filename and a small unsaved dot. Save uses `.exlt`; import from XLSX starts an unsaved native project. Export dialog identifies format and scope. Its compatibility summary states exactly what is omitted or converted: for example, “3 Python blocks will be exported as cached values; 1 is stale. Python code is retained only in the Excel-Lite project.” User can cancel or continue knowingly.


On save failure retain the dirty indicator and previous valid file. Message: “Could not save Budget.exlt because the file is in use. Close the other application or choose Save As.” Never show a successful save toast before replacement completes.

## 8. Accessibility and visual quality gate

Expose grid coordinate, value/formula, selected/editable state, error/freshness, row/column headers, and Python markers through accessible names/patterns supported by the chosen grid. Keyboard-only navigation is required but is not a substitute for screen-reader verification. Native dialogs and menus retain standard automation semantics.

Respect reduced-motion settings. Prefer built-in control transitions; no decorative movement. Busy indicators may animate, but avoid flashing selections or layout shifts. Notifications should remain available in a status/details surface after transient display.

Review these populated states at 100%, 150%, and 200% OS scaling, dark/light/high contrast, plus narrow layout:

1. Normal data sheet with currency/date/text formatting and selected range.
2. Formula edit with a long expression, cross-sheet reference, and visible error.
3. Filtered sheet with hidden rows, frozen headers, and horizontal scrolling.
4. Cell-first =PY editor with discovered xl references and a scalar result.
5. Large table preview, plot result, stale output, and traceback.
6. XLSX export compatibility report.
7. Save failure, cancellation, unsaved-close dialog, and first-run trust prompt.

Pass criteria: black title/command/formula bars with no white title strip in either standard theme, including inactive-window state; no clipped labels or unreachable primary actions; consistent focus restoration; all state distinctions readable without color; authored cell styles preserved across themes; no hidden dependency on internet-hosted fonts/icons/editor assets; no unsupported buttons. Report actual checks in the chat; saved screenshots are not required. Accessibility and visual criteria do not count as verification by themselves.

## Quick Access and Clipboard styling

Quick Access uses 32-pixel-wide icon controls with transparent resting surfaces on the black title bar. Home Clipboard uses a large Paste icon and a separate options arrow beside three compact icon/label rows for Cut, Copy and Format Painter. These controls retain native hover/pressed/focus states, tooltips and accessible names. The active painter has a blue-tinted surface and a brush cursor in the sheet. Styling is tuned for the black command area. The application identity is a cyan/teal spreadsheet tile with a blue folded corner; the multi-resolution ICO is embedded in the executable and explicitly assigned to AppWindow for the window/taskbar.

The title bar displays the app icon at 24 effective pixels before Save. Format Painter uses the supplied Excel brush silhouette with a narrow handle, ferrule and yellow bristles, scaled into a 20-pixel square. Clipboard/Font vector outlines inherit the control foreground; authored icon accents remain fixed in standard themes and follow system ink in High Contrast.

## Alignment controls

Home Alignment uses two rows of three icon-only buttons: Top/Middle/Bottom above Left/Center/Right. Use the Excel line silhouettes with inherited app foregrounds, native checked/hover/focus states, tooltips and accessible names. Wrap Text uses lowercase ab above c with a thin blue return arrow beside its label; explicit compact padding keeps both within the 28-pixel button. Orientation sits beside the vertical alignment row; its button and angled menu commands use tilted lowercase ab lettering with thin blue directional arrows. Angle Clockwise rotates the entire Angle Counterclockwise drawing 90 degrees clockwise around the shared viewport center. Format Cell Alignment adds a blue baseline beneath the upward arrow. Lettering inherits the control foreground, and the complete drawing scales into the native icon slot. Decrease/Increase Indent sit beside the horizontal row. Wrap Text and the labeled merge split control sit to their right. The group caption includes an Alignment dialog launcher. The merge menu uses fine grid outlines with a blue middle band: two-ended arrows for Merge & Center, a right arrow for Merge Across, an empty merged band for Merge Cells and separate blue cells for Unmerge Cells. Vertical Text stacks actual letter outlines; Rotate Text Up/Down rotate the lettering beside blue vertical arrows. Indent uses full-width top/bottom rules, two right-aligned middle rules and a blue horizontal arrow in the left gap. Alignment flyouts use a 22-pixel square icon slot with an 8-pixel label gap instead of the native 16-pixel slot. Alignment menu items use an explicit two-column template with a fixed icon column; icon/label spacing is established before first measure, never patched through Loaded or native visual-tree searches. Orientation toggles indicate selection around the icon and retain native toggle semantics, hover/pressed/disabled states and keyboard focus. All Alignment drawings share a square vector viewport and scale into the native icon slot. The merge menu retains only the four approved commands. Keep the existing black top bars and blue accents. Sheet cells default to bottom vertical alignment; wrapped content clips within authored row heights, including the combined height of a merged range.

## Cells and Editing controls

Home Cells uses large grid icons over Insert, Delete and Format. Editing stacks the AutoSum split control, Fill and Clear beside large Sort & Filter and Find & Select controls. Fill uses a blue arrow inside a sheet frame, Clear uses the pink eraser, Sort & Filter pairs A/Z lettering with a funnel, and Find & Select uses a magnifier angled down-left. Local vector outlines inherit the themed foreground and accents follow system ink in High Contrast. Keep native hover/focus states, accessible names and tooltips. AutoSum has the same five-option menu in Home and Formulas; its main button runs Sum.

Cells and the refreshed Editing controls use 24-unit SVG drawings in 18- or 32-pixel ribbon slots and fixed 22-pixel menu slots. Sort & Filter and Find & Select use two-line captions with the chevron beside the lower line. AutoSum retains its path and split action. Anchor its menu below the left edge of the whole control, aligned like Fill.

Fill menu arrows point in the action direction (Down, Right, Up, Left). Replace uses two small outlined cells with opposing bent arrows. Sort menus use A-over-Z/Z-over-A with a descending arrow, opposing arrows in a square for Custom Sort, a funnel for Filter and funnel overlays for Clear/Reapply; all inherit the themed menu foreground.

Disabled menu labels, icons, shortcut hints and checkmarks use the native theme's muted disabled-state brushes; High Contrast retains system colors. Sort & Filter separates sorting from filter commands. Clear is unavailable without active criteria and Reapply requires a filter range. Fill rotates its framed SVG for each direction. Clear All uses the pink eraser and Clear Formats separates its percent mark from the miniature eraser. Clear Contents aligns in the same label column.

## Number formatting

The number-format selector displays the active cell’s style and keeps its accessible Number Format name. Comma Style is an icon-only command between Percentage and Increase Decimal. Format Cells has Number, Alignment, Font, Border and Fill tabs; Number uses a left category list and right sample/options area, using the shared native dialog theme. Only controls relevant to the selected category appear.

## Date/time locales and editing

Date/Time format options place the localized Type list above Locale (location), following the supplied Excel reference. The cell edit buffer uses readable localized date/time text and retains the cell’s horizontal alignment; temporal values with General alignment edit right-aligned. Time edits include any hidden date component.

## Cell editing and Special formats

Worksheet cells use equal 2 px vertical padding and a consistent text line box. Top and Bottom keep matching boundary insets; Middle centers the rendered block, including merged cells. F2/double-click enter caret editing: Left/Right use native input caret/word/selection behavior, Up/Down move to the value’s start/end and Shift extends the selection. Typing directly into a selected cell retains arrow-to-commit navigation. Format Cells restores worksheet focus after closing.

Worksheet scrollbars use slim tracks with rounded thumbs and centered arrows. Page Up and Page Down move the selection and viewport together.

The worksheet inline editor uses a solid worksheet-colored background, a square accent border, and the cell’s font and horizontal alignment. Long values such as timestamps can widen the editor within the viewport. Row and column dividers use distinct resize cursors on hover and during drag. On the worksheet, Top, Middle and Bottom align the visible Vertical Text stack after number formatting. If the stack is taller than the cell, the selected alignment determines which portion is visible.

## Font and Fill formatting

Format Cells uses five native tabs: Number, Alignment, Font, Border and Fill. Font has aligned family/style/size columns with lists and text fields, underline/color controls, Effects and a live preview. Fill places background controls at left, pattern controls at right and Sample below. Color dropdowns use inline custom-color controls; Fill Effects uses a draft flyout within the existing dialog, avoiding nested ContentDialog instances. Colors, patterns and gradients update previews before acceptance; Cancel never changes cells.

## Format Cells control layout

Use non-closable native TabView tabs for Format Cells, with visible tab boundaries and normal control-size labels. Format list rows are compact and have no informational hover tooltips; selected font/size entries scroll into view. Number Sample matches the 32 px editor height. Negative examples stay visible in a list and render red as color rather than a text suffix. Font color selectors show Automatic or the color alone; labels belong above the selector. Effects and Preview use captioned outline groups. Font preview uses the selected family name and converts point sizes to effective pixels; font/size lists stay synchronized with typed values and Normal font reset. Font and pattern-color dropdowns include More Colors with inline hex validation and Apply; Fill also exposes a separate More Colors button; picking/typing a custom color does not apply it until Apply.

Format Cells keeps a fixed viewport across categories and tabs (bounded by the host window), with scrollable bodies, the Find/Replace 8 px gap below tabs, and equal left/right content insets. Center checkbox labels within the input row as well as the checkbox surface. Fill uses a plain No Color button matching the swatch grid width, explicit More Colors, a graphical pattern dropdown, and a sample that paints the complete inner frame.

Pattern Style and Pattern Color open below their controls; constrain tall color palettes to the available space and scroll them. Pattern thumbnails must visibly show distinct Excel-style dots/lines/grids in the chosen ink color. Verify rendered ink and paper pixels rather than successful image loading alone. Keep native checkbox label gaps consistent across inline and grouped checkboxes.

Fill Effects retains the top color controls, puts shading radio options beside clickable gradient variant thumbnails, and places a full-width final sample below both. Show the selected variant clearly and refresh thumbnails/sample when either color or shading changes. Solids, gradients and None previews must paint the complete content rectangle; only the normal control padding and dropdown chevron area remain outside the preview.

Font uses the same column widths/gaps for its upper fields, Underline/Color row and Effects/Preview divider. Fill Effects has a draggable title and shared blue outline; shading offers three linear variants, four corner variants and one center variant. Corner/center shading follows Excel’s rectangular path shape, consistently in previews and sheet cells. Repeated tab loading must preserve native control geometry without nesting layout wrappers.


Place Border between Font and Fill. Group line style/color at left, presets and preview at right; keep individual controls keyboard accessible and show mixed edges as indeterminate. Style/color changes update enabled uniform edges and the pen for subsequent clicks; leave mixed edges unchanged until targeted. Use the same Automatic/color palette and inline More Colors validation as Pattern Color.

Match Excel’s Border layout: two seven-row style columns without scrollbars, graphic presets with captions below, Top/Inside horizontal/Bottom icon buttons to the left, and Diagonal up/Left/Inside vertical/Right/Diagonal down beneath the diagram. Use centered text in a 1x1, 1x2, 2x1 or 2x2 preview. Show only corner/midpoint guides for absent edges. Match the worksheet canvas and use contrasting Automatic ink.

The Home Borders dropdown uses Excel’s grouping/order, drawn border icons, a Draw Borders section and nested color/style choices. Line Color reuses theme/standard colors and the shared inline More Colors picker, with recent custom colors. Keep drawing tools persistent with a visible pencil/eraser cursor, range feedback and Escape exit. Drawing must not enter text editing or change the selection; erasing targets shared line segments rather than clearing whole cell styles.

Build the Borders submenu rows ahead of opening the parent menu. Create line samples when their submenu loads, and reuse rendered border icons while their foreground and contrast mode match.
Use an opaque Borders flyout background from the worksheet theme so the Windows backdrop cannot override it. In High Contrast, use the system background.

The Borders primary button must execute its displayed action without opening a menu; the separate arrow opens the menu. Retain a selected drawing/erase icon after Escape for direct reactivation. Preview straight grid-edge drawing with a single styled line, switching to a rectangle only when the gesture leaves that edge. Use distinct, outlined blue pencil and grid-pencil cursors with tip hotspots. Match the Erase Border cursor to the ribbon's pink eraser icon and center a small contrasting cross on the erasing hotspot.

Home Font uses compact AutoSum-style dropdown segments: 10 px chevrons in 18 px arrow areas. Underline, Borders, Fill Color and Font Color keep separate 28 px primary actions. Editable font-family/size controls are 140/72 px wide with compact arrow areas.

Compact Font controls override native split-column minimums and editable-combo right padding. Keep chevrons centered and font dropdown rows at a compact 28 px minimum height. The Font and Font Size lists use the shared downward placement and native scrolling, with the native split opening animation removed.

The font-size field reserves enough width for complete two-digit and fractional values. Apply editable-text padding when the native ComboBox lazily creates its editor, and detach the layout callback once sizing is applied.

The Format Cells Alignment pane uses four captioned group boxes from the shared FormatGroupBox helper: Text alignment above Text control/Right-to-left on the left, with Orientation spanning both rows on the right. The outlines have consistent neutral borders and aligned outer edges; group bodies and checkboxes use the same padding, default control sizing and four-pixel row gaps as Font Effects. The pane fits without scrolling at the maximized reference size; shorter windows retain accessible scrolling. Indent sits beside Horizontal; the Horizontal and Vertical combo boxes share the same left/right edges. Text control and Right-to-left sit side by side beneath the alignment fields. The orientation dial, numeric degrees and stacked-text toggle share one draft; valid typed degree values update both the dial needle and text preview immediately without committing or moving the caret. Inapplicable controls disable with the selected alignment. Preserve the existing dialog tabs, margins, native focus/checked states and accessible control names. Text transforms affect value content rather than cell borders/fills; row heights remain authored and may clip rotated/wrapped content. Vertical Distributed spreads words along the cell height for ±90° and stacked Vertical Text, and centers other single rotated runs; unrotated multi-line values retain even line distribution. Vertical Justify starts a single run at the top.

Function hints use an ellipsis for repeating arguments. Clicking or dragging worksheet cells while entering an argument inserts a cell or range reference and outlines the picked range. Further clicks replace that reference until typing continues. The formula and hint stay in the editor where the draft started. A pick from the formula bar returns keyboard focus there after release. The formula retains its original target, Enter commits and Escape cancels.

Home Clipboard and Font use local vector artwork matching the supplied Excel references: yellow clipboard, paper copy, blue scissor handles, a bristled paintbrush, letter-based font commands and blue size arrows. Border presets use dotted guide grids with solid authored edges; drawing/color tools use blue pencils and Erase Border uses a pink eraser. Paste, Underline and Borders share the fixed-column menu template, including native submenus. Alignment and Font dialog launchers use a transparent 16-pixel square button with a 12-pixel corner-arrow drawing and no rounded resting tile.

Font’s launcher glyph has a two-pixel optical offset down and right within its existing hit target. Fill Color uses a tilted handled pail with a blue paint drop above the selected-color bar. Erase Border is narrow, with an unfilled upper half and pink lower half; Line Color uses an outlined blue pencil, a corner guide and a solid ink bar. Font/Fill color palettes reuse a shared lightweight Button template, inherit theme resources from their panel, cache presenter styling by theme/contrast and open without animation.

Number format choices use 32-pixel icons, bold names and an active-cell preview on a second line, formatted with the existing preset renderer; General shows No specific format. The selector keeps its chevron at the right edge. Number toolbar icons use thin money/coin, percentage, comma and blue decimal-arrow artwork. Cells uses fine sheet grids with blue insertion arrows, red deletion crosses and a blue width bracket for Format; row/column/sheet menu entries have distinct drawings. Format headings and separators follow Excel and omit unsupported commands. The Alignment launcher uses square-ended corner/diagonal strokes and a right offset that compensates for the group body padding, plus a three-pixel downward offset, placing the glyph within eight pixels of the group’s bottom/right edges.

Merge & Center retains its existing height and primary-segment sizing/padding; only the dropdown segment is narrowed to 18 pixels with a centered compact chevron, matching the Font controls.
