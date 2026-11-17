# Excel-Lite architecture

Excel-Lite is a local Windows 11 x64 application. WinUI 3 owns the shell, dialogs and commands. Win2D draws the worksheet, and WebView2 hosts the local Monaco editor. C# evaluates native formulas, and a separate Python worker receives immutable input snapshots. The workbook session owns all authored data and committed results.

## Project boundaries

| Component | Responsibility |
| --- | --- |
| `ExcelLite.Core` | Immutable workbook/cell/reference/style/Python contracts, validation and document limits |
| `ExcelLite.Application` | `WorkbookSession`, commands, history, saved identities and adoption/commit contracts |
| `ExcelLite.Calculation` | Native dependency graph, parser-backed reference services and ClosedXML evaluator adapter |
| `ExcelLite.Persistence` | Validated native packages, XLSX exchange, reports and atomic writes |
| `ExcelLite.Python` | Worker discovery/lifecycle, framed protocol, conversion and host-side result validation |
| `ExcelLite.Foundation` | Composition and mixed native/Python calculation |
| `ExcelLite.App` | WinUI shell, Win2D worksheet and Monaco editor, UI scheduling, settings and composition root |
| `worker/` | Private CPython request handler and result/figure normalization |

Core/Application do not depend on WinUI, Win2D or file-library objects. Grid edits emit intents; only session commands mutate the workbook. Calculation and I/O use detached immutable snapshots off the UI thread; completion validates its basis before committing.

## Document and history

Workbooks have stable IDs and ordered sheets; sheets use stable IDs, sparse cells, dimensions, manual visibility, filter, freeze and merged-range state. Cells distinguish blank, empty text, finite number, Boolean, native formula and Python anchor. Authored source is separate from derived caches. Python results retain source/input/runtime fingerprints, last success, freshness and latest execution status.

Coordinates at model boundaries are one-based. References are parsed tokens, never arbitrary address-string replacement. Copy/fill translates relative dimensions; cut/move retains identity and redirects dependents. Structural edits shift/expand/shrink ranges, including absolute references; deleted targets become `#REF!`. Sheet copy creates independent Python block/name identities and retargets supported references. Undo restores prior identities and references.

Commands validate before mutation and commit one forward/inverse transaction. Compound paste, Replace All, sort, structure and Python result generations undo atomically. Native derived recalculation adds no history. Python Undo/Redo restores cached state without executing code; external script side effects cannot be undone. History is bounded to 100 transactions or 64 MiB; oversized edits are rejected rather than silently losing Undo.

A session ID, monotonically increasing revision, generation and authored-content identity guard asynchronous work. Undo never rewinds revision. Saved identity survives history eviction; derived native caches do not dirty authored content. Saves capture immutable snapshots: edits during a save stay dirty, and an old session's completion cannot mark a replacement workbook saved.

Model bounds are 10 sheets, 10,000 rows by 100 columns per sheet, 200,000 populated cells and 10,000 formulas. The worksheet projection is A1:Z1000 per sheet; model bounds do not imply that the full model is navigable in the current viewport.

## Worksheet rendering and input

The worksheet host publishes the active workbook to Win2D in process. It caches sheet views by immutable workbook identity, draws visible cells and headings, and clips each frozen pane to its viewport. Cell layouts use authored dimensions, number formats, styles, merges, and derived filter visibility. Floating charts use worksheet coordinates and cached images. UI Automation exposes grid and selection providers, with offscreen cells available on demand.

Sheet switching flushes pending edits, resets the selection and scroll position, and reuses the last checked calculation state when workbook and runtime fingerprints match. The worksheet maintains an active rectangle and additional Ctrl+click areas. The shell uses those areas for formatting, clearing, and selection statistics.

Worksheet zoom is view state (10–400%). The ribbon, status controls, and Zoom dialog share the same value. Fit selection uses authored dimensions, visible headings and frozen panes to compute the scale and reveal the selection. Formula drafts retain their original cell while the user picks references across sheets.

## Native calculation and data operations

`NativeCalculator` uses ClosedXML.Parser to discover cross-sheet rectangular dependencies, detects cycles before evaluation, and evaluates eligible expressions against detached literals/completed predecessors. Circular cells carry `#REF!` with a circular-reference detail; malformed source uses `#VALUE!`, unsupported functions `#NAME?`. Unsupported imported formula text is retained. References in inactive branches still contribute graph edges.

The supported function list is in [Python capabilities](Python-Capabilities.md#native-formulas). IF/IFERROR evaluate only the selected value branch; AND/OR evaluate their arguments. Aggregate range values ignore text/booleans where appropriate; direct scalar arguments follow the evaluator's coercion rules. Exact VLOOKUP requires FALSE/0. Native dynamic arrays, iterative calculation and external workbook links do not execute. Python tables use separate host-owned spilling.

Native Automatic evaluates affected dirty nodes; Manual retains explicitly stale caches until recalculation. The native evaluator never launches Python. `MixedPythonCalculation` keeps one native calculation context for a generation. It evaluates native prerequisites only when blocks have unrefreshed inputs or require execution, then reuses that ClosedXML workbook. A changed Python result, or a prerequisite whose calculated value changed, invalidates the affected native formulas and those formulas are recalculated in place. Unrelated native formulas keep their caches. Volatile `TODAY` formulas and their native dependents refresh once per generation. The edit builds one dependency graph and the generation reuses it, rebuilding only when spill topology changes and replanning when those edges change. A placed chart is rasterized once at its placement size during that execution; a later resize still redraws the retained Figure. Detached steps validate the changed anchor, spill and figure records without hashing the whole workbook. The live session then commits the generation as one undo and validates the document. The context combines native dependencies, literal `xl()` ranges/names/tables and spill ownership; the UI scheduler supplies execution authorization. See [Python capabilities](Python-Capabilities.md) for Automatic/Manual, prerequisite execution and result freshness.

Persisted formulas use invariant English function names and comma argument separators. `FormulaPresentation` adapts function argument separators to Windows regional entry/display without rewriting text, quoted strings, sheet names, unions or array separators. This is not full function-name/decimal localization.

Sort is stable, single-key and top-to-bottom: numbers, text, Booleans, errors, blanks; blanks remain last in either direction. Text uses CurrentCulture with optional case sensitivity. Headers stay fixed. Content/styles and relative formulas move with rows; external formulas retain their coordinates. Hidden rows participate; dimensions/manual hidden flags remain attached to worksheet indices. Partial-block sorting offers expansion before mutation.

One filter rectangle per sheet stores allowed formatted-display values per column, combined with AND. Reapply/edits derive filter visibility separately from manual hiding. Each dropdown derives its choices with the same criteria evaluator, skipping its own column. The checklist virtualizes its controls and keeps selection state by value so scrolling and searching preserve it. Clear criteria retains dropdowns and Filter off removes the definition. Visible selection statistics exclude either hidden state. Freeze stores worksheet-relative row and column counts in the sheet's authored state. The grid pins the original cells and headings using authored row heights and visible column widths and reserves space for scrolling if a saved pane exceeds the viewport. Native packages retain the legacy single-axis shortcut values and store other boundaries as a pair. Tables add explicit extents, cached schema, styles and optional summary rows. See [Menu structure](Menu-Structure.md#tables) for their interaction contract.

## Python and editor boundaries

One managed private CPython worker serializes execution. Inherited pipes carry length-prefixed UTF-8 JSON with protocol/request identities; stdout/stderr are captured separately. Handshake validates runtime/package versions including pyarrow. Eligible DataFrame/Series results use Apache Arrow IPC in session files that the host maps. The default input transport is JSON. The worker client supports an explicit Arrow input mode for benchmarks, with JSON fallback for unsupported tables. Matplotlib figures write PNG bytes into that directory and return the relative path rather than a base64 JSON string. The host reads the result files, validates the result, then deletes the payload. Handshake, scalars, errors and small named inputs remain JSON. The host enforces timeout, bounds, stale-completion rejection and result commit. Windows job containment controls worker/descendant lifetime; it is not a sandbox. Scripts can access files/network as the current account.

Each run has its own namespace. Literal `xl()` references are discovered without executing Python; the worker validates syntax, headers and conversion before user statements. Scalars/tables/figures are validated again by the host. Output limits and conversion details are in [Python capabilities](Python-Capabilities.md).

`PythonScheduling` coalesces committed edits and serializes detached mixed generations. An isolated single-cell edit uses a 75 ms debounce. Paste, compound edits, and edits during a pending or running generation use 220 ms. A matching generation commits atomically; superseded work is discarded. Saved local Automatic authorizes enablement at workbook adoption; Manual resets session authorization. Workbook files never grant execution authority. Selection, Undo/Redo and Restart never execute source.

Cell lookups share a lazy index per immutable sheet instance. Mixed calculation reuses verified upstream Python results within the generation and clears those checks when native caches or spill topology change. The native worksheet context reuses its evaluation order and skips allocating cells for authored blanks.

Monaco uses a restricted local virtual host and versioned JSON messages. Drafts flush before selection/file/run barriers; stale editor messages cannot overwrite another selected source. Monaco keeps text Undo; workbook history owns committed edits. Selecting an ordinary cell clears the Python editor/results, including after background output updates.

## Floating figures

Each placement belongs to one Python figure result and stores a target sheet, cell anchor, unzoomed offsets and independent width/height. Legacy cell-extent placements remain readable. Placement bounds are validated independently of cell contents and other figures. Movement, resizing and deletion use session history and preserve covered cells. A source may have multiple differently sized placements.

A placed chart's execution rasterizes that placement in the same worker request and retains the untouched Figure under the returned image id. On resize release the worker clones that Figure, sets the target canvas/DPI and recomputes standard subplot layout. Crowded x labels rotate or thin on that disposable copy; custom manually positioned axes retain their authored layout. The original is never progressively resized. Scoped counter-copy support handles embedded Python's `itertools.count` copying limitation. Requested/returned PNG dimensions must match; per-placement image caches include a renderer-versioned source identity.

The worker retains up to 32 Figures and bounded compressed snapshots in a random host-created session directory (64 MiB raw object, 256 MiB total disk). These worker-produced objects can restore across worker restart; they never come from workbook files or protocol-supplied pickle data. The directory is cleaned on client disposal. Nonserializable custom artists may require explicit source execution after restart.

Automatic generations prepare placed chart sources lacking live/archive state even when their saved PNG is current; other current caches are skipped. Resize can wait for queued calculation. If the retained Figure is unavailable, resizing can run its source and prerequisites after workbook consent. Geometry/image commit requires the same session, placement geometry and source identity. A fitted cached image is the preview/error fallback; redraw failures remain visible in status and the owner's Errors panel. Current verification limits are in [Development](Development.md#verification-boundaries).

## Native files and lifecycle

`.exlt` is a version-1 ZIP containing `manifest.json`, `workbook.json` and `sheets/<guid>.json`. Workbook JSON includes ordered sheets and names/table definitions; sheet cells embed Python source, bindings, results, spills, figure placements and optional rendered PNG variants. There are no separate `python/`, `results/`, `assets/` or `view.json` entries. Execution authorization and temporary serialized Figure objects are never saved.

The reader validates entry names, duplicates, streamed sizes, schema and reference integrity before detached adoption; it never extracts arbitrary archive paths. Default limits are 128 MiB compressed, 512 MiB expanded, 32 MiB per entry and 10,000 entries. Invalid/newer incompatible packages fail without replacing the current document.

Save flushes drafts, writes a unique same-directory temporary file, flushes it, reloads and checks document equivalence, then replaces/moves it atomically. Failure preserves the previous destination and dirty state. File locks and edits during saving must not become false success; filesystem behavior does not imply stronger power-loss guarantees.

Startup/New creates a clean blank workbook. New/Open/Close flush drafts and use Save/Don't Save/Cancel. A cancelled picker, failed save/read or changed adoption basis prevents replacement. File Close returns to Home. Recent contains up to 10 paths. Native Save/Save As uses `.exlt`; an imported XLSX source is distinct from the native path and cannot be overwritten by session export. [XLSX compatibility](Xlsx-Compatibility.md) describes reviewed interchange losses; export never evaluates Python. CSV is outside scope.

Theme, calculation preferences, Recent and Find/Replace histories live in validated atomic local settings under `%LOCALAPPDATA%/ExcelLite`. View preferences do not rewrite authored cell formatting. There is no implemented autosave/crash recovery claim.

Formulas → Calculation Options routes Automatic/Manual through the same native formula actions and local preference persistence as Options. Menu checks synchronize from the workbook mode when commands finish and whenever the menu opens; Options reads that same state. The ribbon has no separate calculation-mode model and does not change Python's independent calculation preference.

## Distribution and capability boundaries

SDK/NuGet locks, local Monaco assets, and runtime manifests pin dependencies. The offline package includes application assets, a private Python runtime, prerequisite installers and notices. See [Development](Development.md) for packaging and verification commands.

## Format Painter and application identity

`FormatBrush` in Foundation captures an immutable rectangle of authored `CellStyle` values independently of the system clipboard. It builds one style-only `SetCellsCommand` against current destination content, validates session and destination bounds, and repeats the source pattern for selected rectangles. The worksheet host owns transient single-use and locked modes and applies the brush on pointer release. Workbook replacement invalidates the brush. `Assets/AppIcon.ico` supplies both the executable ApplicationIcon and AppWindow icon; its generated PNG source and prompt live under `Assets/Brand`. The title bar loads that PNG as a 24-pixel Image before the Quick Access commands.

## Alignment and merges

`CellStyle.VerticalAlignment` stores top/center/bottom/justify/distributed (null means bottom); sheet `Merges` stores validated nonoverlapping rectangles. Merge commands build detached structural proposals. The topmost, leftmost populated cell in each merge area supplies the anchor's content. Merge Across evaluates each row separately, and confirmation is required when a merge would discard populated cells. Merges participate in equality, authored identity, history size and structural deltas. Native packages and XLSX retain merge extents and vertical alignment. Copy/move carries extents; whole-axis edits shift/expand/shrink them. Sorting, fill and lane shifts through merged areas require unmerge. Covered coordinates cannot receive independent content. The grid projects spans, expands selections and navigates past covered cells; calculations and Python reads see only the upper-left content. Python spill materialization rejects merged destinations.

## Number formatting

NumberFormats in Foundation owns preset/category catalogs, format classification, validation and decimal-code transformations. Authored CellStyle.NumberFormat remains the single format source for display and persistence. The Number/Alignment/Font/Border/Fill Format Cells UI edits a local code draft, previews through EverydayCommands.Display and applies to a captured selection/snapshot only after OK. Decimal changes preserve literal text, grouping, exponent and negative sections. The renderer derives section colors from the formatted value and preserves authored font colors.

## Date/time locales and editing

TemporalFormats in Core handles locale tags, readable temporal edit text and conversion back to serials. TemporalFormatOptions in Foundation supplies Gregorian locale/type catalogs. CellSnapshot.EditText is separate from canonical Source, so formula/copy serialization stays canonical. Inline editing and formula-bar entry share GridClipboard.EntryText/ParseEntry; accepting an unchanged temporal edit preserves the exact original double and creates no new Undo entry. Formula and Python anchors retain their authored source.

## Cell editing and Special formats

The inline editor distinguishes caret editing from direct typing. Format Cells restores worksheet focus when it closes. Cell context formatting validates the captured snapshot and selection before opening. SpecialFormatOptions supplies an explicit country catalog, locale-tagged Excel codes and style classification; no new numeric type or persistence field is needed.

## Font and Fill formatting

CellStyle adds strikethrough, script, accounting underline kind and optional CellFill details (pattern/color or two-color shading/variant). These participate in identity, history, style copying and native persistence. CellAppearanceDraft tracks edited property groups so changing one Font/Fill control preserves unrelated mixed-selection fields. Number formatting is applied only when changed. Native previews and Win2D render fill effects; XLSX uses ClosedXML for font/pattern styles and a bounded XML pass for gradients that ClosedXML cannot represent.

## Format Cells control layout

FormatDialogControls supplies native tab, list, caption-frame and accessibility-only naming helpers. NumberFormatDialogOptions owns fraction labels, regional negative examples, separator labels and inline RGB validation. Numeric choices store standard locale-tagged Excel codes; matching the current locale uses current Windows separator overrides for display. FormatColorSelector shares the inline palette/hex flow across Font, Fill and gradient colors without nested color-tab dialogs. A dialog completion flag covers final publication/focus restoration; restoring focus preserves an already-open F2 editor, and editor refocus cancels a pending blur commit.

## Shared dialog tabs and field bounds

`DialogTabs.ApplyDialogTabStyle` is shared by Find/Replace and Format Cells, including native TabView colors/borders, focus visuals and suppression of inherited shortcut tooltips. DialogTab constructs scrolling bodies before they are attached to a TabViewItem. Active panes are constrained to the realized tab-content presenter width; long pane bodies scroll vertically with symmetric horizontal insets. Format Cells places Font labels and input surfaces on separate shared rows; compact NumberBox fields explicitly remove conflicting minimum widths while retaining native spin behavior.

Format Cells fixes its tab viewport before opening so category/tab changes cannot resize the dialog. Only the initial pane is built before opening. Other panes are built on first selection and retained until the dialog closes, sharing the dialog's captured selection and formatting drafts. Tab bodies share the Find/Replace top gap and symmetric insets. Pattern Style uses a native dropdown button with a six-column graphical palette. Fill thumbnails and samples share a renderer: solids, tiled patterns and gradients use explicit BGRA bitmaps at the realized image dimensions to avoid native SVG paint/aspect-ratio inconsistencies. Scroll bodies are constrained to their actual viewport width, with symmetric insets.

FillEffectsEditor owns a temporary two-color gradient draft, six exclusive shading radio buttons and exclusive visual variant buttons (three linear, four corner, one center). Shading and variants sit side by side above the final sample; a scrollable flyout with shared dialog outline/title dragging keeps controls reachable in small windows. Flyout OK updates the parent draft; its Cancel discards only the gradient edit.

DialogDragging supports ContentDialog and the Fill Effects overlay surface. Native NumberBox spin reparenting is idempotent across Loaded events. Font group boxes share the upper field column geometry; rectangular corner/center gradients use matching preview and worksheet rendering.


CellStyle.BorderDetails stores six nullable CellBorder edges with a line style and optional RGB (Automatic). CellBorders.From adapts legacy thin/text-colored Borders. BorderFormatDraft records edge operations and applies outer edges to the selection perimeter and inside edges to both sides of internal boundaries. It joins the existing Format Cells transaction. Native persistence and authored identity retain the details; XLSX maps them to border styles/colors and shared diagonal ink. Ribbon border commands clear detailed borders before applying their existing presets.

BorderFormatEditor shares selection aggregation, preview shape, pen choices and enabled-edge state between controls and the applied draft. Border style/color changes preserve the other edge property and do not overwrite mixed edges. Preview and worksheet borders resolve Automatic ink from the active theme. Win2D draws the stroke pattern at a consistent scale across cell sizes.

BorderCommands builds atomic preset and stroke commands from canonical physical border segments, updating both neighboring cell faces while preserving contents and unrelated styles; merged interiors are skipped and visible merged edges map to the anchor. The worksheet host owns the drawing mode and pen for the active session and sheet. Win2D handles pointer capture, drag previews, eraser paths and cancellation. A completed gesture validates its starting snapshot and applies one command. Drawing and Format Painter are mutually exclusive. No tool state is persisted in workbooks.

The Borders split button remembers its last action separately from active drawing state, so cancellation never resets the primary action. Border line feedback and committed segments use the same gesture classification. The worksheet loads Windows cursor assets for the pencil, grid-pencil and eraser and restores the normal pointer when the tool ends.

Alignment styles additionally store eight horizontal modes (null means General), signed −90…90 rotation or 255 for stacked text, 0–250 indentation, shrink-to-fit, reading order and justify-last-line. Validation, authored identity and native/XLSX exchange include these fields. `CellAppearanceDraft` applies only edited alignment fields. `AlignmentCommands.Format` combines optional merge/unmerge and dialog formatting into one detached, validated structural proposal with content-loss reporting, preserving cell identities and caches. Projection carries the actual value kind so General aligns numbers right, booleans/errors centrally and text at the reading-order start. Win2D renders rotation, stacked text, indentation, shrink, fill repetitions, horizontal/vertical distribution and Center Across Selection without changing stored values or merge geometry. Win2D uses DirectWrite text metrics.

Formula reference picking handles worksheet pointer gestures before ordinary cell selection. It retains the original edit target and updates the reference as the pointer moves. The formula bar owns its draft and hint through the drag and regains keyboard focus when it ends. The inline editor owns drafts started in the worksheet. Both font surfaces use the same installed-font catalog.
