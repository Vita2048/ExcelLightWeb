# Expansion Plan

This document defines the cross-phase architecture contract, permanent exclusions, documentation protocol, and detailed sprint breakdown for the active phase of `ExcelLightWeb` expansion. Detailed sprint plans for later phases are added once the preceding phase completes.

Live sprint progress, verification counts, and active risks live in [ExpansionPlanStatus.md](ExpansionPlanStatus.md).

---

## 1. Phase Overview

1. **Phase 1 — Formula Engine + Core Functions** *(Complete; Section 4)*: Pure TypeScript engine (`src/calculation`) with dynamic array spilling, `LET`/`LAMBDA`, structured table references, 3D references, Windows Excel COM differential oracle (`16.0.17932`), 225 implemented functions, Function Library ribbon & parameter dialogs, and F2 color-coded range highlights.
2. **Phase 2 — PivotTables** *(Next; detailed in Section 5)*: Core immutable PivotTable model, aggregation engine, Insert → PivotTable dialog, docked dark PivotTable Fields side pane, grid rendering, in-sheet filter/sort dropdowns, contextual `PivotTable Analyze` & `Design` ribbon tabs, `GETPIVOTDATA`, and `.exlt`/XLSX handling.
3. **Phase 3 — Charts (Full Excel Variety)** *(Planned; detail after Phase 2)*: Floating interactive charts bound to ranges, Tables, dynamic spills, and PivotTables, with Format Chart / Select Data side panes and dark-theme parity across Excel's chart families.
4. **Phase 4 — Remaining Long-Tail Excel Functions** *(Planned; can batch after Phase 1)*: Remaining ~300+ Excel functions in categorized waves verified against the Excel COM oracle.

### Permanent Formula Exclusions
Excluded permanently from native formula support:
`WEBSERVICE`, `STOCKHISTORY`, `RTD`, `CUBE*` (×7), `FIELDVALUE`, `PY` *(handled by the existing Python engine)*, `COPILOT`, `TRANSLATE`, `DETECTLANGUAGE`, `CALL`, `REGISTER.ID`, `IMPORTCSV`, `IMPORTTEXT`, `IMAGE`.

---

## 2. Cross-Phase Contract (Phases 1–4)

Three shared invariants must hold across every phase so earlier work does not cause rework later:

1. **Unified Generated-Region / Spill Ownership (`src/core`)**
   - Generalize generated-cell ownership beyond Python (`'python' | 'formula' | 'pivot'`) via a lightweight kind/region contract on `SheetDocument` / `CellDocument`.
   - All generated cells are read-only, excluded from authored-content equality when derived, and cleaned up atomically when their owner shrinks, moves, or is cleared.
   - Collision checks (`#SPILL!` / Pivot overlap) uniformly guard against populated cells, merges, Tables, Python spills, formula spills, and PivotTables.
2. **Reusable Excel COM Differential Oracle (`scripts/` + `tests/fixtures/formulas/`)**
   - One Windows Excel COM generator script (`scripts/generate-excel-oracle.ps1`) captures `{ type, value, spillMatrix }` using `Formula2` and `CalculateFullRebuild()` plus the Excel build version into category JSON fixtures.
   - CI runs Vitest (`tests/parity/native-formulas.test.ts`) against those checked-in fixtures without needing Excel installed. The same harness serves Phase 1, Phase 4, and optional Phase 2 aggregation cross-checks.
3. **Additive `.exlt` v1 Persistence & Reviewed XLSX Boundaries (`src/persistence`)**
   - All new workbook fields (formula spill metadata in Phase 1, PivotTables in Phase 2, Charts in Phase 3) must remain optional/additive on `.exlt` format version 1 so existing workbooks load cleanly.
   - XLSX export materializes cached scalar/spilled values and continues reporting unsupported features in the reviewed loss report rather than silently dropping data.

---

## 3. Execution & Documentation Protocol

### Per-Sprint Execution Loop
1. The user requests one sprint or slice (e.g., *"Implement Sprint 1.1"*).
2. Implement **only** that sprint, keeping `main` releasable at the end of every sprint.
3. Run the sprint's scoped verification commands (Vitest / Excel oracle / `npm run check` / Playwright).
4. Update [ExpansionPlanStatus.md](ExpansionPlanStatus.md) with status, exact pass/fail/skipped counts, and open risks, plus any triggered product docs below.
5. Report exact verification results and stop for review before advancing.

### Doc Update Cadence

| Document | When to update |
| --- | --- |
| [ExpansionPlanStatus.md](ExpansionPlanStatus.md) | **Every sprint** (status, exact test/oracle counts, risks). |
| [Features.md](Features.md) | **Every sprint that changes user-visible behavior** (new functions, spills, UI, persistence) with test paths. Skip on pure internal refactors. |
| [Architecture.md](Architecture.md) | **Only when contracts change** (model schema, spill ownership, worker protocol, `.exlt`/XLSX format). |
| [Development.md](Development.md) | **Only when commands or test workflows change** (e.g., adding the Excel COM oracle script). |
| [Known-Issues.md](Known-Issues.md) | **Whenever an exclusion/limitation is resolved or a new one is discovered.** |
| [Deployment.md](Deployment.md) & [dependencies.md](dependencies.md) | **Only when build/CI/Pages steps or `package.json` dependencies change.** |

Also perform a full cross-doc consistency check at the end of each Phase.

---

## 4. Phase 1 Detailed Sprints: Formula Engine + Core Functions

### Sprint 1.1 — Excel COM Oracle Harness & Function Registry Refactor
- **Oracle generator (`scripts/generate-excel-oracle.ps1`)**: Drive desktop `Excel.Application` via COM (1900 date system, `Formula2`, `CalculateFullRebuild()`), capture scalar results and 2D `SpillingToRange` matrices with Excel build provenance into `tests/fixtures/formulas/*.json` (plus an `.xlsx` cached-value fallback extractor).
- **Modularize registry (`src/calculation/functions/`)**: Split `src/calculation/formula-functions.ts` into category modules (`logical.ts`, `math.ts`, `text.ts`, `datetime.ts`, `lookup.ts`, `statistical.ts`, `financial.ts`, `dynamic-array.ts`, `info.ts`, `registry.ts`), keeping `nativeFunctions` / `supportedFunctions` re-exported from `src/calculation/formula-functions.ts` and co-locating arity, volatile flags, and formula-bar signature hints.
- **Error catalog (`src/core/workbook.ts`, `src/core/formula-references.ts`)**: Add `#SPILL!` and `#CALC!` to `formulaErrors` and parser error tokens.
- **Verification**: `npm run check` passes with existing 31 functions unchanged; smoke-run `scripts/generate-excel-oracle.ps1` and verify `tests/parity/native-formulas.test.ts` asserts against it.

### Sprint 1.2 — Parser & Reference System Upgrade
- **Syntax & AST (`src/core/formula-references.ts`)**:
  - Spill range operator (`A1#`, `Sheet2!B2#`).
  - Implicit intersection (`@A1:A10`, `@Table1[Col]`).
  - General range (`:`) and intersection (space) operators with reference-returning calls (e.g., `INDEX(...):INDEX(...)`).
  - Higher-order call expressions (`LAMBDA(x, x*2)(5)`, lambda variables in `LET`).
  - Full structured table references (`Table1[Col]`, `Table1[[#Headers],[Col]]`, `Table1[[#Totals],[Col]]`, `Table1[[#Data],[Col1]:[Col2]]`, `[@Col]`, `Table1[@[Col]]`).
  - 3D sheet references (`Sheet1:Sheet3!A1:B5`).
- **Structural rewrites**: Preserve `#`, `@`, structured refs, and 3D refs across copy/cut/fill, row/column insert/delete, sheet move/rename/delete, and table/column rename.
- **Verification**: New unit tests (`tests/unit/formula-parser.test.ts`) plus `tests/integration/structure-commands.test.ts` and `tests/integration/editing-commands.test.ts`.

### Sprint 1.3 — Dynamic Array Spilling, Broadcasting & `LET` / `LAMBDA`
- **Broadcasting & `@` (`src/calculation/formula-values.ts`, `src/calculation/native-calculator.ts`)**: Elementwise broadcasting across scalars, 1×N, N×1, and M×N arrays (`#N/A` on mismatched shapes) plus `@` implicit intersection.
- **Spill lifecycle (`src/core/python-spills.ts`, `src/application/workbook-protocol.ts`, `src/calculation/native-generation.ts`, `src/calculation/mixed-generation.ts`)**:
  - Generalize `src/core/python-spills.ts` to shared generated-cell ownership (`'python' | 'formula' | 'pivot'`) and extend `CacheUpdate` with optional spill payloads committed atomically without advancing `authoredRevision` or adding undo history.
  - Cache `#SPILL!` on collision with populated cells, merges, Tables, or Python spills, clearing stale owned cells.
  - Track `A1#` spill dependencies and replan across native and mixed Python/native generations.
  - Render spilled formula cells read-only in `src/grid` with Excel's spill-range indicator border on selection.
- **Lexical scoping & volatile/dynamic refs**: Implement `LET`, `LAMBDA`, `ISOMITTED`, `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`, `MAKEARRAY` (with recursion depth guard) and dynamic/volatile dependency tracking (`NOW`, `RAND`, `RANDBETWEEN`, `RANDARRAY`, `OFFSET`, `INDIRECT`, `A1#`).
- **Verification**: Excel COM oracle fixture `tests/fixtures/formulas/dynamic-arrays.json`; integration tests for `#SPILL!`, undo/redo, Python↔native spill chaining, `.exlt` round-trip; Playwright e2e test for spill rendering, read-only protection, and spill border.

### Sprint 1.4 — Core Function Wave A: Lookup, Dynamic Arrays, Logical & Info (~50 functions)
- **Lookup & Reference**: `XLOOKUP`, `XMATCH`, `INDEX`, `MATCH`, `HLOOKUP`, `LOOKUP`, `CHOOSE`, `OFFSET`, `INDIRECT`, `ROW`, `ROWS`, `COLUMN`, `COLUMNS`, `ADDRESS`, `AREAS`, `FORMULATEXT`, `HYPERLINK` *(display text)*.
- **Dynamic Arrays & Shaping**: `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `RANDARRAY`, `VSTACK`, `HSTACK`, `TOROW`, `TOCOL`, `WRAPROWS`, `WRAPCOLS`, `TAKE`, `DROP`, `CHOOSEROWS`, `CHOOSECOLS`, `EXPAND`, `TEXTSPLIT`.
- **Logical**: `IFS`, `SWITCH`, `XOR`, `IFNA`, `TRUE`, `FALSE`.
- **Information**: `ISBLANK`, `ISERR`, `ISERROR`, `ISLOGICAL`, `ISNA`, `ISNONTEXT`, `ISNUMBER`, `ISTEXT`, `ISFORMULA`, `ISREF`, `ISEVEN`, `ISODD`, `ERROR.TYPE`, `N`, `NA`, `TYPE`.
- **Verification**: Excel COM oracle fixtures (`tests/fixtures/formulas/lookup.json`, `logical-info.json`, `dynamic-arrays.json`) with automated coverage gate requiring valid + error + edge cases for 100% of `nativeFunctions`.

### Sprint 1.5 — Core Function Wave B: Math, Aggregates, Text, Date/Time & Core Stats/Financial (~85 functions)
- **Conditional & General Aggregates**: `SUMIFS`, `COUNTIFS`, `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS`, `MINIFS`, `COUNTBLANK`, `SUMPRODUCT`, `AGGREGATE`, `PRODUCT`.
- **Math & Trig**: `INT`, `TRUNC`, `MOD`, `POWER`, `SQRT`, `SIGN`, `RAND`, `RANDBETWEEN`, `CEILING`, `CEILING.MATH`, `FLOOR`, `FLOOR.MATH`, `MROUND`, `EXP`, `LN`, `LOG`, `LOG10`, `PI`, `SIN`, `COS`, `TAN`, `ASIN`, `ACOS`, `ATAN`, `ATAN2`, `DEGREES`, `RADIANS`, `FACT`, `COMBIN`, `PERMUT`, `GCD`, `LCM`, `QUOTIENT`, `EVEN`, `ODD`, `SUMSQ`, `MMULT`, `TRANSPOSE`.
- **Text**: `CONCAT`, `CONCATENATE`, `TEXTJOIN`, `TEXT`, `VALUE`, `NUMBERVALUE`, `FIND`, `SEARCH`, `SUBSTITUTE`, `REPLACE`, `REPT`, `PROPER`, `EXACT`, `CLEAN`, `CHAR`, `CODE`, `UNICODE`, `UNICHAR`, `T`, `TEXTBEFORE`, `TEXTAFTER`, `VALUETOTEXT`, `ARRAYTOTEXT`, `FIXED`, `DOLLAR`.
- **Date & Time**: `NOW`, `TIME`, `HOUR`, `MINUTE`, `SECOND`, `DATEVALUE`, `TIMEVALUE`, `EDATE`, `EOMONTH`, `DATEDIF`, `DAYS`, `DAYS360`, `WEEKDAY`, `WEEKNUM`, `ISOWEEKNUM`, `WORKDAY`, `WORKDAY.INTL`, `NETWORKDAYS`, `NETWORKDAYS.INTL`, `YEARFRAC`.
- **Core Statistical & Financial**: `MEDIAN`, `MODE`, `MODE.SNGL`, `STDEV.S`, `STDEV.P`, `VAR.S`, `VAR.P`, `LARGE`, `SMALL`, `RANK`, `RANK.EQ`, `RANK.AVG`, `PERCENTILE.INC`, `PERCENTILE.EXC`, `QUARTILE.INC`, `QUARTILE.EXC`, `CORREL`, `COVARIANCE.S`, `SLOPE`, `INTERCEPT`, `FORECAST.LINEAR`, `PMT`, `PV`, `FV`, `NPV`, `XNPV`, `IRR`, `XIRR`, `RATE`, `NPER`, `IPMT`, `PPMT`, `SLN`, `DB`, `DDB`.
- **Verification**: Excel COM oracle fixtures (`math.json`, `text.json`, `datetime.json`, `stats-financial.json`), `_xlfn.` XLSX import/export tests, `npm run test:performance` (warm 1,000-formula recalc p95 < 500 ms; 10,000-cell edit p95 < 100 ms), full `npm run check`, `npm run test:e2e`, `npm run test:worksheet`, and Phase 1 sync of [Architecture.md](Architecture.md), [Features.md](Features.md), and [Known-Issues.md](Known-Issues.md).

---

## 5. Phase 2 Detailed Sprints: PivotTables (Excel UI)

All PivotTable UI flows follow desktop Excel conventions adapted to ExcelLightWeb's dark palette, forced-colors accessibility, React/canvas split, and session command transactions. Screenshots of Excel's PivotTable dialogs/panes/menus can be used to lock pixel-exact menu labels before executing each UI sprint.

### Sprint 2.1 — PivotTable Core Model, Aggregation Engine & `.exlt` Persistence
- **Core model (`src/core/pivot.ts`, `src/core/workbook.ts`)**:
  - Immutable `PivotTableDocument` definitions on `WorkbookDocument` / `SheetDocument` (stable ID, unique case-insensitive name, source range or named Table reference, target anchor cell, `PivotFieldConfig[]` across `filters` / `columns` / `rows` / `values`, layout options, subtotals/grand-totals flags, collapsed keys, filter exclusions, sort rules, and PivotTable style ID).
  - Reuse the shared generated-region ownership (`spillKind: 'pivot'`) so PivotTable cells are read-only, excluded from authored equality, guarded against overlapping populated cells/merges/Tables/formula spills/other PivotTables, and atomically cleaned on resize or clear.
- **Aggregation engine (`src/calculation/pivot-engine.ts`)**:
  - Pure cached field extraction from source sheets/Tables, multi-key grouping for Rows and Columns, Report Filter slicing, and `Values` placement (`Values` virtual field in Rows or Columns).
  - Core aggregations (`Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `Count Numbers`, `StdDev`, `StdDevp`, `Var`, `Varp`) plus `Show Values As` (`No Calculation`, `% of Grand Total`, `% of Column Total`, `% of Row Total`, `% of Parent Row/Column Total`, `Difference From`, `% Difference From`, `Running Total In`, `Rank Smallest to Largest`, `Rank Largest to Smallest`, `Index`).
- **Session commands & persistence (`src/application/pivot-commands.ts`, `src/persistence/native/`)**:
  - Atomic session commands (`create-pivot`, `update-pivot-layout`, `refresh-pivot`, `toggle-pivot-collapse`, `filter-pivot`, `sort-pivot`, `move-pivot`, `resize-pivot-source`, `delete-pivot`) with full undo/redo and copy/move/sheet/axis reference updates.
  - Optional/additive `.exlt` v1 serialization preserving PivotTable definitions and cached layout without breaking existing `.exlt` files.
- **Verification**: Unit tests for grouping, aggregations, `Show Values As`, and collision/undo invariants + Excel COM oracle spot checks for PivotTable numeric aggregates + `npm run check`.

### Sprint 2.2 — Create PivotTable Dialog, Docked `PivotTable Fields` Pane & Value Field Settings
- **Ribbon entry (`Insert → Tables → PivotTable`)**: Opens the dark **Create PivotTable** modal dialog matching Excel:
  - *Choose the data that you want to analyze*: `Select a table or range` input with `...` worksheet range picker (auto-detecting the contiguous `currentRegion` or active Table around the active cell) and `OK` / `Cancel`.
  - *Choose where you want the PivotTable to be placed*: `New Worksheet` (default) or `Existing Worksheet` with `Location` range picker.
- **Docked dark `PivotTable Fields` side pane (`src/features/pivot/PivotTableFieldsPane.tsx`)**:
  - Appears docked on the right (resizable splitter, toggleable close `×` button) whenever the active cell is inside a PivotTable (if Field List is enabled).
  - Search box (`Search` fields), scrollable field list with checkboxes (check adds to `Rows` or `Values` for numeric columns matching Excel heuristic), and 2×2 drop zones (**Filters**, **Columns**, **Rows**, **Values**).
  - Both pointer drag-and-drop AND full keyboard item menus on each field chip (`Move Up`, `Move Down`, `Move to Beginning`, `Move to End`, `Move to Report Filter` / `Row Labels` / `Column Labels` / `Values`, `Remove Field`, `Value Field Settings...`).
  - Bottom `Defer Layout Update` checkbox + explicit `Update` button.
- **`Field Settings` / `Value Field Settings` modal (`src/features/pivot/FieldSettingsDialog.tsx`)**:
  - `Source Name` (read-only), `Custom Name`, tabs for `Summarize Values By` (11 functions) and `Show Values As` (+ Base Field / Base Item selectors), `Number Format...` button launching `FormatCellsDialog`'s Number tab, and `Subtotals & Filters` tab for row/column fields (`Automatic`, `None`, `Custom`).
- **Verification**: Keyboard/drag-and-drop/axe accessibility and Playwright journeys in Chromium creating PivotTables from ranges/Tables, moving fields across areas, changing aggregations/number formats, and closing/reopening the pane.

### Sprint 2.3 — Grid Layouts, Expand/Collapse, In-Sheet Filter/Sort & Drill-Down
- **Report layouts (`src/calculation/pivot-engine.ts`, `src/grid/draw-worksheet.ts`)**:
  - Support Excel's three layouts: **Compact Form** (default indented hierarchy with `Row Labels` / `Column Labels` headers), **Outline Form**, and **Tabular Form**, plus **Repeat All Item Labels**, **Blank Rows** after items, **Subtotals** (top/bottom/off), and **Grand Totals** (rows/columns/both/off).
  - Render PivotTable headers, subtotal/total bands, indented items, `+/-` outline buttons, and filter dropdown triggers cleanly on the Canvas 2D grid with forced-colors contrast support.
- **In-sheet Pivot interaction (`src/grid/Worksheet.tsx`, `src/features/pivot/PivotFilterDropdown.tsx`)**:
  - Clickable `+/-` expand/collapse toggles on row/column group items.
  - Report Filter row dropdowns above the PivotTable (`(All)`, single-select, or `Select Multiple Items` checkboxes + search).
  - `Row Labels` / `Column Labels` filter & sort dropdowns (`Sort A to Z`, `Sort Z to A`, `More Sort Options...`, field selector when multiple fields share compact column, Label/Value filter checklist & search).
  - Double-clicking a PivotTable data value cell runs **Show Details**, inserting a new worksheet populated with a formatted Excel Table of matching source records (undoable).
- **Verification**: Unit/integration tests for Compact/Outline/Tabular layouts and drill-down record matching; Playwright tests for expand/collapse, in-sheet report filters, sorting, and drill-down.

### Sprint 2.4 — Contextual `PivotTable Analyze` & `Design` Ribbon Tabs + Styles
- **Contextual ribbon chrome (`src/features/shell/Ribbon.tsx`, `desktop-glyphs.ts`)**:
  - Dynamically show green/accent-highlighted **PivotTable Analyze** and **Design** tabs when the active cell is inside a PivotTable (remembering active contextual tab per PivotTable during the session, restoring the prior ordinary tab when leaving).
  - **PivotTable Analyze** groups following Excel:
    - `PivotTable`: `PivotTable Name` editable box, `Options...` (layout/display/totals/zero-for-empty/autofit options dialog).
    - `Active Field`: `Active Field` indicator, `Field Settings`, `Expand Field` / `Collapse Field`.
    - `Data`: `Refresh` / `Refresh All`, `Change Data Source...` (with range picker).
    - `Actions`: `Clear` (`Clear All`, `Clear Filters`), `Select` (`Entire PivotTable`), `Move PivotTable...`.
    - `Show`: toggle buttons for `Field List`, `+/- Buttons`, and `Field Headers`.
  - **Design** groups following Excel:
    - `Layout`: `Subtotals`, `Grand Totals`, `Report Layout`, `Blank Rows` flyout menus with custom SVG glyphs matching desktop parity rules.
    - `PivotTable Style Options`: `Row Headers`, `Column Headers`, `Banded Rows`, `Banded Columns` checkboxes.
    - `PivotTable Styles`: Light/Medium/Dark-adapted swatch gallery respecting authored cell overrides and forced-colors mode.
- **Verification**: Extend `tests/fixtures/parity/ribbon-order.json` and `tests/e2e/ribbon-parity.spec.ts` for contextual PivotTable tab appearance, keyboard navigation, and axe checks.

### Sprint 2.5 — `GETPIVOTDATA`, Mixed Recalc Integration, XLSX Boundary & Phase 2 Verification
- **Formula integration (`src/calculation/functions/`)**:
  - Implement `GETPIVOTDATA(data_field, pivot_table, [field1, item1], ...)` returning aggregated scalar values or `#REF!` when missing/collapsed, registered across `Lookup & Reference` and `FunctionLibrary`.
  - Automatic dependency propagation: editing source data + refreshing the PivotTable triggers downstream formulas (including `GETPIVOTDATA` / `SUM` referencing Pivot cells) and downstream Python `xl(...)` cells in a single generation.
- **XLSX interchange boundary (`src/persistence/xlsx/`)**:
  - Export materialized PivotTable header/value cells and number formats cleanly so workbooks open accurately in desktop Excel without executing scripts, while adding explicit reviewed XLSX loss notices on PivotTable import/export (`pivotTableDefinition` not round-tripped natively).
- **Phase 2 Completion Gate**:
  - Run `npm run check`, Excel COM `GETPIVOTDATA` oracle fixtures, `npm run test:performance` (verify PivotTable refresh & 10k-row aggregation latency), `npm run test:e2e`, and `npm run test:worksheet`.
  - Update [Features.md](Features.md), [Architecture.md](Architecture.md), [Known-Issues.md](Known-Issues.md), and [ExpansionPlanStatus.md](ExpansionPlanStatus.md).
