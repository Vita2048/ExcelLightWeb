import { useEffect, useMemo, useRef, useState, type ClipboardEvent, type FormEvent, type KeyboardEvent, type MouseEvent, type PointerEvent as ReactPointerEvent } from 'react';
import type { EditRange } from '../application/editing-commands';
import { applyNumberFormat, type NumberFormatChange } from '../application/number-format-command';
import { changeAlignment, mergeState, proposeMerge, type AlignmentAction, type AlignmentPatch, type MergeMode } from '../application/alignment-commands';
import type { BorderDraft } from '../core/formats/borders';
import { MergeConfirmation } from '../features/formatting/MergeConfirmation';
import { applyBorderPreset, applyBorderStroke, inspectBorders, mergeAt } from '../application/border-commands';
import { borderModes, borderPen, type BorderAction, type BorderMode, type BorderPreset, type BorderStyle } from '../core/formats/borders';
import type { DrawGesture } from '../grid/border-geometry';
import { FormatCellsDialog, type FormatTab } from '../features/formatting/FormatCellsDialog';
import { activeFormatBrush, captureFormatBrush, formatStyles, paintStyles, type FormatBrush } from '../application/style-commands';
import { underlinePatch, type AppearancePatch } from '../core/formats/appearance';
import { changeDimensions, changeSheet, editAxis } from '../application/structure-commands';
import {
  columnValues, currentRegion, filterRange, filteredRows, headerLabel, localeTextOrder, looksLikeHeader, rangeLabel,
  recordSummary, setFilter, sortData, sortKeyFor, type FilterValue,
} from '../application/data-commands';
import {
  emptySearchHistory, findMatches, matchLabel, rememberSearch, replaceMatches, searchResultRows,
  type SearchHistory, type SearchLook, type SearchMatch, type SearchOptions,
} from '../application/search-commands';
import { freezeFromCaret, setFreeze } from '../application/view-commands';
import { FindReplaceDialog, type SearchOutcome } from '../features/search/FindReplaceDialog';
import { ZoomDialog } from '../features/view/ZoomDialog';
import { columnLetters } from '../core/formula-references';
import type { RowMask } from '../core/row-mask';
import { FilterDropdown } from '../features/data/FilterDropdown';
import { SortDialog, SortOptionsDialog, SortWarningDialog, type SortDraft, type SortExtent } from '../features/data/SortDialog';
import type { WorkbookSnapshot } from '../application/workbook-protocol';
import { loadNativePackage, saveNativePackage } from '../persistence/native';
import {
  defaultWorkspacePreferences, LocalWorkspaceError, LocalWorkspaceStore, WorkspaceConflictError, WorkspaceLease,
  type LocalWorkspacePreferences, type LocalWorkspaceRecord, type RecoveryRecord, type WorkspaceFileHandle,
} from '../persistence/local';
import { downloadNativePackage, downloadXlsxFile, fileVersion, hasNativeSavePicker, hasXlsxSavePicker,
  pickNativeOpenFile, pickNativeSaveFile, pickXlsxSaveFile, sameFileEntry, writeNativeFile,
  type OpenedNativeFile } from '../persistence/local';
import { importXlsx, inspectXlsxExport, writeXlsx, type ExchangeIssue } from '../persistence/xlsx';
import { FileBackstage, type FileAction } from '../features/files/FileBackstage';
import { FileMessageDialog, OptionsDialog, RecoveryDialog, RemoveWorkspaceDialog, SaveAsDialog, UnsavedChangesDialog,
  XlsxReportDialog, type ReplacementChoice, type ReportChoice } from '../features/files/FileDialogs';
import {
  changeTableDesign, createTable, defineName, deleteName, deleteTableSelection, nextTableName, parseNameExtent,
  refersTo, renameTable, resizeTable, selectionHasHeaders, setTotalFunction,
} from '../application/table-commands';
import {
  changePivotSource, clearPivotTable, createPivotTableWithSheet, deletePivotSelection, drillDownPivotCell,
  expandCollapsePivotField, findPivotAt, movePivotTableWithSheet, parsePivotRangeReference, parsePivotTargetCell,
  refreshAllPivotTables, refreshPivotTable, renamePivotTable, togglePivotCollapse, updatePivotLayout,
} from '../application/pivot-commands';
import { computePivotOverlays, resolveActivePivotField } from '../calculation/pivot-engine';
import { CreatePivotTableDialog, type PivotPlacement } from '../features/pivot/CreatePivotTableDialog';
import { FieldSettingsDialog, type FieldSettingsRequest } from '../features/pivot/FieldSettingsDialog';
import { ChangePivotSourceDialog, MovePivotTableDialog, PivotOptionsDialog } from '../features/pivot/PivotDialogs';
import { PivotFilterDropdown } from '../features/pivot/PivotFilterDropdown';
import { PivotTableFieldsPane } from '../features/pivot/PivotTableFieldsPane';
import type { PivotAxisField, PivotTableDocument, PivotValueField } from '../core/pivot';
import type { ChartPlacement, ChartType } from '../core/chart';
import { createChart, deleteChart, moveResizeChart, updateChart } from '../application/chart-commands';
import { formatAbsoluteAreaRef, resolveChartData, resolveFormulaExtent } from '../charts/chart-data';
import { chartRangeHighlights } from '../charts/chart-highlights';
import { InsertChartDialog } from '../features/charts/InsertChartDialog';
import type { CoreChartCategoryTab } from '../features/charts/chart-catalog';
import { extentContains, lastColumn, lastRow, ownsTableFilter, tableDesign, tableInputExtent, tablesOnSheet } from '../core/workbook-names';
import { type DefinitionDraft, NameDefinitionDialog, NameManagerDialog, RangePickerWidget } from '../features/tables/TableDialogs';
import type { FigurePlacement, PythonFigure, TableDesign, WorkbookName } from '../core/workbook';
import { putFigurePlacement, removeFigurePlacement } from '../application/figure-commands';
import { OfflineStatus } from '../offline/OfflineStatus';
import { WorkbookSession, type SetCellsCommand, type StructureCommand, type WorkbookCommand } from '../application/workbook-session';
import type { Axis } from '../core/axis-edit';
import type { NativeFunction } from '../calculation/formula-functions';
import { NativeScheduler, type CalculationMode, type CalculationStatus } from '../calculation/native-scheduler';
import { documentLimits, findCell, findReferencedCell, newWorkbook, resolvedCellStyle, worksheetViewport, type CellReference, type Coordinate, type DocumentId, type WorkbookDocument } from '../core/workbook';
import { requireDocument } from '../core/workbook-validation';
import { Ribbon, type SheetView } from '../features/shell/Ribbon';
import { CellEditor, type CellEditorProps } from '../features/editing/CellEditor';
import { PasteTextDialog } from '../features/editing/PasteTextDialog';
import { localeArgumentSeparator } from '../features/formulas/formula-presentation';
import { InsertFunctionDialog } from '../features/formulas/InsertFunctionDialog';
import { FunctionArgumentsDialog } from '../features/formulas/FunctionArgumentsDialog';
import { formulaReferenceHighlights } from '../features/formulas/formula-highlights';
import { formatArgumentReference, initialFunctionArguments, recordRecentFunction, sanitizeRecentFunctions, type InsertFunctionFilter } from '../features/formulas/function-library';
import { useWorksheetEditing, type AfterEditKey, type EditAction, type EditSurface } from '../features/editing/useWorksheetEditing';
import { Worksheet, type WorksheetHandle } from '../grid/Worksheet';
import type { WorksheetFigure } from '../grid/Worksheet';
import { PythonPanel } from '../features/python/PythonPanel';
import { pythonAnchorAt, type PythonAnchorTarget } from '../features/python/python-editor-model';
import { usePythonEditor } from '../features/python/usePythonEditor';
import { usePythonExecution } from '../features/python/usePythonExecution';
import { PythonTrustDialog } from '../features/python/PythonTrustDialog';
import { pythonInsertCommand, pythonResultText } from '../python/python-results';
import { buildPythonInputs } from '../python/python-inputs';
import { figureCacheKey, figureCanvasForPlacement } from '../python/figure-rendering';
import { StructureDialog, StructureMenu, type DialogAction, type StructureAction, type StructureInput } from '../features/structure/StructureControls';
import { fitDimensions } from '../grid/autofit';
import { address, clamp, gridMetrics } from '../grid/grid-geometry';
import {
  includes, mergeSelection, parseAddress, primaryArea, selectionAddress, selectionAreas, selectionStatistics, singleSelection, type Area, type GridSelection,
} from '../grid/grid-selection';

const defaultSelection = singleSelection();
const defaultView: SheetView = { zoom: 100, headings: true, gridlines: true, formulaBar: true, showFormulas: false };
const statisticsFormat = new Intl.NumberFormat('en-US', { maximumSignificantDigits: 10, useGrouping: false });
const makeId = () => crypto.randomUUID();
const supportSlaTemplateUrl = new URL('../../docs/showcase/Support-SLA.exlt', import.meta.url).href;
async function acquireRecoveryLease(lease: WorkspaceLease, workbookId: string): Promise<boolean> {
  // A reload can race the previous page's asynchronous pagehide lease release.
  // Give that transaction time to commit while still refusing a workbook held by another tab.
  for (let attempt = 0; attempt < 10; attempt++) {
    if (await lease.switchTo(workbookId)) return true;
    await new Promise(resolve => window.setTimeout(resolve, 100));
  }
  return false;
}
interface StructureTarget {
  readonly id: string;
  readonly ranges: readonly EditRange[];
  readonly header?: number;
}
const dialogActions: ReadonlySet<StructureAction> = new Set(['insert-cells', 'delete-cells', 'delete-sheet', 'place-sheet', 'size-row', 'size-column']);

interface ActiveFileContext {
  readonly name: string;
  readonly source: 'untitled' | 'native' | 'xlsx' | 'recovery';
  /** Imported XLSX provenance survives a native Save As for this session, so Export stays XLSX. */
  readonly importedXlsx?: boolean;
  readonly sourceFileName?: string;
  readonly fileHandle?: WorkspaceFileHandle;
  /** Imported XLSX file identity only; XLSX Export refuses to write over this file. */
  readonly sourceFileHandle?: WorkspaceFileHandle;
  readonly fileSize?: number;
  readonly fileLastModified?: number;
}

interface PendingFigureResize {
  readonly ownerId: string;
  readonly requested: FigurePlacement;
  readonly original: FigurePlacement;
  readonly target: PythonAnchorTarget;
  readonly baseRevision: number;
  readonly runningObserved: boolean;
}

function lowerBound(indexes: readonly number[], target: number): number {
  let low = 0;
  let high = indexes.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (indexes[middle]! < target) low = middle + 1;
    else high = middle;
  }
  return low;
}

function hiddenIndexes(dimensions: WorkbookDocument['sheets'][number]['rows']): readonly number[] {
  return dimensions.filter(dimension => dimension.hidden).map(dimension => dimension.index)
    .sort((left, right) => left - right);
}

function visibleIndex(index: number, manual: readonly number[], limit: number,
  filtered?: RowMask): number {
  const manualHas = (row: number) => {
    const at = lowerBound(manual, row);
    return at < manual.length && manual[at] === row;
  };
  const skipManual = (row: number, direction: 1 | -1) => {
    if (!manualHas(row)) return row;
    let at = lowerBound(manual, row);
    if (direction > 0) {
      let end = manual[at]!;
      while (at + 1 < manual.length && manual[at + 1] === end + 1) end = manual[++at]!;
      return end + 1;
    }
    let start = manual[at]!;
    while (at > 0 && manual[at - 1] === start - 1) start = manual[--at]!;
    return start - 1;
  };
  if (!manualHas(index) && !filtered?.has(index)) return index;
  const seek = (start: number, direction: 1 | -1, bound: number) => {
    let cursor = start;
    const guardLimit = manual.length + (filtered?.exceptions ?? 0) + 2;
    for (let guard = 0; guard < guardLimit; guard++) {
      if (direction > 0 ? cursor > bound : cursor < bound) return undefined;
      const opened = filtered ? direction > 0 ? filtered.rowAtOrAfter(cursor) : filtered.rowAtOrBefore(cursor) : cursor;
      if (direction > 0 ? opened > bound : opened < bound) return undefined;
      const landed = skipManual(opened, direction);
      if (landed === opened) return opened;
      if (direction > 0 ? landed <= cursor : landed >= cursor) return undefined;
      cursor = landed;
    }
    return undefined;
  };
  return seek(index + 1, 1, limit) ?? seek(index - 1, -1, 1) ?? index;
}

function NameBox({ selection, goTo, cancel }: {
  selection: GridSelection;
  goTo: (value: string) => boolean;
  cancel: () => void;
}) {
  const [draft, setDraft] = useState<string>();
  const selected = selectionAddress(selection);
  function submit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const input = event.currentTarget.elements.namedItem('address') as HTMLInputElement;
    if (goTo(input.value)) {
      setDraft(undefined);
    } else {
      input.setCustomValidity(`Enter a visible cell address from A1 to ${worksheetViewport.address.slice(3)}.`);
      input.reportValidity();
    }
  }
  return (
    <form onSubmit={submit}>
      <input className="name-box" name="address" aria-label="Name box" title="Go to a cell address"
        value={draft ?? selected} onFocus={event => event.currentTarget.select()}
        onChange={event => { event.currentTarget.setCustomValidity(''); setDraft(event.currentTarget.value); }}
        onBlur={event => { event.currentTarget.setCustomValidity(''); setDraft(undefined); }}
        onKeyDown={event => {
          if (event.key === 'Escape') {
            event.currentTarget.setCustomValidity('');
            cancel();
          }
        }} />
    </form>
  );
}

type Icon = 'save' | 'undo' | 'redo' | 'previous' | 'next';
const iconPath: Record<Icon, string> = {
  save: 'M3 2.5h11.5L18 6v11.5H2V2.5h1Z M5 2.5v5h9v-5 M5 17.5v-6h10v6',
  undo: 'm6 5-4 4 4 4 M2 9h11a4 4 0 0 1 0 8H9',
  redo: 'm14 5 4 4-4 4 M18 9H7a4 4 0 0 0 0 8h4',
  previous: 'm12 5-5 5 5 5',
  next: 'm8 5 5 5-5 5',
};
function ShellIcon({ icon }: { icon: Icon }) {
  return <svg width="17" height="17" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="1.5"
    strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d={iconPath[icon]} /></svg>;
}

export function App({ initialWorkbook }: { initialWorkbook?: WorkbookDocument }) {
  const [session] = useState(() => new WorkbookSession(initialWorkbook ?? newWorkbook(makeId), { makeId }));
  const [snapshot, setSnapshot] = useState(() => session.capture());
  const [localStore] = useState(() => new LocalWorkspaceStore());
  const [preferences, setPreferences] = useState<LocalWorkspacePreferences>(defaultWorkspacePreferences);
  const [preferencesReady, setPreferencesReady] = useState(false);
  const [activeFile, setActiveFile] = useState<ActiveFileContext>({ name: 'Book1', source: 'untitled' });
  const [recentWorkbooks, setRecentWorkbooks] = useState<readonly LocalWorkspaceRecord[]>([]);
  const [backstageOpen, setBackstageOpen] = useState(false);
  const [fileBusy, setFileBusy] = useState(false);
  const fileBusyRef = useRef(false);
  const recoveryEpoch = useRef(0);
  const [fileStatus, setFileStatus] = useState('');
  const [recoveryStatus, setRecoveryStatus] = useState('');
  const [storageError, setStorageError] = useState('');
  const [pendingRecovery, setPendingRecovery] = useState<RecoveryRecord>();
  const [workspaceConflict, setWorkspaceConflict] = useState('');
  const [replacementPrompt, setReplacementPrompt] = useState(false);
  const replacementResolver = useRef<((choice: ReplacementChoice) => void) | undefined>(undefined);
  const [optionsOpen, setOptionsOpen] = useState(false);
  const optionsResolver = useRef<((options: { readonly nativeCalculation: CalculationMode; readonly pythonCalculation: CalculationMode } | null) => void) | undefined>(undefined);
  const [saveAsOpen, setSaveAsOpen] = useState(false);
  const [saveNamePurpose, setSaveNamePurpose] = useState<'browser-save' | 'local-save-as'>('local-save-as');
  const saveAsResolver = useRef<((name: string | null) => void) | undefined>(undefined);
  const [workspaceToRemove, setWorkspaceToRemove] = useState<LocalWorkspaceRecord>();
  const [fileMessage, setFileMessage] = useState<{ title: string; message: string }>();
  const [xlsxReport, setXlsxReport] = useState<{ title: string; issues: readonly ExchangeIssue[] }>();
  const [selectedChartId, setSelectedChartId] = useState<DocumentId | undefined>(undefined);
  const [insertChartDraft, setInsertChartDraft] = useState<{
    readonly tab: 'recommended' | 'all';
    readonly category?: CoreChartCategoryTab;
    readonly sourceFormula: string;
  } | undefined>(undefined);
  const xlsxReportResolver = useRef<((choice: ReportChoice) => void) | undefined>(undefined);
  const workspaceLease = useRef<WorkspaceLease | undefined>(undefined);
  const workspaceConflictRef = useRef<(id: string) => void>(() => {});
  workspaceConflictRef.current = id => {
    setWorkspaceConflict(id);
    setFileStatus('This workbook is open in another tab. Editing is paused to protect its local copy.');
  };
  if (!workspaceLease.current) workspaceLease.current = new WorkspaceLease(localStore, id => workspaceConflictRef.current(id));
  const [calculation, setCalculation] = useState<CalculationStatus>(() => ({ mode: 'automatic', pending: false,
    calculating: new Set(), stale: new Set() }));
  const scheduler = useRef<NativeScheduler>(null);
  const [locale] = useState(() => navigator.language);
  const separator = useMemo(() => localeArgumentSeparator(locale), [locale]);
  useEffect(() => {
    if (!preferencesReady) return;
    const native = new NativeScheduler(session, (snapshot, status) => {
      setSnapshot(snapshot);
      setCalculation(status);
    }, locale);
    scheduler.current = native;
    native.setMode(preferences.nativeCalculation);
    if (session.active.sheets.some(sheet => sheet.cells.some(cell => cell.content.kind === 'formula'))) native.changed(false);
    return () => { native.dispose(); scheduler.current = null; };
  }, [session, locale, preferencesReady]);
  const sheets = snapshot.document.sheets;
  const [activeSheetId, setActiveSheetId] = useState(() => sheets[0]!.id);
  const [selections, setSelections] = useState<ReadonlyMap<DocumentId, GridSelection>>(() => new Map());
  const [view, setView] = useState(defaultView);
  const worksheet = useRef<WorksheetHandle>(null);
  const root = useRef<HTMLDivElement>(null);
  const requestedFocus = useRef(false);
  const [menu, setMenu] = useState<{ x: number; y: number; target: StructureTarget; actions: readonly StructureAction[] }>();
  const [mergeConfirmation, setMergeConfirmation] = useState<{ confirm: () => void }>();
  const [formatDraft, setFormatDraft] = useState<{ basis: WorkbookSnapshot; ranges: readonly EditRange[]; cell: CellReference; tab: FormatTab }>();
  const [lineStyle, setLineStyle] = useState<BorderStyle>('thin');
  const [lineColor, setLineColor] = useState<string>();
  const [borderAction, rememberBorder] = useState<BorderAction>('bottom');
  const [tool, setBorderTool] = useState<{ mode: BorderMode; sessionId: string; sheetId: string }>();
  const [brush, setBrush] = useState<(FormatBrush & { readonly locked: boolean })>();
  const painter = activeFormatBrush(snapshot, brush);
  const [structure, setStructure] = useState<{ action: DialogAction; target: StructureTarget; basis: WorkbookSnapshot }>();
  const [rename, setRename] = useState<{ id: string; basis: WorkbookSnapshot }>();
  const [sortDialog, setSortDialog] = useState<{ basis: WorkbookSnapshot; range: SortExtent; draft: SortDraft }>();
  const [sortOptions, setSortOptions] = useState<{ basis: WorkbookSnapshot; range: SortExtent; draft: SortDraft }>();
  const [sortWarning, setSortWarning] = useState<{
    selection: SortExtent; region: SortExtent; resume: (expand: boolean | null) => void;
  }>();
  const [filterMenu, setFilterMenu] = useState<{
    basis: WorkbookSnapshot; column: number; header: string; textColumn: boolean; values: readonly FilterValue[];
    selected: readonly string[]; hasColumnFilter: boolean; anchor: { x: number; y: number };
  }>();
  const [searchDialog, setSearchDialog] = useState<{ replace: boolean; sheetId: string }>();
  const [searchHistory, setSearchHistory] = useState<SearchHistory>(emptySearchHistory);
  const [searchLook, setSearchLook] = useState<SearchLook>('values');
  const [zoomDialog, setZoomDialog] = useState(false);
  const [insertFunctionDialog, setInsertFunctionDialog] = useState<{ category: InsertFunctionFilter }>();
  const [functionArgsDialog, setFunctionArgsDialog] = useState<{
    readonly name: NativeFunction;
    readonly target: CellReference;
    readonly values: readonly string[];
    readonly activeIndex: number;
  }>();
  const recentFunctions = useMemo(
    () => sanitizeRecentFunctions(preferences.recentFunctions),
    [preferences.recentFunctions],
  );
  const textOrder = useMemo(() => localeTextOrder(locale), [locale]);
  const usedNumberCodes = useMemo(() => [...new Set(formatDraft?.basis.document.sheets.flatMap(item => [
    ...item.cells.flatMap(cell => cell.style?.numberFormat ?? []),
    ...item.rows.flatMap(dim => dim.style?.numberFormat ?? []),
    ...item.columns.flatMap(dim => dim.style?.numberFormat ?? []),
  ]))], [formatDraft]);
  const renaming = useRef(false);
  const sheet = sheets.find(candidate => candidate.id === activeSheetId) ?? sheets[0]!;
  const sheetIndex = sheets.indexOf(sheet);
  const borderTool = tool?.sessionId === snapshot.sessionId && tool.sheetId === sheet.id ? tool : undefined;
  const borderEditor = useMemo(() => formatDraft?.ranges.length === 1 ?
    inspectBorders(formatDraft.basis, formatDraft.ranges[0]!) : undefined, [formatDraft]);
  const selected = selections.get(sheet.id) ?? defaultSelection;
  const sheetFiltered = useMemo(() => filteredRows(sheet), [sheet]);
  const hiddenRowIndexes = useMemo(() => hiddenIndexes(sheet.rows), [sheet.rows]);
  const hiddenColumnIndexes = useMemo(() => hiddenIndexes(sheet.columns), [sheet.columns]);
  const selection = useMemo(() => {
    const row = visibleIndex(selected.anchor.row, hiddenRowIndexes, worksheetViewport.rows, sheetFiltered);
    const column = visibleIndex(selected.anchor.column, hiddenColumnIndexes, worksheetViewport.columns);
    return mergeSelection(sheet, row === selected.anchor.row && column === selected.anchor.column ? selected : singleSelection({ row, column }));
  }, [selected, sheet, hiddenRowIndexes, hiddenColumnIndexes, sheetFiltered]);
  const stats = useMemo(() => selectionStatistics(sheet, selectionAreas(selection), sheetFiltered), [sheet, selection, sheetFiltered]);
  const activeTarget = { sheetId: sheet.id, ...selection.anchor };
  const pythonExecution = usePythonExecution(session, locale,
    snapshot => setSnapshot(snapshot), paused => scheduler.current?.setSuspended(paused));
  const python = usePythonEditor({ session, changed: compound => workbookChanged(compound) });
  const [pythonOpen, setPythonOpen] = useState(false);
  const [pythonTrustPrompt, setPythonTrustPrompt] = useState<{ target: PythonAnchorTarget; all?: boolean }>();
  const [pendingFigureOwner, setPendingFigureOwner] = useState<string>();
  const [figureStatus, setFigureStatus] = useState<{ readonly ownerId: string; readonly message: string }>();
  const [pendingFigureResize, setPendingFigureResize] = useState<PendingFigureResize>();
  const trustedPythonSession = useRef<string | undefined>(undefined);
  const [pythonHiddenFor, setPythonHiddenFor] = useState<string>();
  const [pythonWidth, setPythonWidth] = useState(380);
  const [pythonPick, setPythonPick] = useState<{ readonly sheetId: string; readonly area: Area }>();
  const pythonResize = useRef<{ readonly id: number; readonly startX: number; readonly startWidth: number } | undefined>(undefined);
  const pythonTarget = useMemo(() => pythonAnchorAt(snapshot, sheet.id, selection.anchor),
    [snapshot, sheet.id, selection.anchor.row, selection.anchor.column]);
  const pythonTargetKey = pythonTarget ? pythonTarget.sheetId + ':' + pythonTarget.cellId : undefined;
  const pythonPanelTarget = python.context ? pythonAnchorAt(snapshot, python.context.sheetId, python.context) : undefined;
  const pythonPanelCell = pythonPanelTarget && snapshot.document.sheets.find(item => item.id === pythonPanelTarget.sheetId)?.cells
    .find(cell => cell.id === pythonPanelTarget.cellId);
  const pythonBlock = pythonPanelCell?.content.kind === 'python' ? pythonPanelCell.content.block : undefined;
  const worksheetFigures = useMemo<readonly WorksheetFigure[]>(() => snapshot.document.sheets.flatMap(sourceSheet =>
    sourceSheet.cells.flatMap(cell => {
      if (cell.content.kind !== 'python' || !cell.content.block?.lastFigure) return [];
      const block = cell.content.block;
      return (block.placements ?? []).filter(placement => placement.sheetId === sheet.id).map(placement => ({
        ownerId: block.id, figure: block.lastFigure!, placement, label: `Figure from ${sourceSheet.name}!${address(cell)}`,
      }));
    })), [snapshot.document, sheet.id]);
  const pendingFigure = useMemo(() => {
    if (!pendingFigureOwner) return undefined;
    for (const sourceSheet of snapshot.document.sheets) for (const cell of sourceSheet.cells) {
      if (cell.content.kind === 'python' && cell.content.block?.id === pendingFigureOwner && cell.content.block.lastFigure) {
        return { ownerId: pendingFigureOwner, figure: cell.content.block.lastFigure };
      }
    }
    return undefined;
  }, [snapshot.document, pendingFigureOwner]);
  useEffect(() => {
    if (pendingFigureOwner && !pendingFigure) setPendingFigureOwner(undefined);
  }, [pendingFigureOwner, pendingFigure]);
  const panelFigureStatus = pythonBlock && figureStatus?.ownerId === pythonBlock.id ? figureStatus.message : undefined;
  const pythonInputFresh = useMemo(() => {
    if (!pythonPanelTarget || !pythonPanelCell || pythonPanelCell.content.kind !== 'python' || !pythonBlock?.hasResult ||
        !pythonBlock.inputFingerprint || pythonBlock.status !== 'succeeded') return false;
    try {
      return buildPythonInputs(snapshot, { sheetId: pythonPanelTarget.sheetId, row: pythonPanelTarget.row,
        column: pythonPanelTarget.column, blockId: pythonPanelTarget.blockId, source: pythonPanelCell.content.source },
      calculation.stale).fingerprint === pythonBlock.inputFingerprint;
    } catch { return false; }
  }, [snapshot, pythonPanelTarget?.sheetId, pythonPanelTarget?.cellId, pythonPanelTarget?.blockId,
    pythonBlock, pythonPanelCell, calculation.stale]);
  const pythonExecutionState = pythonExecution.state;
  const calculatingCells = useMemo(() => new Set([
    ...calculation.calculating,
    ...(pythonExecutionState?.running ? pythonExecutionState.calculatingCellIds ?? [] : []),
    ...(pythonExecutionState?.running && pythonExecutionState.currentCellId ? [pythonExecutionState.currentCellId] : []),
  ]), [calculation.calculating, pythonExecutionState?.running, pythonExecutionState?.calculatingCellIds,
    pythonExecutionState?.currentCellId]);
  const pythonResultExecution = pythonPanelTarget && pythonExecutionState?.key === pythonPanelTarget.sheetId + ':' + pythonPanelTarget.cellId
    ? pythonExecutionState : undefined;
  useEffect(() => {
    if (pythonTargetKey && pythonTargetKey === pythonHiddenFor) return;
    if (pythonHiddenFor) setPythonHiddenFor(undefined);
    // A cross-sheet xl pick keeps the original anchor even though another sheet is active.
    if (python.context && python.wantsReference() && python.context.sheetId !== sheet.id) return;
    python.sync(pythonTarget);
    if (pythonTarget) setPythonOpen(true);
  }, [pythonTarget, pythonTargetKey, pythonHiddenFor, python.sync, python.context, python.wantsReference]);
  const focusGrid = () => worksheet.current?.focus();
  const scheduleFocusGrid = () => {
    const origin = document.activeElement;
    requestAnimationFrame(() => {
      const current = document.activeElement;
      if (current && current !== origin && current !== document.body && current !== document.documentElement) return;
      worksheet.current?.focus();
    });
  };
  const editing = useWorksheetEditing(session, (compound, history) => workbookChanged(compound, history), separator, locale, makeId, focusGrid, surface => {
    root.current?.querySelector<HTMLTextAreaElement>('[data-cell-editor="' + surface + '"]')?.focus({ preventScroll: true });
  });

  const tables = useMemo(() => tablesOnSheet(snapshot.document, sheet.id), [snapshot.document, sheet.id]);
  const activeTable = useMemo(() => tables.find(t => extentContains(t.area, selection.anchor)), [tables, selection.anchor]);
  const [definitionDraft, setDefinitionDraft] = useState<DefinitionDraft | undefined>();
  const [nameManagerOpen, setNameManagerOpen] = useState(false);
  const [pickedManagerExtent, setPickedManagerExtent] = useState<string | undefined>();
  const [rangePicker, setRangePicker] = useState<{
    initial: string;
    onDone: (result: string | null) => void;
  }>();
  const [createPivotDraft, setCreatePivotDraft] = useState<{
    readonly sourceText: string;
    readonly placement: PivotPlacement;
    readonly locationText: string;
  }>();
  const [pivotFieldListOpen, setPivotFieldListOpen] = useState(true);
  const [pivotPaneWidth, setPivotPaneWidth] = useState(310);
  const pivotResize = useRef<{ readonly id: number; readonly startX: number; readonly startWidth: number } | undefined>(undefined);
  const [fieldSettingsModal, setFieldSettingsModal] = useState<{
    readonly pivotId: DocumentId;
    readonly request: FieldSettingsRequest;
  }>();
  const [pivotOptionsOpen, setPivotOptionsOpen] = useState(false);
  const [changePivotSourceDraft, setChangePivotSourceDraft] = useState<{
    readonly pivotId: DocumentId;
    readonly sourceText: string;
  }>();
  const [movePivotDraft, setMovePivotDraft] = useState<{
    readonly pivotId: DocumentId;
    readonly placement: PivotPlacement;
    readonly locationText: string;
  }>();
  const [pivotFilterMenu, setPivotFilterMenu] = useState<{
    readonly pivotId: DocumentId;
    readonly kind: 'reportFilter' | 'rowLabel' | 'colLabel';
    readonly fieldNames: readonly string[];
    readonly anchor: { readonly x: number; readonly y: number };
  }>();
  const activePivot = useMemo(
    () => findPivotAt(snapshot.document, { sheetId: sheet.id, row: selection.anchor.row, column: selection.anchor.column }),
    [snapshot.document, sheet.id, selection.anchor.row, selection.anchor.column],
  );
  const activePivotField = useMemo(
    () => activePivot
      ? resolveActivePivotField(snapshot.document, activePivot, { sheetId: sheet.id, row: selection.anchor.row, column: selection.anchor.column })
      : undefined,
    [activePivot, snapshot.document, sheet.id, selection.anchor.row, selection.anchor.column],
  );
  const pivotOverlays = useMemo(
    () => computePivotOverlays(snapshot.document, sheet.id),
    [snapshot.document, sheet.id],
  );
  const pivotOpen = Boolean(activePivot && pivotFieldListOpen);
  const settingsPivot = fieldSettingsModal
    ? snapshot.document.pivots?.find(item => item.id === fieldSettingsModal.pivotId)
    : undefined;
  const optionsPivot = pivotOptionsOpen ? activePivot : undefined;
  const menuPivot = pivotFilterMenu
    ? snapshot.document.pivots?.find(item => item.id === pivotFilterMenu.pivotId)
    : undefined;
  useEffect(() => {
    if (activePivot?.id) setPivotFieldListOpen(true);
    else setPivotOptionsOpen(false);
  }, [activePivot?.id]);

  useEffect(() => {
    let live = true;
    const lease = workspaceLease.current!;
    void (async () => {
      try {
        const [loadedPreferences, recent, recoveries] = await Promise.all([
          localStore.loadPreferences(), localStore.listRecent(), localStore.listRecoveries(),
        ]);
        if (!live) return;
        setPreferences(loadedPreferences);
        setView(loadedPreferences.view);
        setSearchHistory(loadedPreferences.searchHistory);
        setSearchLook(loadedPreferences.searchLook);
        setRecentWorkbooks(recent);
        const blankId = session.active.id;
        if (!await lease.switchTo(blankId)) {
          setWorkspaceConflict(blankId);
          return;
        }
        const recovery = initialWorkbook ? undefined : recoveries[0];
        if (recovery) {
          if (!await acquireRecoveryLease(lease, recovery.workbookId)) {
            setFileMessage({ title: 'Recovery already open', message: 'This workbook is open in another tab. That tab owns the recovery copy; your current workbook was kept.' });
            return;
          }
          const saved = await localStore.getWorkspace(recovery.workbookId);
          if (saved && saved.updatedAt >= recovery.savedAt) {
            await localStore.clearRecovery(recovery.workbookId, lease.owner);
            await lease.switchTo(blankId);
          } else if (live) {
            setPendingRecovery(recovery);
          }
        }
      } catch (error) {
        if (live) {
          const message = error instanceof Error ? error.message : String(error);
          setStorageError(message);
          setFileStatus('Browser-local storage is unavailable. Use Export to download a backup.');
        }
      } finally {
        if (live) setPreferencesReady(true);
      }
    })();
    return () => {
      live = false;
      void lease.release().catch(() => {});
    };
  }, [initialWorkbook, localStore, session]);

  useEffect(() => {
    if (!preferencesReady) return;
    const timer = window.setTimeout(() => {
      void localStore.savePreferences(preferences).then(() => setStorageError(current =>
        current.startsWith('Local settings could not be saved.') ? '' : current)).catch(error => {
        const message = error instanceof Error ? error.message : String(error);
        setStorageError('Local settings could not be saved. Changes apply for this session. ' + message);
      });
    }, 250);
    return () => window.clearTimeout(timer);
  }, [localStore, preferences, preferencesReady]);

  useEffect(() => {
    if (preferencesReady) pythonExecution.setAutomatic(preferences.pythonCalculation === 'automatic');
  }, [preferencesReady, preferences.pythonCalculation, pythonExecution.setAutomatic]);

  useEffect(() => {
    if (!preferencesReady || preferences.view === view) return;
    setPreferences(current => current.view === view ? current : { ...current, view });
  }, [preferences.view, preferencesReady, view]);

  useEffect(() => {
    if (!preferencesReady || preferences.searchHistory.find === searchHistory.find &&
        preferences.searchHistory.replace === searchHistory.replace) return;
    setPreferences(current => current.searchHistory.find === searchHistory.find &&
      current.searchHistory.replace === searchHistory.replace ? current : { ...current, searchHistory });
  }, [preferences.searchHistory.find, preferences.searchHistory.replace, preferencesReady, searchHistory]);

  useEffect(() => {
    if (!preferencesReady || preferences.searchLook === searchLook) return;
    setPreferences(current => current.searchLook === searchLook ? current : { ...current, searchLook });
  }, [preferences.searchLook, preferencesReady, searchLook]);

  useEffect(() => {
    const beforeUnload = (event: globalThis.BeforeUnloadEvent) => {
      if (!session.isDirty) return;
      event.preventDefault();
      event.returnValue = '';
    };
    const suspend = () => { void workspaceLease.current?.suspend().catch(() => {}); };
    const resume = () => { void workspaceLease.current?.resume().catch(() => {}); };
    window.addEventListener('beforeunload', beforeUnload);
    window.addEventListener('pagehide', suspend);
    window.addEventListener('pageshow', resume);
    return () => {
      window.removeEventListener('beforeunload', beforeUnload);
      window.removeEventListener('pagehide', suspend);
      window.removeEventListener('pageshow', resume);
    };
  }, [session, snapshot]);

  useEffect(() => {
    if (!preferencesReady || !session.isDirty || workspaceConflict || pendingRecovery ||
        workspaceLease.current?.activeId !== snapshot.document.id) return;
    const captured = snapshot;
    const epoch = ++recoveryEpoch.current;
    let live = true;
    const timer = window.setTimeout(() => {
      void (async () => {
        try {
          const packageBlob = await saveNativePackage(captured.document);
          if (!live || epoch !== recoveryEpoch.current || !session.isCurrent(captured) || !session.isDirty) return;
          await localStore.writeRecovery({
            id: 'recovery:' + captured.document.id,
            workbookId: captured.document.id,
            name: activeFile.name,
            package: packageBlob,
            savedAt: Date.now(),
            revision: captured.authoredRevision,
          }, workspaceLease.current!.owner);
          if (live && epoch === recoveryEpoch.current) setRecoveryStatus('Local recovery copy updated.');
        } catch (error) {
          if (!live || epoch !== recoveryEpoch.current) return;
          const message = error instanceof Error ? error.message : String(error);
          setStorageError('Local recovery is unavailable. Export a backup to keep a portable copy. ' + message);
        }
      })();
    }, 900);
    return () => { live = false; window.clearTimeout(timer); };
  }, [activeFile.name, localStore, pendingRecovery, preferencesReady, session, snapshot, workspaceConflict]);

  function askReplacement(): Promise<ReplacementChoice> {
    return new Promise(resolve => {
      replacementResolver.current = resolve;
      setReplacementPrompt(true);
    });
  }

  function chooseReplacement(choice: ReplacementChoice) {
    const resolve = replacementResolver.current;
    replacementResolver.current = undefined;
    setReplacementPrompt(false);
    resolve?.(choice);
  }

  function askSaveAsName(purpose: 'browser-save' | 'local-save-as' = 'local-save-as'): Promise<string | null> {
    return new Promise(resolve => {
      saveAsResolver.current = resolve;
      setSaveNamePurpose(purpose);
      setSaveAsOpen(true);
    });
  }

  function finishSaveAsName(name: string | null) {
    const resolve = saveAsResolver.current;
    saveAsResolver.current = undefined;
    setSaveAsOpen(false);
    resolve?.(name);
  }

  function askOptions(): Promise<{ readonly nativeCalculation: CalculationMode; readonly pythonCalculation: CalculationMode } | null> {
    return new Promise(resolve => {
      optionsResolver.current = resolve;
      setOptionsOpen(true);
    });
  }

  function finishOptions(options: { readonly nativeCalculation: CalculationMode; readonly pythonCalculation: CalculationMode } | null) {
    const resolve = optionsResolver.current;
    optionsResolver.current = undefined;
    setOptionsOpen(false);
    resolve?.(options);
  }

  function workbookChanged(compound: boolean, history = false) {
    if (history) pythonExecution.historyRestored();
    else if (pythonExecution.changed(compound)) scheduler.current?.setSuspended(true);
    scheduler.current?.changed(compound);
  }

  function askXlsxReport(title: string, issues: readonly ExchangeIssue[]): Promise<ReportChoice> {
    return new Promise(resolve => {
      xlsxReportResolver.current = resolve;
      setXlsxReport({ title, issues });
    });
  }

  function finishXlsxReport(choice: ReportChoice) {
    const resolve = xlsxReportResolver.current;
    xlsxReportResolver.current = undefined;
    setXlsxReport(undefined);
    resolve?.(choice);
  }

  async function refreshRecent() {
    try { setRecentWorkbooks(await localStore.listRecent()); }
    catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setStorageError('Recent workbooks could not be loaded. ' + message);
    }
  }

  async function clearRecoveryFor(workbookId: string) {
    const recoveries = await localStore.listRecoveries();
    if (recoveries.some(item => item.workbookId === workbookId)) {
      await localStore.clearRecovery(workbookId, workspaceLease.current!.owner);
    }
    setRecoveryStatus('');
  }

  function resetWorkbookUi(document: WorkbookDocument) {
    setSnapshot(session.capture());
    setRecoveryStatus('');
    setActiveSheetId(document.sheets[0]!.id);
    setSelections(new Map());
    setMenu(undefined);
    setBorderTool(undefined);
    setBrush(undefined);
    editing.selectionChanged();
    editing.clearCut();
    editing.clearStatus();
    scheduler.current?.changed(true);
    requestedFocus.current = true;
    requestAnimationFrame(() => {
      const focused = globalThis.document.activeElement;
      if (!(focused instanceof HTMLElement && focused.dataset.cellEditor)) {
        worksheet.current?.focus();
      }
    });
  }

  async function verifyActiveWorkspace(workbookId: string) {
    const lease = workspaceLease.current!;
    if (lease.activeId !== workbookId || lease.isLost || !await localStore.renewLease(workbookId, lease.owner)) {
      setWorkspaceConflict(workbookId);
      throw new WorkspaceConflictError(workbookId);
    }
  }

  async function saveCurrent(): Promise<boolean> {
    recoveryEpoch.current++;
    const captured = session.capture();
    await verifyActiveWorkspace(captured.document.id);
    const packageBlob = await saveNativePackage(captured.document);
    const now = Date.now();
    const handle = activeFile.source === 'native' ? activeFile.fileHandle : undefined;
    let fileWasWritten = false;
    let fileWriteWarning = '';
    let nextContext = activeFile;
    if (handle) {
      try {
        await verifyActiveWorkspace(captured.document.id);
        const version = await writeNativeFile(handle, packageBlob,
          activeFile.fileSize === undefined || activeFile.fileLastModified === undefined ? undefined : {
            size: activeFile.fileSize, lastModified: activeFile.fileLastModified,
          });
        fileWasWritten = true;
        nextContext = { ...activeFile, fileSize: version.size, fileLastModified: version.lastModified };
      } catch (error) {
        fileWriteWarning = error instanceof Error ? error.message : String(error);
      }
    }

    let localWasWritten = false;
    try {
      const record: LocalWorkspaceRecord = {
        id: captured.document.id,
        name: activeFile.name,
        package: packageBlob,
        createdAt: now,
        updatedAt: now,
        lastUsedAt: now,
        source: activeFile.source === 'xlsx' ? 'xlsx' : 'native',
        ...(activeFile.sourceFileName && { sourceFileName: activeFile.sourceFileName }),
        ...(handle && { fileHandle: handle }),
        ...(nextContext.fileSize !== undefined && { fileSize: nextContext.fileSize }),
        ...(nextContext.fileLastModified !== undefined && { fileLastModified: nextContext.fileLastModified }),
      };
      await localStore.saveWorkspace(record, workspaceLease.current!.owner);
      localWasWritten = true;
      await refreshRecent();
    } catch (error) {
      if (error instanceof WorkspaceConflictError) throw error;
      const message = error instanceof Error ? error.message : String(error);
      setStorageError('The local workspace could not be updated. ' + message);
    }

    if (!fileWasWritten && !localWasWritten) {
      throw new LocalWorkspaceError(storageError || 'The workbook could not be saved to this browser. Use Export to download a backup.');
    }
    if (!session.markSaved(captured)) {
      throw new LocalWorkspaceError('The save completed for an earlier workbook snapshot. The active workbook was not marked saved.');
    }
    setActiveFile(nextContext);
    if (!session.isDirty) {
      try { await localStore.clearRecovery(captured.document.id, workspaceLease.current!.owner); }
      catch (error) {
        if (!(error instanceof WorkspaceConflictError)) {
          setStorageError('The recovery copy could not be cleared yet. ' + (error instanceof Error ? error.message : String(error)));
        } else throw error;
      }
    if (!session.isDirty) setRecoveryStatus('');
    }
    const dirty = session.isDirty;
    setFileStatus(fileWriteWarning ?
      (localWasWritten ? 'Saved to this browser. The linked file was not changed: ' : fileWasWritten ? 'Saved to the linked file. Local storage was unavailable: ' : '') + fileWriteWarning :
      dirty ? 'Saved snapshot; newer edits remain modified.' : fileWasWritten ? 'Saved ' + handle!.name : 'Saved to this browser.');
    setBackstageOpen(false);
    return true;
  }

  async function saveAsToHandle(handle: WorkspaceFileHandle): Promise<boolean> {
    recoveryEpoch.current++;
    const captured = session.capture();
    await verifyActiveWorkspace(captured.document.id);
    const packageBlob = await saveNativePackage(captured.document);
    const version = await writeNativeFile(handle, packageBlob);
    const name = handle.name.replace(/\.exlt$/iu, '') || handle.name;
    const context: ActiveFileContext = {
      name, source: 'native', fileHandle: handle, fileSize: version.size, fileLastModified: version.lastModified,
      ...(activeFile.importedXlsx && { importedXlsx: true }),
      ...(activeFile.sourceFileHandle && { sourceFileHandle: activeFile.sourceFileHandle }),
      sourceFileName: handle.name,
    };
    let localWarning = '';
    try {
      await localStore.saveWorkspace({
        id: captured.document.id, name, package: packageBlob, createdAt: Date.now(), updatedAt: Date.now(),
        lastUsedAt: Date.now(), source: 'native', sourceFileName: handle.name, fileHandle: handle,
        fileSize: version.size, fileLastModified: version.lastModified,
      }, workspaceLease.current!.owner);
      await refreshRecent();
    } catch (error) {
      if (error instanceof WorkspaceConflictError) throw error;
      localWarning = error instanceof Error ? error.message : String(error);
      setStorageError('The file was saved, but its local Recent entry could not be updated. ' + localWarning);
    }
    if (!session.markSaved(captured)) throw new LocalWorkspaceError('The file was saved, but the active workbook changed.');
    setActiveFile(context);
    if (!session.isDirty) {
      try { await localStore.clearRecovery(captured.document.id, workspaceLease.current!.owner); setRecoveryStatus(''); }
      catch { /* The file save is already confirmed. */ }
    }
    setFileStatus(localWarning ? 'Saved ' + handle.name + '. Local storage could not be updated.' : 'Saved ' + handle.name + '.');
    setBackstageOpen(false);
    return true;
  }

  async function saveAsLocally(name: string): Promise<boolean> {
    recoveryEpoch.current++;
    const captured = session.capture();
    const packageBlob = await saveNativePackage(captured.document);
    const now = Date.now();
    const context: ActiveFileContext = { name, source: activeFile.source === 'xlsx' ? 'xlsx' : 'native',
      ...(activeFile.importedXlsx && { importedXlsx: true }),
      ...(activeFile.sourceFileHandle && { sourceFileHandle: activeFile.sourceFileHandle }),
      ...(activeFile.sourceFileName && { sourceFileName: activeFile.sourceFileName }) };
    await localStore.saveWorkspace({
      id: captured.document.id, name, package: packageBlob, createdAt: now, updatedAt: now,
      lastUsedAt: now, source: context.source === 'xlsx' ? 'xlsx' : 'native',
      ...(context.sourceFileName && { sourceFileName: context.sourceFileName }),
    }, workspaceLease.current!.owner);
    if (!session.markSaved(captured)) throw new LocalWorkspaceError('The save completed for an earlier workbook snapshot.');
    setActiveFile(context);
    await refreshRecent();
    if (!session.isDirty) {
      try { await localStore.clearRecovery(captured.document.id, workspaceLease.current!.owner); setRecoveryStatus(''); }
      catch { /* The local snapshot remains saved. */ }
    }
    setFileStatus('Saved to this browser as ' + name + '.');
    setBackstageOpen(false);
    return true;
  }

  async function saveCurrentWithName(): Promise<boolean> {
    if (activeFile.source !== 'untitled') return saveCurrent();
    const name = await askSaveAsName('browser-save');
    return name === null ? false : saveAsLocally(name);
  }

  async function confirmReplacement(): Promise<boolean> {
    if (!session.isDirty) return true;
    const choice = await askReplacement();
    if (choice === 'cancel') {
      setFileStatus('Operation cancelled.');
      return false;
    }
    if (choice === 'save') return saveCurrentWithName();
    recoveryEpoch.current++;
    try { await clearRecoveryFor(session.active.id); }
    catch (error) {
      setFileMessage({ title: 'Recovery copy could not be removed', message: error instanceof Error ? error.message : String(error) });
      return false;
    }
    return true;
  }

  async function adoptOpened(document: WorkbookDocument, record: LocalWorkspaceRecord, saved = true,
    sourceFileHandle?: WorkspaceFileHandle) {
    const basis = session.capture();
    if (!await confirmReplacement()) return;
    const lease = workspaceLease.current!;
    if (!await lease.switchTo(document.id)) throw new WorkspaceConflictError(document.id);
    if (!session.tryAdopt(basis, document, saved)) {
      await lease.switchTo(basis.document.id);
      throw new LocalWorkspaceError('Workbook changed while opening. Your current workbook was kept; try Open again.');
    }
    pythonExecution.adopted();
    const active: ActiveFileContext = {
      name: record.name, source: saved ? (record.source === 'xlsx' ? 'xlsx' : 'native') : 'recovery',
      ...(record.source === 'xlsx' && saved ? { importedXlsx: true } : {}),
      ...(record.fileHandle && { fileHandle: record.fileHandle }),
      ...(sourceFileHandle && { sourceFileHandle }),
      ...(record.fileSize !== undefined && { fileSize: record.fileSize }),
      ...(record.fileLastModified !== undefined && { fileLastModified: record.fileLastModified }),
      ...(record.sourceFileName && { sourceFileName: record.sourceFileName }),
    };
    setActiveFile(active);
    resetWorkbookUi(document);
    setBackstageOpen(false);
    try {
      if (saved) await localStore.saveWorkspace(record, lease.owner);
      await refreshRecent();
      setStorageError('');
    } catch (error) {
      if (error instanceof WorkspaceConflictError) throw error;
      setStorageError('The workbook opened, but its local Recent entry could not be updated. ' + (error instanceof Error ? error.message : String(error)));
    }
    setFileStatus((saved ? 'Opened ' : 'Recovered ') + record.name + '.');
  }

  async function openExternalWorkbook() {
    const chosen = await pickNativeOpenFile();
    if (!chosen) { setFileStatus('Open cancelled.'); return; }
    const fileName = chosen.file.name.toLowerCase();
    if (fileName.endsWith('.xlsx')) return openXlsxWorkbook(chosen);
    if (!fileName.endsWith('.exlt')) {
      throw new LocalWorkspaceError('Open an .exlt or .xlsx workbook.');
    }
    setFileStatus('Opening ' + chosen.file.name + '…');
    const document = await loadNativePackage(chosen.file);
    const version = fileVersion(chosen.file);
    const record: LocalWorkspaceRecord = {
      id: document.id,
      name: chosen.file.name.replace(/\.exlt$/iu, '') || chosen.file.name,
      package: chosen.file,
      createdAt: Date.now(), updatedAt: Date.now(), lastUsedAt: Date.now(), source: 'native',
      sourceFileName: chosen.file.name,
      ...(chosen.handle && { fileHandle: chosen.handle }),
      fileSize: version.size, fileLastModified: version.lastModified,
    };
    await adoptOpened(document, record);
  }

  async function openXlsxWorkbook(chosen: OpenedNativeFile) {
    setFileStatus('Opening ' + chosen.file.name + '…');
    const { document, report } = await importXlsx(chosen.file);
    if (report.issues.length > 0) {
      const choice = await askXlsxReport('XLSX import report', report.issues);
      if (choice === 'cancel') {
        setFileStatus('XLSX import cancelled. Your current workbook was kept.');
        return;
      }
    }
    const version = fileVersion(chosen.file);
    const name = chosen.file.name.replace(/\.xlsx$/iu, '') || chosen.file.name;
    const packageBlob = await saveNativePackage(document);
    const record: LocalWorkspaceRecord = {
      id: document.id, name, package: packageBlob, createdAt: Date.now(), updatedAt: Date.now(),
      lastUsedAt: Date.now(), source: 'xlsx', sourceFileName: chosen.file.name,
      fileSize: version.size, fileLastModified: version.lastModified,
    };
    await adoptOpened(document, record, true, chosen.handle);
    setFileStatus('Opened ' + chosen.file.name + '. This workbook is stored in this browser; Export writes a separate .xlsx copy.');
  }

  async function openRecentWorkbook(record: LocalWorkspaceRecord) {
    const latest = await localStore.getWorkspace(record.id);
    if (!latest) throw new LocalWorkspaceError('This Recent workbook is no longer available in this browser.');
    const document = await loadNativePackage(latest.package);
    if (document.id !== latest.id) throw new LocalWorkspaceError('The stored Recent workbook identity is invalid.');
    await adoptOpened(document, { ...latest, lastUsedAt: Date.now() });
  }

  async function createBlankWorkbook(keepBackstage = false) {
    const oldId = session.active.id;
    if (!await confirmReplacement()) return;
    const document = newWorkbook(makeId);
    const lease = workspaceLease.current!;
    if (!await lease.switchTo(document.id)) throw new WorkspaceConflictError(document.id);
    const basis = session.capture();
    if (!session.tryAdopt(basis, document, true)) {
      await lease.switchTo(oldId);
      throw new LocalWorkspaceError('Workbook changed while creating a new workbook. Try New again.');
    }
    pythonExecution.adopted();
    setActiveFile({ name: 'Book1', source: 'untitled' });
    resetWorkbookUi(document);
    if (keepBackstage) setBackstageOpen(true); else setBackstageOpen(false);
    setFileStatus(keepBackstage ? 'Workbook closed.' : 'New blank workbook.');
  }

  async function createSupportSlaWorkbook() {
    const response = await fetch(supportSlaTemplateUrl);
    if (!response.ok) throw new LocalWorkspaceError('The Support SLA template could not be loaded.');
    const template = await loadNativePackage(await response.blob());
    const document = { ...template, id: makeId() };
    const previousId = session.active.id;
    if (!await confirmReplacement()) return;
    const lease = workspaceLease.current!;
    if (!await lease.switchTo(document.id)) throw new WorkspaceConflictError(document.id);
    const basis = session.capture();
    if (!session.tryAdopt(basis, document, false)) {
      await lease.switchTo(previousId);
      throw new LocalWorkspaceError('Workbook changed while creating this template copy. Your current workbook was kept; try again.');
    }
    pythonExecution.adopted();
    setActiveFile({ name: 'Book1', source: 'untitled' });
    resetWorkbookUi(document);
    setBackstageOpen(false);
    setFileStatus('New workbook from Support SLA Dashboard.');
  }

  async function exportWorkbook() {
    if (activeFile.importedXlsx || activeFile.source === 'xlsx') return exportXlsxWorkbook();
    const captured = session.capture();
    const packageBlob = await saveNativePackage(captured.document);
    const suggestedName = (activeFile.name || 'Book1') + '.exlt';
    if (hasNativeSavePicker()) {
      const handle = await pickNativeSaveFile(suggestedName);
      if (!handle) { setFileStatus('Export cancelled.'); return; }
      await writeNativeFile(handle, packageBlob);
      setFileStatus('Exported ' + handle.name + '. The local saved state is unchanged.');
    } else {
      downloadNativePackage(packageBlob, suggestedName);
      setFileStatus('Export download started. This does not change the local saved state.');
    }
    setBackstageOpen(false);
  }

  async function exportXlsxWorkbook() {
    const captured = session.capture();
    const report = inspectXlsxExport(captured.document);
    const choice = await askXlsxReport('XLSX export report', report.issues);
    if (choice === 'cancel') {
      setFileStatus('Export cancelled. Your workbook is unchanged.');
      return;
    }
    const { blob } = await writeXlsx(captured.document);
    const suggestedName = (activeFile.name || 'Book1') + '.xlsx';
    if (hasXlsxSavePicker()) {
      const handle = await pickXlsxSaveFile(suggestedName);
      if (!handle) { setFileStatus('Export cancelled.'); return; }
      if (activeFile.sourceFileHandle && await sameFileEntry(handle, activeFile.sourceFileHandle)) {
        throw new LocalWorkspaceError('Choose a separate XLSX export destination; the imported source cannot be overwritten.');
      }
      await writeNativeFile(handle, blob);
      setFileStatus('Exported ' + handle.name + '. The local saved state is unchanged.');
    } else {
      downloadXlsxFile(blob, suggestedName);
      setFileStatus('XLSX export download started. This does not change the local saved state.');
    }
    setBackstageOpen(false);
  }

  async function runFileOperation(title: string, action: () => Promise<void>) {
    if (fileBusyRef.current) return;
    if (!preferencesReady) { setFileStatus('Local workspace is still starting. Try again in a moment.'); return; }
    if (workspaceConflict) return;
    if (root.current?.querySelector('dialog[open]')) return;
    if (!python.barrier() || !editing.commit()) return;
    fileBusyRef.current = true;
    setFileBusy(true);
    try { await action(); }
    catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      setFileMessage({ title: error instanceof WorkspaceConflictError ? 'Workbook already open' : title + ' failed', message });
      setFileStatus(title + ' failed. Your current workbook was kept.');
    } finally {
      fileBusyRef.current = false;
      setFileBusy(false);
    }
  }

  async function decideWorkspaceRemoval(record: LocalWorkspaceRecord, remove: boolean) {
    setWorkspaceToRemove(undefined);
    if (!remove) return;
    await new Promise<void>(resolve => requestAnimationFrame(() => resolve()));
    if (record.id === session.active.id) {
      setFileMessage({ title: 'Close this workbook first', message: 'Close this workbook before removing its browser-local saved copy.' });
      return;
    }
    await runFileOperation('Remove local copy', async () => {
      await localStore.removeWorkspace(record.id);
      await refreshRecent();
      setFileStatus('Removed ' + record.name + ' from this browser. Files on disk were not changed.');
    });
  }

  async function runFileAction(action: FileAction) {
    if (action === 'options') {
      if (root.current?.querySelector('dialog[open]')) return;
      const options = await askOptions();
      if (options) {
        scheduler.current?.setMode(options.nativeCalculation);
        pythonExecution.setAutomatic(options.pythonCalculation === 'automatic');
        setPreferences(current => ({ ...current, ...options }));
        setFileStatus('Calculation preferences saved.');
      }
      return;
    }
    await runFileOperation(action === 'saveAs' ? 'Save As' : action[0]!.toUpperCase() + action.slice(1), async () => {
      switch (action) {
        case 'new': await createBlankWorkbook(); break;
        case 'example': await createSupportSlaWorkbook(); break;
        case 'open': await openExternalWorkbook(); break;
        case 'save':
          await saveCurrentWithName();
          break;
        case 'saveAs':
          if (hasNativeSavePicker()) {
            const handle = await pickNativeSaveFile(activeFile.name + '.exlt');
            if (!handle) { setFileStatus(''); return; }
            await saveAsToHandle(handle);
          } else {
            const name = await askSaveAsName('local-save-as');
            if (name !== null) await saveAsLocally(name);
          }
          break;
        case 'export': await exportWorkbook(); break;
        case 'close': await createBlankWorkbook(true); break;
      }
    });
  }

  async function chooseRecovery(choice: 'recover' | 'discard') {
    const recovery = pendingRecovery;
    if (!recovery) return;
    setPendingRecovery(undefined);
    if (choice === 'discard') {
      try {
        await localStore.clearRecovery(recovery.workbookId, workspaceLease.current!.owner);
        setRecoveryStatus('');
        const blank = session.active.id;
        if (!await workspaceLease.current!.switchTo(blank)) throw new WorkspaceConflictError(blank);
        setFileStatus('Recovery copy discarded.');
      } catch (error) {
        setFileMessage({ title: 'Recovery copy could not be removed', message: error instanceof Error ? error.message : String(error) });
      }
      return;
    }
    try {
      const document = await loadNativePackage(recovery.package);
      if (document.id !== recovery.workbookId) throw new LocalWorkspaceError('Recovery workbook identity is invalid.');
      const basis = session.capture();
      if (!session.tryAdopt(basis, document, false)) throw new LocalWorkspaceError('The current workbook changed during recovery.');
      pythonExecution.adopted();
      setActiveFile({ name: recovery.name, source: 'recovery' });
      resetWorkbookUi(document);
      setFileStatus('Recovered ' + recovery.name + '. Save the workbook to keep these changes.');
    } catch (error) {
      setFileMessage({ title: 'Recovery failed', message: error instanceof Error ? error.message : String(error) });
    }
  }

  function setSelection(next: GridSelection): boolean {
    if (!rangePicker && !python.barrier()) return false;
    if (rangePicker) {
      setSelections(current => new Map(current).set(sheet.id, next));
      const selArea = selectionAreas(next)[0];
      if (selArea) {
        const areaExtent = {
          sheetId: sheet.id,
          row: selArea.top,
          column: selArea.left,
          rows: selArea.bottom - selArea.top + 1,
          columns: selArea.right - selArea.left + 1,
        };
        const refText = refersTo(snapshot.document, areaExtent);
        setRangePicker(prev => prev ? { ...prev, initial: refText } : undefined);
      }
      return true;
    }
    if (!editing.commit()) return false;
    editing.selectionChanged();
    setSelections(current => new Map(current).set(sheet.id, next));
    return true;
  }

  function editorSelection() {
    const input = document.activeElement;
    if (input instanceof HTMLTextAreaElement && input.dataset.cellEditor && root.current?.contains(input)) {
      return { start: input.selectionStart, end: input.selectionEnd };
    }
  }

  function keepEditorOnSheetClick(event: MouseEvent<HTMLButtonElement>) {
    const caret = editorSelection();
    if (event.button === 0 && caret && editing.pickSlot(caret)) event.preventDefault();
  }

  function chooseSheet(nextSheetId: DocumentId) {
    if (python.wantsReference()) {
      setActiveSheetId(nextSheetId);
      return;
    }
    if (!python.barrier()) return;
    const caret = editorSelection();
    if (caret && editing.pickSlot(caret)) {
      if (nextSheetId !== sheet.id) {
        editing.showCrossSheetDraft(caret);
        setActiveSheetId(nextSheetId);
      }
      return;
    }
    if (!editing.commit()) return;
    if (nextSheetId === sheet.id) {
      worksheet.current?.focus();
      return;
    }
    editing.selectionChanged();
    if (!painter?.locked) setBrush(undefined);
    setBorderTool(undefined);
    requestedFocus.current = true;
    setActiveSheetId(nextSheetId);
  }

  function finishEditor(target: typeof activeTarget, key?: AfterEditKey, reverse?: boolean) {
    function finish() {
      if (key) worksheet.current?.moveAfterEdit(target, key, !!reverse);
      else worksheet.current?.goTo(target);
      focusGrid();
    }
    if (target.sheetId === sheet.id) finish();
    else {
      setActiveSheetId(target.sheetId);
      requestAnimationFrame(finish);
    }
  }

  function chooseAdjacent(direction: number) {
    const next = sheets[sheetIndex + direction];
    if (next) chooseSheet(next.id);
  }

  function onView(patch: Partial<SheetView>) {
    if (!editing.commit()) return;
    setView(current => ({ ...current, ...patch }));
  }

  function setCalculationMode(mode: CalculationMode) {
    if (!editing.commit()) return;
    scheduler.current?.setMode(mode);
    setPreferences(current => ({ ...current, nativeCalculation: mode }));
  }

  function goTo(value: string) {
    const point = parseAddress(value);
    return !!point && !!worksheet.current?.goTo(point);
  }

  function ranges(): EditRange[] {
    return selectionAreas(selection).map(area => ({
      sheetId: sheet.id, row: area.top, column: area.left, rows: area.bottom - area.top + 1, columns: area.right - area.left + 1,
    }));
  }

  function applyStructure(basis: WorkbookSnapshot, command: StructureCommand, active?: string) {
    requireDocument(session.isCurrent(basis), 'The workbook changed. Open the command again.');
    const revision = session.revision;
    requireDocument(session.tryExecute(basis, command), 'The workbook changed. Open the command again.');
    editing.clearStatus();
    if (session.revision !== revision) {
      editing.selectionChanged();
      editing.clearCut();
      workbookChanged(true);
    }
    setBorderTool(undefined);
    if (active && active !== sheet.id) {
      if (!painter?.locked) setBrush(undefined);
      requestedFocus.current = true;
      setActiveSheetId(active);
    } else scheduleFocusGrid();
  }

  function openCreateDialog(table: boolean) {
    if (!editing.commit()) return;
    const basis = session.capture();
    const selArea = selectionAreas(selection)[0];
    const initialArea = selArea ? {
      sheetId: sheet.id,
      row: selArea.top,
      column: selArea.left,
      rows: selArea.bottom - selArea.top + 1,
      columns: selArea.right - selArea.left + 1,
    } : { sheetId: sheet.id, row: selection.anchor.row, column: selection.anchor.column, rows: 1, columns: 1 };

    const name = table ? nextTableName(basis.document) : '';
    const extentText = refersTo(basis.document, initialArea);
    const hasHeaders = table ? selectionHasHeaders(basis.document, initialArea) : false;
    setDefinitionDraft({
      name,
      comment: '',
      extentText,
      isTable: table,
      isEdit: false,
      hasHeaders,
    });
  }

  function openEditDialog(entry: WorkbookName, lockExtent = false) {
    if (!editing.commit()) return;
    const basis = session.capture();
    setDefinitionDraft({
      id: entry.id,
      name: entry.name,
      comment: entry.comment ?? '',
      extentText: refersTo(basis.document, entry.area),
      isTable: !!entry.isTable,
      isEdit: true,
      lockExtent: lockExtent && !!entry.isTable,
    });
  }

  function handleDefinitionSubmit(name: string, comment: string, extentText: string, headers: boolean): string | undefined {
    try {
      const basis = session.capture();
      const isTable = definitionDraft?.isTable ?? false;
      const isEdit = definitionDraft?.isEdit ?? false;
      const id = definitionDraft?.id ?? makeId();
      const prior = basis.document.names.find(entry => entry.id === id);
      const keepInvalid = isEdit && prior?.area.invalid && extentText.trim().replace(/^=/u, '') === '#REF!';
      const area = definitionDraft?.lockExtent && prior ? prior.area :
        keepInvalid ? prior.area : parseNameExtent(basis.document, sheet.id, extentText);
      const entry: WorkbookName = {
        id,
        name,
        comment,
        area,
        isTable,
        ...(prior?.design && { design: prior.design }),
        ...(prior?.headerLabels && { headerLabels: prior.headerLabels }),
      };
      let command: StructureCommand;
      if (isTable) {
        if (!isEdit) {
          command = createTable(basis, entry, headers, makeId);
        } else {
          command = resizeTable(basis, entry, makeId);
        }
      } else {
        command = defineName(basis, entry);
      }
      applyStructure(basis, command);
      setDefinitionDraft(undefined);
      return undefined;
    } catch (err: unknown) {
      return err instanceof Error ? err.message : String(err);
    }
  }

  function handleApplyManagerExtent(id: string, extentText: string): string | undefined {
    try {
      const basis = session.capture();
      const entry = basis.document.names.find(n => n.id === id);
      requireDocument(entry, 'Select a valid name.');
      requireDocument(!entry.isTable, 'Resize the table from the worksheet or Table Design.');
      const area = parseNameExtent(basis.document, entry.area.sheetId, extentText);
      const updated = { ...entry, area };
      const command = entry.isTable ? resizeTable(basis, updated, makeId) : defineName(basis, updated);
      applyStructure(basis, command);
      return undefined;
    } catch (err: unknown) {
      return err instanceof Error ? err.message : String(err);
    }
  }

  function handleDeleteName(id: string) {
    try {
      const basis = session.capture();
      applyStructure(basis, deleteName(basis, id));
    } catch (err: unknown) {
      editing.reject(err instanceof Error ? err : new Error(String(err)));
    }
  }

  function showNamedExtent(extentText: string) {
    let area;
    try {
      area = parseNameExtent(session.capture().document, sheet.id, extentText);
    } catch {
      scheduleFocusGrid();
      return;
    }
    const next: GridSelection = {
      anchor: { row: area.row, column: area.column },
      end: { row: lastRow(area), column: lastColumn(area) },
      kind: 'cells',
      additional: [],
    };
    setSelections(current => new Map(current).set(area.sheetId, next));
    if (area.sheetId !== sheet.id) {
      requestedFocus.current = true;
      setActiveSheetId(area.sheetId);
    }
    requestAnimationFrame(() => {
      worksheet.current?.reveal(next.anchor);
      worksheet.current?.focus();
    });
  }

  function handlePickRangeFromDefinition(name: string, comment: string, extentText: string) {
    setDefinitionDraft(current => current && { ...current, name, comment, extentText });
    showNamedExtent(extentText);
    setRangePicker({
      initial: extentText,
      onDone: picked => {
        setRangePicker(undefined);
        if (picked) setDefinitionDraft(current => current && { ...current, extentText: picked });
      },
    });
  }

  function handlePickRangeFromManager(extentText: string) {
    showNamedExtent(extentText);
    setRangePicker({
      initial: extentText,
      onDone: picked => {
        setRangePicker(undefined);
        if (picked) setPickedManagerExtent(picked);
      },
    });
  }

  function handlePickRangeFromFunctionArgs(index: number, currentValue: string) {
    if (!functionArgsDialog) return;
    const originSheet = snapshot.document.sheets.find(item => item.id === functionArgsDialog.target.sheetId)?.name ?? sheet.name;
    const initial = currentValue.trim()
      ? (currentValue.trim().startsWith('=') ? currentValue.trim() : `=${currentValue.trim()}`)
      : '';
    setRangePicker({
      initial,
      onDone: picked => {
        setRangePicker(undefined);
        if (!picked) return;
        const formatted = formatArgumentReference(picked, originSheet);
        setFunctionArgsDialog(current => current ? {
          ...current,
          values: current.values.map((val, i) => (i === index ? formatted : val)),
          activeIndex: index,
        } : undefined);
      },
    });
  }

  function handleConfirmFunctionArgs(localizedFormula: string) {
    if (!functionArgsDialog) return;
    const { target } = functionArgsDialog;
    if (!editing.commitFormula(target, localizedFormula)) return;
    setFunctionArgsDialog(undefined);
    scheduleFocusGrid();
  }

  function resolveChartSourceFormula(doc: WorkbookDocument): string {
    const selArea = selectionAreas(selection)[0];
    if (activeTable) {
      return refersTo(doc, activeTable.area);
    }
    if (selArea && (selArea.bottom > selArea.top || selArea.right > selArea.left)) {
      return refersTo(doc, {
        sheetId: sheet.id,
        row: selArea.top,
        column: selArea.left,
        rows: selArea.bottom - selArea.top + 1,
        columns: selArea.right - selArea.left + 1,
      });
    }
    const region = currentRegion(doc, {
      sheetId: sheet.id,
      row: selection.anchor.row,
      column: selection.anchor.column,
      rows: 1,
      columns: 1,
    });
    return refersTo(doc, region);
  }

  function insertChartWithType(chartType: ChartType, explicitSourceFormula?: string, explicitSeriesBy?: 'columns' | 'rows') {
    if (!editing.commit()) return;
    const basis = session.capture();
    const formula = explicitSourceFormula ?? resolveChartSourceFormula(basis.document);
    const sourceArea = resolveFormulaExtent(basis.document, sheet.id, formula);
    const selArea = selectionAreas(selection)[0];
    const anchorRow = Math.max(1, sourceArea?.row ?? selArea?.top ?? selection.anchor.row);
    const anchorCol = Math.min(
      documentLimits.columns,
      (sourceArea ? sourceArea.column + sourceArea.columns : (selArea?.right ?? selection.anchor.column) + 1) + 1,
    );
    const prevCount = basis.document.charts?.length ?? 0;
    const cmd = createChart(basis, {
      chartType,
      source: { kind: 'range', formula },
      ...(explicitSeriesBy && { seriesBy: explicitSeriesBy }),
      placement: { sheetId: sheet.id, row: anchorRow, column: anchorCol, offsetX: 12, offsetY: 8, width: 480, height: 300 },
    }, makeId);
    applyStructure(basis, cmd);
    const inserted = cmd.document.charts?.[prevCount];
    if (inserted) setSelectedChartId(inserted.id);
    setInsertChartDraft(undefined);
  }

  function openInsertChartDialog(tab: 'recommended' | 'all', category?: CoreChartCategoryTab) {
    if (!editing.commit()) return;
    const basis = session.capture();
    const sourceFormula = resolveChartSourceFormula(basis.document);
    setInsertChartDraft({ tab, ...(category && { category }), sourceFormula });
  }

  function handleChartPlacement(chartId: DocumentId, placement: Partial<ChartPlacement>) {
    if (!editing.commit()) return;
    const basis = session.capture();
    applyStructure(basis, moveResizeChart(basis, chartId, placement));
  }

  function handleChartDelete(chartId: DocumentId) {
    if (!editing.commit()) return;
    const basis = session.capture();
    applyStructure(basis, deleteChart(basis, chartId));
    if (selectedChartId === chartId) setSelectedChartId(undefined);
  }

  function openCreatePivotDialog() {
    if (!editing.commit()) return;
    const basis = session.capture();
    const selArea = selectionAreas(selection)[0];
    let sourceText = '';
    if (activeTable) {
      sourceText = activeTable.name;
    } else if (selArea && (selArea.bottom > selArea.top || selArea.right > selArea.left)) {
      sourceText = refersTo(basis.document, {
        sheetId: sheet.id,
        row: selArea.top,
        column: selArea.left,
        rows: selArea.bottom - selArea.top + 1,
        columns: selArea.right - selArea.left + 1,
      }).replace(/^=/u, '');
    } else {
      const region = currentRegion(basis.document, {
        sheetId: sheet.id,
        row: selection.anchor.row,
        column: selection.anchor.column,
        rows: 1,
        columns: 1,
      });
      sourceText = refersTo(basis.document, region).replace(/^=/u, '');
    }
    const locationText = refersTo(basis.document, {
      sheetId: sheet.id,
      row: selection.anchor.row,
      column: selection.anchor.column,
      rows: 1,
      columns: 1,
    }).replace(/^=/u, '');
    setCreatePivotDraft({ sourceText, placement: 'new', locationText });
  }

  function handlePickPivotSource() {
    if (!createPivotDraft) return;
    const trimmed = createPivotDraft.sourceText.trim();
    setRangePicker({
      initial: trimmed ? (trimmed.startsWith('=') ? trimmed : `=${trimmed}`) : '',
      onDone: picked => {
        setRangePicker(undefined);
        if (picked) {
          setCreatePivotDraft(current => current ? { ...current, sourceText: picked.replace(/^=/u, '') } : undefined);
        }
      },
    });
  }

  function handlePickPivotLocation() {
    if (!createPivotDraft) return;
    const trimmed = createPivotDraft.locationText.trim();
    setRangePicker({
      initial: trimmed ? (trimmed.startsWith('=') ? trimmed : `=${trimmed}`) : '',
      onDone: picked => {
        setRangePicker(undefined);
        if (picked) {
          setCreatePivotDraft(current => current ? { ...current, placement: 'existing', locationText: picked.replace(/^=/u, '') } : undefined);
        }
      },
    });
  }

  function handleCreatePivotSubmit(): string | undefined {
    if (!createPivotDraft) return undefined;
    try {
      const basis = session.capture();
      const source = parsePivotRangeReference(basis.document, sheet.id, createPivotDraft.sourceText);
      requireDocument(source, 'Enter a valid table name or worksheet cell range.');
      let created: ReturnType<typeof createPivotTableWithSheet>;
      let targetCell: { row: number; column: number };
      if (createPivotDraft.placement === 'new') {
        created = createPivotTableWithSheet(basis, { source, newSheet: true, insertBeforeSheetId: sheet.id }, makeId);
        targetCell = { row: 3, column: 1 };
      } else {
        const target = parsePivotTargetCell(basis.document, sheet.id, createPivotDraft.locationText);
        requireDocument(target, 'Enter a valid destination cell reference.');
        created = createPivotTableWithSheet(basis, { source, target }, makeId);
        targetCell = { row: target.row, column: target.column };
      }
      applyStructure(basis, created.command, created.sheetId);
      setSelections(current => new Map(current).set(created.sheetId, singleSelection(targetCell)));
      setPivotFieldListOpen(true);
      setCreatePivotDraft(undefined);
      return undefined;
    } catch (err: unknown) {
      return err instanceof Error ? err.message : String(err);
    }
  }

  function handleUpdatePivotLayout(
    pivotId: DocumentId,
    patch: Partial<Omit<PivotTableDocument, 'id' | 'extent'>>,
  ) {
    try {
      const basis = session.capture();
      applyStructure(basis, updatePivotLayout(basis, pivotId, patch, makeId));
    } catch (err: unknown) {
      editing.reject(err instanceof Error ? err : new Error(String(err)));
    }
  }

  function handleSaveValueFieldSettings(updated: PivotValueField) {
    if (!fieldSettingsModal || fieldSettingsModal.request.mode !== 'value') return;
    const { pivotId, request } = fieldSettingsModal;
    const currentPivot = snapshot.document.pivots?.find(item => item.id === pivotId);
    if (!currentPivot) { setFieldSettingsModal(undefined); return; }
    const values = currentPivot.values.map(item => (item.id === request.field.id ? updated : item));
    handleUpdatePivotLayout(pivotId, { values });
    setFieldSettingsModal(undefined);
  }

  function handleSaveAxisFieldSettings(area: 'rows' | 'columns', updated: PivotAxisField) {
    if (!fieldSettingsModal || fieldSettingsModal.request.mode !== 'axis') return;
    const { pivotId, request } = fieldSettingsModal;
    const currentPivot = snapshot.document.pivots?.find(item => item.id === pivotId);
    if (!currentPivot) { setFieldSettingsModal(undefined); return; }
    const list = (area === 'rows' ? currentPivot.rows : currentPivot.columns)
      .map(item => (item.fieldName === request.field.fieldName ? updated : item));
    handleUpdatePivotLayout(pivotId, area === 'rows' ? { rows: list } : { columns: list });
    setFieldSettingsModal(undefined);
  }

  function handlePivotToggle(pivotId: DocumentId, key: string) {
    if (!editing.commit()) return;
    try {
      const basis = session.capture();
      applyStructure(basis, togglePivotCollapse(basis, pivotId, key, makeId));
    } catch (err: unknown) {
      editing.reject(err instanceof Error ? err : new Error(String(err)));
    }
  }

  function handlePivotFilterApply(
    pivotId: DocumentId,
    kind: 'reportFilter' | 'rowLabel' | 'colLabel',
    fieldName: string,
    hiddenItems: readonly string[],
  ) {
    setPivotFilterMenu(undefined);
    const basis = session.capture();
    const currentPivot = basis.document.pivots?.find(item => item.id === pivotId);
    if (!currentPivot) return;
    const updateItems = <T extends { readonly fieldName: string; readonly hiddenItems?: readonly string[] }>(items: readonly T[]) =>
      items.map(item => {
        if (item.fieldName !== fieldName) return item;
        const { hiddenItems: _prev, ...rest } = item;
        return (hiddenItems.length > 0 ? { ...rest, hiddenItems } : rest) as T;
      });
    if (kind === 'reportFilter') {
      handleUpdatePivotLayout(pivotId, { filters: updateItems(currentPivot.filters) });
    } else if (kind === 'rowLabel') {
      handleUpdatePivotLayout(pivotId, { rows: updateItems(currentPivot.rows) });
    } else {
      handleUpdatePivotLayout(pivotId, { columns: updateItems(currentPivot.columns) });
    }
  }

  function handlePivotFilterSort(
    pivotId: DocumentId,
    kind: 'rowLabel' | 'colLabel',
    fieldName: string,
    direction: 'asc' | 'desc',
  ) {
    setPivotFilterMenu(undefined);
    const currentPivot = snapshot.document.pivots?.find(item => item.id === pivotId);
    if (!currentPivot) return;
    const list = (kind === 'rowLabel' ? currentPivot.rows : currentPivot.columns).map(item =>
      item.fieldName === fieldName
        ? { ...item, sort: { direction } }
        : item,
    );
    handleUpdatePivotLayout(pivotId, kind === 'rowLabel' ? { rows: list } : { columns: list });
  }

  function handlePivotMoreSortOptions(
    pivotId: DocumentId,
    kind: 'rowLabel' | 'colLabel',
    fieldName: string,
  ) {
    setPivotFilterMenu(undefined);
    const currentPivot = snapshot.document.pivots?.find(item => item.id === pivotId);
    if (!currentPivot) return;
    const area = kind === 'colLabel' ? 'columns' : 'rows';
    const field = (area === 'columns' ? currentPivot.columns : currentPivot.rows).find(item => item.fieldName === fieldName);
    if (!field) return;
    setFieldSettingsModal({ pivotId, request: { mode: 'axis', area, field } });
  }

  function handlePivotDrillDown(target: CellReference): boolean {
    const pivot = findPivotAt(snapshot.document, target);
    if (!pivot) return false;
    if (!editing.commit()) return true;
    try {
      const basis = session.capture();
      const result = drillDownPivotCell(basis, pivot.id, target, makeId);
      if (!result) return false;
      applyStructure(basis, result.command, result.sheetId);
      setSelections(current => new Map(current).set(result.sheetId, singleSelection({ row: 1, column: 1 })));
      return true;
    } catch (err: unknown) {
      editing.reject(err instanceof Error ? err : new Error(String(err)));
      return true;
    }
  }

  function handlePivotRename(name: string) {
    if (!activePivot || !editing.commit()) return;
    try {
      const basis = session.capture();
      applyStructure(basis, renamePivotTable(basis, activePivot.id, name, makeId));
    } catch (err: unknown) {
      editing.reject(err instanceof Error ? err : new Error(String(err)));
    }
  }

  function handleOpenPivotFieldSettings() {
    if (!activePivot || !activePivotField) return;
    if (activePivotField.valueField) {
      setFieldSettingsModal({ pivotId: activePivot.id, request: { mode: 'value', field: activePivotField.valueField } });
    } else if (activePivotField.axisField && (activePivotField.area === 'rows' || activePivotField.area === 'columns')) {
      setFieldSettingsModal({
        pivotId: activePivot.id,
        request: { mode: 'axis', area: activePivotField.area, field: activePivotField.axisField },
      });
    }
  }

  function handlePivotExpandCollapse(expand: boolean) {
    if (!activePivot || !activePivotField?.canExpandCollapse || !editing.commit()) return;
    try {
      const basis = session.capture();
      applyStructure(basis, expandCollapsePivotField(basis, activePivot.id, activePivotField.fieldName, expand, makeId));
    } catch (err: unknown) {
      editing.reject(err instanceof Error ? err : new Error(String(err)));
    }
  }

  function handlePivotRefresh() {
    if (!activePivot || !editing.commit()) return;
    try {
      const basis = session.capture();
      applyStructure(basis, refreshPivotTable(basis, activePivot.id, makeId));
    } catch (err: unknown) {
      editing.reject(err instanceof Error ? err : new Error(String(err)));
    }
  }

  function handlePivotRefreshAll() {
    if (!editing.commit()) return;
    try {
      const basis = session.capture();
      applyStructure(basis, refreshAllPivotTables(basis, makeId));
    } catch (err: unknown) {
      editing.reject(err instanceof Error ? err : new Error(String(err)));
    }
  }

  function handlePivotClear(mode: 'all' | 'filters') {
    if (!activePivot || !editing.commit()) return;
    try {
      const basis = session.capture();
      applyStructure(basis, clearPivotTable(basis, activePivot.id, mode, makeId));
    } catch (err: unknown) {
      editing.reject(err instanceof Error ? err : new Error(String(err)));
    }
  }

  function handlePivotSelectEntire() {
    if (!activePivot?.extent || !editing.commit()) return;
    const ext = activePivot.extent;
    setSelections(current => new Map(current).set(activePivot.target.sheetId, {
      ...singleSelection({ row: ext.row, column: ext.column }),
      end: { row: ext.row + ext.rows - 1, column: ext.column + ext.columns - 1 },
    }));
    scheduleFocusGrid();
  }

  function handleOpenChangePivotSource() {
    if (!activePivot || !editing.commit()) return;
    const src = activePivot.source;
    const sourceText = src.kind === 'table'
      ? (snapshot.document.names.find(entry => entry.id === src.tableId)?.name ?? '')
      : refersTo(snapshot.document, src.area).replace(/^=/u, '');
    setChangePivotSourceDraft({ pivotId: activePivot.id, sourceText });
  }

  function handlePickChangePivotSource() {
    if (!changePivotSourceDraft) return;
    const trimmed = changePivotSourceDraft.sourceText.trim();
    setRangePicker({
      initial: trimmed ? (trimmed.startsWith('=') ? trimmed : `=${trimmed}`) : '',
      onDone: picked => {
        setRangePicker(undefined);
        if (picked) {
          setChangePivotSourceDraft(current => current ? { ...current, sourceText: picked.replace(/^=/u, '') } : undefined);
        }
      },
    });
  }

  function handleChangePivotSourceSubmit(): string | undefined {
    if (!changePivotSourceDraft) return undefined;
    try {
      const basis = session.capture();
      const source = parsePivotRangeReference(basis.document, sheet.id, changePivotSourceDraft.sourceText);
      requireDocument(source, 'Enter a valid table name or worksheet cell range.');
      applyStructure(basis, changePivotSource(basis, changePivotSourceDraft.pivotId, source, makeId));
      setChangePivotSourceDraft(undefined);
      scheduleFocusGrid();
      return undefined;
    } catch (err: unknown) {
      return err instanceof Error ? err.message : String(err);
    }
  }

  function handleOpenMovePivot() {
    if (!activePivot || !editing.commit()) return;
    const locationText = refersTo(snapshot.document, {
      sheetId: activePivot.target.sheetId,
      row: activePivot.target.row,
      column: activePivot.target.column,
      rows: 1,
      columns: 1,
    }).replace(/^=/u, '');
    setMovePivotDraft({ pivotId: activePivot.id, placement: 'existing', locationText });
  }

  function handlePickMovePivotLocation() {
    if (!movePivotDraft) return;
    const trimmed = movePivotDraft.locationText.trim();
    setRangePicker({
      initial: trimmed ? (trimmed.startsWith('=') ? trimmed : `=${trimmed}`) : '',
      onDone: picked => {
        setRangePicker(undefined);
        if (picked) {
          setMovePivotDraft(current => current ? { ...current, placement: 'existing', locationText: picked.replace(/^=/u, '') } : undefined);
        }
      },
    });
  }

  function handleMovePivotSubmit(): string | undefined {
    if (!movePivotDraft) return undefined;
    try {
      const basis = session.capture();
      const currentPivot = basis.document.pivots?.find(item => item.id === movePivotDraft.pivotId);
      requireDocument(currentPivot, 'PivotTable no longer exists.');
      let moved: ReturnType<typeof movePivotTableWithSheet>;
      let targetAnchor: Coordinate;
      if (movePivotDraft.placement === 'new') {
        moved = movePivotTableWithSheet(basis, movePivotDraft.pivotId, { newSheet: true, insertBeforeSheetId: sheet.id }, makeId);
        targetAnchor = { row: 3, column: 1 };
      } else {
        const target = parsePivotTargetCell(basis.document, sheet.id, movePivotDraft.locationText);
        requireDocument(target, 'Enter a valid destination cell reference.');
        moved = movePivotTableWithSheet(basis, movePivotDraft.pivotId, { target }, makeId);
        targetAnchor = { row: target.row, column: target.column };
      }
      applyStructure(basis, moved.command, moved.sheetId);
      setSelections(current => new Map(current).set(moved.sheetId, singleSelection(targetAnchor)));
      setMovePivotDraft(undefined);
      scheduleFocusGrid();
      return undefined;
    } catch (err: unknown) {
      return err instanceof Error ? err.message : String(err);
    }
  }

  function handleTableRename(name: string) {
    if (!activeTable) return;
    try {
      const basis = session.capture();
      applyStructure(basis, renameTable(basis, activeTable.id, name));
    } catch (err: unknown) {
      editing.reject(err instanceof Error ? err : new Error(String(err)));
    }
  }

  function handleTableChangeDesign(patch: Partial<TableDesign>) {
    if (!activeTable) return;
    try {
      const basis = session.capture();
      applyStructure(basis, changeTableDesign(basis, activeTable.id, patch, makeId));
    } catch (err: unknown) {
      editing.reject(err instanceof Error ? err : new Error(String(err)));
    }
  }

  function handleTableResize(table: WorkbookName, rows: number, cols: number) {
    try {
      const basis = session.capture();
      const nextArea = { ...table.area, rows, columns: cols };
      applyStructure(basis, resizeTable(basis, { ...table, area: nextArea }, makeId));
    } catch (err: unknown) {
      editing.reject(err instanceof Error ? err : new Error(String(err)));
    }
  }

  function handleTotalFunction(tableId: string, col: number, code: number) {
    try {
      const basis = session.capture();
      applyStructure(basis, setTotalFunction(basis, tableId, col, code, makeId));
    } catch (err: unknown) {
      editing.reject(err instanceof Error ? err : new Error(String(err)));
    }
  }

  function runSort(range: SortExtent, draft: SortDraft) {
    const basis = session.capture();
    const outcome = session.execute(sortData(basis, { sheetId: sheet.id, ...range }, draft.column,
      draft.header, draft.ascending, draft.caseSensitive, textOrder));
    if (outcome.changed) {
      editing.selectionChanged();
      editing.clearCut();
      workbookChanged(true);
    }
    editing.report('Sorted ' + rangeLabel(range) + ' by column ' + columnLetters(draft.column) + ', ' +
      (draft.ascending ? 'smallest to largest' : 'largest to smallest') + '. Records moved together.');
    scheduleFocusGrid();
  }

  function requestSort(ascending: boolean, dialog: boolean, choice?: 'expand' | 'current') {
    if (!editing.commit()) return;
    let selectedRange: EditRange;
    try {
      const chosen = ranges();
      requireDocument(chosen.length === 1, 'Select one contiguous range to sort.');
      selectedRange = chosen[0]!;
    } catch (error) { editing.reject(error); return; }
    try {
      const basis = session.capture();
      const origin = activeTable ? activeTable.area : currentRegion(basis.document, selectedRange);
      const top = Math.min(origin.row, selectedRange.row);
      const left = Math.min(origin.column, selectedRange.column);
      const region: SortExtent = { row: top, column: left,
        rows: Math.max(lastRow(origin), lastRow(selectedRange)) - top + 1,
        columns: Math.max(lastColumn(origin), lastColumn(selectedRange)) - left + 1 };
      let range: SortExtent;
      if (selectedRange.rows * selectedRange.columns === 1) range = region;
      else if (region.row < selectedRange.row || lastRow(region) > lastRow(selectedRange) ||
          region.column < selectedRange.column || lastColumn(region) > lastColumn(selectedRange)) {
        if (choice === undefined) {
          setSortWarning({ selection: selectedRange, region, resume: expand => {
            setSortWarning(undefined);
            if (expand !== null) requestSort(ascending, dialog, expand ? 'expand' : 'current');
          } });
          return;
        }
        range = choice === 'expand' ? region : selectedRange;
      } else range = selectedRange;
      const filterHeader = sheet.filter && range.row === sheet.filter.row && lastRow(range) === lastRow(sheet.filter);
      const applied = filterHeader && sheet.filter?.sort && sheet.filter.sort.column >= range.column &&
        sheet.filter.sort.column <= lastColumn(range) ? sheet.filter.sort : undefined;
      const draft: SortDraft = {
        column: dialog && applied ? applied.column : clamp(selection.anchor.column, range.column, lastColumn(range)),
        header: activeTable ? tableDesign(activeTable).headerRow :
          !!(filterHeader || looksLikeHeader(basis, { sheetId: sheet.id, ...range })),
        ascending: dialog && applied ? applied.ascending : ascending,
        caseSensitive: false,
      };
      if (dialog) { setSortDialog({ basis, range, draft }); return; }
      runSort(range, draft);
    } catch (error) { editing.reject(error); }
  }

  function submitSort(draft: SortDraft) {
    const open = sortDialog;
    if (!open) return;
    setSortDialog(undefined);
    try {
      requireDocument(session.isCurrentAuthored(open.basis), 'The workbook changed. Open Sort again.');
      runSort(open.range, draft);
    } catch (error) { editing.reject(error); }
  }

  function openSortOptions(draft: SortDraft) {
    if (!sortDialog) return;
    setSortOptions({ ...sortDialog, draft });
    setSortDialog(undefined);
  }

  function finishSortOptions(caseSensitive: boolean) {
    if (!sortOptions) return;
    setSortDialog({ ...sortOptions, draft: { ...sortOptions.draft, caseSensitive } });
    setSortOptions(undefined);
  }

  function cancelSortOptions() {
    if (!sortOptions) return;
    setSortDialog(sortOptions);
    setSortOptions(undefined);
  }

  function toggleFilter() {
    if (!editing.commit()) return;
    try {
      if (activeTable) {
        const design = tableDesign(activeTable);
        requireDocument(design.headerRow, "Show the table's Header Row before enabling filter buttons.");
        handleTableChangeDesign({ filterButton: !design.filterButton });
        return;
      }
      const basis = session.capture();
      const current = basis.document.sheets.find(item => item.id === sheet.id)!;
      const ownedByTable = !!current.filter && tables.some(table => ownsTableFilter(current.filter, table));
      if (current.filter && !ownedByTable) {
        session.execute(setFilter(basis, sheet.id, null));
        editing.selectionChanged();
        editing.clearCut();
        workbookChanged(true);
        editing.report('Filter removed. Manually hidden rows stay hidden.');
        scheduleFocusGrid();
        return;
      }
      const chosen = ranges();
      requireDocument(chosen.length === 1, 'Select one contiguous range to filter.');
      const target = filterRange(basis.document, chosen[0]!);
      requireDocument(current.cells.some(cell => extentContains(target, { row: cell.row, column: cell.column })),
        'Select a range with data before turning on the filter.');
      // Outside a table the toggle reads off even while that table's filter runs, so it starts a range
      // filter on the caret's data instead of removing the table's. The sheet holds one filter, so this
      // replaces it the way the desktop ToggleFilterAsync does.
      session.execute(setFilter(basis, sheet.id, { row: target.row, column: target.column,
        rows: target.rows, columns: target.columns, criteria: [] }));
      editing.selectionChanged();
      editing.clearCut();
      workbookChanged(true);
      editing.report('Filter applied to ' + rangeLabel(target) + '. Choose values from a header funnel.');
      scheduleFocusGrid();
    } catch (error) { editing.reject(error); }
  }

  function clearFilter() {
    if (!editing.commit()) return;
    try {
      const basis = session.capture();
      const current = basis.document.sheets.find(item => item.id === sheet.id)!;
      requireDocument(current.filter, 'There is no filter to clear.');
      session.execute(setFilter(basis, sheet.id, { ...current.filter, criteria: [] }));
      editing.selectionChanged();
      editing.clearCut();
      workbookChanged(true);
      editing.report('Filter cleared. Manually hidden rows stay hidden.');
      scheduleFocusGrid();
    } catch (error) { editing.reject(error); }
  }

  function reapplyFilter() {
    if (!editing.commit()) return;
    if (!sheet.filter) { editing.reject(new Error('There is no filter to reapply.')); return; }
    workbookChanged(true);
    editing.report('Filter reapplied. ' + (recordSummary(sheet) ?? 'All records are visible.'));
    scheduleFocusGrid();
  }

  function openFilter(row: number, column: number, x: number, y: number) {
    setFilterMenu(undefined);
    if (!editing.commit()) return;
    try {
      let basis = session.capture();
      let current = basis.document.sheets.find(item => item.id === sheet.id)!;
      const table = tablesOnSheet(basis.document, sheet.id).find(item => item.area.row === row &&
        column >= item.area.column && column < lastColumn(item.area));
      if (table) {
        const design = tableDesign(table);
        if (!design.headerRow || !design.filterButton) return;
        const input = tableInputExtent(table);
        if (current.filter && (current.filter.row !== input.row || current.filter.column !== input.column) &&
            current.filter.criteria.length > 0) {
          throw new Error("Clear the other range's filter before filtering this table.");
        }
        if (!current.filter || current.filter.row !== input.row || current.filter.column !== input.column) {
          session.execute(setFilter(basis, sheet.id, { row: input.row, column: input.column,
            rows: input.rows, columns: input.columns, criteria: [] }));
          basis = session.capture();
          current = basis.document.sheets.find(item => item.id === sheet.id)!;
          workbookChanged(true);
        }
      }
      const filter = current.filter;
      if (!filter || column < filter.column || column > lastColumn(filter)) return;
      const values = columnValues(current, filter, column, textOrder);
      const criterion = filter.criteria.find(item => item.column === column);
      const textKeys = current.cells.filter(cell => cell.column === column && cell.row > filter.row &&
        cell.row <= lastRow(filter)).map(cell => sortKeyFor(cell, cell.row)).filter(key => key.kind !== 'blank');
      setFilterMenu({ basis, column, header: headerLabel(current, filter.row, column),
        textColumn: textKeys.length > 0 && textKeys.every(key => key.kind === 'text'),
        values, selected: criterion ? [...criterion.allowed] : values.map(value => value.key),
        hasColumnFilter: !!criterion, anchor: { x, y } });
    } catch (error) { editing.reject(error); }
  }

  function applyFilter(selected: readonly string[]) {
    const open = filterMenu;
    if (!open) return;
    setFilterMenu(undefined);
    try {
      requireDocument(session.isCurrentAuthored(open.basis), 'Workbook changed. Open the filter again to choose current values.');
      const latest = session.capture();
      const live = latest.document.sheets.find(item => item.id === sheet.id)?.filter;
      requireDocument(live, 'The filter was removed.');
      const valueKeys = new Set(open.values.map(value => value.key));
      const criterion = live.criteria.find(item => item.column === open.column);
      const chosen = new Set(selected);
      const picked = open.values.filter(value => chosen.has(value.key)).map(value => value.key);
      const kept = criterion ? criterion.allowed.filter(key => !valueKeys.has(key)) : [];
      const criteria = live.criteria.filter(item => item.column !== open.column);
      if (picked.length !== open.values.length || (criterion && criteria.length > 0)) {
        criteria.push({ column: open.column, allowed: [...picked, ...kept] });
      }
      criteria.sort((left, right) => left.column - right.column);
      const outcome = session.execute(setFilter(latest, sheet.id, { ...live, criteria }));
      if (outcome.changed) {
        editing.selectionChanged();
        editing.clearCut();
        workbookChanged(true);
      }
      const applied = session.active.sheets.find(item => item.id === sheet.id)!;
      editing.report('Filtered ' + open.header + '. ' + (recordSummary(applied) ?? ''));
    } catch (error) { editing.reject(error); }
    scheduleFocusGrid();
  }

  function sortFilterColumn(ascending: boolean) {
    const open = filterMenu;
    if (!open) return;
    setFilterMenu(undefined);
    try {
      const basis = session.capture();
      const live = basis.document.sheets.find(item => item.id === sheet.id)?.filter;
      requireDocument(live, 'There is no filter to sort.');
      const outcome = session.execute(sortData(basis, { sheetId: sheet.id, row: live.row, column: live.column,
        rows: live.rows, columns: live.columns }, open.column, true, ascending, false, textOrder));
      if (outcome.changed) {
        editing.selectionChanged();
        editing.clearCut();
        workbookChanged(true);
      }
      editing.report('Sorted the filtered records. Hidden rows were included and the filter still applies.');
    } catch (error) { editing.reject(error); }
    scheduleFocusGrid();
  }

  function clearFilterColumn() {
    const open = filterMenu;
    if (!open) return;
    setFilterMenu(undefined);
    try {
      const basis = session.capture();
      const live = basis.document.sheets.find(item => item.id === sheet.id)?.filter;
      requireDocument(live, 'There is no filter to clear.');
      const outcome = session.execute(setFilter(basis, sheet.id, { ...live,
        criteria: live.criteria.filter(item => item.column !== open.column) }));
      if (outcome.changed) {
        editing.selectionChanged();
        editing.clearCut();
        workbookChanged(true);
      }
      editing.report('Cleared the filter for this column. Manually hidden rows stay hidden.');
    } catch (error) { editing.reject(error); }
    scheduleFocusGrid();
  }

  function errorText(error: unknown): string {
    return error instanceof Error ? error.message : String(error);
  }

  function rememberSearchValue(value: string, replacement = false) {
    setSearchHistory(current => replacement ?
      { ...current, replace: rememberSearch(current.replace, value, true) } :
      { ...current, find: rememberSearch(current.find, value) });
  }

  function openSearch(replace: boolean) {
    if (!editing.commit()) return;
    if (root.current?.querySelector('dialog[open]')) return;
    setSearchDialog({ replace, sheetId: sheet.id });
  }

  const searchRequest = useRef<(replace: boolean) => void>(() => {});
  searchRequest.current = openSearch;
  useEffect(() => {
    function keydown(event: globalThis.KeyboardEvent) {
      const key = event.key.toLowerCase();
      if (!(event.ctrlKey || event.metaKey) || event.altKey || (key !== 'f' && key !== 'h')) return;
      event.preventDefault();
      searchRequest.current(key === 'h');
    }
    document.addEventListener('keydown', keydown);
    return () => document.removeEventListener('keydown', keydown);
  }, []);

  function searchScope(basis: WorkbookSnapshot, options: SearchOptions): readonly string[] {
    return options.within === 'workbook' ? basis.document.sheets.map(item => item.id) :
      [searchDialog?.sheetId ?? sheet.id];
  }

  function searchMatches(options: SearchOptions) {
    const basis = session.capture();
    return { basis, matches: findMatches(basis.document, searchScope(basis, options), options) };
  }

  function showSearchMatch(match: SearchMatch) {
    setSelections(current => new Map(current).set(match.sheetId, singleSelection({ row: match.row, column: match.column })));
    if (match.sheetId !== activeSheetId) setActiveSheetId(match.sheetId);
    requestAnimationFrame(() => worksheet.current?.reveal({ row: match.row, column: match.column }, false));
  }

  function runFindAll(options: SearchOptions): SearchOutcome {
    try {
      const { basis, matches } = searchMatches(options);
      rememberSearchValue(options.query);
      return { status: matches.length + ' cell(s) found.',
        rows: searchResultRows(basis, activeFile.name, matches, options) };
    } catch (error) { return { status: errorText(error) }; }
  }

  function runFindNext(options: SearchOptions, current: SearchMatch | undefined): SearchOutcome {
    try {
      const { basis, matches } = searchMatches(options);
      rememberSearchValue(options.query);
      if (!matches.length) return { status: 'No matches.' };
      let next: SearchMatch | undefined;
      const index = current ? matches.findIndex(match => match.sheetId === current.sheetId &&
        match.row === current.row && match.column === current.column) : -1;
      if (index >= 0) next = matches[(index + 1) % matches.length];
      else {
        const anchor = selection.anchor;
        next = matches.find(match => match.sheetId === sheet.id && (options.order === 'columns' ?
          match.column > anchor.column || match.column === anchor.column && match.row > anchor.row :
          match.row > anchor.row || match.row === anchor.row && match.column > anchor.column)) ?? matches[0];
      }
      if (next) showSearchMatch(next);
      return { status: matches.length + ' match(es). ' + (next ? matchLabel(basis.document, next, options.within) : ''),
        ...(next && { match: next }) };
    } catch (error) { return { status: errorText(error) }; }
  }

  function runReplace(options: SearchOptions, replacement: string, target: 'all' | 'one',
    selected: SearchMatch | undefined, current: SearchMatch | undefined, listShown: boolean): SearchOutcome {
    try {
      const { basis, matches } = searchMatches(options);
      rememberSearchValue(options.query);
      rememberSearchValue(replacement, true);
      const targets = target === 'all' ? matches : listShown && selected ? [selected] :
        listShown && current ? [current] :
        matches.filter(match => match.sheetId === sheet.id && match.row === selection.anchor.row &&
          match.column === selection.anchor.column);
      if (!targets.length) return { status: matches.length === 0 ? 'No matches.' : 'Find a matching cell first.' };
      requireDocument(session.isCurrent(basis), 'Workbook changed. Find again.');
      const outcome = session.execute(replaceMatches(basis, targets, options, replacement));
      if (outcome.changed) {
        editing.selectionChanged();
        editing.clearCut();
        workbookChanged(true);
      }
      let next: SearchMatch | undefined;
      if (target === 'one') {
        const index = matches.findIndex(match => match.sheetId === targets[0]!.sheetId &&
          match.row === targets[0]!.row && match.column === targets[0]!.column);
        // Replacing the last match does not wrap, matching the desktop Search dialog.
        if (index >= 0 && index + 1 < matches.length) {
          next = matches[index + 1];
          if (next) showSearchMatch(next);
        }
      }
      const refreshed = session.capture();
      const remaining = findMatches(refreshed.document, searchScope(refreshed, options), options);
      return { status: 'Replaced ' + targets.length + ' cell(s).',
        ...(listShown && { rows: searchResultRows(refreshed, activeFile.name, remaining, options) }),
        ...(next && { match: next }) };
    } catch (error) { return { status: errorText(error) }; }
  }

  function applyZoom(zoom: number | 'fit') {
    setZoomDialog(false);
    if (zoom === 'fit') {
      const fitted = worksheet.current?.zoomToSelection();
      if (fitted !== undefined) onView({ zoom: fitted });
    } else onView({ zoom });
    scheduleFocusGrid();
  }

  function runFreeze(mode: 'panes' | 'topRow' | 'firstColumn' | 'none') {
    if (!editing.commit()) return;
    const sheetId = sheet.id;
    try {
      const panes = mode === 'panes' ? freezeFromCaret(selection.anchor) :
        mode === 'topRow' ? { rows: 1, columns: 0 } :
        mode === 'firstColumn' ? { rows: 0, columns: 1 } : { rows: 0, columns: 0 };
      if (sheet.frozenPanes.rows === panes.rows && sheet.frozenPanes.columns === panes.columns) {
        editing.report('Pane boundaries are unchanged.');
        scheduleFocusGrid();
        return;
      }
      const basis = session.capture();
      applyStructure(basis, setFreeze(basis, sheetId, panes));
      editing.report(panes.rows === 0 && panes.columns === 0 ? 'Unfroze panes.' :
        mode === 'topRow' ? 'Froze the top row.' :
        mode === 'firstColumn' ? 'Froze the first column.' : 'Froze panes.');
    } catch (error) { editing.reject(error); }
    scheduleFocusGrid();
  }

  const tableFilter = !!sheet.filter && tables.some(table => ownsTableFilter(sheet.filter, table));
  const filterActive = tableFilter ? !!sheet.filter && !!activeTable && ownsTableFilter(sheet.filter, activeTable) : !!sheet.filter;
  const filterClearable = !!sheet.filter?.criteria.length;
  const filterReapplyable = !!sheet.filter;
  const filterCommands = useMemo(() => ({
    active: filterActive, clearable: filterClearable, reapplyable: filterReapplyable,
  }), [filterActive, filterClearable, filterReapplyable]);

  function currentAuthoredBasis(basis: WorkbookSnapshot) {
    requireDocument(session.isCurrentAuthored(basis), 'The workbook changed. Open the command again.');
    return session.capture();
  }

  function applyStyleCommand(command: WorkbookCommand) {
    const outcome = session.execute(command);
    editing.clearStatus();
    if (outcome.changed) {
      editing.selectionChanged();
      editing.clearCut();
      workbookChanged(true);
    }
  }

  function requestMerge(captured: WorkbookSnapshot, range: EditRange, mode: MergeMode,
    formatting?: { patch: AppearancePatch; code: string | undefined; borders: BorderDraft; alignment: AlignmentPatch }) {
    const basis = currentAuthoredBasis(captured);
    const proposal = proposeMerge(basis, range, mode, makeId);
    function apply() {
      const current = currentAuthoredBasis(captured);
      const { command, range: expanded } = current === basis ? proposal : proposeMerge(current, range, mode, makeId);
      if (formatting) {
        const formatted = formatStyles({ ...current, document: command.document }, [expanded],
          formatting.patch, formatting.code, formatting.borders, formatting.alignment);
        applyStructure(current, formatted.kind === 'set-cells'
          ? { ...command, label: 'Format Cells', edits: formatted.edits }
          : { ...command, label: 'Format Cells', document: formatted.document, ...(formatted.edits && { edits: formatted.edits }) });
        setFormatDraft(undefined);
      } else applyStructure(current, command);
      setMergeConfirmation(undefined);
      setSelections(values => new Map(values).set(range.sheetId, { ...singleSelection({ row: expanded.row, column: expanded.column }),
        end: { row: expanded.row + expanded.rows - 1, column: expanded.column + expanded.columns - 1 } }));
    }
    if (proposal.discards) setMergeConfirmation({ confirm: apply });
    else apply();
  }

  function runMerge(mode: MergeMode) {
    if (!editing.commit()) return;
    try {
      const selected = ranges();
      requireDocument(selected.length === 1, 'Select a single range before merging cells.');
      requestMerge(session.capture(), selected[0]!, mode);
    } catch (error) { editing.reject(error); }
  }

  function runAlignment(action: AlignmentAction) {
    if (!editing.commit()) return;
    try {
      const basis = session.capture();
      const targetSheet = basis.document.sheets.find(item => item.id === activeTarget.sheetId);
      const style = targetSheet ? resolvedCellStyle(targetSheet, activeTarget) : findReferencedCell(basis.document, activeTarget)?.style;
      applyStyleCommand(changeAlignment(basis, ranges(), action, style));
      scheduleFocusGrid();
    } catch (error) { editing.reject(error); }
  }

  function runNumberFormat(change: NumberFormatChange) {
    if (!editing.commit()) return;
    setBorderTool(undefined);
    setBrush(undefined);
    try { applyStyleCommand(applyNumberFormat(session.capture(), ranges(), change)); }
    catch (error) { editing.reject(error); }
    scheduleFocusGrid();
  }

  function runAppearance(patch: AppearancePatch, focus = true) {
    if (!editing.commit()) return;
    if (Object.keys(patch).length) {
      setBrush(undefined);
      setBorderTool(undefined);
      try { applyStyleCommand(formatStyles(session.capture(), ranges(), patch)); }
      catch (error) { editing.reject(error); }
    }
    if (focus) scheduleFocusGrid();
  }

  function armPainter(locked: boolean) {
    if (!editing.commit()) return;
    setBorderTool(undefined);
    try {
      if (painter && (!locked || painter.locked)) setBrush(undefined);
      else {
        const selected = ranges();
        requireDocument(selected.length === 1, 'Format Painter needs a contiguous source range.');
        setBrush({ ...(painter ?? captureFormatBrush(session.capture(), selected[0]!)), locked });
      }
      editing.clearStatus();
    } catch (error) { editing.reject(error); }
    scheduleFocusGrid();
  }

  function paint(area: Area) {
    if (!painter || !editing.commit()) return;
    try {
      const target = { sheetId: sheet.id, row: area.top, column: area.left, rows: area.bottom - area.top + 1,
        columns: area.right - area.left + 1 };
      if (target.rows === 1 && target.columns === 1) {
        requireDocument(target.row + painter.rows - 1 <= worksheetViewport.rows &&
          target.column + painter.columns - 1 <= worksheetViewport.columns, 'Format Painter destination is outside the worksheet view.');
      }
      applyStyleCommand(paintStyles(session.capture(), target, painter));
      if (!painter.locked) setBrush(undefined);
    } catch (error) { editing.reject(error); }
  }

  function runBorder(action: BorderAction) {
    if (!editing.commit()) return;
    try {
      const basis = session.capture();
      if (Object.hasOwn(borderModes, action)) {
        setBrush(undefined);
        setBorderTool({ mode: action as BorderMode, sheetId: sheet.id, sessionId: basis.sessionId });
        editing.clearStatus();
      } else {
        setBorderTool(undefined);
        const selected = ranges();
        requireDocument(selected.length === 1, 'Select one contiguous range for borders.');
        applyStyleCommand(applyBorderPreset(basis, selected[0]!, action as BorderPreset, borderPen(lineStyle, lineColor), lineColor));
      }
      rememberBorder(action);
    } catch (error) { editing.reject(error); }
    scheduleFocusGrid();
  }

  function beginBorderStroke() {
    if (!borderTool || !editing.commit()) return;
    const basis = session.capture();
    const pen = borderTool.mode === 'erase' ? null : borderPen(lineStyle, lineColor);
    return (gesture: DrawGesture) => {
      try {
        const current = currentAuthoredBasis(basis);
        applyStyleCommand('area' in gesture ? applyBorderPreset(current, { sheetId: sheet.id, ...gesture.area },
          borderTool.mode === 'grid' ? 'all' : 'outside', pen) : applyBorderStroke(current, sheet.id, gesture.segments, pen));
      } catch (error) { editing.reject(error); }
    };
  }

  function targetHeader(axis: Axis, index: number): StructureTarget {
    const area = primaryArea(selection);
    const first = axis === 'row' ? area.top : area.left;
    const last = axis === 'row' ? area.bottom : area.right;
    const selectedTrack = selection.kind === (axis === 'row' ? 'rows' : 'columns') && index >= first && index <= last;
    return { id: sheet.id, ranges: selectedTrack ? ranges() : [{
      sheetId: sheet.id, row: axis === 'row' ? index : 1, column: axis === 'column' ? index : 1, rows: 1, columns: 1,
    }], header: index };
  }

  function indexesFor(target: StructureTarget, axis: Axis): number[] {
    const indexes = new Set<number>();
    const count = axis === 'row' ? 'rows' : 'columns';
    for (const range of target.ranges) {
      for (let i = range[axis]; i < range[axis] + range[count]; i++) indexes.add(i);
    }
    return [...indexes];
  }

  function performStructure(action: StructureAction, target: StructureTarget, captured: WorkbookSnapshot, input?: StructureInput) {
    const basis = currentAuthoredBasis(captured);
    if (action.endsWith('-sheet')) {
      const change = action === 'insert-sheet' ? { kind: 'add', after: false } as const :
        action === 'delete-sheet' ? { kind: 'delete' } as const :
        { kind: 'place', before: input!.before, copy: input!.copy } as const;
      const next = changeSheet(basis, target.id, change, makeId);
      applyStructure(basis, next.command, next.active);
      return;
    }
    const axis: Axis = input && action.endsWith('-cells') ? input.axis : action.endsWith('-row') ? 'row' : 'column';
    const count = axis === 'row' ? 'rows' : 'columns';
    if (action.startsWith('insert-') || action.startsWith('delete-')) {
      requireDocument(target.ranges.length === 1, 'Select one contiguous area for insertion or deletion.');
      const range = target.ranges[0]!;
      const other = axis === 'row' ? 'column' : 'row';
      const otherCount = axis === 'row' ? 'columns' : 'rows';
      applyStructure(basis, editAxis(basis, target.id, {
        axis, start: range[axis], count: range[count], delete: action.startsWith('delete-'),
        ...(action.endsWith('-cells') && !input?.entire && { lane: [range[other], range[other] + range[otherCount] - 1] }),
      }));
      return;
    }
    const source = basis.document.sheets.find(item => item.id === target.id)!;
    const span = target.ranges.reduce((sum, range) => sum + range[count], 0);
    requireDocument(span <= documentLimits.cells, 'Selection is too large.');
    const indexes = indexesFor(target, axis);
    if (action.startsWith('show-') && target.header !== undefined) {
      const hidden = new Set(source[count].filter(dimension => dimension.hidden).map(dimension => dimension.index));
      for (let index = Math.min(...indexes) - 1; hidden.has(index); index--) indexes.push(index);
      for (let index = Math.max(...indexes) + 1; hidden.has(index); index++) indexes.push(index);
    }
    const patches = action.startsWith('fit-') ? fitDimensions(source, axis, indexes, view.showFormulas, separator) : indexes.map(index => ({
      index, ...(action.startsWith('size-') ? { size: input!.size } : { hidden: action.startsWith('hide-') }),
    }));
    applyStructure(basis, changeDimensions(basis, target.id, axis, patches));
  }

  function openStructure(action: StructureAction, target: StructureTarget = { id: sheet.id, ranges: ranges() }, tab: FormatTab = 'Number') {
    setMenu(undefined);
    if (!python.barrier() || !editing.commit()) return;
    editing.selectionChanged();
    const basis = session.capture();
    if (action === 'format-cells') {
      setBrush(undefined);
      setBorderTool(undefined);
      if (tab === 'Border' && target.ranges.length !== 1) {
        editing.reject(new Error('Select one contiguous range for borders.'));
        scheduleFocusGrid();
        return;
      }
      const first = target.ranges[0]!;
      const inside = target.ranges.some(range => activeTarget.sheetId === range.sheetId &&
        activeTarget.row >= range.row && activeTarget.row < range.row + range.rows &&
        activeTarget.column >= range.column && activeTarget.column < range.column + range.columns);
      setFormatDraft({ basis, tab, ranges: target.ranges, cell: inside ? activeTarget :
        { sheetId: first.sheetId, row: first.row, column: first.column } });
    } else if (action === 'rename-sheet') {
      renaming.current = true;
      setRename({ id: target.id, basis });
    } else if (dialogActions.has(action)) {
      setStructure({ action: action as DialogAction, target, basis });
    } else {
      try { performStructure(action, target, basis); } catch (error) { editing.reject(error); }
    }
  }

  function finishRename(value?: string) {
    if (!rename || !renaming.current) return;
    renaming.current = false;
    setRename(undefined);
    if (value === undefined) { scheduleFocusGrid(); return; }
    try {
      const basis = currentAuthoredBasis(rename.basis);
      const next = changeSheet(basis, rename.id, { kind: 'rename', name: value }, makeId);
      applyStructure(basis, next.command, next.active);
    } catch (error) { editing.reject(error); scheduleFocusGrid(); }
  }

  function openGridMenu(x: number, y: number, axis?: Axis, index?: number, point?: Coordinate) {
    if (!editing.commit()) return;
    const cell = point && (mergeAt(sheet, point.row, point.column) ?? { ...point, rows: 1, columns: 1 });
    const target = axis && index !== undefined ? targetHeader(axis, index) : { id: sheet.id, ranges:
      point && cell && !selectionAreas(selection).some(area => includes(area, point))
        ? [{ sheetId: sheet.id, ...cell }] : ranges() };
    setMenu({ x, y, target, actions: axis ? [
      `insert-${axis}`, `delete-${axis}`, `size-${axis}`, `fit-${axis}`, `hide-${axis}`, `show-${axis}`,
    ] : ['insert-cells', 'delete-cells', 'format-cells'] });
  }

  function run(action: EditAction, event?: DataTransfer) {
    if (!python.barrier()) return;
    setBrush(undefined);
    void editing.run(action, ranges(), event).then(target => { if (target) worksheet.current?.goTo(target); });
  }

  function moveHistory(undo: boolean) {
    if (!python.barrier()) return;
    editing.moveHistory(undo);
  }

  function requestPythonRun() {
    if (!python.barrier()) return;
    const context = python.context;
    const target = context && pythonAnchorAt(session.capture(), context.sheetId, context);
    if (!target) return;
    setFigureStatus(undefined);
    if (trustedPythonSession.current === target.sessionId) {
      void pythonExecution.run(target, calculation.stale);
    } else setPythonTrustPrompt({ target });
  }

  function requestFigurePlacement() {
    if (!python.barrier() || !editing.commit()) return;
    const context = python.context;
    const target = context && pythonAnchorAt(session.capture(), context.sheetId, context);
    const cell = target && session.active.sheets.find(item => item.id === target.sheetId)?.cells.find(item => item.id === target.cellId);
    const block = cell?.content.kind === 'python' ? cell.content.block : undefined;
    if (!block?.lastFigure) {
      setFileStatus('Run a Python Figure before placing it on the worksheet.');
      return;
    }
    setPendingFigureOwner(current => current === block.id ? undefined : block.id);
    worksheet.current?.focus();
    setFileStatus(pendingFigureOwner === block.id
      ? 'Figure placement cancelled.' : 'Move the pointer over the worksheet, then click or drag to place the Figure. Press Escape to cancel.');
  }

  async function saveFigureImage() {
    const figure = pythonBlock?.lastFigure;
    if (!figure) return;
    try {
      const binary = atob(figure.pngBase64);
      const bytes = new Uint8Array(binary.length);
      for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index);
      const url = URL.createObjectURL(new Blob([bytes], { type: 'image/png' }));
      const link = document.createElement('a');
      link.href = url;
      link.download = 'Python Figure.png';
      link.click();
      window.setTimeout(() => URL.revokeObjectURL(url), 1000);
      setFileStatus('Saved the last successful Python Figure image.');
    } catch (error) { setFileStatus(error instanceof Error ? error.message : 'The Figure image could not be saved.'); }
  }

  async function saveFigurePlacement(ownerId: string, requested: FigurePlacement, original?: FigurePlacement) {
    if (!python.barrier() || !editing.commit()) return;
    const basis = session.capture();
    let source: PythonFigure | null | undefined;
    let placements: readonly FigurePlacement[] = [];
    let sourceTarget: PythonAnchorTarget | undefined;
    for (const sourceSheet of basis.document.sheets) for (const cell of sourceSheet.cells) {
      if (cell.content.kind === 'python' && cell.content.block?.id === ownerId) {
        source = cell.content.block.lastFigure;
        placements = cell.content.block.placements ?? [];
        sourceTarget = pythonAnchorAt(basis, sourceSheet.id, { row: cell.row, column: cell.column });
      }
    }
    if (!source) { setFileStatus('The Python Figure source is no longer available.'); return; }
    const prior = original && placements.find(item => item.id === original.id);
    if (original && (!prior || ['sheetId', 'row', 'column', 'rows', 'columns', 'offsetX', 'offsetY', 'width', 'height']
      .some(key => prior?.[key as keyof FigurePlacement] !== original[key as keyof FigurePlacement]))) {
      const message = 'The Figure placement changed while Python was running. The requested resize was not applied.';
      setFigureStatus({ ownerId, message });
      setFileStatus(message);
      return;
    }
    let placement = requested;
    const targetSheet = basis.document.sheets.find(item => item.id === requested.sheetId);
    if (!targetSheet) { setFileStatus('The destination worksheet no longer exists.'); return; }
    try {
      const newCanvas = figureCanvasForPlacement(requested, targetSheet, window.devicePixelRatio || 1);
      let placedFromCachedImage = false;
      let cacheReusable = false;
      if (prior?.renderedFigure && prior.renderedSourceId) {
        const oldCanvas = figureCanvasForPlacement(prior, targetSheet, window.devicePixelRatio || 1);
        cacheReusable = oldCanvas.width === newCanvas.width && oldCanvas.height === newCanvas.height && oldCanvas.dpi === newCanvas.dpi;
      }
      if (cacheReusable && prior?.renderedFigure && prior.renderedSourceId) {
        placement = { ...requested, renderedFigure: prior.renderedFigure, renderedSourceId: prior.renderedSourceId };
      } else {
        setFigureStatus({ ownerId, message: 'Redrawing the Figure at its placement size…' });
        const redraw = await pythonExecution.redrawFigure(source, newCanvas.width, newCanvas.height, newCanvas.dpi);
        if (!redraw.figure) {
          if (original) {
            if (pythonExecutionState?.running) {
              setFigureStatus({ ownerId, message: 'Wait for the current Python run to finish before resizing this Figure.' });
              setFileStatus('Figure resize paused until the current Python run finishes. The existing placement is unchanged.');
              return;
            }
            if (!sourceTarget) {
              setFigureStatus({ ownerId, message: 'The Python source for this Figure is no longer available.' });
              setFileStatus('The Figure could not be recalculated because its Python source is unavailable.');
              return;
            }
            setPendingFigureResize({ ownerId, requested, original, target: sourceTarget,
              baseRevision: basis.revision, runningObserved: false });
            setFigureStatus({ ownerId, message: 'Recalculating the Figure at its new size.' });
            setFileStatus('Recalculating the resized Figure from its Python source.');
            // The resize gesture is the explicit request to rebuild this Figure; ordinary Run remains consent-gated.
            void pythonExecution.run(sourceTarget, calculation.stale);
            return;
          }
          // Placement and geometry edits remain available after worker loss by retaining the last PNG.
          placedFromCachedImage = true;
          placement = { ...requested, renderedFigure: prior?.renderedFigure ?? source,
            renderedSourceId: prior?.renderedSourceId ?? await figureCacheKey(source) };
        } else {
          placement = { ...requested, renderedFigure: redraw.figure, renderedSourceId: await figureCacheKey(source) };
        }
      }
      const document = putFigurePlacement(basis.document, ownerId, placement);
      const isMove = !!prior;
      if (!session.tryExecute(basis, { kind: 'structure', document, label: isMove ? 'Move or resize Figure' : 'Place Figure' })) {
        const message = 'The workbook changed during Figure redraw. The previous placement remains.';
        setFigureStatus({ ownerId, message });
        setFileStatus(message);
        return;
      }
      setSnapshot(session.capture());
      workbookChanged(true);
      setFigureStatus(placedFromCachedImage ? { ownerId,
        message: 'Retained the cached Figure image. Rebuild Figure to redraw it at this size.' } : undefined);
      setPendingFigureOwner(undefined);
      setFileStatus(placedFromCachedImage
        ? isMove ? 'Figure placement geometry updated with its cached image. Rebuild Figure to redraw at this size.'
          : 'Figure placed from its cached image. Rebuild Figure to redraw at another size.'
        : isMove ? 'Figure placement updated.'
        : 'Figure placed on the worksheet.');
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Figure placement failed.';
      setFigureStatus({ ownerId, message });
      setFileStatus(message);
    }
  }

  useEffect(() => {
    const pending = pendingFigureResize;
    const execution = pythonExecutionState;
    if (!pending || !execution || execution.key !== pending.target.sheetId + ':' + pending.target.cellId) return;
    if (execution.running) {
      if (!pending.runningObserved) setPendingFigureResize({ ...pending, runningObserved: true });
      return;
    }
    if (!pending.runningObserved) return;
    if (execution.message === 'Python run completed.') {
      const current = session.capture();
      if (current.sessionId === pending.target.sessionId && current.revision <= pending.baseRevision) return;
      setPendingFigureResize(undefined);
      const hasFigure = current.sessionId === pending.target.sessionId && current.document.sheets.some(sourceSheet =>
        sourceSheet.cells.some(cell => cell.content.kind === 'python' && cell.content.block?.id === pending.ownerId &&
          cell.content.block.lastFigure));
      if (hasFigure) {
        void saveFigurePlacement(pending.ownerId, pending.requested, pending.original);
        return;
      }
      const message = 'Python completed without a Figure; the previous placement remains unchanged.';
      setFigureStatus({ ownerId: pending.ownerId, message });
      setFileStatus(message);
      return;
    }
    setPendingFigureResize(undefined);
    const detail = execution.error ?? execution.message;
    const message = 'Python could not recalculate the resized Figure; the previous placement remains unchanged.' +
      (detail ? ' ' + detail : '');
    setFigureStatus({ ownerId: pending.ownerId, message });
    setFileStatus(message);
  }, [pendingFigureResize, pythonExecutionState, session, snapshot.revision]);

  function deleteFigurePlacement(placementId: string) {
    if (!python.barrier() || !editing.commit()) return;
    const basis = session.capture();
    try {
      const document = removeFigurePlacement(basis.document, placementId);
      if (!session.tryExecute(basis, { kind: 'structure', document, label: 'Delete Figure placement' })) {
        setFileStatus('The workbook changed before the Figure placement could be deleted.');
        return;
      }
      setSnapshot(session.capture());
      workbookChanged(true);
      setFileStatus('Figure placement deleted.');
    } catch (error) { setFileStatus(error instanceof Error ? error.message : 'Figure placement could not be deleted.'); }
  }

  function rebuildFigure() {
    setFigureStatus(undefined);
    requestPythonRun();
  }

  function requestPythonRunAll() {
    if (!python.barrier()) return;
    const first = session.capture().document.sheets.flatMap(current => current.cells.map(cell => ({ current, cell })))
      .find(item => item.cell.content.kind === 'python');
    const target = first && pythonAnchorAt(session.capture(), first.current.id,
      { row: first.cell.row, column: first.cell.column });
    if (!target) return;
    if (trustedPythonSession.current === target.sessionId) pythonExecution.runAll();
    else setPythonTrustPrompt({ target, all: true });
  }

  function confirmPythonRun(trustSession: boolean) {
    const prompt = pythonTrustPrompt;
    const target = prompt?.target;
    setPythonTrustPrompt(undefined);
    if (!target || target.sessionId !== session.sessionId) return;
    if (trustSession) trustedPythonSession.current = target.sessionId;
    if (prompt?.all) pythonExecution.runAll();
    else void pythonExecution.run(target, calculation.stale);
  }

  async function copyPythonResult() {
    if (!pythonBlock?.hasResult) return;
    const text = pythonResultText(pythonBlock);
    if (text === undefined) return;
    try {
      await navigator.clipboard.writeText(text);
      setFileStatus('Python result copied to the clipboard.');
    } catch {
      const fallback = document.createElement('textarea');
      fallback.value = text;
      fallback.setAttribute('aria-hidden', 'true');
      fallback.style.position = 'fixed';
      fallback.style.opacity = '0';
      document.body.append(fallback);
      fallback.select();
      const copied = document.execCommand('copy');
      fallback.remove();
      setFileStatus(copied ? 'Python result copied to the clipboard.' : 'Clipboard access is unavailable in this browser context.');
    }
  }

  function insertPythonValues() {
    if (!python.barrier()) return;
    const context = python.context;
    const current = context && pythonAnchorAt(session.capture(), context.sheetId, context);
    const cell = current && session.active.sheets.find(item => item.id === current.sheetId)?.cells.find(item => item.id === current.cellId);
    const block = cell?.content.kind === 'python' ? cell.content.block : undefined;
    if (!current || !block?.hasResult) return;
    try {
      const command = pythonInsertCommand(session.active, { sheetId: current.sheetId, row: current.row, column: current.column }, block, makeId);
      if (!python.leave()) return;
      session.execute(command);
      setSnapshot(session.capture());
      workbookChanged(true);
      setPythonOpen(false);
      setPythonPick(undefined);
    } catch (error) { setFileStatus(error instanceof Error ? error.message : 'Python values could not be inserted.'); }
  }

  function hidePython() {
    if (!python.leave()) return;
    setPythonHiddenFor(pythonTargetKey);
    setPythonOpen(false);
    setPythonPick(undefined);
    scheduleFocusGrid();
  }

  function togglePython() {
    if (pythonOpen) hidePython();
    else {
      setPythonHiddenFor(undefined);
      setPythonOpen(true);
    }
  }

  function beginPythonResize(event: ReactPointerEvent<HTMLDivElement>) {
    if (event.button !== 0) return;
    event.preventDefault();
    event.currentTarget.setPointerCapture(event.pointerId);
    pythonResize.current = { id: event.pointerId, startX: event.clientX, startWidth: pythonWidth };
  }

  function movePythonResize(event: ReactPointerEvent<HTMLDivElement>) {
    const active = pythonResize.current;
    if (!active || active.id !== event.pointerId) return;
    setPythonWidth(clamp(active.startWidth - (event.clientX - active.startX), 280, 720));
  }

  function endPythonResize(event: ReactPointerEvent<HTMLDivElement>) {
    if (pythonResize.current?.id === event.pointerId) pythonResize.current = undefined;
  }

  function resizePythonKeys(event: KeyboardEvent<HTMLDivElement>) {
    const step = event.shiftKey ? 48 : 16;
    if (event.key === 'ArrowLeft') setPythonWidth(width => clamp(width + step, 280, 720));
    else if (event.key === 'ArrowRight') setPythonWidth(width => clamp(width - step, 280, 720));
    else if (event.key === 'Home') setPythonWidth(280);
    else if (event.key === 'End') setPythonWidth(720);
    else return;
    event.preventDefault();
  }

  function beginPivotResize(event: ReactPointerEvent<HTMLDivElement>) {
    if (event.button !== 0) return;
    event.preventDefault();
    event.currentTarget.setPointerCapture(event.pointerId);
    pivotResize.current = { id: event.pointerId, startX: event.clientX, startWidth: pivotPaneWidth };
  }

  function movePivotResize(event: ReactPointerEvent<HTMLDivElement>) {
    const active = pivotResize.current;
    if (!active || active.id !== event.pointerId) return;
    setPivotPaneWidth(clamp(active.startWidth - (event.clientX - active.startX), 240, 560));
  }

  function endPivotResize(event: ReactPointerEvent<HTMLDivElement>) {
    if (pivotResize.current?.id === event.pointerId) pivotResize.current = undefined;
  }

  function resizePivotKeys(event: KeyboardEvent<HTMLDivElement>) {
    const step = event.shiftKey ? 48 : 16;
    if (event.key === 'ArrowLeft') setPivotPaneWidth(width => clamp(width + step, 240, 560));
    else if (event.key === 'ArrowRight') setPivotPaneWidth(width => clamp(width - step, 240, 560));
    else if (event.key === 'Home') setPivotPaneWidth(240);
    else if (event.key === 'End') setPivotPaneWidth(560);
    else return;
    event.preventDefault();
  }

  function clipboardEvent(event: ClipboardEvent<HTMLDivElement>, kind: 'copy' | 'cut' | 'paste') {
    event.preventDefault();
    run(kind === 'paste' ? { kind, mode: 'all' } : { kind }, event.clipboardData);
  }

  function beginInline(text?: string) {
    if (!python.barrier()) return;
    setBrush(undefined);
    setBorderTool(undefined);
    editing.begin(activeTarget, 'inline', text);
  }

  function shortcut(event: KeyboardEvent<HTMLDivElement>): boolean {
    const key = event.key.toLowerCase();
    const control = event.ctrlKey || event.metaKey;
    if (key === 'contextmenu' || key === 'f10' && event.shiftKey) {
      event.preventDefault();
      const bounds = event.currentTarget.getBoundingClientRect();
      openGridMenu(bounds.left + 44, bounds.top + 26);
      return true;
    }
    if (event.nativeEvent.isComposing || key === 'process' || key === 'dead' || event.nativeEvent.keyCode === 229) {
      beginInline();
      return true;
    }
    if (key === 'escape') {
      editing.clearCut();
      return false;
    }
    if (event.altKey && key !== '=' && !event.getModifierState('AltGraph')) return false;
    let handle: (() => void) | undefined;
    if (painter && key === 'enter' && !control && !selection.additional.length) {
      handle = () => paint(primaryArea(selection));
    } else if (control && ['b', 'i', 'u'].includes(key)) {
      const style = resolvedCellStyle(sheet, selection.anchor);
      handle = () => runAppearance(key === 'b' ? { bold: style?.bold ? undefined : true } :
        key === 'i' ? { italic: style?.italic ? undefined : true } : underlinePatch(style?.underline ? 'None' : 'Single'));
    } else if (control && key === '1') {
      handle = () => openStructure('format-cells');
    } else if (control && (key === 'z' || key === 'y')) {
      handle = () => moveHistory(key === 'z' && !event.shiftKey);
    } else if (control && (key === 'd' || key === 'r')) {
      handle = () => run({ kind: 'fill', direction: key === 'd' ? 'down' : 'right' });
    } else if (control && event.shiftKey && key === 'v') {
      handle = () => run({ kind: 'paste', mode: 'values' });
    } else if (event.altKey && key === '=') {
      handle = () => run({ kind: 'autosum', aggregate: 'SUM' });
    } else if (!control && key === 'delete') {
      handle = () => {
        if (!editing.commit()) return;
        const basis = session.capture();
        const selectionRanges = ranges();
        try {
          const deletePivotCmd = deletePivotSelection(basis, selectionRanges);
          if (deletePivotCmd) {
            applyStructure(basis, deletePivotCmd);
            return;
          }
        } catch (err: unknown) {
          editing.reject(err instanceof Error ? err : new Error(String(err)));
          return;
        }
        const deleteTableCmd = deleteTableSelection(basis, selectionRanges);
        if (deleteTableCmd) {
          applyStructure(basis, deleteTableCmd);
        } else {
          run({ kind: 'clear', mode: 'contents' });
        }
      };
    } else if (!control && key === 'backspace') {
      handle = () => run({ kind: 'clear', mode: 'contents' });
    } else if (!control && key === 'f2') {
      handle = () => beginInline();
    } else if ((!control || event.getModifierState('AltGraph')) && [...event.key].length === 1 &&
        !(event.shiftKey && key === ' ')) {
      handle = () => beginInline(event.key);
    }
    if (!handle) return false;
    event.preventDefault();
    handle();
    return true;
  }

  const sheetCharts = useMemo(
    () => (snapshot.document.charts ?? [])
      .filter(chart => chart.placement.sheetId === sheet.id)
      .map(chart => ({ chart, data: resolveChartData(snapshot.document, chart) })),
    [snapshot.document, sheet.id],
  );
  const selectedChart = useMemo(
    () => (selectedChartId ? snapshot.document.charts?.find(c => c.id === selectedChartId && c.placement.sheetId === sheet.id) : undefined),
    [snapshot.document.charts, selectedChartId, sheet.id],
  );
  const chartHighlights = useMemo(
    () => (selectedChart && !editing.draft ? chartRangeHighlights(snapshot.document, selectedChart) : []),
    [selectedChart, editing.draft, snapshot.document],
  );
  const chartRangeHandle = useMemo(() => {
    if (!selectedChart || editing.draft || selectedChart.source.kind !== 'range') return undefined;
    const valueHighlights = chartHighlights.filter(h => h.sheetId === sheet.id && h.kind === 'values');
    if (valueHighlights.length === 0) return undefined;
    const combinedValuesArea: Area = {
      top: Math.min(...valueHighlights.map(h => h.area.top)),
      left: Math.min(...valueHighlights.map(h => h.area.left)),
      bottom: Math.max(...valueHighlights.map(h => h.area.bottom)),
      right: Math.max(...valueHighlights.map(h => h.area.right)),
    };
    return {
      area: combinedValuesArea,
      color: valueHighlights[0]!.color,
      onResizeArea: (nextArea: Area) => {
        const basis = session.capture();
        const current = basis.document.charts?.find(c => c.id === selectedChart.id);
        if (!current || current.source.kind !== 'range') return;
        const nextRowCount = nextArea.bottom - nextArea.top + 1;
        const nextColCount = nextArea.right - nextArea.left + 1;
        const catHighlight = chartHighlights.find(h => h.sheetId === sheet.id && h.kind === 'categories');
        const nextSeries = current.series.flatMap((s, idx) => {
          const vh = valueHighlights.find(h => h.seriesIndex === idx);
          if (!vh) return [s];
          if (current.seriesBy === 'columns') {
            if (valueHighlights.length > 1 && vh.area.left > nextArea.right) return [];
            const colCount = valueHighlights.length === 1 ? nextColCount : vh.area.right - vh.area.left + 1;
            return [{
              ...s,
              valuesRef: formatAbsoluteAreaRef(sheet.name, nextArea.top, vh.area.left, nextRowCount, colCount),
              ...(catHighlight && idx === 0
                ? { categoriesRef: formatAbsoluteAreaRef(sheet.name, nextArea.top, catHighlight.area.left, nextRowCount, catHighlight.area.right - catHighlight.area.left + 1) }
                : {}),
            }];
          }
          if (valueHighlights.length > 1 && vh.area.top > nextArea.bottom) return [];
          const rowCount = valueHighlights.length === 1 ? nextRowCount : vh.area.bottom - vh.area.top + 1;
          return [{
            ...s,
            valuesRef: formatAbsoluteAreaRef(sheet.name, vh.area.top, nextArea.left, rowCount, nextColCount),
            ...(catHighlight && idx === 0
              ? { categoriesRef: formatAbsoluteAreaRef(sheet.name, catHighlight.area.top, nextArea.left, catHighlight.area.bottom - catHighlight.area.top + 1, nextColCount) }
              : {}),
          }];
        });
        if (nextSeries.length > 0) {
          applyStructure(basis, updateChart(basis, current.id, { series: nextSeries }));
        }
      },
    };
  }, [selectedChart, editing.draft, chartHighlights, sheet.id, sheet.name, session]);
  const formulaHighlights = useMemo(
    () => (editing.draft && editing.draft.originalStyle?.numberFormat !== '@' && editing.draft.text.startsWith('=')
      ? formulaReferenceHighlights(editing.draft.text, snapshot.document, editing.draft.target.sheetId)
      : []),
    [editing.draft, snapshot.document],
  );
  const sheetFormulaHighlights = useMemo(
    () => (formulaHighlights.length > 0
      ? formulaHighlights.filter(item => item.sheetId === sheet.id)
      : chartHighlights.filter(item => item.sheetId === sheet.id)),
    [formulaHighlights, chartHighlights, sheet.id],
  );

  function editorProps(surface: EditSurface): CellEditorProps {
    const cell = editing.draft ? editing.draft.original : findCell(sheet, selection.anchor);
    return {
      surface,
      active: editing.draft?.surface === surface,
      value: editing.draft?.text ?? editing.editorText(cell),
      label: surface === 'inline' ? 'Edit ' + address(editing.draft!.target) : 'Formula bar',
      onStart: () => editing.begin(activeTarget, surface),
      onChange: editing.update,
      onCommit: (key, reverse) => {
        const target = editing.draft?.target ?? activeTarget;
        if (editing.commit()) finishEditor(target, key, reverse);
      },
      onBlur: () => editing.blur(surface),
      onCancel: () => {
        const target = editing.draft?.target ?? activeTarget;
        editing.cancel();
        if (target.sheetId !== sheet.id) finishEditor(target);
      },
      onComposition: editing.compose,
      isComposing: editing.composing,
      direct: surface === 'inline' && !!editing.draft?.direct,
      functionDraft: editing.draft?.surface === surface && editing.draft.functionDraft,
      ...(editing.draft?.surface === surface && editing.draft.selectText && { selectText: editing.draft.selectText }),
      separator,
      textFormat: (editing.draft ? editing.draft.originalStyle : resolvedCellStyle(sheet, selection.anchor, cell))?.numberFormat === '@',
      highlights: formulaHighlights,
      onKeepCaret: editing.keepCaret,
    };
  }

  const calculationMessage = pythonExecutionState?.running ? pythonExecutionState.message : calculation.error ??
    (calculation.pending ? calculation.pending === 'observe' ? 'Checking calculation status…' : 'Calculating…' :
      calculation.stale.size ? 'Calculation needed' : '');

  async function saveRecoveryBeforeUpdate(): Promise<boolean> {
    if (!session.isDirty) return true;
    if (!preferencesReady || pendingRecovery || workspaceConflict || !python.barrier() || !editing.commit()) return false;
    const captured = session.capture();
    const lease = workspaceLease.current;
    if (!lease || lease.activeId !== captured.document.id) throw new LocalWorkspaceError('The workbook recovery lease is not active.');
    const packageBlob = await saveNativePackage(captured.document);
    if (!session.isCurrent(captured) || !session.isDirty) return false;
    await localStore.writeRecovery({
      id: 'recovery:' + captured.document.id,
      workbookId: captured.document.id,
      name: activeFile.name,
      package: packageBlob,
      savedAt: Date.now(),
      revision: captured.authoredRevision,
    }, lease.owner);
    recoveryEpoch.current++;
    setRecoveryStatus('Local recovery copy updated before the app update.');
    return true;
  }

  const ribbonLatest = useRef({
    onFileBackstage: () => {},
    onAlign: runAlignment,
    onMerge: runMerge,
    onBorderAction: runBorder,
    onBorderStyle: (_style: BorderStyle) => {},
    onBorderColor: (_color?: string) => {},
    onBorderDialog: () => {},
    onNumberFormat: runNumberFormat,
    onAppearance: runAppearance,
    onAppearanceError: (_error: string, _focus: boolean) => {},
    onFormatDialog: (_tab: FormatTab) => {},
    onPainter: armPainter,
    onView,
    onEdit: run,
    onStructure: openStructure,
    onCalculationMode: setCalculationMode,
    onPython: togglePython,
    onCalculate: (_scope: 'sheet' | 'workbook') => {},
    onInsertFunction: (_name: NativeFunction) => {},
    onOpenInsertFunctionDialog: (_category?: InsertFunctionFilter) => {},
    onZoomToSelection: () => {},
    onZoomDialog: () => {},
    onFreeze: runFreeze,
    onTableRename: handleTableRename,
    onTableResize: () => {},
    onTableChangeDesign: handleTableChangeDesign,
    onTableClick: () => {},
    onPivotTableClick: () => {},
    onInsertChart: (_chartType: ChartType) => {},
    onRecommendedCharts: () => {},
    onAllCharts: (_category?: CoreChartCategoryTab) => {},
    onFind: () => {},
    onReplace: () => {},
    onNameManagerClick: () => {},
    onDefineNameClick: () => {},
    onSort: requestSort,
    onToggleFilter: toggleFilter,
    onClearFilter: clearFilter,
    onReapplyFilter: reapplyFilter,
  });
  ribbonLatest.current = {
    onFileBackstage: () => setBackstageOpen(true),
    onAlign: runAlignment,
    onMerge: runMerge,
    onBorderAction: runBorder,
    onBorderStyle: style => { setLineStyle(style); runBorder(borderTool?.mode ?? 'draw'); },
    onBorderColor: color => { setLineColor(color); runBorder(borderTool?.mode ?? 'draw'); },
    onBorderDialog: () => openStructure('format-cells', undefined, 'Border'),
    onNumberFormat: runNumberFormat,
    onAppearance: runAppearance,
    onAppearanceError: (error, focus) => { editing.reject(new Error(error)); if (focus) scheduleFocusGrid(); },
    onFormatDialog: tab => openStructure('format-cells', undefined, tab),
    onPainter: armPainter,
    onView,
    onEdit: run,
    onStructure: openStructure,
    onCalculationMode: setCalculationMode,
    onPython: togglePython,
    onCalculate: scope => {
      if (python.barrier() && editing.commit()) scheduler.current?.calculate(scope === 'sheet' ? { sheetId: sheet.id } : 'workbook');
    },
    onInsertFunction: name => {
      if (!editing.commit()) return;
      setPreferences(current => ({
        ...current,
        recentFunctions: recordRecentFunction(sanitizeRecentFunctions(current.recentFunctions), name),
      }));
      setFunctionArgsDialog({
        name,
        target: activeTarget,
        values: initialFunctionArguments(name).map(() => ''),
        activeIndex: 0,
      });
    },
    onOpenInsertFunctionDialog: (category = 'Most Recently Used') => {
      if (!editing.commit()) return;
      setInsertFunctionDialog({ category });
    },
    onZoomToSelection: () => {
      if (!editing.commit()) return;
      const zoom = worksheet.current?.zoomToSelection();
      if (zoom !== undefined) onView({ zoom });
      worksheet.current?.focus();
    },
    onZoomDialog: () => {
      if (!editing.commit()) return;
      setZoomDialog(true);
    },
    onFreeze: runFreeze,
    onTableRename: handleTableRename,
    onTableResize: () => { if (activeTable) openEditDialog(activeTable); },
    onTableChangeDesign: handleTableChangeDesign,
    onTableClick: () => openCreateDialog(true),
    onPivotTableClick: openCreatePivotDialog,
    onInsertChart: (chartType: ChartType) => insertChartWithType(chartType),
    onRecommendedCharts: () => openInsertChartDialog('recommended'),
    onAllCharts: (category?: CoreChartCategoryTab) => openInsertChartDialog('all', category),
    onFind: () => openSearch(false),
    onReplace: () => openSearch(true),
    onNameManagerClick: () => { if (editing.commit()) setNameManagerOpen(true); },
    onDefineNameClick: () => openCreateDialog(false),
    onSort: requestSort,
    onToggleFilter: toggleFilter,
    onClearFilter: clearFilter,
    onReapplyFilter: reapplyFilter,
  };
  const ribbonHandlers = useMemo(() => ({
    onFileBackstage: () => ribbonLatest.current.onFileBackstage(),
    onAlign: (action: AlignmentAction) => ribbonLatest.current.onAlign(action),
    onMerge: (mode: MergeMode) => ribbonLatest.current.onMerge(mode),
    onBorderAction: (action: BorderAction) => ribbonLatest.current.onBorderAction(action),
    onBorderStyle: (style: BorderStyle) => ribbonLatest.current.onBorderStyle(style),
    onBorderColor: (color?: string) => ribbonLatest.current.onBorderColor(color),
    onBorderDialog: () => ribbonLatest.current.onBorderDialog(),
    onNumberFormat: (change: NumberFormatChange) => ribbonLatest.current.onNumberFormat(change),
    onAppearance: (patch: AppearancePatch, focus?: boolean) => ribbonLatest.current.onAppearance(patch, focus),
    onAppearanceError: (error: string, focus: boolean) => ribbonLatest.current.onAppearanceError(error, focus),
    onFormatDialog: (tab: FormatTab) => ribbonLatest.current.onFormatDialog(tab),
    onPainter: (locked: boolean) => ribbonLatest.current.onPainter(locked),
    onView: (patch: Partial<SheetView>) => ribbonLatest.current.onView(patch),
    onEdit: (action: EditAction) => ribbonLatest.current.onEdit(action),
    onStructure: (action: StructureAction) => ribbonLatest.current.onStructure(action),
    onCalculationMode: (mode: CalculationMode) => ribbonLatest.current.onCalculationMode(mode),
    onPython: () => ribbonLatest.current.onPython(),
    onCalculate: (scope: 'sheet' | 'workbook') => ribbonLatest.current.onCalculate(scope),
    onInsertFunction: (name: NativeFunction) => ribbonLatest.current.onInsertFunction(name),
    onOpenInsertFunctionDialog: (category?: InsertFunctionFilter) => ribbonLatest.current.onOpenInsertFunctionDialog(category),
    onZoomToSelection: () => ribbonLatest.current.onZoomToSelection(),
    onZoomDialog: () => ribbonLatest.current.onZoomDialog(),
    onFreeze: (mode: 'panes' | 'topRow' | 'firstColumn' | 'none') => ribbonLatest.current.onFreeze(mode),
    onTableRename: (name: string) => ribbonLatest.current.onTableRename(name),
    onTableResize: () => ribbonLatest.current.onTableResize(),
    onTableChangeDesign: (patch: Partial<TableDesign>) => ribbonLatest.current.onTableChangeDesign(patch),
    onTableClick: () => ribbonLatest.current.onTableClick(),
    onPivotTableClick: () => ribbonLatest.current.onPivotTableClick(),
    onInsertChart: (chartType: ChartType) => ribbonLatest.current.onInsertChart(chartType),
    onRecommendedCharts: () => ribbonLatest.current.onRecommendedCharts(),
    onAllCharts: (category?: CoreChartCategoryTab) => ribbonLatest.current.onAllCharts(category),
    onFind: () => ribbonLatest.current.onFind(),
    onReplace: () => ribbonLatest.current.onReplace(),
    onNameManagerClick: () => ribbonLatest.current.onNameManagerClick(),
    onDefineNameClick: () => ribbonLatest.current.onDefineNameClick(),
    onSort: (ascending: boolean, dialog: boolean) => ribbonLatest.current.onSort(ascending, dialog),
    onToggleFilter: () => ribbonLatest.current.onToggleFilter(),
    onClearFilter: () => ribbonLatest.current.onClearFilter(),
    onReapplyFilter: () => ribbonLatest.current.onReapplyFilter(),
  }), []);
  const ribbonBorders = useMemo(() => ({
    action: borderAction, style: lineStyle, color: lineColor, mode: borderTool?.mode,
    onAction: ribbonHandlers.onBorderAction, onStyle: ribbonHandlers.onBorderStyle,
    onColor: ribbonHandlers.onBorderColor, onDialog: ribbonHandlers.onBorderDialog,
  }), [borderAction, lineStyle, lineColor, borderTool?.mode, ribbonHandlers]);
  const pivotLatest = useRef({
    onRename: handlePivotRename,
    onOpenOptions: () => setPivotOptionsOpen(true),
    onOpenFieldSettings: handleOpenPivotFieldSettings,
    onExpandField: () => handlePivotExpandCollapse(true),
    onCollapseField: () => handlePivotExpandCollapse(false),
    onRefresh: handlePivotRefresh,
    onRefreshAll: handlePivotRefreshAll,
    onOpenChangeSource: handleOpenChangePivotSource,
    onClear: handlePivotClear,
    onSelectEntire: handlePivotSelectEntire,
    onOpenMove: handleOpenMovePivot,
    onToggleFieldList: () => setPivotFieldListOpen(open => !open),
    onToggleButtons: () => { if (activePivot) handleUpdatePivotLayout(activePivot.id, { showButtons: activePivot.showButtons === false }); },
    onToggleFieldHeaders: () => { if (activePivot) handleUpdatePivotLayout(activePivot.id, { showFieldHeaders: activePivot.showFieldHeaders === false }); },
    onUpdateLayout: (patch: Partial<Omit<PivotTableDocument, 'id' | 'extent'>>) => {
      if (activePivot) handleUpdatePivotLayout(activePivot.id, patch);
    },
  });
  pivotLatest.current = {
    onRename: handlePivotRename,
    onOpenOptions: () => { if (activePivot) setPivotOptionsOpen(true); },
    onOpenFieldSettings: handleOpenPivotFieldSettings,
    onExpandField: () => handlePivotExpandCollapse(true),
    onCollapseField: () => handlePivotExpandCollapse(false),
    onRefresh: handlePivotRefresh,
    onRefreshAll: handlePivotRefreshAll,
    onOpenChangeSource: handleOpenChangePivotSource,
    onClear: handlePivotClear,
    onSelectEntire: handlePivotSelectEntire,
    onOpenMove: handleOpenMovePivot,
    onToggleFieldList: () => setPivotFieldListOpen(open => !open),
    onToggleButtons: () => { if (activePivot) handleUpdatePivotLayout(activePivot.id, { showButtons: activePivot.showButtons === false }); },
    onToggleFieldHeaders: () => { if (activePivot) handleUpdatePivotLayout(activePivot.id, { showFieldHeaders: activePivot.showFieldHeaders === false }); },
    onUpdateLayout: (patch: Partial<Omit<PivotTableDocument, 'id' | 'extent'>>) => {
      if (activePivot) handleUpdatePivotLayout(activePivot.id, patch);
    },
  };
  const pivotCallbacks = useMemo(() => ({
    onRename: (name: string) => pivotLatest.current.onRename(name),
    onOpenOptions: () => pivotLatest.current.onOpenOptions(),
    onOpenFieldSettings: () => pivotLatest.current.onOpenFieldSettings(),
    onExpandField: () => pivotLatest.current.onExpandField(),
    onCollapseField: () => pivotLatest.current.onCollapseField(),
    onRefresh: () => pivotLatest.current.onRefresh(),
    onRefreshAll: () => pivotLatest.current.onRefreshAll(),
    onOpenChangeSource: () => pivotLatest.current.onOpenChangeSource(),
    onClear: (mode: 'all' | 'filters') => pivotLatest.current.onClear(mode),
    onSelectEntire: () => pivotLatest.current.onSelectEntire(),
    onOpenMove: () => pivotLatest.current.onOpenMove(),
    onToggleFieldList: () => pivotLatest.current.onToggleFieldList(),
    onToggleButtons: () => pivotLatest.current.onToggleButtons(),
    onToggleFieldHeaders: () => pivotLatest.current.onToggleFieldHeaders(),
    onUpdateLayout: (patch: Partial<Omit<PivotTableDocument, 'id' | 'extent'>>) => pivotLatest.current.onUpdateLayout(patch),
  }), []);
  const pivotRibbonProps = useMemo(() => (
    activePivot
      ? {
        pivot: activePivot,
        activeField: activePivotField,
        fieldListOpen: pivotFieldListOpen,
        ...pivotCallbacks,
      }
      : undefined
  ), [activePivot, activePivotField, pivotFieldListOpen, pivotCallbacks]);
  const activeCell = findCell(sheet, selection.anchor);
  const activeStyle = resolvedCellStyle(sheet, selection.anchor, activeCell);
  const activeMerged = !!mergeAt(sheet, selection.anchor.row, selection.anchor.column);
  const painterMode = painter ? painter.locked ? 'locked' : 'single' : undefined;
  const sheetFrozen = sheet.frozenPanes.rows > 0 || sheet.frozenPanes.columns > 0;

  return (
    <div className="app-shell" ref={root} onKeyDownCapture={event => {
      if ((event.ctrlKey || event.metaKey) && !event.altKey && !(event.target instanceof Element && event.target.closest('dialog[open]'))) {
        const key = event.key.toLowerCase();
        const action: FileAction | undefined = key === 'n' && !event.shiftKey ? 'new' :
          key === 'o' && !event.shiftKey ? 'open' : key === 's' && event.shiftKey ? 'saveAs' :
          key === 's' ? 'save' : key === 'w' && !event.shiftKey ? 'close' : undefined;
        if (action) {
          event.preventDefault();
          event.stopPropagation();
          void runFileAction(action);
          return;
        }
      }
      if (event.key === 'Escape' && (painter || borderTool) && !(event.target instanceof Element && event.target.closest('dialog, [popover]:popover-open, [role="menu"]'))) {
        event.preventDefault();
        event.stopPropagation();
        setBrush(undefined);
        setBorderTool(undefined);
        editing.clearStatus();
      }
    }}>
      <OfflineStatus dirty={session.isDirty} onSaveRecovery={saveRecoveryBeforeUpdate} />
      <header className="title-bar" inert={backstageOpen || !!workspaceConflict}>
        <span className="brand-mark" aria-hidden="true"><svg viewBox="0 0 20 20" width="16" height="16" fill="none"
          stroke="currentColor" strokeWidth="1.4"><rect x="3" y="3" width="14" height="14" rx="1" /><path d="M3 8h14M3 12h14M8 3v14" /></svg></span>
        <div className="quick-access" role="group" aria-label="Quick Access">
          <button type="button" className="icon-button" aria-label="Save" title="Save (Ctrl+S)"
            disabled={!preferencesReady || fileBusy || !!workspaceConflict} onClick={() => { void runFileAction('save'); }}>
            <ShellIcon icon="save" />
          </button>
          <button type="button" className="icon-button" aria-label="Undo" title={session.canUndo ? 'Undo' : 'Nothing to undo'}
            disabled={!session.canUndo} onClick={() => moveHistory(true)}><ShellIcon icon="undo" /></button>
          <button type="button" className="icon-button" aria-label="Redo" title={session.canRedo ? 'Redo' : 'Nothing to redo'}
            disabled={!session.canRedo} onClick={() => moveHistory(false)}><ShellIcon icon="redo" /></button>
        </div>
        <span className="workbook-title">{activeFile.name}{session.isDirty ? ' *' : ''}</span>
        <span className="product-title">ExcelLight Web</span>
      </header>

      <Ribbon inert={backstageOpen || !!workspaceConflict} fileBackstageOpen={backstageOpen} onFileBackstage={ribbonHandlers.onFileBackstage}
        onAlign={ribbonHandlers.onAlign} onMerge={ribbonHandlers.onMerge} merged={activeMerged} borders={ribbonBorders}
        cell={activeCell} style={activeStyle} onNumberFormat={ribbonHandlers.onNumberFormat} onAppearance={ribbonHandlers.onAppearance}
        onAppearanceError={ribbonHandlers.onAppearanceError} onFormatDialog={ribbonHandlers.onFormatDialog}
        painter={painterMode} onPainter={ribbonHandlers.onPainter} view={view} onView={ribbonHandlers.onView}
        onEdit={ribbonHandlers.onEdit} onStructure={ribbonHandlers.onStructure}
        calculationMode={calculation.mode} onCalculationMode={ribbonHandlers.onCalculationMode} pythonOpen={pythonOpen}
        onPython={ribbonHandlers.onPython} onCalculate={ribbonHandlers.onCalculate}
        recentFunctions={recentFunctions} onInsertFunction={ribbonHandlers.onInsertFunction}
        onOpenInsertFunctionDialog={ribbonHandlers.onOpenInsertFunctionDialog}
        onZoomToSelection={ribbonHandlers.onZoomToSelection} onZoomDialog={ribbonHandlers.onZoomDialog}
        frozen={sheetFrozen} onFreeze={ribbonHandlers.onFreeze} activeTable={activeTable}
        onTableRename={ribbonHandlers.onTableRename} onTableResize={ribbonHandlers.onTableResize}
        onTableChangeDesign={ribbonHandlers.onTableChangeDesign} onTableClick={ribbonHandlers.onTableClick}
        onPivotTableClick={ribbonHandlers.onPivotTableClick}
        onInsertChart={ribbonHandlers.onInsertChart}
        onRecommendedCharts={ribbonHandlers.onRecommendedCharts}
        onAllCharts={ribbonHandlers.onAllCharts}
        {...(pivotRibbonProps && { pivotRibbon: pivotRibbonProps })}
        onFind={ribbonHandlers.onFind} onReplace={ribbonHandlers.onReplace}
        onNameManagerClick={ribbonHandlers.onNameManagerClick} onDefineNameClick={ribbonHandlers.onDefineNameClick}
        filter={filterCommands} onSort={ribbonHandlers.onSort} onToggleFilter={ribbonHandlers.onToggleFilter}
        onClearFilter={ribbonHandlers.onClearFilter} onReapplyFilter={ribbonHandlers.onReapplyFilter}
      />

      <main className="workbook" aria-label="Workbook" inert={backstageOpen || !!workspaceConflict}>
        {(view.formulaBar || editing.draft?.surface === 'formula') && <div className="formula-bar">
          <NameBox selection={selection} goTo={goTo} cancel={() => worksheet.current?.focus()} />
          <span className="formula-symbol" aria-hidden="true">ƒx</span>
          <CellEditor {...editorProps('formula')} />
        </div>}
        <div className="worksheet-row">
          <Worksheet key={sheet.id} ref={handle => {
            worksheet.current = handle;
            if (handle && requestedFocus.current) {
              handle.focus();
              requestedFocus.current = false;
            }
          }} sheet={sheet} selection={selection} onSelection={setSelection} onAdjacentSheet={chooseAdjacent}
            paintActive={!!painter} onPaint={paint}
            borderTool={borderTool?.mode} borderPen={borderPen(lineStyle, lineColor)} onBeginBorder={beginBorderStroke}
            onExitBorder={() => setBorderTool(undefined)}
            referenceArea={pythonPick?.sheetId === sheet.id ? pythonPick.area :
              editing.draft?.pick?.sheetId === sheet.id ? editing.draft.pick.area : undefined}
            formulaHighlights={sheetFormulaHighlights}
            onReferencePick={() => {
              if (python.wantsReference()) {
                let area: Area | undefined;
                return {
                  update: next => { area = next; setPythonPick({ sheetId: sheet.id, area: next }); },
                  finish: () => {
                    if (area && python.insertReference(sheet, area)) {
                      const anchor = python.context;
                      if (anchor && anchor.sheetId !== sheet.id) {
                        setActiveSheetId(anchor.sheetId);
                        setSelections(current => new Map(current).set(anchor.sheetId,
                          singleSelection({ row: anchor.row, column: anchor.column })));
                      }
                    }
                    setPythonPick(undefined);
                  },
                  cancel: () => setPythonPick(undefined),
                };
              }
              const caret = editorSelection();
              return caret && editing.beginReferencePick(caret, sheet);
            }}
            onEdit={beginInline} onShortcut={shortcut} onClipboard={clipboardEvent}
            onStructureMenu={openGridMenu}
            figures={worksheetFigures}
            {...(pendingFigure && { pendingFigure })}
            {...(figureStatus && { figureStatus })}
            onFigurePlacement={(ownerId, placement, original) => { void saveFigurePlacement(ownerId, placement, original); }}
            onFigureDelete={deleteFigurePlacement}
            onFigureCancel={() => setPendingFigureOwner(undefined)}
            charts={sheetCharts}
            {...(selectedChartId !== undefined && { selectedChartId })}
            {...(chartRangeHandle !== undefined && { chartRangeHandle })}
            onSelectChart={setSelectedChartId}
            onChartPlacement={handleChartPlacement}
            onChartDelete={handleChartDelete}
            onHeaderFit={(axis, index) => openStructure(`fit-${axis}`, targetHeader(axis, index))}
            onHeaderResize={(axis, index) => {
              if (!python.barrier() || !editing.commit()) return;
              const captured = session.capture();
              const indexes = indexesFor(targetHeader(axis, index), axis);
              return size => {
                try {
                  const basis = currentAuthoredBasis(captured);
                  applyStructure(basis, changeDimensions(basis, sheet.id, axis, indexes.map(index => ({ index, size }))));
                }
                catch (error) { editing.reject(error); }
              };
            }}
            editor={editing.draft?.surface === 'inline' ? { at: editing.draft.target, input: <CellEditor {...editorProps('inline')} /> } : undefined}
            zoom={view.zoom} headings={view.headings} gridlines={view.gridlines} showFormulas={view.showFormulas}
            staleCells={calculation.stale} calculatingCells={calculatingCells} separator={separator}
            tables={tables}
            rangePicking={rangePicker !== undefined}
            onRangePickFinish={accept => rangePicker?.onDone(accept ? rangePicker.initial : null)}
            onResizeTable={handleTableResize}
            onTotalFunction={handleTotalFunction}
            onFilter={openFilter}
            pivotOverlays={pivotOverlays}
            onPivotToggle={handlePivotToggle}
            onPivotFilter={(pivotId, kind, fieldNames, anchor) => {
              if (!editing.commit()) return;
              setFilterMenu(undefined);
              setPivotFilterMenu({ pivotId, kind, fieldNames, anchor });
            }}
            onPivotDrillDown={handlePivotDrillDown}
          />
          {pythonOpen && <>
            <div className="python-divider" role="separator" aria-label="Resize Python workspace"
              aria-orientation="vertical" aria-valuemin={280} aria-valuemax={720} aria-valuenow={pythonWidth}
              tabIndex={0} onPointerDown={beginPythonResize} onPointerMove={movePythonResize}
              onPointerUp={endPythonResize} onPointerCancel={endPythonResize} onKeyDown={resizePythonKeys} />
            <PythonPanel width={pythonWidth} context={python.context} insertRequest={python.insertRequest}
              status={python.status} pythonInputFresh={pythonInputFresh} {...(pythonBlock && { block: pythonBlock })}
              {...(pythonExecutionState && { execution: pythonExecutionState })}
              {...(pythonResultExecution && { resultExecution: pythonResultExecution })}
              placingFigure={!!pythonBlock && pendingFigureOwner === pythonBlock.id}
              {...(panelFigureStatus && { figureStatus: panelFigureStatus })}
              onDraft={python.receiveDraft} onPickable={python.setPickable}
              automatic={preferences.pythonCalculation === 'automatic'} onHide={hidePython} onRun={requestPythonRun}
              onRunAll={requestPythonRunAll} onStop={pythonExecution.stop}
              onCopy={() => { void copyPythonResult(); }} onInsertValues={insertPythonValues}
              onPlaceFigure={requestFigurePlacement} onSaveImage={() => { void saveFigureImage(); }}
              onRebuildFigure={rebuildFigure} />
          </>}
          {pivotOpen && activePivot && <>
            <div className="python-divider pivot-divider" role="separator" aria-label="Resize PivotTable Fields pane"
              aria-orientation="vertical" aria-valuemin={240} aria-valuemax={560} aria-valuenow={pivotPaneWidth}
              tabIndex={0} onPointerDown={beginPivotResize} onPointerMove={movePivotResize}
              onPointerUp={endPivotResize} onPointerCancel={endPivotResize} onKeyDown={resizePivotKeys} />
            <PivotTableFieldsPane
              width={pivotPaneWidth}
              document={snapshot.document}
              pivot={activePivot}
              makeId={makeId}
              onUpdateLayout={patch => handleUpdatePivotLayout(activePivot.id, patch)}
              onClose={() => { setPivotFieldListOpen(false); scheduleFocusGrid(); }}
              onOpenSettings={request => setFieldSettingsModal({ pivotId: activePivot.id, request })}
            />
          </>}
        </div>
        <nav className="sheet-strip" aria-label="Worksheets">
          <div className="sheet-navigation">
            <button type="button" className="icon-button" aria-label="Previous sheet" disabled={sheetIndex === 0}
              onMouseDown={keepEditorOnSheetClick} onClick={() => chooseAdjacent(-1)}><ShellIcon icon="previous" /></button>
            <button type="button" className="icon-button" aria-label="Next sheet" disabled={sheetIndex === sheets.length - 1}
              onMouseDown={keepEditorOnSheetClick} onClick={() => chooseAdjacent(1)}><ShellIcon icon="next" /></button>
          </div>
          <div className="sheet-tabs">
            {sheets.map(item => rename?.id === item.id
              ? <input key={item.id} className="sheet-rename" aria-label="Sheet name" defaultValue={item.name} maxLength={31}
                autoFocus onFocus={event => event.currentTarget.select()} onBlur={event => finishRename(event.currentTarget.value)}
                onKeyDown={event => {
                  if (event.nativeEvent.isComposing) return;
                  if (event.key === 'Enter' || event.key === 'Escape') {
                    event.preventDefault();
                    finishRename(event.key === 'Enter' ? event.currentTarget.value : undefined);
                  }
                }} />
              : <button type="button" key={item.id} aria-current={item.id === sheet.id ? 'true' : undefined}
                onMouseDown={keepEditorOnSheetClick} onClick={() => chooseSheet(item.id)}
                onDoubleClick={() => openStructure('rename-sheet', { id: item.id, ranges: [] })}
                onContextMenu={event => {
                  event.preventDefault();
                  if (editing.commit()) setMenu({ x: event.clientX, y: event.clientY, target: { id: item.id, ranges: [] },
                    actions: ['insert-sheet', 'delete-sheet', 'rename-sheet', 'place-sheet'] });
                }}
                onKeyDown={event => {
                  if (event.key === 'ContextMenu' || event.key === 'F10' && event.shiftKey) {
                    event.preventDefault();
                    const bounds = event.currentTarget.getBoundingClientRect();
                    if (editing.commit()) setMenu({ x: bounds.left, y: bounds.bottom, target: { id: item.id, ranges: [] },
                      actions: ['insert-sheet', 'delete-sheet', 'rename-sheet', 'place-sheet'] });
                  }
                }}>{item.name}</button>)}
            <button type="button" aria-label="New sheet" disabled={sheets.length >= 10} onClick={() => {
              if (!editing.commit()) return;
              const basis = session.capture();
              try {
                const next = changeSheet(basis, sheet.id, { kind: 'add', after: true }, makeId);
                applyStructure(basis, next.command, next.active);
              } catch (error) { editing.reject(error); }
            }}>+</button>
          </div>
        </nav>
      </main>

      <footer className="status-bar" aria-label="Worksheet status" inert={backstageOpen || !!workspaceConflict}>
        <span className={editing.status || fileStatus || recoveryStatus || calculationMessage ? 'edit-status' : 'ready'} role="status"
          data-calculating={pythonExecutionState?.running || calculation.pending === 'calculate' ? 'true' : undefined}>
          {[editing.status, fileStatus, recoveryStatus, painter && (painter.locked ? 'Format Painter locked' : 'Format Painter active') +
            ' · Select a destination or press Enter · Esc cancels', borderTool && borderModes[borderTool.mode] + ' · Esc exits', calculationMessage].filter(Boolean).join(' · ') || 'Ready'}
        </span>
        {!editing.status && !fileStatus && !recoveryStatus && !calculationMessage && <span className="grid-keyboard-hint" aria-hidden="true">F2 to edit · Esc, then Tab to leave</span>}
        <output className="selection-stats" aria-label="Selection statistics" aria-live="off">
          <span>Sum: {statisticsFormat.format(stats.sum)}</span>
          <span>Average: {stats.average === undefined ? '—' : statisticsFormat.format(stats.average)}</span>
          <span>Count: {stats.count}</span>
        </output>
        <div className="zoom-controls" role="group" aria-label="Sheet zoom">
          <button type="button" className="icon-button" aria-label="Zoom out" disabled={view.zoom <= 10}
            onClick={() => onView({ zoom: clamp(view.zoom - 10, 10, 400) })}>−</button>
          <input type="range" min="10" max="400" value={view.zoom} aria-label="Zoom"
            aria-valuetext={view.zoom + '%'} onChange={event => onView({ zoom: event.currentTarget.valueAsNumber })} />
          <button type="button" className="icon-button" aria-label="Zoom in" disabled={view.zoom >= 400}
            onClick={() => onView({ zoom: clamp(view.zoom + 10, 10, 400) })}>+</button>
          <button type="button" className="zoom-value" aria-label={'Zoom percentage, ' + view.zoom + '%'} title="Zoom"
            onClick={() => { if (editing.commit()) setZoomDialog(true); }}>{view.zoom}%</button>
        </div>
      </footer>
      <FileBackstage open={backstageOpen} busy={fileBusy || !preferencesReady || !!workspaceConflict}
        recent={recentWorkbooks} activeWorkspaceId={session.active.id} onBack={() => setBackstageOpen(false)}
        onAction={action => { void runFileAction(action); }}
        onRecent={record => { void runFileOperation('Open', () => openRecentWorkbook(record)); }}
        onRemoveRecent={record => setWorkspaceToRemove(record)}
        {...(storageError ? { storageError } : {})} />
      {pendingRecovery && <RecoveryDialog recovery={pendingRecovery} onChoice={choice => { void chooseRecovery(choice); }} />}
      {replacementPrompt && <UnsavedChangesDialog name={activeFile.name} onChoice={chooseReplacement} />}
      {optionsOpen && <OptionsDialog preferences={preferences} onApply={finishOptions} onCancel={() => finishOptions(null)} />}
      {saveAsOpen && <SaveAsDialog initialName={activeFile.name} purpose={saveNamePurpose}
        onSave={finishSaveAsName} onCancel={() => finishSaveAsName(null)} />}
      {workspaceToRemove && <RemoveWorkspaceDialog name={workspaceToRemove.name}
        onChoice={remove => { void decideWorkspaceRemoval(workspaceToRemove, remove); }} />}
      {pythonTrustPrompt && <PythonTrustDialog {...(pythonTrustPrompt.all !== undefined && { all: pythonTrustPrompt.all })}
        onRunOnce={() => confirmPythonRun(false)}
        onTrustSession={() => confirmPythonRun(true)} onCancel={() => {
          if (pendingFigureResize) {
            setPendingFigureResize(undefined);
            setFigureStatus({ ownerId: pendingFigureResize.ownerId,
              message: 'Resize cancelled; the existing Figure placement was kept.' });
            setFileStatus('Figure resize cancelled. No Python code was run.');
          }
          setPythonTrustPrompt(undefined);
        }} />}
      {fileMessage && <FileMessageDialog title={fileMessage.title} message={fileMessage.message} onClose={() => setFileMessage(undefined)} />}
      {xlsxReport && <XlsxReportDialog title={xlsxReport.title} issues={xlsxReport.issues}
        onChoice={finishXlsxReport} />}
      {workspaceConflict && <div className="workspace-conflict-cover" role="alertdialog" aria-modal="true" aria-labelledby="workspace-conflict-title">
        <section>
          <h2 id="workspace-conflict-title">Workbook open in another tab</h2>
          <p>This browser-local workbook is being edited in another tab. Editing here is paused to protect its saved and recovery copies.</p>
          <button type="button" className="file-primary" onClick={() => window.location.reload()}>Reload workbook</button>
        </section>
      </div>}
      {editing.manualPaste && <PasteTextDialog mode={editing.manualPaste.mode} error={editing.status}
        onSubmit={editing.submitManualPaste} onClose={editing.closeManualPaste} />}
      {menu && <StructureMenu {...menu} onAction={action => openStructure(action, menu.target)}
        onClose={() => { setMenu(undefined); scheduleFocusGrid(); }} />}
      {formatDraft && (() => {
        const draftSheet = formatDraft.basis.document.sheets.find(item => item.id === formatDraft.cell.sheetId);
        const draftCell = findReferencedCell(snapshot.document, formatDraft.cell);
        const draftStyle = draftSheet ? resolvedCellStyle(draftSheet, formatDraft.cell, draftCell) : draftCell?.style;
        return <FormatCellsDialog cell={draftCell} style={draftStyle} locale={locale}
          usedCodes={usedNumberCodes} initialTab={formatDraft.tab} borderEditor={borderEditor}
          mergeEnabled={formatDraft.ranges.length === 1} merged={formatDraft.ranges.length === 1 ?
            mergeState(draftSheet!, formatDraft.ranges[0]!) : undefined}
        onCancel={() => { setFormatDraft(undefined); scheduleFocusGrid(); }}
        onSubmit={(patch, code, borders, alignment, merge) => {
          try {
            if (merge !== undefined) {
              requestMerge(formatDraft.basis, formatDraft.ranges[0]!, merge ? 'cells' : 'unmerge', { patch, code, borders, alignment });
              return;
            }
            if (code !== undefined || Object.keys(patch).length || Object.keys(borders).length || Object.keys(alignment).length) {
              const basis = currentAuthoredBasis(formatDraft.basis);
              applyStyleCommand(formatStyles(basis, formatDraft.ranges, patch, code, borders, alignment));
            }
            setFormatDraft(undefined);
            scheduleFocusGrid();
          } catch (error) { return error instanceof Error ? error.message : 'The number format could not be applied.'; }
        }} />;
      })()}
      {mergeConfirmation && <MergeConfirmation onConfirm={mergeConfirmation.confirm} onCancel={() => {
        setMergeConfirmation(undefined);
        if (!formatDraft) scheduleFocusGrid();
      }} />}
      {structure && <StructureDialog action={structure.action} sheets={structure.basis.document.sheets}
        target={structure.basis.document.sheets.find(item => item.id === structure.target.id)!}
        initialSize={(() => {
          const axis = structure.action === 'size-row' ? 'row' : 'column';
          const item = structure.basis.document.sheets.find(item => item.id === structure.target.id)!;
          const index = structure.target.header ?? structure.target.ranges[0]?.[axis];
          return item[axis === 'row' ? 'rows' : 'columns'].find(dimension => dimension.index === index)?.size ?? gridMetrics[axis];
        })()}
        onCancel={() => { setStructure(undefined); scheduleFocusGrid(); }}
        onSubmit={input => {
          try { performStructure(structure.action, structure.target, structure.basis, input); setStructure(undefined); }
          catch (error) { return error instanceof Error ? error.message : 'The structural edit failed.'; }
        }} />}
      {definitionDraft && <NameDefinitionDialog definition={definitionDraft}
        suspended={rangePicker !== undefined}
        onPickRange={handlePickRangeFromDefinition}
        onSubmit={handleDefinitionSubmit}
        onCancel={() => { setDefinitionDraft(undefined); scheduleFocusGrid(); }} />}
      {nameManagerOpen && <NameManagerDialog document={snapshot.document}
        suspended={rangePicker !== undefined || definitionDraft !== undefined}
        pickedExtent={pickedManagerExtent}
        onConsumePick={() => setPickedManagerExtent(undefined)}
        onNew={() => openCreateDialog(false)}
        onEdit={entry => openEditDialog(entry, true)}
        onDelete={handleDeleteName}
        onPickRange={handlePickRangeFromManager}
        onApplyExtent={handleApplyManagerExtent}
        onClose={() => { setNameManagerOpen(false); setPickedManagerExtent(undefined); scheduleFocusGrid(); }} />}
      {insertChartDraft && <InsertChartDialog
        document={snapshot.document}
        sheetId={sheet.id}
        sourceFormula={insertChartDraft.sourceFormula}
        initialTab={insertChartDraft.tab}
        {...(insertChartDraft.category && { initialCategory: insertChartDraft.category })}
        onSubmit={(chartType, seriesBy) => insertChartWithType(chartType, insertChartDraft.sourceFormula, seriesBy)}
        onCancel={() => { setInsertChartDraft(undefined); scheduleFocusGrid(); }}
      />}
      {createPivotDraft && <CreatePivotTableDialog
        sourceText={createPivotDraft.sourceText}
        placement={createPivotDraft.placement}
        locationText={createPivotDraft.locationText}
        suspended={rangePicker !== undefined}
        onChangeSource={sourceText => setCreatePivotDraft(current => current ? { ...current, sourceText } : undefined)}
        onChangePlacement={placement => setCreatePivotDraft(current => current ? { ...current, placement } : undefined)}
        onChangeLocation={locationText => setCreatePivotDraft(current => current ? { ...current, locationText } : undefined)}
        onPickSource={handlePickPivotSource}
        onPickLocation={handlePickPivotLocation}
        onSubmit={handleCreatePivotSubmit}
        onCancel={() => { setCreatePivotDraft(undefined); scheduleFocusGrid(); }}
      />}
      {fieldSettingsModal && settingsPivot && <FieldSettingsDialog
        key={`${fieldSettingsModal.pivotId}:${fieldSettingsModal.request.mode}:${fieldSettingsModal.request.mode === 'value' ? fieldSettingsModal.request.field.id : `${fieldSettingsModal.request.area}:${fieldSettingsModal.request.field.fieldName}`}`}
        document={snapshot.document}
        pivot={settingsPivot}
        request={fieldSettingsModal.request}
        locale={locale}
        usedCodes={usedNumberCodes}
        onSaveValueField={handleSaveValueFieldSettings}
        onSaveAxisField={handleSaveAxisFieldSettings}
        onCancel={() => { setFieldSettingsModal(undefined); scheduleFocusGrid(); }}
      />}
      {optionsPivot && <PivotOptionsDialog
        pivot={optionsPivot}
        onSave={patch => {
          handleUpdatePivotLayout(optionsPivot.id, patch);
          setPivotOptionsOpen(false);
          scheduleFocusGrid();
        }}
        onCancel={() => { setPivotOptionsOpen(false); scheduleFocusGrid(); }}
      />}
      {changePivotSourceDraft && <ChangePivotSourceDialog
        sourceText={changePivotSourceDraft.sourceText}
        suspended={rangePicker !== undefined}
        onChangeSource={sourceText => setChangePivotSourceDraft(current => current ? { ...current, sourceText } : undefined)}
        onPickSource={handlePickChangePivotSource}
        onSubmit={handleChangePivotSourceSubmit}
        onCancel={() => { setChangePivotSourceDraft(undefined); scheduleFocusGrid(); }}
      />}
      {movePivotDraft && <MovePivotTableDialog
        placement={movePivotDraft.placement}
        locationText={movePivotDraft.locationText}
        suspended={rangePicker !== undefined}
        onChangePlacement={placement => setMovePivotDraft(current => current ? { ...current, placement } : undefined)}
        onChangeLocation={locationText => setMovePivotDraft(current => current ? { ...current, locationText } : undefined)}
        onPickLocation={handlePickMovePivotLocation}
        onSubmit={handleMovePivotSubmit}
        onCancel={() => { setMovePivotDraft(undefined); scheduleFocusGrid(); }}
      />}
      {rangePicker && <RangePickerWidget initialRange={rangePicker.initial}
        onConfirm={range => rangePicker.onDone(range)}
        onCancel={() => rangePicker.onDone(null)} />}
      {sortDialog && <SortDialog sheet={sortDialog.basis.document.sheets.find(item => item.id === sheet.id)!}
        range={sortDialog.range} draft={sortDialog.draft} onSubmit={submitSort} onOptions={openSortOptions}
        onCancel={() => { setSortDialog(undefined); scheduleFocusGrid(); }} />}
      {sortOptions && <SortOptionsDialog caseSensitive={sortOptions.draft.caseSensitive}
        onSubmit={finishSortOptions} onCancel={cancelSortOptions} />}
      {sortWarning && <SortWarningDialog selection={sortWarning.selection} region={sortWarning.region}
        onChoose={expand => sortWarning.resume(expand)} onCancel={() => sortWarning.resume(null)} />}
      {searchDialog && <FindReplaceDialog replace={searchDialog.replace} history={searchHistory}
        initialLookIn={searchLook} onLookIn={setSearchLook}
        onFindAll={runFindAll} onFindNext={runFindNext} onReplace={runReplace} onSelect={showSearchMatch}
        onClose={() => { setSearchDialog(undefined); scheduleFocusGrid(); }} />}
      {zoomDialog && <ZoomDialog current={view.zoom} onApply={applyZoom}
        onCancel={() => { setZoomDialog(false); scheduleFocusGrid(); }} />}
      {insertFunctionDialog && <InsertFunctionDialog
        initialCategory={insertFunctionDialog.category}
        recentFunctions={recentFunctions}
        separator={separator}
        onInsert={name => {
          setInsertFunctionDialog(undefined);
          ribbonHandlers.onInsertFunction(name);
        }}
        onCancel={() => { setInsertFunctionDialog(undefined); scheduleFocusGrid(); }}
      />}
      {functionArgsDialog && <FunctionArgumentsDialog
        functionName={functionArgsDialog.name}
        document={snapshot.document}
        target={functionArgsDialog.target}
        separator={separator}
        locale={locale}
        suspended={rangePicker !== undefined}
        values={functionArgsDialog.values}
        activeIndex={functionArgsDialog.activeIndex}
        onChangeValues={next => setFunctionArgsDialog(current => current ? { ...current, values: next } : undefined)}
        onChangeActiveIndex={index => setFunctionArgsDialog(current => current ? { ...current, activeIndex: index } : undefined)}
        onPickRange={handlePickRangeFromFunctionArgs}
        onConfirm={handleConfirmFunctionArgs}
        onCancel={() => { setFunctionArgsDialog(undefined); scheduleFocusGrid(); }}
      />}
      {filterMenu && <FilterDropdown anchor={filterMenu.anchor} header={filterMenu.header}
        textColumn={filterMenu.textColumn} values={filterMenu.values} selected={filterMenu.selected}
        hasColumnFilter={filterMenu.hasColumnFilter} onApply={applyFilter} onClearColumn={clearFilterColumn}
        onSort={sortFilterColumn}
        onCancel={() => { setFilterMenu(undefined); scheduleFocusGrid(); }} />}
      {pivotFilterMenu && menuPivot && <PivotFilterDropdown
        key={`${pivotFilterMenu.pivotId}-${pivotFilterMenu.kind}-${pivotFilterMenu.fieldNames.join('\u0000')}`}
        anchor={pivotFilterMenu.anchor}
        kind={pivotFilterMenu.kind}
        fieldNames={pivotFilterMenu.fieldNames}
        pivot={menuPivot}
        document={snapshot.document}
        onApplyFilter={(fieldName, hidden) => handlePivotFilterApply(menuPivot.id, pivotFilterMenu.kind, fieldName, hidden)}
        onSort={(fieldName, direction) => {
          if (pivotFilterMenu.kind !== 'reportFilter') {
            handlePivotFilterSort(menuPivot.id, pivotFilterMenu.kind, fieldName, direction);
          }
        }}
        onMoreSortOptions={(fieldName) => {
          if (pivotFilterMenu.kind !== 'reportFilter') {
            handlePivotMoreSortOptions(menuPivot.id, pivotFilterMenu.kind, fieldName);
          }
        }}
        onCancel={() => { setPivotFilterMenu(undefined); scheduleFocusGrid(); }}
      />}
    </div>
  );
}
