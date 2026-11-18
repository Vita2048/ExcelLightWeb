import { useEffect, useMemo, useRef, useState, type ClipboardEvent, type FormEvent, type KeyboardEvent, type MouseEvent } from 'react';
import type { EditRange } from '../application/editing-commands';
import { applyNumberFormat, type NumberFormatChange } from '../application/number-format-command';
import { FormatCellsDialog, type FormatTab } from '../features/formatting/FormatCellsDialog';
import { activeFormatBrush, captureFormatBrush, formatStyles, paintStyles, type FormatBrush } from '../application/style-commands';
import { underlinePatch, type AppearancePatch } from '../core/formats/appearance';
import { changeDimensions, changeSheet, editAxis } from '../application/structure-commands';
import type { WorkbookSnapshot } from '../application/workbook-protocol';
import { WorkbookSession, type SetCellsCommand, type StructureCommand } from '../application/workbook-session';
import type { Axis } from '../core/axis-edit';
import { NativeScheduler, type CalculationStatus } from '../calculation/native-scheduler';
import { findCell, findReferencedCell, newWorkbook, worksheetViewport, type CellReference, type Coordinate, type DocumentId, type WorkbookDocument } from '../core/workbook';
import { requireDocument } from '../core/workbook-validation';
import { Ribbon, type SheetView } from '../features/shell/Ribbon';
import { CellEditor, type CellEditorProps } from '../features/editing/CellEditor';
import { PasteTextDialog } from '../features/editing/PasteTextDialog';
import { localeArgumentSeparator } from '../features/formulas/formula-presentation';
import { useWorksheetEditing, type AfterEditKey, type EditAction, type EditSurface } from '../features/editing/useWorksheetEditing';
import { Worksheet, type WorksheetHandle } from '../grid/Worksheet';
import { StructureDialog, StructureMenu, type DialogAction, type StructureAction, type StructureInput } from '../features/structure/StructureControls';
import { fitDimensions } from '../grid/autofit';
import { address, clamp, gridMetrics } from '../grid/grid-geometry';
import {
  includes, parseAddress, primaryArea, selectionAddress, selectionAreas, selectionStatistics, singleSelection, type Area, type GridSelection,
} from '../grid/grid-selection';

const defaultSelection = singleSelection();
const defaultView: SheetView = { zoom: 100, headings: true, gridlines: true, formulaBar: true, showFormulas: false };
const statisticsFormat = new Intl.NumberFormat('en-US', { maximumSignificantDigits: 10, useGrouping: false });
const makeId = () => crypto.randomUUID();
interface StructureTarget {
  readonly id: string;
  readonly ranges: readonly EditRange[];
  readonly header?: number;
}
const dialogActions: ReadonlySet<StructureAction> = new Set(['insert-cells', 'delete-cells', 'delete-sheet', 'place-sheet', 'size-row', 'size-column']);

function visibleIndex(index: number, dimensions: WorkbookDocument['sheets'][number]['rows'], limit: number): number {
  const hidden = new Set(dimensions.filter(dimension => dimension.hidden).map(dimension => dimension.index));
  if (!hidden.has(index)) return index;
  for (let next = index + 1; next <= limit; next++) if (!hidden.has(next)) return next;
  for (let next = index - 1; next >= 1; next--) if (!hidden.has(next)) return next;
  return index;
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
      input.setCustomValidity('Enter a visible cell address from A1 to Z1000.');
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

type Icon = 'undo' | 'redo' | 'previous' | 'next';
const iconPath: Record<Icon, string> = {
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
  const [calculation, setCalculation] = useState<CalculationStatus>(() => ({ mode: 'automatic', pending: false, stale: new Set() }));
  const scheduler = useRef<NativeScheduler>(null);
  const [locale] = useState(() => navigator.language);
  const separator = useMemo(() => localeArgumentSeparator(locale), [locale]);
  useEffect(() => {
    const native = new NativeScheduler(session, (snapshot, status) => {
      setSnapshot(snapshot);
      setCalculation(status);
    }, locale);
    scheduler.current = native;
    if (session.active.sheets.some(sheet => sheet.cells.some(cell => cell.content.kind === 'formula'))) native.changed(false);
    return () => { native.dispose(); scheduler.current = null; };
  }, [session, locale]);
  const sheets = snapshot.document.sheets;
  const [activeSheetId, setActiveSheetId] = useState(() => sheets[0]!.id);
  const [selections, setSelections] = useState<ReadonlyMap<DocumentId, GridSelection>>(() => new Map());
  const [view, setView] = useState(defaultView);
  const worksheet = useRef<WorksheetHandle>(null);
  const root = useRef<HTMLDivElement>(null);
  const requestedFocus = useRef(false);
  const [menu, setMenu] = useState<{ x: number; y: number; target: StructureTarget; actions: readonly StructureAction[] }>();
  const [formatDraft, setFormatDraft] = useState<{ basis: WorkbookSnapshot; ranges: readonly EditRange[]; cell: CellReference; tab: FormatTab }>();
  const [brush, setBrush] = useState<(FormatBrush & { readonly locked: boolean })>();
  const painter = activeFormatBrush(snapshot, brush);
  const [structure, setStructure] = useState<{ action: DialogAction; target: StructureTarget; basis: WorkbookSnapshot }>();
  const [rename, setRename] = useState<{ id: string; basis: WorkbookSnapshot }>();
  const usedNumberCodes = useMemo(() => [...new Set(formatDraft?.basis.document.sheets.flatMap(item =>
    item.cells.flatMap(cell => cell.style?.numberFormat ?? [])))], [formatDraft]);
  const renaming = useRef(false);
  const sheet = sheets.find(candidate => candidate.id === activeSheetId) ?? sheets[0]!;
  const sheetIndex = sheets.indexOf(sheet);
  const selected = selections.get(sheet.id) ?? defaultSelection;
  const selection = useMemo(() => {
    const row = visibleIndex(selected.anchor.row, sheet.rows, worksheetViewport.rows);
    const column = visibleIndex(selected.anchor.column, sheet.columns, worksheetViewport.columns);
    return row === selected.anchor.row && column === selected.anchor.column ? selected : singleSelection({ row, column });
  }, [selected, sheet.rows, sheet.columns]);
  const stats = useMemo(() => selectionStatistics(sheet, selectionAreas(selection)), [sheet, selection]);
  const activeTarget = { sheetId: sheet.id, ...selection.anchor };
  const focusGrid = () => worksheet.current?.focus();
  const editing = useWorksheetEditing(session, compound => scheduler.current?.changed(compound), separator, locale, focusGrid, surface => {
    root.current?.querySelector<HTMLTextAreaElement>('[data-cell-editor="' + surface + '"]')?.focus({ preventScroll: true });
  });

  function setSelection(next: GridSelection): boolean {
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
      scheduler.current?.changed(true);
    }
    if (active && active !== sheet.id) {
      if (!painter?.locked) setBrush(undefined);
      requestedFocus.current = true;
      setActiveSheetId(active);
    } else requestAnimationFrame(focusGrid);
  }

  function currentAuthoredBasis(basis: WorkbookSnapshot) {
    requireDocument(session.isCurrentAuthored(basis), 'The workbook changed. Open the command again.');
    return session.capture();
  }

  function applyStyleCommand(command: SetCellsCommand) {
    const outcome = session.execute(command);
    editing.clearStatus();
    if (outcome.changed) {
      editing.selectionChanged();
      editing.clearCut();
      scheduler.current?.changed(true);
    }
  }

  function runNumberFormat(change: NumberFormatChange) {
    if (!editing.commit()) return;
    setBrush(undefined);
    try { applyStyleCommand(applyNumberFormat(session.capture(), ranges(), change)); }
    catch (error) { editing.reject(error); }
    requestAnimationFrame(focusGrid);
  }

  function runAppearance(patch: AppearancePatch, focus = true) {
    if (!editing.commit()) return;
    if (Object.keys(patch).length) {
      setBrush(undefined);
      try { applyStyleCommand(formatStyles(session.capture(), ranges(), patch)); }
      catch (error) { editing.reject(error); }
    }
    if (focus) requestAnimationFrame(focusGrid);
  }

  function armPainter(locked: boolean) {
    if (!editing.commit()) return;
    try {
      if (painter && (!locked || painter.locked)) setBrush(undefined);
      else {
        const selected = ranges();
        requireDocument(selected.length === 1, 'Format Painter needs a contiguous source range.');
        setBrush({ ...(painter ?? captureFormatBrush(session.capture(), selected[0]!)), locked });
      }
      editing.clearStatus();
    } catch (error) { editing.reject(error); }
    requestAnimationFrame(focusGrid);
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
    if (!editing.commit()) return;
    editing.selectionChanged();
    const basis = session.capture();
    if (action === 'format-cells') {
      setBrush(undefined);
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
    if (value === undefined) { requestAnimationFrame(focusGrid); return; }
    try {
      const basis = currentAuthoredBasis(rename.basis);
      const next = changeSheet(basis, rename.id, { kind: 'rename', name: value }, makeId);
      applyStructure(basis, next.command, next.active);
    } catch (error) { editing.reject(error); requestAnimationFrame(focusGrid); }
  }

  function openGridMenu(x: number, y: number, axis?: Axis, index?: number, point?: Coordinate) {
    if (!editing.commit()) return;
    const target = axis && index !== undefined ? targetHeader(axis, index) : { id: sheet.id, ranges:
      point && !selectionAreas(selection).some(area => includes(area, point))
        ? [{ sheetId: sheet.id, ...point, rows: 1, columns: 1 }] : ranges() };
    setMenu({ x, y, target, actions: axis ? [
      `insert-${axis}`, `delete-${axis}`, `size-${axis}`, `fit-${axis}`, `hide-${axis}`, `show-${axis}`,
    ] : ['insert-cells', 'delete-cells', 'format-cells'] });
  }

  function run(action: EditAction, event?: DataTransfer) {
    setBrush(undefined);
    void editing.run(action, ranges(), event).then(target => { if (target) worksheet.current?.goTo(target); });
  }

  function clipboardEvent(event: ClipboardEvent<HTMLDivElement>, kind: 'copy' | 'cut' | 'paste') {
    event.preventDefault();
    run(kind === 'paste' ? { kind, mode: 'all' } : { kind }, event.clipboardData);
  }

  function beginInline(text?: string) {
    setBrush(undefined);
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
    if (painter && key === 'enter' && !control && selection.kind === 'cells' && !selection.additional.length) {
      handle = () => paint(primaryArea(selection));
    } else if (control && ['b', 'i', 'u'].includes(key)) {
      const style = findCell(sheet, selection.anchor)?.style;
      handle = () => runAppearance(key === 'b' ? { bold: style?.bold ? undefined : true } :
        key === 'i' ? { italic: style?.italic ? undefined : true } : underlinePatch(style?.underline ? 'None' : 'Single'));
    } else if (control && key === '1') {
      handle = () => openStructure('format-cells');
    } else if (control && (key === 'z' || key === 'y')) {
      handle = () => editing.moveHistory(key === 'z' && !event.shiftKey);
    } else if (control && (key === 'd' || key === 'r')) {
      handle = () => run({ kind: 'fill', direction: key === 'd' ? 'down' : 'right' });
    } else if (control && event.shiftKey && key === 'v') {
      handle = () => run({ kind: 'paste', mode: 'values' });
    } else if (event.altKey && key === '=') {
      handle = () => run({ kind: 'autosum', aggregate: 'SUM' });
    } else if (!control && (key === 'delete' || key === 'backspace')) {
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
      textFormat: cell?.style?.numberFormat === '@',
      onKeepCaret: editing.keepCaret,
    };
  }

  const calculationMessage = calculation.error ??
    (calculation.pending ? calculation.pending === 'observe' ? 'Checking calculation status…' : 'Calculating…' :
      calculation.stale.size ? 'Calculation needed' : '');

  return (
    <div className="app-shell" ref={root} onKeyDownCapture={event => {
      if (event.key === 'Escape' && painter && !(event.target instanceof Element && event.target.closest('dialog'))) {
        event.preventDefault();
        event.stopPropagation();
        setBrush(undefined);
        editing.clearStatus();
      }
    }}>
      <header className="title-bar">
        <span className="brand-mark" aria-hidden="true"><svg viewBox="0 0 20 20" width="16" height="16" fill="none"
          stroke="currentColor" strokeWidth="1.4"><rect x="3" y="3" width="14" height="14" rx="1" /><path d="M3 8h14M3 12h14M8 3v14" /></svg></span>
        <div className="quick-access" role="group" aria-label="Quick Access">
          <button type="button" className="icon-button" aria-label="Undo" title={session.canUndo ? 'Undo' : 'Nothing to undo'}
            disabled={!session.canUndo} onClick={() => editing.moveHistory(true)}><ShellIcon icon="undo" /></button>
          <button type="button" className="icon-button" aria-label="Redo" title={session.canRedo ? 'Redo' : 'Nothing to redo'}
            disabled={!session.canRedo} onClick={() => editing.moveHistory(false)}><ShellIcon icon="redo" /></button>
        </div>
        <span className="workbook-title">Book1</span>
        <span className="product-title">ExcelLight Web</span>
      </header>

      <Ribbon cell={findCell(sheet, selection.anchor)} onNumberFormat={runNumberFormat} onAppearance={runAppearance}
        onAppearanceError={(error, focus) => { editing.reject(new Error(error)); if (focus) requestAnimationFrame(focusGrid); }}
        onFormatDialog={tab => openStructure('format-cells', undefined, tab)}
        painter={painter ? painter.locked ? 'locked' : 'single' : undefined} onPainter={armPainter} view={view} onView={onView} onEdit={run} onStructure={openStructure}
        calculationMode={calculation.mode} onCalculationMode={mode => {
          if (editing.commit()) scheduler.current?.setMode(mode);
        }} onCalculate={scope => {
          if (editing.commit()) scheduler.current?.calculate(scope === 'sheet' ? { sheetId: sheet.id } : 'workbook');
        }} onInsertFunction={name => {
          if (editing.commit()) editing.begin(activeTarget, view.formulaBar ? 'formula' : 'inline', '=' + name + '()', true);
        }} onZoomToSelection={() => {
        if (!editing.commit()) return;
        const zoom = worksheet.current?.zoomToSelection();
        if (zoom !== undefined) onView({ zoom });
        worksheet.current?.focus();
      }} />

      <main className="workbook" aria-label="Workbook">
        {(view.formulaBar || editing.draft?.surface === 'formula') && <div className="formula-bar">
          <NameBox selection={selection} goTo={goTo} cancel={() => worksheet.current?.focus()} />
          <span className="formula-symbol" aria-hidden="true">ƒx</span>
          <CellEditor {...editorProps('formula')} />
        </div>}
        <Worksheet key={sheet.id} ref={handle => {
          worksheet.current = handle;
          if (handle && requestedFocus.current) {
            handle.focus();
            requestedFocus.current = false;
          }
        }} sheet={sheet} selection={selection} onSelection={setSelection} onAdjacentSheet={chooseAdjacent}
          paintActive={!!painter} onPaint={paint}
          referenceArea={editing.draft?.pick?.sheetId === sheet.id ? editing.draft.pick.area : undefined}
          onReferencePick={() => {
            const caret = editorSelection();
            return caret && editing.beginReferencePick(caret, sheet);
          }}
          onEdit={beginInline} onShortcut={shortcut} onClipboard={clipboardEvent}
          onStructureMenu={openGridMenu}
          onHeaderFit={(axis, index) => openStructure(`fit-${axis}`, targetHeader(axis, index))}
          onHeaderResize={(axis, index) => {
            if (!editing.commit()) return;
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
          staleCells={calculation.stale} separator={separator} />
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

      <footer className="status-bar" aria-label="Worksheet status">
        <span className={editing.status || calculationMessage ? 'edit-status' : 'ready'} role="status">
          {[editing.status, painter && (painter.locked ? 'Format Painter locked' : 'Format Painter active') +
            ' · Select a destination or press Enter · Esc cancels', calculationMessage].filter(Boolean).join(' · ') || 'Ready'}
        </span>
        {!editing.status && !calculationMessage && <span className="grid-keyboard-hint" aria-hidden="true">F2 to edit · Esc, then Tab to leave</span>}
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
          <span className="zoom-value">{view.zoom}%</span>
        </div>
      </footer>
      {editing.manualPaste && <PasteTextDialog mode={editing.manualPaste.mode} error={editing.status}
        onSubmit={editing.submitManualPaste} onClose={editing.closeManualPaste} />}
      {menu && <StructureMenu {...menu} onAction={action => openStructure(action, menu.target)}
        onClose={() => { setMenu(undefined); requestAnimationFrame(focusGrid); }} />}
      {formatDraft && <FormatCellsDialog cell={findReferencedCell(snapshot.document, formatDraft.cell)} locale={locale}
        usedCodes={usedNumberCodes} initialTab={formatDraft.tab}
        onCancel={() => { setFormatDraft(undefined); requestAnimationFrame(focusGrid); }}
        onSubmit={(patch, code) => {
          try {
            if (code !== undefined || Object.keys(patch).length) {
              const basis = currentAuthoredBasis(formatDraft.basis);
              applyStyleCommand(formatStyles(basis, formatDraft.ranges, patch, code));
            }
            setFormatDraft(undefined);
            requestAnimationFrame(focusGrid);
          } catch (error) { return error instanceof Error ? error.message : 'The number format could not be applied.'; }
        }} />}
      {structure && <StructureDialog action={structure.action} sheets={structure.basis.document.sheets}
        target={structure.basis.document.sheets.find(item => item.id === structure.target.id)!}
        initialSize={(() => {
          const axis = structure.action === 'size-row' ? 'row' : 'column';
          const item = structure.basis.document.sheets.find(item => item.id === structure.target.id)!;
          const index = structure.target.header ?? structure.target.ranges[0]?.[axis];
          return item[axis === 'row' ? 'rows' : 'columns'].find(dimension => dimension.index === index)?.size ?? gridMetrics[axis];
        })()}
        onCancel={() => { setStructure(undefined); requestAnimationFrame(focusGrid); }}
        onSubmit={input => {
          try { performStructure(structure.action, structure.target, structure.basis, input); setStructure(undefined); }
          catch (error) { return error instanceof Error ? error.message : 'The structural edit failed.'; }
        }} />}
    </div>
  );
}
