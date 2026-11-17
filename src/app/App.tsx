import { useMemo, useRef, useState, type FormEvent } from 'react';
import { WorkbookSession } from '../application/workbook-session';
import { findCell, newWorkbook, type DocumentId, type WorkbookDocument } from '../core/workbook';
import { Ribbon, type SheetView } from '../features/shell/Ribbon';
import { Worksheet, type WorksheetHandle } from '../grid/Worksheet';
import { clamp } from '../grid/grid-geometry';
import {
  cellDisplay, parseAddress, selectionAddress, selectionAreas, selectionStatistics, singleSelection, type GridSelection,
} from '../grid/grid-selection';

const defaultSelection = singleSelection();
const defaultView: SheetView = { zoom: 100, headings: true, gridlines: true, formulaBar: true };
const statisticsFormat = new Intl.NumberFormat('en-US', { maximumSignificantDigits: 10, useGrouping: false });
const makeId = () => crypto.randomUUID();

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
  const sheets = snapshot.document.sheets;
  const [activeSheetId, setActiveSheetId] = useState(() => sheets[0]!.id);
  const [selections, setSelections] = useState<ReadonlyMap<DocumentId, GridSelection>>(() => new Map());
  const [view, setView] = useState(defaultView);
  const worksheet = useRef<WorksheetHandle>(null);
  const requestedFocus = useRef(false);
  const sheet = sheets.find(candidate => candidate.id === activeSheetId) ?? sheets[0]!;
  const sheetIndex = sheets.indexOf(sheet);
  const selection = selections.get(sheet.id) ?? defaultSelection;
  const stats = useMemo(() => selectionStatistics(sheet, selectionAreas(selection)), [sheet, selection]);

  function setSelection(next: GridSelection) {
    setSelections(current => new Map(current).set(sheet.id, next));
  }

  function chooseSheet(nextSheetId: DocumentId) {
    if (nextSheetId === sheet.id) {
      worksheet.current?.focus();
      return;
    }
    requestedFocus.current = true;
    setActiveSheetId(nextSheetId);
  }

  function chooseAdjacent(direction: number) {
    const next = sheets[sheetIndex + direction];
    if (next) chooseSheet(next.id);
  }

  function onView(patch: Partial<SheetView>) {
    setView(current => ({ ...current, ...patch }));
  }

  function goTo(value: string) {
    const point = parseAddress(value);
    return !!point && !!worksheet.current?.goTo(point);
  }

  function moveHistory(undo: boolean) {
    if (undo ? session.undo() : session.redo()) setSnapshot(session.capture());
    worksheet.current?.focus();
  }

  return (
    <div className="app-shell">
      <header className="title-bar">
        <span className="brand-mark" aria-hidden="true"><svg viewBox="0 0 20 20" width="16" height="16" fill="none"
          stroke="currentColor" strokeWidth="1.4"><rect x="3" y="3" width="14" height="14" rx="1" /><path d="M3 8h14M3 12h14M8 3v14" /></svg></span>
        <div className="quick-access" role="group" aria-label="Quick Access">
          <button type="button" className="icon-button" aria-label="Undo" title={session.canUndo ? 'Undo' : 'Nothing to undo'}
            disabled={!session.canUndo} onClick={() => moveHistory(true)}><ShellIcon icon="undo" /></button>
          <button type="button" className="icon-button" aria-label="Redo" title={session.canRedo ? 'Redo' : 'Nothing to redo'}
            disabled={!session.canRedo} onClick={() => moveHistory(false)}><ShellIcon icon="redo" /></button>
        </div>
        <span className="workbook-title">Book1</span>
        <span className="product-title">ExcelLight Web</span>
      </header>

      <Ribbon view={view} onView={onView} onZoomToSelection={() => {
        const zoom = worksheet.current?.zoomToSelection();
        if (zoom !== undefined) onView({ zoom });
        worksheet.current?.focus();
      }} />

      <main className="workbook" aria-label="Workbook">
        {view.formulaBar && <div className="formula-bar">
          <NameBox selection={selection} goTo={goTo} cancel={() => worksheet.current?.focus()} />
          <span className="formula-symbol" aria-hidden="true">ƒx</span>
          <input readOnly aria-label="Formula bar" className="formula-value"
            value={cellDisplay(findCell(sheet, selection.anchor), true)} />
        </div>}
        <Worksheet key={sheet.id} ref={handle => {
          worksheet.current = handle;
          if (handle && requestedFocus.current) {
            handle.focus();
            requestedFocus.current = false;
          }
        }} sheet={sheet} selection={selection} onSelection={setSelection} onAdjacentSheet={chooseAdjacent}
          zoom={view.zoom} headings={view.headings} gridlines={view.gridlines} />
        <nav className="sheet-strip" aria-label="Worksheets">
          <div className="sheet-navigation">
            <button type="button" className="icon-button" aria-label="Previous sheet" disabled={sheetIndex === 0}
              onClick={() => chooseAdjacent(-1)}><ShellIcon icon="previous" /></button>
            <button type="button" className="icon-button" aria-label="Next sheet" disabled={sheetIndex === sheets.length - 1}
              onClick={() => chooseAdjacent(1)}><ShellIcon icon="next" /></button>
          </div>
          <div className="sheet-tabs">
            {sheets.map(item => <button type="button" key={item.id} aria-current={item.id === sheet.id ? 'true' : undefined}
              onClick={() => chooseSheet(item.id)}>{item.name}</button>)}
          </div>
        </nav>
      </main>

      <footer className="status-bar" aria-label="Worksheet status">
        <span className="ready">Ready</span>
        <span className="grid-keyboard-hint" aria-hidden="true">Esc, then Tab to leave the grid</span>
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
    </div>
  );
}
