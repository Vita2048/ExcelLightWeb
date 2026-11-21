import { useEffect, useLayoutEffect, useRef, useState, type KeyboardEvent, type PointerEvent } from 'react';
import type {
  SearchHistory, SearchLook, SearchMatch, SearchOptions, SearchOrder, SearchResultRow, SearchScope,
} from '../../application/search-commands';
import { DialogTitleBar } from '../dialogs/DialogTitleBar';

export interface SearchOutcome {
  readonly status: string;
  readonly rows?: readonly SearchResultRow[];
  readonly match?: SearchMatch;
}

export interface FindReplaceDialogProps {
  readonly replace: boolean;
  readonly history: SearchHistory;
  /** Session default: the page starts on Values, then keeps the last choice made in the Find tab. */
  readonly initialLookIn: SearchLook;
  readonly onLookIn: (lookIn: SearchLook) => void;
  readonly onFindAll: (options: SearchOptions) => SearchOutcome;
  readonly onFindNext: (options: SearchOptions, current: SearchMatch | undefined) => SearchOutcome;
  readonly onReplace: (options: SearchOptions, replacement: string, target: 'all' | 'one', selected: SearchMatch | undefined,
    current: SearchMatch | undefined, listShown: boolean) => SearchOutcome;
  readonly onSelect: (match: SearchMatch) => void;
  readonly onClose: () => void;
}

const columnLabels = ['Book', 'Sheet', 'Name', 'Cell', 'Value', 'Formula'] as const;
const defaultWidths = [110, 90, 60, 70, 100, 250] as const;

function HistoryCombo({ label, value, history, autoFocus, onChange, onEnter }: {
  label: string; value: string; history: readonly string[]; autoFocus?: boolean;
  onChange: (value: string) => void; onEnter: () => void;
}) {
  const [open, setOpen] = useState(false);
  const id = label.replace(/\s+/gu, '-').toLowerCase();
  return <div className="find-field">
    <label htmlFor={id}>{label}:</label>
    <div className="find-combo">
      <input id={id} type="text" autoFocus={autoFocus} autoComplete="off" spellCheck={false} value={value}
        aria-label={label}
        onChange={event => onChange(event.currentTarget.value)}
        onKeyDown={event => {
          if (event.key === 'Enter' && !open) {
            event.preventDefault();
            onEnter();
          } else if (event.key === 'ArrowDown') {
            event.preventDefault();
            setOpen(true);
          } else if (event.key === 'Escape' && open) {
            event.stopPropagation();
            setOpen(false);
          }
        }} />
      <button type="button" className="find-combo-button" aria-label={'Show ' + label + ' history'}
        aria-expanded={open} aria-haspopup="listbox" title={label + ' history'}
        onClick={() => setOpen(current => !current)}>▾</button>
      {open && <ul className="find-history" role="listbox" aria-label={label + ' history'}>
        {history.length === 0 ? <li className="find-history-empty">No recent values</li> : history.map(entry =>
          <li key={entry} role="option" aria-selected={entry === value}>
            <button type="button" onClick={() => { onChange(entry); setOpen(false); }}>{entry || '(blank)'}</button></li>)}
      </ul>}
    </div>
  </div>;
}

/** Find and Replace, following SearchDialog.cs: the Find/Replace tabs, Options, Look in/Within/Search and the results table. */
export function FindReplaceDialog({ replace, history, initialLookIn, onLookIn, onFindAll, onFindNext, onReplace, onSelect, onClose }: FindReplaceDialogProps) {
  const dialog = useRef<HTMLDialogElement>(null);
  const [tab, setTab] = useState<'find' | 'replace'>(replace ? 'replace' : 'find');
  // The desktop dialog reopens with the text from the last session's fields.
  const [query, setQuery] = useState(history.find[0] ?? '');
  const [replacement, setReplacement] = useState(history.replace[0] ?? '');
  // Only the Find tab offers Values; the Replace tab always searches formulas, as on the desktop.
  const [lookIn, setLookIn] = useState<SearchLook>(replace ? 'formulas' : initialLookIn);
  const [within, setWithin] = useState<SearchScope>('sheet');
  const [order, setOrder] = useState<SearchOrder>('rows');
  const [matchCase, setMatchCase] = useState(false);
  const [wholeCell, setWholeCell] = useState(false);
  const [expanded, setExpanded] = useState(false);
  const [status, setStatus] = useState('');
  const [rows, setRows] = useState<readonly SearchResultRow[]>([]);
  const [listShown, setListShown] = useState(false);
  const [selected, setSelected] = useState<SearchMatch>();
  const [current, setCurrent] = useState<SearchMatch>();
  const [widths, setWidths] = useState<readonly number[]>(defaultWidths);

  useLayoutEffect(() => { if (!dialog.current!.open) dialog.current!.showModal(); }, []);

  // The layout follows SearchDialog's narrow rule: the options become one column when the pane is small.
  const panel = useRef<HTMLDivElement>(null);
  const [narrow, setNarrow] = useState(false);
  useEffect(() => {
    const node = panel.current;
    if (!node) return;
    const observer = new ResizeObserver(() => setNarrow(node.clientWidth < 500));
    observer.observe(node);
    return () => observer.disconnect();
  }, []);

  function options(): SearchOptions {
    return { query, lookIn: tab === 'replace' ? 'formulas' : lookIn, within, order, matchCase, wholeCell,
      includeConstants: tab === 'replace' };
  }

  function show(outcome: SearchOutcome) {
    setStatus(outcome.status);
    if (outcome.rows !== undefined) {
      setRows(outcome.rows);
      setListShown(true);
    }
  }

  function findAll() {
    const outcome = onFindAll(options());
    show(outcome);
    if (outcome.rows?.length) {
      setSelected(outcome.rows[0]!.match);
      setCurrent(outcome.rows[0]!.match);
    }
  }

  function findNext() {
    const outcome = onFindNext(options(), current);
    show(outcome);
    if (outcome.match) {
      setCurrent(outcome.match);
      if (listShown) setSelected(outcome.match);
    }
  }

  function runReplace(target: 'all' | 'one') {
    const outcome = onReplace(options(), replacement, target, selected, current, listShown);
    show(outcome);
    if (target === 'one') {
      if (outcome.match) {
        setCurrent(outcome.match);
        if (listShown) setSelected(outcome.match);
      } else if (listShown) setSelected(undefined);
    }
  }

  function switchTab(next: 'find' | 'replace') {
    if (next === tab) return;
    // Replace only searches sources, as on the desktop, so Values clamps back to Formulas.
    if (next === 'replace' && lookIn === 'values') {
      setLookIn('formulas');
      onLookIn('formulas');
    }
    setTab(next);
  }

  function resizeColumn(index: number, event: PointerEvent<HTMLSpanElement>) {
    event.preventDefault();
    const handle = event.currentTarget;
    const start = event.clientX;
    const initial = widths[index] ?? 100;
    handle.setPointerCapture(event.pointerId);
    function move(moveEvent: globalThis.PointerEvent) {
      setWidths(currentWidths => currentWidths.map((width, position) =>
        position === index ? Math.max(40, Math.round(initial + moveEvent.clientX - start)) : width));
    }
    function up() {
      handle.removeEventListener('pointermove', move);
      handle.removeEventListener('pointerup', up);
      handle.removeEventListener('pointercancel', up);
    }
    handle.addEventListener('pointermove', move);
    handle.addEventListener('pointerup', up);
    handle.addEventListener('pointercancel', up);
  }

  function keyResize(index: number, event: KeyboardEvent<HTMLSpanElement>) {
    const step = event.key === 'ArrowLeft' ? -8 : event.key === 'ArrowRight' ? 8 : 0;
    if (!step) return;
    event.preventDefault();
    setWidths(currentWidths => currentWidths.map((width, position) =>
      position === index ? Math.max(40, width + step) : width));
  }

  function chooseRow(row: SearchResultRow) {
    setSelected(row.match);
    setCurrent(row.match);
    onSelect(row.match);
  }

  const width = widths.reduce((total, value) => total + value, 0);
  return <dialog ref={dialog} className="paste-dialog find-dialog" aria-label="Find and Replace"
    onCancel={event => { event.preventDefault(); onClose(); }}>
    <DialogTitleBar><h2>Find and Replace</h2>
      <button type="button" className="format-dialog-close" aria-label="Close Find and Replace" onClick={onClose}>×</button>
    </DialogTitleBar>
    <div className="find-tabs" role="tablist" aria-label="Find and Replace tabs">
      {(['find', 'replace'] as const).map(name => <button key={name} type="button" role="tab" id={'find-tab-' + name}
        aria-selected={tab === name} aria-controls="find-pane" tabIndex={tab === name ? 0 : -1}
        onClick={() => switchTab(name)}>{name === 'find' ? 'Find' : 'Replace'}</button>)}
    </div>
    <div ref={panel} id="find-pane" role="tabpanel" aria-labelledby={'find-tab-' + tab} className="find-pane">
      <HistoryCombo label="Find what" value={query} history={history.find} autoFocus
        onChange={value => setQuery(value)} onEnter={findNext} />
      {tab === 'replace' && <HistoryCombo label="Replace with" value={replacement} history={history.replace}
        onChange={value => setReplacement(value)} onEnter={findNext} />}
      {expanded && <div className={'find-options' + (narrow ? ' find-options-narrow' : '')}>
        <div className="find-field find-option-within"><label htmlFor="find-within">Within:</label>
          <select id="find-within" aria-label="Within" value={within}
            onChange={event => setWithin(event.currentTarget.value as SearchScope)}>
            <option value="sheet">Sheet</option><option value="workbook">Workbook</option>
          </select></div>
        <div className="find-field find-option-order"><label htmlFor="find-order">Search:</label>
          <select id="find-order" aria-label="Search" value={order}
            onChange={event => setOrder(event.currentTarget.value as SearchOrder)}>
            <option value="rows">By Rows</option><option value="columns">By Columns</option>
          </select></div>
        <label className="find-check find-option-case"><input type="checkbox" aria-label="Match case" checked={matchCase}
          onChange={event => setMatchCase(event.currentTarget.checked)} /> Match case</label>
        <label className="find-check find-option-whole"><input type="checkbox" aria-label="Match entire cell contents" checked={wholeCell}
          onChange={event => setWholeCell(event.currentTarget.checked)} /> Match entire cell contents</label>
        <div className="find-field find-option-look"><label htmlFor="find-look">Look in:</label>
          <select id="find-look" aria-label="Look in" value={tab === 'replace' ? 'formulas' : lookIn}
            disabled={tab === 'replace'}
            onChange={event => {
              const value = event.currentTarget.value as SearchLook;
              setLookIn(value);
              onLookIn(value);
            }}>
            <option value="formulas">Formulas</option>
            {tab === 'find' && <option value="values">Values</option>}
          </select></div>
      </div>}
      <div className="find-status-row">
        <span className="find-status" role="status">{status}</span>
        <button type="button" className="find-options-toggle" aria-expanded={expanded}
          onClick={() => setExpanded(currentValue => !currentValue)}>{expanded ? 'Options <<' : 'Options >>'}</button>
      </div>
      {listShown && <div className="find-result-table">
        <table style={{ width }} aria-label="Find All results">
          <colgroup>{widths.map((columnWidth, index) => <col key={index} style={{ width: columnWidth }} />)}</colgroup>
          <thead><tr>{columnLabels.map((label, index) => <th key={label} scope="col">
            {label}
            <span className="find-resize" role="separator" aria-label={'Resize ' + label + ' column'} tabIndex={0}
              onPointerDown={event => resizeColumn(index, event)}
              onKeyDown={event => keyResize(index, event)} />
          </th>)}</tr></thead>
          <tbody>{rows.map(row => <tr key={row.match.sheetId + '-' + row.match.row + '-' + row.match.column}
            aria-selected={selected?.sheetId === row.match.sheetId && selected.row === row.match.row &&
              selected.column === row.match.column} tabIndex={0}
            onClick={() => chooseRow(row)}
            onKeyDown={event => {
              if (event.key === 'Enter' || event.key === ' ') {
                event.preventDefault();
                chooseRow(row);
              }
            }}>
            <td>{row.book}</td><td>{row.sheet}</td><td>{row.name}</td><td>{row.cell}</td>
            <td>{row.value}</td><td>{row.formula}</td>
          </tr>)}</tbody>
        </table>
      </div>}
      <div className="find-actions">
        <div className="find-actions-group">
          {tab === 'replace' && <>
            <button type="button" onClick={() => runReplace('all')}>Replace All</button>
            <button type="button" onClick={() => runReplace('one')}>Replace</button>
          </>}
        </div>
        <div className="find-actions-group">
          <button type="button" onClick={findAll}>Find All</button>
          <button type="button" className="find-primary" onClick={findNext}>Find Next</button>
          <button type="button" onClick={onClose}>Close</button>
        </div>
      </div>
    </div>
  </dialog>;
}
