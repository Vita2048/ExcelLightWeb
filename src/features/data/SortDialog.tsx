import { useLayoutEffect, useRef, useState, type FormEvent, type RefObject } from 'react';
import { findCell, type SheetDocument } from '../../core/workbook';
import { sortKeyFor } from '../../application/data-commands';
import { columnLetters } from '../../core/formula-references';
import { lastColumn, lastRow } from '../../core/workbook-names';
import { DialogTitleBar } from '../dialogs/DialogTitleBar';
import { DialogActions } from '../dialogs/DialogActions';

export interface SortDraft {
  readonly column: number;
  readonly ascending: boolean;
  readonly header: boolean;
  readonly caseSensitive: boolean;
}
export interface SortExtent {
  readonly row: number;
  readonly column: number;
  readonly rows: number;
  readonly columns: number;
}

function useModal(): RefObject<HTMLDialogElement | null> {
  const dialog = useRef<HTMLDialogElement>(null);
  useLayoutEffect(() => { if (!dialog.current!.open) dialog.current!.showModal(); }, []);
  return dialog;
}

function headerLabel(sheet: SheetDocument, row: number, column: number): string {
  const cell = sheet.cells.find(item => item.row === row && item.column === column);
  return cell?.content.kind === 'text' && cell.content.value.length > 0 ? cell.content.value : 'Column ' + columnLetters(column);
}

/** The desktop Sort, Sort Options and Sort Warning dialogs. */
export function SortDialog({ sheet, range, draft, onSubmit, onOptions, onCancel }: {
  sheet: SheetDocument;
  range: SortExtent;
  draft: SortDraft;
  onSubmit: (draft: SortDraft) => void;
  onOptions: (draft: SortDraft) => void;
  onCancel: () => void;
}) {
  const dialog = useModal();
  const [column, setColumn] = useState(draft.column);
  const [ascending, setAscending] = useState(draft.ascending);
  const [header, setHeader] = useState(draft.header);
  const first = header ? range.row + 1 : range.row;
  let allText = true, anyValue = false;
  for (let row = first; row <= lastRow(range); row++) {
    const key = sortKeyFor(findCell(sheet, { row, column }), row);
    if (key.kind === 'blank') continue;
    anyValue = true;
    if (key.kind !== 'text') allText = false;
  }
  const textOrder = anyValue && allText;
  const labels = new Map<number, string>();
  const used = new Map<string, number>();
  for (let index = range.column; index <= lastColumn(range); index++) {
    const label = header ? headerLabel(sheet, range.row, index) : 'Column ' + columnLetters(index);
    labels.set(index, label);
    used.set(label.toLowerCase(), (used.get(label.toLowerCase()) ?? 0) + 1);
  }
  const submit = (event: FormEvent) => {
    event.preventDefault();
    onSubmit({ column, ascending, header, caseSensitive: draft.caseSensitive });
  };
  return <dialog ref={dialog} className="paste-dialog sort-dialog" aria-label="Sort"
    onCancel={event => { event.preventDefault(); onCancel(); }}>
    <form onSubmit={submit}>
      <DialogTitleBar><h2>Sort</h2>
        <button type="button" className="format-dialog-close" aria-label="Close Sort" onClick={onCancel}>×</button>
      </DialogTitleBar>
      <div className="sort-tools">
        <button type="button" className="sort-options" onClick={() => onOptions({ column, ascending, header, caseSensitive: draft.caseSensitive })}>Options...</button>
        <label><input type="checkbox" aria-label="My data has headers" checked={header}
          onChange={event => setHeader(event.currentTarget.checked)} /> My data has headers</label>
      </div>
      <div className="sort-grid">
        <span />
        <span>Column</span><span>Sort On</span><span>Order</span>
        <span>Sort by</span>
        <select aria-label="Sort column" value={column} autoFocus onChange={event => setColumn(Number(event.currentTarget.value))}>
          {[...labels].map(([index, label]) => <option key={index} value={index}>
            {(used.get(label.toLowerCase()) ?? 0) > 1 ? label + ' (' + columnLetters(index) + ')' : label}
          </option>)}
        </select>
        <select aria-label="Sort on" defaultValue="values">
          <option value="values">Cell Values</option>
        </select>
        <select aria-label="Sort order" value={ascending ? 'asc' : 'desc'}
          onChange={event => setAscending(event.currentTarget.value === 'asc')}>
          <option value="asc">{textOrder ? 'A to Z' : 'Smallest to Largest'}</option>
          <option value="desc">{textOrder ? 'Z to A' : 'Largest to Smallest'}</option>
        </select>
      </div>
      <DialogActions onCancel={onCancel} />
    </form>
  </dialog>;
}

export function SortOptionsDialog({ caseSensitive, onSubmit, onCancel }: {
  caseSensitive: boolean;
  onSubmit: (caseSensitive: boolean) => void;
  onCancel: () => void;
}) {
  const dialog = useModal();
  const [sensitive, setSensitive] = useState(caseSensitive);
  return <dialog ref={dialog} className="paste-dialog sort-options-dialog" aria-label="Sort Options"
    onCancel={event => { event.preventDefault(); onCancel(); }}>
    <form onSubmit={event => { event.preventDefault(); onSubmit(sensitive); }}>
      <DialogTitleBar><h2>Sort Options</h2>
        <button type="button" className="format-dialog-close" aria-label="Close Sort Options" onClick={onCancel}>×</button>
      </DialogTitleBar>
      <label className="sort-check"><input type="checkbox" aria-label="Case sensitive" checked={sensitive}
        onChange={event => setSensitive(event.currentTarget.checked)} /> Case sensitive</label>
      <h3 className="sort-heading">Orientation</h3>
      <label className="sort-check"><input type="radio" name="sort-orientation" aria-label="Sort top to bottom"
        defaultChecked /> Sort top to bottom</label>
      <DialogActions onCancel={onCancel} />
    </form>
  </dialog>;
}

export function SortWarningDialog({ selection, region, onChoose, onCancel }: {
  selection: SortExtent;
  region: SortExtent;
  onChoose: (expand: boolean) => void;
  onCancel: () => void;
}) {
  const dialog = useModal();
  const [expand, setExpand] = useState(true);
  const selectionBox = columnLetters(selection.column) + selection.row + ':' +
    columnLetters(lastColumn(selection)) + lastRow(selection);
  const regionBox = columnLetters(region.column) + region.row + ':' + columnLetters(lastColumn(region)) + lastRow(region);
  return <dialog ref={dialog} className="paste-dialog sort-warning-dialog" aria-label="Sort Warning"
    onCancel={event => { event.preventDefault(); onCancel(); }}>
    <form onSubmit={event => { event.preventDefault(); onChoose(expand); }}>
      <DialogTitleBar><h2>Sort Warning</h2>
        <button type="button" className="format-dialog-close" aria-label="Close Sort Warning" onClick={onCancel}>×</button>
      </DialogTitleBar>
      <p>There is data next to {selectionBox}. Choose which data to sort.</p>
      <label className="sort-check"><input type="radio" name="sort-range" aria-label="Expand the selection"
        checked={expand} onChange={() => setExpand(true)} /> Expand the selection</label>
      <label className="sort-check"><input type="radio" name="sort-range" aria-label="Continue with the current selection"
        checked={!expand} onChange={() => setExpand(false)} /> Continue with the current selection</label>
      <p className="sort-preview">{expand ? 'Sort range: ' + regionBox :
        'Sort range: ' + selectionBox + '. Adjacent columns will stay in place.'}</p>
      <DialogActions confirmLabel="Sort" onCancel={onCancel} />
    </form>
  </dialog>;
}
