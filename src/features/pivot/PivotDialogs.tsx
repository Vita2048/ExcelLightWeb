import { useId, useLayoutEffect, useRef, useState } from 'react';
import type { PivotTableDocument } from '../../core/pivot';
import { DialogActions } from '../dialogs/DialogActions';
import { DialogTitleBar } from '../dialogs/DialogTitleBar';
import type { PivotPlacement } from './CreatePivotTableDialog';

export function PivotOptionsDialog({
  pivot,
  onSave,
  onCancel,
}: {
  readonly pivot: PivotTableDocument;
  readonly onSave: (patch: Partial<Omit<PivotTableDocument, 'id' | 'extent'>>) => void;
  readonly onCancel: () => void;
}) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const emptyInputId = useId();
  const [emptyCellText, setEmptyCellText] = useState(pivot.emptyCellText ?? '');
  const [autoFitColumns, setAutoFitColumns] = useState(pivot.autoFitColumns !== false);
  const [showButtons, setShowButtons] = useState(pivot.showButtons !== false);
  const [showFieldHeaders, setShowFieldHeaders] = useState(pivot.showFieldHeaders !== false);
  const [grandTotalRows, setGrandTotalRows] = useState(pivot.grandTotals.rows);
  const [grandTotalCols, setGrandTotalCols] = useState(pivot.grandTotals.columns);
  const [repeatLabels, setRepeatLabels] = useState(Boolean(pivot.repeatLabels));

  useLayoutEffect(() => {
    const dialog = dialogRef.current;
    if (dialog && !dialog.open) dialog.showModal();
  }, []);

  return (
    <dialog
      ref={dialogRef}
      className="paste-dialog pivot-options-dialog"
      aria-label="PivotTable Options"
      onCancel={event => {
        event.preventDefault();
        onCancel();
      }}
    >
      <form
        onSubmit={event => {
          event.preventDefault();
          onSave({
            emptyCellText,
            autoFitColumns,
            showButtons,
            showFieldHeaders,
            grandTotals: { rows: grandTotalRows, columns: grandTotalCols },
            repeatLabels,
          });
        }}
      >
        <DialogTitleBar>
          <h2>PivotTable Options — {pivot.name}</h2>
          <button type="button" className="format-dialog-close" aria-label="Close PivotTable Options" onClick={onCancel}>
            ×
          </button>
        </DialogTitleBar>

        <fieldset className="pivot-dialog-group">
          <legend>Layout &amp; Format</legend>
          <div className="pivot-options-row">
            <label htmlFor={emptyInputId}>For empty cells show:</label>
            <input
              id={emptyInputId}
              type="text"
              value={emptyCellText}
              placeholder="Leave blank"
              onChange={event => setEmptyCellText(event.currentTarget.value)}
            />
          </div>
          <label className="pivot-options-check">
            <input type="checkbox" checked={autoFitColumns} onChange={event => setAutoFitColumns(event.currentTarget.checked)} />
            Autofit column widths on update
          </label>
          <label className="pivot-options-check">
            <input type="checkbox" checked={showButtons} onChange={event => setShowButtons(event.currentTarget.checked)} />
            Show expand/collapse buttons
          </label>
          <label className="pivot-options-check">
            <input type="checkbox" checked={showFieldHeaders} onChange={event => setShowFieldHeaders(event.currentTarget.checked)} />
            Display field captions and filter drop downs
          </label>
          <label className="pivot-options-check">
            <input type="checkbox" checked={repeatLabels} onChange={event => setRepeatLabels(event.currentTarget.checked)} />
            Repeat all item labels
          </label>
        </fieldset>

        <fieldset className="pivot-dialog-group">
          <legend>Totals &amp; Filters</legend>
          <label className="pivot-options-check">
            <input type="checkbox" checked={grandTotalRows} onChange={event => setGrandTotalRows(event.currentTarget.checked)} />
            Show grand totals for rows
          </label>
          <label className="pivot-options-check">
            <input type="checkbox" checked={grandTotalCols} onChange={event => setGrandTotalCols(event.currentTarget.checked)} />
            Show grand totals for columns
          </label>
        </fieldset>

        <DialogActions confirmLabel="OK" onCancel={onCancel} />
      </form>
    </dialog>
  );
}

export function ChangePivotSourceDialog({
  sourceText,
  suspended = false,
  onChangeSource,
  onPickSource,
  onSubmit,
  onCancel,
}: {
  readonly sourceText: string;
  readonly suspended?: boolean;
  readonly onChangeSource: (value: string) => void;
  readonly onPickSource: () => void;
  readonly onSubmit: () => string | undefined;
  readonly onCancel: () => void;
}) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const sourceId = useId();
  const [error, setError] = useState('');

  useLayoutEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog) return;
    if (suspended) {
      if (dialog.open) dialog.close();
      return;
    }
    if (!dialog.open) dialog.showModal();
    inputRef.current?.focus();
  }, [suspended]);

  return (
    <dialog
      ref={dialogRef}
      className="paste-dialog create-pivot-dialog"
      aria-label="Change PivotTable Data Source"
      onCancel={event => {
        event.preventDefault();
        if (!suspended) onCancel();
      }}
    >
      <form
        onSubmit={event => {
          event.preventDefault();
          const err = onSubmit();
          if (err) setError(err);
        }}
      >
        <DialogTitleBar>
          <h2>Change PivotTable Data Source</h2>
          <button type="button" className="format-dialog-close" aria-label="Close Change Data Source" onClick={onCancel}>
            ×
          </button>
        </DialogTitleBar>

        <fieldset className="pivot-dialog-group">
          <legend>Select a table or range</legend>
          <div className="pivot-dialog-range-row">
            <label htmlFor={sourceId}>Table/Range:</label>
            <input
              ref={inputRef}
              id={sourceId}
              type="text"
              value={sourceText}
              onChange={event => {
                setError('');
                onChangeSource(event.currentTarget.value);
              }}
            />
            <button
              type="button"
              className="refers-control-btn"
              aria-label="Pick Table/Range"
              title="Select range from worksheet"
              onClick={onPickSource}
            >
              ↑
            </button>
          </div>
        </fieldset>

        {error && <p className="dialog-error" role="alert">{error}</p>}
        <DialogActions confirmLabel="OK" onCancel={onCancel} />
      </form>
    </dialog>
  );
}

export function MovePivotTableDialog({
  placement,
  locationText,
  suspended = false,
  onChangePlacement,
  onChangeLocation,
  onPickLocation,
  onSubmit,
  onCancel,
}: {
  readonly placement: PivotPlacement;
  readonly locationText: string;
  readonly suspended?: boolean;
  readonly onChangePlacement: (value: PivotPlacement) => void;
  readonly onChangeLocation: (value: string) => void;
  readonly onPickLocation: () => void;
  readonly onSubmit: () => string | undefined;
  readonly onCancel: () => void;
}) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const locationId = useId();
  const newSheetId = useId();
  const existingSheetId = useId();
  const [error, setError] = useState('');

  useLayoutEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog) return;
    if (suspended) {
      if (dialog.open) dialog.close();
      return;
    }
    if (!dialog.open) dialog.showModal();
  }, [suspended]);

  return (
    <dialog
      ref={dialogRef}
      className="paste-dialog create-pivot-dialog"
      aria-label="Move PivotTable"
      onCancel={event => {
        event.preventDefault();
        if (!suspended) onCancel();
      }}
    >
      <form
        onSubmit={event => {
          event.preventDefault();
          const err = onSubmit();
          if (err) setError(err);
        }}
      >
        <DialogTitleBar>
          <h2>Move PivotTable</h2>
          <button type="button" className="format-dialog-close" aria-label="Close Move PivotTable" onClick={onCancel}>
            ×
          </button>
        </DialogTitleBar>

        <fieldset className="pivot-dialog-group">
          <legend>Choose where you want the PivotTable to be placed</legend>
          <div className="pivot-dialog-radio">
            <input
              id={newSheetId}
              type="radio"
              name="move-pivot-placement"
              checked={placement === 'new'}
              onChange={() => {
                setError('');
                onChangePlacement('new');
              }}
            />
            <label htmlFor={newSheetId}>New Worksheet</label>
          </div>
          <div className="pivot-dialog-radio">
            <input
              id={existingSheetId}
              type="radio"
              name="move-pivot-placement"
              checked={placement === 'existing'}
              onChange={() => {
                setError('');
                onChangePlacement('existing');
              }}
            />
            <label htmlFor={existingSheetId}>Existing Worksheet</label>
          </div>
          <div className="pivot-dialog-range-row pivot-dialog-location-row">
            <label htmlFor={locationId}>Location:</label>
            <input
              id={locationId}
              type="text"
              disabled={placement !== 'existing'}
              value={locationText}
              onChange={event => {
                setError('');
                onChangeLocation(event.currentTarget.value);
              }}
            />
            <button
              type="button"
              className="refers-control-btn"
              aria-label="Pick Location"
              title="Select destination cell from worksheet"
              disabled={placement !== 'existing'}
              onClick={onPickLocation}
            >
              ↑
            </button>
          </div>
        </fieldset>

        {error && <p className="dialog-error" role="alert">{error}</p>}
        <DialogActions confirmLabel="OK" onCancel={onCancel} />
      </form>
    </dialog>
  );
}
