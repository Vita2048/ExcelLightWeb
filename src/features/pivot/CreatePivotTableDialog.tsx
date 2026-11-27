import { useId, useLayoutEffect, useRef, useState } from 'react';
import { DialogActions } from '../dialogs/DialogActions';
import { DialogTitleBar } from '../dialogs/DialogTitleBar';

export type PivotPlacement = 'new' | 'existing';

export function CreatePivotTableDialog({
  sourceText,
  placement,
  locationText,
  suspended = false,
  onChangeSource,
  onChangePlacement,
  onChangeLocation,
  onPickSource,
  onPickLocation,
  onSubmit,
  onCancel,
}: {
  readonly sourceText: string;
  readonly placement: PivotPlacement;
  readonly locationText: string;
  readonly suspended?: boolean;
  readonly onChangeSource: (value: string) => void;
  readonly onChangePlacement: (value: PivotPlacement) => void;
  readonly onChangeLocation: (value: string) => void;
  readonly onPickSource: () => void;
  readonly onPickLocation: () => void;
  readonly onSubmit: () => string | undefined;
  readonly onCancel: () => void;
}) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const sourceInputRef = useRef<HTMLInputElement>(null);
  const sourceId = useId();
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
    sourceInputRef.current?.focus();
  }, [suspended]);

  return (
    <dialog
      ref={dialogRef}
      className="paste-dialog create-pivot-dialog"
      aria-label="PivotTable from table or range"
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
          <h2>PivotTable from table or range</h2>
          <button
            type="button"
            className="format-dialog-close"
            aria-label="Close Create PivotTable"
            onClick={onCancel}
          >
            ×
          </button>
        </DialogTitleBar>

        <fieldset className="pivot-dialog-group">
          <legend>Select a table or range</legend>
          <div className="pivot-dialog-range-row">
            <label htmlFor={sourceId}>Table/Range:</label>
            <input
              ref={sourceInputRef}
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

        <fieldset className="pivot-dialog-group">
          <legend>Choose where you want the PivotTable to be placed</legend>
          <div className="pivot-dialog-radio">
            <input
              id={newSheetId}
              type="radio"
              name="pivot-placement"
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
              name="pivot-placement"
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
