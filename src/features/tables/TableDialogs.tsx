import { useEffect, useLayoutEffect, useRef, useState, type FormEvent } from 'react';
import type { CellDocument, WorkbookDocument, WorkbookName } from '../../core/workbook';
import { refersTo } from '../../application/table-commands';
import { DialogActions } from '../dialogs/DialogActions';
import { DialogTitleBar } from '../dialogs/DialogTitleBar';
import { formatCellValue } from '../../core/formats/cell-format';

export interface DefinitionDraft {
  readonly id?: string;
  readonly name: string;
  readonly comment: string;
  readonly extentText: string;
  readonly isTable: boolean;
  readonly isEdit: boolean;
  readonly hasHeaders?: boolean;
  readonly lockExtent?: boolean;
}

export function NameDefinitionDialog({
  definition,
  suspended = false,
  onPickRange,
  onSubmit,
  onCancel,
}: {
  readonly definition: DefinitionDraft;
  readonly suspended?: boolean;
  readonly onPickRange: (name: string, comment: string, extentText: string) => void;
  readonly onSubmit: (name: string, comment: string, extentText: string, headers: boolean) => string | undefined;
  readonly onCancel: () => void;
}) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const [error, setError] = useState('');
  const [extent, setExtent] = useState(definition.extentText);
  const [headers, setHeaders] = useState(definition.hasHeaders ?? false);

  useEffect(() => {
    setExtent(definition.extentText);
  }, [definition.extentText]);

  useEffect(() => {
    if (definition.hasHeaders !== undefined) setHeaders(definition.hasHeaders);
  }, [definition.hasHeaders]);

  const { isTable, isEdit } = definition;
  const lockExtent = definition.lockExtent === true;
  const title = isTable ? (isEdit ? 'Edit Table' : 'Create Table') : (isEdit ? 'Edit Name' : 'New Name');

  useLayoutEffect(() => {
    const dialog = dialogRef.current!;
    if (suspended) {
      if (dialog.open) dialog.close();
      return;
    }
    if (!dialog.open) dialog.showModal();
    dialog.querySelector<HTMLInputElement>('[name="name"]')?.focus();
  }, [title, suspended]);

  function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const data = new FormData(event.currentTarget);
    const name = String(data.get('name') ?? '').trim();
    const comment = String(data.get('comment') ?? '');
    const currentExtent = String(data.get('extent') ?? extent).trim();
    const err = onSubmit(name, comment, currentExtent, headers);
    if (err) setError(err);
  }

  return (
    <dialog
      ref={dialogRef}
      className="paste-dialog name-definition-dialog"
      aria-label={title}
      onCancel={event => { event.preventDefault(); if (!suspended) onCancel(); }}
    >
      <form onSubmit={handleSubmit}>
        <DialogTitleBar>
          <h2>{title}</h2>
          <button type="button" className="format-dialog-close" aria-label={'Close ' + title} onClick={onCancel}>×</button>
        </DialogTitleBar>

        <div className="name-dialog-fields">
          <label className="name-dialog-row">
            <span className="name-dialog-label">Name:</span>
            <input
              type="text"
              name="name"
              aria-label="Name"
              defaultValue={definition.name}
              maxLength={255}
              required
            />
          </label>

          <div className="name-dialog-row">
            <span className="name-dialog-label">Scope:</span>
            <span className="name-dialog-scope">Workbook</span>
          </div>

          <label className="name-dialog-row">
            <span className="name-dialog-label">Comment:</span>
            <textarea
              name="comment"
              aria-label="Comment"
              defaultValue={definition.comment}
              maxLength={32767}
              rows={3}
            />
          </label>

          <label className="name-dialog-row">
            <span className="name-dialog-label">{isTable ? 'Range:' : 'Refers to:'}</span>
            <div className="range-input-row">
              <input
                type="text"
                name="extent"
                aria-label={isTable ? 'Range' : 'Refers to'}
                value={extent}
                disabled={lockExtent}
                onChange={e => setExtent(e.target.value)}
                required
              />
              <button
                type="button"
                className="range-pick-btn"
                aria-label="Select range from worksheet"
                title="Select range from worksheet"
                disabled={lockExtent}
                onClick={event => {
                  const data = new FormData(event.currentTarget.form!);
                  onPickRange(String(data.get('name') ?? ''), String(data.get('comment') ?? ''), extent);
                }}
              >
                ...
              </button>
            </div>
          </label>

          {isTable && !isEdit && (
            <label className="structure-checkbox table-headers-check">
              <input
                type="checkbox"
                name="headers"
                aria-label="My table has headers"
                checked={headers}
                onChange={e => setHeaders(e.target.checked)}
              />
              <span>My table has headers</span>
            </label>
          )}

          {error && <p className="dialog-error" role="alert">{error}</p>}
        </div>

        <DialogActions confirmLabel="OK" onCancel={onCancel} />
      </form>
    </dialog>
  );
}

export function NameManagerDialog({
  document,
  suspended = false,
  onNew,
  onEdit,
  onDelete,
  onApplyExtent,
  onPickRange,
  pickedExtent,
  onConsumePick,
  onClose,
}: {
  readonly document: WorkbookDocument;
  readonly suspended?: boolean;
  readonly onNew: () => void;
  readonly onEdit: (name: WorkbookName) => void;
  readonly onDelete: (id: string) => void;
  readonly onApplyExtent: (id: string, extentText: string) => string | undefined;
  readonly onPickRange: (extentText: string) => void;
  readonly pickedExtent?: string | undefined;
  readonly onConsumePick: () => void;
  readonly onClose: () => void;
}) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const [selectedId, setSelectedId] = useState<string | undefined>(() => document.names[0]?.id);
  const [filterIndex, setFilterIndex] = useState(0); // 0: All names, 1: Named ranges, 2: Tables
  const [refersDraft, setRefersDraft] = useState('');
  const [error, setError] = useState('');

  const selectedName = document.names.find(n => n.id === selectedId);
  const skipReset = useRef(false);

  useEffect(() => {
    if (!pickedExtent || !selectedName) return;
    skipReset.current = true;
    setRefersDraft(pickedExtent);
    setError('');
    onConsumePick();
  }, [pickedExtent, onConsumePick, selectedName]);

  useEffect(() => {
    if (skipReset.current) {
      skipReset.current = false;
      return;
    }
    setRefersDraft(selectedName ? refersTo(document, selectedName.area) : '');
    setError('');
  }, [selectedId, document, selectedName]);

  useLayoutEffect(() => {
    const dialog = dialogRef.current!;
    if (suspended) {
      if (dialog.open) dialog.close();
      return;
    }
    if (!dialog.open) dialog.showModal();
  }, [suspended]);

  const extentLocked = !selectedName || selectedName.isTable === true;
  const filteredNames = document.names.filter(n => {
    if (filterIndex === 1) return !n.isTable;
    if (filterIndex === 2) return n.isTable;
    return true;
  }).sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));

  function cellDisplayValue(name: WorkbookName): string {
    if (name.area.invalid) return '#REF!';
    if (name.isTable) {
      return `${Math.max(0, name.area.rows - 1)} rows × ${name.area.columns} columns`;
    }
    if (name.area.rows === 1 && name.area.columns === 1) {
      const sheet = document.sheets.find(s => s.id === name.area.sheetId);
      const cell = sheet?.cells.find(c => c.row === name.area.row && c.column === name.area.column);
      return formatCellValue(cell).text;
    }
    return `${name.area.rows} × ${name.area.columns}`;
  }

  function handleApply() {
    if (!selectedName || selectedName.isTable) return;
    const err = onApplyExtent(selectedName.id, refersDraft);
    if (err) setError(err);
    else setError('');
  }

  function handleRevert() {
    if (selectedName) {
      setRefersDraft(refersTo(document, selectedName.area));
      setError('');
    }
  }

  return (
    <dialog
      ref={dialogRef}
      className="paste-dialog name-manager-dialog"
      aria-label="Name Manager"
      onCancel={event => { event.preventDefault(); if (!suspended) onClose(); }}
    >
      <DialogTitleBar>
        <h2>Name Manager</h2>
        <button type="button" className="format-dialog-close" aria-label="Close Name Manager" onClick={onClose}>×</button>
      </DialogTitleBar>

      <div className="name-manager-body">
        <div className="name-manager-toolbar">
          <div className="name-manager-actions">
            <button type="button" className="name-manager-btn" onClick={onNew}>New...</button>
            <button
              type="button"
              className="name-manager-btn"
              disabled={!selectedName}
              onClick={() => { if (selectedName) onEdit(selectedName); }}
            >
              Edit...
            </button>
            <button
              type="button"
              className="name-manager-btn"
              disabled={!selectedName || selectedName.isTable}
              onClick={() => { if (selectedName && !selectedName.isTable) onDelete(selectedName.id); }}
            >
              Delete
            </button>
          </div>
          <div className="name-manager-filter">
            <select
              aria-label="Filter names"
              value={filterIndex}
              onChange={e => setFilterIndex(Number(e.target.value))}
            >
              <option value={0}>All names</option>
              <option value={1}>Named ranges</option>
              <option value={2}>Tables</option>
            </select>
          </div>
        </div>

        <div className="name-manager-table-wrapper" tabIndex={0} aria-label="Workbook names and tables"
          onKeyDown={event => {
            if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
            event.preventDefault();
            const index = filteredNames.findIndex(item => item.id === selectedId);
            const next = filteredNames[index + (event.key === 'ArrowDown' ? 1 : -1)];
            if (next) setSelectedId(next.id);
          }}>
          <table className="name-manager-grid">
            <thead>
              <tr>
                <th style={{ width: '22%' }}>Name</th>
                <th style={{ width: '22%' }}>Value</th>
                <th style={{ width: '28%' }}>Refers To</th>
                <th style={{ width: '12%' }}>Scope</th>
                <th style={{ width: '16%' }}>Comment</th>
              </tr>
            </thead>
            <tbody>
              {filteredNames.map(item => {
                const isSelected = item.id === selectedId;
                const refText = refersTo(document, item.area);
                return (
                  <tr
                    key={item.id}
                    className={isSelected ? 'selected' : ''}
                    aria-selected={isSelected}
                    aria-label={`${item.name}, ${item.isTable ? 'table' : 'named range'}, ${refText}`}
                    onClick={() => setSelectedId(item.id)}
                    onDoubleClick={() => onEdit(item)}
                  >
                    <td>{(item.isTable ? '▦ ' : '') + item.name}</td>
                    <td>{cellDisplayValue(item)}</td>
                    <td title={refText}>{refText}</td>
                    <td>Workbook</td>
                    <td>{item.comment ?? ''}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        <div className="name-manager-refers-section">
          <span>Refers to:</span>
          <div className="name-manager-refers-row">
            <button
              type="button"
              className="refers-control-btn"
              aria-label="Cancel refers to edit"
              title="Cancel refers to edit"
              disabled={extentLocked}
              onClick={handleRevert}
            >
              ✕
            </button>
            <button
              type="button"
              className="refers-control-btn"
              aria-label="Apply refers to"
              title="Apply refers to"
              disabled={extentLocked}
              onClick={handleApply}
            >
              ✓
            </button>
            <input
              type="text"
              className="name-manager-refers-input"
              aria-label="Refers to"
              disabled={extentLocked}
              value={refersDraft}
              onChange={e => setRefersDraft(e.target.value)}
            />
            <button
              type="button"
              className="range-pick-btn"
              aria-label="Select range from worksheet"
              title="Select range from worksheet"
              disabled={extentLocked}
              onClick={() => { if (!extentLocked) onPickRange(refersDraft); }}
            >
              ...
            </button>
          </div>
          {error && <p className="dialog-error" role="alert">{error}</p>}
        </div>

        <div className="dialog-actions">
          <button type="button" onClick={onClose}>Close</button>
        </div>
      </div>
    </dialog>
  );
}

export function RangePickerWidget({
  initialRange,
  onConfirm,
  onCancel,
}: {
  readonly initialRange: string;
  readonly onConfirm: (range: string) => void;
  readonly onCancel: () => void;
}) {
  const [rangeText, setRangeText] = useState(initialRange);

  useEffect(() => {
    setRangeText(initialRange);
  }, [initialRange]);

  return (
    <div className="range-picker-floating-panel">
      <div className="range-picker-heading">Select a worksheet range</div>
      <div className="range-picker-input-row">
        <input
          type="text"
          className="range-picker-input"
          aria-label="Selected worksheet range"
          value={rangeText}
          onChange={e => setRangeText(e.target.value)}
        />
        <button type="button" className="range-picker-action-btn" onClick={() => onConfirm(rangeText)}>OK</button>
        <button type="button" className="range-picker-action-btn" onClick={onCancel}>Cancel</button>
      </div>
    </div>
  );
}
