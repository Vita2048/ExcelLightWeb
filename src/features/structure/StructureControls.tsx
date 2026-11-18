import { useLayoutEffect, useRef, useState, type FormEvent } from 'react';
import type { Axis } from '../../core/axis-edit';
import type { SheetDocument } from '../../core/workbook';
import { DialogActions } from '../dialogs/DialogActions';

export type StructureAction = 'insert-cells' | 'delete-cells' | 'insert-row' | 'insert-column' |
  'delete-row' | 'delete-column' | 'insert-sheet' | 'delete-sheet' | 'rename-sheet' | 'place-sheet' |
  'size-row' | 'size-column' | 'fit-row' | 'fit-column' | 'hide-row' | 'hide-column' | 'show-row' | 'show-column';

export const structureLabels: Record<StructureAction, string> = {
  'insert-cells': 'Insert Cells…', 'delete-cells': 'Delete Cells…',
  'insert-row': 'Insert Sheet Rows', 'insert-column': 'Insert Sheet Columns',
  'delete-row': 'Delete Sheet Rows', 'delete-column': 'Delete Sheet Columns',
  'insert-sheet': 'Insert Sheet', 'delete-sheet': 'Delete Sheet', 'rename-sheet': 'Rename Sheet',
  'place-sheet': 'Move or Copy Sheet…', 'size-row': 'Row Height…', 'size-column': 'Column Width…',
  'fit-row': 'AutoFit Row Height', 'fit-column': 'AutoFit Column Width',
  'hide-row': 'Hide Rows', 'hide-column': 'Hide Columns', 'show-row': 'Unhide Rows', 'show-column': 'Unhide Columns',
};

export function StructureRibbon({ onAction }: { onAction: (action: StructureAction) => void }) {
  const groups: readonly [string, StructureAction[]][] = [
    ['Insert', ['insert-cells', 'insert-row', 'insert-column', 'insert-sheet']],
    ['Delete', ['delete-cells', 'delete-row', 'delete-column', 'delete-sheet']],
    ['Format', ['size-row', 'fit-row', 'size-column', 'fit-column', 'hide-row', 'hide-column',
      'show-row', 'show-column', 'rename-sheet', 'place-sheet']],
  ];
  return <fieldset><legend>Cells</legend>{groups.map(([label, actions]) =>
    <select key={label} aria-label={label} value="" onChange={event => onAction(event.target.value as StructureAction)}>
      <option value="" disabled>{label}</option>
      {actions.map(action => <option key={action} value={action}>{structureLabels[action]}</option>)}
    </select>)}</fieldset>;
}

export type DialogAction = 'insert-cells' | 'delete-cells' | 'delete-sheet' | 'place-sheet' | 'size-row' | 'size-column';
export interface StructureInput { readonly axis: Axis; readonly entire: boolean; readonly before: string; readonly copy: boolean; readonly size: number }
export function StructureDialog({ action, target, sheets, initialSize, onSubmit, onCancel }: {
  action: DialogAction;
  target: SheetDocument;
  sheets: readonly SheetDocument[];
  initialSize: number;
  onSubmit: (input: StructureInput) => string | undefined;
  onCancel: () => void;
}) {
  const ref = useRef<HTMLDialogElement>(null);
  const [error, setError] = useState('');
  const shift = action === 'insert-cells' || action === 'delete-cells';
  const place = action === 'place-sheet';
  const size = action === 'size-row' || action === 'size-column';
  const deleting = action === 'delete-cells';

  useLayoutEffect(() => {
    const dialog = ref.current!;
    if (!dialog.open) dialog.showModal();
    if (action === 'delete-sheet') dialog.querySelector<HTMLButtonElement>('.dialog-actions button[type="button"]')?.focus();
  }, [action]);
  function submit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const data = new FormData(event.currentTarget);
    const mode = String(data.get('shift') ?? '');
    setError(onSubmit({
      axis: mode.endsWith('row') ? 'row' : 'column', entire: mode.startsWith('entire'),
      before: String(data.get('before') ?? ''), copy: data.has('copy'), size: Number(data.get('size')),
    }) ?? '');
  }
  return <dialog ref={ref} className="paste-dialog structure-dialog" aria-label={structureLabels[action]}
    onCancel={event => { event.preventDefault(); onCancel(); }}>
    <form onSubmit={submit}>
      <h2>{structureLabels[action].replace('…', '')}</h2>
      {shift && <label>Cell movement<select name="shift" autoFocus>
        <option value="lane-column">Shift cells {deleting ? 'left' : 'right'}</option>
        <option value="lane-row">Shift cells {deleting ? 'up' : 'down'}</option>
        <option value="entire-row">Entire row</option><option value="entire-column">Entire column</option>
      </select></label>}
      {place && <>
        <label>Before worksheet<select name="before" defaultValue={target.id} autoFocus>
          {sheets.map(sheet => <option key={sheet.id} value={sheet.id}>{sheet.name}</option>)}
          <option value="">Move to end</option>
        </select></label>
        <label className="structure-checkbox"><input type="checkbox" name="copy" /> Create a copy</label>
      </>}
      {size && <label>Size in pixels<input name="size" type="number" min="4" max="4096" step="any" required autoFocus defaultValue={initialSize} /></label>}
      {action === 'delete-sheet' && <p>Delete {target.name} and its contents? References to it become #REF!.</p>}
      {error && <p role="alert">{error}</p>}
      <DialogActions confirmLabel={action === 'delete-sheet' ? 'Delete' : 'OK'} onCancel={onCancel} />
    </form>
  </dialog>;
}

export function StructureMenu({ x, y, actions, onAction, onClose }: {
  x: number; y: number; actions: readonly StructureAction[];
  onAction: (action: StructureAction) => void; onClose: () => void;
}) {
  const ref = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    const node = ref.current!;
    node.style.left = Math.max(0, Math.min(x, innerWidth - node.offsetWidth)) + 'px';
    node.style.top = Math.max(0, Math.min(y, innerHeight - node.offsetHeight)) + 'px';
    node.querySelector('button')?.focus();
    function dismiss(event: PointerEvent) {
      if (!node.contains(event.target as Node)) onClose();
    }
    document.addEventListener('pointerdown', dismiss);
    return () => document.removeEventListener('pointerdown', dismiss);
  }, [x, y]);
  return <div ref={ref} className="structure-menu" role="group" aria-label="Structure commands" onKeyDown={event => {
    if (event.key === 'Escape') { event.preventDefault(); onClose(); }
    if (event.key === 'ArrowUp' || event.key === 'ArrowDown') {
      event.preventDefault();
      const buttons = [...event.currentTarget.querySelectorAll('button')];
      const index = buttons.indexOf(document.activeElement as HTMLButtonElement);
      buttons[(index + buttons.length + (event.key === 'ArrowUp' ? -1 : 1)) % buttons.length]?.focus();
    }
  }}>
    {actions.map(action => <button key={action} type="button" onClick={() => onAction(action)}>{structureLabels[action]}</button>)}
  </div>;
}
