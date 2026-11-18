import { useLayoutEffect, useRef, useState, type FormEvent } from 'react';
import type { Axis } from '../../core/axis-edit';
import type { SheetDocument } from '../../core/workbook';
import { RibbonFlyout, type RibbonItem } from '../shell/RibbonFlyout';
import { DialogTitleBar } from '../dialogs/DialogTitleBar';
import { DialogActions } from '../dialogs/DialogActions';

export type StructureAction = 'insert-cells' | 'delete-cells' | 'insert-row' | 'insert-column' |
  'delete-row' | 'delete-column' | 'insert-sheet' | 'delete-sheet' | 'rename-sheet' | 'place-sheet' |
  'format-cells' | 'size-row' | 'size-column' | 'fit-row' | 'fit-column' | 'hide-row' | 'hide-column' | 'show-row' | 'show-column';

export const structureLabels: Record<StructureAction, string> = {
  'format-cells': 'Format Cells…',
  'insert-cells': 'Insert Cells…', 'delete-cells': 'Delete Cells…',
  'insert-row': 'Insert Sheet Rows', 'insert-column': 'Insert Sheet Columns',
  'delete-row': 'Delete Sheet Rows', 'delete-column': 'Delete Sheet Columns',
  'insert-sheet': 'Insert Sheet', 'delete-sheet': 'Delete Sheet', 'rename-sheet': 'Rename Sheet',
  'place-sheet': 'Move or Copy Sheet…', 'size-row': 'Row Height…', 'size-column': 'Column Width…',
  'fit-row': 'AutoFit Row Height', 'fit-column': 'AutoFit Column Width',
  'hide-row': 'Hide Rows', 'hide-column': 'Hide Columns', 'show-row': 'Unhide Rows', 'show-column': 'Unhide Columns',
};

export function StructureRibbon({ onAction }: { onAction: (action: StructureAction) => void }) {
  function entry(action: StructureAction, group?: string): RibbonItem {
    const label = structureLabels[action];
    return { label, icon: label.replace('…', '...'), ...(group && { group }), select: () => onAction(action) };
  }
  return <fieldset className="cells-ribbon"><legend>Cells</legend>
    <RibbonFlyout label="Insert" icon="Insert" large items={(['insert-cells', 'insert-row', 'insert-column', 'insert-sheet'] as const).map(action => entry(action))} />
    <RibbonFlyout label="Delete" icon="Delete" large items={(['delete-cells', 'delete-row', 'delete-column', 'delete-sheet'] as const).map(action => entry(action))} />
    <RibbonFlyout label="Format" icon="Format" large items={[
      ...(['size-row', 'fit-row', 'size-column', 'fit-column'] as const).map(action => entry(action, 'Cell Size')),
      { label: 'Hide & Unhide', group: 'Visibility', children: (['hide-row', 'hide-column', 'show-row', 'show-column'] as const).map(action => entry(action)) },
      entry('rename-sheet', 'Organize Sheets'), entry('place-sheet', 'Organize Sheets'), entry('format-cells'),
    ]} />
  </fieldset>;
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
  const title = shift ? deleting ? 'Delete' : 'Insert' : structureLabels[action].replace('…', '');

  useLayoutEffect(() => {
    const dialog = ref.current!;
    if (!dialog.open) dialog.showModal();
    if (action === 'insert-cells' || action === 'delete-cells') dialog.querySelector<HTMLInputElement>('[name="shift"]')?.focus();
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
  return <dialog ref={ref} className={'paste-dialog structure-dialog' + (shift ? ' shift-dialog' : '')} aria-label={title}
    onCancel={event => { event.preventDefault(); onCancel(); }}>
    <form onSubmit={submit}>
      {shift ? <DialogTitleBar><h2>{title}</h2>
        <button type="button" className="format-dialog-close" aria-label={'Close ' + title} onClick={onCancel}>×</button>
      </DialogTitleBar> : <DialogTitleBar spaced><h2>{title}</h2></DialogTitleBar>}
      {shift && <fieldset className="shift-options" aria-label="Cell movement">
        {[['lane-column', 'Shift cells ' + (deleting ? 'left' : 'right')], ['lane-row', 'Shift cells ' + (deleting ? 'up' : 'down')],
          ['entire-row', 'Entire row'], ['entire-column', 'Entire column']].map(([value, label], index) =>
          <label key={value}><input type="radio" name="shift" value={value} defaultChecked={index === 0} />{label}</label>)}
      </fieldset>}
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
