import { useLayoutEffect, useRef, useState } from 'react';
import type { CalculationMode } from '../../calculation/native-scheduler';
import type { LocalWorkspacePreferences, RecoveryRecord } from '../../persistence/local';
import { DialogActions } from '../dialogs/DialogActions';
import { DialogTitleBar } from '../dialogs/DialogTitleBar';

function useModalDialog() {
  const dialog = useRef<HTMLDialogElement>(null);
  useLayoutEffect(() => {
    if (!dialog.current?.open) dialog.current?.showModal();
  }, []);
  return dialog;
}

export type ReplacementChoice = 'save' | 'discard' | 'cancel';

export function UnsavedChangesDialog({ name, onChoice }: {
  readonly name: string;
  readonly onChoice: (choice: ReplacementChoice) => void;
}) {
  const dialog = useModalDialog();
  return <dialog ref={dialog} className="paste-dialog file-modal unsaved-dialog" aria-labelledby="unsaved-title"
    onCancel={event => { event.preventDefault(); onChoice('cancel'); }}>
    <DialogTitleBar><h2 id="unsaved-title">Save changes to {name}?</h2></DialogTitleBar>
    <p>Your changes will be lost if you don't save them.</p>
    <div className="dialog-actions">
      <button type="button" className="file-primary" onClick={() => onChoice('save')}>Save</button>
      <button type="button" onClick={() => onChoice('discard')}>Don't Save</button>
      <button type="button" onClick={() => onChoice('cancel')}>Cancel</button>
    </div>
  </dialog>;
}

export function RecoveryDialog({ recovery, onChoice }: {
  readonly recovery: RecoveryRecord;
  readonly onChoice: (choice: 'recover' | 'discard') => void;
}) {
  const dialog = useModalDialog();
  const time = new Intl.DateTimeFormat(undefined, { dateStyle: 'medium', timeStyle: 'short' }).format(recovery.savedAt);
  return <dialog ref={dialog} className="paste-dialog file-modal recovery-dialog" aria-labelledby="recovery-title"
    onCancel={event => { event.preventDefault(); onChoice('discard'); }}>
    <DialogTitleBar><h2 id="recovery-title">Recover your workbook?</h2></DialogTitleBar>
    <p>A local recovery copy of <strong>{recovery.name}</strong> was saved {time}.</p>
    <p>Recover opens that copy with its unsaved changes. Discard removes the recovery copy.</p>
    <div className="dialog-actions">
      <button type="button" className="file-primary" onClick={() => onChoice('recover')}>Recover</button>
      <button type="button" onClick={() => onChoice('discard')}>Discard</button>
    </div>
  </dialog>;
}

export function OptionsDialog({ preferences, onApply, onCancel }: {
  readonly preferences: LocalWorkspacePreferences;
  readonly onApply: (mode: CalculationMode) => void;
  readonly onCancel: () => void;
}) {
  const dialog = useModalDialog();
  const [category, setCategory] = useState<'general' | 'formulas'>('general');
  const [mode, setMode] = useState<CalculationMode>(preferences.nativeCalculation);
  return <dialog ref={dialog} className="paste-dialog file-modal options-dialog" aria-labelledby="options-title"
    onCancel={event => { event.preventDefault(); onCancel(); }}>
    <form onSubmit={event => { event.preventDefault(); onApply(mode); }}>
      <DialogTitleBar><h2 id="options-title">Excel-Lite Options</h2>
        <button type="button" className="format-dialog-close" aria-label="Close Excel-Lite Options" onClick={onCancel}>×</button>
      </DialogTitleBar>
      <div className="options-body">
        <div className="options-categories" role="tablist" aria-label="Options category" aria-orientation="vertical">
          {(['general', 'formulas'] as const).map(item => <button type="button" role="tab" key={item}
            aria-selected={category === item} onClick={() => setCategory(item)}>
            {item === 'general' ? 'General' : 'Formulas'}
          </button>)}
        </div>
        <section className="options-pane" role="tabpanel" aria-label={category === 'general' ? 'General' : 'Formulas'}>
          {category === 'general' ? <>
            <h3>Personalize your copy of Excel-Lite</h3>
            <label htmlFor="office-theme">Office Theme</label>
            <select id="office-theme" value="dark" disabled>
              <option value="dark">Dark</option>
            </select>
            <p>ExcelLight uses a dark-only theme.</p>
          </> : <>
            <h3>Calculation options</h3>
            <fieldset className="option-group">
              <legend>Workbook Calculation</legend>
              <label><input type="radio" name="native-calculation" value="automatic"
                checked={mode === 'automatic'} onChange={() => setMode('automatic')} /> Automatic</label>
              <label><input type="radio" name="native-calculation" value="manual"
                checked={mode === 'manual'} onChange={() => setMode('manual')} /> Manual</label>
            </fieldset>
            <fieldset className="option-group python-option-group" disabled>
              <legend>Python Calculation</legend>
              <label><input type="radio" checked={preferences.pythonAutomaticConsent} readOnly name="python-calculation" /> Automatic</label>
              <label><input type="radio" checked={!preferences.pythonAutomaticConsent} readOnly name="python-calculation" /> Manual</label>
            </fieldset>
            <p>Python calculation is not available yet. Automatic execution stays off until you explicitly choose it after Python is available.</p>
          </>}
        </section>
      </div>
      <DialogActions onCancel={onCancel} />
    </form>
  </dialog>;
}

export function SaveAsDialog({ initialName, onSave, onCancel }: {
  readonly initialName: string;
  readonly onSave: (name: string) => void;
  readonly onCancel: () => void;
}) {
  const dialog = useModalDialog();
  const [name, setName] = useState(initialName);
  const [error, setError] = useState('');
  return <dialog ref={dialog} className="paste-dialog file-modal save-as-dialog" aria-labelledby="saveas-title"
    onCancel={event => { event.preventDefault(); onCancel(); }}>
    <form onSubmit={event => {
      event.preventDefault();
      const clean = name.trim();
      if (!clean || clean.length > 120 || /[\\/\0]/u.test(clean)) { setError('Enter a workbook name without path separators.'); return; }
      onSave(clean);
    }}>
      <DialogTitleBar><h2 id="saveas-title">Save As</h2>
        <button type="button" className="format-dialog-close" aria-label="Close Save As" onClick={onCancel}>×</button>
      </DialogTitleBar>
      <label className="save-as-name">Workbook name
        <input autoFocus value={name} maxLength={120} onChange={event => { setName(event.currentTarget.value); setError(''); }} />
      </label>
      <p className={error ? 'file-dialog-error' : 'file-dialog-hint'} role="status">{error || 'Saved to this browser on this device.'}</p>
      <DialogActions onCancel={onCancel} confirmLabel="Save" />
    </form>
  </dialog>;
}

export function FileMessageDialog({ title, message, onClose }: {
  readonly title: string;
  readonly message: string;
  readonly onClose: () => void;
}) {
  const dialog = useModalDialog();
  return <dialog ref={dialog} className="paste-dialog file-modal file-message-dialog" aria-labelledby="file-message-title"
    onCancel={event => { event.preventDefault(); onClose(); }}>
    <DialogTitleBar><h2 id="file-message-title">{title}</h2>
      <button type="button" className="format-dialog-close" aria-label={'Close ' + title} onClick={onClose}>×</button>
    </DialogTitleBar>
    <p>{message}</p>
    <div className="dialog-actions"><button type="button" className="file-primary" autoFocus onClick={onClose}>Close</button></div>
  </dialog>;
}
