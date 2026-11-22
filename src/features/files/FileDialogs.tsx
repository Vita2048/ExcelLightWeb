import { useLayoutEffect, useRef, useState } from 'react';
import type { CalculationMode } from '../../calculation/native-scheduler';
import type { LocalWorkspacePreferences, RecoveryRecord } from '../../persistence/local';
import type { ExchangeIssue } from '../../persistence/xlsx';
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
export type ReportChoice = 'continue' | 'cancel';

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
  readonly onApply: (options: { readonly nativeCalculation: CalculationMode; readonly pythonAutomaticConsent: boolean }) => void;
  readonly onCancel: () => void;
}) {
  const dialog = useModalDialog();
  const [category, setCategory] = useState<'general' | 'formulas'>('general');
  const [mode, setMode] = useState<CalculationMode>(preferences.nativeCalculation);
  const [pythonAutomatic, setPythonAutomatic] = useState(preferences.pythonAutomaticConsent);
  return <dialog ref={dialog} className="paste-dialog file-modal options-dialog" aria-labelledby="options-title"
    onCancel={event => { event.preventDefault(); onCancel(); }}>
    <form onSubmit={event => { event.preventDefault(); onApply({ nativeCalculation: mode, pythonAutomaticConsent: pythonAutomatic }); }}>
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
            <fieldset className="option-group python-option-group">
              <legend>Python Calculation</legend>
              <label><input type="radio" value="automatic" checked={pythonAutomatic} name="python-calculation"
                onChange={() => setPythonAutomatic(true)} /> Automatic</label>
              <label><input type="radio" value="manual" checked={!pythonAutomatic} name="python-calculation"
                onChange={() => setPythonAutomatic(false)} /> Manual</label>
            </fieldset>
            <p>Automatic is a local preference for this browser. With it enabled, Python scripts run after relevant edits and when you open or import a workbook. Scripts run locally and can access the network if their code requests it. Choose Manual to run code only when you select Run.</p>
          </>}
        </section>
      </div>
      <DialogActions onCancel={onCancel} />
    </form>
  </dialog>;
}

export function SaveAsDialog({ initialName, purpose, onSave, onCancel }: {
  readonly initialName: string;
  readonly purpose: 'browser-save' | 'local-save-as';
  readonly onSave: (name: string) => void;
  readonly onCancel: () => void;
}) {
  const dialog = useModalDialog();
  const [name, setName] = useState(initialName);
  const [error, setError] = useState('');
  const title = purpose === 'browser-save' ? 'Save to this browser' : 'Save As';
  const description = purpose === 'browser-save' ?
    'This saves the workbook in this browser, not to a disk folder. Use Export when you want an .exlt file on disk.' :
    'This saves in this browser under the name shown. It does not write to a disk folder; use Export to download an .exlt file.';
  return <dialog ref={dialog} className="paste-dialog file-modal save-as-dialog" aria-labelledby="saveas-title"
    onCancel={event => { event.preventDefault(); onCancel(); }}>
    <form onSubmit={event => {
      event.preventDefault();
      const clean = name.trim();
      if (!clean || clean.length > 120 || /[\\/\0]/u.test(clean)) { setError('Enter a workbook name without path separators.'); return; }
      onSave(clean);
    }}>
      <DialogTitleBar><h2 id="saveas-title">{title}</h2>
        <button type="button" className="format-dialog-close" aria-label={'Close ' + title} onClick={onCancel}>×</button>
      </DialogTitleBar>
      <p className="save-as-explanation">{description}</p>
      <label className="save-as-name">Workbook name
        <input autoFocus value={name} maxLength={120} onChange={event => { setName(event.currentTarget.value); setError(''); }} />
      </label>
      <p className={error ? 'file-dialog-error' : 'file-dialog-hint'} role="status">{error || 'Choose the name shown in File → Home → Recent.'}</p>
      <DialogActions onCancel={onCancel} confirmLabel="Save" />
    </form>
  </dialog>;
}

export function RemoveWorkspaceDialog({ name, onChoice }: {
  readonly name: string;
  readonly onChoice: (remove: boolean) => void;
}) {
  const dialog = useModalDialog();
  return <dialog ref={dialog} className="paste-dialog file-modal remove-workspace-dialog" aria-labelledby="remove-workspace-title"
    onCancel={event => { event.preventDefault(); onChoice(false); }}>
    <DialogTitleBar><h2 id="remove-workspace-title">Remove {name} from this browser?</h2></DialogTitleBar>
    <p>This deletes the browser-local saved copy and its recovery data.</p>
    <p>It does not delete an .exlt file in Downloads or any file on disk.</p>
    <div className="dialog-actions">
      <button type="button" className="file-danger" autoFocus onClick={() => onChoice(true)}>Remove</button>
      <button type="button" onClick={() => onChoice(false)}>Cancel</button>
    </div>
  </dialog>;
}

export function XlsxReportDialog({ title, issues, onChoice }: {
  readonly title: string;
  readonly issues: readonly ExchangeIssue[];
  readonly onChoice: (choice: ReportChoice) => void;
}) {
  const dialog = useModalDialog();
  return <dialog ref={dialog} className="paste-dialog file-modal xlsx-report-dialog" aria-labelledby="xlsx-report-title"
    onCancel={event => { event.preventDefault(); onChoice('cancel'); }}>
    <DialogTitleBar><h2 id="xlsx-report-title">{title}</h2></DialogTitleBar>
    <p>Review the unsupported content below. Continue keeps everything ExcelLight supports; Cancel leaves your current workbook unchanged.</p>
    <ul className="xlsx-report-list">
      {issues.map(issue => <li key={issue.code}>
        <strong>{issue.code}</strong><span>{issue.detail}</span>
      </li>)}
    </ul>
    <div className="dialog-actions">
      <button type="button" className="file-primary" autoFocus onClick={() => onChoice('continue')}>Continue</button>
      <button type="button" onClick={() => onChoice('cancel')}>Cancel</button>
    </div>
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
