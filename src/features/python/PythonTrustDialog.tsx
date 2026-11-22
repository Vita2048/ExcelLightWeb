import { useEffect, useRef } from 'react';

export function PythonTrustDialog({ all = false, onRunOnce, onTrustSession, onCancel }: {
  readonly all?: boolean;
  readonly onRunOnce: () => void;
  readonly onTrustSession: () => void;
  readonly onCancel: () => void;
}) {
  const dialog = useRef<HTMLDialogElement>(null);
  useEffect(() => {
    if (!dialog.current?.open) dialog.current?.showModal();
  }, []);
  return <dialog ref={dialog} className="python-trust-dialog" aria-labelledby="python-trust-title"
    aria-describedby="python-trust-description" onCancel={event => { event.preventDefault(); onCancel(); }}>
    <h2 id="python-trust-title">Run Python in this workbook?</h2>
    <p id="python-trust-description">{all ? 'Every Python cell and its prerequisites in this workbook will run locally in your browser using the bundled Python runtime.' :
      'This source and its prerequisites will run locally in your browser using the bundled Python runtime.'} Python code can access the network if the source requests it.</p>
    <div className="python-trust-actions">
      <button type="button" className="file-primary" autoFocus onClick={onRunOnce}>{all ? 'Run All once' : 'Run once'}</button>
      <button type="button" className="file-secondary" onClick={onTrustSession}>{all ? 'Trust this session and Run All' : 'Trust this session and Run'}</button>
      <button type="button" className="file-secondary" onClick={onCancel}>Cancel</button>
    </div>
  </dialog>;
}
