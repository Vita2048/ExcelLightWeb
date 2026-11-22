import { useEffect, useRef } from 'react';

export function PythonTrustDialog({ onRunOnce, onTrustSession, onCancel }: {
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
    <p id="python-trust-description">This source will run locally in your browser using the bundled Python runtime. It runs only after you choose Run; opening or importing a workbook never runs its Python source. Python code can access the network if the source requests it.</p>
    <div className="python-trust-actions">
      <button type="button" className="file-primary" autoFocus onClick={onRunOnce}>Run once</button>
      <button type="button" className="file-secondary" onClick={onTrustSession}>Trust this session and Run</button>
      <button type="button" className="file-secondary" onClick={onCancel}>Cancel</button>
    </div>
  </dialog>;
}
