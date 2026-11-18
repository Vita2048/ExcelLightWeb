import { useLayoutEffect, useRef, useState } from 'react';
import { DialogTitleBar } from '../dialogs/DialogTitleBar';
import { DialogActions } from '../dialogs/DialogActions';

export function MergeConfirmation({ onConfirm, onCancel }: {
  onConfirm: () => void; onCancel: () => void;
}) {
  const dialog = useRef<HTMLDialogElement>(null);
  const [error, setError] = useState('');
  useLayoutEffect(() => { if (!dialog.current!.open) dialog.current!.showModal(); }, []);
  return <dialog ref={dialog} className="paste-dialog merge-confirmation" aria-label="Merge cells" onCancel={event => {
    event.preventDefault(); onCancel();
  }}>
    <form onSubmit={event => {
      event.preventDefault();
      try { onConfirm(); } catch (error) { setError(error instanceof Error ? error.message : 'The merge failed.'); }
    }}>
      <DialogTitleBar><h2>Merge cells</h2>
        <button type="button" className="format-dialog-close" aria-label="Close Merge cells" onClick={onCancel}>×</button>
      </DialogTitleBar>
      <p>Each merged area keeps the leftmost value in its first populated row. Other values in that area will be discarded.</p>
      {error && <p role="alert">{error}</p>}
      <DialogActions onCancel={onCancel} />
    </form>
  </dialog>;
}
