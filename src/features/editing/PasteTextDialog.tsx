import { useId, useLayoutEffect, useRef, useState } from 'react';
import type { PasteMode } from '../../application/editing-commands';
import { DialogTitleBar } from '../dialogs/DialogTitleBar';
import { DialogActions } from '../dialogs/DialogActions';

export function PasteTextDialog({ mode, error, onSubmit, onClose }: {
  mode: PasteMode;
  error: string;
  onSubmit: (text: string) => void;
  onClose: () => void;
}) {
  const dialog = useRef<HTMLDialogElement>(null);
  const input = useRef<HTMLTextAreaElement>(null);
  const id = useId();
  const [text, setText] = useState('');
  useLayoutEffect(() => {
    const node = dialog.current!;
    if (!node.open) node.showModal();
    input.current?.focus();
  }, []);
  const title = mode === 'values' ? 'Paste Values' : mode === 'formulas' ? 'Paste Formulas' : 'Paste text';
  return <dialog className="paste-dialog" ref={dialog} aria-labelledby={id} onCancel={event => {
    event.preventDefault();
    onClose();
  }}>
    <form onSubmit={event => { event.preventDefault(); onSubmit(text); }}>
      <DialogTitleBar spaced><h2 id={id}>{title}</h2></DialogTitleBar>
      <label>Paste clipboard text here<textarea ref={input} value={text} rows={5}
        onChange={event => setText(event.currentTarget.value)} /></label>
      <p role="status">{error}</p>
      <DialogActions confirmLabel="Paste" onCancel={onClose} />
    </form>
  </dialog>;
}
