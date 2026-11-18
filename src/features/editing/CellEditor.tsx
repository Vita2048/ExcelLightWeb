import { useLayoutEffect, useRef, type KeyboardEvent } from 'react';
import type { AfterEditKey, EditSurface } from './useWorksheetEditing';

export interface CellEditorProps {
  readonly surface: EditSurface;
  readonly value: string;
  readonly label: string;
  readonly onStart: () => void;
  readonly onChange: (text: string) => void;
  readonly onCommit: (move?: AfterEditKey, reverse?: boolean) => void;
  readonly onBlur: () => void;
  readonly onCancel: () => void;
  readonly onComposition: (active: boolean) => void;
  readonly isComposing: () => boolean;
  readonly direct: boolean;
  readonly onKeepCaret: () => void;
}

export function CellEditor(props: CellEditorProps) {
  const node = useRef<HTMLTextAreaElement>(null);
  useLayoutEffect(() => {
    if (props.surface !== 'inline') return;
    const input = node.current!;
    input.focus({ preventScroll: true });
    input.setSelectionRange(input.value.length, input.value.length);
  }, [props.surface]);

  function keyDown(event: KeyboardEvent<HTMLTextAreaElement>) {
    const key = event.key;
    if (event.nativeEvent.isComposing || event.nativeEvent.keyCode === 229) return;
    if (props.isComposing()) {
      if (key === 'Enter' || key === 'Tab' || key === 'Escape') event.preventDefault();
      return;
    }
    if (key === 'Escape') {
      event.preventDefault();
      props.onCancel();
    } else if (key === 'Enter' || key === 'Tab') {
      event.preventDefault();
      props.onCommit(key, event.shiftKey);
    } else if (props.direct && key === 'F2') {
      event.preventDefault();
      props.onKeepCaret();
    } else if (props.direct && !event.shiftKey && !event.ctrlKey && !event.metaKey && !event.altKey &&
        (key === 'ArrowUp' || key === 'ArrowDown' || key === 'ArrowLeft' || key === 'ArrowRight')) {
      event.preventDefault();
      props.onCommit(key);
    }
  }

  return <textarea ref={node} rows={1} className={props.surface === 'inline' ? 'inline-cell-input' : 'formula-value'}
    aria-label={props.label} data-cell-editor={props.surface} spellCheck={false}
    value={props.value} onFocus={props.onStart} onChange={event => props.onChange(event.currentTarget.value)}
    onKeyDown={keyDown} onBlur={event => {
      if (!(event.relatedTarget instanceof HTMLElement) || !event.relatedTarget.dataset.cellEditor) props.onBlur();
    }} onCompositionStart={() => props.onComposition(true)}
    onCompositionEnd={event => { props.onChange(event.currentTarget.value); props.onComposition(false); }} />;
}
