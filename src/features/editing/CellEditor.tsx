import { useId, useLayoutEffect, useRef, useState, type KeyboardEvent } from 'react';
import { createPortal } from 'react-dom';
import { formulaHint } from './formula-hints';
import type { AfterEditKey, EditSurface } from './useWorksheetEditing';

export interface CellEditorProps {
  readonly surface: EditSurface;
  readonly active: boolean;
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
  readonly textFormat: boolean;
  readonly onKeepCaret: () => void;
}

export function CellEditor(props: CellEditorProps) {
  const node = useRef<HTMLTextAreaElement>(null);
  const tip = useRef<HTMLDivElement>(null);
  const tipId = useId();
  const [caret, setCaret] = useState<number | null>(null);
  const hint = props.active && caret !== null && !props.textFormat ? formulaHint(props.value, caret) : undefined;

  useLayoutEffect(() => {
    if (props.surface !== 'inline') return;
    const input = node.current!;
    input.focus({ preventScroll: true });
    input.setSelectionRange(input.value.length, input.value.length);
    setCaret(input.selectionStart);
  }, [props.surface]);

  useLayoutEffect(() => {
    if (!hint) return;
    const input = node.current!;
    const popup = tip.current!;
    const clip = input.closest('.inline-editor-layer');
    function position() {
      const editor = input.getBoundingClientRect();
      const viewport = document.documentElement;
      const visible = (clip ?? viewport).getBoundingClientRect();
      popup.style.visibility = editor.bottom > visible.top && editor.top < visible.bottom &&
        editor.right > visible.left && editor.left < visible.right ? 'visible' : 'hidden';
      popup.style.left = Math.max(4, Math.min(editor.left, viewport.clientWidth - popup.offsetWidth - 4)) + 'px';
      const below = editor.bottom + 2;
      popup.style.top = (below + popup.offsetHeight <= viewport.clientHeight - 4
        ? below : Math.max(4, editor.top - popup.offsetHeight - 2)) + 'px';
    }
    position();
    let frame = 0;
    function schedule() {
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(position);
    }
    window.addEventListener('resize', schedule);
    window.addEventListener('scroll', schedule, true);
    return () => {
      cancelAnimationFrame(frame);
      window.removeEventListener('resize', schedule);
      window.removeEventListener('scroll', schedule, true);
    };
  }, [hint]);

  function trackCaret(input: HTMLTextAreaElement) {
    if (document.activeElement === input && !props.isComposing()) setCaret(input.selectionStart);
  }

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

  return <>
    <textarea ref={node} rows={1} className={props.surface === 'inline' ? 'inline-cell-input' : 'formula-value'}
      aria-label={props.label} aria-describedby={hint ? tipId : undefined} data-cell-editor={props.surface} spellCheck={false}
      value={props.value} onFocus={event => { props.onStart(); trackCaret(event.currentTarget); }}
      onSelect={event => trackCaret(event.currentTarget)}
      onChange={event => { props.onChange(event.currentTarget.value); trackCaret(event.currentTarget); }}
      onKeyDown={keyDown} onBlur={event => {
        setCaret(null);
        if (!(event.relatedTarget instanceof HTMLElement) || !event.relatedTarget.dataset.cellEditor) props.onBlur();
      }} onCompositionStart={() => { props.onComposition(true); setCaret(null); }}
      onCompositionEnd={event => {
        props.onChange(event.currentTarget.value);
        props.onComposition(false);
        trackCaret(event.currentTarget);
      }} />
    {hint && createPortal(
      <div ref={tip} id={tipId} role="tooltip" className="formula-hint">
        {hint.name}(
        {hint.parameters.map((parameter, index) => <span key={parameter}>
          {index > 0 && ', '}
          {index === hint.active ? <strong>{parameter}</strong> : parameter}
        </span>)}
        {hint.variadic && ', …'})
      </div>,
      document.body,
    )}
  </>;
}
