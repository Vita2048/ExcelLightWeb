import { useId, useLayoutEffect, useMemo, useRef, useState, type KeyboardEvent, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import type { TextSelection } from '../formulas/reference-pick';
import { formulaHint } from './formula-hints';
import type { AfterEditKey, EditSurface } from './useWorksheetEditing';

export interface CellEditorHighlight {
  readonly start: number;
  readonly end: number;
  readonly color: string;
}

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
  readonly functionDraft?: boolean;
  readonly selectText?: TextSelection;
  readonly separator: ',' | ';';
  readonly textFormat: boolean;
  readonly highlights?: readonly CellEditorHighlight[] | undefined;
  readonly onKeepCaret: () => void;
}

export function CellEditor(props: CellEditorProps) {
  const node = useRef<HTMLTextAreaElement>(null);
  const backdrop = useRef<HTMLDivElement>(null);
  const tip = useRef<HTMLDivElement>(null);
  const tipId = useId();
  const [caret, setCaret] = useState<number | null>(null);
  const hint = props.active && caret !== null && !props.textFormat ? formulaHint(props.value, caret, props.separator) : undefined;
  const activeHighlights = useMemo(
    () => (props.active && !props.textFormat && props.highlights?.length ? props.highlights : []),
    [props.active, props.textFormat, props.highlights],
  );

  useLayoutEffect(() => {
    if (props.surface !== 'inline' && !props.functionDraft) return;
    const input = node.current!;
    input.focus({ preventScroll: true });
    const position = input.value.length - (props.functionDraft ? 1 : 0);
    input.setSelectionRange(position, position);
    setCaret(input.selectionStart);
  }, [props.surface, props.functionDraft]);

  useLayoutEffect(() => {
    if (!props.selectText) return;
    const input = node.current!;
    input.focus({ preventScroll: true });
    input.setSelectionRange(props.selectText.start, props.selectText.end);
    setCaret(props.selectText.start);
  }, [props.selectText]);

  useLayoutEffect(() => {
    if (backdrop.current && node.current) {
      backdrop.current.scrollLeft = node.current.scrollLeft;
    }
  }, [props.value, activeHighlights, props.selectText]);

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

  const hasHighlights = activeHighlights.length > 0;
  const renderedTokens = useMemo<ReactNode>(() => {
    if (!hasHighlights) return null;
    const parts: ReactNode[] = [];
    let cursor = 0;
    const sorted = [...activeHighlights].sort((a, b) => a.start - b.start);
    for (let i = 0; i < sorted.length; i++) {
      const span = sorted[i]!;
      const start = Math.max(cursor, Math.min(span.start, props.value.length));
      const end = Math.max(start, Math.min(span.end, props.value.length));
      if (start > cursor) parts.push(props.value.slice(cursor, start));
      if (end > start) {
        parts.push(
          <span key={`${start}-${end}-${i}`} className="formula-highlight-token" style={{ color: span.color }}>
            {props.value.slice(start, end)}
          </span>,
        );
      }
      cursor = end;
    }
    if (cursor < props.value.length) parts.push(props.value.slice(cursor));
    return parts;
  }, [hasHighlights, activeHighlights, props.value]);

  return <>
    <div className={props.surface === 'inline' ? 'inline-cell-editor-wrap' : 'formula-editor-wrap'}>
      {props.surface === 'inline' && (
        <span key="sizer" className="inline-cell-sizer" aria-hidden="true">{props.value || ' '}</span>
      )}
      {hasHighlights && <div key="backdrop" ref={backdrop} className="formula-highlight-backdrop" aria-hidden="true">{renderedTokens}</div>}
      <textarea key="editor" ref={node} rows={1}
        className={(props.surface === 'inline' ? 'inline-cell-input' : 'formula-value') + (hasHighlights ? ' has-highlights' : '')}
        aria-label={props.label} aria-describedby={hint ? tipId : undefined} data-cell-editor={props.surface} spellCheck={false}
        value={props.value} onFocus={event => { props.onStart(); trackCaret(event.currentTarget); }}
        onSelect={event => trackCaret(event.currentTarget)}
        onScroll={event => { if (backdrop.current) backdrop.current.scrollLeft = event.currentTarget.scrollLeft; }}
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
    </div>
    {hint && createPortal(
      <div ref={tip} id={tipId} role="tooltip" className="formula-hint">
        {hint.name}(
        {hint.parameters.map((parameter, index) => <span key={parameter}>
          {index > 0 && props.separator + ' '}
          {index === hint.active ? <strong>{parameter}</strong> : parameter}
        </span>)}
        {hint.variadic && props.separator + ' …'})
      </div>,
      document.body,
    )}
  </>;
}
