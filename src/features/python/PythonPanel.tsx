import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from 'react';
import { xlReferenceInsertion, xlReferenceSlot } from './xl-pick';
import type * as Monaco from 'monaco-editor/editor/editor.api';
import type { PythonBlock } from '../../core/workbook';
import type { PythonExecutionState } from './usePythonExecution';

export interface PythonEditorContext {
  readonly contextId: string;
  readonly source: string;
  readonly label: string;
  readonly sheetId: string;
  readonly row: number;
  readonly column: number;
}

export interface PythonInsertRequest {
  readonly contextId: string;
  readonly reference: string;
  readonly serial: number;
}

export interface PythonEditorHandle {
  draft(): { readonly contextId: string; readonly text: string } | undefined;
  wantsReference(): boolean;
  insert(reference: string): boolean;
  focus(): void;
}

export interface PythonPanelProps {
  readonly width: number;
  readonly context: PythonEditorContext | null;
  readonly insertRequest: PythonInsertRequest | undefined;
  readonly status: string;
  readonly automatic: boolean;
  readonly pythonInputFresh: boolean;
  readonly block?: PythonBlock;
  readonly execution?: PythonExecutionState;
  readonly resultExecution?: PythonExecutionState;
  readonly placingFigure?: boolean;
  readonly figureStatus?: string;
  readonly onDraft: (contextId: string, text: string) => void;
  readonly onPickable: (pickable: boolean) => void;
  readonly onHide: () => void;
  readonly onRun: () => void;
  readonly onRunAll: () => void;
  readonly onStop: () => void;
  readonly onCopy: () => void;
  readonly onInsertValues: () => void;
  readonly onPlaceFigure: () => void;
  readonly onSaveImage: () => void;
  readonly onRebuildFigure: () => void;
}

type Editor = Monaco.editor.IStandaloneCodeEditor;

export const PythonPanel = forwardRef<PythonEditorHandle, PythonPanelProps>(function PythonPanel(props, ref) {
  const host = useRef<HTMLDivElement>(null);
  const editor = useRef<Editor | null>(null);
  const context = useRef<string | null>(null);
  const applying = useRef(false);
  const insertSerial = useRef(0);
  const propsRef = useRef(props);
  propsRef.current = props;
  const [state, setState] = useState<'loading' | 'ready' | 'failed'>('loading');

  function pickable() {
    const instance = editor.current;
    const model = instance?.getModel();
    const position = instance?.getPosition();
    if (!instance || !model || !position) return false;
    return !!xlReferenceSlot(model.getValue(), model.getOffsetAt(position));
  }

  function publishDraft() {
    const instance = editor.current;
    if (!instance || applying.current || !context.current) return;
    propsRef.current.onDraft(context.current, instance.getValue());
    propsRef.current.onPickable(pickable());
  }

  useImperativeHandle(ref, () => ({
    draft: () => {
      const instance = editor.current;
      return context.current && instance ? { contextId: context.current, text: instance.getValue() } : undefined;
    },
    wantsReference: pickable,
    insert(reference: string) {
      const instance = editor.current;
      const model = instance?.getModel();
      const position = instance?.getPosition();
      if (!instance || !model || !position) return false;
      const insertion = xlReferenceInsertion(model.getValue(), model.getOffsetAt(position), reference);
      if (!insertion) return false;
      const start = model.getPositionAt(insertion.start);
      const end = model.getPositionAt(insertion.end);
      applying.current = true;
      instance.executeEdits('xl-reference', [{
        range: { startLineNumber: start.lineNumber, startColumn: start.column,
          endLineNumber: end.lineNumber, endColumn: end.column },
        text: insertion.inserted,
        forceMoveMarkers: true,
      }]);
      applying.current = false;
      instance.setPosition(model.getPositionAt(insertion.caret));
      instance.focus();
      publishDraft();
      return true;
    },
    focus: () => editor.current?.focus(),
  }), []);

  useEffect(() => {
    let cancelled = false;
    void import('./monaco-runtime').then(({ monaco }) => {
      if (cancelled || !host.current) return;
      const instance = monaco.editor.create(host.current, {
        value: props.context?.source ?? '',
        language: 'python',
        theme: 'vs-dark',
        fontSize: 14,
        minimap: { enabled: false },
        automaticLayout: true,
        scrollBeyondLastLine: false,
        lineNumbers: 'on',
        ariaLabel: 'Python source',
        tabSize: 4,
        insertSpaces: true,
        padding: { top: 8 },
      });
      editor.current = instance;
      context.current = propsRef.current.context?.contextId ?? null;
      instance.onDidChangeModelContent(publishDraft);
      instance.onDidChangeCursorPosition(publishDraft);
      setState('ready');
      if (propsRef.current.context) instance.setValue(propsRef.current.context.source);
    }).catch(() => { if (!cancelled) setState('failed'); });
    return () => {
      cancelled = true;
      editor.current?.dispose();
      editor.current = null;
    };
  }, []);

  useEffect(() => {
    const instance = editor.current;
    if (!instance) return;
    context.current = props.context?.contextId ?? null;
    applying.current = true;
    instance.setValue(props.context?.source ?? '');
    applying.current = false;
    propsRef.current.onPickable(false);
  }, [props.context?.contextId]);

  useEffect(() => {
    if (!props.insertRequest || props.insertRequest.serial === insertSerial.current) return;
    if (props.insertRequest.contextId !== context.current) return;
    insertSerial.current = props.insertRequest.serial;
    editorRefInsert(props.insertRequest.reference);
  }, [props.insertRequest?.serial]);

  function editorRefInsert(reference: string) {
    const instance = editor.current;
    const model = instance?.getModel();
    const position = instance?.getPosition();
    if (!instance || !model || !position) return;
    const insertion = xlReferenceInsertion(model.getValue(), model.getOffsetAt(position), reference);
    if (!insertion) return;
    const start = model.getPositionAt(insertion.start);
    const end = model.getPositionAt(insertion.end);
    applying.current = true;
    instance.executeEdits('xl-reference', [{
      range: { startLineNumber: start.lineNumber, startColumn: start.column,
        endLineNumber: end.lineNumber, endColumn: end.column },
      text: insertion.inserted,
      forceMoveMarkers: true,
    }]);
    applying.current = false;
    instance.setPosition(model.getPositionAt(insertion.caret));
    instance.focus();
    publishDraft();
  }

  return (
    <aside className="python-panel" aria-label="Python workspace" style={{ width: props.width }}>
      <div className="python-panel-header">
        <span className="python-panel-title">Python</span>
        {props.context && <span className="python-anchor" aria-label="Python anchor">{props.context.label}</span>}
        <button type="button" className="ribbon-command python-hide" aria-label="Hide Python workspace"
          onClick={props.onHide}>Hide</button>
      </div>
      {state === 'loading' && <p className="python-loading" role="status">Loading Python editor…</p>}
      {state === 'failed' && <p className="python-loading" role="alert">The local Python editor could not be loaded.</p>}
      <div className="python-editor" ref={host} />
      {state === 'ready' && !props.context && <p className="python-empty">Select a Python cell or type =PY on the worksheet.</p>}
      {(props.context || props.execution?.running) && <div className="python-run-controls">
        {props.context && <button type="button" className="file-primary" onClick={props.onRun} disabled={!!props.execution?.running || state !== 'ready'}>Run</button>}
        {props.context && <button type="button" className="file-secondary" onClick={props.onRunAll} disabled={!!props.execution?.running || state !== 'ready'}>Run All</button>}
        <button type="button" className="file-secondary" onClick={props.onStop} disabled={!props.execution?.running}>Stop</button>
        <span className="python-status" role="status">{props.execution?.message ?? props.status}</span>
      </div>}
      {props.context && props.block && <PythonResult block={props.block} inputFresh={props.pythonInputFresh} execution={props.resultExecution}
        placing={!!props.placingFigure} figureStatus={props.figureStatus}
        onCopy={props.onCopy} onInsertValues={props.onInsertValues} onPlaceFigure={props.onPlaceFigure}
        onSaveImage={props.onSaveImage} onRebuildFigure={props.onRebuildFigure} />}
      <div className="python-panel-footer">
        {!props.context && <span className="python-status" role="status">{props.status}</span>}
        <span className="python-note">{props.automatic ? 'Automatic runs Python cells when a workbook opens and after relevant edits.' :
          'Manual runs only when you choose Run. Opening a workbook never executes its Python source.'}</span>
      </div>
    </aside>
  );
});

function displayCache(value: { readonly kind: string; readonly value?: string | number | boolean } | null | undefined) {
  if (!value) return '';
  if (value.kind === 'boolean') return value.value ? 'TRUE' : 'FALSE';
  return String(value.value ?? '');
}

function PythonResult({ block, inputFresh, execution, placing, figureStatus, onCopy, onInsertValues,
  onPlaceFigure, onSaveImage, onRebuildFigure }: {
  block: PythonBlock;
  inputFresh: boolean;
  execution: PythonExecutionState | undefined;
  placing: boolean;
  figureStatus: string | undefined;
  onCopy: () => void;
  onInsertValues: () => void;
  onPlaceFigure: () => void;
  onSaveImage: () => void;
  onRebuildFigure: () => void;
}) {
  const table = block.spillActive ? block.lastTable : null;
  const rows = table ? Math.min(table.rows, 500) : 0;
  const columns = table?.columns ?? 0;
  const successful = block.hasResult;
  return <section className="python-result" aria-label="Python result">
    <div className="python-result-header">
      <strong>Result</strong>
      {successful && <span>{inputFresh ? 'Latest run' : 'Previous successful result · stale'}</span>}
      {successful && <div className="python-result-actions">
        <button type="button" className="ribbon-command" onClick={onCopy}>Copy result</button>
        <button type="button" className="ribbon-command" onClick={onInsertValues} disabled={!!execution?.running}>Insert as Values</button>
      </div>}
    </div>
    {execution?.error && <pre className="python-error" role="alert">{execution.error}</pre>}
    {!execution?.error && block.error && block.status !== 'succeeded' && <pre className="python-error" role="alert">{block.error}</pre>}
    {!successful && !execution?.error && !block.error && <p className="python-empty-result">Run Python to see its result.</p>}
    {successful && block.lastFigure && <>
      <img className="python-figure" alt="Python generated figure"
        width={block.lastFigure.width} height={block.lastFigure.height}
        src={`data:image/png;base64,${block.lastFigure.pngBase64}`} />
      <div className="python-figure-actions" role="group" aria-label="Figure actions">
        <button type="button" className="ribbon-command" onClick={onSaveImage}>Save Image</button>
        <button type="button" className="ribbon-command" onClick={onPlaceFigure} disabled={!!execution?.running}>
          {placing ? 'Cancel placement' : 'Place on Sheet'}
        </button>
        {(figureStatus || !block.lastFigure.figureId) && <button type="button" className="ribbon-command"
          onClick={onRebuildFigure} disabled={!!execution?.running}>Rebuild Figure</button>}
      </div>
      {figureStatus && <p className="python-figure-status" role="status">{figureStatus}</p>}
    </>}
    {successful && table && <div className="python-table-scroll"><table className="python-result-table"><tbody>
      {Array.from({ length: rows }, (_, row) => <tr key={row}>{Array.from({ length: columns }, (_, column) => {
        const value = table.cells[row * columns + column];
        const Tag = row === 0 ? 'th' : 'td';
        return <Tag key={column}>{displayCache(value)}</Tag>;
      })}</tr>)}
    </tbody></table>{table.rows > rows && <p>Showing the first {rows} of {table.rows} rows.</p>}</div>}
    {successful && !table && !block.lastFigure && <output className="python-scalar-result">{displayCache(block.lastValue)}</output>}
    {successful && block.lastTable?.notices.map((notice, index) => <p className="python-note" key={index}>{notice}</p>)}
    {execution?.stdout && <details><summary>Python output</summary><pre>{execution.stdout}</pre></details>}
    {execution?.stderr && <details><summary>Python errors</summary><pre>{execution.stderr}</pre></details>}
    {execution?.outputTruncated && <p className="python-note">Python output was truncated at 1 MiB.</p>}
  </section>;
}
