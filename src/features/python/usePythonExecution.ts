import { useCallback, useEffect, useRef, useState } from 'react';
import type { WorkbookSession } from '../../application/workbook-session';
import { findCell, type PythonBlock } from '../../core/workbook';
import { applyPythonResult } from '../../core/python-spills';
import { buildPythonInputs, type PythonInputAnchor } from '../../python/python-inputs';
import { PythonRuntimeClient, PythonRuntimeError } from '../../python/python-runtime-client';
import { pythonBlockFromOutput } from '../../python/python-results';
import type { PythonRuntimeOutput } from '../../python/runtime-protocol';
import type { PythonAnchorTarget } from './python-editor-model';

export interface PythonExecutionState {
  readonly key: string;
  readonly running: boolean;
  readonly stdout: string;
  readonly stderr: string;
  readonly error?: string;
  readonly outputTruncated: boolean;
  readonly message: string;
}

function keyOf(target: PythonAnchorTarget) { return target.sheetId + ':' + target.cellId; }
function errorMessage(error: unknown) { return error instanceof Error ? error.message : String(error); }

export function usePythonExecution(session: WorkbookSession, changed: () => void) {
  const client = useRef<PythonRuntimeClient | undefined>(undefined);
  const abort = useRef<AbortController | undefined>(undefined);
  const requestSerial = useRef(0);
  const [state, setState] = useState<PythonExecutionState>();

  useEffect(() => () => {
    abort.current?.abort();
    client.current?.dispose();
  }, []);

  const stop = useCallback(() => {
    abort.current?.abort();
    client.current?.stop();
  }, []);

  const run = useCallback(async (target: PythonAnchorTarget, staleFormulaIds: ReadonlySet<string>) => {
    if (abort.current) return;
    const request = ++requestSerial.current;
    const key = keyOf(target);
    const basis = session.capture();
    const sheet = basis.document.sheets.find(item => item.id === target.sheetId);
    const cell = sheet && findCell(sheet, { row: target.row, column: target.column });
    if (!cell || cell.id !== target.cellId || cell.content.kind !== 'python' || cell.content.blockId !== target.blockId || !cell.content.block) {
      setState({ key, running: false, stdout: '', stderr: '', outputTruncated: false, message: 'Python anchor is no longer current.' });
      return;
    }
    const block = cell.content.block;
    const anchor: PythonInputAnchor = { sheetId: target.sheetId, row: target.row, column: target.column,
      blockId: target.blockId, source: cell.content.source };
    const controller = new AbortController();
    abort.current = controller;
    setState({ key, running: true, stdout: '', stderr: '', outputTruncated: false, message: 'Starting local Python runtime…' });

    const commit = (next: PythonBlock, label: string) => {
      if (request !== requestSerial.current) return false;
      const document = applyPythonResult(basis.document, { sheetId: target.sheetId, row: target.row, column: target.column }, next,
        () => crypto.randomUUID());
      const committed = session.tryExecute(basis, { kind: 'structure', document, label });
      if (committed) changed();
      return committed;
    };

    try {
      const built = buildPythonInputs(basis, anchor, staleFormulaIds);
      const runtime = client.current ??= new PythonRuntimeClient();
      setState({ key, running: true, stdout: '', stderr: '', outputTruncated: false, message: 'Running Python locally…' });
      const result = await runtime.run(anchor.source, built.input, controller.signal);
      const next = pythonBlockFromOutput(block, result.output, built.fingerprint, result.runtimeFingerprint);
      if (!commit(next, 'Run Python')) {
        setState({ key, running: false, stdout: result.output.stdout, stderr: result.output.stderr,
          outputTruncated: result.output.outputTruncated, message: 'Workbook changed while Python was running. Run it again.' });
        return;
      }
      setState({ key, running: false, stdout: result.output.stdout, stderr: result.output.stderr,
        ...(result.output.error && { error: result.output.error }), outputTruncated: result.output.outputTruncated,
        message: result.output.status === 'succeeded' ? 'Python run completed.' : 'Python run failed.' });
    } catch (error) {
      const status = error instanceof PythonRuntimeError ? error.status : 'failed';
      const message = errorMessage(error).slice(-32767);
      try {
        const next = { ...block, status, error: message };
        const committed = commit(next, 'Python run failed');
        setState({ key, running: false, stdout: '', stderr: '', error: message, outputTruncated: false,
          message: committed ? status === 'cancelled' ? 'Python run stopped.' : status === 'timedOut' ? 'Python run timed out.' : 'Python run failed.' :
            'Workbook changed while Python was running. Run it again.' });
      } catch (commitError) {
        const failure = errorMessage(commitError);
        setState({ key, running: false, stdout: '', stderr: '', error: message, outputTruncated: false, message: failure });
      }
    } finally {
      if (request === requestSerial.current) abort.current = undefined;
    }
  }, [changed, session]);

  return { state, run, stop };
}
