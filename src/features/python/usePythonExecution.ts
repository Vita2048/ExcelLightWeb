import { useCallback, useEffect, useRef, useState } from 'react';
import type { WorkbookSnapshot } from '../../application/workbook-protocol';
import type { WorkbookSession } from '../../application/workbook-session';
import type { PythonAnchorTarget } from './python-editor-model';
import { MixedPythonScheduler } from '../../calculation/mixed-scheduler';

export interface PythonExecutionState {
  readonly key: string;
  readonly running: boolean;
  readonly stdout: string;
  readonly stderr: string;
  readonly outputTruncated: boolean;
  readonly message: string;
  readonly error?: string;
}

export function usePythonExecution(session: WorkbookSession, locale: string,
  onSnapshot: (snapshot: WorkbookSnapshot) => void, pauseNative: (paused: boolean) => void) {
  const scheduler = useRef<MixedPythonScheduler | undefined>(undefined);
  const callbacks = useRef({ onSnapshot, pauseNative });
  callbacks.current = { onSnapshot, pauseNative };
  const [state, setState] = useState<PythonExecutionState>();

  useEffect(() => {
    const instance = new MixedPythonScheduler(session, locale, setState,
      snapshot => callbacks.current.onSnapshot(snapshot), paused => callbacks.current.pauseNative(paused));
    scheduler.current = instance;
    return () => {
      instance.dispose();
      if (scheduler.current === instance) scheduler.current = undefined;
    };
  }, [session, locale]);

  const run = useCallback((target: PythonAnchorTarget, _staleFormulaIds: ReadonlySet<string>) => {
    scheduler.current?.run(target.cellId, true);
  }, []);
  const runAll = useCallback(() => scheduler.current?.runAll(), []);
  const stop = useCallback(() => scheduler.current?.stop(), []);
  const changed = useCallback((compound: boolean) => scheduler.current?.changed(compound) ?? false, []);
  const setAutomatic = useCallback((automatic: boolean) => scheduler.current?.setAutomatic(automatic), []);
  const adopted = useCallback(() => scheduler.current?.adopted(), []);
  const historyRestored = useCallback(() => scheduler.current?.historyRestored(), []);

  return { state, run, runAll, stop, changed, setAutomatic, adopted, historyRestored };
}
