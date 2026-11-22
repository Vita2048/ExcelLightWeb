import { useCallback, useRef, useState } from 'react';
import type { WorkbookSession } from '../../application/workbook-session';
import { columnLetters, quoteSheet } from '../../core/formula-references';
import type { SheetDocument } from '../../core/workbook';
import {
  acceptsPythonEditorMessage, currentPythonSource, pythonSourceCommand, samePythonTarget,
  type PythonAnchorTarget,
} from './python-editor-model';
import type { PythonEditorContext, PythonInsertRequest } from './PythonPanel';

export interface PythonEditorHost {
  readonly session: WorkbookSession;
  readonly changed: (compound: boolean) => void;
}

export function usePythonEditor({ session, changed }: PythonEditorHost) {
  const target = useRef<PythonAnchorTarget | null>(null);
  const contextId = useRef<string | null>(null);
  const draft = useRef<{ contextId: string; text: string } | null>(null);
  const pickable = useRef(false);
  const serial = useRef(0);
  const [context, setContext] = useState<PythonEditorContext | null>(null);
  const [insertRequest, setInsertRequest] = useState<PythonInsertRequest>();
  const [status, setStatus] = useState('');

  const open = useCallback((next: PythonAnchorTarget, source: string) => {
    const id = crypto.randomUUID();
    target.current = { ...next, source };
    contextId.current = id;
    draft.current = { contextId: id, text: source };
    pickable.current = false;
    setContext({ contextId: id, source, label: label(next), sheetId: next.sheetId, row: next.row, column: next.column });
    setStatus('');
  }, []);

  const clear = useCallback(() => {
    target.current = null;
    contextId.current = null;
    draft.current = null;
    pickable.current = false;
    setContext(null);
  }, []);

  function label(next: PythonAnchorTarget): string {
    return columnLetters(next.column) + next.row;
  }

  const flush = useCallback((): boolean => {
    const current = target.current;
    const pending = draft.current;
    if (!current || !pending || pending.contextId !== contextId.current) return true;
    if (pending.text === current.source) return true;
    try {
      const snapshot = session.capture();
      const live = currentPythonSource(snapshot, current);
      if (live !== current.source) {
        // A structural edit, undo or adoption replaced the authored source. The newer
        // workbook wins; never commit an older editor draft over it.
        if (live === undefined) clear();
        else open({ ...current, source: live }, live);
        return false;
      }
      const command = pythonSourceCommand(snapshot, current, pending.text);
      if (!command) return true;
      session.execute(command);
      target.current = { ...current, source: pending.text };
      draft.current = { contextId: pending.contextId, text: pending.text };
      setStatus('Source saved.');
      changed(false);
      return true;
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error));
      return false;
    }
  }, [changed, clear, open, session]);

  const sync = useCallback((next: PythonAnchorTarget | undefined) => {
    if (!next) {
      if (target.current) {
        flush();
        clear();
      }
      return;
    }
    const same = samePythonTarget(target.current ?? undefined, next);
    if (same) {
      const live = currentPythonSource(session.capture(), target.current!);
      if (live === undefined) {
        clear();
        return;
      }
      if (live !== target.current!.source) {
        // Undo, structure/name refactors and package adoption replace the source.
        open(next, live);
      }
      return;
    }
    flush();
    const snapshot = session.capture();
    const live = currentPythonSource(snapshot, next);
    if (live === undefined) {
      clear();
      return;
    }
    open({ ...next, source: live }, live);
  }, [clear, flush, open, session]);

  const receiveDraft = useCallback((incomingContext: string, text: string) => {
    if (!acceptsPythonEditorMessage(contextId.current ?? undefined, { contextId: incomingContext, text })) return;
    if (!target.current) return;
    draft.current = { contextId: incomingContext, text };
    if (text !== target.current.source) setStatus('');
  }, []);

  const setPickable = useCallback((next: boolean) => {
    pickable.current = next;
  }, []);

  const wantsReference = useCallback(() => {
    return pickable.current && target.current !== null;
  }, []);

  const insertReference = useCallback((sheet: SheetDocument, area: {
    top: number; left: number; bottom: number; right: number;
  }): boolean => {
    const current = target.current;
    if (!current || !pickable.current) return false;
    const reference = (sheet.id === current.sheetId ? '' : quoteSheet(sheet.name)) +
      columnLetters(area.left) + area.top +
      (area.top === area.bottom && area.left === area.right ? '' :
        ':' + columnLetters(area.right) + area.bottom);
    setInsertRequest({ contextId: contextId.current!, reference, serial: ++serial.current });
    return true;
  }, []);

  /** Persist a pending draft before a save, calculation, file or history barrier. */
  const barrier = useCallback(() => flush(), [flush]);

  /** Persist the draft and close the editor. Hiding the panel must not drop pending source. */
  const leave = useCallback((): boolean => {
    if (!flush()) return false;
    clear();
    return true;
  }, [flush, clear]);

  return {
    context, insertRequest, status, sync, receiveDraft, setPickable, wantsReference, insertReference,
    barrier, leave,
  };
}
