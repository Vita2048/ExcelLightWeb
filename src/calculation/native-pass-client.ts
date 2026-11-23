import type { NativeCalculationResult, NativeWorkerReply, NativeWorkerRequest, WorkbookSnapshot } from '../application/workbook-protocol';
import { workbookProtocolVersion } from '../application/workbook-protocol';
import { documentLimits } from '../core/workbook';

class NativeWorkerError extends Error {}

/** Owns a calculation worker for one detached mixed generation. */
export class NativePassClient {
  #worker: Worker | undefined;
  #serial = 0;
  #acceptedRequestId: string | undefined;
  #active: { requestId: string; resolve: (reply: NativeCalculationResult) => void; reject: (error: Error) => void;
    signal: AbortSignal; abort: () => void; onCalculating?: (cellIds: readonly string[]) => void } | undefined;

  constructor(private readonly basis: WorkbookSnapshot, private readonly locale: string) {}

  calculate(document: WorkbookSnapshot['document'], targets: readonly string[], signal: AbortSignal,
    onCalculating?: (cellIds: readonly string[]) => void): Promise<NativeCalculationResult> {
    signal.throwIfAborted();
    if (this.#active) return Promise.reject(new Error('Native calculation worker is busy.'));
    const worker = this.#worker ??= this.#createWorker();
    const requestId = `mixed-native-${++this.#serial}`;
    const request: NativeWorkerRequest = { kind: 'calculate-native', protocolVersion: workbookProtocolVersion,
      requestId, sessionId: this.basis.sessionId, revision: this.basis.revision, generation: this.basis.generation,
      sourceFingerprint: this.basis.sourceFingerprint, runtimeFingerprint: 'native-formulas-v1', document,
      scope: { targets }, locale: this.locale, ...(this.#acceptedRequestId && { acceptedRequestId: this.#acceptedRequestId }) };
    return new Promise((resolve, reject) => {
      const abort = () => {
        if (this.#active?.requestId === requestId) this.#clear(new NativeWorkerError('Mixed calculation stopped.'));
        this.dispose();
      };
      this.#active = { requestId, resolve, reject, signal, abort, ...(onCalculating && { onCalculating }) };
      signal.addEventListener('abort', abort, { once: true });
      try { worker.postMessage(request); }
      catch (error) { this.#clear(error instanceof Error ? error : new Error(String(error))); this.dispose(); }
    });
  }

  dispose() {
    this.#worker?.terminate();
    this.#worker = undefined;
    this.#acceptedRequestId = undefined;
    if (this.#active) this.#clear(new NativeWorkerError('Native calculation worker stopped.'));
  }

  #createWorker() {
    const worker = new Worker(new URL('../workers/native-calculation.worker.ts', import.meta.url), { type: 'module' });
    worker.onmessage = ({ data }: MessageEvent<NativeWorkerReply>) => {
      const active = this.#active;
      if (!active || data?.requestId !== active.requestId) return;
      if (data.protocolVersion !== workbookProtocolVersion ||
          data.sessionId !== this.basis.sessionId || data.revision !== this.basis.revision ||
          data.generation !== this.basis.generation || data.sourceFingerprint !== this.basis.sourceFingerprint ||
          data.runtimeFingerprint !== 'native-formulas-v1') {
        this.#clear(new NativeWorkerError('Native calculation worker returned an invalid generation.'));
        this.dispose();
        return;
      }
      if (data.kind === 'native-calculating') {
        if (!Array.isArray(data.calculatingCellIds) || data.calculatingCellIds.length > documentLimits.formulas ||
            data.calculatingCellIds.some(id => typeof id !== 'string')) {
          this.#clear(new NativeWorkerError('Native calculation worker returned invalid progress.'));
          this.dispose();
          return;
        }
        active.onCalculating?.(data.calculatingCellIds);
        return;
      }
      if (data.kind !== 'native-calculated' || !Array.isArray(data.updates) || !Array.isArray(data.staleCellIds)) {
        this.#clear(new NativeWorkerError('Native calculation worker returned an invalid generation.'));
        this.dispose();
        return;
      }
      this.#acceptedRequestId = data.requestId;
      this.#clear(undefined, data);
    };
    worker.onerror = worker.onmessageerror = () => {
      this.#clear(new NativeWorkerError('Native calculation worker failed.'));
      this.dispose();
    };
    return worker;
  }

  #clear(error?: Error, reply?: NativeCalculationResult) {
    const active = this.#active;
    if (!active) return;
    this.#active = undefined;
    active.signal.removeEventListener('abort', active.abort);
    if (error) active.reject(error);
    else if (reply) active.resolve(reply);
    else active.reject(new NativeWorkerError('Native calculation worker returned no result.'));
  }
}

