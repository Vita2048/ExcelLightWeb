import type { NativeCalculationScope, NativeWorkerReply, NativeWorkerRequest, WorkbookSnapshot } from '../application/workbook-protocol';
import type { WorkbookSession } from '../application/workbook-session';
import { documentLimits } from '../core/workbook';

export type CalculationMode = 'automatic' | 'manual';
type ExplicitScope = 'workbook' | { readonly sheetId: string };
export interface CalculationStatus {
  readonly mode: CalculationMode;
  readonly pending: 'calculate' | 'observe' | false;
  readonly stale: ReadonlySet<string>;
  readonly error?: string;
}

export class NativeScheduler {
  #worker: Worker | undefined;
  #request: NativeWorkerRequest | undefined;
  #superseded = false;
  #acceptedRequestId: string | undefined;
  #timer?: ReturnType<typeof setTimeout>;
  #queued = false;
  #suspended = false;
  #explicit: ExplicitScope | undefined;
  #status: CalculationStatus = { mode: 'automatic', pending: false, stale: new Set() };

  constructor(private readonly session: WorkbookSession,
    private readonly publish: (snapshot: WorkbookSnapshot, status: CalculationStatus) => void,
    private readonly locale: string) {}

  changed(compound: boolean) {
    this.#schedule(this.#status.mode === 'manual' || this.#explicit ? 0 : compound || this.#request ? 220 : 75);
  }

  setMode(mode: CalculationMode) {
    if (mode === this.#status.mode) return;
    this.#status = { ...this.#status, mode };
    this.#schedule(0);
  }

  calculate(scope: ExplicitScope) {
    this.#explicit = scope;
    this.#schedule(0);
  }

  setSuspended(suspended: boolean) {
    if (this.#suspended === suspended) return;
    this.#suspended = suspended;
    if (suspended) {
      clearTimeout(this.#timer);
      this.session.cancelNativeCalculation();
      this.#superseded = !!this.#request;
      this.#queued = true;
      const { error: _error, ...current } = this.#status;
      this.#update({ ...current, pending: 'observe' });
    } else this.#schedule(0);
  }

  dispose() {
    clearTimeout(this.#timer);
    this.session.cancelNativeCalculation();
    this.#worker?.terminate();
    this.#worker = undefined;
  }

  #update(status: CalculationStatus) {
    this.#status = status;
    this.publish(this.session.capture(), status);
  }

  #schedule(delay: number) {
    this.session.cancelNativeCalculation();
    this.#superseded = !!this.#request;
    clearTimeout(this.#timer);
    if (this.#suspended) {
      this.#queued = true;
      const { error: _error, ...current } = this.#status;
      this.#update({ ...current, pending: 'observe' });
      return;
    }
    this.#queued = false;
    const { error: _error, ...current } = this.#status;
    this.#update({ ...current, pending: this.#explicit || current.mode === 'automatic' ? 'calculate' : 'observe' });
    this.#timer = setTimeout(() => {
      this.#queued = true;
      this.#send();
    }, delay);
  }

  #send() {
    if (!this.#queued || this.#request || this.#suspended) return;
    this.#queued = false;
    const explicit = this.#explicit;
    if (typeof explicit === 'object' && !this.session.active.sheets.some(sheet => sheet.id === explicit.sheetId)) {
      this.#explicit = undefined;
    }
    const scope: NativeCalculationScope = this.#explicit ?? (this.#status.mode === 'manual' ? 'observe' : 'automatic');
    try {
      if (!this.#worker) {
        const worker = new Worker(new URL('../workers/native-calculation.worker.ts', import.meta.url), { type: 'module' });
        this.#worker = worker;
        worker.onmessage = ({ data }: MessageEvent<NativeWorkerReply>) => {
          if (this.#worker === worker) this.#receive(data);
        };
        worker.onerror = worker.onmessageerror = () => {
          if (this.#worker === worker) this.#fail();
        };
      }
      this.#request = { ...this.session.beginNativeCalculation('native-formulas-v1'),
        scope, locale: this.locale, ...(this.#acceptedRequestId && { acceptedRequestId: this.#acceptedRequestId }) };
      this.#worker.postMessage(this.#request);
    } catch {
      this.#fail();
    }
  }

  #receive(reply: NativeWorkerReply) {
    const request = this.#request;
    if (!request || reply?.requestId !== request.requestId) return;
    this.#request = undefined;
    if (this.#superseded) {
      this.#superseded = false;
      this.#send();
      return;
    }
    try {
      if (!Array.isArray(reply.staleCellIds) || reply.staleCellIds.length > documentLimits.formulas ||
          reply.staleCellIds.some(id => typeof id !== 'string')) {
        this.#fail();
        return;
      }
      if (this.session.tryCommitNativeCalculation(reply)) {
        this.#acceptedRequestId = reply.requestId;
        if (request.scope === this.#explicit) this.#explicit = undefined;
        this.#update({ mode: this.#status.mode, pending: false, stale: new Set(reply.staleCellIds) });
      } else this.#fail();
    } catch {
      this.#fail();
      return;
    }
    this.#send();
  }

  #fail() {
    this.dispose();
    this.#request = undefined;
    this.#acceptedRequestId = undefined;
    this.#queued = false;
    this.#explicit = undefined;
    this.#update({ ...this.#status, pending: false, error: 'Calculation failed. Choose Calculate Now to retry.' });
  }
}
