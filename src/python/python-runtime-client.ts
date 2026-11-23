import lock from './runtime-lock.json';
import { pythonRuntimeProtocolVersion, type PythonRuntimeInput, type PythonRuntimeOutput,
  type PythonRuntimeReply, type PythonRuntimeRequest } from './runtime-protocol';

export class PythonRuntimeError extends Error {
  constructor(message: string, readonly status: 'failed' | 'cancelled' | 'timedOut' | 'workerExited' = 'failed') {
    super(message);
    this.name = 'PythonRuntimeError';
  }
}

export class PythonRuntimeClient {
  private worker: Worker | undefined;
  private ready: Promise<string> | undefined;
  private requestId = 0;
  private workerGeneration = 0;
  private exchange: ((reply: PythonRuntimeReply | undefined, error?: PythonRuntimeError) => void) | undefined;

  get generation() { return this.workerGeneration; }
  get isReady() { return !!this.worker && !!this.ready; }

  start(signal?: AbortSignal, onProgress?: (message: string) => void): Promise<string> {
    if (this.ready) return this.ready;
    const worker = new Worker(new URL('../workers/python-runtime.worker.ts', import.meta.url), { type: 'module' });
    this.worker = worker;
    this.workerGeneration++;
    const requestId = this.nextRequestId();
    const runtimeUrl = new URL(`./runtime/pyodide/${lock.version}/`, document.baseURI).href;
    const request: PythonRuntimeRequest = { kind: 'init', protocolVersion: pythonRuntimeProtocolVersion,
      requestId, runtimeUrl };
    this.ready = this.exchangeWith(worker, request, 180_000, signal, onProgress).then(reply => {
      if (reply.kind === 'error') throw new PythonRuntimeError(reply.message);
      if (reply.kind !== 'ready' || reply.runtimeFingerprint.length > 128) {
        throw new PythonRuntimeError('Python runtime handshake was invalid.');
      }
      return reply.runtimeFingerprint;
    }).catch(error => {
      this.terminate(error instanceof PythonRuntimeError ? error : new PythonRuntimeError(message(error), 'workerExited'));
      throw error;
    });
    return this.ready;
  }

  async run(source: string, input: PythonRuntimeInput, signal?: AbortSignal, timeoutMs = 30_000,
    onProgress?: (message: string) => void): Promise<{
    readonly runtimeFingerprint: string;
    readonly output: PythonRuntimeOutput;
  }> {
    if (new TextEncoder().encode(source).length > 1024 * 1024) throw new PythonRuntimeError('Source exceeds 1 MiB.');
    const runtimeFingerprint = await this.start(signal, onProgress);
    onProgress?.('Running Python calculation…');
    const worker = this.worker;
    if (!worker) throw new PythonRuntimeError('Python runtime is stopped.', 'workerExited');
    const request: PythonRuntimeRequest = { kind: 'run', protocolVersion: pythonRuntimeProtocolVersion,
      requestId: this.nextRequestId(), source, input };
    const reply = await this.exchangeWith(worker, request, timeoutMs, signal);
    if (reply.kind === 'error') throw new PythonRuntimeError(reply.message, 'workerExited');
    if (reply.kind !== 'result') throw new PythonRuntimeError('Python worker returned an invalid reply.', 'workerExited');
    return { runtimeFingerprint, output: reply.output };
  }

  /** Redraws only a Figure retained by the already-running worker; it never starts Python execution. */
  async renderFigure(figureId: string, width: number, height: number, dpi: number,
    signal?: AbortSignal, timeoutMs = 120_000): Promise<PythonRuntimeOutput> {
    const worker = this.worker;
    const ready = this.ready;
    if (!worker || !ready) throw new PythonRuntimeError('The Python worker no longer retains this Figure.', 'workerExited');
    await ready;
    if (worker !== this.worker) throw new PythonRuntimeError('The Python worker changed before the Figure redraw.', 'workerExited');
    if (!/^[0-9A-F]{64}$/i.test(figureId) || !Number.isSafeInteger(width) || !Number.isSafeInteger(height) ||
        width < 1 || height < 1 || width > 8192 || height > 8192 || width * height > 16_000_000 ||
        !Number.isFinite(dpi) || dpi <= 0 || dpi > 384) {
      throw new PythonRuntimeError('Invalid figure render dimensions.');
    }
    const request: PythonRuntimeRequest = { kind: 'render-figure', protocolVersion: pythonRuntimeProtocolVersion,
      requestId: this.nextRequestId(), figureId: figureId.toUpperCase(), width, height, dpi };
    const reply = await this.exchangeWith(worker, request, timeoutMs, signal);
    if (reply.kind === 'error') throw new PythonRuntimeError(reply.message, 'workerExited');
    if (reply.kind !== 'result') throw new PythonRuntimeError('Python worker returned an invalid Figure redraw.', 'workerExited');
    return reply.output;
  }

  stop() {
    this.terminate(new PythonRuntimeError('Python run stopped.', 'cancelled'));
  }

  dispose() { this.stop(); }

  private nextRequestId() { return `python-${++this.requestId}`; }

  private exchangeWith(worker: Worker, request: PythonRuntimeRequest, timeoutMs: number,
    signal?: AbortSignal, onProgress?: (message: string) => void): Promise<PythonRuntimeReply> {
    if (this.exchange) return Promise.reject(new PythonRuntimeError('Python worker is busy.'));
    if (signal?.aborted) return Promise.reject(new PythonRuntimeError('Python run stopped.', 'cancelled'));
    return new Promise((resolve, reject) => {
      let finished = false;
      const finish = (reply?: PythonRuntimeReply, error?: PythonRuntimeError) => {
        if (finished) return;
        finished = true;
        clearTimeout(timer);
        signal?.removeEventListener('abort', abort);
        if (worker.onmessage === receive) worker.onmessage = null;
        if (worker.onerror === fail) worker.onerror = null;
        if (this.exchange === complete) this.exchange = undefined;
        if (error) reject(error);
        else if (reply) resolve(reply);
        else reject(new PythonRuntimeError('Python worker did not reply.', 'workerExited'));
      };
      const complete = (reply?: PythonRuntimeReply, error?: PythonRuntimeError) => finish(reply, error);
      const abort = () => {
        const error = new PythonRuntimeError('Python run stopped.', 'cancelled');
        finish(undefined, error);
        this.terminate(error);
      };
      const timer = setTimeout(() => {
        const error = new PythonRuntimeError(`Python timed out after ${timeoutMs} ms.`, 'timedOut');
        finish(undefined, error);
        this.terminate(error);
      }, timeoutMs);
      const receive = ({ data }: MessageEvent<PythonRuntimeReply>) => {
        if (data?.requestId !== request.requestId) {
          finish(undefined, new PythonRuntimeError('Python worker response identity did not match.', 'workerExited'));
          this.terminate(new PythonRuntimeError('Python worker response identity did not match.', 'workerExited'));
          return;
        }
        if (data.protocolVersion !== pythonRuntimeProtocolVersion) {
          finish(undefined, new PythonRuntimeError('Python worker protocol version did not match.', 'workerExited'));
          this.terminate(new PythonRuntimeError('Python worker protocol version did not match.', 'workerExited'));
          return;
        }
        if (data.kind === 'progress') {
          onProgress?.(data.message);
          return;
        }
        finish(data);
      };
      const fail = (event: ErrorEvent) => {
        const error = new PythonRuntimeError(event.message || 'Python worker exited.', 'workerExited');
        finish(undefined, error);
        this.terminate(error);
      };
      this.exchange = complete;
      worker.onmessage = receive;
      worker.onerror = fail;
      signal?.addEventListener('abort', abort, { once: true });
      try { worker.postMessage(request); }
      catch (error) {
        const failure = new PythonRuntimeError(message(error), 'workerExited');
        finish(undefined, failure);
        this.terminate(failure);
      }
    });
  }

  private terminate(reason: PythonRuntimeError) {
    const worker = this.worker;
    this.worker = undefined;
    this.ready = undefined;
    worker?.terminate();
    this.exchange?.(undefined, reason);
    this.exchange = undefined;
  }
}

function message(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
