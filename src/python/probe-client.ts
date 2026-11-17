import lock from './runtime-lock.json';
import type { ProbeRequest, ProbeResponse, Snapshot } from './probe-protocol';

export class PythonProbe {
  private worker: Worker | undefined;
  private cancel: ((error: Error) => void) | undefined;

  async start() {
    this.stop();
    this.worker = new Worker(new URL('../workers/python-probe.worker.ts', import.meta.url), { type: 'module' });
    const runtimeUrl = new URL(`runtime/pyodide/${lock.version}/`, document.baseURI).href;
    // Cold WASM compilation can dominate a probe run on slower browser engines.
    const reply = await this.request({ kind: 'init', runtimeUrl }, 180_000);
    if (reply.kind !== 'ready') throw new Error('Expected worker handshake');
    return reply.versions;
  }

  async run(source: string, snapshots: Record<string, Snapshot> = {}, timeoutMs = 30_000) {
    const reply = await this.request({ kind: 'run', source, snapshots }, timeoutMs);
    if (reply.kind !== 'result') throw new Error('Expected Python result');
    return reply.output;
  }

  async loadArrow() {
    const reply = await this.request({ kind: 'load-arrow' }, 180_000);
    if (reply.kind !== 'arrow-ready') throw new Error('Expected Arrow version');
    return reply.version;
  }

  stop(reason = new Error('Python worker stopped')) {
    this.worker?.terminate();
    this.worker = undefined;
    this.cancel?.(reason);
  }

  private request(message: ProbeRequest, timeoutMs: number): Promise<ProbeResponse> {
    const worker = this.worker;
    if (!worker) return Promise.reject(new Error('Python worker has stopped'));
    if (this.cancel) return Promise.reject(new Error('Python worker is busy'));
    return new Promise((resolve, reject) => {
      const finish = () => {
        clearTimeout(timer);
        worker.onmessage = null;
        worker.onerror = null;
        this.cancel = undefined;
      };
      const timer = window.setTimeout(() => this.stop(new Error(`Python timed out after ${timeoutMs} ms`)), timeoutMs);
      this.cancel = error => { finish(); reject(error); };
      worker.onerror = event => this.stop(new Error(event.message));
      worker.onmessage = ({ data }: MessageEvent<ProbeResponse>) => {
        finish();
        if (data.kind === 'error') reject(new Error(data.message));
        else resolve(data);
      };
      try {
        worker.postMessage(message);
      } catch (error) {
        finish();
        reject(error);
      }
    });
  }
}
