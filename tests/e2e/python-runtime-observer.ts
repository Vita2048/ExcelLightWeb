import { expect, type Page } from '@playwright/test';

export interface PythonRuntimeTrace {
  initRequests: number;
  runRequests: number;
  readyReplies: number;
  progressMessages: string[];
  errors: string[];
}

export async function observePythonRuntime(page: Page) {
  await page.addInitScript(() => {
    const trace: PythonRuntimeTrace = {
      initRequests: 0, runRequests: 0, readyReplies: 0, progressMessages: [], errors: [],
    };
    (window as Window & { __pythonRuntimeTrace?: PythonRuntimeTrace }).__pythonRuntimeTrace = trace;
    const postMessage = Worker.prototype.postMessage as (...args: unknown[]) => void;
    Worker.prototype.postMessage = function(this: Worker, ...args: unknown[]) {
      const request = args[0] as { protocolVersion?: unknown; kind?: unknown } | undefined;
      if (request?.protocolVersion === 1 && request.kind === 'init') {
        trace.initRequests++;
        this.addEventListener('message', event => {
          const reply = (event as MessageEvent<{ kind?: string; message?: string }>).data;
          if (reply?.kind === 'ready') trace.readyReplies++;
          if (reply?.kind === 'progress' && reply.message) trace.progressMessages.push(reply.message);
          if (reply?.kind === 'error' && reply.message) trace.errors.push(reply.message);
        });
      } else if (request?.protocolVersion === 1 && request.kind === 'run') {
        trace.runRequests++;
      }
      return Reflect.apply(postMessage, this, args);
    } as typeof Worker.prototype.postMessage;
  });
}

export async function readPythonRuntimeTrace(page: Page): Promise<PythonRuntimeTrace> {
  return page.evaluate(() => {
    const trace = (window as Window & { __pythonRuntimeTrace?: PythonRuntimeTrace }).__pythonRuntimeTrace;
    return trace ? { ...trace, progressMessages: [...trace.progressMessages], errors: [...trace.errors] } : {
      initRequests: 0, runRequests: 0, readyReplies: 0, progressMessages: [], errors: [],
    };
  });
}

export async function waitForPythonRuntime(page: Page) {
  await expect.poll(async () => (await readPythonRuntimeTrace(page)).readyReplies, { timeout: 240_000 }).toBe(1);
  const trace = await readPythonRuntimeTrace(page);
  expect(trace.initRequests).toBe(1);
  expect(trace.errors).toEqual([]);
  return trace;
}
