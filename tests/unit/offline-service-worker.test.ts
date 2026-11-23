import { createHash, webcrypto } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { runInNewContext } from 'node:vm';
import { MessageChannel } from 'node:worker_threads';
import { describe, expect, it } from 'vitest';

const template = await readFile(resolve('scripts/service-worker.template.js'), 'utf8');
const scope = 'https://workbook.test/ExcelLightWeb/';
const buildId = 'test-build-id';
const hash = (value: string | Uint8Array) => createHash('sha256').update(value).digest('hex');

function cacheUrl(value: string | URL | Request) {
  if (typeof value === 'string') return new URL(value, scope).href;
  return value instanceof URL ? value.href : value.url;
}

function createWorker({ peerDirty = false, quota = false } = {}) {
  const bytesByUrl = new Map<string, Uint8Array>();
  const offline = new Set<string>();
  let quotaPending = quota;
  const listeners = new Map<string, (event: any) => void>();
  const caches = new Map<string, Map<string, Response>>();
  let skipWaitingCount = 0;
  const runtimeA = new TextEncoder().encode('runtime wasm module A');
  const runtimeB = new TextEncoder().encode('runtime package B');
  const addAsset = (url: string, bytes: Uint8Array) => {
    bytesByUrl.set(new URL(url, scope).href, bytes);
    return { url, bytes: bytes.byteLength, sha256: hash(bytes) };
  };
  const shellAsset = addAsset('index.html', new TextEncoder().encode('<main>ExcelLight</main>'));
  const runtimeAssets = [
    addAsset('runtime/pyodide/314.0.8/pyodide.mjs', runtimeA),
    addAsset('runtime/pyodide/314.0.8/package.whl', runtimeB),
  ];
  const manifest = { schemaVersion: 1, buildId, runtimeFingerprint: hash('runtime-v1'), shell: [shellAsset], runtime: runtimeAssets };
  const manifestHref = new URL('offline-manifest.json', scope).href;
  const fetcher = async (input: RequestInfo | URL) => {
    const url = cacheUrl(input as string | URL | Request);
    if (offline.has(url)) throw new TypeError('Network is offline');
    if (url === manifestHref) return new Response(JSON.stringify(manifest));
    const bytes = bytesByUrl.get(url);
    if (!bytes) return new Response('missing', { status: 404 });
    return new Response(bytes.slice().buffer, { headers: { 'Content-Type': 'application/octet-stream' } });
  };
  const cacheStorage = {
    async open(name: string) {
      let cache = caches.get(name);
      if (!cache) { cache = new Map(); caches.set(name, cache); }
      return {
        async match(key: string | URL | Request) { return cache!.get(cacheUrl(key))?.clone(); },
        async put(key: string | URL | Request, response: Response) {
          if (quotaPending && name.includes('-python-')) {
            quotaPending = false;
            throw Object.assign(new Error('Storage quota exceeded'), { name: 'QuotaExceededError' });
          }
          cache!.set(cacheUrl(key), response.clone());
        },
        async delete(key: string | URL | Request) { return cache!.delete(cacheUrl(key)); },
      };
    },
    async keys() { return [...caches.keys()]; },
    async delete(name: string) { return caches.delete(name); },
  };
  const source = { id: 'initiator', url: scope };
  const peer = {
    id: 'peer', url: scope,
    postMessage(message: { type: string }, ports?: MessagePort[]) {
      if (message.type === 'CHECK_WORKBOOK_DIRTY') queueMicrotask(() => ports?.[0]?.postMessage({ dirty: peerDirty }));
    },
  };
  const clients = {
    async matchAll() { return [source, peer]; },
    async claim() {},
  };
  const self = {
    registration: { scope },
    navigator: { storage: { async estimate() { return undefined; } } },
    clients,
    addEventListener(type: string, listener: (event: any) => void) { listeners.set(type, listener); },
    async skipWaiting() { skipWaitingCount++; },
  };
  runInNewContext(template.replaceAll('__BUILD_ID__', buildId), {
    self, caches: cacheStorage, fetch: fetcher, crypto: webcrypto, Response, Request, URL,
    TextEncoder, Headers, MessageChannel, setTimeout, clearTimeout,
  });

  async function dispatch(type: string, data?: object, eventSource: unknown = source) {
    const pending: Promise<unknown>[] = [];
    const output: unknown[] = [];
    const port = { postMessage(value: unknown) { output.push(value); } };
    listeners.get(type)?.({ data, source: eventSource, ports: [port], waitUntil(value: Promise<unknown>) { pending.push(value); } });
    await Promise.all(pending);
    return output as { type: string; code?: string; ready?: boolean; cached?: number; total?: number; reason?: string }[];
  }

  return {
    manifest, offline, bytesByUrl, caches, dispatch,
    get skipWaitingCount() { return skipWaitingCount; },
    set peerDirty(value: boolean) { peerDirty = value; },
  };
}

async function install(worker: ReturnType<typeof createWorker>) {
  await worker.dispatch('install');
}

describe('production service worker cache contracts', () => {
  it('keeps an old runtime usable when a new cache is interrupted, then completes on retry', async () => {
    const worker = createWorker();
    await install(worker);
    worker.offline.add(new URL(worker.manifest.runtime[1]!.url, scope).href);

    const interrupted = await worker.dispatch('message', { type: 'PREPARE_RUNTIME' });
    expect(interrupted.at(-1)).toMatchObject({ type: 'CACHE_FAILED', code: 'offline-or-stale' });
    const partial = await worker.dispatch('message', { type: 'GET_STATUS' });
    expect(partial.at(-1)).toMatchObject({ type: 'STATUS', ready: false, cached: 1, total: 2 });

    worker.offline.clear();
    const resumed = await worker.dispatch('message', { type: 'PREPARE_RUNTIME' });
    expect(resumed.at(-1)).toMatchObject({ type: 'CACHE_COMPLETE', ready: true, cached: 2, total: 2 });
    expect((await worker.dispatch('message', { type: 'GET_STATUS' })).at(-1)).toMatchObject({ type: 'STATUS', ready: true });
  });

  it('reports quota and checksum failures without marking a partial cache ready', async () => {
    const worker = createWorker({ quota: true });
    await install(worker);
    const quota = await worker.dispatch('message', { type: 'PREPARE_RUNTIME' });
    expect(quota.at(-1)).toMatchObject({ type: 'CACHE_FAILED', code: 'storage-quota' });
    expect((await worker.dispatch('message', { type: 'GET_STATUS' })).at(-1)).toMatchObject({ type: 'STATUS', ready: false });

    const assetUrl = new URL(worker.manifest.runtime[0]!.url, scope).href;
    worker.bytesByUrl.set(assetUrl, new TextEncoder().encode('stale runtime bytes'));
    const stale = await worker.dispatch('message', { type: 'PREPARE_RUNTIME' });
    expect(stale.at(-1)).toMatchObject({ type: 'CACHE_FAILED', code: 'offline-or-stale' });
    expect((await worker.dispatch('message', { type: 'GET_STATUS' })).at(-1)).toMatchObject({ type: 'STATUS', ready: false });
  });

  it('defers an update while another same-scope tab has dirty workbook state', async () => {
    const worker = createWorker({ peerDirty: true });
    await install(worker);
    const deferred = await worker.dispatch('message', { type: 'ACTIVATE_UPDATE' });
    expect(deferred.at(-1)).toMatchObject({ type: 'UPDATE_DEFERRED' });
    expect(worker.skipWaitingCount).toBe(0);

    worker.peerDirty = false;
    const activated = await worker.dispatch('message', { type: 'ACTIVATE_UPDATE' });
    expect(activated.at(-1)).toMatchObject({ type: 'UPDATE_ACTIVATING' });
    expect(worker.skipWaitingCount).toBe(1);
  });
});
