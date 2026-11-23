const BUILD_ID = '__BUILD_ID__';
const SCOPE = new URL(self.registration.scope);
const SCOPE_KEY = encodeURIComponent(SCOPE.pathname);
const PREFIX = `excellight-offline-${SCOPE_KEY}-`;
const SHELL_CACHE = `${PREFIX}shell-${BUILD_ID}`;
const encoder = new TextEncoder();
let manifestPromise;

function manifestUrl() {
  return new URL('offline-manifest.json', SCOPE);
}

function runtimeCacheName(manifest) {
  return `${PREFIX}python-${manifest.runtimeFingerprint}`;
}

function readyMarkerUrl() {
  return new URL('__excellight_offline__/python-ready', SCOPE);
}

function inScope(url) {
  return url.origin === SCOPE.origin && url.pathname.startsWith(SCOPE.pathname);
}

async function digest(bytes) {
  const hash = await crypto.subtle.digest('SHA-256', bytes);
  return [...new Uint8Array(hash)].map(value => value.toString(16).padStart(2, '0')).join('');
}

async function checkedResponse(response, asset) {
  if (!response.ok) throw new Error(`Asset request failed (${response.status}) for ${asset.url}.`);
  const bytes = await response.arrayBuffer();
  if (bytes.byteLength !== asset.bytes || await digest(bytes) !== asset.sha256) {
    throw new Error(`Asset integrity check failed for ${asset.url}.`);
  }
  const headers = new Headers(response.headers);
  headers.set('X-ExcelLight-SHA256', asset.sha256);
  headers.set('Content-Length', String(bytes.byteLength));
  return new Response(bytes, { status: response.status, statusText: response.statusText, headers });
}

async function validateManifest(value) {
  if (!value || value.schemaVersion !== 1 || value.buildId !== BUILD_ID ||
      !Array.isArray(value.shell) || !Array.isArray(value.runtime) || !value.runtimeFingerprint) {
    throw new Error('The offline asset manifest is missing or stale. Reload the site while online.');
  }
  for (const asset of [...value.shell, ...value.runtime]) {
    const url = new URL(asset.url, SCOPE);
    if (!inScope(url) || !Number.isSafeInteger(asset.bytes) || asset.bytes < 0 ||
        !/^[a-f0-9]{64}$/.test(asset.sha256)) {
      throw new Error('The offline asset manifest contains an invalid path or checksum.');
    }
  }
  return value;
}

async function loadManifest() {
  if (!manifestPromise) {
    manifestPromise = (async () => {
      const shell = await caches.open(SHELL_CACHE);
      const cached = await shell.match(manifestUrl().href);
      if (cached) return validateManifest(await cached.json());
      const response = await fetch(manifestUrl(), { cache: 'no-store' });
      if (!response.ok) throw new Error('Offline manifest could not be loaded.');
      return validateManifest(await response.json());
    })();
  }
  return manifestPromise;
}

async function cacheAsset(cache, url, asset, { reload = false, verifyCached = false } = {}) {
  const cached = await cache.match(url.href);
  if (cached) {
    const expectedHeader = cached.headers.get('X-ExcelLight-SHA256');
    if (expectedHeader === asset.sha256 && Number(cached.headers.get('Content-Length')) === asset.bytes) {
      if (!verifyCached) return cached;
      try {
        const verified = await checkedResponse(cached.clone(), asset);
        await cache.put(url.href, verified.clone());
        return verified;
      } catch {
        await cache.delete(url.href);
      }
    }
    try {
      const verified = await checkedResponse(cached, asset);
      await cache.put(url.href, verified.clone());
      return verified;
    } catch {
      await cache.delete(url.href);
    }
  }
  const response = await fetch(url, { cache: reload ? 'reload' : 'default' });
  const verified = await checkedResponse(response, asset);
  await cache.put(url.href, verified.clone());
  return verified;
}

async function installShell() {
  const response = await fetch(manifestUrl(), { cache: 'reload' });
  if (!response.ok) throw new Error('Offline manifest could not be downloaded.');
  const manifest = await validateManifest(await response.clone().json());
  const cache = await caches.open(SHELL_CACHE);
  await cache.put(manifestUrl().href, response);
  for (const asset of manifest.shell) {
    await cacheAsset(cache, new URL(asset.url, SCOPE), asset, { reload: true });
  }
}

self.addEventListener('install', event => {
  event.waitUntil(installShell());
});

self.addEventListener('activate', event => {
  event.waitUntil((async () => {
    await loadManifest();
  })());
});

function assetMap(manifest) {
  return new Map([...manifest.shell, ...manifest.runtime].map(asset => [new URL(asset.url, SCOPE).href, asset]));
}

async function runtimeAssetResponse(request, asset, manifest) {
  const cache = await caches.open(runtimeCacheName(manifest));
  try {
    const cached = await cache.match(request.url);
    if (cached) {
      if (cached.headers.get('X-ExcelLight-SHA256') === asset.sha256 &&
          Number(cached.headers.get('Content-Length')) === asset.bytes) return cached;
      await cache.delete(request.url);
    }
    return await cacheAsset(cache, new URL(asset.url, SCOPE), asset);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return new Response(message, { status: 503, headers: { 'Content-Type': 'text/plain; charset=utf-8', 'X-ExcelLight-Offline-Miss': 'true' } });
  }
}

self.addEventListener('fetch', event => {
  const request = event.request;
  if (request.method !== 'GET') return;
  const url = new URL(request.url);
  if (!inScope(url)) return;
  event.respondWith((async () => {
    const manifest = await loadManifest();
    const asset = assetMap(manifest).get(url.href);
    if (asset && asset.url.startsWith('runtime/pyodide/')) {
      return runtimeAssetResponse(request, asset, manifest);
    }
    const shell = await caches.open(SHELL_CACHE);
    const cached = await shell.match(request);
    if (cached) return cached;
    if (request.mode === 'navigate') {
      const index = await shell.match(new URL('index.html', SCOPE).href);
      if (index) return index;
    }
    if (asset) {
      try { return await cacheAsset(shell, new URL(asset.url, SCOPE), asset); }
      catch { return new Response('This app asset is not cached. Reconnect and reload ExcelLight.', { status: 503 }); }
    }
    try { return await fetch(request); }
    catch { return new Response('This resource is unavailable offline.', { status: 503 }); }
  })());
});

function send(port, message) {
  try { port?.postMessage(message); } catch { /* The page may have navigated. */ }
}

async function getStatus(manifest) {
  const cache = await caches.open(runtimeCacheName(manifest));
  const marker = await cache.match(readyMarkerUrl().href);
  const saved = marker ? await marker.json().catch(() => undefined) : undefined;
  let cached = 0;
  for (const asset of manifest.runtime) {
    const response = await cache.match(new URL(asset.url, SCOPE).href);
    if (response && response.headers.get('X-ExcelLight-SHA256') === asset.sha256 &&
        Number(response.headers.get('Content-Length')) === asset.bytes) cached++;
  }
  const ready = saved?.runtimeFingerprint === manifest.runtimeFingerprint && cached === manifest.runtime.length;
  if (!ready && marker) await cache.delete(readyMarkerUrl().href);
  return { ready, cached, total: manifest.runtime.length, buildId: manifest.buildId };
}

function isQuotaError(error) {
  return error?.name === 'QuotaExceededError' || /quota|storage.*full/i.test(error?.message ?? '');
}

async function prepareRuntime(manifest, port) {
  const cache = await caches.open(runtimeCacheName(manifest));
  const estimate = await self.navigator.storage?.estimate?.().catch(() => undefined);
  const required = manifest.runtime.reduce((sum, asset) => sum + asset.bytes, 0);
  if (estimate?.quota !== undefined && estimate.usage !== undefined && estimate.quota - estimate.usage < required) {
    throw Object.assign(new Error('There is not enough browser storage to cache the complete Python runtime.'), { code: 'storage-quota' });
  }
  let completed = 0;
  for (const asset of manifest.runtime) {
    const url = new URL(asset.url, SCOPE);
    await cacheAsset(cache, url, asset, { verifyCached: true });
    completed++;
    send(port, { type: 'CACHE_PROGRESS', completed, total: manifest.runtime.length, current: asset.url });
  }
  await cache.put(readyMarkerUrl().href, new Response(JSON.stringify({ runtimeFingerprint: manifest.runtimeFingerprint })));
  const status = await getStatus(manifest);
  if (!status.ready) throw Object.assign(new Error('A Python runtime file is missing or stale. Retry while online.'), { code: 'stale-cache' });
  for (const name of await caches.keys()) {
    if (name.startsWith(`${PREFIX}python-`) && name !== runtimeCacheName(manifest)) await caches.delete(name);
  }
  return status;
}

function queryClientState(client) {
  return new Promise(resolve => {
    const channel = new MessageChannel();
    const timeout = setTimeout(() => { channel.port1.close(); resolve({ dirty: true, unavailable: true }); }, 1500);
    channel.port1.onmessage = event => {
      clearTimeout(timeout);
      channel.port1.close();
      resolve(event.data);
    };
    client.postMessage({ type: 'CHECK_WORKBOOK_DIRTY' }, [channel.port2]);
  });
}

async function activateIfSafe(source, port) {
  const clients = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
  const scopedOthers = clients.filter(client => client.id !== source?.id && client.url.startsWith(SCOPE.href));
  for (const client of scopedOthers) {
    const status = await queryClientState(client);
    if (status?.dirty || status?.unavailable) {
      send(port, { type: 'UPDATE_DEFERRED', reason: 'Another tab has unsaved workbook changes. Save or close it before updating.' });
      return;
    }
  }
  send(port, { type: 'UPDATE_ACTIVATING' });
  await self.skipWaiting();
}

async function cleanOldCaches(source) {
  const clients = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
  const scoped = clients.filter(client => client.url.startsWith(SCOPE.href));
  if (scoped.length > 1) return false;
  for (const name of await caches.keys()) {
    if (name.startsWith(`${PREFIX}shell-`) && name !== SHELL_CACHE) await caches.delete(name);
  }
  return true;
}

self.addEventListener('message', event => {
  const data = event.data;
  const port = event.ports?.[0];
  if (!data || typeof data.type !== 'string') return;
  event.waitUntil((async () => {
    try {
      const manifest = await loadManifest();
      if (data.type === 'GET_STATUS') send(port, { type: 'STATUS', ...(await getStatus(manifest)) });
      else if (data.type === 'PREPARE_RUNTIME') {
        const status = await prepareRuntime(manifest, port);
        send(port, { type: 'CACHE_COMPLETE', ...status });
      } else if (data.type === 'ACTIVATE_UPDATE') await activateIfSafe(event.source, port);
      else if (data.type === 'CLEAN_OLD_CACHES') send(port, { type: 'CACHES_CLEANED', cleaned: await cleanOldCaches(event.source) });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      const code = error?.code ?? (isQuotaError(error) ? 'storage-quota' : 'offline-or-stale');
      send(port, { type: 'CACHE_FAILED', code, message });
    }
  })());
});
