import { useEffect, useRef, useState } from 'react';
import runtimeLock from '../python/runtime-lock.json';

interface OfflineState {
  supported: boolean;
  shellReady: boolean;
  online: boolean;
  runtimeReady: boolean;
  cached: number;
  total: number;
  progress?: { readonly completed: number; readonly total: number; readonly current: string } | undefined;
  error?: string | undefined;
}

type ServiceWorkerReply = {
  readonly type: string;
  readonly ready?: boolean;
  readonly cached?: number;
  readonly total?: number;
  readonly completed?: number;
  readonly current?: string;
  readonly reason?: string;
  readonly code?: string;
  readonly message?: string;
};

function request(worker: ServiceWorker, message: object,
  onProgress?: (reply: ServiceWorkerReply) => void): Promise<ServiceWorkerReply> {
  return new Promise((resolve, reject) => {
    const channel = new MessageChannel();
    const timer = window.setTimeout(() => {
      channel.port1.close();
      reject(new Error('Offline storage did not respond. Retry while online.'));
    }, 180_000);
    channel.port1.onmessage = event => {
      const reply = event.data as ServiceWorkerReply;
      if (reply.type === 'CACHE_PROGRESS') {
        onProgress?.(reply);
        return;
      }
      window.clearTimeout(timer);
      channel.port1.close();
      resolve(reply);
    };
    worker.postMessage(message, [channel.port2]);
  });
}

const initialState = (): OfflineState => ({
  supported: 'serviceWorker' in navigator,
  shellReady: false,
  online: navigator.onLine,
  runtimeReady: false,
  cached: 0,
  total: 0,
});

export function OfflineStatus({ dirty, onSaveRecovery }: {
  readonly dirty: boolean;
  readonly onSaveRecovery: () => Promise<boolean>;
}) {
  const [state, setState] = useState(initialState);
  const [panelOpen, setPanelOpen] = useState(false);
  const [update, setUpdate] = useState<ServiceWorker | undefined>();
  const [updateBusy, setUpdateBusy] = useState(false);
  const [updateHidden, setUpdateHidden] = useState(false);
  const dirtyRef = useRef(dirty);
  const registrationRef = useRef<ServiceWorkerRegistration | undefined>(undefined);
  dirtyRef.current = dirty;

  useEffect(() => {
    let live = true;
    const online = () => setState(current => ({ ...current, online: true }));
    const offline = () => setState(current => ({ ...current, online: false }));
    const controllerMessage = (event: MessageEvent) => {
      if (event.data?.type === 'CHECK_WORKBOOK_DIRTY') {
        event.ports[0]?.postMessage({ dirty: dirtyRef.current });
      }
    };
    const refreshStatus = async (worker?: ServiceWorker) => {
      if (!worker) return;
      try {
        const reply = await request(worker, { type: 'GET_STATUS' });
        if (live && reply.type === 'STATUS') setState(current => ({ ...current, shellReady: true,
          runtimeReady: !!reply.ready, cached: reply.cached ?? 0, total: reply.total ?? 0, error: undefined }));
      } catch (error) {
        if (live) setState(current => ({ ...current, error: error instanceof Error ? error.message : String(error) }));
      }
    };
    const showWaiting = (registration: ServiceWorkerRegistration) => {
      if (registration.waiting && navigator.serviceWorker.controller) {
        setUpdate(registration.waiting);
        setUpdateHidden(false);
      }
    };
    const stateChanged = (registration: ServiceWorkerRegistration) => {
      const installing = registration.installing;
      if (installing) installing.addEventListener('statechange', () => showWaiting(registration));
      showWaiting(registration);
    };

    window.addEventListener('online', online);
    window.addEventListener('offline', offline);
    navigator.serviceWorker?.addEventListener('message', controllerMessage);
    if ('serviceWorker' in navigator && import.meta.env.PROD) {
      const base = new URL(import.meta.env.BASE_URL, document.baseURI);
      void navigator.serviceWorker.register(new URL('service-worker.js', base), {
        scope: base.pathname,
        updateViaCache: 'none',
      }).then(async registration => {
        if (!live) return;
        registrationRef.current = registration;
        registration.addEventListener('updatefound', () => stateChanged(registration));
        stateChanged(registration);
        await navigator.serviceWorker.ready;
        if (!live) return;
        await refreshStatus(registration.active ?? navigator.serviceWorker.controller ?? undefined);
        await request(registration.active ?? navigator.serviceWorker.controller!, { type: 'CLEAN_OLD_CACHES' }).catch(() => {});
        void registration.update().catch(() => {});
      }).catch(error => {
        if (live) setState(current => ({ ...current, error: error instanceof Error ? error.message : String(error) }));
      });
    }
    return () => {
      live = false;
      window.removeEventListener('online', online);
      window.removeEventListener('offline', offline);
      navigator.serviceWorker?.removeEventListener('message', controllerMessage);
    };
  }, []);

  async function preparePython() {
    const worker = registrationRef.current?.active ?? navigator.serviceWorker.controller;
    if (!worker) {
      setState(current => ({ ...current, error: 'The offline worker is still starting. Try again in a moment.' }));
      return;
    }
    setState(current => ({ ...current, error: undefined, progress: { completed: 0, total: current.total, current: '' } }));
    try {
      const reply = await request(worker, { type: 'PREPARE_RUNTIME' }, progress => setState(current => ({ ...current,
        progress: { completed: progress.completed ?? 0, total: progress.total ?? current.total, current: progress.current ?? '' } })));
      if (reply.type === 'CACHE_COMPLETE') {
        setState(current => ({ ...current, runtimeReady: true, cached: reply.cached ?? current.total,
          total: reply.total ?? current.total, progress: undefined, error: undefined }));
      } else {
        setState(current => ({ ...current, progress: undefined,
          error: reply.code === 'storage-quota'
            ? 'There is not enough browser storage to cache Python. Free local storage and retry.'
            : reply.message ?? 'Python could not be cached. Reconnect and retry.' }));
      }
    } catch (error) {
      setState(current => ({ ...current, progress: undefined,
        error: error instanceof Error ? error.message : String(error) }));
    }
  }

  async function installUpdate() {
    if (!update || updateBusy) return;
    setUpdateBusy(true);
    setState(current => ({ ...current, error: undefined }));
    try {
      if (dirty && !await onSaveRecovery()) {
        setState(current => ({ ...current, error: 'Save or finish the current edit before updating.' }));
        return;
      }
      const reply = await request(update, { type: 'ACTIVATE_UPDATE' });
      if (reply.type === 'UPDATE_DEFERRED') {
        setState(current => ({ ...current, error: reply.reason ?? 'Another tab is still editing a workbook.' }));
        return;
      }
      if (reply.type !== 'UPDATE_ACTIVATING') throw new Error(reply.message ?? 'The update could not be activated.');
      const registration = registrationRef.current;
      for (let attempt = 0; attempt < 80; attempt++) {
        await new Promise(resolve => window.setTimeout(resolve, 100));
        if (registration?.waiting === null && registration.active?.state === 'activated') {
          window.location.reload();
          return;
        }
      }
      throw new Error('The update is taking longer than expected. Reload when the workbook is safe.');
    } catch (error) {
      setState(current => ({ ...current, error: error instanceof Error ? error.message : String(error) }));
    } finally {
      setUpdateBusy(false);
    }
  }

  const base = new URL(import.meta.env.BASE_URL, document.baseURI);
  const licenseUrl = new URL(`runtime/pyodide/${runtimeLock.version}/LICENSES.md`, base).href;
  const summary = !state.supported ? 'Offline cache unavailable'
    : !state.shellReady ? 'Offline setup'
      : state.runtimeReady ? 'Offline ready' : 'Offline setup';

  return <>
    {update && !updateHidden && <aside className="offline-update" role="status" aria-live="polite">
      <span>{dirty ? 'An update is ready. Your unsaved workbook will be recovered before you reload.' : 'An update is ready.'}</span>
      <button type="button" disabled={updateBusy} onClick={() => { void installUpdate(); }}>
        {updateBusy ? 'Updating…' : dirty ? 'Save recovery and update' : 'Reload to update'}
      </button>
      <button type="button" className="offline-dismiss" aria-label="Remind me later" onClick={() => setUpdateHidden(true)}>Later</button>
    </aside>}
    <div className="offline-menu">
      <button type="button" className="offline-status-button" aria-expanded={panelOpen}
        aria-controls="offline-readiness-panel" onClick={() => setPanelOpen(value => !value)}>
        {state.online ? summary : 'Offline'}
      </button>
      {panelOpen && <section id="offline-readiness-panel" className="offline-panel" aria-label="Offline availability">
        <h2>Offline availability</h2>
        <p>{!state.supported ? 'This browser does not support offline caching.'
          : !state.shellReady ? 'The app shell is being cached for offline use.'
            : 'The app shell is cached. Offline Python is ' + (state.runtimeReady ? 'ready.' : 'not ready yet.')}</p>
        {state.total > 0 && <p>{state.cached} of {state.total} Python runtime files cached.</p>}
        {state.progress && <p role="status" aria-live="polite">Caching Python {state.progress.completed} of {state.progress.total}…</p>}
        {state.error && <p className="offline-error" role="alert">{state.error}</p>}
        <div className="offline-actions">
          <button type="button" className="file-primary" disabled={!state.supported || !state.shellReady || !state.online || !!state.progress || state.runtimeReady}
            onClick={() => { void preparePython(); }}>
            {state.runtimeReady ? 'Python ready offline' : state.progress ? 'Caching Python…' : 'Prepare Python for offline use'}
          </button>
          <a href={licenseUrl} target="_blank" rel="noreferrer">Python package notices</a>
        </div>
        {state.error && <button type="button" className="offline-retry" disabled={!state.online}
          onClick={() => { void preparePython(); }}>Retry Python cache</button>}
      </section>}
    </div>
  </>;
}
