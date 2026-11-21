import type { SearchLook } from '../../application/search-commands';

const databaseName = 'excellight-workspace';
const databaseVersion = 2;
const preferencesKey = 'user';
const recoveryKeyPrefix = 'recovery:';
export const workspaceRecentLimit = 10;
const recoveryLimit = 32;
const leaseDurationMs = 20_000;
const leaseOwnerKey = 'excellight-workspace-lease-owner';
const leaseChannelName = 'excellight-workspace-leases';

export interface PersistedView {
  readonly zoom: number;
  readonly headings: boolean;
  readonly gridlines: boolean;
  readonly formulaBar: boolean;
  readonly showFormulas: boolean;
}

export interface LocalWorkspacePreferences {
  readonly schemaVersion: 1;
  readonly id: typeof preferencesKey;
  readonly nativeCalculation: 'automatic' | 'manual';
  /** Remains opt-in; loading a workbook never enables Python execution. */
  readonly pythonAutomaticConsent: boolean;
  readonly view: PersistedView;
  readonly searchHistory: { readonly find: readonly string[]; readonly replace: readonly string[] };
  readonly searchLook: SearchLook;
}

export interface WorkspaceFileHandle {
  readonly name: string;
  getFile(): Promise<File>;
  createWritable(options?: { readonly keepExistingData?: boolean }): Promise<WorkspaceWritableFile>;
  queryPermission?(descriptor: { readonly mode: 'readwrite' }): Promise<PermissionState>;
  requestPermission?(descriptor: { readonly mode: 'readwrite' }): Promise<PermissionState>;
}

export interface WorkspaceWritableFile {
  write(data: Blob): Promise<void>;
  close(): Promise<void>;
  abort?(reason?: unknown): Promise<void>;
}

export interface LocalWorkspaceRecord {
  readonly id: string;
  readonly name: string;
  readonly package: Blob;
  readonly createdAt: number;
  readonly updatedAt: number;
  readonly lastUsedAt: number;
  readonly source: 'native' | 'xlsx';
  readonly sourceFileName?: string;
  readonly fileHandle?: WorkspaceFileHandle;
  readonly fileSize?: number;
  readonly fileLastModified?: number;
}

export interface RecoveryRecord {
  readonly id: string;
  readonly workbookId: string;
  readonly name: string;
  readonly package: Blob;
  readonly savedAt: number;
  readonly revision: number;
}

interface WorkspaceLeaseRecord {
  readonly id: string;
  readonly owner: string;
  readonly expiresAt: number;
}

export class LocalWorkspaceError extends Error {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = 'LocalWorkspaceError';
  }
}

export class WorkspaceConflictError extends LocalWorkspaceError {
  constructor(readonly workbookId: string) {
    super('This workbook is already open in another tab. Your current workbook was kept.');
    this.name = 'WorkspaceConflictError';
  }
}

export const defaultWorkspacePreferences: LocalWorkspacePreferences = Object.freeze({
  schemaVersion: 1,
  id: preferencesKey,
  nativeCalculation: 'automatic',
  pythonAutomaticConsent: false,
  view: Object.freeze({ zoom: 100, headings: true, gridlines: true, formulaBar: true, showFormulas: false }),
  searchHistory: Object.freeze({ find: Object.freeze([]), replace: Object.freeze([]) }),
  searchLook: 'values',
});

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function validHistory(value: unknown): value is readonly string[] {
  return Array.isArray(value) && value.length <= 15 && value.every(item => typeof item === 'string' &&
    item.length <= 32_767 && !item.includes('\0')) && new Set(value).size === value.length;
}

export function migrateLegacyPreferences(value: unknown): LocalWorkspacePreferences {
  if (value === undefined) return defaultWorkspacePreferences;
  if (!isRecord(value)) throw new LocalWorkspaceError('Local settings are malformed; existing workspace data was kept.');

  const nativeCalculation = value.nativeCalculation ?? (value.manual === true ? 'manual' :
    value.manual === false || value.manual === undefined ? 'automatic' : undefined);
  if (nativeCalculation !== 'automatic' && nativeCalculation !== 'manual') {
    throw new LocalWorkspaceError('Local calculation preference is invalid; existing workspace data was kept.');
  }
  const pythonAutomaticConsent = value.pythonAutomaticConsent ?? value.pythonAutomatic ?? false;
  if (typeof pythonAutomaticConsent !== 'boolean') {
    throw new LocalWorkspaceError('Local Python preference is invalid; existing workspace data was kept.');
  }
  const find = value.findHistory ?? [];
  const replace = value.replacementHistory ?? [];
  if (!validHistory(find) || !validHistory(replace)) {
    throw new LocalWorkspaceError('Local search history is invalid; existing workspace data was kept.');
  }
  return {
    ...defaultWorkspacePreferences,
    nativeCalculation,
    pythonAutomaticConsent,
    searchHistory: { find, replace },
  };
}

function validatePreferences(value: unknown): LocalWorkspacePreferences {
  if (!isRecord(value) || value.schemaVersion !== 1 || value.id !== preferencesKey ||
      (value.nativeCalculation !== 'automatic' && value.nativeCalculation !== 'manual') ||
      typeof value.pythonAutomaticConsent !== 'boolean' ||
      (value.searchLook !== 'values' && value.searchLook !== 'formulas') || !isRecord(value.view) ||
      !Number.isInteger(value.view.zoom) || (value.view.zoom as number) < 10 || (value.view.zoom as number) > 400 ||
      typeof value.view.headings !== 'boolean' || typeof value.view.gridlines !== 'boolean' ||
      typeof value.view.formulaBar !== 'boolean' || typeof value.view.showFormulas !== 'boolean' ||
      !isRecord(value.searchHistory) || !validHistory(value.searchHistory.find) || !validHistory(value.searchHistory.replace)) {
    throw new LocalWorkspaceError('Local settings are invalid; existing workspace data was kept.');
  }
  return value as unknown as LocalWorkspacePreferences;
}

function requestResult<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new LocalWorkspaceError('IndexedDB request failed.'));
  });
}

function transactionDone(transaction: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onabort = transaction.onerror = () => reject(transaction.error ?? new LocalWorkspaceError('IndexedDB transaction failed.'));
  });
}

function restorePackage(value: unknown): Blob {
  if (value instanceof Blob) return value;
  if (value instanceof ArrayBuffer) return new Blob([value], { type: 'application/x-excellight' });
  throw new LocalWorkspaceError('A locally stored workbook package is invalid.');
}

function restoreWorkspace(value: unknown): LocalWorkspaceRecord {
  if (!isRecord(value)) throw new LocalWorkspaceError('A locally stored workbook record is invalid.');
  return { ...value, package: restorePackage(value.package) } as unknown as LocalWorkspaceRecord;
}

function restoreRecovery(value: unknown): RecoveryRecord {
  if (!isRecord(value)) throw new LocalWorkspaceError('A local recovery record is invalid.');
  return { ...value, package: restorePackage(value.package) } as unknown as RecoveryRecord;
}

function idbAvailable(): IDBFactory {
  if (typeof indexedDB === 'undefined') throw new LocalWorkspaceError('Browser-local storage is unavailable. Use Export to download a backup.');
  return indexedDB;
}

export class LocalWorkspaceStore {
  #database: Promise<IDBDatabase> | undefined;

  #open(): Promise<IDBDatabase> {
    if (this.#database) return this.#database;
    const factory = idbAvailable();
    this.#database = new Promise<IDBDatabase>((resolve, reject) => {
      const request = factory.open(databaseName, databaseVersion);
      let upgradeError: unknown;
      request.onupgradeneeded = event => {
        const db = request.result;
        const transaction = request.transaction!;
        try {
          if (event.oldVersion < 1) {
            db.createObjectStore('workspaces', { keyPath: 'id' });
            db.createObjectStore('settings', { keyPath: 'id' });
            db.createObjectStore('recovery', { keyPath: 'id' });
            db.createObjectStore('locks', { keyPath: 'id' });
          }
          if (event.oldVersion < 2) {
            const workspaces = transaction.objectStore('workspaces');
            if (!workspaces.indexNames.contains('lastUsedAt')) workspaces.createIndex('lastUsedAt', 'lastUsedAt');
            const preferences = db.createObjectStore('preferences', { keyPath: 'id' });
            const settings = transaction.objectStore('settings');
            const legacy = settings.get(preferencesKey);
            legacy.onsuccess = () => {
              try {
                preferences.put(migrateLegacyPreferences(legacy.result));
                db.deleteObjectStore('settings');
              } catch (error) {
                upgradeError = error;
                transaction.abort();
              }
            };
            legacy.onerror = () => {
              upgradeError = legacy.error;
              transaction.abort();
            };
          }
        } catch (error) {
          upgradeError = error;
          transaction.abort();
        }
      };
      request.onsuccess = () => {
        const db = request.result;
        db.onversionchange = () => db.close();
        resolve(db);
      };
      request.onerror = () => reject(upgradeError instanceof Error ? upgradeError :
        request.error ?? new LocalWorkspaceError('Could not open browser-local storage.'));
      request.onblocked = () => reject(new LocalWorkspaceError('Close another ExcelLight tab to finish the local workspace update.'));
    }).catch(error => {
      this.#database = undefined;
      throw error;
    });
    return this.#database;
  }

  async loadPreferences(): Promise<LocalWorkspacePreferences> {
    const db = await this.#open();
    const transaction = db.transaction('preferences', 'readonly');
    const done = transactionDone(transaction);
    const value = await requestResult(transaction.objectStore('preferences').get(preferencesKey));
    await done;
    return value === undefined ? defaultWorkspacePreferences : validatePreferences(value);
  }

  async savePreferences(value: LocalWorkspacePreferences): Promise<void> {
    const preferences = validatePreferences(value);
    const db = await this.#open();
    const transaction = db.transaction('preferences', 'readwrite');
    const done = transactionDone(transaction);
    transaction.objectStore('preferences').put(preferences);
    await done;
  }

  async listRecent(): Promise<readonly LocalWorkspaceRecord[]> {
    const db = await this.#open();
    const transaction = db.transaction('workspaces', 'readonly');
    const done = transactionDone(transaction);
    const values = await requestResult(transaction.objectStore('workspaces').getAll());
    await done;
    return values.map(restoreWorkspace).sort((left, right) => right.lastUsedAt - left.lastUsedAt)
      .slice(0, workspaceRecentLimit);
  }

  async getWorkspace(id: string): Promise<LocalWorkspaceRecord | undefined> {
    const db = await this.#open();
    const transaction = db.transaction('workspaces', 'readonly');
    const done = transactionDone(transaction);
    const value = await requestResult(transaction.objectStore('workspaces').get(id));
    await done;
    return value === undefined ? undefined : restoreWorkspace(value);
  }

  async removeWorkspace(id: string, now = Date.now()): Promise<void> {
    const db = await this.#open();
    const transaction = db.transaction(['locks', 'workspaces', 'recovery'], 'readwrite');
    const done = transactionDone(transaction);
    const locks = transaction.objectStore('locks');
    const lockRequest = locks.get(id);
    let conflict = false;
    let requestError: DOMException | null = null;
    lockRequest.onerror = () => { requestError = lockRequest.error; };
    lockRequest.onsuccess = () => {
      const lock = lockRequest.result as WorkspaceLeaseRecord | undefined;
      if (lock && lock.expiresAt > now) {
        conflict = true;
        transaction.abort();
        return;
      }
      transaction.objectStore('workspaces').delete(id);
      transaction.objectStore('recovery').delete(recoveryKeyPrefix + id);
      if (lock) locks.delete(id);
    };
    try { await done; }
    catch (error) {
      if (conflict) throw new WorkspaceConflictError(id);
      if (requestError) throw requestError;
      throw error;
    }
  }

  async listRecoveries(): Promise<readonly RecoveryRecord[]> {
    const db = await this.#open();
    const transaction = db.transaction('recovery', 'readonly');
    const done = transactionDone(transaction);
    const values = await requestResult(transaction.objectStore('recovery').getAll());
    await done;
    return values.map(restoreRecovery).sort((left, right) => right.savedAt - left.savedAt).slice(0, recoveryLimit);
  }

  async acquireLease(id: string, owner: string, now = Date.now()): Promise<boolean> {
    const db = await this.#open();
    const transaction = db.transaction('locks', 'readwrite');
    const done = transactionDone(transaction);
    const locks = transaction.objectStore('locks');
    const currentRequest = locks.get(id);
    let acquired = false;
    currentRequest.onsuccess = () => {
      const current = currentRequest.result as WorkspaceLeaseRecord | undefined;
      if (current && current.owner !== owner && current.expiresAt > now) return;
      locks.put({ id, owner, expiresAt: now + leaseDurationMs } satisfies WorkspaceLeaseRecord);
      acquired = true;
    };
    await done;
    return acquired;
  }

  async renewLease(id: string, owner: string, now = Date.now()): Promise<boolean> {
    const db = await this.#open();
    const transaction = db.transaction('locks', 'readwrite');
    const done = transactionDone(transaction);
    const locks = transaction.objectStore('locks');
    const currentRequest = locks.get(id);
    let renewed = false;
    currentRequest.onsuccess = () => {
      const current = currentRequest.result as WorkspaceLeaseRecord | undefined;
      if (!current || current.owner !== owner || current.expiresAt <= now) return;
      locks.put({ id, owner, expiresAt: now + leaseDurationMs } satisfies WorkspaceLeaseRecord);
      renewed = true;
    };
    await done;
    return renewed;
  }

  async releaseLease(id: string, owner: string): Promise<void> {
    const db = await this.#open();
    const transaction = db.transaction('locks', 'readwrite');
    const done = transactionDone(transaction);
    const locks = transaction.objectStore('locks');
    const currentRequest = locks.get(id);
    currentRequest.onsuccess = () => {
      const current = currentRequest.result as WorkspaceLeaseRecord | undefined;
      if (current?.owner === owner) locks.delete(id);
    };
    await done;
  }

  async saveWorkspace(record: LocalWorkspaceRecord, owner: string, now = Date.now()): Promise<LocalWorkspaceRecord> {
    if (!(record.package instanceof Blob) || !record.id || !record.name.trim()) {
      throw new LocalWorkspaceError('A valid workbook snapshot and name are required.');
    }
    // WebKit can reject Blob-backed IndexedDB records. Persist portable bytes instead,
    // then restore a Blob at the adapter boundary for the native package reader.
    const packageBytes = await record.package.arrayBuffer();
    const db = await this.#open();
    const transaction = db.transaction(['locks', 'workspaces'], 'readwrite');
    const done = transactionDone(transaction);
    const locks = transaction.objectStore('locks');
    const workspaces = transaction.objectStore('workspaces');
    const lockRequest = locks.get(record.id);
    let saved: LocalWorkspaceRecord | undefined;
    let conflict = false;
    let requestError: DOMException | null = null;
    lockRequest.onerror = () => { requestError = lockRequest.error; };
    lockRequest.onsuccess = () => {
      const lock = lockRequest.result as WorkspaceLeaseRecord | undefined;
      if (!lock || lock.owner !== owner || lock.expiresAt <= now) {
        conflict = true;
        transaction.abort();
        return;
      }
      const oldRequest = workspaces.get(record.id);
      oldRequest.onerror = () => { requestError = oldRequest.error; };
      oldRequest.onsuccess = () => {
        const old = oldRequest.result as LocalWorkspaceRecord | undefined;
        saved = {
          ...record,
          createdAt: old?.createdAt ?? record.createdAt ?? now,
          updatedAt: now,
          lastUsedAt: now,
        };
        const writeRequest = workspaces.put({ ...saved, package: packageBytes });
        writeRequest.onerror = () => { requestError = writeRequest.error; };
      };
    };
    try {
      await done;
    } catch (error) {
      if (conflict) throw new WorkspaceConflictError(record.id);
      if (requestError) throw requestError;
      throw error;
    }
    if (!saved) throw new LocalWorkspaceError('Workbook could not be saved to this browser.');
    return saved;
  }

  async writeRecovery(record: RecoveryRecord, owner: string, now = Date.now()): Promise<void> {
    if (!(record.package instanceof Blob) || record.id !== recoveryKeyPrefix + record.workbookId) {
      throw new LocalWorkspaceError('Recovery snapshot is invalid.');
    }
    const packageBytes = await record.package.arrayBuffer();
    const db = await this.#open();
    const transaction = db.transaction(['locks', 'recovery'], 'readwrite');
    const done = transactionDone(transaction);
    const locks = transaction.objectStore('locks');
    const recovery = transaction.objectStore('recovery');
    const lockRequest = locks.get(record.workbookId);
    let conflict = false;
    lockRequest.onsuccess = () => {
      const lock = lockRequest.result as WorkspaceLeaseRecord | undefined;
      if (!lock || lock.owner !== owner || lock.expiresAt <= now) {
        conflict = true;
        transaction.abort();
        return;
      }
      recovery.put({ ...record, package: packageBytes });
    };
    try { await done; }
    catch (error) { if (conflict) throw new WorkspaceConflictError(record.workbookId); throw error; }
  }

  async clearRecovery(workbookId: string, owner: string, now = Date.now()): Promise<void> {
    const db = await this.#open();
    const transaction = db.transaction(['locks', 'recovery'], 'readwrite');
    const done = transactionDone(transaction);
    const locks = transaction.objectStore('locks');
    const recovery = transaction.objectStore('recovery');
    const lockRequest = locks.get(workbookId);
    let conflict = false;
    lockRequest.onsuccess = () => {
      const lock = lockRequest.result as WorkspaceLeaseRecord | undefined;
      if (!lock || lock.owner !== owner || lock.expiresAt <= now) {
        conflict = true;
        transaction.abort();
        return;
      }
      recovery.delete(recoveryKeyPrefix + workbookId);
    };
    try { await done; }
    catch (error) { if (conflict) throw new WorkspaceConflictError(workbookId); throw error; }
  }
}

export class WorkspaceLease {
  owner: string;
  #activeId: string | undefined;
  #lastActiveId: string | undefined;
  #heartbeat: ReturnType<typeof setInterval> | undefined;
  #lost = false;
  #suspended = false;
  #ownerChecked = false;
  #channel: BroadcastChannel | undefined;
  #pendingProbe: { readonly nonce: string; readonly found: () => void } | undefined;

  constructor(private readonly store: LocalWorkspaceStore, private readonly onLost: (workbookId: string) => void) {
    let owner: string = crypto.randomUUID();
    try {
      owner = sessionStorage.getItem(leaseOwnerKey) || owner;
      sessionStorage.setItem(leaseOwnerKey, owner);
    } catch { /* Storage may be disabled; retain a per-page owner. */ }
    this.owner = owner;
    this.#openChannel();
  }

  #openChannel() {
    if (typeof BroadcastChannel !== 'undefined') {
      this.#channel = new BroadcastChannel(leaseChannelName);
      this.#channel.addEventListener('message', event => {
        const message = event.data as { type?: unknown; owner?: unknown; nonce?: unknown };
        if (!message || typeof message !== 'object' || typeof message.owner !== 'string' || typeof message.nonce !== 'string') return;
        if (message.type === 'probe' && message.owner === this.owner && !this.#suspended) {
          this.#channel?.postMessage({ type: 'ack', owner: this.owner, nonce: message.nonce });
        } else if (message.type === 'ack' && message.nonce === this.#pendingProbe?.nonce && message.owner === this.owner) {
          this.#pendingProbe.found();
        }
      });
    }
  }

  get activeId() { return this.#activeId; }
  get isLost() { return this.#lost; }

  async suspend(): Promise<void> {
    this.#suspended = true;
    this.#lastActiveId = this.#activeId;
    this.#channel?.close();
    this.#channel = undefined;
    await this.release();
  }

  async resume(): Promise<void> {
    this.#openChannel();
    this.#suspended = false;
    const id = this.#lastActiveId;
    this.#lastActiveId = undefined;
    if (id && !await this.switchTo(id)) this.onLost(id);
  }

  async #checkOwnerIdentity(): Promise<void> {
    if (this.#ownerChecked) return;
    this.#ownerChecked = true;
    const channel = this.#channel;
    if (!channel) return;
    const nonce = crypto.randomUUID();
    let duplicate = false;
    await new Promise<void>(resolve => {
      const timer = setTimeout(() => {
        this.#pendingProbe = undefined;
        resolve();
      }, 100);
      this.#pendingProbe = { nonce, found: () => {
        duplicate = true;
        clearTimeout(timer);
        this.#pendingProbe = undefined;
        resolve();
      } };
      channel.postMessage({ type: 'probe', owner: this.owner, nonce });
    });
    if (duplicate) {
      this.owner = crypto.randomUUID();
      try { sessionStorage.setItem(leaseOwnerKey, this.owner); } catch { /* Lease still uses the unique in-memory owner. */ }
    }
  }

  async switchTo(workbookId: string): Promise<boolean> {
    await this.#checkOwnerIdentity();
    if (this.#activeId === workbookId && !this.#lost) return true;
    const acquired = await this.store.acquireLease(workbookId, this.owner);
    if (!acquired) return false;
    const prior = this.#activeId;
    clearInterval(this.#heartbeat);
    this.#activeId = workbookId;
    this.#lost = false;
    this.#heartbeat = setInterval(() => { void this.#renew(workbookId); }, 5_000);
    if (prior) await this.store.releaseLease(prior, this.owner);
    return true;
  }

  async release(): Promise<void> {
    clearInterval(this.#heartbeat);
    this.#heartbeat = undefined;
    const id = this.#activeId;
    this.#activeId = undefined;
    this.#lost = false;
    if (id) await this.store.releaseLease(id, this.owner);
  }

  async #renew(workbookId: string) {
    if (this.#activeId !== workbookId || this.#lost) return;
    try {
      if (await this.store.renewLease(workbookId, this.owner)) return;
    } catch { /* A failed renewal is treated as a lost lease. */ }
    this.#lost = true;
    clearInterval(this.#heartbeat);
    this.#heartbeat = undefined;
    this.onLost(workbookId);
  }
}
