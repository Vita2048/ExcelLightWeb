import { describe, expect, it, vi } from 'vitest';
import { defaultWorkspacePreferences, LocalWorkspaceError, LocalWorkspaceStore, migrateLegacyPreferences, WorkspaceLease } from '../../src/persistence/local';

describe('local workspace preference migration', () => {
  it('starts native and Python calculation in Automatic mode', () => {
    expect(defaultWorkspacePreferences.nativeCalculation).toBe('automatic');
    expect(defaultWorkspacePreferences.pythonCalculation).toBe('automatic');
  });

  it('migrates desktop settings fields with both calculation modes defaulting to Automatic', () => {
    const migrated = migrateLegacyPreferences({
      id: 'user', version: 1, manual: true, findHistory: ['SLA'], replacementHistory: ['Open'],
    });
    expect(migrated.nativeCalculation).toBe('manual');
    expect(migrated.pythonCalculation).toBe('automatic');
    expect(migrated.searchHistory).toEqual({ find: ['SLA'], replace: ['Open'] });
    expect(migrated.view).toEqual(defaultWorkspacePreferences.view);
  });

  it('migrates old Python Automatic consent and mode values', () => {
    expect(migrateLegacyPreferences({ schemaVersion: 1, id: 'user', nativeCalculation: 'automatic',
      pythonAutomaticConsent: true }).pythonCalculation).toBe('automatic');
    expect(migrateLegacyPreferences({ schemaVersion: 1, id: 'user', nativeCalculation: 'automatic',
      pythonAutomaticConsent: false }).pythonCalculation).toBe('automatic');
    expect(migrateLegacyPreferences({ manual: false, pythonAutomatic: true }).pythonCalculation).toBe('automatic');
    expect(migrateLegacyPreferences({ manual: false, pythonAutomatic: false }).pythonCalculation).toBe('manual');
  });

  it.each([
    { manual: 'manual' },
    { manual: false, pythonAutomatic: 'yes' },
    { findHistory: Array.from({ length: 16 }, (_, index) => String(index)) },
    { findHistory: ['same', 'same'] },
    { recentFunctions: ['sum'] },
    { recentFunctions: ['SUM', 'SUM'] },
    { recentFunctions: Array.from({ length: 11 }, (_, index) => `F${index}`) },
  ])('rejects invalid preference data without normalizing it silently: %o', value => {
    expect(() => migrateLegacyPreferences(value)).toThrow(LocalWorkspaceError);
  });

  it('preserves valid recentFunctions on preference migration', () => {
    const migrated = migrateLegacyPreferences({
      ...defaultWorkspacePreferences,
      recentFunctions: ['XLOOKUP', 'SUM', 'IF'],
    });
    expect(migrated.recentFunctions).toEqual(['XLOOKUP', 'SUM', 'IF']);
  });
});

describe('workspace tab leases', () => {
  it('waits for pagehide lease release before reacquiring on pageshow', async () => {
    vi.stubGlobal('BroadcastChannel', undefined);
    vi.useFakeTimers();
    const operations: string[] = [];
    let finishRelease!: () => void;
    const releaseGate = new Promise<void>(resolve => { finishRelease = resolve; });
    let holdRelease = true;
    const store = {
      acquireLease: vi.fn(async () => { operations.push('acquire'); return true; }),
      renewLease: vi.fn(async () => true),
      releaseLease: vi.fn(async () => {
        operations.push('release-start');
        if (holdRelease) { holdRelease = false; await releaseGate; }
        operations.push('release-done');
      }),
    } as unknown as LocalWorkspaceStore;
    const lease = new WorkspaceLease(store, vi.fn());
    try {
      expect(await lease.switchTo('book')).toBe(true);
      const suspending = lease.suspend();
      const resuming = lease.resume();
      await Promise.resolve();
      expect(operations).toEqual(['acquire', 'release-start']);

      finishRelease();
      await Promise.all([suspending, resuming]);
      expect(operations).toEqual(['acquire', 'release-start', 'release-done', 'acquire']);
      expect(lease.activeId).toBe('book');
    } finally {
      await lease.release();
      vi.useRealTimers();
      vi.unstubAllGlobals();
    }
  });

  it('reclaims an expired background lease when unclaimed and reports a real takeover', async () => {
    vi.stubGlobal('BroadcastChannel', undefined);
    vi.useFakeTimers();
    let claimable = true;
    const store = {
      acquireLease: vi.fn(async () => claimable),
      renewLease: vi.fn(async () => false),
      releaseLease: vi.fn(async () => {}),
    } as unknown as LocalWorkspaceStore;
    const onLost = vi.fn();
    const lease = new WorkspaceLease(store, onLost);
    try {
      expect(await lease.switchTo('book')).toBe(true);
      await vi.advanceTimersByTimeAsync(5_000);
      expect(store.acquireLease).toHaveBeenCalledTimes(2);
      expect(lease.isLost).toBe(false);
      expect(onLost).not.toHaveBeenCalled();

      claimable = false;
      await vi.advanceTimersByTimeAsync(5_000);
      expect(lease.isLost).toBe(true);
      expect(onLost).toHaveBeenCalledWith('book');
    } finally {
      await lease.release();
      vi.useRealTimers();
      vi.unstubAllGlobals();
    }
  });
});
